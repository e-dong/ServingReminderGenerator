/**
 * This Program creates a serving reminder as a google doc
 *
 * Author: Eric Dong
 * Creation Date: 2/22/2020
 * Last Modfied: 2/24/2020
 *
 */

// IDs to access the documents
const googleDocTemplateId = "1GXnHc9b5--0qzySO0V7uBRf0veZ4pdipwuOMcqx01qo";
const oifScheduleSheetId = "1paaTofVVKhezdiHoPbQn8yqcDZjnf4X3GlGKA7J8Tac";

const oifSchedule = SpreadsheetApp.openById(oifScheduleSheetId);

const templateTexts = [
  "Speaker",
  "Prayer Leader",
  "Announcer",
  "Music Team Don't Edit, automatically linked",
  "Tech Team",
  "Welcome Team",
  "Bulletin",
  "CBT Teacher",
  "Lunch Service",
  "Communion"
];

const createServingReminder = () => {
  const data = oifSchedule.getDataRange().getValues();

  // Get column names
  const columnNames = data[0];

  // Get the appropriate Sunday Date for the reminder
  // Get current Date
  const currentDate = new Date();

  // Get data for the upcoming Sunday
  const currentRow = data
    .filter(
      (d, i) => currentDate <= getValueFromColumnName("Date", d, columnNames)
    )[0]
    .map(row => row);

  const dataMap = columnNames.reduce((acc, name, index) => {
    // Parentheses are problematic with replacing text in the template
    name = name.replace(/[()]/g, "");
    acc[name] = currentRow[index];

    // Check if there is Communion
    if (checkCommunion(name, currentRow[index])) acc.Communion = "yes";
    return acc;
  }, {});

  const id = buildGoogleDocsFromTemplate(dataMap);
  sendDocument(id, "eric2043@gmail.com", "test");

  // Remove copy
  Drive.Files.remove(id);
};

// Helper Functions

const buildGoogleDocsFromTemplate = dataMap => {
  // Create a copy of the google doc template
  const templateCopy = DriveApp.getFileById(googleDocTemplateId).makeCopy();
  const document = DocumentApp.openById(templateCopy.getId());
  const body = document.getBody();

  // Replace template text with the data in OIF Schedule
  templateTexts.forEach(text => {
    let value = dataMap[text];
    if (!value) value = "";
    body.replaceText(`{${text}}`, value);
  });

  document.saveAndClose();

  return document.getId();
};

/**
 * Checks if there is Communion for the upcoming Sunday.
 *
 * @param {string} columnName The column name to check for Communion
 * @param {string} content The content of the cell
 */
const checkCommunion = (columnName, content) => {
  if (columnName === "Topic") {
    return content.match(/(c|C)ommunion/g);
  }
  return false;
};

/**
 * Gets the value from the column name.
 *
 * @param {string} name The column name
 * @param {string []} row The array of values
 * @param {string []} ColumnNames The array of column names
 * @returns {string} The value from the column name
 */
const getValueFromColumnName = (name, row, ColumnNames) => {
  const index = ColumnNames.indexOf(name);
  if (index !== -1) {
    return row[index];
  }
  return undefined;
};

function sendDocument(documentId, recipient, subject) {
  var html = convertToHtml(documentId);
  html = inlineCss(html);
  GmailApp.createDraft(recipient, subject, null, {
    htmlBody: html
  });
}

function convertToHtml(fileId) {
  var file = Drive.Files.get(fileId);
  var htmlExportLink = file.exportLinks["text/html"];
  if (!htmlExportLink) {
    throw "File cannot be converted to HTML.";
  }
  var oAuthToken = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(htmlExportLink, {
    headers: {
      Authorization: "Bearer " + oAuthToken
    },
    muteHttpExceptions: true
  });
  if (!response.getResponseCode() == 200) {
    throw "Error converting to HTML: " + response.getContentText();
  }
  return response.getContentText();
}

function inlineCss(html) {
  var apikey = CacheService.getPublicCache().get("mailchimp.apikey");
  if (!apikey) {
    apikey = PropertiesService.getScriptProperties().getProperty(
      "mailchimp.apikey"
    );
    CacheService.getPublicCache().put("mailchimp.apikey", apikey);
  }
  var datacenter = apikey.split("-")[1];
  var url = Utilities.formatString(
    "https://%s.api.mailchimp.com/2.0/helper/inline-css",
    datacenter
  );
  var response = UrlFetchApp.fetch(url, {
    method: "post",
    payload: {
      apikey: apikey,
      html: html,
      strip_css: true
    }
  });
  var output = JSON.parse(response.getContentText());
  if (!response.getResponseCode() == 200) {
    throw "Error inlining CSS: " + output["error"];
  }
  return output["html"];
}
