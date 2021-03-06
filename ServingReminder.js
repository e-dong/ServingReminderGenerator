/**
 * This Program creates a serving reminder as a google doc
 *
 * Author: Eric Dong
 * Creation Date: 2/22/2020
 * Last Modfied: 2/25/2020
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

  return columnNames.reduce((acc, name, index) => {
    // Parentheses are problematic with replacing text in the template
    name = name.replace(/[()]/g, "");
    acc[name] = currentRow[index];

    // Check if there is Communion
    if (checkCommunion(name, currentRow[index])) acc.Communion = "yes";
    return acc;
  }, {});
};

// Helper Functions

const buildGoogleDocFromTemplate = dataMap => {
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

  return document;
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

const deleteFileById = id => {
  Drive.Files.remove(id);
};

function doGet(e) {
  const output = HtmlService.createTemplateFromFile("index");
  const data = createServingReminder();
  const document = buildGoogleDocFromTemplate(data);
  output.reminderUrl = document.getUrl();
  output.reminderId = document.getId();
  return output.evaluate();
}
