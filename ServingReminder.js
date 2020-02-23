/**
 * This Program creates a serving reminder as a google doc
 * 
 * Author: Eric Dong
 * Creation Date: 2/22/2020
 * Last Modfied: 2/23/2020
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
"Lunch Service"
];
  
const createServingReminder = () => {
  const data = oifSchedule.getDataRange().getValues();

  // Get column names
  const columnNames = data[0];

  // Get the appropriate Sunday Date for the reminder
  // Get current Date
  const currentDate = new Date();
  
  // Get data for the upcoming Sunday
  const currentRow = data.filter((d, i) =>
    currentDate <= d[0] 
  )[0].map(row => row);
  const dataMap = columnNames.reduce((acc, name, index) => {
    // Parentheses are problematic with replacing text in the template
    name = name.replace(/[()]/g, '');
    acc[name] = currentRow[index];
    return acc;
  }, {});

  // Create a copy of the google doc template
  const templateCopy = DriveApp.getFileById(googleDocTemplateId).makeCopy();
  const body = DocumentApp.openById(templateCopy.getId()).getBody();

  // Replace template text with the data in OIF Schedule
  templateTexts.forEach(text => {
    const templateString = "{" + text + "}";
    body.replaceText(templateString, dataMap[text]);
  });
}