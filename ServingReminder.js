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

const templateCopy = DriveApp.getFileById(googleDocTemplateId).makeCopy();
const body = DocumentApp.openById(templateCopy.getId())
      .getBody();
  
const createServingReminder = () => {
  const data = oifSchedule.getDataRange().getValues().map(function(row) {
    return row.map(function(value){ return value.toString()});
  });

  // Get column names
  const columnNames = data[0];

  // Get the appropriate Sunday Date for the reminder
  const sundayDate = new Date();
  sundayDate.setMonth(1);
  sundayDate.setDate(23);
  sundayDate.setYear(2020);
  sundayDate.setHours(0, 0, 0);
  
  // Find the current date
  const currentRow = data.filter((d, i) =>
    d.indexOf(sundayDate.toString()) !== -1)[0].map(row => row);
  const dataMap = columnNames.reduce((acc, name, index) => {
    name = name.replace(/[()]/g, '');
    acc[name] = currentRow[index];
    return acc;
  }, {});

  // Replace template text with the data in OIF Schedule
  templateTexts.forEach(text => {
    const templateString = "{" + text + "}";
    body.replaceText(templateString, dataMap[text]);
  });
}