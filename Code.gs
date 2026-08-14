const SHEET_NAME = "Cards";

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Visiting Cards Hub");
}

function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function getCards() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers = data[0];

  return data.slice(1)
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {

      const card = {};

      headers.forEach((header, index) => {
        card[header] = row[index] || "";
      });

      return card;

    });
}