const SHEET_NAME = "Cards";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Visiting Cards Hub");
}

function getCards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
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

function searchCards(searchText) {
  const cards = getCards();

  if (!searchText || !searchText.trim()) {
    return cards;
  }

  const search = searchText.toLowerCase().trim();

  return cards.filter(card => {
    return Object.values(card).some(value =>
      String(value).toLowerCase().includes(search)
    );
  });
}