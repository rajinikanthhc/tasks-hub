const SHEET_NAME = "Cards";

const GITHUB_OWNER = "rajinikanthhc";
const GITHUB_REPO = "images";
const GITHUB_FOLDER = "visiting-cards";


function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Visiting Cards Hub");

}


function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =================================
   GET ALL CARDS
================================= */

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


/* =================================
   ADD NEW CARD
   USE SMALLEST AVAILABLE ID
================================= */

function addCard(card, imageData) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


    const newId = getSmallestAvailableId(sheet);


    /* -----------------------------
       UPLOAD IMAGE
    ----------------------------- */

    let imageName = "";


    if (imageData && imageData.base64) {

      imageName =
        createImageName(
          card.Name,
          card.Company
        );

      uploadToGitHub(
        imageName,
        imageData.base64
      );

    }


    /* -----------------------------
       SAVE TO SHEET
    ----------------------------- */

    sheet.appendRow([

      newId,

      card.Name || "",

      card.Company || "",

      card.Designation || "",

      card.Area || "",

      card.Mobile || "",

      card.Email || "",

      imageName,

      card.Notes || ""

    ]);


    return newId;

  }

  finally {

    lock.releaseLock();

  }

}


/* =================================
   FIND SMALLEST AVAILABLE ID
================================= */

function getSmallestAvailableId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {

    return "VC0001";

  }


  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();


  const usedNumbers = new Set();


  ids.forEach(id => {

    const match =
      String(id).match(/^VC(\d+)$/i);

    if (match) {

      usedNumbers.add(
        Number(match[1])
      );

    }

  });


  let number = 1;


  while (usedNumbers.has(number)) {

    number++;

  }


  return "VC" +
    String(number).padStart(4, "0");

}


/* =================================
   EDIT CARD
================================= */

function updateCard(card, imageData) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);


  const data =
    sheet.getDataRange().getValues();


  const headers =
    data[0];


  const idColumn =
    headers.indexOf("ID");


  const rowIndex =
    data.findIndex((row, index) => {

      if (index === 0) {
        return false;
      }

      return String(row[idColumn]) ===
        String(card.ID);

    });


  if (rowIndex === -1) {

    throw new Error(
      "Card not found: " + card.ID
    );

  }


  const sheetRow =
    rowIndex + 1;


  /* -----------------------------
     KEEP EXISTING IMAGE
  ----------------------------- */

  let imageName =
    card.Image || "";


  /* -----------------------------
     UPLOAD NEW IMAGE IF SELECTED
  ----------------------------- */

  if (imageData && imageData.base64) {

    imageName =
      createImageName(
        card.Name,
        card.Company
      );


    uploadToGitHub(
      imageName,
      imageData.base64
    );

  }


  /* -----------------------------
     UPDATE ROW
  ----------------------------- */

  sheet.getRange(sheetRow, 1, 1, 9)
    .setValues([[
      card.ID || "",
      card.Name || "",
      card.Company || "",
      card.Designation || "",
      card.Area || "",
      card.Mobile || "",
      card.Email || "",
      imageName,
      card.Notes || ""
    ]]);


  return card.ID;

}


/* =================================
   CREATE IMAGE NAME
================================= */

function createImageName(name, company) {

  let filename =
    (name || "Unknown") +
    "-" +
    (company || "Company");


  filename = filename
    .replace(/[\/\\:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();


  return filename + ".png";

}


/* =================================
   UPLOAD TO GITHUB
================================= */

function uploadToGitHub(filename, base64Data) {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty("GITHUB_TOKEN");


  if (!token) {

    throw new Error(
      "GitHub token not found in Script Properties."
    );

  }


  const path =
    GITHUB_FOLDER +
    "/" +
    filename;


  const url =
    "https://api.github.com/repos/" +
    GITHUB_OWNER +
    "/" +
    GITHUB_REPO +
    "/contents/" +
    encodeURIComponent(path);


  const payload = {

    message:
      "Add visiting card - " + filename,

    content:
      base64Data,

    branch:
      "main"

  };


  const response =
    UrlFetchApp.fetch(url, {

      method: "put",

      contentType:
        "application/json",

      headers: {

        Authorization:
          "Bearer " + token,

        Accept:
          "application/vnd.github+json"

      },

      payload:
        JSON.stringify(payload),

      muteHttpExceptions:
        true

    });


  const code =
    response.getResponseCode();


  if (code < 200 || code >= 300) {

    throw new Error(
      "GitHub upload failed: " +
      response.getContentText()
    );

  }

}