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
   ADD CARD + UPLOAD IMAGE
================================= */

function addCard(card, imageData) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);


  /* -----------------------------
     GENERATE ID
  ----------------------------- */

  const lastRow = sheet.getLastRow();

  let nextNumber = 1;

  if (lastRow > 1) {

    const ids = sheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .flat();

    const numbers = ids
      .map(id => {

        const match =
          String(id).match(/VC(\d+)/i);

        return match
          ? Number(match[1])
          : 0;

      });

    nextNumber =
      Math.max(...numbers) + 1;

  }


  const newId =
    "VC" + String(nextNumber).padStart(4, "0");


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
function authorizeGitHub() {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty("GITHUB_TOKEN");

  const response =
    UrlFetchApp.fetch("https://api.github.com/user", {

      method: "get",

      headers: {
        Authorization: "Bearer " + token,
        Accept: "application/vnd.github+json"
      },

      muteHttpExceptions: true

    });

  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());

}