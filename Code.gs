const SHEET_NAME = "Tasks";


/* ================================
   OPEN WEB APP
================================ */

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Tasks Hub")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* ================================
   INCLUDE HTML FILES
================================ */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* ================================
   GET TASKS
================================ */

function getTasks() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {

    throw new Error("Tasks sheet not found.");

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  const data =
    sheet
      .getRange(2, 1, lastRow - 1, 4)
      .getValues();


  return data

    .filter(function (row) {

      return row[1] !== "";

    })

    .map(function (row) {

      return {

        id: row[0],

        task: row[1],

        dueDate: row[2]
          ? Utilities.formatDate(
              new Date(row[2]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : "",

        status: row[3] || "Pending"

      };

    });

}


/* ================================
   ADD TASK
================================ */

function addTask(task, dueDate) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  if (!sheet) {

    throw new Error("Tasks sheet not found.");

  }


  if (!task || task.trim() === "") {

    throw new Error("Please enter a task.");

  }


  const lastRow =
    sheet.getLastRow();


  let newId = 1;


  if (lastRow >= 2) {

    const ids =
      sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues()
        .flat()
        .filter(function (id) {

          return id !== "";

        });


    if (ids.length > 0) {

      newId =
        Math.max(
          ...ids.map(Number)
        ) + 1;

    }

  }


  sheet.appendRow([

    newId,

    task.trim(),

    dueDate
      ? new Date(dueDate)
      : "",

    "Pending"

  ]);


  return true;

}


/* ================================
   UPDATE TASK
================================ */

function updateTask(id, task, dueDate) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  if (!sheet) {

    throw new Error("Tasks sheet not found.");

  }


  if (!task || task.trim() === "") {

    throw new Error("Please enter a task.");

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return false;

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (let i = 0; i < ids.length; i++) {

    if (
      String(ids[i][0]) ===
      String(id)
    ) {

      /* Update Task */

      sheet
        .getRange(i + 2, 2)
        .setValue(task.trim());


      /* Update Due Date */

      sheet
        .getRange(i + 2, 3)
        .setValue(
          dueDate
            ? new Date(dueDate)
            : ""
        );


      return true;

    }

  }


  return false;

}


/* ================================
   COMPLETE / UNCOMPLETE TASK
================================ */

function toggleTask(id, status) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  if (!sheet) {

    throw new Error("Tasks sheet not found.");

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return false;

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (let i = 0; i < ids.length; i++) {

    if (
      String(ids[i][0]) ===
      String(id)
    ) {

      sheet
        .getRange(i + 2, 4)
        .setValue(status);


      return true;

    }

  }


  return false;

}


/* ================================
   DELETE TASK
================================ */

function deleteTask(id) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  if (!sheet) {

    throw new Error("Tasks sheet not found.");

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return false;

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (let i = 0; i < ids.length; i++) {

    if (
      String(ids[i][0]) ===
      String(id)
    ) {

      sheet.deleteRow(i + 2);

      return true;

    }

  }


  return false;

}