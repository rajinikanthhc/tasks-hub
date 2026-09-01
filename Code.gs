const SHEET_NAME = "Tasks";


/* ================================
   REMINDER EMAIL
================================ */

const REMINDER_EMAIL = "rajinikanthcrps@gmail.com";


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

  /* Make sure Repeat column exists */

  if (!sheet.getRange(1, 7).getValue()) {
    sheet.getRange(1, 7).setValue("Repeat");
  }

  const data =
    sheet
      .getRange(2, 1, lastRow - 1, 7)
      .getValues();

  return data

    .filter(function (row) {

      return row[1] !== "";

    })

    .map(function (row) {

      let timeValue = "";

      if (row[3] instanceof Date) {

        timeValue =
          Utilities.formatDate(
            row[3],
            Session.getScriptTimeZone(),
            "HH:mm"
          );

      }

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

        time: timeValue,

        status:
          row[4] ||
          "Pending",

        reminderSent:
          row[5] === true ||
          row[5] === "TRUE",

        repeat:
          row[6] ||
          "Does not repeat"

      };

    });

}


/* ================================
   ADD TASK
================================ */

function addTask(
  task,
  dueDate,
  time,
  repeat
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Tasks sheet not found.");
  }

  if (
    !task ||
    task.trim() === ""
  ) {

    throw new Error(
      "Please enter a task."
    );

  }

  const lastRow =
    sheet.getLastRow();

  let newId = 1;

  if (lastRow >= 2) {

    const ids =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
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

  let timeValue = "";

  if (time) {

    timeValue =
      new Date(
        "1970-01-01T" +
        time +
        ":00"
      );

  }

  const repeatValue =
    repeat ||
    "Does not repeat";


  sheet.appendRow([

    newId,

    task.trim(),

    dueDate
      ? new Date(dueDate)
      : "",

    timeValue,

    "Pending",

    false,

    repeatValue

  ]);


  return {

    id: newId,

    task: task.trim(),

    dueDate: dueDate || "",

    time: time || "",

    status: "Pending",

    reminderSent: false,

    repeat: repeatValue

  };

}


/* ================================
   UPDATE TASK
================================ */

function updateTask(
  id,
  task,
  dueDate,
  time,
  repeat
) {

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

  let rowNumber = -1;

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0]) ===
      String(id)
    ) {

      rowNumber =
        i + 2;

      break;

    }

  }

  if (rowNumber === -1) {
    return false;
  }


  let timeValue = "";

  if (time) {

    timeValue =
      new Date(
        "1970-01-01T" +
        time +
        ":00"
      );

  }


  const status =
    sheet
      .getRange(
        rowNumber,
        5
      )
      .getValue() ||
      "Pending";


  const repeatValue =
    repeat ||
    "Does not repeat";


  /*
     Update B:G together
  */

  sheet
    .getRange(
      rowNumber,
      2,
      1,
      6
    )
    .setValues([[
      task.trim(),

      dueDate
        ? new Date(dueDate)
        : "",

      timeValue,

      status,

      false,

      repeatValue

    ]]);


  return {

    id: id,

    task: task.trim(),

    dueDate: dueDate || "",

    time: time || "",

    status: status,

    reminderSent: false,

    repeat: repeatValue

  };

}


/* ================================
   NEXT REPEAT DATE
================================ */

function getNextRepeatDate(
  currentDate,
  repeat
) {

  const nextDate =
    new Date(currentDate);

  if (
    repeat === "Daily"
  ) {

    nextDate.setDate(
      nextDate.getDate() + 1
    );

  }

  else if (
    repeat === "Weekly"
  ) {

    nextDate.setDate(
      nextDate.getDate() + 7
    );

  }

  else if (
    repeat === "Monthly"
  ) {

    nextDate.setMonth(
      nextDate.getMonth() + 1
    );

  }

  return nextDate;

}


/* ================================
   COMPLETE / UNCOMPLETE
================================ */

function toggleTask(
  id,
  status
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {

    throw new Error(
      "Tasks sheet not found."
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getValues();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      String(id)
    ) {

      const rowNumber =
        i + 2;

      const repeat =
        data[i][6] ||
        "Does not repeat";


      /*
         RECURRING TASK
      */

      if (
        status === "Completed" &&
        repeat !== "Does not repeat"
      ) {

        const currentDate =
          data[i][2];

        if (
          currentDate instanceof Date
        ) {

          const nextDate =
            getNextRepeatDate(
              currentDate,
              repeat
            );

          sheet
            .getRange(
              rowNumber,
              3
            )
            .setValue(nextDate);

        }


        /*
           Keep it pending for next occurrence
        */

        sheet
          .getRange(
            rowNumber,
            5
          )
          .setValue("Pending");


        /*
           Reset reminder
        */

        sheet
          .getRange(
            rowNumber,
            6
          )
          .setValue(false);


        return {

          recurring: true,

          status: "Pending"

        };

      }


      /*
         NORMAL TASK
      */

      sheet
        .getRange(
          rowNumber,
          5
        )
        .setValue(status);


      return {

        recurring: false,

        status: status

      };

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

    throw new Error(
      "Tasks sheet not found."
    );

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

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0]) ===
      String(id)
    ) {

      sheet.deleteRow(
        i + 2
      );

      return true;

    }

  }

  return false;

}


/* ================================
   SEND TASK REMINDERS
================================ */

function sendTaskReminders() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {
    return;
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getValues();

  const now =
    new Date();


  data.forEach(
    function (row, index) {

      const task =
        row[1];

      const dueDate =
        row[2];

      const time =
        row[3];

      const status =
        row[4];

      const reminderSent =
        row[5];

      const repeat =
        row[6] ||
        "Does not repeat";


      if (!task) {
        return;
      }


      if (
        String(status)
          .toLowerCase() !==
        "pending"
      ) {

        return;

      }


      if (
        reminderSent === true ||
        reminderSent === "TRUE"
      ) {

        return;

      }


      if (
        !(dueDate instanceof Date) ||
        !(time instanceof Date)
      ) {

        return;

      }


      const taskDateTime =
        new Date(dueDate);

      taskDateTime.setHours(
        time.getHours(),
        time.getMinutes(),
        0,
        0
      );


      if (
        now < taskDateTime
      ) {

        return;

      }


      /*
         SEND EMAIL
      */

      MailApp.sendEmail({

        to:
          REMINDER_EMAIL,

        subject:
          "Task Reminder: " +
          task,

        htmlBody:

          "<h3>Task Reminder</h3>" +

          "<p><b>" +
          escapeEmailHtml(task) +
          "</b></p>" +

          "<p>Due: " +
          Utilities.formatDate(
            taskDateTime,
            Session.getScriptTimeZone(),
            "dd-MM-yyyy hh:mm a"
          ) +
          "</p>"

      });


      /*
         RECURRING TASK
      */

      if (
        repeat !==
        "Does not repeat"
      ) {

        const nextDate =
          getNextRepeatDate(
            dueDate,
            repeat
          );


        sheet
          .getRange(
            index + 2,
            3
          )
          .setValue(nextDate);


        /*
           Reset reminder
        */

        sheet
          .getRange(
            index + 2,
            6
          )
          .setValue(false);

      }

      else {

        /*
           Normal task
        */

        sheet
          .getRange(
            index + 2,
            6
          )
          .setValue(true);

      }

    }
  );

}


/* ================================
   ESCAPE EMAIL HTML
================================ */

function escapeEmailHtml(text) {

  return String(text)

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}


/* ================================
   CREATE REMINDER TRIGGER
================================ */

function setupReminderTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(
    function (trigger) {

      if (
        trigger
          .getHandlerFunction() ===
        "sendTaskReminders"
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );

      }

    }
  );


  ScriptApp
    .newTrigger(
      "sendTaskReminders"
    )
    .timeBased()
    .everyMinutes(1)
    .create();

}


/* ================================
   GOOGLE SHEET URL
================================ */

function getSheetUrl() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getUrl();

}


/* ================================
   TEST EMAIL
================================ */

function testTaskEmail() {

  MailApp.sendEmail({

    to: REMINDER_EMAIL,

    subject:
      "Tasks Hub - Test Email",

    body:
      "This is a test email from your Tasks Hub."

  });

}