/*******************************************************
 * FOODCLIQUE WORLD FOOD DAY 2026
 * ONE-SHEET REGISTRATION + QR + ATTENDANCE SYSTEM
 *******************************************************/

const CONFIG = {
  SHEET_NAME: "Form Responses 1",

  REGISTRATION_PREFIX: "WFD26-",

  MATCH_EVENT:
    "Zero Hunger Match — 10 October 2026",

  FOOD_EVENT:
    "World Food Day Food Distribution — 16 October 2026",

  TIMEZONE:
    "Africa/Lagos",

  ORGANISATION:
    "FoodClique Support Initiative",

  EMAIL_NAME:
    "FoodClique Support Initiative"
};


/*******************************************************
 * COLUMN NUMBERS
 *
 * Google Sheets columns start at 1
 *******************************************************/

const COL = {

  TIMESTAMP: 1,
  NAME: 2,
  EMAIL: 3,
  PHONE: 4,
  ORGANISATION: 5,
  TYPE: 6,
  EVENTS: 7,
  SOURCE: 8,

  REG_ID: 9,
  QR: 10,
  REG_EMAIL: 11,

  MATCH_ATTENDANCE: 12,
  MATCH_CHECKIN: 13,
  MATCH_THANKYOU: 14,
  MATCH_SENT: 15,

  FOOD_ATTENDANCE: 16,
  FOOD_CHECKIN: 17,
  FOOD_THANKYOU: 18,
  FOOD_SENT: 19
};


/*******************************************************
 * FIRST-TIME SETUP
 *******************************************************/

function setupSystem() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Could not find the Form Responses sheet."
    );
  }

  const headers = [

    "Timestamp",
    "Full Name",
    "Email Address",
    "Phone Number",
    "Organisation / Institution",
    "Participant Type",
    "Event(s) Attending",
    "How Did You Hear About Us?",

    "Registration ID",
    "QR Code",
    "Registration Email",

    "Match Attendance",
    "Match Check-in Time",
    "Match Thank-you",
    "Match Thank-you Sent",

    "Food Distribution Attendance",
    "Food Distribution Check-in Time",
    "Food Distribution Thank-you",
    "Food Distribution Thank-you Sent"
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  sheet.autoResizeColumns(
    1,
    headers.length
  );

  createTrigger();

  SpreadsheetApp.getUi().alert(
    "FoodClique World Food Day system is ready."
  );
}


/*******************************************************
 * FORM SUBMISSION
 *******************************************************/

function onFormSubmit(e) {

  const sheet =
    e.range.getSheet();

  if (
    sheet.getName() !== CONFIG.SHEET_NAME
  ) {
    return;
  }

  const row =
    e.range.getRow();

  processRegistration(row);
}


/*******************************************************
 * PROCESS NEW REGISTRATION
 *******************************************************/

function processRegistration(row) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const rowData =
    sheet
      .getRange(
        row,
        1,
        1,
        COL.FOOD_SENT
      )
      .getValues()[0];

  const name =
    rowData[COL.NAME - 1];

  const email =
    rowData[COL.EMAIL - 1];

  const phone =
    rowData[COL.PHONE - 1];

  const organisation =
    rowData[COL.ORGANISATION - 1];

  const participantType =
    rowData[COL.TYPE - 1];

  const events =
    rowData[COL.EVENTS - 1];

  // Don't process twice
  if (
    rowData[COL.REG_ID - 1]
  ) {
    return;
  }

  const registrationId =
    generateRegistrationId();

  const qrUrl =
    generateQrUrl(registrationId);

  // Write registration information
  sheet
    .getRange(row, COL.REG_ID)
    .setValue(registrationId);

  sheet
    .getRange(row, COL.QR)
    .setValue(qrUrl);

  sheet
    .getRange(row, COL.REG_EMAIL)
    .setValue("SENDING");

  try {

    sendRegistrationEmail({
      name,
      email,
      phone,
      organisation,
      participantType,
      events,
      registrationId,
      qrUrl
    });

    sheet
      .getRange(row, COL.REG_EMAIL)
      .setValue("SENT");

  } catch (error) {

    sheet
      .getRange(row, COL.REG_EMAIL)
      .setValue(
        "FAILED: " +
        error.message
      );
  }
}


/*******************************************************
 * REGISTRATION ID
 *******************************************************/

function generateRegistrationId() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const lastRow =
    sheet.getLastRow();

  if (lastRow <= 1) {

    return CONFIG.REGISTRATION_PREFIX +
      "00001";
  }

  const ids =
    sheet
      .getRange(
        2,
        COL.REG_ID,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();

  let highest = 0;

  ids.forEach(id => {

    if (!id) return;

    const number =
      parseInt(
        String(id)
          .replace(
            CONFIG.REGISTRATION_PREFIX,
            ""
          ),
        10
      );

    if (
      !isNaN(number) &&
      number > highest
    ) {
      highest = number;
    }
  });

  return CONFIG.REGISTRATION_PREFIX +
    String(highest + 1)
      .padStart(5, "0");
}


/*******************************************************
 * QR CODE
 *******************************************************/

function generateQrUrl(registrationId) {

  return (
    "https://quickchart.io/qr?size=300&text=" +
    encodeURIComponent(registrationId)
  );
}


/*******************************************************
 * REGISTRATION EMAIL
 *******************************************************/

function sendRegistrationEmail(data) {

  const qrBlob =
    UrlFetchApp
      .fetch(data.qrUrl)
      .getBlob()
      .setName("FoodClique-WFD-QR.png");

  const subject =
    "FoodClique World Food Day 2026 — Registration Confirmed";

  const html = `

  <div style="
    font-family:Arial,sans-serif;
    max-width:650px;
    margin:auto;
    line-height:1.6;
  ">

    <h2 style="color:#16833b;">
      Registration Confirmed
    </h2>

    <p>
      Dear ${escapeHtml(data.name)},
    </p>

    <p>
      Thank you for registering to participate in
      <strong>FoodClique Support Initiative's
      World Food Day 2026 Campaign</strong>,
      themed <strong>“Together Towards Zero Hunger.”</strong>
    </p>

    <p>
      Your registration details are:
    </p>

    <table
      cellpadding="8"
      cellspacing="0"
      style="border-collapse:collapse;"
    >

      <tr>
        <td><strong>Registration ID</strong></td>
        <td>${escapeHtml(data.registrationId)}</td>
      </tr>

      <tr>
        <td><strong>Participant Type</strong></td>
        <td>${escapeHtml(data.participantType)}</td>
      </tr>

      <tr>
        <td><strong>Event(s)</strong></td>
        <td>${escapeHtml(data.events)}</td>
      </tr>

    </table>

    <h3>Your Entry QR Code</h3>

    <p>
      Please keep this QR code and present it at the
      event entrance for attendance verification.
    </p>

    <div>
      <img
        src="cid:qrcode"
        width="250"
        alt="FoodClique QR Code"
      >
    </div>

    <p>
      <strong>
        Please do not share your QR code with another person.
      </strong>
    </p>

    <p>
      We look forward to having you join us as we work
      together towards zero hunger.
    </p>

    <p>
      Warm regards,<br>

      <strong>
        FoodClique Support Initiative
      </strong><br>

      info@foodclique.org
    </p>

  </div>

  `;

  GmailApp.sendEmail(
    data.email,
    subject,
    "Your FoodClique World Food Day registration is confirmed.",
    {
      htmlBody: html,
      name: CONFIG.EMAIL_NAME,

      inlineImages: {
        qrcode: qrBlob
      }
    }
  );
}


/*******************************************************
 * WEB APP / SCANNER
 *******************************************************/

function doGet(e) {

  const params = e.parameter || {};

  // Test endpoint
  if (params.api === "test") {

    return jsonpResponse(
      {
        success: true,
        message: "FoodClique WFD API is working."
      },
      params.callback
    );
  }


  // Attendance endpoint
  if (params.api === "checkin") {

    return handleCheckInApi(params);
  }


  return jsonResponse({
    success: true,
    service: "FoodClique World Food Day 2026 Attendance API",
    status: "online"
  });
}


function handleCheckInApi(params) {

  const registrationId =
    String(params.registrationId || "").trim();

  const eventName =
    String(params.event || "").trim();

  const callback =
    String(params.callback || "").trim();


  if (!registrationId) {

    return jsonpResponse(
      {
        success: false,
        message: "Registration ID is missing."
      },
      callback
    );
  }


  if (!eventName) {

    return jsonpResponse(
      {
        success: false,
        message: "Event is missing."
      },
      callback
    );
  }


  try {

    const result =
      checkInParticipantApi(
        registrationId,
        eventName
      );

    return jsonpResponse(
      result,
      callback
    );

  } catch (error) {

    return jsonpResponse(
      {
        success: false,
        message:
          "System error: " +
          error.message
      },
      callback
    );
  }
}

function checkInParticipantApi(
  registrationId,
  eventName
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();


  let participant = null;
  let participantRow = null;


  // Find registration
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const storedId =
      String(
        data[i][COL.REG_ID - 1]
      ).trim();


    if (
      storedId === registrationId
    ) {

      participantRow = i + 1;

      participant = {

        name:
          data[i][COL.NAME - 1],

        email:
          data[i][COL.EMAIL - 1],

        phone:
          data[i][COL.PHONE - 1],

        organisation:
          data[i][COL.ORGANISATION - 1],

        type:
          data[i][COL.TYPE - 1],

        events:
          String(
            data[i][COL.EVENTS - 1]
          ),

        registrationId:
          registrationId
      };

      break;
    }
  }


  // Registration doesn't exist
  if (!participant) {

    return {

      success: false,

      message:
        "Registration ID not found."
    };
  }


  /***************************************************
   * ZERO HUNGER MATCH
   ***************************************************/

  if (
    eventName ===
    CONFIG.MATCH_EVENT
  ) {

    if (
      !participant.events
        .includes("Zero Hunger Match")
    ) {

      return {

        success: false,

        message:
          participant.name +
          " did not register for the Zero Hunger Match."
      };
    }


    const currentAttendance =
      sheet
        .getRange(
          participantRow,
          COL.MATCH_ATTENDANCE
        )
        .getValue();


    if (
      String(currentAttendance)
        .toUpperCase() === "YES"
    ) {

      const previousTime =
        sheet
          .getRange(
            participantRow,
            COL.MATCH_CHECKIN
          )
          .getValue();


      return {

        success: false,

        duplicate: true,

        participant: {

          name:
            participant.name,

          type:
            participant.type
        },

        previousTime:
          formatDate(previousTime)
      };
    }


    const now = new Date();


    sheet
      .getRange(
        participantRow,
        COL.MATCH_ATTENDANCE
      )
      .setValue("YES");


    sheet
      .getRange(
        participantRow,
        COL.MATCH_CHECKIN
      )
      .setValue(now);


    sheet
      .getRange(
        participantRow,
        COL.MATCH_THANKYOU
      )
      .setValue("PENDING");


    return {

      success: true,

      event:
        "MATCH",

      participant: {

        name:
          participant.name,

        type:
          participant.type,

        registrationId:
          participant.registrationId
      },

      time:
        formatDate(now)
    };
  }


  /***************************************************
   * FOOD DISTRIBUTION
   ***************************************************/

  if (
    eventName ===
    CONFIG.FOOD_EVENT
  ) {

    if (
      !participant.events
        .includes(
          "World Food Day Food Distribution"
        )
    ) {

      return {

        success: false,

        message:
          participant.name +
          " did not register for the World Food Day Food Distribution."
      };
    }


    const currentAttendance =
      sheet
        .getRange(
          participantRow,
          COL.FOOD_ATTENDANCE
        )
        .getValue();


    if (
      String(currentAttendance)
        .toUpperCase() === "YES"
    ) {

      const previousTime =
        sheet
          .getRange(
            participantRow,
            COL.FOOD_CHECKIN
          )
          .getValue();


      return {

        success: false,

        duplicate: true,

        participant: {

          name:
            participant.name,

          type:
            participant.type
        },

        previousTime:
          formatDate(previousTime)
      };
    }


    const now = new Date();


    sheet
      .getRange(
        participantRow,
        COL.FOOD_ATTENDANCE
      )
      .setValue("YES");


    sheet
      .getRange(
        participantRow,
        COL.FOOD_CHECKIN
      )
      .setValue(now);


    sheet
      .getRange(
        participantRow,
        COL.FOOD_THANKYOU
      )
      .setValue("PENDING");


    return {

      success: true,

      event:
        "FOOD",

      participant: {

        name:
          participant.name,

        type:
          participant.type,

        registrationId:
          participant.registrationId
      },

      time:
        formatDate(now)
    };
  }


  return {

    success: false,

    message:
      "Unknown event."
  };
}


function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


function jsonpResponse(
  data,
  callback
) {

  // If no callback is supplied,
  // return normal JSON.
  if (!callback) {

    return jsonResponse(data);
  }


  // Only permit a simple JavaScript
  // function name.
  if (
    !/^[A-Za-z_$][A-Za-z0-9_$]*$/
      .test(callback)
  ) {

    return jsonResponse({

      success: false,

      message:
        "Invalid callback."
    });
  }


  return ContentService

    .createTextOutput(
      callback +
      "(" +
      JSON.stringify(data) +
      ");"
    )

    .setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
}


/*******************************************************
 * CHECK IN PARTICIPANT
 *******************************************************/

function checkInParticipant(
  registrationId,
  eventName
) {

  registrationId =
    String(registrationId).trim();

  if (!registrationId) {

    return {
      success: false,
      message:
        "No registration ID detected."
    };
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  let participant = null;
  let participantRow = null;

  // Find participant
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(
        data[i][COL.REG_ID - 1]
      ).trim() === registrationId
    ) {

      participantRow = i + 1;

      participant = {

        name:
          data[i][COL.NAME - 1],

        email:
          data[i][COL.EMAIL - 1],

        phone:
          data[i][COL.PHONE - 1],

        organisation:
          data[i][COL.ORGANISATION - 1],

        type:
          data[i][COL.TYPE - 1],

        events:
          data[i][COL.EVENTS - 1],

        registrationId:
          registrationId
      };

      break;
    }
  }

  if (!participant) {

    return {
      success: false,
      message:
        "Registration ID not found."
    };
  }


  /*****************************************************
   * MATCH
   *****************************************************/

  if (
    eventName === CONFIG.MATCH_EVENT
  ) {

    // Verify registration
    if (
      !String(participant.events)
        .includes("Zero Hunger Match")
    ) {

      return {
        success: false,
        message:
          participant.name +
          " did not register for the Zero Hunger Match."
      };
    }

    const attendance =
      sheet
        .getRange(
          participantRow,
          COL.MATCH_ATTENDANCE
        )
        .getValue();

    if (
      String(attendance)
        .toUpperCase() === "YES"
    ) {

      const previousTime =
        sheet
          .getRange(
            participantRow,
            COL.MATCH_CHECKIN
          )
          .getValue();

      return {
        success: false,
        duplicate: true,
        participant,
        previousTime:
          formatDate(previousTime)
      };
    }

    sheet
      .getRange(
        participantRow,
        COL.MATCH_ATTENDANCE
      )
      .setValue("YES");

    const now = new Date();

    sheet
      .getRange(
        participantRow,
        COL.MATCH_CHECKIN
      )
      .setValue(now);

    sheet
      .getRange(
        participantRow,
        COL.MATCH_THANKYOU
      )
      .setValue("PENDING");

    return {
      success: true,
      event: "MATCH",
      participant,
      time: formatDate(now)
    };
  }


  /*****************************************************
   * FOOD DISTRIBUTION
   *****************************************************/

  if (
    eventName === CONFIG.FOOD_EVENT
  ) {

    if (
      !String(participant.events)
        .includes("World Food Day Food Distribution")
    ) {

      return {
        success: false,
        message:
          participant.name +
          " did not register for the World Food Day Food Distribution."
      };
    }

    const attendance =
      sheet
        .getRange(
          participantRow,
          COL.FOOD_ATTENDANCE
        )
        .getValue();

    if (
      String(attendance)
        .toUpperCase() === "YES"
    ) {

      const previousTime =
        sheet
          .getRange(
            participantRow,
            COL.FOOD_CHECKIN
          )
          .getValue();

      return {
        success: false,
        duplicate: true,
        participant,
        previousTime:
          formatDate(previousTime)
      };
    }

    sheet
      .getRange(
        participantRow,
        COL.FOOD_ATTENDANCE
      )
      .setValue("YES");

    const now = new Date();

    sheet
      .getRange(
        participantRow,
        COL.FOOD_CHECKIN
      )
      .setValue(now);

    sheet
      .getRange(
        participantRow,
        COL.FOOD_THANKYOU
      )
      .setValue("PENDING");

    return {
      success: true,
      event: "FOOD",
      participant,
      time: formatDate(now)
    };
  }


  return {
    success: false,
    message:
      "Unknown event."
  };
}


/*******************************************************
 * SEND THANK-YOU EMAILS
 *
 * Runs hourly.
 * Sends once 24 hours have passed.
 *******************************************************/

function sendDueThankYouEmails() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const data =
    sheet.getDataRange().getValues();

  const now =
    new Date();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];

    const name =
      row[COL.NAME - 1];

    const email =
      row[COL.EMAIL - 1];

    const participantType =
      row[COL.TYPE - 1];


    /***************************************************
     * MATCH THANK-YOU
     ***************************************************/

    const matchAttendance =
      row[COL.MATCH_ATTENDANCE - 1];

    const matchCheckin =
      row[COL.MATCH_CHECKIN - 1];

    const matchStatus =
      row[COL.MATCH_THANKYOU - 1];


    if (
      String(matchAttendance)
        .toUpperCase() === "YES" &&

      matchCheckin &&

      String(matchStatus)
        .toUpperCase() !== "SENT"
    ) {

      const hours =
        hoursBetween(
          new Date(matchCheckin),
          now
        );

      if (hours >= 24) {

        try {

          sendThankYouEmail({
            name,
            email,
            participantType,
            eventName:
              "Zero Hunger Match",
            eventDate:
              "10 October 2026"
          });

          sheet
            .getRange(
              i + 1,
              COL.MATCH_THANKYOU
            )
            .setValue("SENT");

          sheet
            .getRange(
              i + 1,
              COL.MATCH_SENT
            )
            .setValue(new Date());

        } catch (error) {

          sheet
            .getRange(
              i + 1,
              COL.MATCH_THANKYOU
            )
            .setValue(
              "FAILED"
            );
        }
      }
    }


    /***************************************************
     * FOOD DISTRIBUTION THANK-YOU
     ***************************************************/

    const foodAttendance =
      row[COL.FOOD_ATTENDANCE - 1];

    const foodCheckin =
      row[COL.FOOD_CHECKIN - 1];

    const foodStatus =
      row[COL.FOOD_THANKYOU - 1];


    if (
      String(foodAttendance)
        .toUpperCase() === "YES" &&

      foodCheckin &&

      String(foodStatus)
        .toUpperCase() !== "SENT"
    ) {

      const hours =
        hoursBetween(
          new Date(foodCheckin),
          now
        );

      if (hours >= 24) {

        try {

          sendThankYouEmail({
            name,
            email,
            participantType,
            eventName:
              "World Food Day Food Distribution",
            eventDate:
              "16 October 2026"
          });

          sheet
            .getRange(
              i + 1,
              COL.FOOD_THANKYOU
            )
            .setValue("SENT");

          sheet
            .getRange(
              i + 1,
              COL.FOOD_SENT
            )
            .setValue(new Date());

        } catch (error) {

          sheet
            .getRange(
              i + 1,
              COL.FOOD_THANKYOU
            )
            .setValue(
              "FAILED"
            );
        }
      }
    }
  }
}


/*******************************************************
 * THANK-YOU EMAIL
 *******************************************************/

function sendThankYouEmail(data) {

  const subject =
    "Thank You for Attending — " +
    data.eventName +
    " | FoodClique";

  const html = `

  <div style="
    font-family:Arial,sans-serif;
    max-width:650px;
    margin:auto;
    line-height:1.6;
  ">

    <h2 style="color:#16833b;">
      Thank You for Being Part of It
    </h2>

    <p>
      Dear ${escapeHtml(data.name)},
    </p>

    <p>
      Thank you for attending the
      <strong>
        ${escapeHtml(data.eventName)}
      </strong>
      on ${escapeHtml(data.eventDate)}
      as a
      <strong>
        ${escapeHtml(data.participantType)}
      </strong>.
    </p>

    <p>
      Your presence and support contributed to
      FoodClique Support Initiative's
      World Food Day 2026 campaign,
      themed
      <strong>
        “Together Towards Zero Hunger.”
      </strong>
    </p>

    <p>
      We sincerely appreciate your time,
      energy and commitment towards supporting
      efforts to combat hunger and improve
      food security in vulnerable communities.
    </p>

    <p>
      Thank you for standing with FoodClique.
    </p>

    <p>
      Warm regards,<br>

      <strong>
        FoodClique Support Initiative
      </strong><br>

      info@foodclique.org
    </p>

  </div>

  `;

  GmailApp.sendEmail(
    data.email,
    subject,
    "Thank you for attending " +
      data.eventName,

    {
      htmlBody: html,
      name: CONFIG.EMAIL_NAME
    }
  );
}


/*******************************************************
 * HOURS BETWEEN TWO DATES
 *******************************************************/

function hoursBetween(
  start,
  end
) {

  return (
    end.getTime() -
    start.getTime()
  ) /
  (1000 * 60 * 60);
}


/*******************************************************
 * DATE FORMAT
 *******************************************************/

function formatDate(date) {

  if (!date) {
    return "";
  }

  return Utilities.formatDate(
    new Date(date),
    CONFIG.TIMEZONE,
    "dd MMM yyyy, hh:mm a"
  );
}


/*******************************************************
 * ESCAPE HTML
 *******************************************************/

function escapeHtml(text) {

  if (
    text === null ||
    text === undefined
  ) {
    return "";
  }

  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


/*******************************************************
 * CREATE FORM SUBMISSION + HOURLY TRIGGERS
 *******************************************************/

function createTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    const functionName =
      trigger.getHandlerFunction();

    if (
      functionName === "onFormSubmit" ||
      functionName === "sendDueThankYouEmails"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }
  });


  // Form submission
  ScriptApp
    .newTrigger("onFormSubmit")
    .forSpreadsheet(
      SpreadsheetApp.getActive()
    )
    .onFormSubmit()
    .create();


  // Check hourly for 24-hour emails
  ScriptApp
    .newTrigger(
      "sendDueThankYouEmails"
    )
    .timeBased()
    .everyHours(1)
    .create();
}
