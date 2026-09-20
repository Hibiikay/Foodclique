/***************************************************************
 * FOODCLIQUE SUPPORT INITIATIVE
 * WORLD FOOD DAY 2026 REGISTRATION & ATTENDANCE SYSTEM
 *
 * System:
 * Google Form
 *     ↓
 * Google Sheet
 *     ↓
 * Apps Script
 *     ├── Registration ID
 *     ├── QR Code
 *     ├── Confirmation Email
 *     ├── Calendar Buttons
 *     └── Attendance API
 *             ↓
 *        Netlify Scanner
 *             ↓
 *        Google Sheet
 *             ↓
 *     24-Hour Thank You Email
 *
 * Timezone: Africa/Lagos
 ***************************************************************/


/* ============================================================
   1. CONFIGURATION
   ============================================================ */

const CONFIG = {

  // Google Sheet tab receiving the Google Form responses
  SHEET_NAME: "Form Responses 1",

  // Nigeria timezone
  TIMEZONE: "Africa/Lagos",

  // Organisation
  ORGANIZATION: "FoodClique Support Initiative",

  // Email
  EMAIL_FROM_NAME: "FoodClique Support Initiative",

  // Campaign
  CAMPAIGN_NAME: "World Food Day 2026",
  CAMPAIGN_THEME: "Together Towards Zero Hunger",

  // Google Form event choices
  MATCH_EVENT: "Zero Hunger Match — 10 October 2026",
  FOOD_EVENT: "World Food Day Food Distribution — 16 October 2026",


  /* ----------------------------------------------------------
     CALENDAR DETAILS
     ---------------------------------------------------------- */

  CALENDAR: {

    MATCH: {

      title:
        "FoodClique Zero Hunger Match — World Food Day 2026",

      date:
        "2026-10-10",

      // 11:00 AM - 5:00 PM
      startTime:
        "11:00",

      endTime:
        "17:00",

      venue:
        "Elegbata Sport Complex, Lagos Island, Lagos",

      description:
        "FoodClique Support Initiative's Zero Hunger Match " +
        "as part of the World Food Day 2026 Campaign.\n\n" +

        "Theme: Together Towards Zero Hunger.\n\n" +

        "Join FoodClique Support Initiative, volunteers, " +
        "partners and supporters as we mobilize action " +
        "against hunger and food insecurity."
    },


    FOOD: {

      title:
        "FoodClique World Food Day Cooked Meal Distribution — 2026",

      date:
        "2026-10-16",

      // 11:00 AM - 5:00 PM
      startTime:
        "11:00",

      endTime:
        "17:00",

      venue:
        "Yaba College of Technology, Lagos",

      description:
        "FoodClique Support Initiative's World Food Day 2026 " +
        "Cooked Meal Distribution.\n\n" +

        "Theme: Together Towards Zero Hunger.\n\n" +

        "FoodClique will distribute cooked meals as part of " +
        "its World Food Day campaign to raise awareness " +
        "and mobilize action against hunger and food insecurity."
    }
  },


  /* ----------------------------------------------------------
     COLUMN NUMBERS
     ----------------------------------------------------------

     Google Form columns:

     A Timestamp
     B Full Name
     C Email Address
     D Phone Number
     E Organisation / Institution
     F Participant Type
     G Event(s) Attending
     H How Did You Hear About Us

     Automation columns:

     I Registration ID
     J QR Code
     K Registration Email

     L Match Attendance
     M Match Check-in Time
     N Match Thank-you
     O Match Thank-you Sent

     P Food Distribution Attendance
     Q Food Distribution Check-in Time
     R Food Distribution Thank-you
     S Food Distribution Thank-you Sent
     ---------------------------------------------------------- */

  COL: {

    TIMESTAMP: 1,
    FULL_NAME: 2,
    EMAIL: 3,
    PHONE: 4,
    ORGANIZATION: 5,
    PARTICIPANT_TYPE: 6,
    EVENTS: 7,
    SOURCE: 8,

    REGISTRATION_ID: 9,
    QR_CODE: 10,
    REGISTRATION_EMAIL: 11,

    MATCH_ATTENDANCE: 12,
    MATCH_CHECKIN_TIME: 13,
    MATCH_THANKYOU: 14,
    MATCH_THANKYOU_SENT: 15,

    FOOD_ATTENDANCE: 16,
    FOOD_CHECKIN_TIME: 17,
    FOOD_THANKYOU: 18,
    FOOD_THANKYOU_SENT: 19
  }
};


/* ============================================================
   2. SETUP SYSTEM
   ============================================================ */

/**
 * Run this ONCE after pasting the complete code.
 *
 * It:
 * - Creates the automation headers
 * - Creates the form-submit trigger
 * - Creates the hourly thank-you trigger
 */
function setupSystem() {

  const sheet =
    getRegistrationSheet();

  ensureSheetColumns(sheet);

  createTriggers();

  Logger.log(
    "FoodClique World Food Day system setup completed."
  );
}


/* ============================================================
   3. GET REGISTRATION SHEET
   ============================================================ */

function getRegistrationSheet() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );

  if (!sheet) {

    throw new Error(
      'Sheet "' +
      CONFIG.SHEET_NAME +
      '" was not found.'
    );
  }

  return sheet;
}


/* ============================================================
   4. ENSURE SHEET COLUMNS EXIST
   ============================================================ */

function ensureSheetColumns(sheet) {

  const headers = [

    "Timestamp",
    "Full Name",
    "Email Address",
    "Phone Number",
    "Organisation / Institution",
    "Participant Type",
    "Event(s) Attending",
    "How Did You Hear About Us",

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


  // Make sure the sheet has at least 19 columns
  if (
    sheet.getMaxColumns() <
    headers.length
  ) {

    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length -
      sheet.getMaxColumns()
    );
  }


  // Only set missing headers.
  // Existing data is not overwritten.
  const currentHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .getValues()[0];


  headers.forEach(
    function(header, index) {

      if (!currentHeaders[index]) {

        sheet
          .getRange(
            1,
            index + 1
          )
          .setValue(header);
      }
    }
  );
}


/* ============================================================
   5. CREATE TRIGGERS
   ============================================================ */

function createTriggers() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const triggers =
    ScriptApp.getProjectTriggers();


  // Remove duplicate triggers created by
  // previous setup runs.
  triggers.forEach(
    function(trigger) {

      const handler =
        trigger.getHandlerFunction();

      if (
        handler === "onFormSubmit" ||
        handler === "sendDueThankYouEmails"
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );


  // Form submission trigger
  ScriptApp
    .newTrigger("onFormSubmit")
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();


  // Hourly thank-you trigger
  ScriptApp
    .newTrigger("sendDueThankYouEmails")
    .timeBased()
    .everyHours(1)
    .create();


  Logger.log(
    "Triggers created successfully."
  );
}


/* ============================================================
   6. FORM SUBMISSION
   ============================================================ */

/**
 * Runs automatically when a Google Form response
 * is submitted.
 */
function onFormSubmit(e) {

  if (
    !e ||
    !e.range
  ) {

    throw new Error(
      "onFormSubmit must be triggered by a Google Form submission."
    );
  }


  const sheet =
    e.range.getSheet();

  const rowNumber =
    e.range.getRow();

  processRegistration(
    sheet,
    rowNumber
  );
}


/* ============================================================
   7. PROCESS REGISTRATION
   ============================================================ */

function processRegistration(
  sheet,
  rowNumber
) {

  const row =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        CONFIG.COL.FOOD_THANKYOU_SENT
      )
      .getValues()[0];


  const fullName =
    String(
      row[
        CONFIG.COL.FULL_NAME - 1
      ] || ""
    ).trim();


  const email =
    String(
      row[
        CONFIG.COL.EMAIL - 1
      ] || ""
    ).trim();


  const phone =
    String(
      row[
        CONFIG.COL.PHONE - 1
      ] || ""
    ).trim();


  const organisation =
    String(
      row[
        CONFIG.COL.ORGANIZATION - 1
      ] || ""
    ).trim();


  const participantType =
    String(
      row[
        CONFIG.COL.PARTICIPANT_TYPE - 1
      ] || ""
    ).trim();


  const events =
    String(
      row[
        CONFIG.COL.EVENTS - 1
      ] || ""
    ).trim();


  if (!fullName) {

    throw new Error(
      "Full Name is missing on row " +
      rowNumber
    );
  }


  if (!email) {

    throw new Error(
      "Email Address is missing on row " +
      rowNumber
    );
  }


  /* ----------------------------------------------------------
     CHECK IF REGISTRATION ALREADY EXISTS
     ---------------------------------------------------------- */

  let registrationId =
    String(
      sheet
        .getRange(
          rowNumber,
          CONFIG.COL.REGISTRATION_ID
        )
        .getValue() || ""
    ).trim();


  if (!registrationId) {

    registrationId =
      generateRegistrationId();

    sheet
      .getRange(
        rowNumber,
        CONFIG.COL.REGISTRATION_ID
      )
      .setValue(
        registrationId
      );
  }


  /* ----------------------------------------------------------
     CREATE QR CODE
     ---------------------------------------------------------- */

  const qrUrl =
    generateQrUrl(
      registrationId
    );


  sheet
    .getRange(
      rowNumber,
      CONFIG.COL.QR_CODE
    )
    .setValue(qrUrl);


  /* ----------------------------------------------------------
     STORE REGISTRATION EMAIL
     ---------------------------------------------------------- */

  sheet
    .getRange(
      rowNumber,
      CONFIG.COL.REGISTRATION_EMAIL
    )
    .setValue(email);


  /* ----------------------------------------------------------
     SEND CONFIRMATION EMAIL
     ---------------------------------------------------------- */

  const data = {

    rowNumber:
      rowNumber,

    fullName:
      fullName,

    email:
      email,

    phone:
      phone,

    organisation:
      organisation,

    participantType:
      participantType,

    events:
      events,

    registrationId:
      registrationId,

    qrUrl:
      qrUrl
  };


  sendRegistrationEmail(
    data
  );
}


/* ============================================================
   8. GENERATE REGISTRATION ID
   ============================================================ */

function generateRegistrationId() {

  const timestamp =
    new Date()
      .getTime()
      .toString(36)
      .toUpperCase();


  const random =
    Math.random()
      .toString(36)
      .substring(2, 7)
      .toUpperCase();


  return (
    "FC-WFD-" +
    timestamp +
    "-" +
    random
  );
}


/* ============================================================
   9. GENERATE QR CODE URL
   ============================================================ */

/**
 * The QR code contains ONLY the registration ID.
 *
 * It does not contain:
 * - name
 * - phone
 * - email
 * - organisation
 *
 * This keeps personal information out of the QR code.
 */
function generateQrUrl(
  registrationId
) {

  const encoded =
    encodeURIComponent(
      registrationId
    );


  return (
    "https://quickchart.io/qr" +
    "?size=500" +
    "&margin=2" +
    "&text=" +
    encoded
  );
}


/* ============================================================
   10. REGISTRATION EMAIL
   ============================================================ */

function sendRegistrationEmail(
  data
) {

  const participantName =
    escapeHtml(
      data.fullName
    );


  const participantType =
    escapeHtml(
      data.participantType
    );


  const events =
    escapeHtml(
      data.events
    );


  const registrationId =
    escapeHtml(
      data.registrationId
    );


  let calendarSection = "";


  /* ----------------------------------------------------------
     MATCH CALENDAR
     ---------------------------------------------------------- */

  if (
    eventSelectionContains(
      data.events,
      CONFIG.MATCH_EVENT
    )
  ) {

    const event =
      CONFIG.CALENDAR.MATCH;


    calendarSection += `

      <div style="
        background:#f7f7f7;
        border:1px solid #e5e5e5;
        border-radius:10px;
        padding:20px;
        margin:20px 0;
      ">

        <h3 style="
          margin-top:0;
          color:#198754;
        ">
          ⚽ Zero Hunger Match
        </h3>

        <p>
          <strong>Date:</strong>
          ${formatCalendarDisplayDate(event.date)}
        </p>

        <p>
          <strong>Time:</strong>
          11:00 AM – 5:00 PM
        </p>

        <p>
          <strong>Venue:</strong>
          ${escapeHtml(event.venue)}
        </p>

        <p>
          Add this event to your personal calendar
          so you don't miss it.
        </p>

        ${createCalendarButtons(event)}

      </div>
    `;
  }


  /* ----------------------------------------------------------
     FOOD DISTRIBUTION CALENDAR
     ---------------------------------------------------------- */

  if (
    eventSelectionContains(
      data.events,
      CONFIG.FOOD_EVENT
    )
  ) {

    const event =
      CONFIG.CALENDAR.FOOD;


    calendarSection += `

      <div style="
        background:#f7f7f7;
        border:1px solid #e5e5e5;
        border-radius:10px;
        padding:20px;
        margin:20px 0;
      ">

        <h3 style="
          margin-top:0;
          color:#198754;
        ">
          🍛 World Food Day Cooked Meal Distribution
        </h3>

        <p>
          <strong>Date:</strong>
          ${formatCalendarDisplayDate(event.date)}
        </p>

        <p>
          <strong>Time:</strong>
          11:00 AM – 5:00 PM
        </p>

        <p>
          <strong>Venue:</strong>
          ${escapeHtml(event.venue)}
        </p>

        <p>
          Add this event to your personal calendar
          so you don't miss it.
        </p>

        ${createCalendarButtons(event)}

      </div>
    `;
  }


  /* ----------------------------------------------------------
     EMAIL SUBJECT
     ---------------------------------------------------------- */

  const subject =
    "FoodClique World Food Day 2026 — Registration Confirmed";


  /* ----------------------------------------------------------
     HTML EMAIL
     ---------------------------------------------------------- */

  const htmlBody = `

  <div style="
    font-family:Arial,Helvetica,sans-serif;
    max-width:680px;
    margin:0 auto;
    color:#333;
    background:#ffffff;
  ">

    <div style="
      background:#198754;
      padding:28px 24px;
      text-align:center;
      color:#ffffff;
      border-radius:10px 10px 0 0;
    ">

      <h1 style="
        margin:0;
        font-size:25px;
      ">
        FoodClique Support Initiative
      </h1>

      <p style="
        margin:8px 0 0;
        font-size:17px;
      ">
        World Food Day 2026
      </p>

    </div>


    <div style="
      padding:28px 24px;
    ">

      <h2 style="
        color:#198754;
      ">
        Registration Confirmed
      </h2>


      <p>
        Dear ${participantName},
      </p>


      <p>
        Thank you for registering to participate in the
        <strong>
          FoodClique World Food Day 2026 Campaign
        </strong>.
      </p>


      <div style="
        background:#f1f8f4;
        border-left:4px solid #198754;
        padding:15px;
        margin:20px 0;
      ">

        <strong>
          Campaign Theme:
        </strong>

        <br>

        Together Towards Zero Hunger

      </div>


      <h3>
        Registration Details
      </h3>


      <table style="
        width:100%;
        border-collapse:collapse;
        margin-bottom:20px;
      ">

        <tr>

          <td style="
            padding:8px;
            font-weight:bold;
            border-bottom:1px solid #eee;
          ">
            Registration ID
          </td>

          <td style="
            padding:8px;
            border-bottom:1px solid #eee;
          ">
            ${registrationId}
          </td>

        </tr>


        <tr>

          <td style="
            padding:8px;
            font-weight:bold;
            border-bottom:1px solid #eee;
          ">
            Participant Type
          </td>

          <td style="
            padding:8px;
            border-bottom:1px solid #eee;
          ">
            ${participantType}
          </td>

        </tr>

      </table>


      <h3>
        Registered Event(s)
      </h3>


      <p>
        ${events}
      </p>


      <h3 style="
        margin-top:30px;
        color:#198754;
      ">
        Add Your Event(s) to Calendar
      </h3>


      <p>
        Use the buttons below to save your registered
        event(s) to your calendar.
      </p>


      ${calendarSection}


      <div style="
        text-align:center;
        margin:35px 0;
        padding:25px;
        background:#fafafa;
        border-radius:10px;
      ">

        <h3>
          Your Event QR Code
        </h3>


        <p>
          Please keep this QR code.
          It will be scanned at event check-in.
        </p>


        <img
          src="${data.qrUrl}"
          alt="FoodClique Registration QR Code"
          style="
            width:220px;
            height:220px;
            border:1px solid #ddd;
            padding:10px;
            background:#ffffff;
          "
        />


        <p style="
          font-size:13px;
          color:#666;
          margin-bottom:0;
        ">
          Registration ID:<br>
          <strong>${registrationId}</strong>
        </p>

      </div>


      <div style="
        background:#fff8e1;
        border-radius:8px;
        padding:16px;
        margin:20px 0;
      ">

        <strong>
          Important:
        </strong>

        <p style="
          margin-bottom:0;
        ">
          Please keep this email and your QR code.
          You will need the QR code for event check-in.
        </p>

      </div>


      <p style="
        margin-top:30px;
      ">

        We look forward to having you join us as we work
        <strong>
          Together Towards Zero Hunger.
        </strong>

      </p>


      <p>

        Warm regards,<br>

        <strong>
          FoodClique Support Initiative
        </strong>

      </p>

    </div>


    <div style="
      background:#f5f5f5;
      padding:20px;
      text-align:center;
      font-size:12px;
      color:#666;
      border-radius:0 0 10px 10px;
    ">

      FoodClique Support Initiative<br>

      info@foodclique.org

    </div>

  </div>

  `;


  /* ----------------------------------------------------------
     PLAIN TEXT FALLBACK
     ---------------------------------------------------------- */

  const plainBody = `

FoodClique Support Initiative

WORLD FOOD DAY 2026

Theme:
Together Towards Zero Hunger


Dear ${data.fullName},

Your registration has been confirmed.


Registration ID:
${data.registrationId}


Participant Type:
${data.participantType}


Registered Event(s):
${data.events}


EVENT INFORMATION

Zero Hunger Match:
Saturday, 10 October 2026
11:00 AM - 5:00 PM
Elegbata Sport Complex, Lagos Island, Lagos


World Food Day Cooked Meal Distribution:
Friday, 16 October 2026
11:00 AM - 5:00 PM
Yaba College of Technology, Lagos


Please keep this email and your QR code
for event check-in.


Thank you for supporting FoodClique Support Initiative.


Together Towards Zero Hunger.

  `;


  MailApp.sendEmail({

    to:
      data.email,

    subject:
      subject,

    body:
      plainBody,

    htmlBody:
      htmlBody,

    name:
      CONFIG.EMAIL_FROM_NAME
  });
}


/* ============================================================
   11. EVENT SELECTION CHECK
   ============================================================ */

/**
 * Normalizes event text so differences such as:
 *
 * Zero Hunger Match — 10 October 2026
 *
 * and
 *
 * Zero Hunger Match - 10 October 2026
 *
 * are treated consistently.
 */
function normalizeEventText(
  value
) {

  return String(
    value || ""
  )
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}


/**
 * Determines whether a participant selected an event.
 *
 * This is intentionally tolerant of small differences
 * in Google Form checkbox text.
 */
function eventSelectionContains(
  selectedEvents,
  eventName
) {

  const selected =
    normalizeEventText(
      selectedEvents
    );

  const target =
    normalizeEventText(
      eventName
    );


  // Exact / normal substring match
  if (
    selected.indexOf(target) !== -1
  ) {

    return true;
  }


  /* ----------------------------------------------------------
     ZERO HUNGER MATCH
     ---------------------------------------------------------- */

  if (
    target.indexOf(
      "zero hunger match"
    ) !== -1
  ) {

    return (
      selected.indexOf(
        "zero hunger match"
      ) !== -1 &&

      (
        selected.indexOf(
          "10 october"
        ) !== -1 ||

        selected.indexOf(
          "10 oct"
        ) !== -1
      )
    );
  }


  /* ----------------------------------------------------------
     FOOD DISTRIBUTION
     ---------------------------------------------------------- */

  if (
    target.indexOf(
      "food distribution"
    ) !== -1
  ) {

    return (
      selected.indexOf(
        "food distribution"
      ) !== -1 &&

      (
        selected.indexOf(
          "16 october"
        ) !== -1 ||

        selected.indexOf(
          "16 oct"
        ) !== -1
      )
    );
  }


  return false;
}


/* ============================================================
   12. GOOGLE CALENDAR URL
   ============================================================ */

function createGoogleCalendarUrl(
  event
) {

  const start =
    calendarDateTime(
      event.date,
      event.startTime
    );


  const end =
    calendarDateTime(
      event.date,
      event.endTime
    );


  const params = [

    "action=TEMPLATE",

    "text=" +
      encodeURIComponent(
        event.title
      ),

    "dates=" +
      start +
      "/" +
      end,

    "details=" +
      encodeURIComponent(
        event.description
      ),

    "location=" +
      encodeURIComponent(
        event.venue
      ),

    "ctz=Africa/Lagos"

  ].join("&");


  return (
    "https://calendar.google.com/calendar/render?" +
    params
  );
}


/* ============================================================
   13. OUTLOOK CALENDAR URL
   ============================================================ */

function createOutlookCalendarUrl(
  event
) {

  const start =
    event.date +
    "T" +
    event.startTime +
    ":00";


  const end =
    event.date +
    "T" +
    event.endTime +
    ":00";


  const params = [

    "rru=addevent",

    "startdt=" +
      encodeURIComponent(
        start
      ),

    "enddt=" +
      encodeURIComponent(
        end
      ),

    "subject=" +
      encodeURIComponent(
        event.title
      ),

    "body=" +
      encodeURIComponent(
        event.description
      ),

    "location=" +
      encodeURIComponent(
        event.venue
      )

  ].join("&");


  return (
    "https://outlook.live.com/calendar/0/deeplink/compose?" +
    params
  );
}


/* ============================================================
   14. CALENDAR BUTTON HTML
   ============================================================ */

function createCalendarButtons(
  event
) {

  const googleUrl =
    createGoogleCalendarUrl(
      event
    );


  const outlookUrl =
    createOutlookCalendarUrl(
      event
    );


  return `

    <div style="
      margin:20px 0 5px;
    ">

      <a
        href="${googleUrl}"
        target="_blank"
        style="
          display:inline-block;
          background:#1a73e8;
          color:#ffffff;
          text-decoration:none;
          padding:12px 18px;
          border-radius:6px;
          font-weight:bold;
          margin:5px 5px 5px 0;
        "
      >
        📅 Add to Google Calendar
      </a>


      <a
        href="${outlookUrl}"
        target="_blank"
        style="
          display:inline-block;
          background:#0078d4;
          color:#ffffff;
          text-decoration:none;
          padding:12px 18px;
          border-radius:6px;
          font-weight:bold;
          margin:5px;
        "
      >
        📅 Add to Outlook Calendar
      </a>

    </div>

  `;
}


/* ============================================================
   15. CALENDAR DATE/TIME
   ============================================================ */

function calendarDateTime(
  dateString,
  timeString
) {

  return (

    dateString.replace(
      /-/g,
      ""
    ) +

    "T" +

    timeString.replace(
      ":",
      ""
    ) +

    "00"

  );
}


/* ============================================================
   16. DISPLAY DATE
   ============================================================ */

function formatCalendarDisplayDate(
  dateString
) {

  const date =
    new Date(
      dateString +
      "T00:00:00"
    );


  return Utilities.formatDate(

    date,

    CONFIG.TIMEZONE,

    "EEEE, d MMMM yyyy"

  );
}


/* ============================================================
   17. WEB APP / NETLIFY API
   ============================================================ */

/**
 * This endpoint is used by the Netlify QR scanner.
 *
 * Supported:
 *
 * ?api=test
 *
 * ?api=checkin
 * &registrationId=FC-WFD-XXXX
 * &event=match
 *
 * or:
 *
 * ?api=checkin
 * &registrationId=FC-WFD-XXXX
 * &event=food
 */
function doGet(e) {

  const params =
    (
      e &&
      e.parameter
    )
      ? e.parameter
      : {};


  /* ----------------------------------------------------------
     API TEST
     ---------------------------------------------------------- */

  if (
    params.api === "test"
  ) {

    const response = {

      success:
        true,

      message:
        "FoodClique World Food Day 2026 API is working.",

      timestamp:
        new Date().toISOString()
    };


    return jsonpResponse(
      response,
      params.callback
    );
  }


  /* ----------------------------------------------------------
     CHECK-IN
     ---------------------------------------------------------- */

  if (
    params.api === "checkin"
  ) {

    return handleCheckInApi(
      params
    );
  }


  /* ----------------------------------------------------------
     DEFAULT
     ---------------------------------------------------------- */

  return jsonResponse({

    success:
      true,

    service:
      "FoodClique World Food Day 2026 Attendance API",

    status:
      "online"

  });
}


/* ============================================================
   /* ============================================================
   18. CHECK-IN API
   ============================================================ */

/**
 * Converts the event received from the Netlify scanner
 * into the internal event name used by the system.
 *
 * Accepted values:
 *   match
 *   food
 *   Zero Hunger Match
 *   Food Distribution
 *   Cooked Meal Distribution
 */
function normalizeApiEvent(value) {

  const event =
    String(value || "")
      .toLowerCase()
      .replace(/[–—]/g, "-")
      .replace(/\s+/g, " ")
      .trim();


  if (
    event === "match" ||
    event.indexOf("zero hunger match") !== -1
  ) {
    return "match";
  }


  if (
    event === "food" ||
    event.indexOf("food distribution") !== -1 ||
    event.indexOf("cooked meal") !== -1
  ) {
    return "food";
  }


  return "";
}


/**
 * Handles a check-in request from the Netlify scanner.
 */
function handleCheckInApi(params) {

  const registrationId =
    String(
      params.registrationId || ""
    ).trim();


  const event =
    normalizeApiEvent(
      params.event
    );


  /* ----------------------------------------------------------
     VALIDATE REGISTRATION ID
     ---------------------------------------------------------- */

  if (!registrationId) {

    return apiResponse(

      {
        success: false,
        error: "Registration ID is required."
      },

      params.callback
    );
  }


  /* ----------------------------------------------------------
     VALIDATE EVENT
     ---------------------------------------------------------- */

  if (
    event !== "match" &&
    event !== "food"
  ) {

    return apiResponse(

      {
        success: false,
        error:
          "Invalid event. Use match or food."
      },

      params.callback
    );
  }


  /* ----------------------------------------------------------
     CHECK PARTICIPANT
     ---------------------------------------------------------- */

  try {

    const result =
      checkInParticipantApi(
        registrationId,
        event
      );


    return apiResponse(
      result,
      params.callback
    );


  } catch (error) {

    return apiResponse(

      {
        success: false,
        error:
          error.message
      },

      params.callback
    );
  }
}


/* ============================================================
   19. CHECK-IN PARTICIPANT
   ============================================================ */


/* ============================================================
   19. CHECK-IN API
   ============================================================ */

function handleCheckInApi(
  params
) {

  const registrationId =
    String(
      params.registrationId ||
      ""
    ).trim();


  const event =
    normalizeApiEvent(
      params.event
    );


  /* ----------------------------------------------------------
     VALIDATE REGISTRATION ID
     ---------------------------------------------------------- */

  if (!registrationId) {

    return apiResponse(

      {
        success:
          false,

        error:
          "Registration ID is required."
      },

      params.callback
    );
  }


  /* ----------------------------------------------------------
     VALIDATE EVENT
     ---------------------------------------------------------- */

  if (
    event !== "match" &&
    event !== "food"
  ) {

    return apiResponse(

      {
        success:
          false,

        error:
          "Invalid event. Use match or food."
      },

      params.callback
    );
  }


  /* ----------------------------------------------------------
     PROCESS CHECK-IN
     ---------------------------------------------------------- */

  try {

    const result =
      checkInParticipantApi(
        registrationId,
        event
      );


    return apiResponse(
      result,
      params.callback
    );

  } catch (error) {

    return apiResponse(

      {
        success:
          false,

        error:
          error.message
      },

      params.callback
    );
  }
}


/* ============================================================
   20. CHECK-IN PARTICIPANT
   ============================================================ */

function checkInParticipantApi(
  registrationId,
  event
) {

  const sheet =
    getRegistrationSheet();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return {

      success:
        false,

      error:
        "No registrations found."
    };
  }


  /* ----------------------------------------------------------
     SEARCH REGISTRATION ID
     ---------------------------------------------------------- */

  const idColumn =
    sheet.getRange(

      2,

      CONFIG.COL.REGISTRATION_ID,

      lastRow - 1,

      1

    ).getValues();


  let matchedRow =
    -1;


  for (
    let i = 0;
    i < idColumn.length;
    i++
  ) {

    if (

      String(
        idColumn[i][0]
      )
        .trim()
        .toUpperCase() ===

      String(
        registrationId
      )
        .trim()
        .toUpperCase()

    ) {

      matchedRow =
        i + 2;

      break;
    }
  }


  if (
    matchedRow === -1
  ) {

    return {

      success:
        false,

      error:
        "Registration ID not found."
    };
  }


  /* ----------------------------------------------------------
     GET PARTICIPANT
     ---------------------------------------------------------- */

  const row =
    sheet
      .getRange(
        matchedRow,
        1,
        1,
        CONFIG.COL.FOOD_THANKYOU_SENT
      )
      .getValues()[0];


  const participantName =
    String(
      row[
        CONFIG.COL.FULL_NAME - 1
      ] || ""
    ).trim();


  const participantType =
    String(
      row[
        CONFIG.COL.PARTICIPANT_TYPE - 1
      ] || ""
    ).trim();


  const registeredEvents =
    String(
      row[
        CONFIG.COL.EVENTS - 1
      ] || ""
    ).trim();


  /* ----------------------------------------------------------
     MAKE SURE PARTICIPANT REGISTERED FOR EVENT
     ---------------------------------------------------------- */

  let registeredForEvent =
    false;


  if (
    event === "match"
  ) {

    registeredForEvent =
      eventSelectionContains(
        registeredEvents,
        CONFIG.MATCH_EVENT
      );

  } else if (
    event === "food"
  ) {

    registeredForEvent =
      eventSelectionContains(
        registeredEvents,
        CONFIG.FOOD_EVENT
      );
  }


  if (
    !registeredForEvent
  ) {

    return {

      success:
        false,

      error:
        "This participant did not register for this event.",

      participantName:
        participantName
    };
  }


  /* ----------------------------------------------------------
     EVENT COLUMN SELECTION
     ---------------------------------------------------------- */

  let attendanceColumn;
  let checkinColumn;
  let thankYouColumn;
  let thankYouSentColumn;


  if (
    event === "match"
  ) {

    attendanceColumn =
      CONFIG.COL.MATCH_ATTENDANCE;

    checkinColumn =
      CONFIG.COL.MATCH_CHECKIN_TIME;

    thankYouColumn =
      CONFIG.COL.MATCH_THANKYOU;

    thankYouSentColumn =
      CONFIG.COL.MATCH_THANKYOU_SENT;

  } else {

    attendanceColumn =
      CONFIG.COL.FOOD_ATTENDANCE;

    checkinColumn =
      CONFIG.COL.FOOD_CHECKIN_TIME;

    thankYouColumn =
      CONFIG.COL.FOOD_THANKYOU;

    thankYouSentColumn =
      CONFIG.COL.FOOD_THANKYOU_SENT;
  }


  /* ----------------------------------------------------------
     CHECK DUPLICATE
     ---------------------------------------------------------- */

  const existingAttendance =
    String(
      sheet
        .getRange(
          matchedRow,
          attendanceColumn
        )
        .getValue() || ""
    ).trim();


  const existingCheckinTime =
    sheet
      .getRange(
        matchedRow,
        checkinColumn
      )
      .getValue();


  if (
    existingAttendance ===
    "Present"
  ) {

    return {

      success:
        true,

      duplicate:
        true,

      message:
        "Participant has already been checked in.",

      participantName:
        participantName,

      participantType:
        participantType,

      registrationId:
        registrationId,

      event:
        event,

      checkInTime:
        formatDate(
          existingCheckinTime
        )
    };
  }


  /* ----------------------------------------------------------
     RECORD ATTENDANCE
     ---------------------------------------------------------- */

  const now =
    new Date();


  sheet
    .getRange(
      matchedRow,
      attendanceColumn
    )
    .setValue(
      "Present"
    );


  sheet
    .getRange(
      matchedRow,
      checkinColumn
    )
    .setValue(
      now
    );


  /* ----------------------------------------------------------
     MARK THANK-YOU AS PENDING
     ---------------------------------------------------------- */

  sheet
    .getRange(
      matchedRow,
      thankYouColumn
    )
    .setValue(
      "Pending"
    );


  sheet
    .getRange(
      matchedRow,
      thankYouSentColumn
    )
    .setValue(
      "No"
    );


  /* ----------------------------------------------------------
     RETURN RESULT
     ---------------------------------------------------------- */

  return {

    success:
      true,

    duplicate:
      false,

    message:
      "Check-in successful.",

    participantName:
      participantName,

    participantType:
      participantType,

    registrationId:
      registrationId,

    event:
      event,

    checkInTime:
      formatDate(
        now
      )
  };
}


/* ============================================================
   21. API RESPONSE
   ============================================================ */

function apiResponse(
  data,
  callback
) {

  if (callback) {

    return jsonpResponse(
      data,
      callback
    );
  }


  return jsonResponse(
    data
  );
}


/* ============================================================
   22. JSON RESPONSE
   ============================================================ */

function jsonResponse(
  data
) {

  return ContentService

    .createTextOutput(
      JSON.stringify(
        data
      )
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/* ============================================================
   23. JSONP RESPONSE
   ============================================================ */

function jsonpResponse(
  data,
  callback
) {

  /*
   * Only allow normal JavaScript function names.
   * This prevents arbitrary JavaScript being injected
   * as the callback.
   */

  const safeCallback =
    /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/
      .test(
        callback || ""
      )
      ? callback
      : null;


  if (
    !safeCallback
  ) {

    return jsonResponse(
      data
    );
  }


  return ContentService

    .createTextOutput(

      safeCallback +
      "(" +
      JSON.stringify(
        data
      ) +
      ")"

    )

    .setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
}


/* ============================================================
   24. 24-HOUR THANK-YOU AUTOMATION
   ============================================================ */

/**
 * Runs every hour.
 *
 * It checks whether an attendee has been checked in
 * for at least 24 hours.
 *
 * Each event is processed independently.
 *
 * Therefore:
 *
 * Person attends Match
 *       ↓
 * Match thank-you after 24 hours
 *
 * Person attends Food Distribution
 *       ↓
 * Food Distribution thank-you after 24 hours
 *
 * Person attends BOTH
 *       ↓
 * Two separate thank-you emails
 */
function sendDueThankYouEmails() {

  const sheet =
    getRegistrationSheet();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return;
  }


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        CONFIG.COL.FOOD_THANKYOU_SENT
      )
      .getValues();


  const now =
    new Date();


  data.forEach(
    function(row, index) {

      const rowNumber =
        index + 2;


      const email =
        String(
          row[
            CONFIG.COL.EMAIL - 1
          ] || ""
        ).trim();


      const fullName =
        String(
          row[
            CONFIG.COL.FULL_NAME - 1
          ] || ""
        ).trim();


      const registrationId =
        String(
          row[
            CONFIG.COL.REGISTRATION_ID - 1
          ] || ""
        ).trim();


      if (
        !email ||
        !registrationId
      ) {

        return;
      }


      /* ------------------------------------------------------
         MATCH THANK-YOU
         ------------------------------------------------------ */

      processThankYouForEvent({

        sheet:
          sheet,

        rowNumber:
          rowNumber,

        row:
          row,

        event:
          "match",

        email:
          email,

        fullName:
          fullName,

        registrationId:
          registrationId,

        now:
          now
      });


      /* ------------------------------------------------------
         FOOD DISTRIBUTION THANK-YOU
         ------------------------------------------------------ */

      processThankYouForEvent({

        sheet:
          sheet,

        rowNumber:
          rowNumber,

        row:
          row,

        event:
          "food",

        email:
          email,

        fullName:
          fullName,

        registrationId:
          registrationId,

        now:
          now
      });

    }
  );
}


/* ============================================================
   25. PROCESS ONE EVENT THANK-YOU
   ============================================================ */

function processThankYouForEvent(
  data
) {

  const sheet =
    data.sheet;


  let attendanceColumn;
  let checkinColumn;
  let thankYouColumn;
  let thankYouSentColumn;


  if (
    data.event === "match"
  ) {

    attendanceColumn =
      CONFIG.COL.MATCH_ATTENDANCE;

    checkinColumn =
      CONFIG.COL.MATCH_CHECKIN_TIME;

    thankYouColumn =
      CONFIG.COL.MATCH_THANKYOU;

    thankYouSentColumn =
      CONFIG.COL.MATCH_THANKYOU_SENT;

  } else {

    attendanceColumn =
      CONFIG.COL.FOOD_ATTENDANCE;

    checkinColumn =
      CONFIG.COL.FOOD_CHECKIN_TIME;

    thankYouColumn =
      CONFIG.COL.FOOD_THANKYOU;

    thankYouSentColumn =
      CONFIG.COL.FOOD_THANKYOU_SENT;
  }


  const attendance =
    String(
      data.row[
        attendanceColumn - 1
      ] || ""
    ).trim();


  const checkinTime =
    data.row[
      checkinColumn - 1
    ];


  const thankYouSent =
    String(
      data.row[
        thankYouSentColumn - 1
      ] || ""
    )
      .trim()
      .toLowerCase();


  /* ----------------------------------------------------------
     ONLY PROCESS PRESENT PARTICIPANTS
     ---------------------------------------------------------- */

  if (
    attendance !== "Present"
  ) {

    return;
  }


  /* ----------------------------------------------------------
     DON'T SEND TWICE
     ---------------------------------------------------------- */

  if (
    thankYouSent === "yes"
  ) {

    return;
  }


  /* ----------------------------------------------------------
     VALID CHECK-IN TIME
     ---------------------------------------------------------- */

  if (
    !checkinTime ||
    !(checkinTime instanceof Date)
  ) {

    return;
  }


  /* ----------------------------------------------------------
     WAIT AT LEAST 24 HOURS
     ---------------------------------------------------------- */

  const hours =
    hoursBetween(
      checkinTime,
      data.now
    );


  if (
    hours < 24
  ) {

    return;
  }


  /* ----------------------------------------------------------
     SEND THANK-YOU
     ---------------------------------------------------------- */

  sendThankYouEmail({

    email:
      data.email,

    fullName:
      data.fullName,

    registrationId:
      data.registrationId,

    event:
      data.event
  });


  /* ----------------------------------------------------------
     MARK SENT
     ---------------------------------------------------------- */

  sheet
    .getRange(
      data.rowNumber,
      thankYouColumn
    )
    .setValue(
      "Sent"
    );


  sheet
    .getRange(
      data.rowNumber,
      thankYouSentColumn
    )
    .setValue(
      "Yes"
    );
}


/* ============================================================
   26. SEND THANK-YOU EMAIL
   ============================================================ */

function sendThankYouEmail(
  data
) {

  let eventTitle;
  let eventDate;
  let eventVenue;


  if (
    data.event === "match"
  ) {

    eventTitle =
      "Zero Hunger Match";

    eventDate =
      "Saturday, 10 October 2026";

    eventVenue =
      CONFIG.CALENDAR.MATCH.venue;

  } else {

    eventTitle =
      "World Food Day Cooked Meal Distribution";

    eventDate =
      "Friday, 16 October 2026";

    eventVenue =
      CONFIG.CALENDAR.FOOD.venue;
  }


  const subject =
    "Thank You for Joining FoodClique — " +
    eventTitle;


  const htmlBody = `

    <div style="
      font-family:Arial,Helvetica,sans-serif;
      max-width:650px;
      margin:0 auto;
      color:#333;
    ">


      <div style="
        background:#198754;
        color:#fff;
        padding:25px;
        text-align:center;
        border-radius:8px 8px 0 0;
      ">

        <h1 style="margin:0;">
          Thank You!
        </h1>

        <p style="margin-bottom:0;">
          FoodClique Support Initiative
        </p>

      </div>


      <div style="
        padding:25px;
      ">

        <p>
          Dear ${escapeHtml(data.fullName)},
        </p>


        <p>

          Thank you for joining us for
          <strong>
            ${escapeHtml(eventTitle)}
          </strong>
          as part of the FoodClique World Food Day
          2026 Campaign.

        </p>


        <div style="
          background:#f1f8f4;
          border-left:4px solid #198754;
          padding:15px;
          margin:20px 0;
        ">

          <strong>
            Together Towards Zero Hunger
          </strong>

        </div>


        <p>

          Your participation, support and commitment
          contributed to our collective effort to raise
          awareness and take action against hunger
          and food insecurity.

        </p>


        <p>

          <strong>
            Event:
          </strong>

          ${escapeHtml(eventTitle)}

          <br>

          <strong>
            Date:
          </strong>

          ${escapeHtml(eventDate)}

          <br>

          <strong>
            Venue:
          </strong>

          ${escapeHtml(eventVenue)}

        </p>


        <p>

          We sincerely appreciate your support and hope
          to welcome you again at future FoodClique
          programmes.

        </p>


        <p>

          Warm regards,<br>

          <strong>
            FoodClique Support Initiative
          </strong>

        </p>

      </div>


      <div style="
        background:#f5f5f5;
        padding:18px;
        text-align:center;
        font-size:12px;
        color:#666;
        border-radius:0 0 8px 8px;
      ">

        FoodClique Support Initiative<br>
        info@foodclique.org

      </div>

    </div>

  `;


  const plainBody = `

FoodClique Support Initiative

Thank You for Joining Us


Dear ${data.fullName},

Thank you for joining us for ${eventTitle}
as part of the FoodClique World Food Day
2026 Campaign.


Theme:
Together Towards Zero Hunger


Event:
${eventTitle}

Date:
${eventDate}

Venue:
${eventVenue}


Your participation and support contributed
to our collective effort to raise awareness
and take action against hunger and food insecurity.


We sincerely appreciate your support.


Warm regards,

FoodClique Support Initiative

  `;


  MailApp.sendEmail({

    to:
      data.email,

    subject:
      subject,

    body:
      plainBody,

    htmlBody:
      htmlBody,

    name:
      CONFIG.EMAIL_FROM_NAME
  });
}


/* ============================================================
   27. TIME DIFFERENCE
   ============================================================ */

function hoursBetween(
  startDate,
  endDate
) {

  const milliseconds =
    endDate.getTime() -
    startDate.getTime();


  return (
    milliseconds /
    (1000 * 60 * 60)
  );
}


/* ============================================================
   28. DATE FORMAT
   ============================================================ */

function formatDate(
  date
) {

  if (
    !date ||
    !(date instanceof Date)
  ) {

    return "";
  }


  return Utilities.formatDate(

    date,

    CONFIG.TIMEZONE,

    "d MMMM yyyy, h:mm a"

  );
}


/* ============================================================
   29. HTML ESCAPE
   ============================================================ */

function escapeHtml(
  value
) {

  return String(
    value || ""
  )

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


/* ============================================================
   30. TEST REGISTRATION EMAIL
   ============================================================ */

/**
 * OPTIONAL TEST FUNCTION
 *
 * Change TEST_EMAIL to your own email before running.
 *
 * This does NOT create a real registration in the sheet.
 */
function testRegistrationEmail() {

  const TEST_EMAIL =
    "olarinde.ramat@gmail.com";


  const data = {

    fullName:
      "FoodClique Test Participant",

    email:
      TEST_EMAIL,

    phone:
      "08000000000",

    organisation:
      "FoodClique",

    participantType:
      "Supporter",

    events:
      CONFIG.MATCH_EVENT +
      ", " +
      CONFIG.FOOD_EVENT,

    registrationId:
      "FC-WFD-TEST-001",

    qrUrl:
      generateQrUrl(
        "FC-WFD-TEST-001"
      )
  };


  sendRegistrationEmail(
    data
  );


  Logger.log(
    "Test registration email sent."
  );
}


/* ============================================================
   31. TEST THANK-YOU EMAIL
   ============================================================ */

/**
 * OPTIONAL TEST FUNCTION.
 *
 * Change TEST_EMAIL before running.
 */
function testThankYouEmail() {

  const TEST_EMAIL =
    "olarinde.ramat@gmail.com";


  sendThankYouEmail({

    email:
      TEST_EMAIL,

    fullName:
      "FoodClique Test Participant",

    registrationId:
      "FC-WFD-TEST-001",

    event:
      "match"
  });


  Logger.log(
    "Test thank-you email sent."
  );
}


/* ============================================================
   32. API TEST FUNCTION
   ============================================================ */

/**
 * Optional Apps Script-side test.
 */
function testApiConfiguration() {

  Logger.log(
    "FoodClique WFD API configuration is ready."
  );


  Logger.log(
    "Web app endpoint should use /exec."
  );
}
