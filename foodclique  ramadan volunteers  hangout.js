function onFormSubmit(e) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

var row = sheet.getLastRow();

var data = e.values;

var fullName = data[1];
var phone = data[2];
var email = data[3];
var tshirtSize = data[6];


  // Generate Volunteer ID
  var volunteerId = "RVH2026-" + Utilities.getUuid().substring(0,8).toUpperCase();

  // Event Details
  var eventTitle = "FoodClique Ramadan Volunteers Hangout 2026";

  var venue = "FoodClique Office, Olusola Olude Close, Gbagada Phase 2, Lagos";

  var eventDate = new Date("2026-06-20T13:00:00"); // Adjust date
  function formatDate(date) {
  var options = { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' };
  return date.toLocaleDateString('en-GB', options);
  }

  var readableDate = formatDate(eventDate);

  // Create Calendar Event
  var calendarLink =
  "https://calendar.google.com/calendar/render?action=TEMPLATE" +
  "&text=" + encodeURIComponent("FoodClique Ramadan Volunteers Hangout 2026") +
  "&location=" + encodeURIComponent("FoodClique Office, 6b Olusola Olude Close, Gbagada Phase 2, Lagos") +
  "&dates=20260620T130000/20260620T180000" +
  "&details=" + encodeURIComponent("Ramadan Volunteers Hangout 2026");

  // QR Data
  var qrData =
    "Volunteer ID: " + volunteerId +
    "\nName: " + fullName +
    "\nTshirt Size: " + tshirtSize +
    "\nEvent: Ramadan Volunteers Hangout 2026";

  var qrCode =
    "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
    encodeURIComponent(qrData);

  // Save Details
  sheet.getRange(row,14).setValue(volunteerId);
  sheet.getRange(row,15).setValue(qrCode);
  sheet.getRange(row,16).setValue(calendarLink);

  // Email Design
  var htmlBody = `
  <div style="font-family:Arial,sans-serif;background:#ffffff;padding:20px">

    <div style="
      background:#008751;
      color:white;
      padding:25px;
      text-align:center;
      border-radius:10px 10px 0 0;">

      <h1>FoodClique Support Initiative</h1>
      <h2>Ramadan Volunteers Hangout 2026</h2>
    </div>

    <div style="
      border-left:5px solid #008751;
      border-right:5px solid #008751;
      padding:25px;
      background:#fff;">

      <p>Dear <strong>${fullName}</strong>,</p>

      <p>
      Thank you for registering for the Ramadan Volunteers Hangout 2026.
      We are excited to celebrate and appreciate your contributions during Ramadan 2026.
      </p>

      <table style="border-collapse:collapse;width:100%;">
        <tr>
          <td><strong>Volunteer ID</strong></td>
          <td>${volunteerId}</td>
        </tr>
        <tr>
          <td><strong>T-Shirt Size</strong></td>
          <td>${tshirtSize}</td>
        </tr>
        <tr>
          <td><strong>Venue</strong></td>
          <td>${venue}</td>
        </tr>
        <tr>
          <td><strong>Date</strong></td>
          <td>${readableDate}</td>
        </tr>
        <tr>
          <td><strong>Time</strong></td>
          <td>1:00 PM</td>
        </tr>
      </table>

      <br>

      <h3 style="color:#008751;">Admission QR Code</h3>

      <p>
      Kindly present this QR Code at the entrance for admission.
      </p>

      <img src="${qrCode}" width="250">

      <br><br>

      <a href="${calendarLink}"
      style="
      background:#008751;
      color:white;
      padding:12px 20px;
      text-decoration:none;
      border-radius:5px;">
      Add to Calendar
      </a>

      <br><br>

      <p>
      We look forward to seeing you at the event.
      </p>

      <p>
      Warm regards,
      </p>

      <p>
      <strong>FoodClique Support Initiative</strong><br>
      Fighting Hunger, Spreading Hope
      </p>

    </div>

    <div style="
    background:#E31B23;
    color:white;
    text-align:center;
    padding:12px;
    border-radius:0 0 10px 10px;">
      Ramadan Volunteers Hangout 2026
    </div>

  </div>
  `;
  if (!email) {
  throw new Error("Email address is blank");
}


  MailApp.sendEmail({
    to: email,
    subject: "Ramadan Volunteers Hangout 2026 Registration Confirmation",
    htmlBody: htmlBody
  });

  sheet.getRange(row,17).setValue("Sent");
}
 
