function sendThankYouEmail(e) {
  var responses = e.namedValues;
  if (!responses) return;

  // IMPROVED MATCHER: Strips all hidden spaces and special characters
  function getResponseValue(targetName) {
    var keys = Object.keys(responses);
    for (var i = 0; i < keys.length; i++) {
      // This regex removes everything except letters and numbers for a perfect match
      var cleanKey = keys[i].replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      var cleanTarget = targetName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      
      if (cleanKey === cleanTarget) {
        return responses[keys[i]][0];
      }
    }
    return null;
  }

  // Get values using the robust matcher
  var name = getResponseValue("Full Name") || "Volunteer";
  var email = getResponseValue("Email Address");
  var program = getResponseValue("Program / Event Name") || "our program";

  // Stop if no email was found
  if (!email) {
    console.error("FAILED: Could not find Email Address. Columns found: " + Object.keys(responses).join(", "));
    return;
  }

  var brandColor = "#1B8E3E"; // FoodClique Green
  var subject = "Thank You for Volunteering with FoodClique ❤️";

  var htmlBody = `
  <div style="font-family: Arial, sans-serif; background-color:#f9f9f9; padding:20px;">
    <div style="max-width:600px; margin:auto; background:#ffffff; border-radius:8px; overflow:hidden; border: 1px solid #ddd;">
      
      <div style="background:${brandColor}; padding:20px; text-align:center; color:white;">
        <h2 style="margin:0;">FoodClique Support Initiative</h2>
      </div>

      <div style="padding:30px; color:#333; line-height: 1.6;">
        <p>Dear <strong>${name}</strong>,</p>

        <p>
          Thank you for volunteering with <strong>FoodClique Support Initiative</strong> during our 
          <strong>${program}</strong>.
        </p>

        <p>
          Your time, energy, and commitment help us fight hunger and support communities in need.
          We truly appreciate your impact and dedication.
        </p>

        <p>
          Together, we are making food accessible and restoring hope — one meal at a time.
        </p>

        <p style="margin-top:30px;">
          Warm regards,<br>
          <strong>FoodClique Support Initiative Team</strong>
        </p>
      </div>

      <div style="background:#f1f1f1; padding:15px; text-align:center; font-size:12px; color:#666;">
        © ${new Date().getFullYear()} FoodClique Support Initiative <br>
        <em>Fighting Hunger, Restoring Hope.</em>
      </div>

    </div>
  </div>
  `;

  try {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "FoodClique Support Initiative"
    });
    console.log("Email successfully sent to " + email);
  } catch (err) {
    console.error("Gmail Error: " + err.toString());
  }
}

// --- NEW FUNCTION: RUNS EVERY MORNING ---
function sendBirthdayEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indexes (adjust these if your column names are different)
  var nameIdx = headers.indexOf("Full Name");
  var emailIdx = headers.indexOf("Email Address");
  var dobIdx = headers.indexOf("Date of Birth");

  var today = new Date();
  var todayMonthDay = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM-dd");

  // Loop through rows (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var bday = new Date(row[dobIdx]);
    
    // Check if bday is valid and matches today's Month and Day
    if (bday instanceof Date && !isNaN(bday)) {
      var bdayMonthDay = Utilities.formatDate(bday, Session.getScriptTimeZone(), "MM-dd");
      
      if (todayMonthDay === bdayMonthDay) {
        var name = row[nameIdx];
        var email = row[emailIdx];
        
        sendCelebratoryEmail(email, name);
      }
    }
  }
}

// Helper to send the actual Birthday Email
function sendCelebratoryEmail(email, name) {
  var subject = "Happy Birthday from Foodclique Support Initiative!";
  var brandColor = "#1B8E3E";
  
  // Using HTML entities (e.g., &#127874; for cake) ensures they show up correctly
  var htmlBody = `
  <div style="font-family: Arial, sans-serif; background-color:#f9f9f9; padding:20px;">
    <div style="max-width:600px; margin:auto; background:#ffffff; border-radius:8px; overflow:hidden; border: 1px solid #ddd;">
      <div style="background:${brandColor}; padding:20px; text-align:center; color:white;">
        <h1 style="margin:0;">Happy Birthday, ${name}! &#127880;</h1>
      </div>
      <div style="padding:30px; color:#333; line-height: 1.6; text-align: center;">
        <p style="font-size: 18px;">Today, we celebrate <strong>YOU</strong>!</p>
        <p>On behalf of the entire <strong>Foodclique Support Initiative</strong> team, we wish you a day filled with joy, laughter, and everything you love.</p>
        <p>Thank you for being a vital part of our mission to fight hunger. Your heart for service makes the world a better place.</p>
        <p style="font-size: 30px;">&#127874; &#127873; &#127881;</p>
        <p style="margin-top:30px;">Warmest wishes,<br><strong>Foodclique Support Initiative Team</strong></p>
      </div>
    </div>
  </div>`;

  try {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "FoodClique Support Initiative"
    });
    console.log("Birthday email sent successfully to: " + email);
  } catch (e) {
    console.error("Error sending birthday mail: " + e.message);
  }
}

  
function testBirthdayEmail() {
  // Replace this with your email to see the result
  var testEmail = "olarinde.ramat@gmail.com"; 
  var testName = "Ramat";
  
  console.log("Running birthday test...");
  sendCelebratoryEmail(testEmail, testName);
  console.log("Test email sent! Check your inbox.");
}
