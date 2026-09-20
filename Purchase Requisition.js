/**
 * FOODCLIQUE PURCHASE REQUISITION - FINAL FIXED VERSION
 * Optimized for Account Details and flexible field mapping.
 */

/* ========== CONFIGURATION ========== */
const HEAD_ADMIN_EMAIL = "foodclique01@gmail.com";
const ADMIN_OFFICER_EMAIL = "Olarinde.ramat@gmail.com";
const RESPONSE_SHEET_NAME = "Requisition Form (Responses)"; 

// IMPORTANT: Paste your new Web App URL from the "Deploy" menu here
const MASTER_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwLNA3OZFm4JTNd2CEfOpcaBdbmnc-5LLlx0hklmiqiKsBxjjmGoQAD_TwqajDnoJRP/exec"; 

/* ========== TRIGGER: ON FORM SUBMIT ========== */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = e.namedValues;
  const rowIndex = e.range.getRow();

  function getVal(key) {
    const foundKey = Object.keys(rowData).find(k => k.toLowerCase().trim() === key.toLowerCase().trim());
    return foundKey ? rowData[foundKey][0] : null;
  }

  const staffName = getVal("Staff Name") || "Unknown Staff";
  const department = getVal("Department") || "N/A";
  const staffEmail = getVal("Email Address") || "";
  const accName = getVal("Account Name") || "Not Provided";
  const accNum = getVal("Account Number") || "Not Provided";
  const accBank = getVal("Bank") || "Not Provided";
  
  const rawTotal = getVal("Overall Estimated Total") || "0";
  const overallTotal = Number(rawTotal.toString().replace(/[^0-9.-]+/g, ""));

  let itemRowsHtml = "";
  let itemsArray = []; 

  for (let i = 1; i <= 7; i++) {
    const name  = getVal(`Item ${i}`); 
    const qty   = getVal(`Item ${i} quantity`) || getVal(`Item ${i} Quantity`) || "0";
    const rawUnit = getVal(`Item ${i} Unit Price`) || "0";
    const rawAmt  = getVal(`Item ${i} Amount`) || "0";
  
    // Convert to numbers safely
    const unit = Number(rawUnit.toString().replace(/[^0-9.-]+/g, ""));
    const amt  = Number(rawAmt.toString().replace(/[^0-9.-]+/g, ""));

    if (name && name.trim() !== "") {
      itemsArray.push({ name, qty, unit, amt });
      itemRowsHtml += `
        <tr>
          <td style="padding:10px 8px; border-bottom:1px solid #eee;">${name}</td>
          <td style="padding:10px 8px; border-bottom:1px solid #eee; text-align:center;">${qty}</td>
          <td style="padding:10px 8px; border-bottom:1px solid #eee; text-align:right;">₦${unit.toLocaleString()}</td>
          <td style="padding:10px 8px; border-bottom:1px solid #eee; text-align:right;"><b>₦${amt.toLocaleString()}</b></td>
        </tr>`;
    }
  }

  const statusCol = headers.indexOf("Approval Status") + 1;
  if (statusCol > 0) {
    sheet.getRange(rowIndex, statusCol).setValue("Awaiting Approval");
  }

  // Generate PDF
  const pdfBlob = generatePDF(staffName, department, overallTotal, accName, accNum, accBank, itemsArray);

  sendAdminNotification(staffName, staffEmail, department, overallTotal, itemRowsHtml, rowIndex, accName, accNum, accBank, pdfBlob);
}

/* ========== PDF GENERATOR ========== */
function generatePDF(name, dept, total, accName, accNum, accBank, items) {
  let html = `
    <html><body style="font-family: Arial;">
      <h1 style="text-align:center; color: #2c3e50;">FoodClique Requisition Receipt</h1>
      <hr>
      <p><b>Staff:</b> ${name} | <b>Dept:</b> ${dept}</p>
      <p><b>Payment To:</b> ${accName} | ${accNum} (${accBank})</p>
      <table border="1" style="width:100%; border-collapse: collapse; margin-top:20px;">
        <tr style="background:#f8f9fa;"><th>Item</th><th>Qty</th><th>Unit Price</th><th>Total</th></tr>`;
  
  items.forEach(item => {
    html += `<tr><td>${item.name}</td><td align="center">${item.qty}</td><td align="right">₦${item.unit.toLocaleString()}</td><td align="right">₦${item.amt.toLocaleString()}</td></tr>`;
  });

  html += `</table><h2 style="text-align:right;">Grand Total: ₦${total.toLocaleString()}</h2></body></html>`;
  return Utilities.newBlob(html, 'text/html', `Requisition_${name}.html`).getAs('application/pdf');
}

/* ========== ADMIN NOTIFICATION EMAIL ========== */
function sendAdminNotification(name, email, dept, total, tableRows, row, accName, accNum, accBank, pdfBlob) {
  const approveUrl = `${MASTER_WEB_APP_URL}?action=approve&row=${row}&amt=${total}&email=${encodeURIComponent(email)}`;
  const rejectUrl = `${MASTER_WEB_APP_URL}?action=reject&row=${row}&email=${encodeURIComponent(email)}`;

  const body = `
    <div style="font-family: Arial; max-width: 600px; border: 1px solid #ddd; padding: 20px;">
      <h2 style="color: #2c3e50;">Purchase Requisition Approval</h2>
      <p><b>Staff:</b> ${name}<br><b>Dept:</b> ${dept}</p>
      
      <div style="background-color: #f9f9f9; border-left: 4px solid #2980b9; padding: 15px; margin: 15px 0;">
        <h4 style="margin: 0 0 10px 0; color: #2980b9; text-transform: uppercase; font-size: 11px;">Payment Account Details</h4>
        <p style="margin: 2px 0;"><b>Name:</b> ${accName}</p>
        <p style="margin: 2px 0;"><b>No:</b> ${accNum}</p>
        <p style="margin: 2px 0;"><b>Bank:</b> ${accBank}</p>
      </div>

      <table style="width:100%; border-collapse: collapse;">
        <tr style="background: #f8f9fa;"><th align="left">Item</th><th>Qty</th><th align="right">Unit</th><th align="right">Total</th></tr>
        ${tableRows}
      </table>
      <h3 style="text-align:right;">Grand Total: ₦${total.toLocaleString()}</h3>
      
      <div style="margin-top: 30px; text-align: center;">
        <a href="${approveUrl}" style="background: #27ae60; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">APPROVE</a>
        &nbsp;
        <a href="${rejectUrl}" style="background: #e74c3c; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">REJECT</a>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: `${HEAD_ADMIN_EMAIL}, ${ADMIN_OFFICER_EMAIL}`,
    subject: `Approval Req: ${name} (₦${total.toLocaleString()})`,
    htmlBody: body,
    attachments: [pdfBlob]
  });
}

/* ========== WEB APP HANDLER (Approvals) ========== */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const row = Number(e.parameter.row);
    const approvedAmt = e.parameter.amt; 
    const staffEmail = e.parameter.email;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const statusCol = headers.indexOf("Approval Status") + 1;
    const amtCol = headers.indexOf("Approved Amount") + 1;

    if (action === "approve") {
      if (statusCol > 0) sheet.getRange(row, statusCol).setValue("Approved");
      if (amtCol > 0) sheet.getRange(row, amtCol).setValue(approvedAmt);
      if (staffEmail) notifyStaff(staffEmail, "Approved", Number(approvedAmt));

      return HtmlService.createHtmlOutput("<h2 style='color:green; text-align:center; padding-top:50px;'>✔ REQUISITION APPROVED</h2>")
             .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } 
    
    if (action === "reject") {
      if (statusCol > 0) sheet.getRange(row, statusCol).setValue("Rejected");
      if (amtCol > 0) sheet.getRange(row, amtCol).setValue(0);
      if (staffEmail) notifyStaff(staffEmail, "Rejected", 0);

      return HtmlService.createHtmlOutput("<h2 style='color:red; text-align:center; padding-top:50px;'>✘ REQUISITION REJECTED</h2>")
             .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    return HtmlService.createHtmlOutput("Error: " + err.message);
  }
}

function testSubmitWithBank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  
  const e = {
    range: range,
    namedValues: {
      "Staff Name": ["John Doe"],
      "Email Address": ["your-email@example.com"],
      "Department": ["Kitchen"],
      "Account Name": ["John FoodClique Doe"],
      "Account Number": ["0123456789"],
      "Bank": ["GTBank"],
      "Overall Estimated Total": ["₦15,000"],
      // Updated keys to match the logic in onFormSubmit
      "Item 1 Name": ["Rice Bag"], "Item 1 Quantity": ["1"], "Item 1 Unit Price": ["5000"], "Item 1 Amount": ["5000"],
      "Item 2 Name": ["Beans Bag"], "Item 2 Quantity": ["1"], "Item 2 Unit Price": ["5000"], "Item 2 Amount": ["5000"],
      "Item 3 Name": ["Oil Keg"], "Item 3 Quantity": ["1"], "Item 3 Unit Price": ["5000"], "Item 3 Amount": ["5000"]
    }
  };
  
  onFormSubmit(e);
}

function notifyStaff(email, status, amt) {
  const subject = `Update: Your Requisition has been ${status}`;
  const statusColor = status === "Approved" ? "#27ae60" : "#e74c3c";
  
  const body = `
    <div style="font-family: sans-serif; padding: 20px; border: 1px solid #eee;">
      <h2 style="color: ${statusColor};">Requisition ${status}</h2>
      <p>Hello,</p>
      <p>Your recent purchase requisition has been reviewed and marked as <b>${status}</b>.</p>
      ${status === "Approved" ? `<p><b>Approved Amount:</b> ₦${amt.toLocaleString()}</p>` : `<p>Reason: Budget constraints or missing details. Please contact the Admin Office for further clarification.</p>`}
      <br>
      <p>Best Regards,<br><b>FoodClique Admin Team</b></p>
    </div>
  `;

  MailApp.sendEmail({ to: email, cc: `${ADMIN_OFFICER_EMAIL}, ibrahim@foodclique.org`, subject: subject, htmlBody: body });
}

