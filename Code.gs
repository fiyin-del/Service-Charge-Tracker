function calculateCoverage() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Service Charge Tracker");
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {

    const weeklyRate = sheet.getRange(row, 5).getValue(); // Column E
    const paymentDate = sheet.getRange(row, 6).getValue(); // Column F
    const amountPaid = sheet.getRange(row, 7).getValue(); // Column G

    // Skip empty rows
    if (!weeklyRate || !paymentDate || !amountPaid) continue;

    // 1. Daily rate
    const dailyRate = weeklyRate / 7;

    // 2. Full days covered (rounded down)
    const fullDays = Math.floor(amountPaid / dailyRate);

    // 3. Coverage start (copy payment date)
    const coverageStart = new Date(paymentDate);

    // 4. Coverage end
    const coverageEnd = new Date(coverageStart);
    coverageEnd.setDate(coverageEnd.getDate() + fullDays - 1);

    // 5. Next payment due (day after coverage end)
    const nextPaymentDue = new Date(coverageEnd);
    nextPaymentDue.setDate(nextPaymentDue.getDate() + 1);

    // 6. Balance carried from previous row
    let balanceAtWeekStart = 0;

    if (row > 2) {
    balanceAtWeekStart = sheet.getRange(row - 1, 12).getValue(); // Column L
}

    // 7. Expected cost
    const expectedCost = fullDays * dailyRate;

    // 8. New balance after this payment
    const balanceAfterPayment = Math.round(
    (balanceAtWeekStart + amountPaid - expectedCost) * 100)/100;

    // 9. Arrears (only if balance is negative)
    const arrears =
    balanceAfterPayment < 0 ? Math.abs(balanceAfterPayment) : 0;


    // 9. Write results back to sheet
    sheet.getRange(row, 8).setValue(fullDays);          // Column H
    sheet.getRange(row, 9).setValue(coverageStart);    // Column I
    sheet.getRange(row, 10).setValue(coverageEnd);     // Column J
    sheet.getRange(row, 11).setValue(balanceAtWeekStart);  // Column K
    sheet.getRange(row, 12).setValue(balanceAfterPayment);  // Column L
    sheet.getRange(row, 13).setValue(nextPaymentDue);  // Column M
    sheet.getRange(row, 14).setValue(arrears); // Column N


  }
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Salvation Army Tools")
    .addItem("Calculate Coverage", "calculateCoverage")
    .addToUi();
}

function sendPaymentReminders() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Service Charge Tracker");

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const millisPerDay = 1000 * 60 * 60 * 24;

  for (let row = 2; row <= lastRow; row++) {

    const residentName = sheet.getRange(row, 1).getValue(); // A
    const email = sheet.getRange(row, 3).getValue();        // C
    const nextPaymentDue = sheet.getRange(row, 13).getValue(); // M
    const arrears = sheet.getRange(row, 14).getValue();        // N
    const reminderSent = sheet.getRange(row, 16).getValue();  // P

    // Skip if already reminded
    if (reminderSent === "YES") continue;

    if (!email || !(nextPaymentDue instanceof Date)) continue;

    const dueDate = new Date(nextPaymentDue);
    dueDate.setHours(0, 0, 0, 0);

    const daysUntilDue = Math.floor((dueDate - today) / millisPerDay);

    // Send reminder 1 day before OR on due date
    if (daysUntilDue === 1 || daysUntilDue === 0) {

      const subject = "Service Charge Payment Reminder";

      let message =
        `Dear ${residentName},\n\n` +
        `This is a reminder that your service charge payment is due on ` +
        `${dueDate.toLocaleDateString("en-GB")}.\n\n`;

      if (arrears > 0) {
        message +=
          `Our records show outstanding arrears of £${arrears.toFixed(2)}.\n\n`;
      }

      message +=
        `Please contact staff if you have any questions.\n\n` +
        `Kind regards,\n` +
        `The Salvation Army`;

      MailApp.sendEmail(email, subject, message);

      // Mark reminder as sent
      sheet.getRange(row, 16).setValue("YES");
    }
  }
}


function submitPayment() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = ss.getSheetByName("Data Entry");
  const trackerSheet = ss.getSheetByName("Service Charge Tracker");

  if (!entrySheet || !trackerSheet) {
    SpreadsheetApp.getUi().alert("Required sheets not found.");
    return;
  }

  // Read inputs
  const residentName = entrySheet.getRange("B2").getValue();
  const residentId   = entrySheet.getRange("B3").getValue();
  const email        = entrySheet.getRange("B4").getValue();
  const phone        = entrySheet.getRange("B5").getValue();
  const weeklyRate   = entrySheet.getRange("B6").getValue();
  const paymentDate  = entrySheet.getRange("B7").getValue();
  const amountPaid   = entrySheet.getRange("B8").getValue();
  const notes        = entrySheet.getRange("B9").getValue();

  // Validation — prevents silent failure
  if (!residentName || !weeklyRate || !paymentDate || !amountPaid) {
    SpreadsheetApp.getUi().alert(
      "Submission blocked",
      "Please complete all required fields before submitting.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Append row (O = Notes)
  trackerSheet.appendRow([
    residentName,  // A
    residentId,    // B
    email,         // C
    phone,         // D
    weeklyRate,    // E
    paymentDate,   // F
    amountPaid,    // G
    "", "", "", "", "", "", "", // H–N (calculated)
    notes          // O
  ]);

  // Run calculations AFTER append
  calculateCoverage();

  // Clear inputs ONLY after success
  entrySheet.getRange("B2:B9").clearContent();

  SpreadsheetApp.getActive().toast("Payment submitted successfully");
}
function testEmailNow() {
  MailApp.sendEmail(
    "rolufala@gmail.com",
    "TEST – Email system working",
    "If you receive this, MailApp works."
  );
}
