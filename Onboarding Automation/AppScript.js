function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Offer Letter Tools")
    .addItem("Generate and Send Offer Letters", "processOfferLetters")
    .addToUi();
}
function formatDate(date) {
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return new Date(date).toLocaleDateString('en-US', options);
}
function processOfferLetters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Offer Letters");
  const data = sheet.getDataRange().getValues();

  // Column indices (update these if the column positions in your sheet change)
  const headers = data[0];
  const statusIndex = headers.indexOf("Status");
  const nameIndex = headers.indexOf("Candidate Name");
  const roleIndex = headers.indexOf("Role Title");
  const yearQIndex = headers.indexOf("Yr-Q");
  const emailIndex = headers.indexOf("Email");
  const startDateIndex = headers.indexOf("Start Date");
  const endDateIndex = headers.indexOf("End Date");
  const programLengthIndex = headers.indexOf("Program Length");
  const hoursIndex = headers.indexOf("Hours");
  const teamIndex = headers.indexOf("Team");
  const responsibility1Index = headers.indexOf("Responsibility 1");
  const r1TaskAIndex = headers.indexOf("R1-Task A");
  const r1TaskBIndex = headers.indexOf("R1-Task B");
  const r1TaskCIndex = headers.indexOf("R1-Task C");
  const responsibility2Index = headers.indexOf("Responsibility 2");
  const r2TaskAIndex = headers.indexOf("R2-Task A");
  const r2TaskBIndex = headers.indexOf("R2-Task B");
  const r2TaskCIndex = headers.indexOf("R2-Task C");
  const responsibility3Index = headers.indexOf("Responsibility 3");
  const r3TaskAIndex = headers.indexOf("R3-Task A");
  const r3TaskBIndex = headers.indexOf("R3-Task B");
  const r3TaskCIndex = headers.indexOf("R3-Task C");
  const olDateIndex = headers.indexOf("OL Date");
  const recruiter = headers.indexOf("Recruiter");
  const recruiterRole = headers.indexOf("Recruiter Role");
  const recruiterEmail = headers.indexOf("Recruiter Email");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[statusIndex] === "completed") continue; // Skip if already completed

    // Replace variables in the template
    const template = DriveApp.getFileById("1Ul5Uyx73NSeZl1YVfTsN3l7XC8zvcXvHxgKUjLap7BY"); // Replace with the Google Doc template ID
    const doc = DocumentApp.openById(template.makeCopy().getId());
    const body = doc.getBody();
    body.replaceText("{{Candidate Name}}", row[nameIndex]);
    body.replaceText("{{Role Title}}", row[roleIndex]);
    body.replaceText("{{Yr-Q}}", row[yearQIndex]);
    body.replaceText("{{Program Length}}", row[programLengthIndex]);
    body.replaceText("{{Hours}}", row[hoursIndex]);
    body.replaceText("{{Team}}", row[teamIndex]);
    body.replaceText("{{Responsibility 1}}", row[responsibility1Index]);
    body.replaceText("{{R1-Task A}}", row[r1TaskAIndex]);
    body.replaceText("{{R1-Task B}}", row[r1TaskBIndex]);
    body.replaceText("{{R1-Task C}}", row[r1TaskCIndex]);
    body.replaceText("{{Responsibility 2}}", row[responsibility2Index]);
    body.replaceText("{{R2-Task A}}", row[r2TaskAIndex]);
    body.replaceText("{{R2-Task B}}", row[r2TaskBIndex]);
    body.replaceText("{{R2-Task C}}", row[r2TaskCIndex]);
    body.replaceText("{{Responsibility 3}}", row[responsibility3Index]);
    body.replaceText("{{R3-Task A}}", row[r3TaskAIndex]);
    body.replaceText("{{R3-Task B}}", row[r3TaskBIndex]);
    body.replaceText("{{R3-Task C}}", row[r3TaskCIndex]);
    body.replaceText("{{Start Date}}", formatDate(row[startDateIndex]));
    body.replaceText("{{End Date}}", formatDate(row[endDateIndex]));
    body.replaceText("{{OL Date}}", formatDate(row[olDateIndex]));
    doc.saveAndClose();
    // Save as PDF
    const pdfBlob = doc.getAs(MimeType.PDF);
    const pdfFileName = `PROJXON Offer Letter - ${row[nameIndex]}`;
    const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfFileName);

    // Send email
    const emailSubject = `Your Offer Letter is here ${row[nameIndex]}`;
    const emailBodyHtml = `
    <p>Dear ${row[nameIndex]},</p>

    <p>We are delighted to formally extend an offer for the position of <strong>${row[roleIndex]}</strong> at our organization. Congratulations on being selected for this exciting opportunity!</p>

    <p>Attached, you will find your official offer letter. Please take the time to review it thoroughly, and don’t hesitate to reach out if you have any questions or need clarification.</p>

    <p>We kindly ask that you sign and return the offer letter within <strong>1 week</strong> to confirm your acceptance.</p>

    <p>Welcome aboard, and we look forward to having you as part of our team!</p>

    <p>P.S. If you need any adjustments to your offer letter please contact me directly.</p>

    <br>

    <p>Best regards,</p>
    <p><strong><i>${row[recruiter]}</i></strong><br>
    <i>${row[recruiterRole]}</i><br>
    <strong>PROJXON</strong><br>
    <a href="mailto:${row[recruiterEmail]}">${row[recruiterEmail]}</a></p>
    `;

    GmailApp.sendEmail(row[emailIndex], emailSubject, '', {
      htmlBody: emailBodyHtml,
      cc: 'recruiting@projxon.com',
      attachments: [pdfFile],
    });


    // Mark as completed
    sheet.getRange(i + 1, statusIndex + 1).setValue("completed");

    // Remove temporary Google Doc
    DriveApp.getFileById(doc.getId()).setTrashed(true);
  }
}

