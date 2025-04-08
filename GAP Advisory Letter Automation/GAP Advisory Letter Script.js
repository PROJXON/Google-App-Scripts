function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Advisor Tools')
      .addItem('Generate and Send Documents', 'processAdvisoryDocs')
      .addToUi();
  }
  
  function processAdvisoryDocs() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0]; // First row as headers
    
    var templateId = "1qgdKi32b3zphmxBN-9s2WfHcuzLitZ2zY2Z39z6Tjv0"; // Template ID
    var folderId = "13GBaIxh3_g1BGmJsgF8ZYLUkquP3ge-M"; // Correct Google Drive folder ID
    
    var docDraftedColumn = headers.indexOf("Doc Drafted");
    var emailSentColumn = headers.indexOf("Email Sent");
    var advisorEmailColumn = headers.indexOf("Advisor Email");
  
    var folder = DriveApp.getFolderById(folderId); // Target folder
  
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
  
      // Skip row if already drafted or sent
      if (
        (docDraftedColumn !== -1 && row[docDraftedColumn] === "Drafted!") || 
        (emailSentColumn !== -1 && row[emailSentColumn] === "Sent!")
      ) {
        Logger.log(`Skipping row ${i + 1} because it's already processed.`);
        continue; // Skip this row
      }
  
      // Create a new document from the template
      var advisorName = row[headers.indexOf("Advisor Name")] || "Unknown_Advisor";
      var doc = DriveApp.getFileById(templateId).makeCopy(`Advisor_Document_${advisorName}`, folder); // Save directly to the target folder
      var docId = doc.getId();
      var docFile = DocumentApp.openById(docId);
      var body = docFile.getBody();
  
      // Replace placeholders in the document
      for (var j = 0; j < headers.length; j++) {
        var placeholder = `{{${headers[j]}}}`;
        body.replaceText(placeholder, row[j] ? row[j].toString() : '');
      }
  
      docFile.saveAndClose();
  
      // Convert to PDF and save in the target folder
      var pdfBlob = doc.getAs('application/pdf');
      var pdfFile = folder.createFile(pdfBlob.setName(`Advisor_Document_${advisorName}.pdf`));
      
      // Send email if email is available
      var recipient = row[advisorEmailColumn];
      if (recipient) {
        var subject = "Your Advisory Program Document";
        var bodyText = `Hello,\n\nPlease find attached your advisory document.\n\nBest regards.`;
        MailApp.sendEmail({
          to: recipient,
          subject: subject,
          body: bodyText,
          attachments: [pdfFile]
        });
      }
  
      // Update the spreadsheet to reflect status
      if (docDraftedColumn !== -1) {
        sheet.getRange(i + 1, docDraftedColumn + 1).setValue("Drafted!");
      }
      if (emailSentColumn !== -1) {
        sheet.getRange(i + 1, emailSentColumn + 1).setValue("Sent!");
      }
    }
  }