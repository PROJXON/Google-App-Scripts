function setFolderPermissions() {
  // Open the Google Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get the range with emails (e.g., column A starting from row 2)
  var emailRange = sheet.getRange("A2:A" + sheet.getLastRow());
  var emails = emailRange.getValues(); // Get emails from the range
  
  // Specify the folder ID or the folder object you want to share
  var folderId = 'YOUR_FOLDER_ID'; // Replace with your actual folder ID
  var folder = DriveApp.getFolderById(folderId);
  
  // Loop through the emails and give permissions
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i][0];
    if (email) {
      // Add permission to the folder for each email
      folder.addEditor(email); // You can use .addViewer(email) for view-only access
    }
  }
}
