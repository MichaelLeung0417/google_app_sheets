function onSpreadsheetChange() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0];
  Logger.log(`get sheet values success, ${values}`);

  var folderId = "1D1S7YCB7yKvrHsGRM8Y0jpvb0j1DGDOS";
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName(`${lastRow}-${values[0]}.pdf`);
  if (files.hasNext() && files.next().getMimeType() === "application/pdf") {
    Logger.log("PDF existed.");
    return; // PDF file with the same name exists
  } else {
    // Load the template document
    var templateId = "1mgEo2yEXQH8e-uRY-9_dwqae_SgHxifTOgaYVVpAh_g"; // Replace with the ID of your template document
    var template = DriveApp.getFileById(templateId);

    // Create a copy of the template as a new document
    var newDocFile = template.makeCopy();
    var newDoc = DocumentApp.openById(newDocFile.getId());
    var body = newDoc.getBody();

    // Replace the placeholders with actual values
    body.replaceText("<<CHINESENAME>>", values[0]);
    body.replaceText("<<ENGLISHNAME>>", values[1]);
    body.replaceText("<<GENDER>>", values[2]);
    body.replaceText("<<PHONENUMBER>>", values[3]);
    body.replaceText("<<EMAIL>>", values[4]);
    body.replaceText("<<CHURCH>>", values[5]);
    body.replaceText("<<SERVICE>>", values[6]);
    // Replace more placeholders with corresponding values

    // Save and close the new document
    newDoc.saveAndClose();

    // Convert the new document to PDF
    var pdfFile = DriveApp.getFileById(newDocFile.getId()).getAs(
      "application/pdf"
    );
    folder.createFile(pdfFile).setName(`${lastRow}-${values[0]}.pdf`);
    DriveApp.getFileById(newDocFile.getId()).setTrashed(true);
    Logger.log("PDF created.");
  }
}
