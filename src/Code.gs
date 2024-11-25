function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Tools')
    .addItem('Attach File', 'showFilePicker')
    .addToUi();
}

function showFilePicker() {
  var html = HtmlService.createHtmlOutputFromFile('FilePicker')
    .setWidth(600)
    .setHeight(425)
    .setTitle('Select or Upload a File');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select or Upload a File');
}

function getOAuthToken() {
  DriveApp.getRootFolder(); // This line ensures the script has the necessary scope
  return ScriptApp.getOAuthToken();
}

function processSelectedFile(fileId) {
  var file = DriveApp.getFileById(fileId);
  var fileUrl = file.getUrl();
  var fileName = file.getName();
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  
  cell.setValue(fileName);
  cell.setFormula('=HYPERLINK("' + fileUrl + '", "' + fileName + '")');
  
  return 'File "' + fileName + '" linked successfully';
}