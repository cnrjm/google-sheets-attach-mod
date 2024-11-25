# Google Sheets File Attachment Add-on

## Overview
A Google Apps Script add-on that enables users to attach Google Drive files to cells in Google Sheets. The script creates hyperlinks in selected cells pointing to chosen files, streamlining document referencing within spreadsheets.

## Features
- Custom menu integration in Google Sheets
- Google Drive file selection
- New file upload support
- Clickable hyperlinks in cells
- Visible file names
- Real-time status updates

## Prerequisites
- Google account with Sheets and Drive access
- Google Cloud Console access
- Basic Google Apps Script knowledge

## Source Code

### appsscript.json
```json
{
  "timeZone": "Europe/London",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

### Code.gs
```javascript
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
```

### FilePicker.html
```html
<!DOCTYPE html>
<html>
  <head>
    <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
    <script type="text/javascript">
      var DEVELOPER_KEY = 'YOUR_DEVELOPER_KEY';
      var DIALOG_DIMENSIONS = {width: 600, height: 425};
      var pickerApiLoaded = false;

      function loadPickerApi() {
        gapi.load('picker', {'callback': onPickerApiLoad});
      }

      function onPickerApiLoad() {
        pickerApiLoaded = true;
        document.getElementById('status').innerHTML = 'Picker API loaded';
        document.getElementById('select-file').disabled = false;
      }

      function getOAuthToken() {
        document.getElementById('status').innerHTML = 'Getting OAuth token...';
        google.script.run
          .withSuccessHandler(createPicker)
          .withFailureHandler(showError)
          .getOAuthToken();
      }

      function createPicker(token) {
        if (pickerApiLoaded && token) {
          document.getElementById('status').innerHTML = 'Creating picker...';
          var view = new google.picker.View(google.picker.ViewId.DOCS);
          var uploadView = new google.picker.DocsUploadView();
          var picker = new google.picker.PickerBuilder()
            .addView(view)
            .addView(uploadView)
            .setOAuthToken(token)
            .setDeveloperKey(DEVELOPER_KEY)
            .setCallback(pickerCallback)
            .setOrigin(google.script.host.origin)
            .setSize(DIALOG_DIMENSIONS.width - 2, DIALOG_DIMENSIONS.height - 2)
            .build();
          picker.setVisible(true);
        } else {
          showError('Unable to load the file picker. Picker API loaded: ' + pickerApiLoaded + ', Token received: ' + (token ? 'Yes' : 'No'));
        }
      }

      function pickerCallback(data) {
        if (data[google.picker.Response.ACTION] == google.picker.Action.PICKED) {
          var doc = data[google.picker.Response.DOCUMENTS][0];
          var id = doc[google.picker.Document.ID];
          document.getElementById('status').innerHTML = 'File selected, processing...';
          google.script.run
            .withSuccessHandler(closeDialog)
            .withFailureHandler(showError)
            .processSelectedFile(id);
        } else if (data[google.picker.Response.ACTION] == google.picker.Action.CANCEL) {
          google.script.host.close();
        }
      }

      function closeDialog(message) {
        document.getElementById('status').innerHTML = message || 'Done!';
        setTimeout(function() {
          google.script.host.close();
        }, 2000);
      }

      function showError(message) {
        document.getElementById('error').innerHTML = 'Error: ' + message;
        console.error('Error:', message);
      }

      window.onload = function() {
        loadPickerApi();
      }
    </script>
    <script src="https://apis.google.com/js/api.js?onload=loadPickerApi"></script>
  </head>
  <body>
    <div>
      <button id="select-file" onclick="getOAuthToken()" disabled>Select or Upload a file</button>
    </div>
    <div id="status"></div>
    <div id="error" style="color: red;"></div>
  </body>
</html>
```

## Installation

### 1. Create Apps Script Project
1. Open Google Sheets
2. Navigate to `Extensions` > `Apps Script`
3. Clear default code
4. Create three files:
   ```
   Code.gs
   FilePicker.html
   appsscript.json
   ```

### 2. Google Cloud Console Setup
1. Visit [Google Cloud Console](https://console.cloud.google.com)
2. Create/select a project
3. Enable APIs:
   - Google Picker API
   - Google Drive API
4. Create credentials:
   ```
   API Key:
   - Navigate to Credentials
   - Create Credentials > API Key
   - Copy key for later use

   OAuth 2.0 Client ID:
   - Create Credentials > OAuth 2.0 Client ID
   - Choose "Web application"
   - Add authorized origins:
     * https://script.google.com
     * https://sheets.google.com
   ```

### 3. Code Configuration

1. Update `FilePicker.html`:
   ```javascript
   var DEVELOPER_KEY = 'YOUR_DEVELOPER_KEY';
   ```

2. Copy the code blocks above to their respective files:
   - `Code.gs`: Main script logic
   - `FilePicker.html`: UI components
   - `appsscript.json`: Project configuration

### 4. Deployment
1. Select `Deploy` > `New deployment`
2. Choose `Web app`
3. Configure:
   - Execute as: `User accessing the web app`
   - Access: `Anyone`
4. Complete deployment
5. Grant necessary permissions
6. Save deployment ID

## Usage
1. Refresh Google Sheet
2. Find `Custom Tools` menu
3. Select target cell
4. Click `Custom Tools` > `Attach File`
5. Select/upload file
6. Hyperlink appears in cell

## Troubleshooting

### Common Issues
- **Missing Menu**
  - Refresh sheet
  - Check authorization

- **Picker Issues**
  - Verify API credentials
  - Check OAuth setup

- **Authorization Errors**
  - Confirm OAuth scopes in `appsscript.json`

- **Upload Problems**
  - Verify Drive API activation

## Security

### Considerations
- Secure API credential storage
- User-level permissions
- Operation logging enabled
- Standard Google security protocols

### Permissions
The add-on:
- Runs with user's permissions
- Accesses only authorized files
- Logs all operations

## Limitations
- Single file per cell
- Requires internet connection
- Google account required
- Fixed picker size (600x425px)
- Google Drive files only

## Contributing
Contributions welcome through pull requests. Please:
- Maintain existing code style
- Document new features
- Test thoroughly

## License
MIT License - Modify and use freely

---

*For questions or issues, please open a GitHub issue or contact the maintainer.*