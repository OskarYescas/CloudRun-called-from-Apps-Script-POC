// --- CONFIGURATION ---
// Retrieve the URL securely from Project Settings > Script Properties
const properties = PropertiesService.getScriptProperties();
const CLOUD_RUN_URL = properties.getProperty('CLOUD_RUN_URL');

/**
 * Main function triggered by the menu.
 * Sends the Spreadsheet ID to Cloud Run to initiate the "Pull" process.
 */
function createPdfFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Validation: Ensure URL property is set
  if (!CLOUD_RUN_URL) {
    SpreadsheetApp.getUi().alert("Configuration Error: CLOUD_RUN_URL script property is missing.");
    return;
  }

  // 1. Package Metadata ("Pull Pattern")
  const payload = { 
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName()
  };

  // 2. Get User's OAuth Token
  const token = ScriptApp.getOAuthToken();
  if (!token) {
    SpreadsheetApp.getUi().alert("Error: Could not retrieve Access Token.");
    return;
  }

  // 3. Prepare Request
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
    timeoutSeconds: 180 // 3 minutes
  };

  try {
    // 4. Call Cloud Run
    const response = UrlFetchApp.fetch(CLOUD_RUN_URL, options);
    
    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());

      if (result.status === 'success') {
        // UX Improvement: Determine the parent folder URL locally
        // Since Python saves to the same parent, we can just look up our own parent here.
        const parentFolderUrl = getParentFolderUrl(ss.getId());
        
        showSuccessModal(result.url, result.file_id, parentFolderUrl);
      } else {
        SpreadsheetApp.getUi().alert("Server Error: " + (result.details || "Unknown error"));
      }
    } else {
      SpreadsheetApp.getUi().alert(`HTTP Error (${response.getResponseCode()}): ${response.getContentText()}`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Connection Failed: ${e.toString()}`);
  }
}

/**
 * Helper to get the URL of the immediate parent folder.
 * This ensures the "Open Folder" button goes to where the PDF was actually saved.
 */
function getParentFolderUrl(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    if (parents.hasNext()) {
      return parents.next().getUrl();
    }
    // Fallback to Root if no parent found (rare)
    return "https://drive.google.com/drive/my-drive";
  } catch (e) {
    console.error("Error fetching parent folder:", e);
    return "https://drive.google.com/drive/my-drive";
  }
}

/**
 * Displays the success modal with links to the specific file and folder.
 */
function showSuccessModal(fileUrl, fileId, folderUrl) {
  const htmlTemplate = `
    <div style="font-family: 'Google Sans', Roboto, sans-serif; text-align: center; padding: 20px;">
      <h2 style="color: #188038; margin-top: 0;">Success!</h2>
      <p style="color: #3c4043;">Your PDF has been generated and saved.</p>
      
      <div style="margin: 25px 0 15px 0;">
        <a href="${fileUrl}" target="_blank" style="
            display: inline-block;
            background-color: #1a73e8; 
            color: white; 
            padding: 10px 24px; 
            text-decoration: none; 
            border-radius: 4px; 
            font-weight: 500;
            font-size: 14px;
            box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3);">
          OPEN PDF REPORT
        </a>
      </div>

      <div style="margin-bottom: 25px;">
        <a href="${folderUrl}" target="_blank" style="
            display: inline-block;
            background-color: #f1f3f4; 
            color: #3c4043; 
            padding: 8px 16px; 
            text-decoration: none;
            border-radius: 4px; 
            font-size: 13px;
            border: 1px solid #dadce0;">
          Open Folder
        </a>
      </div>

      <p style="color: #5f6368; font-size: 11px;">File ID: ${fileId}</p>
      <br>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
      .setWidth(400)
      .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Report Generator');
}

/**
 * Creates the menu item in Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('POC PDF')
    .addItem('Generate Report', 'createPdfFromSheets')
    .addToUi();
}