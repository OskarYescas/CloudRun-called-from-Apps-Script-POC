# Serverless PDF Generator for Google Sheets (GCP Cloud Run + Apps Script)

A scalable, secure solution to generate professional "Fit-to-Page" PDF reports from Google Sheets using Google Cloud Run (Python) and Google Apps Script.

## 🏗 Architecture: The "Pull" Pattern

This project solves the "Payload Limit" and "Timeout" constraints often faced when processing large datasets directly in Google Apps Script.

**The Workflow:**
1.  **Trigger:** User clicks a menu item in Google Sheets.
2.  **Signal:** Apps Script sends a lightweight signal (Spreadsheet ID only) to Cloud Run. **No row data is transmitted by the client.**
3.  **Pull (Backend):** Cloud Run authenticates via the user's OAuth token (Pass-through Auth), connects directly to the Google Sheets API, and "pulls" the data in a single batch request.
4.  **Render:** A Python engine (`ReportLab`) calculates dynamic column widths to fit a standard US Letter Landscape page.
5.  **Save:** The PDF is uploaded directly to the **same folder** as the source Spreadsheet using the Drive API.

## ✨ Features

* **Zero-Trust Security:** Validates OAuth2 Access Tokens and enforces domain restrictions (`ALLOWED_DOMAIN`) at the application layer.
* **High Performance:** Uses `spreadsheets.values.batchGet` to fetch all tabs in a single API call.
* **Smart Layout:** Automatically calculates column widths to prevent data truncation ("Fit-to-Width").
* **User Experience:** PDF appears in the same Drive folder as the source Sheet; UI Modal provides direct links.

## 🚀 Deployment Guide

### Phase 1: Google Cloud Platform (Cloud Run)

1.  **Create a GCP Project** and enable:
    * Cloud Run Admin API
    * Google Drive API
    * Google Sheets API
2.  **Deploy the Python Service:**
    ```bash
    gcloud run deploy pdf-generator \
      --source . \
      --platform managed \
      --region us-central1 \
      --allow-unauthenticated \
      --set-env-vars ALLOWED_DOMAIN="your-organization.com"
    ```
    * *Note: We use `--allow-unauthenticated` because we validate the User Access Token manually in the code (Pass-Through Auth).*
3.  **Copy the Service URL** (e.g., `https://pdf-generator-xyz.a.run.app`).

### Phase 2: Google Apps Script

1.  Create a new Apps Script project attached to a Google Sheet.
2.  **Copy the Code:**
    * Paste `Code.gs` into the script editor.
    * Update `appsscript.json` with the required Manifest Scopes.
3.  **Set Configuration:**
    * Go to **Project Settings > Script Properties**.
    * Add a property: `CLOUD_RUN_URL` = `YOUR_SERVICE_URL_FROM_PHASE_1`.
4.  **Reload the Sheet:** You will see a custom menu "POC PDF".

## 🛠 Local Development

**Requirements:**
* Python 3.9+
* `pip install -r requirements.txt`

**Environment Variables:**
* `ALLOWED_DOMAIN`: The email domain allowed to generate reports (e.g., `example.com`).
