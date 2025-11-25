import os
import io
import requests
import logging
from flask import Flask, request, jsonify

# ReportLab Imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Google API Imports
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Setup Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# --- CONFIGURATION ---
# Security: Retrieve strictly from Environment Variable.
# If not set in Cloud Run, this defaults to None (which disables the domain check unless handled).
ALLOWED_DOMAIN = os.environ.get('ALLOWED_DOMAIN') 
# ---------------------

def verify_user(request):
    """
    Validates the User Access Token and checks the Domain.
    """
    auth_header = request.headers.get('Authorization')
    if not auth_header: 
        return None, "Missing Authorization Header", None
    
    token = auth_header.split(' ')[1] if len(auth_header.split(' ')) > 1 else auth_header
    
    try:
        # Validate token against Google UserInfo API
        response = requests.get('https://www.googleapis.com/oauth2/v3/userinfo', headers={'Authorization': f'Bearer {token}'})
        
        if response.status_code != 200: 
            logger.warning(f"Invalid Token provided. Status: {response.status_code}")
            return None, "Invalid Token", None
        
        id_info = response.json()
        user_email = id_info.get('email', '')

        # Domain Validation
        if ALLOWED_DOMAIN:
            if not user_email.endswith(f"@{ALLOWED_DOMAIN}"):
                logger.warning(f"Unauthorized domain access attempt: {user_email}")
                return None, f"Unauthorized domain: {user_email}", None
        
        return id_info, None, token

    except Exception as e: 
        logger.error(f"Auth Exception: {str(e)}")
        return None, str(e), None

@app.route('/', methods=['POST'])
def generate_and_save_pdf():
    # 1. Security Check
    user_info, error_msg, token = verify_user(request)
    if error_msg: 
        return jsonify({"error": "Unauthorized", "details": error_msg}), 401

    try:
        data = request.get_json()
        spreadsheet_id = data.get('spreadsheetId')
        spreadsheet_name = data.get('spreadsheetName', 'Report')
        
        # 2. Setup Google Services
        creds = Credentials(token=token)
        sheets_service = build('sheets', 'v4', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)
        
        # 3. Fetch Data & Metadata
        # Get Sheet structure
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        
        # Get Parent Folder (UX Improvement)
        # We fetch the file's parents so we can save the PDF in the same location.
        file_meta = drive_service.files().get(fileId=spreadsheet_id, fields='parents').execute()
        parents = file_meta.get('parents', [])
        parent_folder_id = parents[0] if parents else None

        # 4. OPTIMIZATION: Batch Fetch Data
        # Instead of looping API calls, we define all ranges and fetch once.
        sheet_titles = [s.get('properties', {}).get('title') for s in sheets]
        
        if not sheet_titles:
            return jsonify({'error': 'No visible sheets found'}), 400

        # One single HTTP request for all data
        batch_result = sheets_service.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id, 
            ranges=sheet_titles
        ).execute()
        
        value_ranges = batch_result.get('valueRanges', [])

        # 5. PDF Generation (In-Memory)
        buffer = io.BytesIO()
        # Standard US Letter Landscape
        doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        # Styles
        header_style = ParagraphStyle('H', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.whitesmoke)
        cell_style = ParagraphStyle('C', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8)

        # Title
        elements.append(Paragraph(f"Report: {spreadsheet_name}", styles['Title']))
        elements.append(Spacer(1, 12))

        # Loop through the BATCH results
        for i, sheet_data in enumerate(value_ranges):
            # valueRanges order matches the requested ranges order
            title = sheet_titles[i]
            rows = sheet_data.get('values', [])
            
            if not rows: continue

            elements.append(Paragraph(f"Tab: {title}", styles['Heading2']))
            elements.append(Spacer(1, 5))

            # Table Logic
            num_cols = max(len(r) for r in rows)
            if num_cols == 0: continue
            
            # Dynamic Width Calculation
            col_width = 750 / num_cols 

            table_data = []
            
            # Header Processing
            header_row = rows[0] + [''] * (num_cols - len(rows[0]))
            table_data.append([Paragraph(str(c), header_style) for c in header_row])
            
            # Body Processing
            for row in rows[1:]:
                padded_row = row + [''] * (num_cols - len(row))
                table_data.append([Paragraph(str(c), cell_style) for c in padded_row])

            # Style Table
            t = Table(table_data, colWidths=[col_width] * num_cols, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.darkgrey),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ]))
            
            elements.append(t)
            elements.append(PageBreak())

        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        # 6. Upload to Drive
        file_metadata = {
            'name': f"{spreadsheet_name}.pdf", 
            'mimeType': 'application/pdf'
        }
        
        # If we found a parent folder, save the PDF there.
        if parent_folder_id:
            file_metadata['parents'] = [parent_folder_id]

        file = drive_service.files().create(
            body=file_metadata,
            media_body=MediaIoBaseUpload(buffer, mimetype='application/pdf', resumable=True),
            fields='id, webViewLink'
        ).execute()

        logger.info(f"PDF generated successfully: {file.get('id')}")
        
        return jsonify({'status': 'success', 'file_id': file.get('id'), 'url': file.get('webViewLink')})

    except Exception as e:
        logger.error(f"Processing Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == "__main__":
    # Ensure PORT is handled for Cloud Run
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))