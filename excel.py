import requests
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


class QRUploadSheet:
    def __init__(self, token_path):
        self.token_path = token_path
        self.SCOPES = [
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/spreadsheets'
        ]


    def get_services(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file(self.token_path, self.SCOPES)
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.SCOPES)
            creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        return drive_service, sheets_service

    # Upload QR code to Google Drive
    def upload_to_drive(self, drive_service, filename):
        file_metadata = {
            'name': filename,
            'mimeType': 'image/png'
        }
        media = MediaFileUpload(filename, mimetype='image/png')
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        # Make the file publicly accessible
        drive_service.permissions().create(
            fileId=file['id'],
            body={'role': 'reader', 'type': 'anyone'}
        ).execute()
        return f"https://drive.google.com/uc?export=download&id={file['id']}"


    # Insert image into Google Sheet
    def insert_image_to_sheet(self, sheets_service, spreadsheet_id, sheet_name, image_url, qr_id, cell_id='A2', cell_image='B2', width_px=200, height_px=200):
        sheet_id = self.get_sheet_id(sheets_service, spreadsheet_id, sheet_name)
            
        # Insert QR ID (JSON string) in cell A2
        id_request = {
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': int(cell_id[1:]) - 1,
                    'endRowIndex': int(cell_id[1:]),
                    'startColumnIndex': ord(cell_id[0].upper()) - ord('A'),
                    'endColumnIndex': ord(cell_id[0].upper()) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {'stringValue': qr_id}
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        }

        # Insert =IMAGE() formula in cell B2
        formula = f'=IMAGE("{image_url}")'
        image_request = {
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': int(cell_image[1:]) - 1,
                    'endRowIndex': int(cell_image[1:]),
                    'startColumnIndex': ord(cell_image[0].upper()) - ord('A'),
                    'endColumnIndex': ord(cell_image[0].upper()) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {'formulaValue': formula}
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        }

        # Set column width (for column B)
        column_request = {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': ord(cell_image[0].upper()) - ord('A'),
                    'endIndex': ord(cell_image[0].upper()) - ord('A') + 1
                },
                'properties': {
                    'pixelSize': width_px
                },
                'fields': 'pixelSize'
            }
        }

        # Set row height (for row 2)
        row_request = {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': int(cell_image[1:]) - 1,
                    'endIndex': int(cell_image[1:])
                },
                'properties': {
                    'pixelSize': height_px
                },
                'fields': 'pixelSize'
            }
        }

        # Batch the requests
        batch_update_request = {
            'requests': [id_request, image_request, column_request, row_request]
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=batch_update_request
        ).execute()

    def get_sheet_id(self, sheets_service, spreadsheet_id, sheet_name):
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        raise ValueError(f"Sheet '{sheet_name}' not found")

    # Main execution
    def main(self, excel_id, qr_id, qr_filename, cell_id):
        # Replace with your Google Sheet ID and sheet name
        SPREADSHEET_ID = excel_id
        SHEET_NAME = 'Sheet1'
        
        CUSTOM_WIDTH = 200  # Custom width in pixels
        CUSTOM_HEIGHT = 200  # Custom height in pixels

        drive_service, sheets_service = self.get_services()
        image_url = self.upload_to_drive(drive_service, qr_filename)
        self.insert_image_to_sheet(
            sheets_service,
            SPREADSHEET_ID,
            SHEET_NAME,
            image_url,
            qr_id=qr_id,
            cell_id=f'A{cell_id}',
            cell_image=f'B{cell_id}',
            width_px=CUSTOM_WIDTH,
            height_px=CUSTOM_HEIGHT
        )

if __name__ == '__main__':
    from uuid import uuid4
    qr_id = str(uuid4())
    qr_filename="/Users/datdq98/Desktop/experiments/qr-testing/qrcode_from_json.png"
    qr_upload = QRUploadSheet(token_path="/Users/datdq98/Desktop/experiments/qr-testing/token.json")
    qr_upload.main(excel_id="1iAOi7sCC8u7DSEf093UJzwxrTRif3X5ClWSzkadmdps",qr_id=qr_id, qr_filename=qr_filename, cell_id=3)