from qr import QRGenerator
from excel import QRUploadSheet

class QRSheet:
    def __init__(self, token_path):
        self.qr_generator = QRGenerator()
        self.qr_upload_sheet = QRUploadSheet(token_path=token_path)
    
    def execute(self, data, excel_id, qr_filename, cell_id):
        self.qr_generator.generate(data, qr_filename)
        qr_id = data.get("qr_id")
        self.qr_upload_sheet.main(excel_id=excel_id, qr_id=qr_id, qr_filename=qr_filename, cell_id=cell_id)


if __name__ == "__main__":
    from uuid import uuid4
    data = {
        "qr_id": str(uuid4()),
        "name": "Dat Dao",
        "email": "datdq11@gmail.com",
        "phone": "0123456789",
        "website": "https://example.com"
    }
    excel_id="1iAOi7sCC8u7DSEf093UJzwxrTRif3X5ClWSzkadmdps" # gg sheet id, which can find in link https://docs.google.com/spreadsheets/d/<sheet-id>/edit?gid=0#gid=0
    qr_filename = "/Users/datdq98/Desktop/experiments/qr-testing/qrcode_from_json.png" # change location to save qr code
    cell_id=2 # row number where the data saved

    qr_sheet = QRSheet(
        token_path="/Users/datdq98/Desktop/experiments/qr-testing/token.json" # token generating when run this code
    )

    qr_sheet.execute(data=data, excel_id=excel_id, qr_filename=qr_filename, cell_id=cell_id)
