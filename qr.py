import qrcode
import json

class QRGenerator:
    def __init__(self):
        self.qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

    def generate(self, data: dict, qr_filename: str):
        json_string = json.dumps(data)
        # Add JSON data to QR code
        self.qr.add_data(json_string)
        self.qr.make(fit=True)

        # Create an image from the QR Code instance
        img = self.qr.make_image(fill_color="black", back_color="white")

        # Save the QR code image
        img.save(qr_filename)


if __name__ == "__main__":
    data = {
        "name": "Dat Dao",
        "email": "datdq11@gmail.com",
        "phone": "0123456789",
        "website": "https://example.com"
    }

    saved_tmp_path = "qrcode_from_json.png"

    qr_generator = QRGenerator()

    qr_generator.generate(data=data, qr_filename=saved_tmp_path)
