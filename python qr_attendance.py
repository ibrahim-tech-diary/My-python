import qrcode
import os


output_dir = "qr_codes"
os.makedirs(output_dir, exist_ok=True)

students = [
    {"id": "239211", "name": "Ibrahim",  "dept": "CST", "sem": "5th"},
    {"id": "203077", "name": "Rahim",  "dept": "TAXTILE", "sem": "3rd"},
    {"id": "203108", "name": "Karim",  "dept": "ET", "sem": "2nd"},
    {"id": "239822", "name": "Sakib",  "dept": "CST", "sem": "6th"},
    {"id": "239212", "name": "Nabil",  "dept": "CIVIL", "sem": "1st"},
]

for student in students:
   
    data = f"ID:{student['id']};Name:{student['name']};Dept:{student['dept']};Semester:{student['sem']}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    filename = os.path.join(output_dir, f"QR_{student['id']}.png")
    img.save(filename)
    print(f"✅ Saved: {filename} | Data: {data}")

print("\n🎉 All QR Codes Generated!")
print(f"📁 Folder: {os.path.abspath(output_dir)}")