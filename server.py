from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_socketio import SocketIO
import qrcode
import io
import base64
import threading
import time
import random
import string
import pandas as pd
import os

app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*")

EXCEL_FILE = "students.xlsx"

# ✅ التأكد من أن ملف Excel موجود ومهيأ بشكل صحيح
def ensure_excel_file():
    """تأكد من وجود ملف Excel، وإذا لم يكن موجودًا يتم إنشاؤه"""
    if not os.path.exists(EXCEL_FILE) or os.stat(EXCEL_FILE).st_size == 0:
        df = pd.DataFrame(columns=["ID", "Name", "Scanned QR Code"])
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        print("✅ تم إنشاء ملف students.xlsx.")

ensure_excel_file()

# ✅ تحديث QR Code كل ثانيتين وإرساله إلى الصفحة باستخدام WebSocket
def generate_qr():
    """تحديث QR Code كل ثانيتين وإرساله إلى الصفحة باستخدام WebSocket"""
    while True:
        with app.app_context():
            random_data = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            scan_url = f"http://192.168.1.11:5000/scan/{random_data}"  # عدّل IP حسب شبكتك

            qr = qrcode.make(scan_url)
            buffer = io.BytesIO()
            qr.save(buffer, format="PNG")
            buffer.seek(0)
            qr_base64 = base64.b64encode(buffer.getvalue()).decode()

            socketio.emit('update_qr', {"qr_code": f"data:image/png;base64,{qr_base64}"})  
            time.sleep(2)

threading.Thread(target=generate_qr, daemon=True).start()

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/students")
def students():
    ensure_excel_file()
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    students_list = df.to_dict(orient="records")
    return render_template("students.html", students=students_list)

@app.route("/scan/<qr_data>", methods=["GET", "POST"])
def scan(qr_data):
    if request.method == "POST":
        student_id = request.form.get("student_id")
        student_name = request.form.get("student_name")

        if not student_id or not student_name:
            return "❌ يرجى إدخال جميع البيانات.", 400

        ensure_excel_file()

        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        new_data = pd.DataFrame({"ID": [student_id], "Name": [student_name], "Scanned QR Code": [qr_data]})
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

        print(f"✅ تم تسجيل الطالب: {student_name} (ID: {student_id})")

        return redirect(url_for("students"))

    return render_template("scan.html", qr_data=qr_data)

# ✅ تحميل جميع الأسماء عند تنزيل ملف Excel
@app.route("/download_excel")
def download_excel():
    ensure_excel_file()
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    
    # ✅ حفظ الملف قبل التنزيل لضمان وجود جميع البيانات
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
    
    print("📥 يتم تحميل ملف Excel بمحتوى:", df)
    
    return send_file(EXCEL_FILE, as_attachment=True)

@app.route("/delete_student/<student_id>")
def delete_student(student_id):
    ensure_excel_file()
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

    df = df[df["ID"].astype(str) != str(student_id)]
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

    return redirect(url_for("students"))

@app.route("/delete_all_students")
def delete_all_students():
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
        print("🗑️ تم حذف جميع الطلاب وإعادة إنشاء الملف")
    ensure_excel_file()
    return redirect(url_for("students"))

if __name__ == "__main__":
    socketio.run(app, host="0.0.0.0", port=5000, debug=True)
