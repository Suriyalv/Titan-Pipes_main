from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
import os

app = Flask(__name__)
CORS(app)

EXCEL_FILE = "contacts.xlsx"
HTML_FILE = "contact1.html"

# Create Excel file and headers if not exists
def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "ContactForm"
        sheet.append(["Name", "Phone Number", "Company Name", "Email", "Message"])
        wb.save(EXCEL_FILE)

# Serve contact.html directly from same directory
@app.route("/")
def contact_page():
    return send_file(HTML_FILE)

@app.route("/submit-form", methods=["POST"])
def submit_form():
    try:
        data = request.get_json()
        name = data.get("name", "")
        phone = data.get("phone", "")
        company = data.get("company", "")
        email = data.get("email", "")
        message = data.get("message", "")

        if not name or not phone:
            return jsonify({"error": "Name and Phone Number are required."}), 400

        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
        sheet.append([name, phone, company, email, message])
        wb.save(EXCEL_FILE)

        return jsonify({"message": "Form submitted successfully!"}), 200

    except Exception as e:
        print("Error:", e)
        return jsonify({"error": "Failed to process the form."}), 500

if __name__ == "__main__":
    initialize_excel()
    app.run(debug=True)
