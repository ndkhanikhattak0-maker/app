from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook, load_workbook
from pathlib import Path
from threading import RLock
import io

app = Flask(__name__)
FILE_NAME = Path(__file__).with_name("students_data.xlsx")
LOCK = RLock()

FIELDS = [
    "Student Name", "FatherName", "Gender", "RollNo", "Admission Date",
    "Nationality", "CNIC Number", "Passport Number", "Date of Birth",
    "PhoneNumber", "Email", "Domicile District", "Domicile Province",
    "Mailing Address", "City (Mailing Address)", "SSC Degree Name",
    "SSC Board Name", "SSC Total Marks", "SSC Obtained Marks",
    "HSSC Degree Name", "HSSC Degree Nomenclature", "HSSC Board Name",
    "HSSC Total Marks", "HSSC Obtained Marks"
]

def initialize_excel():
    with LOCK:
        if not FILE_NAME.exists():
            wb = Workbook()
            ws = wb.active
            ws.title = "Student Records"
            ws.append(FIELDS)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            wb.save(FILE_NAME)

def workbook():
    initialize_excel()
    return load_workbook(FILE_NAME)

def clean_record(payload):
    return {field: str(payload.get(field, "") or "").strip() for field in FIELDS}

def validate_record(data, require_name=True):
    if require_name and not data["Student Name"]:
        return "Student Name is required."
    if not data["RollNo"]:
        return "Roll Number is required."
    for total_field, obtained_field in [
        ("SSC Total Marks", "SSC Obtained Marks"),
        ("HSSC Total Marks", "HSSC Obtained Marks")
    ]:
        total, obtained = data[total_field], data[obtained_field]
        if total or obtained:
            try:
                total_num, obtained_num = float(total), float(obtained)
                if total_num <= 0 or obtained_num < 0:
                    return f"Enter valid positive total marks and non-negative obtained marks for {total_field}."
                if obtained_num > total_num:
                    return f"{obtained_field} cannot exceed {total_field}."
            except ValueError:
                return f"Please enter numeric values for {total_field} and {obtained_field}."
    return None

def read_records(ws):
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(value is not None for value in row):
            continue
        record = {}
        for i, field in enumerate(FIELDS):
            value = row[i] if i < len(row) else None
            record[field] = "" if value is None else str(value)
        records.append(record)
    return records

@app.get("/")
def index():
    return render_template("index.html", fields=FIELDS)

@app.get("/api/students")
def list_students():
    with LOCK:
        wb = workbook()
        records = read_records(wb.active)
        wb.close()
    query = request.args.get("q", "").strip().lower()
    if query:
        records = [r for r in records if query in r["RollNo"].lower()
                   or query in r["Student Name"].lower()]
    return jsonify(records)

@app.post("/api/students")
def create_student():
    data = clean_record(request.get_json(silent=True) or {})
    error = validate_record(data)
    if error:
        return jsonify(error=error), 400
    with LOCK:
        wb = workbook()
        ws = wb.active
        if any(str(ws.cell(row, FIELDS.index("RollNo") + 1).value or "").strip() == data["RollNo"]
               for row in range(2, ws.max_row + 1)):
            wb.close()
            return jsonify(error="This Roll Number already exists."), 409
        ws.append([data[field] for field in FIELDS])
        wb.save(FILE_NAME)
        wb.close()
    return jsonify(message="Student record saved successfully.")

@app.put("/api/students/<path:original_roll>")
def update_student(original_roll):
    data = clean_record(request.get_json(silent=True) or {})
    error = validate_record(data)
    if error:
        return jsonify(error=error), 400
    roll_col = FIELDS.index("RollNo") + 1
    with LOCK:
        wb = workbook()
        ws = wb.active
        target = None
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row, roll_col).value or "").strip() == original_roll:
                target = row
                break
        if target is None:
            wb.close()
            return jsonify(error="Student record not found."), 404
        for row in range(2, ws.max_row + 1):
            if row != target and str(ws.cell(row, roll_col).value or "").strip() == data["RollNo"]:
                wb.close()
                return jsonify(error="Another student already uses this Roll Number."), 409
        for col, field in enumerate(FIELDS, start=1):
            ws.cell(target, col).value = data[field]
        wb.save(FILE_NAME)
        wb.close()
    return jsonify(message="Student record updated successfully.")

@app.delete("/api/students/<path:roll_no>")
def delete_student(roll_no):
    roll_col = FIELDS.index("RollNo") + 1
    with LOCK:
        wb = workbook()
        ws = wb.active
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row, roll_col).value or "").strip() == roll_no:
                ws.delete_rows(row, 1)
                wb.save(FILE_NAME)
                wb.close()
                return jsonify(message="Student record deleted successfully.")
        wb.close()
    return jsonify(error="Student record not found."), 404

@app.get("/download-excel")
def download_excel():
    initialize_excel()
    return send_file(FILE_NAME, as_attachment=True, download_name="students_data.xlsx")

if __name__ == "__main__":
    initialize_excel()
    app.run(host="127.0.0.1", port=5000, debug=True)
