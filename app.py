from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.page import PageMargins
from pathlib import Path
import json, io

app = Flask(__name__)
DATA_FILE = Path("companies.json")

FIELDS = [
    "company_name", "company_location", "contact_person", "designation",
    "mobile_number", "email_id", "domain", "industry"
]
INDUSTRIES = ["IT", "Electrical & Electronics", "Manufacturing", "Construction", "Fashion", "Others"]

def load_data():
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []

def save_data(data):
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

@app.route("/")
def index():
    return render_template("index.html", industries=INDUSTRIES)

@app.get("/api/companies")
def get_companies():
    return jsonify(load_data())

@app.post("/api/companies")
def add_company():
    payload = request.get_json(force=True)
    row = {f: str(payload.get(f, "")).strip() for f in FIELDS}
    if not row["company_name"]:
        return jsonify({"error": "Company Name is required"}), 400
    data = load_data()
    row["id"] = (max([x.get("id", 0) for x in data] or [0]) + 1)
    data.append(row)
    save_data(data)
    return jsonify(row)

@app.put("/api/companies/<int:cid>")
def update_company(cid):
    payload = request.get_json(force=True)
    data = load_data()
    for i, row in enumerate(data):
        if row.get("id") == cid:
            updated = {f: str(payload.get(f, "")).strip() for f in FIELDS}
            if not updated["company_name"]:
                return jsonify({"error": "Company Name is required"}), 400
            updated["id"] = cid
            data[i] = updated
            save_data(data)
            return jsonify(updated)
    return jsonify({"error": "Company not found"}), 404

@app.delete("/api/companies/<int:cid>")
def delete_company(cid):
    data = load_data()
    new_data = [x for x in data if x.get("id") != cid]
    if len(new_data) == len(data):
        return jsonify({"error": "Company not found"}), 404
    save_data(new_data)
    return jsonify({"ok": True})

@app.get("/export")
def export_xlsx():
    data = load_data()
    wb = Workbook()
    ws = wb.active
    ws.title = "Company Database"

    headers = ["S.No", "Company Name", "Company Location", "Contact Person",
               "Designation", "Mobile Number", "Email ID", "Domain", "Industry"]
    ws.append(headers)

    for i, row in enumerate(data, 1):
        ws.append([
            i, row.get("company_name",""), row.get("company_location",""),
            row.get("contact_person",""), row.get("designation",""),
            row.get("mobile_number",""), row.get("email_id",""),
            row.get("domain",""), row.get("industry","")
        ])

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [8, 28, 24, 24, 22, 18, 32, 22, 25]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64+idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    ws.print_title_rows = "1:1"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name="Company_Database_V1.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)