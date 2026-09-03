import os
import io
from datetime import datetime

import psycopg
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.page import PageMargins

app = Flask(__name__)

# Supabase PostgreSQL connection is taken from Render's DATABASE_URL.
DATABASE_URL = os.environ.get("DATABASE_URL")

FIELDS = [
    "company_name", "company_location", "contact_person", "designation",
    "mobile_number", "email_id", "domain", "industry"
]

INDUSTRIES = [
    "IT",
    "Electrical & Electronics",
    "Manufacturing",
    "Construction",
    "Fashion",
    "Others",
]


def get_connection():
    if not DATABASE_URL:
        raise RuntimeError(
            "DATABASE_URL is not configured. Add the Supabase PostgreSQL "
            "connection string in the hosting service Environment settings."
        )
    return psycopg.connect(DATABASE_URL)


def get_table_name(conn):
    """
    The Supabase table was created as 'Companies'. PostgreSQL can store
    that name either quoted as Companies or unquoted as companies, so
    detect the actual table and use the correct identifier.
    """
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT
                to_regclass('public."Companies"'),
                to_regclass('public.companies')
            """
        )
        quoted_name, lower_name = cur.fetchone()

    if quoted_name is not None:
        return '"Companies"'
    if lower_name is not None:
        return '"companies"'

    raise RuntimeError(
        "Supabase table 'Companies' was not found. "
        "Please check the table name in Supabase."
    )


def row_to_dict(row):
    return {
        "id": row[0],
        "company_name": row[1] or "",
        "company_location": row[2] or "",
        "contact_person": row[3] or "",
        "designation": row[4] or "",
        "mobile_number": row[5] or "",
        "email_id": row[6] or "",
        "domain": row[7] or "",
        "industry": row[8] or "",
        "year": row[9],
        "is_archived": bool(row[10]),
    }


@app.route("/")
def index():
    return render_template("index.html", industries=INDUSTRIES)


@app.get("/api/companies")
def get_companies():
    try:
        with get_connection() as conn:
            table = get_table_name(conn)
            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT id, company_name, company_location,
                           contact_person, designation, mobile_number,
                           email_id, domain, industry, year, is_archived
                    FROM {table}
                    WHERE COALESCE(is_archived, FALSE) = FALSE
                    ORDER BY id DESC
                    """
                )
                rows = cur.fetchall()

        return jsonify([row_to_dict(row) for row in rows])

    except Exception as exc:
        app.logger.exception("Error loading companies")
        return jsonify({"error": f"Database error: {exc}"}), 500


@app.post("/api/companies")
def add_company():
    try:
        payload = request.get_json(force=True) or {}
        row = {f: str(payload.get(f, "")).strip() for f in FIELDS}

        if not row["company_name"]:
            return jsonify({"error": "Company Name is required"}), 400

        current_year = datetime.now().year

        with get_connection() as conn:
            table = get_table_name(conn)

            # Small multi-user lock so two people cannot receive the same ID.
            with conn.cursor() as cur:
                cur.execute("SELECT pg_advisory_xact_lock(7312026)")
                cur.execute(f"SELECT COALESCE(MAX(id), 0) + 1 FROM {table}")
                new_id = cur.fetchone()[0]

                cur.execute(
                    f"""
                    INSERT INTO {table}
                    (id, company_name, company_location, contact_person,
                     designation, mobile_number, email_id, domain,
                     industry, year, is_archived)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, FALSE)
                    """,
                    (
                        new_id,
                        row["company_name"],
                        row["company_location"],
                        row["contact_person"],
                        row["designation"],
                        row["mobile_number"],
                        row["email_id"],
                        row["domain"],
                        row["industry"],
                        current_year,
                    ),
                )

        row["id"] = new_id
        row["year"] = current_year
        row["is_archived"] = False
        return jsonify(row), 201

    except Exception as exc:
        app.logger.exception("Error adding company")
        return jsonify({"error": f"Database error: {exc}"}), 500


@app.put("/api/companies/<int:cid>")
def update_company(cid):
    try:
        payload = request.get_json(force=True) or {}
        updated = {f: str(payload.get(f, "")).strip() for f in FIELDS}

        if not updated["company_name"]:
            return jsonify({"error": "Company Name is required"}), 400

        with get_connection() as conn:
            table = get_table_name(conn)

            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    UPDATE {table}
                    SET company_name=%s,
                        company_location=%s,
                        contact_person=%s,
                        designation=%s,
                        mobile_number=%s,
                        email_id=%s,
                        domain=%s,
                        industry=%s
                    WHERE id=%s
                      AND COALESCE(is_archived, FALSE)=FALSE
                    RETURNING id, company_name, company_location,
                              contact_person, designation, mobile_number,
                              email_id, domain, industry, year, is_archived
                    """,
                    (
                        updated["company_name"],
                        updated["company_location"],
                        updated["contact_person"],
                        updated["designation"],
                        updated["mobile_number"],
                        updated["email_id"],
                        updated["domain"],
                        updated["industry"],
                        cid,
                    ),
                )
                row = cur.fetchone()

        if row is None:
            return jsonify({"error": "Company not found"}), 404

        return jsonify(row_to_dict(row))

    except Exception as exc:
        app.logger.exception("Error updating company")
        return jsonify({"error": f"Database error: {exc}"}), 500


@app.delete("/api/companies/<int:cid>")
def delete_company(cid):
    """
    IMPORTANT:
    This does NOT permanently delete the company.
    It only marks the record as archived, so the data remains in Supabase.
    """
    try:
        with get_connection() as conn:
            table = get_table_name(conn)

            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    UPDATE {table}
                    SET is_archived=TRUE
                    WHERE id=%s
                      AND COALESCE(is_archived, FALSE)=FALSE
                    RETURNING id
                    """,
                    (cid,),
                )
                row = cur.fetchone()

        if row is None:
            return jsonify({"error": "Company not found"}), 404

        return jsonify({"ok": True, "archived": True})

    except Exception as exc:
        app.logger.exception("Error archiving company")
        return jsonify({"error": f"Database error: {exc}"}), 500


@app.get("/export")
def export_xlsx():
    try:
        with get_connection() as conn:
            table = get_table_name(conn)

            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT company_name, company_location,
                           contact_person, designation, mobile_number,
                           email_id, domain, industry
                    FROM {table}
                    WHERE COALESCE(is_archived, FALSE)=FALSE
                    ORDER BY id ASC
                    """
                )
                data = cur.fetchall()

        wb = Workbook()
        ws = wb.active
        ws.title = "Company Database"

        headers = [
            "S.No", "Company Name", "Company Location", "Contact Person",
            "Designation", "Mobile Number", "Email ID", "Domain", "Industry"
        ]
        ws.append(headers)

        for i, row in enumerate(data, 1):
            ws.append([i, *row])

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        widths = [8, 28, 24, 24, 22, 18, 32, 22, 25]
        for idx, width in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + idx)].width = width

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(
            left=0.25, right=0.25, top=0.5, bottom=0.5,
            header=0.2, footer=0.2
        )
        ws.print_title_rows = "1:1"

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="Company_Database_V1.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception as exc:
        app.logger.exception("Error exporting Excel")
        return jsonify({"error": f"Database error: {exc}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
