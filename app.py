from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
from pathlib import Path
import re
import io
import csv
import pandas as pd
import pypdf
import os
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-me")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("COOKIE_SECURE", "0") == "1"

LOGIN_USERNAME = os.environ.get("LOGIN_USERNAME", "admin")
LOGIN_PASSWORD = os.environ.get("LOGIN_PASSWORD", "change-me")

BASE_DIR = Path(__file__).resolve().parent
CNIC_PDF = Path(os.environ.get("CNIC_PDF_PATH", str(BASE_DIR / "cnic_data.pdf")))
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf"}


def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_student_id(value):
    return clean_text(value).upper()


def normalize_copy_number(value):
    # Keep copy/script numbers as text so leading zeroes are never lost.
    return re.sub(r"\D", "", clean_text(value))


def extract_pdf_lines(pdf_path):
    lines = []
    reader = pypdf.PdfReader(str(pdf_path))

    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = clean_text(line)
            if line:
                lines.append(line)

    return lines


def parse_cnic_pdf(pdf_path):
    """
    Expected CNIC PDF row:
    SR.# ID Student Name Phone CNIC

    Example:
    1 0000534827 RUKYAN BAI 03453865253 44303-9959843-6
    """
    records = {}

    # CNIC pattern is intentionally strict enough to avoid treating other numbers
    # in the PDF as CNIC values.
    row_re = re.compile(
        r"^\s*(\d+)\s+([A-Za-z0-9]+)\s+(.*?)\s+"
        r"(?:\d{10,11})\s+(\d{5}-\d{7}-\d)\s*$"
    )

    for line in extract_pdf_lines(pdf_path):
        m = row_re.match(line)
        if not m:
            continue

        _, student_id, student_name, cnic = m.groups()
        student_id = normalize_student_id(student_id)
        records[student_id] = {
            "student_id": student_id,
            "student_name": clean_text(student_name),
            "cnic": cnic,
        }

    return records


def parse_attendance_pdf(pdf_path):
    """
    Expected attendance row:
    Sr Student ID Name Status Script No Course Code

    Example:
    1 0000052718 DANISH S 346126115 AIOU-5410
    """
    records = []

    # Script/copy number is the key anchor. The name is everything between
    # Student ID and Status. This supports names containing spaces.
    row_re = re.compile(
        r"^\s*(\d+)\s+([A-Za-z0-9]+)\s+(.*?)\s+([A-Za-z])\s+"
        r"(\d{6,20})\s+(AIOU-\d+)\s*$",
        re.IGNORECASE,
    )

    for line in extract_pdf_lines(pdf_path):
        m = row_re.match(line)
        if not m:
            continue

        _, student_id, student_name, status, copy_number, course_code = m.groups()

        records.append({
            "copy_number": normalize_copy_number(copy_number),
            "student_id": normalize_student_id(student_id),
            "student_name": clean_text(student_name),
            "course_code": clean_text(course_code).upper(),
            "status": clean_text(status).upper(),
        })

    return records


def parse_copy_numbers(raw):
    """
    Accept copy numbers separated by commas, spaces, new lines, tabs, or semicolons.
    """
    parts = re.split(r"[,;\s]+", raw or "")
    return [normalize_copy_number(x) for x in parts if normalize_copy_number(x)]


def build_result(attendance_records, cnic_records, requested_copy_numbers):
    selected = set(requested_copy_numbers)
    results = []
    missing_cnic = []
    not_found = []

    attendance_by_copy = {}
    for row in attendance_records:
        attendance_by_copy.setdefault(row["copy_number"], []).append(row)

    for copy_number in requested_copy_numbers:
        matches = attendance_by_copy.get(copy_number, [])

        if not matches:
            not_found.append(copy_number)
            continue

        for row in matches:
            cnic_info = cnic_records.get(row["student_id"])
            cnic = cnic_info["cnic"] if cnic_info else ""

            if not cnic:
                missing_cnic.append(row["student_id"])

            results.append({
                "Copy Number": row["copy_number"],
                "Student ID": row["student_id"],
                "Student Name": row["student_name"],
                "CNIC": cnic,
                "Course Code": row["course_code"],
            })

    return results, sorted(set(missing_cnic)), not_found


def make_excel_bytes(rows):
    columns = ["Copy Number", "Student ID", "Student Name", "CNIC", "Course Code"]
    df = pd.DataFrame(rows, columns=columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Attendance")

        ws = writer.book["Attendance"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Keep IDs/CNIC/copy numbers as text in Excel.
        for col in ["A", "B", "D"]:
            for cell in ws[col][1:]:
                cell.number_format = "@"

        widths = {
            "A": 18,
            "B": 18,
            "C": 28,
            "D": 20,
            "E": 18,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

    output.seek(0)
    return output


def make_tsv_bytes(rows):
    columns = ["Copy Number", "Student ID", "Student Name", "CNIC", "Course Code"]
    output = io.StringIO(newline="")
    writer = csv.writer(output, delimiter="\t", lineterminator="\n")
    writer.writerow(columns)

    for row in rows:
        writer.writerow([row.get(c, "") for c in columns])

    return io.BytesIO(output.getvalue().encode("utf-8-sig"))



def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("logged_in"):
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if username == LOGIN_USERNAME and password == LOGIN_PASSWORD:
            session.clear()
            session["logged_in"] = True
            session["username"] = username
            next_url = request.args.get("next") or url_for("index")
            if not next_url.startswith("/") or next_url.startswith("//"):
                next_url = url_for("index")
            return redirect(next_url)
        flash("Invalid username or password.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        attendance_file = request.files.get("attendance_pdf")
        copy_numbers_raw = request.form.get("copy_numbers", "")

        if not attendance_file or not attendance_file.filename:
            flash("Please select the Attendance Report PDF.", "error")
            return render_template("index.html")

        if not attendance_file.filename.lower().endswith(".pdf"):
            flash("Only PDF files are allowed.", "error")
            return render_template("index.html")

        copy_numbers = parse_copy_numbers(copy_numbers_raw)
        if not copy_numbers:
            flash("Please enter at least one copy number.", "error")
            return render_template("index.html")

        if not CNIC_PDF.exists():
            flash("Permanent CNIC PDF is missing. Put cnic_data.pdf beside app.py.", "error")
            return render_template("index.html")

        filename = secure_filename(attendance_file.filename)
        attendance_path = UPLOAD_DIR / filename
        attendance_file.save(attendance_path)

        try:
            cnic_records = parse_cnic_pdf(CNIC_PDF)
            attendance_records = parse_attendance_pdf(attendance_path)

            rows, missing_cnic, not_found = build_result(
                attendance_records,
                cnic_records,
                copy_numbers,
            )

            if not rows:
                flash(
                    "No requested copy numbers were found in the Attendance PDF. "
                    "Check the copy/script numbers and PDF format.",
                    "error",
                )
                return render_template(
                    "index.html",
                    parsed_attendance=len(attendance_records),
                    cnic_students=len(cnic_records),
                )

            excel_bytes = make_excel_bytes(rows)
            tsv_bytes = make_tsv_bytes(rows)

            # Store latest output on server too.
            excel_path = OUTPUT_DIR / "separated_attendance.xlsx"
            txt_path = OUTPUT_DIR / "separated_attendance.txt"
            excel_path.write_bytes(excel_bytes.getvalue())
            txt_path.write_bytes(tsv_bytes.getvalue())

            return render_template(
                "result.html",
                rows=rows,
                count=len(rows),
                missing_cnic=missing_cnic,
                not_found=not_found,
                excel_ready=True,
                txt_ready=True,
            )

        except Exception as exc:
            flash(
                "Could not process the PDF. Make sure it is a text-readable attendance report. "
                f"Technical error: {exc}",
                "error",
            )
            return render_template("index.html")

    return render_template("index.html")


@app.post("/download/excel")
@login_required
def download_excel():
    rows_json = request.form.get("rows_json", "")
    if not rows_json:
        return "No result data.", 400

    import json
    rows = json.loads(rows_json)
    return send_file(
        make_excel_bytes(rows),
        as_attachment=True,
        download_name="separated_attendance.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/download/txt")
@login_required
def download_txt():
    rows_json = request.form.get("rows_json", "")
    if not rows_json:
        return "No result data.", 400

    import json
    rows = json.loads(rows_json)
    return send_file(
        make_tsv_bytes(rows),
        as_attachment=True,
        download_name="separated_attendance.txt",
        mimetype="text/plain; charset=utf-8",
    )


if __name__ == "__main__":
    app.run(debug=True)
