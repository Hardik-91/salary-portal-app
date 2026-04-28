import os
import json
import zipfile
import re
from io import BytesIO
from datetime import timedelta, date, datetime

import pandas as pd
import requests
from flask import Flask, render_template, request, redirect, session, send_file
from urllib.parse import quote
from supabase import create_client, Client
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change_this_secret_key")
app.permanent_session_lifetime = timedelta(days=7)

# Local fallback values are kept so your PC test works.
# For online deployment, set these as environment variables on Render.
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "salary-slips")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("SUPABASE_URL and SUPABASE_KEY must be set.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

BACKUP_DIR = os.path.join(os.path.dirname(__file__), "backups")
BACKUP_TABLES = ["employees", "salary_slips", "salary_certificates"]

APP_DIR = os.path.dirname(__file__)
COMPANY_SETTINGS_FILE = os.path.join(APP_DIR, "company_settings.json")
COMPANY_LOGO_STORAGE_PATH = "company/company_logo"
DEFAULT_COMPANY_SETTINGS = {
    "company_name": "NRICH SKYOTEL",
    "logo_url": "",
    "logo_storage_path": "",
}


def get_logo_signed_url(storage_path):
    """Return a temporary Supabase signed URL for the company logo."""
    if not storage_path:
        return ""
    try:
        signed = supabase.storage.from_(SUPABASE_BUCKET).create_signed_url(storage_path, 60 * 60 * 24)
        return signed.get("signedURL") or signed.get("signedUrl") or ""
    except Exception:
        return ""


def load_company_settings():
    settings = DEFAULT_COMPANY_SETTINGS.copy()
    try:
        if os.path.exists(COMPANY_SETTINGS_FILE):
            with open(COMPANY_SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            settings.update(data)
    except Exception:
        pass

    # For online deployment, the logo is stored permanently in Supabase Storage.
    # We generate a fresh signed URL every time the dashboard loads.
    if settings.get("logo_storage_path"):
        settings["logo_url"] = get_logo_signed_url(settings["logo_storage_path"])

    return settings


def save_company_settings(company_name=None, logo_file=None):
    settings = load_company_settings()

    if company_name:
        settings["company_name"] = company_name.strip()

    if logo_file and logo_file.filename:
        content_type = logo_file.mimetype or "image/png"

        # Keep a fixed path so replacing the logo is simple and no old logo files collect.
        # The logo file is stored in Supabase Storage, not in the local static/uploads folder.
        try:
            supabase.storage.from_(SUPABASE_BUCKET).remove([COMPANY_LOGO_STORAGE_PATH])
        except Exception:
            pass

        supabase.storage.from_(SUPABASE_BUCKET).upload(
            COMPANY_LOGO_STORAGE_PATH,
            logo_file.read(),
            {"content-type": content_type},
        )

        settings["logo_storage_path"] = COMPANY_LOGO_STORAGE_PATH
        settings["logo_url"] = get_logo_signed_url(COMPANY_LOGO_STORAGE_PATH)

    with open(COMPANY_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)

    return settings


def status_column_missing_error(error):
    text = str(error)
    return "status" in text and ("PGRST204" in text or "schema cache" in text or "Could not find" in text)



def _safe_table_data(table_name):
    try:
        return supabase.table(table_name).select("*").execute().data or []
    except Exception:
        return []


def create_backup_file():
    """Create a local ZIP backup with JSON + CSV copies of important Supabase tables."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"salary_portal_backup_{timestamp}.zip"
    backup_path = os.path.join(BACKUP_DIR, backup_name)

    snapshot = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "tables": {},
    }

    for table in BACKUP_TABLES:
        rows = _safe_table_data(table)
        snapshot["tables"][table] = rows

    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as backup_zip:
        backup_zip.writestr("backup.json", json.dumps(snapshot, indent=2, default=str))
        for table, rows in snapshot["tables"].items():
            if rows:
                df = pd.DataFrame(rows)
            else:
                df = pd.DataFrame()
            backup_zip.writestr(f"{table}.csv", df.to_csv(index=False))
        backup_zip.writestr(
            "README.txt",
            "Salary Portal Backup\n"
            "This ZIP contains backup.json and CSV exports for employees, salary slips, and salary certificates.\n"
            "Keep this file safely. Do not share it publicly because it can contain employee data.\n",
        )

    return backup_path


def latest_backup_info():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backups = sorted(
        [p for p in os.listdir(BACKUP_DIR) if p.endswith(".zip")],
        reverse=True,
    )
    if not backups:
        return {"exists": False, "filename": "No backup created yet", "created_at": "-"}

    latest = os.path.join(BACKUP_DIR, backups[0])
    created_at = datetime.fromtimestamp(os.path.getmtime(latest)).strftime("%d-%m-%Y %I:%M %p")
    return {"exists": True, "filename": backups[0], "created_at": created_at}


def ensure_daily_backup():
    """Automatically create one backup per day when admin opens dashboard."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    today_prefix = f"salary_portal_backup_{datetime.now().strftime('%Y%m%d')}"
    already_done = any(name.startswith(today_prefix) and name.endswith(".zip") for name in os.listdir(BACKUP_DIR))
    if not already_done:
        create_backup_file()


MONTHS = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}


def get_employee_by_email(email):
    res = supabase.table("employees").select("*").eq("email", email).execute()
    return res.data[0] if res.data else None


def get_admin_by_email(email):
    res = supabase.table("admins").select("*").eq("email", email).execute()
    return res.data[0] if res.data else None


def password_match(stored, entered):
    if stored == entered:
        return True
    try:
        return check_password_hash(stored, entered)
    except Exception:
        return False


def get_financial_year(month_text):
    month_num, year_num = parse_month_year(month_text)

    if not month_num or not year_num:
        return "Unknown"

    if month_num >= 4:
        return f"FY {year_num}-{str(year_num + 1)[-2:]}"
    return f"FY {year_num - 1}-{str(year_num)[-2:]}"


def parse_month_year(month_text):
    """
    Supports values like:
    April 2026, Apr-2026, 2026-04, 04-2026
    """
    text = (month_text or "").strip().lower().replace("/", "-").replace("_", "-")

    if not text:
        return None, None

    # HTML month input format: 2026-04
    parts = text.replace("-", " ").split()

    month_num = None
    year_num = None

    for word in parts:
        if word in MONTHS:
            month_num = MONTHS[word]
        elif word.isdigit():
            if len(word) == 4:
                year_num = int(word)
            elif len(word) in (1, 2):
                num = int(word)
                if 1 <= num <= 12:
                    month_num = num

    return month_num, year_num


def month_key(month_text):
    month_num, year_num = parse_month_year(month_text)
    if not month_num or not year_num:
        return None
    return year_num * 100 + month_num


def month_year_key(month_name, year_text):
    month = (month_name or "").strip().lower()
    year = (year_text or "").strip()

    if month not in MONTHS or not year.isdigit():
        return None

    year_num = int(year)
    month_num = MONTHS[month]
    return year_num * 100 + month_num


def latest_slip_sort_key(slip):
    """
    Pick the most recently uploaded slip.
    Priority: created_at/uploaded_at if present, otherwise highest id.
    """
    created_value = slip.get("created_at") or slip.get("uploaded_at") or ""
    slip_id = slip.get("id") or 0
    return (str(created_value), int(slip_id) if str(slip_id).isdigit() else 0)


def current_financial_year():
    today = date.today()
    if today.month >= 4:
        return f"FY {today.year}-{str(today.year + 1)[-2:]}"
    return f"FY {today.year - 1}-{str(today.year)[-2:]}"


def previous_financial_year():
    current = current_financial_year()
    start_year = int(current.split()[1].split("-")[0])
    return f"FY {start_year - 1}-{str(start_year)[-2:]}"


def safe_get_salary_certificates(employee_email):
    """
    This expects a Supabase table named salary_certificates with columns:
    id, employee_email, financial_year, filename

    If the table is not created yet, dashboard will still work and simply show no certificate.
    """
    try:
        return supabase.table("salary_certificates").select("*").eq(
            "employee_email", employee_email
        ).execute().data or []
    except Exception:
        return []


@app.route("/", methods=["GET", "POST"])
def login():
    message = ""

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "").strip()

        user = get_employee_by_email(email)

        if user and password_match(user["password"], password):
            if user.get("status") in ["Inactive", "Blocked"]:
                message = "Your account is inactive. Please contact admin."
                return render_template("login.html", message=message)

            if request.form.get("remember"):
                session.permanent = True

            session["employee_email"] = user["email"]
            session["employee_name"] = user["name"]
            return redirect("/dashboard")

        message = "Invalid email or password"

    return render_template("login.html", message=message)


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    message = ""

    if request.method == "POST":
        message = "Password reset request received. Please contact admin to reset your password."

    return render_template("forgot_password.html", message=message)


@app.route("/admin-forgot-password", methods=["GET", "POST"])
def admin_forgot_password():
    message = ""

    if request.method == "POST":
        message = "Please contact the system owner to reset admin password."

    return render_template("admin_forgot_password.html", message=message)


@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if "employee_email" not in session:
        return redirect("/")

    message = ""
    employee = get_employee_by_email(session["employee_email"])

    if not employee:
        session.clear()
        return redirect("/")

    if request.method == "POST":
        action = request.form.get("action")

        if action == "change_own_password":
            current_password = request.form.get("current_password", "").strip()
            new_password = request.form.get("new_password", "").strip()
            confirm_password = request.form.get("confirm_password", "").strip()

            if not password_match(employee["password"], current_password):
                message = "Current password is incorrect"
            elif not new_password:
                message = "New password cannot be empty"
            elif new_password != confirm_password:
                message = "New password and re-enter password do not match"
            else:
                hashed = generate_password_hash(new_password)
                supabase.table("employees").update(
                    {"password": hashed}
                ).eq("email", session["employee_email"]).execute()

                message = "Password updated successfully"

    # If employee opens dashboard normally, show the LAST UPLOADED salary slip by default.
    # When employee selects a filter and clicks Show Salary Slips, then apply that filter.
    period_type = request.args.get("period_type")
    is_filter_applied = "period_type" in request.args
    if not period_type:
        period_type = "current"

    from_month = request.args.get("from_month", "")
    from_year = request.args.get("from_year", "")
    to_month = request.args.get("to_month", "")
    to_year = request.args.get("to_year", "")
    certificate_fy = request.args.get("certificate_fy", "")

    all_slips = supabase.table("salary_slips").select("*").eq(
        "employee_email", session["employee_email"]
    ).execute().data or []

    for slip in all_slips:
        slip["display_name"] = os.path.basename(slip.get("filename", ""))
        slip["financial_year"] = get_financial_year(slip.get("month", ""))
        slip["month_key"] = month_key(slip.get("month", ""))

    current_fy = current_financial_year()
    previous_fy = previous_financial_year()

    if not is_filter_applied:
        sorted_all_slips = sorted(all_slips, key=latest_slip_sort_key, reverse=True)
        slips = sorted_all_slips[:1]
    elif period_type == "previous":
        slips = [slip for slip in all_slips if slip["financial_year"] == previous_fy]
    elif period_type == "specific":
        start_key = month_year_key(from_month, from_year)
        end_key = month_year_key(to_month, to_year)

        if start_key and end_key:
            if start_key > end_key:
                start_key, end_key = end_key, start_key

            slips = [
                slip for slip in all_slips
                if slip.get("month_key") and start_key <= slip["month_key"] <= end_key
            ]
        else:
            slips = []
    else:
        period_type = "current"
        slips = [slip for slip in all_slips if slip["financial_year"] == current_fy]

    slips = sorted(slips, key=lambda x: x.get("month_key") or 0, reverse=True)

    financial_years = sorted(
        list(set(slip["financial_year"] for slip in all_slips if slip["financial_year"] != "Unknown")),
        reverse=True
    )

    # Keep current and previous FY in dropdown even if no slips exist yet.
    for fy in [current_fy, previous_fy]:
        if fy not in financial_years:
            financial_years.append(fy)
    financial_years = sorted(financial_years, reverse=True)

    current_year = date.today().year
    years = list(range(current_year + 1, current_year - 10, -1))

    salary_certificate = None
    certificates = safe_get_salary_certificates(session["employee_email"])

    for cert in certificates:
        cert["display_name"] = os.path.basename(cert.get("filename", ""))
        if certificate_fy and cert.get("financial_year") == certificate_fy:
            salary_certificate = cert
            break

    return render_template(
        "dashboard.html",
        name=employee.get("name", ""),
        email=employee.get("email", ""),
        employee_id=employee.get("employee_id", ""),
        mobile=employee.get("mobile", ""),
        slips=slips,
        financial_years=financial_years,
        period_type=period_type,
        from_month=from_month,
        from_year=from_year,
        to_month=to_month,
        to_year=to_year,
        years=years,
        certificate_fy=certificate_fy,
        salary_certificate=salary_certificate,
        message=message,
        company_settings=load_company_settings(),
    )


@app.route("/view/<int:id>")
def view(id):
    if "employee_email" not in session and "admin_email" not in session:
        return redirect("/")

    res = supabase.table("salary_slips").select("*").eq("id", id).execute()

    if not res.data:
        return "Salary slip not found"

    slip = res.data[0]

    if "employee_email" in session and slip["employee_email"] != session["employee_email"]:
        return "Access Denied"

    url = supabase.storage.from_(SUPABASE_BUCKET).create_signed_url(
        slip["filename"], 300
    )

    return redirect(url["signedURL"])


@app.route("/download/<int:id>")
def download(id):
    if "employee_email" not in session and "admin_email" not in session:
        return redirect("/")

    res = supabase.table("salary_slips").select("*").eq("id", id).execute()

    if not res.data:
        return "Salary slip not found"

    slip = res.data[0]

    if "employee_email" in session and slip["employee_email"] != session["employee_email"]:
        return "Access Denied"

    url = supabase.storage.from_(SUPABASE_BUCKET).create_signed_url(
        slip["filename"], 300
    )

    r = requests.get(url["signedURL"])
    return send_file(BytesIO(r.content), download_name=os.path.basename(slip["filename"]), as_attachment=True)


@app.route("/view_certificate/<int:id>")
def view_certificate(id):
    if "employee_email" not in session and "admin_email" not in session:
        return redirect("/")

    try:
        res = supabase.table("salary_certificates").select("*").eq("id", id).execute()
    except Exception:
        return "Salary certificate table not found"

    if not res.data:
        return "Salary certificate not found"

    cert = res.data[0]

    if "employee_email" in session and cert["employee_email"] != session["employee_email"]:
        return "Access Denied"

    url = supabase.storage.from_(SUPABASE_BUCKET).create_signed_url(
        cert["filename"], 300
    )

    return redirect(url["signedURL"])


@app.route("/download_certificate/<int:id>")
def download_certificate(id):
    if "employee_email" not in session and "admin_email" not in session:
        return redirect("/")

    try:
        res = supabase.table("salary_certificates").select("*").eq("id", id).execute()
    except Exception:
        return "Salary certificate table not found"

    if not res.data:
        return "Salary certificate not found"

    cert = res.data[0]

    if "employee_email" in session and cert["employee_email"] != session["employee_email"]:
        return "Access Denied"

    url = supabase.storage.from_(SUPABASE_BUCKET).create_signed_url(
        cert["filename"], 300
    )

    r = requests.get(url["signedURL"])
    return send_file(BytesIO(r.content), download_name=os.path.basename(cert["filename"]), as_attachment=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect("/admin")


@app.route("/admin", methods=["GET", "POST"])
def admin():
    message = ""

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "").strip()

        admin_user = get_admin_by_email(email)

        if admin_user and password_match(admin_user["password"], password):
            if request.form.get("remember"):
                session.permanent = True

            session["admin_email"] = admin_user["email"]
            return redirect("/admin/dashboard")

        message = "Invalid email or password"

    return render_template("admin_login.html", message=message)


@app.route("/admin/backup/download")
def download_backup():
    if "admin_email" not in session:
        return redirect("/admin")

    backup_path = create_backup_file()
    return send_file(
        backup_path,
        as_attachment=True,
        download_name=os.path.basename(backup_path),
        mimetype="application/zip",
    )


@app.route("/admin/dashboard", methods=["GET", "POST"])
def admin_dashboard():
    if "admin_email" not in session:
        return redirect("/admin")

    try:
        ensure_daily_backup()
    except Exception:
        pass

    message = request.args.get("msg", "")
    active_tab = request.form.get("tab") or request.args.get("tab") or "dashboard"
    if active_tab in ["employees", "excel"]:
        active_tab = "employee_entry"
    if active_tab in ["salary", "certificates"]:
        active_tab = "salary_upload"

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add_employee":
            name = request.form.get("name", "").strip()
            employee_id = request.form.get("employee_id", "").strip()
            mobile = request.form.get("mobile", "").strip()
            email = request.form.get("email", "").strip()
            password = request.form.get("password", "").strip()

            if name and email and password:
                existing = get_employee_by_email(email)

                if existing:
                    message = "Employee already exists"
                else:
                    hashed = generate_password_hash(password)
                    supabase.table("employees").insert(
                        {
                            "name": name,
                            "employee_id": employee_id,
                            "mobile": mobile,
                            "email": email,
                            "password": hashed,
                        }
                    ).execute()
                    message = "Employee added successfully"
            else:
                message = "Please fill full name, email and password"

        elif action == "upload_employee_excel":
            excel_file = request.files.get("employee_excel")

            if not excel_file or not excel_file.filename:
                message = "Please select an Excel file"
            else:
                try:
                    df = pd.read_excel(excel_file)

                    required_columns = [
                        "Full Name",
                        "Employee ID",
                        "Mobile Number",
                        "Email ID",
                        "Password",
                    ]

                    missing = [col for col in required_columns if col not in df.columns]

                    if missing:
                        message = "Missing columns: " + ", ".join(missing)
                    else:
                        added_count = 0
                        skipped_count = 0

                        for _, row in df.iterrows():
                            name = str(row["Full Name"]).strip()
                            employee_id = str(row["Employee ID"]).strip()
                            mobile = str(row["Mobile Number"]).strip()
                            email = str(row["Email ID"]).strip()
                            password = str(row["Password"]).strip()

                            if not name or not email or not password or email.lower() == "nan":
                                skipped_count += 1
                                continue

                            existing = get_employee_by_email(email)

                            if existing:
                                skipped_count += 1
                                continue

                            hashed = generate_password_hash(password)

                            supabase.table("employees").insert(
                                {
                                    "name": name,
                                    "employee_id": employee_id if employee_id.lower() != "nan" else "",
                                    "mobile": mobile if mobile.lower() != "nan" else "",
                                    "email": email,
                                    "password": hashed,
                                }
                            ).execute()

                            added_count += 1

                        message = f"Excel upload complete. Added: {added_count}, Skipped: {skipped_count}"

                except Exception as e:
                    message = f"Excel upload failed: {str(e)}"

        elif action == "change_employee_password":
            email = request.form.get("employee_email", "").strip()
            new_password = request.form.get("new_password", "").strip()

            if email and new_password:
                hashed = generate_password_hash(new_password)

                supabase.table("employees").update(
                    {"password": hashed}
                ).eq("email", email).execute()

                message = "Employee password updated successfully"
            else:
                message = "Please select employee and enter new password"

        elif action == "delete_employee":
            email = request.form.get("employee_email", "").strip()

            if email:
                slips = supabase.table("salary_slips").select("*").eq(
                    "employee_email", email
                ).execute().data or []

                for slip in slips:
                    try:
                        supabase.storage.from_(SUPABASE_BUCKET).remove([slip["filename"]])
                    except Exception:
                        pass

                supabase.table("salary_slips").delete().eq(
                    "employee_email", email
                ).execute()

                supabase.table("employees").delete().eq("email", email).execute()

                message = "Employee deleted successfully"
            else:
                message = "Please select employee"

        elif action == "upload_slip":
            employee_email = request.form.get("employee_email", "").strip()
            month = request.form.get("month", "").strip()
            pdf_file = request.files.get("pdf_file")

            if employee_email and month and pdf_file and pdf_file.filename:
                filename = pdf_file.filename.replace(" ", "_")
                storage_path = f"{employee_email}/{filename}"

                try:
                    supabase.storage.from_(SUPABASE_BUCKET).upload(
                        storage_path,
                        pdf_file.read(),
                        {"content-type": "application/pdf"},
                    )

                    supabase.table("salary_slips").insert(
                        {
                            "employee_email": employee_email,
                            "month": month,
                            "filename": storage_path,
                        }
                    ).execute()

                    message = "Salary slip uploaded successfully"

                except Exception as e:
                    message = f"Salary slip upload failed: {str(e)}"
            else:
                message = "Please fill all salary slip details"

        elif action == "delete_slip":
            slip_id = request.form.get("slip_id", "").strip()

            if slip_id:
                res = supabase.table("salary_slips").select("*").eq(
                    "id", int(slip_id)
                ).execute()

                if res.data:
                    slip = res.data[0]

                    try:
                        supabase.storage.from_(SUPABASE_BUCKET).remove([slip["filename"]])
                    except Exception:
                        pass

                    supabase.table("salary_slips").delete().eq(
                        "id", int(slip_id)
                    ).execute()

                    message = "Salary slip deleted successfully"
                else:
                    message = "Salary slip not found"

        elif action == "upload_certificate":
            employee_email = request.form.get("employee_email", "").strip()
            financial_year = request.form.get("financial_year", "").strip()
            certificate_file = request.files.get("certificate_file")

            if employee_email and financial_year and certificate_file and certificate_file.filename:
                filename = certificate_file.filename.replace(" ", "_")
                storage_path = f"{employee_email}/certificates/{financial_year.replace(' ', '_')}_{filename}"

                try:
                    supabase.storage.from_(SUPABASE_BUCKET).upload(
                        storage_path,
                        certificate_file.read(),
                        {"content-type": "application/pdf"},
                    )

                    # Remove old certificate for same employee + same financial year from DB/storage.
                    old_certs = supabase.table("salary_certificates").select("*").eq(
                        "employee_email", employee_email
                    ).eq("financial_year", financial_year).execute().data or []

                    for old_cert in old_certs:
                        try:
                            supabase.storage.from_(SUPABASE_BUCKET).remove([old_cert["filename"]])
                        except Exception:
                            pass
                        supabase.table("salary_certificates").delete().eq("id", old_cert["id"]).execute()

                    supabase.table("salary_certificates").insert(
                        {
                            "employee_email": employee_email,
                            "financial_year": financial_year,
                            "filename": storage_path,
                        }
                    ).execute()

                    message = "Salary certificate uploaded successfully"

                except Exception as e:
                    message = f"Salary certificate upload failed: {str(e)}"
            else:
                message = "Please fill all salary certificate details"

        elif action == "delete_certificate":
            certificate_id = request.form.get("certificate_id", "").strip()

            if certificate_id:
                try:
                    res = supabase.table("salary_certificates").select("*").eq(
                        "id", int(certificate_id)
                    ).execute()
                except Exception as e:
                    res = None
                    message = f"Salary certificate table error: {str(e)}"

                if res and res.data:
                    cert = res.data[0]

                    try:
                        supabase.storage.from_(SUPABASE_BUCKET).remove([cert["filename"]])
                    except Exception:
                        pass

                    supabase.table("salary_certificates").delete().eq(
                        "id", int(certificate_id)
                    ).execute()

                    message = "Salary certificate deleted successfully"
                elif not message:
                    message = "Salary certificate not found"
            else:
                message = "Please select salary certificate"

        elif action == "block_employee":
            email = request.form.get("employee_email", "").strip()
            if email:
                try:
                    supabase.table("employees").update({"status": "Inactive"}).eq("email", email).execute()
                    message = "Employee access blocked successfully"
                except Exception as e:
                    if status_column_missing_error(e):
                        message = "Block employee failed: Please add the status column in Supabase first. Run the SQL file included in this ZIP."
                    else:
                        message = f"Block employee failed: {str(e)}"
            else:
                message = "Please select employee"

        elif action == "unblock_employee":
            email = request.form.get("employee_email", "").strip()
            if email:
                try:
                    supabase.table("employees").update({"status": "Active"}).eq("email", email).execute()
                    message = "Employee access unblocked successfully"
                except Exception as e:
                    if status_column_missing_error(e):
                        message = "Unblock employee failed: Please add the status column in Supabase first. Run the SQL file included in this ZIP."
                    else:
                        message = f"Unblock employee failed: {str(e)}"
            else:
                message = "Please select employee"

        elif action == "manual_backup":
            try:
                create_backup_file()
                message = "Backup created successfully"
            except Exception as e:
                message = f"Backup failed: {str(e)}"

        elif action == "change_company":
            company_name = request.form.get("company_name", "").strip()
            company_logo = request.files.get("company_logo")

            if not company_name and (not company_logo or not company_logo.filename):
                message = "Please enter company name or select company logo"
            else:
                try:
                    save_company_settings(company_name=company_name, logo_file=company_logo)
                    message = "Company name and logo updated successfully"
                except Exception as e:
                    message = f"Company update failed: {str(e)}"

        elif action == "change_admin_email":
            new_admin_email = request.form.get("new_admin_email", "").strip()

            if not new_admin_email:
                message = "Please enter new admin email"
            else:
                try:
                    supabase.table("admins").update({"email": new_admin_email}).eq("email", session["admin_email"]).execute()
                    session["admin_email"] = new_admin_email
                    message = "Admin email updated successfully"
                except Exception as e:
                    message = f"Admin email update failed: {str(e)}"

        elif action == "change_admin_password":
            current_password = request.form.get("current_admin_password", "").strip()
            new_password = request.form.get("new_admin_password", "").strip()
            confirm_password = request.form.get("confirm_admin_password", "").strip()

            admin_user = get_admin_by_email(session["admin_email"])

            if not admin_user:
                message = "Admin account not found"
            elif not password_match(admin_user["password"], current_password):
                message = "Current admin password is incorrect"
            elif not new_password:
                message = "New password cannot be empty"
            elif new_password != confirm_password:
                message = "New password and confirm password do not match"
            else:
                try:
                    hashed = generate_password_hash(new_password)
                    supabase.table("admins").update({"password": hashed}).eq("email", session["admin_email"]).execute()
                    message = "Admin password updated successfully"
                except Exception as e:
                    message = f"Admin password update failed: {str(e)}"

        elif action == "bulk_upload_slip":
            bulk_file = request.files.get("bulk_salary_file")

            if not bulk_file or not bulk_file.filename:
                message = "Please select a ZIP file"
            elif not bulk_file.filename.lower().endswith(".zip"):
                message = "Bulk upload currently supports ZIP files only. Put PDFs inside ZIP and name each PDF like employee@email.com_April_2025.pdf"
            else:
                try:
                    employees_data = supabase.table("employees").select("email, employee_id").execute().data or []
                    employees_lookup = {
                        emp["email"].lower(): {
                            "email": emp["email"],
                            "employee_id": (emp.get("employee_id") or "").strip(),
                        }
                        for emp in employees_data
                        if emp.get("email")
                    }
                    added_count = 0
                    skipped_count = 0
                    errors = []

                    with zipfile.ZipFile(BytesIO(bulk_file.read())) as zf:
                        for item in zf.infolist():
                            if item.is_dir() or not item.filename.lower().endswith(".pdf"):
                                continue

                            base_name = os.path.basename(item.filename).replace(" ", "_")
                            base_lower = base_name.lower()

                            matched_employee = None
                            for emp_lower, emp_info in employees_lookup.items():
                                if emp_lower in base_lower:
                                    matched_employee = emp_info
                                    break

                            if not matched_employee:
                                skipped_count += 1
                                errors.append(f"Skipped {base_name}: employee email not found in file name")
                                continue

                            matched_email = matched_employee["email"]
                            emp_id = matched_employee.get("employee_id") or matched_email.split("@")[0]
                            emp_id = re.sub(r"[^A-Za-z0-9_-]", "", emp_id)

                            month_text = re.sub(re.escape(matched_email), "", os.path.splitext(base_name)[0], flags=re.IGNORECASE)
                            month_text = month_text.replace("_", " ").replace("-", " ").strip() or "Bulk Upload"

                            month_parts = month_text.split()
                            if len(month_parts) >= 2 and month_parts[-1].isdigit():
                                clean_month = f"{month_parts[0].title()}{month_parts[-1]}"
                            else:
                                clean_month = re.sub(r"[^A-Za-z0-9]", "", month_text.title()) or "BulkUpload"

                            final_filename = f"{emp_id}_{clean_month}.pdf"
                            storage_path = f"{matched_email}/{final_filename}"

                            pdf_bytes = zf.read(item)

                            # If same employee/month already exists in storage, remove it first so re-upload works cleanly.
                            try:
                                supabase.storage.from_(SUPABASE_BUCKET).remove([storage_path])
                            except Exception:
                                pass

                            supabase.storage.from_(SUPABASE_BUCKET).upload(
                                storage_path,
                                pdf_bytes,
                                {"content-type": "application/pdf"},
                            )
                            supabase.table("salary_slips").insert(
                                {
                                    "employee_email": matched_email,
                                    "month": month_text,
                                    "filename": storage_path,
                                }
                            ).execute()
                            added_count += 1

                    if added_count:
                        message = f"Bulk upload complete. Added: {added_count}, Skipped: {skipped_count}"
                    else:
                        details = "; ".join(errors[:3])
                        message = f"No salary slips uploaded. Skipped: {skipped_count}. {details}"

                except zipfile.BadZipFile:
                    message = "Bulk upload failed: selected file is not a valid ZIP"
                except Exception as e:
                    message = f"Bulk salary slip upload failed: {str(e)}"

        # Auto-refresh dashboard data after every admin action and show message as toast.
        # This avoids old stats/list data after add/delete/upload actions.
        if message:
            return redirect(f"/admin/dashboard?tab={active_tab}&msg={quote(message)}")

    employees = supabase.table("employees").select("*").execute().data or []
    slips = supabase.table("salary_slips").select("*").execute().data or []

    try:
        certificates = supabase.table("salary_certificates").select("*").execute().data or []
    except Exception:
        certificates = []
        if not message:
            message = "Salary certificate table not found. Please create table salary_certificates."

    for slip in slips:
        slip["display_name"] = os.path.basename(slip.get("filename", ""))

    for cert in certificates:
        cert["display_name"] = os.path.basename(cert.get("filename", ""))

    total_active = len([emp for emp in employees if emp.get("status") not in ["Inactive", "Blocked"]])
    total_inactive = len([emp for emp in employees if emp.get("status") in ["Inactive", "Blocked"]])

    company_settings = load_company_settings()

    return render_template(
        "admin_dashboard.html",
        employees=employees,
        slips=slips,
        certificates=certificates,
        message=message,
        total_employees=len(employees),
        total_active=total_active,
        total_inactive=total_inactive,
        total_slips=len(slips),
        total_certificates=len(certificates),
        backup_info=latest_backup_info(),
        active_tab=active_tab,
        company_settings=company_settings,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050, debug=False, use_reloader=False)
