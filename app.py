import os
import re
import secrets
import sqlite3
from contextlib import closing
from functools import wraps
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024
app.config["DATABASE"] = Path(__file__).with_name("Register.db")
app.config["LEGACY_WORKBOOK"] = Path(__file__).with_name("Register.xlsx")
app.config["ADMIN_PASSWORD"] = os.environ.get("ADMIN_PASSWORD", "")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

HEADERS = ("Овог", "Нэр", "Утасны дугаар", "Имэйл")
EXPORT_HEADERS = ("№", *HEADERS, "Бүртгэсэн огноо")
EMAIL_PATTERN = re.compile(
    r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+"
    r"[A-Za-z]{2,63}"
)


def valid_email(email):
    if len(email) > 254 or not EMAIL_PATTERN.fullmatch(email):
        return False
    local_part = email.split("@", 1)[0]
    return not (local_part.startswith(".") or local_part.endswith(".") or ".." in local_part)


def initialize_database():
    database_path = Path(app.config["DATABASE"])
    database_path.parent.mkdir(parents=True, exist_ok=True)

    with closing(sqlite3.connect(database_path)) as connection:
        table_exists = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'registrations'"
        ).fetchone() is not None
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS registrations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                surname TEXT NOT NULL,
                first_name TEXT NOT NULL,
                phone TEXT NOT NULL,
                email TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        legacy_path = Path(app.config["LEGACY_WORKBOOK"])
        if not table_exists and legacy_path.exists():
            try:
                workbook = load_workbook(legacy_path, read_only=True, data_only=False)
            except (OSError, ValueError, BadZipFile, InvalidFileException):
                app.logger.exception("Could not read the previous registration workbook")
                raise

            try:
                worksheet = workbook.active
                headers = next(
                    worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
                    (),
                )
                if tuple(headers[:len(HEADERS)]) != HEADERS:
                    app.logger.warning("Legacy workbook has an unexpected header; migration skipped.")
                else:
                    registrations = []
                    for row_number, row in enumerate(
                        worksheet.iter_rows(min_row=2, values_only=True),
                        start=2,
                    ):
                        values = row[:len(HEADERS)]
                        if not values or all(value is None for value in values):
                            continue
                        if len(values) != len(HEADERS) or any(value is None for value in values):
                            app.logger.warning(
                                "Skipping incomplete legacy registration row %s.",
                                row_number,
                            )
                            continue
                        registrations.append(tuple(str(value) for value in values))

                    connection.executemany(
                        """
                        INSERT INTO registrations (surname, first_name, phone, email)
                        VALUES (?, ?, ?, ?)
                        """,
                        registrations,
                    )
                    app.logger.info(
                        "Imported %s registration(s) from the previous workbook.",
                        len(registrations),
                    )
            finally:
                workbook.close()

        connection.commit()


def save_registration(surname, first_name, phone, email):
    with closing(sqlite3.connect(app.config["DATABASE"])) as connection:
        connection.execute(
            """
            INSERT INTO registrations (surname, first_name, phone, email)
            VALUES (?, ?, ?, ?)
            """,
            (surname, first_name, phone, email),
        )
        connection.commit()


def get_registrations():
    with closing(sqlite3.connect(app.config["DATABASE"])) as connection:
        connection.row_factory = sqlite3.Row
        return connection.execute(
            """
            SELECT id, surname, first_name, phone, email, created_at
            FROM registrations
            ORDER BY id DESC
            """
        ).fetchall()


def csrf_token():
    token = session.get("_csrf_token")
    if token is None:
        token = secrets.token_urlsafe(32)
        session["_csrf_token"] = token
    return token


app.jinja_env.globals["csrf_token"] = csrf_token


@app.after_request
def prevent_admin_cache(response):
    if request.path.startswith("/admin"):
        response.headers["Cache-Control"] = "no-store"
        response.headers["Pragma"] = "no-cache"
    return response


def admin_required(view_function):
    @wraps(view_function)
    def wrapped(*args, **kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin_login"))
        return view_function(*args, **kwargs)

    return wrapped


def valid_csrf_token():
    expected = session.get("_csrf_token", "")
    supplied = request.form.get("csrf_token", "")
    return bool(expected) and secrets.compare_digest(supplied, expected)


@app.route("/", methods=["GET", "POST"])
def register():
    values = {"surname": "", "first_name": "", "phone": "", "email": ""}

    if request.method == "POST":
        values = {key: request.form.get(key, "").strip() for key in values}
        if not all(values.values()):
            flash("Бүх талбарыг бөглөнө үү.", "error")
        elif not valid_email(values["email"]):
            flash("Имэйл хаягийн формат буруу байна. Жишээ: bat@example.com", "error")
        elif not re.fullmatch(r"[+\d\s()-]+", values["phone"]) or not 8 <= sum(
            character.isdigit() for character in values["phone"]
        ) <= 15:
            flash("Утасны дугаараа зөв оруулна уу. 8–15 оронтой байж болно.", "error")
        else:
            try:
                save_registration(**values)
            except sqlite3.Error:
                app.logger.exception("Could not save registration to the SQLite database")
                flash("Бүртгэл хадгалагдсангүй. Дахин оролдоно уу.", "error")
            else:
                flash("Таны бүртгэл амжилттай хадгалагдлаа.", "success")
                return redirect(url_for("register"))

    return render_template("index.html", values=values)


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    password_configured = bool(app.config["ADMIN_PASSWORD"])
    if request.method == "POST":
        if not valid_csrf_token():
            abort(400)
        if not password_configured:
            abort(503, description="Set the ADMIN_PASSWORD environment variable.")
        password = request.form.get("password", "")
        if secrets.compare_digest(password, app.config["ADMIN_PASSWORD"]):
            session.clear()
            session["admin_authenticated"] = True
            return redirect(url_for("admin_dashboard"))
        flash("Нууц үг буруу байна.", "error")

    return render_template(
        "admin_login.html",
        password_configured=password_configured,
    )


@app.route("/admin")
@admin_required
def admin_dashboard():
    registrations = get_registrations()
    return render_template(
        "admin.html",
        registrations=registrations,
    )


@app.route("/admin/export.xlsx")
@admin_required
def export_registrations():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Бүртгэл"
    worksheet.append(EXPORT_HEADERS)

    for registration in get_registrations():
        row = worksheet.max_row + 1
        values = (
            registration["id"],
            registration["surname"],
            registration["first_name"],
            registration["phone"],
            registration["email"],
            registration["created_at"],
        )
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row, column=column, value=str(value))
            cell.data_type = "s"

    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="24465B")
    for column, width in enumerate((8, 18, 18, 22, 32, 22), start=1):
        worksheet.column_dimensions[worksheet.cell(row=1, column=column).column_letter].width = width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Register.xlsx",
    )


@app.route("/admin/logout", methods=["POST"])
@admin_required
def admin_logout():
    if not valid_csrf_token():
        abort(400)
    session.clear()
    flash("Админаас амжилттай гарлаа.", "success")
    return redirect(url_for("admin_login"))


if __name__ == "__main__":
    initialize_database()
    app.run()
