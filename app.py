from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_from_directory, abort
)
import sqlite3, os, sys
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "change-this-to-a-strong-secret"

# Paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "users.db")
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLES_DIR = os.path.join(STATIC_DIR, "samples")
os.makedirs(SAMPLES_DIR, exist_ok=True)

def get_conn():
    return sqlite3.connect(DB_PATH)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL
                )''')
    conn.commit()
    conn.close()

init_db()

# --------- Try importing openpyxl once at startup for clear diagnostics ----------
OPENPYXL_IMPORT_ERROR = None
LOAD_WORKBOOK = None
try:
    from openpyxl import load_workbook  # type: ignore
    LOAD_WORKBOOK = load_workbook
except Exception as e:
    OPENPYXL_IMPORT_ERROR = repr(e)

# ------------------- Auth & Core -------------------
@app.route('/')
def home():
    if session.get("username"):
        return redirect(url_for('claims'))
    return render_template('login.html')

@app.route('/s3display')
def s3display():
    return render_template('s3display.html')

@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')

    if not username or not password:
        return render_template('login.html', error="Username and password are required"), 400

    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT password FROM users WHERE username = ?", (username,))
    row = c.fetchone()
    conn.close()

    if row is None:
        return render_template('login.html', error="Invalid username or password"), 400

    stored_hash = row[0]
    if check_password_hash(stored_hash, password):
        session['username'] = username
        return redirect(url_for('claims'))
    else:
        return render_template('login.html', error="Invalid username or password"), 400

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

# ------------------- UI pages -------------------
@app.route('/claims')
def claims():
    return render_template('claims.html')

# ✅ FIXED: use real Flask dynamic segment, not HTML-escaped
@app.route('/review/<claim_id>')
def review_claim(claim_id):
    return render_template('review.html', claim_id=claim_id)

# ------------------- Excel Preview (openpyxl-only) + Download -------------------
def ws_to_html(ws, max_rows=None, max_cols=None):
    """
    Convert an openpyxl worksheet to a simple HTML table string.
    - max_rows/max_cols: optional limits for very large sheets (None = no limit)
    """
    from openpyxl.utils import get_column_letter
    from datetime import datetime, date, time

    def fmt(val):
        if isinstance(val, (datetime, date, time)):
            try:
                # Safe iso-like formatting
                return val.isoformat(sep=' ')
            except TypeError:
                return str(val)
        return "" if val is None else str(val)

    # Determine bounds
    max_row = ws.max_row if not max_rows else min(ws.max_row, max_rows)
    max_col = ws.max_column if not max_cols else min(ws.max_column, max_cols)

    # Header row (A, B, C...)
    thead_cells = [f"<th>{get_column_letter(c)}</th>" for c in range(1, max_col + 1)]
    thead_html = "<thead><tr>" + "".join(thead_cells) + "</tr></thead>"

    # Body rows
    body_rows = []
    for r in range(1, max_row + 1):
        tds = []
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            tds.append(f"<td>{fmt(cell.value)}</td>")
        body_rows.append("<tr>" + "".join(tds) + "</tr>")
    tbody_html = "<tbody>" + "".join(body_rows) + "</tbody>"

    return f'<table class="excel-table">{thead_html}{tbody_html}</table>'

@app.route('/excel-preview/<claim_id>')
def excel_preview(claim_id):
    """
    Server-side preview of static/samples/dummy.xlsx using openpyxl only.
    Renders each sheet as an HTML table (no client-side libs).
    """
    # If import failed at startup, show friendly diagnostics
    if LOAD_WORKBOOK is None:
        details = (
            f"Import error: {OPENPYXL_IMPORT_ERROR}\n"
            f"Python: {sys.version}\n"
            f"Executable: {sys.executable}\n"
            f"Sys.path[0]: {sys.path[0]}"
        )
        return render_template(
            'excel_preview.html',
            claim_id=claim_id,
            sheets=[],
            error="openpyxl import failed. See details below.",
            diag=details
        )

    sample_filename = 'dummy.xlsx'
    sample_abs_path = os.path.join(SAMPLES_DIR, sample_filename)
    if not os.path.exists(sample_abs_path):
        return render_template(
            'excel_preview.html',
            claim_id=claim_id,
            sheets=[],
            error="Dummy Excel not found at static/samples/dummy.xlsx",
            diag=None
        )

    try:
        # read_only=True for performance and to avoid locking
        wb = LOAD_WORKBOOK(sample_abs_path, data_only=True, read_only=True)
        sheets = []
        # Optional: set max_rows/max_cols if needed for huge workbooks
        for ws in wb.worksheets:
            html_table = ws_to_html(ws)  # ws_to_html(ws, max_rows=1000, max_cols=50)
            sheets.append({"name": ws.title, "table_html": html_table})
    except Exception as e:
        # Show the exact failure cause in the UI
        return render_template(
            'excel_preview.html',
            claim_id=claim_id,
            sheets=[],
            error=f"Failed to read dummy.xlsx: {e}",
            diag=f"Python: {sys.version}\nExecutable: {sys.executable}"
        )

    return render_template('excel_preview.html', claim_id=claim_id, sheets=sheets, error=None, diag=None)

@app.route('/excel/download')
def excel_download():
    """
    Downloads the dummy.xlsx file from static/samples/.
    """
    filename = 'dummy.xlsx'
    full_path = os.path.join(SAMPLES_DIR, filename)
    if not os.path.exists(full_path):
        abort(404, description="Dummy Excel not found at static/samples/dummy.xlsx")
    return send_from_directory(SAMPLES_DIR, filename, as_attachment=True)

# -------- Optional: quick diagnostics route --------
@app.route('/diag')
def diag():
    return {
        "python_version": sys.version,
        "executable": sys.executable,
        "openpyxl_import_error": OPENPYXL_IMPORT_ERROR,
        "cwd": os.getcwd(),
        "base_dir": BASE_DIR,
        "static_dir": STATIC_DIR,
        "samples_dir_exists": os.path.isdir(SAMPLES_DIR),
        "dummy_exists": os.path.exists(os.path.join(SAMPLES_DIR, "dummy.xlsx")),
        "sys_path_0": sys.path[0],
    }

if __name__ == '__main__':
    app.run(debug=True)