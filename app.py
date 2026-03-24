import os
import requests
import psycopg2
import psycopg2.extras
from functools import wraps
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, jsonify, g
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-production")

# ── Config from environment ──────────────────────────────────
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
MERAKI_API_KEY = os.environ.get("MERAKI_API_KEY", "")
MERAKI_NET_ID  = os.environ.get("MERAKI_NET_ID", "")
DATABASE_URL   = os.environ.get("DATABASE_URL", "")

# ── Database ─────────────────────────────────────────────────
def get_db():
    db = getattr(g, "_database", None)
    if db is None:
        db = g._database = psycopg2.connect(
            DATABASE_URL,
            cursor_factory=psycopg2.extras.RealDictCursor
        )
    return db

@app.teardown_appcontext
def close_db(exception):
    db = getattr(g, "_database", None)
    if db is not None:
        db.close()

def init_db():
    with app.app_context():
        db = get_db()
        cur = db.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS wifi_requests (
                id           SERIAL PRIMARY KEY,
                conf_name    TEXT NOT NULL,
                ssid         TEXT NOT NULL,
                password     TEXT NOT NULL,
                start_date   TEXT NOT NULL,
                end_date     TEXT NOT NULL,
                notes        TEXT,
                status       TEXT DEFAULT 'pending',
                slot         INTEGER,
                error_msg    TEXT,
                submitted_at TIMESTAMP DEFAULT NOW(),
                pushed_at    TIMESTAMP
            )
        """)
        db.commit()
        cur.close()

# ── Auth decorator ────────────────────────────────────────────
def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("admin"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated

# ══════════════════════════════════════════════════════════════
# MANAGEMENT ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def index():
    return render_template("request.html")

@app.route("/submit", methods=["POST"])
def submit():
    conf_name  = request.form.get("conf_name", "").strip()
    ssid       = request.form.get("ssid", "").strip().replace(" ", "-")
    password   = request.form.get("password", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date   = request.form.get("end_date", "").strip()
    notes      = request.form.get("notes", "").strip()

    errors = []
    if not conf_name:            errors.append("Event name is required.")
    if not ssid:                 errors.append("Network name is required.")
    if len(ssid) > 32:           errors.append("Network name must be 32 characters or less.")
    if len(password) < 8:        errors.append("Password must be at least 8 characters.")
    if not start_date or not end_date: errors.append("Start and end dates are required.")
    if start_date and end_date and end_date < start_date:
        errors.append("End date must be after start date.")

    if errors:
        return render_template("request.html", errors=errors, form=request.form)

    db = get_db()
    cur = db.cursor()
    cur.execute(
        "INSERT INTO wifi_requests (conf_name, ssid, password, start_date, end_date, notes) VALUES (%s,%s,%s,%s,%s,%s)",
        (conf_name, ssid, password, start_date, end_date, notes)
    )
    db.commit()
    cur.close()
    return render_template("request.html", success=True)

# ══════════════════════════════════════════════════════════════
# ADMIN ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect(url_for("admin_panel"))
        error = "Incorrect password."
    return render_template("login.html", error=error)

@app.route("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin_login"))

@app.route("/admin")
@admin_required
def admin_panel():
    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT * FROM wifi_requests ORDER BY submitted_at DESC")
    requests_list = cur.fetchall()
    cur.close()
    return render_template(
        "admin.html",
        requests=requests_list,
        meraki_configured=bool(MERAKI_API_KEY and MERAKI_NET_ID)
    )

@app.route("/admin/push/<int:req_id>", methods=["POST"])
@admin_required
def push_ssid(req_id):
    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT * FROM wifi_requests WHERE id=%s", (req_id,))
    row = cur.fetchone()
    if not row:
        cur.close()
        return jsonify({"ok": False, "error": "Request not found"}), 404

    slot = request.json.get("slot")
    if not slot:
        cur.close()
        return jsonify({"ok": False, "error": "No slot provided"}), 400

    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        cur.close()
        return jsonify({"ok": False, "error": "Meraki API key or Network ID not configured on server"}), 500

    slot_index = int(slot) - 1
    url = f"https://api.meraki.com/api/v1/networks/{MERAKI_NET_ID}/wireless/ssids/{slot_index}"

    try:
        resp = requests.put(
            url,
            headers={
                "X-Cisco-Meraki-API-Key": MERAKI_API_KEY,
                "Content-Type": "application/json"
            },
            json={
                "name": row["ssid"],
                "enabled": True,
                "authMode": "psk",
                "encryptionMode": "wpa",
                "wpaEncryptionMode": "WPA2 only",
                "psk": row["password"]
            },
            timeout=10
        )
        if resp.status_code == 200:
            cur.execute(
                "UPDATE wifi_requests SET status='pushed', slot=%s, pushed_at=NOW(), error_msg=NULL WHERE id=%s",
                (slot, req_id)
            )
            db.commit()
            cur.close()
            return jsonify({"ok": True, "ssid": row["ssid"], "slot": slot})
        else:
            err = resp.json().get("errors", [f"HTTP {resp.status_code}"])
            err_str = ", ".join(err) if isinstance(err, list) else str(err)
            cur.execute("UPDATE wifi_requests SET status='error', slot=%s, error_msg=%s WHERE id=%s", (slot, err_str, req_id))
            db.commit()
            cur.close()
            return jsonify({"ok": False, "error": err_str})
    except Exception as e:
        cur.execute("UPDATE wifi_requests SET status='error', error_msg=%s WHERE id=%s", (str(e), req_id))
        db.commit()
        cur.close()
        return jsonify({"ok": False, "error": str(e)})

@app.route("/admin/delete/<int:req_id>", methods=["POST"])
@admin_required
def delete_request(req_id):
    db = get_db()
    cur = db.cursor()
    cur.execute("DELETE FROM wifi_requests WHERE id=%s", (req_id,))
    db.commit()
    cur.close()
    return jsonify({"ok": True})

@app.route("/admin/status")
@admin_required
def queue_status():
    """Polling endpoint so admin panel auto-refreshes."""
    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT * FROM wifi_requests ORDER BY submitted_at DESC")
    rows = cur.fetchall()
    cur.close()
    return jsonify(rows)

if __name__ == "__main__":
    init_db()
    app.run(debug=True)

# Always init on import (for gunicorn)
init_db()
