import os
import requests
import psycopg2
import psycopg2.extras
import threading
from datetime import datetime, timezone
from functools import wraps
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, jsonify, g
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-production")

# ── Config ───────────────────────────────────────────────────
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
MERAKI_API_KEY = os.environ.get("MERAKI_API_KEY", "")
MERAKI_NET_ID  = os.environ.get("MERAKI_NET_ID", "")
DATABASE_URL   = os.environ.get("DATABASE_URL", "")
MERAKI_BASE    = "https://api.meraki.com/api/v1"

# Tier → true kbps using 1024-based conversion so Meraki reports exact Mbps
# 25 Mbps = 25 * 1024 = 25600 kbps, etc.
TIERS = {
    "1": {"label": "Tier 1 — Standard",  "mbps": 25,  "kbps": 25600},
    "2": {"label": "Tier 2 — Enhanced",  "mbps": 50,  "kbps": 51200},
    "3": {"label": "Tier 3 — Premium",   "mbps": 100, "kbps": 102400},
}

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

def get_scheduler_db():
    """Separate DB connection for the scheduler thread (not Flask g)."""
    return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)

def init_db():
    with app.app_context():
        db = get_db()
        cur = db.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS wifi_requests (
                id               SERIAL PRIMARY KEY,
                conf_name        TEXT NOT NULL,
                ssid             TEXT NOT NULL,
                password         TEXT NOT NULL,
                start_date       TEXT NOT NULL,
                end_date         TEXT NOT NULL,
                notes            TEXT,
                tier             TEXT DEFAULT '1',
                status           TEXT DEFAULT 'pending',
                slot             INTEGER,
                error_msg        TEXT,
                submitted_at     TIMESTAMP DEFAULT NOW(),
                pushed_at        TIMESTAMP,
                enable_at        TIMESTAMP,
                disable_at       TIMESTAMP,
                schedule_status  TEXT DEFAULT 'unscheduled'
            )
        """)
        # Safe migrations for existing DBs
        for col, defn in [
            ("tier",            "TEXT DEFAULT '1'"),
            ("enable_at",       "TIMESTAMP"),
            ("disable_at",      "TIMESTAMP"),
            ("schedule_status", "TEXT DEFAULT 'unscheduled'"),
        ]:
            cur.execute(f"ALTER TABLE wifi_requests ADD COLUMN IF NOT EXISTS {col} {defn}")
        db.commit()
        cur.close()

# ── Auth ──────────────────────────────────────────────────────
def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("admin"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated

# ── Meraki helpers ────────────────────────────────────────────
def meraki_headers():
    return {
        "X-Cisco-Meraki-API-Key": MERAKI_API_KEY,
        "Content-Type": "application/json"
    }

def get_live_ssids():
    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        return None, "Meraki not configured"
    try:
        resp = requests.get(
            f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids",
            headers=meraki_headers(), timeout=10
        )
        if resp.status_code == 200:
            return resp.json(), None
        return None, f"HTTP {resp.status_code}"
    except Exception as e:
        return None, str(e)

def push_ssid_to_meraki(row, slot_index):
    """Core push logic — used by both manual push and scheduler."""
    tier_key  = str(row.get("tier") or "1")
    tier      = TIERS.get(tier_key, TIERS["1"])
    bw_kbps   = tier["kbps"]
    resp = requests.put(
        f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{slot_index}",
        headers=meraki_headers(),
        json={
            "name": row["ssid"],
            "enabled": True,
            "authMode": "psk",
            "encryptionMode": "wpa",
            "wpaEncryptionMode": "WPA2 only",
            "psk": row["password"],
            "perClientBandwidthLimitUp":   bw_kbps,
            "perClientBandwidthLimitDown": bw_kbps,
        },
        timeout=10
    )
    return resp

# ══════════════════════════════════════════════════════════════
# SCHEDULER  (runs in background thread, checks every 60s)
# ══════════════════════════════════════════════════════════════
def scheduler_tick():
    """Called every 60s. Enables/disables SSIDs based on schedule."""
    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        return
    try:
        db  = get_scheduler_db()
        cur = db.cursor()
        now = datetime.now(timezone.utc)

        # ── Auto-enable: pushed rows where enable_at has passed and not yet enabled
        cur.execute("""
            SELECT * FROM wifi_requests
            WHERE status = 'pushed'
              AND slot IS NOT NULL
              AND enable_at IS NOT NULL
              AND enable_at <= %s
              AND (schedule_status IS NULL OR schedule_status NOT IN ('enabled', 'disabled'))
        """, (now,))
        to_enable = cur.fetchall()
        for row in to_enable:
            try:
                slot_index = int(row["slot"]) - 1
                push_ssid_to_meraki(row, slot_index)
                cur.execute(
                    "UPDATE wifi_requests SET schedule_status='enabled' WHERE id=%s",
                    (row["id"],)
                )
                db.commit()
                print(f"[scheduler] Enabled SSID '{row['ssid']}' on slot {row['slot']}")
            except Exception as e:
                print(f"[scheduler] Enable error for id {row['id']}: {e}")

        # ── Auto-disable: pushed rows where disable_at has passed
        cur.execute("""
            SELECT * FROM wifi_requests
            WHERE status = 'pushed'
              AND slot IS NOT NULL
              AND disable_at IS NOT NULL
              AND disable_at <= %s
              AND (schedule_status IS NULL OR schedule_status != 'disabled')
        """, (now,))
        to_disable = cur.fetchall()
        for row in to_disable:
            try:
                slot_index = int(row["slot"]) - 1
                requests.put(
                    f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{slot_index}",
                    headers=meraki_headers(),
                    json={"enabled": False},
                    timeout=10
                )
                cur.execute(
                    "UPDATE wifi_requests SET schedule_status='disabled' WHERE id=%s",
                    (row["id"],)
                )
                db.commit()
                print(f"[scheduler] Disabled SSID '{row['ssid']}' on slot {row['slot']}")
            except Exception as e:
                print(f"[scheduler] Disable error for id {row['id']}: {e}")

        cur.close()
        db.close()
    except Exception as e:
        print(f"[scheduler] tick error: {e}")

def start_scheduler():
    def loop():
        import time
        while True:
            scheduler_tick()
            time.sleep(60)
    t = threading.Thread(target=loop, daemon=True)
    t.start()

# ══════════════════════════════════════════════════════════════
# MANAGEMENT ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def index():
    return render_template("request.html", tiers=TIERS)

@app.route("/submit", methods=["POST"])
def submit():
    conf_name  = request.form.get("conf_name", "").strip()
    ssid       = request.form.get("ssid", "").strip().replace(" ", "-")
    password   = request.form.get("password", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date   = request.form.get("end_date", "").strip()
    notes      = request.form.get("notes", "").strip()
    tier       = request.form.get("tier", "1").strip()

    errors = []
    if not conf_name:      errors.append("Event name is required.")
    if not ssid:           errors.append("Network name is required.")
    if len(ssid) > 32:     errors.append("Network name must be 32 characters or less.")
    if len(password) < 8:  errors.append("Password must be at least 8 characters.")
    if not start_date or not end_date: errors.append("Start and end dates are required.")
    if start_date and end_date and end_date < start_date:
        errors.append("End date must be after start date.")
    if tier not in TIERS:  errors.append("Invalid tier selected.")

    if errors:
        return render_template("request.html", errors=errors, form=request.form, tiers=TIERS)

    db = get_db()
    cur = db.cursor()
    cur.execute(
        "INSERT INTO wifi_requests (conf_name, ssid, password, start_date, end_date, notes, tier) VALUES (%s,%s,%s,%s,%s,%s,%s)",
        (conf_name, ssid, password, start_date, end_date, notes, tier)
    )
    db.commit()
    cur.close()
    return render_template("request.html", success=True, tiers=TIERS)

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
    live_ssids, ssid_error = get_live_ssids()
    return render_template(
        "admin.html",
        requests=requests_list,
        live_ssids=live_ssids,
        ssid_error=ssid_error,
        meraki_configured=bool(MERAKI_API_KEY and MERAKI_NET_ID),
        tiers=TIERS
    )

@app.route("/admin/ssids")
@admin_required
def api_live_ssids():
    ssids, err = get_live_ssids()
    if err:
        return jsonify({"ok": False, "error": err})
    return jsonify({"ok": True, "ssids": ssids})

@app.route("/admin/ssid/toggle", methods=["POST"])
@admin_required
def toggle_ssid():
    data    = request.json or {}
    slot    = data.get("slot")
    enabled = data.get("enabled")
    if slot is None or enabled is None:
        return jsonify({"ok": False, "error": "Missing slot or enabled"}), 400
    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        return jsonify({"ok": False, "error": "Meraki not configured"}), 500
    try:
        resp = requests.put(
            f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{slot}",
            headers=meraki_headers(),
            json={"enabled": bool(enabled)},
            timeout=10
        )
        if resp.status_code == 200:
            return jsonify({"ok": True, "slot": slot, "enabled": enabled})
        return jsonify({"ok": False, "error": f"HTTP {resp.status_code}"}), 500
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

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

    data = request.json or {}
    slot = data.get("slot")
    if not slot:
        cur.close()
        return jsonify({"ok": False, "error": "No slot provided"}), 400
    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        cur.close()
        return jsonify({"ok": False, "error": "Meraki not configured"}), 500

    tier_key = str(row.get("tier") or "1")
    tier     = TIERS.get(tier_key, TIERS["1"])
    slot_index = int(slot) - 1

    try:
        resp = push_ssid_to_meraki(row, slot_index)
        if resp.status_code == 200:
            cur.execute(
                "UPDATE wifi_requests SET status='pushed', slot=%s, pushed_at=NOW(), error_msg=NULL WHERE id=%s",
                (slot, req_id)
            )
            db.commit()
            cur.close()
            return jsonify({"ok": True, "ssid": row["ssid"], "slot": slot, "mbps": tier["mbps"]})
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

@app.route("/admin/schedule/<int:req_id>", methods=["POST"])
@admin_required
def set_schedule(req_id):
    """Set or clear enable_at / disable_at for a pushed request."""
    data       = request.json or {}
    enable_at  = data.get("enable_at")   # ISO string or None
    disable_at = data.get("disable_at")  # ISO string or None

    def parse_dt(val):
        if not val:
            return None
        # Accept "2025-08-01T08:00" from datetime-local input
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(val, fmt).replace(tzinfo=timezone.utc)
            except ValueError:
                continue
        return None

    ea = parse_dt(enable_at)
    da = parse_dt(disable_at)

    db = get_db()
    cur = db.cursor()
    cur.execute(
        "UPDATE wifi_requests SET enable_at=%s, disable_at=%s, schedule_status='scheduled' WHERE id=%s",
        (ea, da, req_id)
    )
    db.commit()
    cur.close()
    return jsonify({"ok": True, "enable_at": str(ea), "disable_at": str(da)})

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
    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT * FROM wifi_requests ORDER BY submitted_at DESC")
    rows = cur.fetchall()
    cur.close()
    return jsonify(rows)

# ── Boot ──────────────────────────────────────────────────────
init_db()
start_scheduler()

if __name__ == "__main__":
    app.run(debug=True)
