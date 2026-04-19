import os
import requests
import psycopg2
import psycopg2.extras
import threading
from datetime import datetime, timezone
from functools import wraps
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, jsonify, g, send_file
)
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

app = Flask(__name__)

# ── Crash loudly if critical secrets are missing ──────────────
_secret_key    = os.environ.get("SECRET_KEY", "")
_admin_pw      = os.environ.get("ADMIN_PASSWORD", "")
if not _secret_key:
    raise RuntimeError("SECRET_KEY environment variable is not set")
if not _admin_pw:
    raise RuntimeError("ADMIN_PASSWORD environment variable is not set")
app.secret_key = _secret_key

# Session expires after 8 hours of inactivity
from datetime import timedelta
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=8)

# Rate limiter — uses in-memory storage (no Redis needed)
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri="memory://"
)

# Security headers on every response
@app.after_request
def set_security_headers(resp):
    resp.headers["X-Frame-Options"]           = "DENY"
    resp.headers["X-Content-Type-Options"]    = "nosniff"
    resp.headers["Referrer-Policy"]           = "strict-origin-when-cross-origin"
    resp.headers["Content-Security-Policy"]   = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com; "
        "img-src 'self' data:; "
        "connect-src 'self';"
    )
    return resp

# ── Config ───────────────────────────────────────────────────
ADMIN_PASSWORD = _admin_pw
MERAKI_API_KEY = os.environ.get("MERAKI_API_KEY", "")
MERAKI_NET_ID  = os.environ.get("MERAKI_NET_ID", "")
DATABASE_URL   = os.environ.get("DATABASE_URL", "")
MERAKI_BASE      = "https://api.meraki.com/api/v1"
SLACK_WEBHOOK    = os.environ.get("SLACK_WEBHOOK", "")
PORTAL_PASSWORD  = os.environ.get("PORTAL_PASSWORD", "")  # shared passphrase for request form

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
            ("archived_at",     "TIMESTAMP"),
        ]:
            cur.execute(f"ALTER TABLE wifi_requests ADD COLUMN IF NOT EXISTS {col} {defn}")
        # Auto-provisioning tables
        cur.execute("""
            CREATE TABLE IF NOT EXISTS auto_slots (
                slot INTEGER PRIMARY KEY
            )
        """)
        for col, defn in [("auto_mode", "BOOLEAN DEFAULT FALSE")]:
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

def portal_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if PORTAL_PASSWORD and not session.get("portal"):
            return redirect(url_for("portal_login"))
        return f(*args, **kwargs)
    return decorated

# ── Slack notifications ───────────────────────────────────
def send_slack(msg):
    if not SLACK_WEBHOOK:
        return
    try:
        requests.post(SLACK_WEBHOOK, json={"text": msg}, timeout=5)
    except Exception as e:
        print(f"[slack] notify failed: {e}")

# ── Auto-slot helpers ────────────────────────────────────────
def get_auto_slots_db():
    try:
        db = get_scheduler_db()
        cur = db.cursor()
        cur.execute("SELECT slot FROM auto_slots ORDER BY slot DESC")
        rows = cur.fetchall()
        cur.close(); db.close()
        return [r["slot"] for r in rows]
    except Exception:
        return []

def find_free_auto_slot(auto_slots):
    if not auto_slots or not MERAKI_API_KEY or not MERAKI_NET_ID:
        return None
    try:
        resp = requests.get(
            f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids",
            headers=meraki_headers(), timeout=10
        )
        if resp.status_code != 200:
            return None
        live = {s["number"] + 1: s for s in resp.json()}
        for slot in auto_slots:
            s = live.get(slot)
            if s and (s["name"].startswith("Unconfigured") or not s.get("enabled", True)):
                return slot
        return None
    except Exception:
        return None

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
    url = f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{slot_index}"

    # Step 1: Configure the SSID fully
    resp = requests.put(
        url,
        headers=meraki_headers(),
        json={
            "name":                      row["ssid"],
            "enabled":                   True,
            "authMode":                  "psk",
            "encryptionMode":            "wpa",
            "wpaEncryptionMode":         "WPA2 only",
            "psk":                       row["password"],
            "perClientBandwidthLimitUp":   bw_kbps,
            "perClientBandwidthLimitDown": bw_kbps,
        },
        timeout=10
    )

    # Step 2: Explicit enable call — Meraki sometimes ignores enabled:True
    # on previously-unconfigured slots until a second call is made
    if resp.status_code == 200:
        import time; time.sleep(0.5)
        requests.put(
            url,
            headers=meraki_headers(),
            json={"enabled": True},
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
                send_slack(
                    f":large_green_circle: *Wi-Fi Network is Now Live!*\n"
                    f">*Event:* {row['conf_name']}\n"
                    f">*SSID:* `{row['ssid']}`\n"
                    f">*Password:* `{row['password']}`\n"
                    f">*Dates:* {row['start_date']} → {row['end_date']}\n"
                    f">_Your scheduled Wi-Fi network has been automatically enabled._"
                )
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

        # ── Auto-disable + archive: disable at 11:59 PM EST (UTC-5) on end_date
        # 23:59 EST = 04:59 UTC next day, so end_date + 28hrs 59min
        # Only acts on records pushed through this portal — never touches pre-existing SSIDs
        cur.execute("""
            SELECT * FROM wifi_requests
            WHERE status = 'pushed'
              AND end_date IS NOT NULL
              AND (end_date::date + interval '28 hours 59 minutes') <= NOW()
              AND (schedule_status IS NULL OR schedule_status != 'disabled')
        """)
        to_archive = cur.fetchall()
        for row in to_archive:
            try:
                # Disable on Meraki first
                if row["slot"] and MERAKI_API_KEY and MERAKI_NET_ID:
                    try:
                        requests.put(
                            f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{int(row['slot']) - 1}",
                            headers=meraki_headers(),
                            json={"enabled": False},
                            timeout=10
                        )
                        print(f"[scheduler] Disabled SSID '{row['ssid']}' on slot {row['slot']} (event ended)")
                        send_slack(
                            f":no_entry: *Wi-Fi Network Automatically Disabled*\n"
                            f">*Event:* {row['conf_name']}\n"
                            f">*SSID:* `{row['ssid']}`\n"
                            f">*End Date:* {row['end_date']}\n"
                            f">_This network has been disabled as the event has concluded._"
                        )
                    except Exception as me:
                        print(f"[scheduler] Meraki disable error for id {row['id']}: {me}")

                # Archive the record
                cur.execute(
                    "UPDATE wifi_requests SET status='archived', archived_at=NOW(), schedule_status='disabled' WHERE id=%s",
                    (row["id"],)
                )
                db.commit()
                print(f"[scheduler] Auto-archived '{row['ssid']}' (end_date passed)")
            except Exception as e:
                print(f"[scheduler] Archive error for id {row['id']}: {e}")

        # ── Auto-provisioning: process 'auto' queue ──────────────
        import datetime as _dt
        today_est = (now + _dt.timedelta(hours=-5)).strftime('%Y-%m-%d')
        auto_slots = get_auto_slots_db()

        cur.execute("""
            SELECT * FROM wifi_requests
            WHERE status = 'auto'
            ORDER BY start_date ASC, submitted_at ASC
            FOR UPDATE SKIP LOCKED
        """)
        auto_queue = cur.fetchall()

        for row in auto_queue:
            try:
                start       = row["start_date"]
                is_same_day = (start <= today_est)
                tier_key    = str(row.get("tier") or "1")
                tier_info   = TIERS.get(tier_key, TIERS["1"])

                slot = find_free_auto_slot(auto_slots)
                if not slot:
                    cur.execute(
                        "UPDATE wifi_requests SET status='needs_slot', error_msg=%s WHERE id=%s",
                        ("No automation slots available — manual assignment required", row["id"])
                    )
                    db.commit()
                    send_slack(
                        f":warning: *Automation Conflict — Manual Action Required*\n"
                        f">*Event:* {row['conf_name']}\n"
                        f">*SSID:* `{row['ssid']}`\n"
                        f">*Start Date:* {row['start_date']}\n"
                        f">_No automation slots available. Assign manually in the <https://aqsarniithotel.up.railway.app/admin|IT Admin Panel>._"
                    )
                    print(f"[auto] No slot for '{row['ssid']}'")
                    continue

                slot_index = int(slot) - 1

                if is_same_day:
                    resp = push_ssid_to_meraki(row, slot_index)
                    if resp.status_code == 200:
                        cur.execute("""
                            UPDATE wifi_requests
                            SET status='pushed', slot=%s, pushed_at=NOW(),
                                error_msg=NULL, schedule_status='enabled', auto_mode=TRUE
                            WHERE id=%s
                        """, (slot, row["id"]))
                        db.commit()
                        send_slack(
                            f":large_green_circle: *Wi-Fi Network is Now Live! (Auto-Provisioned)*\n"
                            f">*Event:* {row['conf_name']}\n"
                            f">*SSID:* `{row['ssid']}`\n"
                            f">*Password:* `{row['password']}`\n"
                            f">*Dates:* {row['start_date']} \u2192 {row['end_date']}\n"
                            f">*Speed:* {tier_info['mbps']} Mbps (Tier {tier_key})\n"
                            f">*Slot:* {slot} (auto-assigned)\n"
                            f">_Network is live and ready for guests._"
                        )
                        print(f"[auto] Provisioned '{row['ssid']}' on slot {slot}")
                    else:
                        cur.execute(
                            "UPDATE wifi_requests SET status='needs_slot', error_msg=%s WHERE id=%s",
                            (f"Meraki push failed: HTTP {resp.status_code}", row["id"])
                        )
                        db.commit()
                else:
                    enable_utc = _dt.datetime.strptime(start, "%Y-%m-%d").replace(
                        hour=5, minute=0, second=0, tzinfo=timezone.utc
                    )
                    requests.put(
                        f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{slot_index}",
                        headers=meraki_headers(),
                        json={
                            "name": row["ssid"], "enabled": False,
                            "authMode": "psk", "encryptionMode": "wpa",
                            "wpaEncryptionMode": "WPA2 only", "psk": row["password"],
                            "perClientBandwidthLimitUp":   tier_info["kbps"],
                            "perClientBandwidthLimitDown": tier_info["kbps"],
                        }, timeout=10
                    )
                    cur.execute("""
                        UPDATE wifi_requests
                        SET status='pushed', slot=%s, pushed_at=NOW(),
                            enable_at=%s, schedule_status='scheduled', auto_mode=TRUE
                        WHERE id=%s
                    """, (slot, enable_utc, row["id"]))
                    db.commit()
                    send_slack(
                        f":calendar: *Wi-Fi Network Scheduled (Auto-Provisioned)*\n"
                        f">*Event:* {row['conf_name']}\n"
                        f">*SSID:* `{row['ssid']}`\n"
                        f">*Password:* `{row['password']}`\n"
                        f">*Starts:* {row['start_date']} at 12:00 AM EST\n"
                        f">*Ends:* {row['end_date']} at 11:59 PM EST\n"
                        f">*Speed:* {tier_info['mbps']} Mbps (Tier {tier_key})\n"
                        f">*Slot:* {slot} (auto-assigned)\n"
                        f">_Pre-configured on Meraki. Will enable automatically._"
                    )
                    print(f"[auto] Scheduled '{row['ssid']}' on slot {slot} for {start}")
            except Exception as e:
                print(f"[auto] Error id {row['id']}: {e}")

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
# PORTAL LOGIN ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/portal/login", methods=["GET", "POST"])
@limiter.limit("10 per hour")
def portal_login():
    if not PORTAL_PASSWORD:
        return redirect(url_for("index"))
    if session.get("portal"):
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        if request.form.get("portal_password", "").strip() == PORTAL_PASSWORD:
            session.permanent = True
            session["portal"] = True
            return redirect(url_for("index"))
        error = "Incorrect access code. Please contact IT or the front desk."
    return render_template("portal_login.html", error=error)

@app.route("/portal/logout")
def portal_logout():
    session.pop("portal", None)
    return redirect(url_for("portal_login"))

# ══════════════════════════════════════════════════════════════
# MANAGEMENT ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/", methods=["GET"])
@portal_required
def index():
    return render_template("request.html", tiers=TIERS)

@app.route("/submit", methods=["POST"])
@portal_required
@limiter.limit("10 per hour")
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
    if not password:       errors.append("Password is required.")
    if not start_date or not end_date: errors.append("Start and end dates are required.")
    if start_date and end_date and end_date < start_date:
        errors.append("End date must be after start date.")
    if tier not in TIERS:  errors.append("Invalid tier selected.")

    if errors:
        return render_template("request.html", errors=errors, form=request.form, tiers=TIERS)

    db = get_db()
    cur = db.cursor()
    cur.execute(
        """INSERT INTO wifi_requests
           (conf_name, ssid, password, start_date, end_date, notes, tier, status, auto_mode)
           VALUES (%s,%s,%s,%s,%s,%s,%s,'auto',TRUE)""",
        (conf_name, ssid, password, start_date, end_date, notes, tier)
    )
    db.commit()
    cur.close()

    tier_info = TIERS.get(tier, TIERS["1"])
    send_slack(
        f":wifi: *New Conference WiFi Request*\n"
        f">*Event:* {conf_name}\n"
        f">*SSID:* `{ssid}`\n"
        f">*Password:* `{password}`\n"
        f">*Dates:* {start_date} → {end_date}\n"
        f">*Tier:* {tier_info['label']} ({tier_info['mbps']} Mbps)\n"
        f">*Notes:* {notes or '—'}\n"
        f">_Review it in the <https://aqsarniithotel.up.railway.app/admin|IT Admin Panel>_"
    )

    return render_template("request.html", success=True, tiers=TIERS)

# ══════════════════════════════════════════════════════════════
# ADMIN ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/admin/login", methods=["GET", "POST"])
@limiter.limit("10 per hour")
def admin_login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session.permanent = True  # apply PERMANENT_SESSION_LIFETIME
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
    cur.execute("""
        SELECT * FROM wifi_requests
        WHERE status IN ('pending','error','auto','needs_slot')
        ORDER BY submitted_at DESC
    """)
    pending_list = cur.fetchall()
    cur.execute("SELECT * FROM wifi_requests WHERE status='pushed' ORDER BY pushed_at DESC")
    active_list = cur.fetchall()
    cur.execute("SELECT slot FROM auto_slots ORDER BY slot")
    auto_slot_list = [r["slot"] for r in cur.fetchall()]
    cur.close()
    live_ssids, ssid_error = get_live_ssids()
    return render_template(
        "admin.html",
        pending_list=pending_list,
        active_list=active_list,
        live_ssids=live_ssids,
        ssid_error=ssid_error,
        meraki_configured=bool(MERAKI_API_KEY and MERAKI_NET_ID),
        tiers=TIERS,
        auto_slot_list=auto_slot_list
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
    if not isinstance(slot, int) or slot < 0 or slot > 14:
        return jsonify({"ok": False, "error": "Invalid slot number"}), 400
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
    slot       = data.get("slot")
    enable_at  = data.get("enable_at")
    disable_at = data.get("disable_at")
    if not slot:
        cur.close()
        return jsonify({"ok": False, "error": "No slot provided"}), 400
    if not MERAKI_API_KEY or not MERAKI_NET_ID:
        cur.close()
        return jsonify({"ok": False, "error": "Meraki not configured"}), 500

    def parse_dt(val):
        if not val: return None
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try: return datetime.strptime(val, fmt).replace(tzinfo=timezone.utc)
            except ValueError: continue
        return None

    ea = parse_dt(enable_at)
    da = parse_dt(disable_at)
    sched_status = "scheduled" if (ea or da) else "unscheduled"

    tier_key = str(row.get("tier") or "1")
    tier     = TIERS.get(tier_key, TIERS["1"])
    slot_index = int(slot) - 1

    try:
        resp = push_ssid_to_meraki(row, slot_index)
        if resp.status_code == 200:
            cur.execute(
                """UPDATE wifi_requests
                   SET status='pushed', slot=%s, pushed_at=NOW(), error_msg=NULL,
                       enable_at=%s, disable_at=%s, schedule_status=%s
                   WHERE id=%s""",
                (slot, ea, da, sched_status, req_id)
            )
            db.commit()
            cur.close()
            send_slack(
                f":large_green_circle: *Wi-Fi Network is Now Live!*\n"
                f">*Event:* {row['conf_name']}\n"
                f">*SSID:* `{row['ssid']}`\n"
                f">*Password:* `{row['password']}`\n"
                f">*Dates:* {row['start_date']} → {row['end_date']}\n"
                f">*Speed:* {tier['mbps']} Mbps (Tier {tier_key})\n"
                f">*Notes:* {row['notes'] or '—'}\n"
                f">_Your dedicated Wi-Fi network is configured and ready to use._"
            )
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

@app.route("/admin/update/<int:req_id>", methods=["POST"])
@admin_required
def update_request(req_id):
    """Update end_date, enable_at, disable_at on an active pushed request."""
    data       = request.json or {}
    end_date   = data.get("end_date", "").strip()
    enable_at  = data.get("enable_at")
    disable_at = data.get("disable_at")

    if not end_date:
        return jsonify({"ok": False, "error": "End date required"}), 400

    def parse_dt(val):
        if not val:
            return None
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try:
                return datetime.strptime(val, fmt).replace(tzinfo=timezone.utc)
            except ValueError:
                continue
        return None

    ea = parse_dt(enable_at)
    da = parse_dt(disable_at)
    sched_status = "scheduled" if (ea or da) else "unscheduled"

    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT start_date FROM wifi_requests WHERE id=%s", (req_id,))
    row = cur.fetchone()
    if not row:
        cur.close()
        return jsonify({"ok": False, "error": "Request not found"}), 404

    cur.execute("""
        UPDATE wifi_requests
        SET end_date=%s, enable_at=%s, disable_at=%s, schedule_status=%s
        WHERE id=%s
    """, (end_date, ea, da, sched_status, req_id))
    db.commit()
    cur.close()
    return jsonify({
        "ok":         True,
        "start_date": row["start_date"],
        "end_date":   end_date
    })

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
    # Exclude password from status polling endpoint — only used for change detection
    cur.execute("SELECT id, status, submitted_at FROM wifi_requests ORDER BY submitted_at DESC")
    rows = cur.fetchall()
    cur.close()
    return jsonify(rows)

@app.route("/admin/archive/<int:req_id>", methods=["POST"])
@admin_required
def archive_request(req_id):
    """Archive a pushed request. Optionally disable on Meraki first."""
    data           = request.json or {}
    disable_meraki = data.get("disable_meraki", False)

    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT * FROM wifi_requests WHERE id=%s", (req_id,))
    row = cur.fetchone()
    if not row:
        cur.close()
        return jsonify({"ok": False, "error": "Not found"}), 404

    meraki_result = None
    if disable_meraki and row["slot"] and MERAKI_API_KEY and MERAKI_NET_ID:
        try:
            resp = requests.put(
                f"{MERAKI_BASE}/networks/{MERAKI_NET_ID}/wireless/ssids/{int(row['slot']) - 1}",
                headers=meraki_headers(),
                json={"enabled": False},
                timeout=10
            )
            meraki_result = "disabled" if resp.status_code == 200 else f"error {resp.status_code}"
        except Exception as e:
            meraki_result = f"error: {str(e)}"

    cur.execute(
        "UPDATE wifi_requests SET status='archived', archived_at=NOW() WHERE id=%s",
        (req_id,)
    )
    db.commit()
    cur.close()
    return jsonify({"ok": True, "meraki": meraki_result})

@app.route("/admin/history")
@admin_required
def admin_history():
    db = get_db()
    cur = db.cursor()
    # Get distinct year-months that have pushed records
    cur.execute("""
        SELECT DISTINCT
            TO_CHAR(pushed_at, 'YYYY-MM') AS month_key,
            TO_CHAR(pushed_at, 'Month YYYY') AS month_label
        FROM wifi_requests
        WHERE status IN ('pushed','archived') AND pushed_at IS NOT NULL
        ORDER BY month_key DESC
    """)
    months = cur.fetchall()

    # Default to current month
    selected = request.args.get("month", "")
    if not selected and months:
        selected = months[0]["month_key"]

    records = []
    if selected:
        cur.execute("""
            SELECT
                id, conf_name, ssid, tier, start_date, end_date,
                notes, slot, pushed_at, disable_at, schedule_status
            FROM wifi_requests
            WHERE status IN ('pushed','archived')
              AND TO_CHAR(pushed_at, 'YYYY-MM') = %s
            ORDER BY pushed_at DESC
        """, (selected,))
        records = cur.fetchall()

    cur.close()
    return render_template(
        "history.html",
        months=months,
        selected=selected,
        records=records,
        tiers=TIERS
    )


def generate_pdf_buffer(month, records, month_label):
    """Shared PDF generator — uses fpdf2 (no system font dependencies)."""
    import io
    from fpdf import FPDF
    from datetime import datetime as dt

    DARK  = (30,  58,  42)
    MID   = (45,  90,  61)
    LIGHT = (232, 245, 236)
    WHITE = (255, 255, 255)
    GREY  = (245, 249, 246)
    RED   = (255, 240, 240)
    T1    = (219, 234, 254)
    T2    = (209, 250, 229)
    T3    = (254, 243, 199)
    DIM   = (118, 131, 144)

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_margins(12, 10, 12)

    # ── Header ──────────────────────────────────────────────
    pdf.set_fill_color(*DARK)
    pdf.rect(0, 0, 297, 18, "F")
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(*WHITE)
    pdf.set_xy(12, 3)
    pdf.cell(0, 8, "Aqsarniit Hotel & Conference Centre", ln=False)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(180, 200, 190)
    pdf.set_xy(12, 11)
    pdf.cell(0, 5, f"Wi-Fi Billing Report  |  {month_label}  |  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {len(records)} record(s)")

    pdf.set_y(22)

    # ── Column definitions ───────────────────────────────────
    cols = [
        ("Organization / Group", 64),
        ("SSID",                 44),
        ("Tier",                 18),
        ("Speed",                18),
        ("Start Date",           28),
        ("End Date",             28),
        ("Days",                 14),
        ("Status",               22),
    ]
    row_h = 8

    # ── Column header row ────────────────────────────────────
    pdf.set_fill_color(*DARK)
    pdf.set_text_color(*WHITE)
    pdf.set_font("Helvetica", "B", 8)
    for label, w in cols:
        pdf.cell(w, 7, label, border=0, align="C", fill=True)
    pdf.ln()

    # ── Green underline ──────────────────────────────────────
    pdf.set_fill_color(*MID)
    pdf.rect(pdf.get_x(), pdf.get_y(), sum(w for _, w in cols), 1, "F")
    pdf.ln(1)

    # ── Data rows ────────────────────────────────────────────
    tier_colors = {"1": T1, "2": T2, "3": T3}

    for i, r in enumerate(records):
        tk   = str(r["tier"] or "1")
        mbps = TIERS.get(tk, TIERS["1"])["mbps"]
        try:
            sd  = dt.strptime(str(r["start_date"]), "%Y-%m-%d")
            ed  = dt.strptime(str(r["end_date"]),   "%Y-%m-%d")
            dur = str((ed - sd).days + 1)
            sd_fmt = sd.strftime("%b %d, %Y")
            ed_fmt = ed.strftime("%b %d, %Y")
        except Exception:
            sd_fmt = str(r["start_date"])
            ed_fmt = str(r["end_date"])
            dur    = "-"

        status = str(r.get("status") or "pushed").upper()
        is_arch = status == "ARCHIVED"
        base_fill = RED if is_arch else (GREY if i % 2 == 0 else WHITE)

        row_data = [
            (r["conf_name"] or "")[:38],
            (r["ssid"] or "")[:24],
            f"Tier {tk}",
            f"{mbps} Mbps",
            sd_fmt, ed_fmt, dur, status,
        ]

        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(30, 40, 35)

        for ci, ((label, w), val) in enumerate(zip(cols, row_data)):
            is_tier = (ci == 2)
            fill_c  = tier_colors.get(tk, WHITE) if is_tier else base_fill
            pdf.set_fill_color(*fill_c)
            align = "C" if ci in (2, 3, 6, 7) else "L"
            pdf.cell(w, row_h, val, border="B", align=align, fill=True)
        pdf.ln()

    # ── Summary bar ──────────────────────────────────────────
    pdf.ln(3)
    pdf.set_fill_color(*LIGHT)
    pdf.set_text_color(*DARK)
    pdf.set_font("Helvetica", "B", 8)
    t1 = sum(1 for r in records if str(r["tier"]) == "1")
    t2 = sum(1 for r in records if str(r["tier"]) == "2")
    t3 = sum(1 for r in records if str(r["tier"]) == "3")
    summary = f"Total: {len(records)}   |   Tier 1 (25 Mbps): {t1}   |   Tier 2 (50 Mbps): {t2}   |   Tier 3 (100 Mbps): {t3}"
    pdf.cell(0, 7, summary, align="C", fill=True, border=0)

    buf = io.BytesIO(pdf.output())
    buf.seek(0)
    return buf

@app.route("/admin/history/export")
@admin_required
def export_history():
    month = request.args.get("month", "")
    if not month:
        return "Month required", 400
    try:
        db = get_db()
        cur = db.cursor()
        cur.execute("""
            SELECT conf_name, ssid, tier, start_date, end_date, notes, pushed_at, status
            FROM wifi_requests
            WHERE status IN ('pushed','archived')
              AND TO_CHAR(pushed_at, 'YYYY-MM') = %s
            ORDER BY start_date ASC
        """, (month,))
        records = cur.fetchall()
        cur.close()
    except Exception as e:
        return f"Database error: {str(e)}", 500
    try:
        from datetime import datetime as dt
        try:
            month_label = dt.strptime(month, "%Y-%m").strftime("%B %Y")
        except Exception:
            month_label = month
        buf = generate_pdf_buffer(month, records, month_label)
        return send_file(buf, mimetype="application/pdf",
                         as_attachment=True,
                         download_name=f"wifi_billing_{month}.pdf")
    except Exception as e:
        import traceback
        print(f"[export] {traceback.format_exc()}")
        return f"Export error: {str(e)}", 500

@app.route("/admin/history/slack", methods=["POST"])
@admin_required
def slack_billing():
    """Post billing report PDF link to Slack, or upload file if bot token available."""
    data  = request.json or {}
    month = data.get("month", "")
    if not month:
        return jsonify({"ok": False, "error": "Month required"}), 400
    if not SLACK_WEBHOOK:
        return jsonify({"ok": False, "error": "Slack not configured on server"}), 500
    try:
        db = get_db()
        cur = db.cursor()
        cur.execute("""
            SELECT conf_name, ssid, tier, start_date, end_date, notes, pushed_at, status
            FROM wifi_requests
            WHERE status IN ('pushed','archived')
              AND TO_CHAR(pushed_at, 'YYYY-MM') = %s
            ORDER BY start_date ASC
        """, (month,))
        records = cur.fetchall()
        cur.close()
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    try:
        from datetime import datetime as dt
        try:
            month_label = dt.strptime(month, "%Y-%m").strftime("%B %Y")
        except Exception:
            month_label = month

        slack_token = os.environ.get("SLACK_BOT_TOKEN", "")
        if slack_token:
            # Upload PDF directly to Slack
            buf = generate_pdf_buffer(month, records, month_label)
            import requests as req_lib
            resp = req_lib.post(
                "https://slack.com/api/files.upload",
                headers={"Authorization": f"Bearer {slack_token}"},
                data={
                    "channels":        os.environ.get("SLACK_CHANNEL", "#general"),
                    "filename":        f"wifi_billing_{month}.pdf",
                    "title":           f"Wi-Fi Billing — {month_label}",
                    "initial_comment": f":page_facing_up: *Aqsarniit Wi-Fi Billing Report — {month_label}*  |  {len(records)} record(s)",
                },
                files={"file": (f"wifi_billing_{month}.pdf", buf, "application/pdf")}
            )
            if resp.json().get("ok"):
                return jsonify({"ok": True, "method": "file"})
        # Fallback: send message with download link
        t1 = sum(1 for r in records if str(r["tier"])=="1")
        t2 = sum(1 for r in records if str(r["tier"])=="2")
        t3 = sum(1 for r in records if str(r["tier"])=="3")
        base = request.host_url.rstrip("/")
        send_slack(
            f":page_facing_up: *Wi-Fi Billing Report — {month_label}*\n"
            f">Total records: *{len(records)}*  |  Tier 1: {t1}  |  Tier 2: {t2}  |  Tier 3: {t3}\n"
            f">:arrow_down: <{base}/admin/history/export?month={month}|Download PDF>"
        )
        return jsonify({"ok": True, "method": "link"})
    except Exception as e:
        import traceback
        print(f"[slack_billing] {traceback.format_exc()}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/admin/history/delete/<int:req_id>", methods=["POST"])
@admin_required
def delete_history_record(req_id):
    """Hard delete a billing record — only use for mistakes/test data."""
    db = get_db()
    cur = db.cursor()
    cur.execute("DELETE FROM wifi_requests WHERE id=%s", (req_id,))
    db.commit()
    cur.close()
    return jsonify({"ok": True})

@app.route("/admin/manual/<int:req_id>", methods=["POST"])
@admin_required
def make_manual(req_id):
    db = get_db(); cur = db.cursor()
    cur.execute("UPDATE wifi_requests SET status='pending', auto_mode=FALSE WHERE id=%s", (req_id,))
    db.commit(); cur.close()
    return jsonify({"ok": True})

@app.route("/admin/settings", methods=["GET", "POST"])
@admin_required
def admin_settings():
    db = get_db(); cur = db.cursor()
    if request.method == "POST":
        data  = request.json or {}
        slots = [int(s) for s in data.get("slots", []) if 1 <= int(s) <= 15]
        cur.execute("DELETE FROM auto_slots")
        for s in slots:
            cur.execute("INSERT INTO auto_slots (slot) VALUES (%s) ON CONFLICT DO NOTHING", (s,))
        db.commit(); cur.close()
        return jsonify({"ok": True, "slots": slots})
    cur.execute("SELECT slot FROM auto_slots ORDER BY slot")
    slots = [r["slot"] for r in cur.fetchall()]
    cur.close()
    return jsonify({"slots": slots})

# ── Boot ──────────────────────────────────────────────────────
init_db()
start_scheduler()

if __name__ == "__main__":
    app.run(debug=False)
