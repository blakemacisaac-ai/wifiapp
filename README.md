# Hotel WiFi Request Portal

A standalone Flask app with two views:
- **`/`** — Clean management form for hotel staff to submit conference WiFi requests
- **`/admin`** — Tech panel for IT to review the queue and push SSIDs directly to Meraki

---

## Local Setup

```bash
pip install -r requirements.txt
python app.py
```

Visit `http://localhost:5000` for the management form.
Visit `http://localhost:5000/admin` for the tech panel.

---

## Deploy to Railway (Free)

### 1. Push to GitHub
```bash
git init
git add .
git commit -m "initial commit"
gh repo create wifi-portal --public --push
# or push to an existing repo
```

### 2. Create Railway project
1. Go to [railway.app](https://railway.app) and sign in with GitHub
2. Click **New Project → Deploy from GitHub repo**
3. Select your `wifi-portal` repo
4. Railway auto-detects Python and deploys

### 3. Add a Postgres database
1. In your Railway project, click **+ New** → **Database** → **Add PostgreSQL**
2. Railway provisions a Postgres instance and automatically injects `DATABASE_URL` into your app's environment — no copy/paste needed

### 4. Set environment variables
In Railway → your app service → **Variables**, add:

| Variable | Value |
|---|---|
| `SECRET_KEY` | any long random string |
| `ADMIN_PASSWORD` | your chosen admin password |
| `MERAKI_API_KEY` | your Meraki dashboard API key |
| `MERAKI_NET_ID` | your Meraki network ID (e.g. `N_xxxxxxxxxx`) |

`DATABASE_URL` is injected automatically by the Postgres plugin — do not set it manually.

### 5. Generate a public domain
Railway → your app service → **Settings → Networking → Generate Domain**

You'll get a URL like `wifi-portal-production.up.railway.app`

---

## Usage

| URL | Who uses it |
|---|---|
| `your-app.up.railway.app/` | Hotel management — submit requests |
| `your-app.up.railway.app/admin` | IT team — review queue, assign slots, push to Meraki |

---

## How the Meraki push works

The server calls the Meraki API directly:
```
PUT https://api.meraki.com/api/v1/networks/{NETWORK_ID}/wireless/ssids/{slot-1}
```
Sets the SSID name, enables it, and applies WPA2-PSK. No CORS issues since it's server-side.

---

## Notes

- SQLite DB (`requests.db`) is created automatically on first run
- The admin panel auto-refreshes every 30 seconds and notifies you of new submissions
- SSID slots 1–2 are flagged in the UI as hotel reserved — use slots 3+ for conferences
- Passwords are stored in the DB (consider encrypting at rest for production use)
