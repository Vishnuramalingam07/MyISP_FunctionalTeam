# Architecture Overview - What Goes Where

## CURRENTLY (BROKEN) ❌

```
User's Browser
    |
    | Opens: https://vishnuramalingam07.github.io/MyISP_FunctionalTeam/
    v
GitHub Pages (Static Files Only)
    |
    | HTML tries to call: /api/generate-daily-status-report
    v
GitHub Pages returns: 404 HTML Error Page ❌
    |
    v
JavaScript expects: JSON
JavaScript gets: HTML ("<DOCTYPE html>...")
    |
    v
ERROR: "Unexpected token '<'" ❌
```

---

## AFTER FIX (WORKING) ✅

```
User's Browser
    |
    +--------------------+--------------------+
    |                                         |
    v                                         v
GitHub Pages                         Render.com
(Frontend - HTML/CSS/JS)            (Backend - Flask/Python)
    |                                         |
    | Serves:                                 | Runs:
    | - index.html                            | - app.py
    | - daily-report.html                     | - PostgreSQL queries
    | - config.js                             | - SharePoint downloads
    | - All static files                      | - Report generation
    |                                         |
    | config.js has:                          |
    | API_BASE_URL =                          |
    | 'https://myisp-tools.onrender.com' -----+
    |                                         |
    | User clicks "Generate Report"           |
    |                                         |
    | JavaScript calls:                       |
    | fetch(getApiUrl('/api/...'))            |
    |                                         |
    | Translates to:                          |
    | https://myisp-tools.onrender.com/api/...
    |                                         |
    +-----------------------------------------+
                        |
                        v
            Returns JSON ✅
                        |
                        v
            Report Generated! 🎉
```

---

## WHERE EACH COMPONENT LIVES

### GitHub Pages (Already Working)
- **URL:** `https://vishnuramalingam07.github.io/MyISP_FunctionalTeam/`
- **Hosts:** All `.html`, `.css`, `.js` files
- **Cost:** FREE forever
- **Limitation:** Cannot run Python/backend code

### Render.com (Need to Deploy)
- **URL:** `https://myisp-tools.onrender.com` (your custom URL)
- **Hosts:** Flask backend (`app.py`)
- **Runs:** Python scripts, database queries, report generation
- **Cost:** FREE (with sleep after 15 min inactivity)
- **Upgrade:** $7/month for always-on

---

## WHAT EACH FILE DOES

### Frontend (GitHub Pages)
```
index.html              → Main dashboard (already on GitHub Pages)
daily-report.html       → Report generation page (already on GitHub Pages)
config.js               → Points to Render.com backend ← YOU NEED TO UPDATE THIS
```

### Backend (Render.com)
```
app.py                  → Flask server (handles API requests)
render.yaml             → Tells Render how to deploy
requirements.txt        → Python dependencies
postgres_client.py      → Database connection
```

---

## SIMPLE ANALOGY

Think of it like a restaurant:

**GitHub Pages = The Dining Room (Frontend)**
- Beautiful tables and menus
- Customers can see and order
- Cannot cook food

**Render.com = The Kitchen (Backend)**
- Invisible to customers
- Does all the work
- Prepares and delivers food

**config.js = The Kitchen Phone Number**
- Tells the dining room how to reach the kitchen
- Without it, orders go nowhere!

---

## YOUR ACTION ITEMS

1. ✅ Frontend is already deployed (GitHub Pages)
2. ⏳ Deploy backend to Render.com ← **DO THIS NOW**
3. ⏳ Update config.js with Render URL
4. ⏳ Push to GitHub
5. ✅ Everything works!

**Next Step:** Open `DEPLOY_BACKEND_STEPS.md` and follow the instructions!
