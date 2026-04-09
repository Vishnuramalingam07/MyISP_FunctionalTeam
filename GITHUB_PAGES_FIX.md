# Quick Fix: Deploy Backend for GitHub Pages

## Problem
Your GitHub Pages site (https://vishnuramalingam07.github.io/MyISP_FunctionalTeam/) cannot run Python/Flask backend code. The error "Network error: Unexpected token '<'" occurs because GitHub Pages returns HTML 404 pages instead of JSON when JavaScript tries to call `/api/...` endpoints.

## Solution Overview
1. ✅ **Created `config.js`** - Central configuration for API URLs
2. ✅ **Updated `daily-report.html`** - Now uses `config.js` 
3. ⏳ **Deploy Flask backend** - Deploy to Render.com (free tier available)
4. ⏳ **Update remaining HTML files** - Point to deployed backend
5. ⏳ **Configure `config.js`** - Set your deployed backend URL

---

## Step 1: Deploy Your Flask Backend to Render.com (FREE)

### Option A: Automatic Deployment (Recommended)

1. **Go to Render.com:**
   - Visit https://render.com
   - Click **"Get Started for Free"**
   - Sign up with GitHub

2. **Deploy from GitHub:**
   - Click **"New +"** → **"Web Service"**
   - Connect your GitHub account
   - Select your repository: `MyISP_FunctionalTeam`
   - Render will **automatically detect** your `render.yaml` configuration!

3. **Configure Environment Variables** (if needed):
   - In Render dashboard, go to **Environment**
   - Add these if using Supabase/PostgreSQL:
     - `SUPABASE_URL` = your_supabase_url
     - `SUPABASE_ANON_KEY` = your_supabase_key
   - Otherwise, leave blank (app will use CSV fallback)

4. **Deploy:**
   - Click **"Create Web Service"**
   - Wait 3-5 minutes for deployment
   - Copy your deployed URL: **`https://myisp-tools.onrender.com`**

5. **IMPORTANT - Free Tier Note:**
   - Free tier "spins down" after 15 minutes of inactivity
   - First request after sleep takes 30-60 seconds to wake up
   - Consider upgrading to $7/month plan for always-on service

---

## Step 2: Update config.js with Your Deployed URL

Open `config.js` and update line 25:

```javascript
// Before (local development):
const API_BASE_URL = '';

// After (production with Render):
const API_BASE_URL = 'https://myisp-tools.onrender.com';  // ← Your Render URL
```

---

## Step 3: Update Remaining HTML Files

I've updated `daily-report.html` for you. Run this PowerShell script to update ALL HTML files:

```powershell
# Update all HTML files to include config.js and use getApiUrl()
# Run this from your MyISP_FunctionalTeam folder

$htmlFiles = @(
    'index.html',
    'hotfix-daily-report.html', 
    'regression-report.html',
    'tc-compare-report.html',
    'm-poc-extref.html',
    'ado-testcase-upload.html',
    'ai-data-transfer.html',
    'missing-data-scope-report.html',
    'missing-filed-report.html',
    'attendance-report.html',
    'defect-status-report.html'
)

foreach ($file in $htmlFiles) {
    $path = Join-Path $PSScriptRoot $file
    if (Test-Path $path) {
        $content = Get-Content $path -Raw
        
        # Check if config.js is already included
        if ($content -notmatch '<script src="config.js"></script>') {
            Write-Host "Updating $file..." -ForegroundColor Yellow
            
            # Add config.js before first <script> tag
            $content = $content -replace '(<script>)', '<script src="config.js"></script>`n$1'
            
            # Update all /api/ fetch calls
            $content = $content -replace "fetch\('(/api/[^']+)'", "fetch(getApiUrl('`$1')"
            $content = $content -replace 'fetch\("(/api/[^"]+)"', 'fetch(getApiUrl("$1")'
            
            Set-Content $path -Value $content -NoNewline
            Write-Host "✓ Updated $file" -ForegroundColor Green
        } else {
            Write-Host "○ $file already updated" -ForegroundColor Gray
        }
    }
}

Write-Host "`n✅ All HTML files updated!" -ForegroundColor Green
```

Or **manually update each HTML file** by:
1. Adding `<script src="config.js"></script>` before the first `<script>` tag
2. Replacing `fetch('/api/xxx')` with `fetch(getApiUrl('/api/xxx'))`

---

## Step 4: Test Locally First

Before deploying to GitHub Pages, test locally:

```powershell
# Start Flask backend
python app.py

# Open http://localhost:5000 in browser
# config.js should have API_BASE_URL = '' for local testing
```

---

## Step 5: Deploy to GitHub Pages

```bash
# Commit changes
git add config.js daily-report.html *.html
git commit -m "Fix: Add backend URL configuration for GitHub Pages deployment"

# Push to GitHub
git push origin main
```

GitHub Pages will auto-update in 1-2 minutes.

---

## Verification Checklist

- [ ] Flask backend deployed to Render.com
- [ ] Received Render URL (e.g., https://myisp-tools.onrender.com)
- [ ] Updated `config.js` with Render URL
- [ ] Updated all HTML files to use `getApiUrl()`
- [ ] Pushed changes to GitHub
- [ ] GitHub Pages updated
- [ ] Tested report generation from GitHub Pages URL

---

## Alternative: Local Network Deployment

If you prefer to keep everything local (no cloud):

1. **Keep Flask running on your computer:**
   ```powershell
   python run_production.py  # Or use Start-All.ps1
   ```

2. **Update config.js with your local IP:**
   ```javascript
   const API_BASE_URL = 'http://192.168.1.2:5000';  // Your computer's IP
   ```

3. **Ensure firewall allows connections:**
   ```powershell
   .\Enable-Firewall.bat
   ```

**Limitation:** Backend must always be running, and only accessible on your network.

---

## Troubleshooting

### Issue: "Network error: Unexpected token '<'"
**Cause:** Frontend is calling `/api/...` but getting HTML instead of JSON
**Fix:** Ensure `API_BASE_URL` in `config.js` points to your deployed backend

### Issue: "CORS error"
**Cause:** Browser blocking cross-origin requests
**Fix:** Already handled in `app.py` - ensure backend has CORS enabled:
```python
from flask_cors import CORS
CORS(app, resources={r"/api/*": {"origins": "*"}})
```

### Issue: "500 Internal Server Error" on Render
**Cause:** Missing environment variables or database connection failure
**Fix:** Check Render logs, ensure environment variables are set

### Issue: Slow first request on Render
**Cause:** Free tier spins down after 15 minutes
**Fix:** Upgrade to paid plan or accept 30-60s startup delay

---

## Need Help?

1. Check Render deployment logs
2. Check browser console (F12) for errors
3. Verify `config.js` has correct URL
4. Ensure all HTML files include `<script src="config.js"></script>`

---

**TL;DR:** 
1. Deploy Flask to Render.com (free)
2. Copy Render URL to `config.js`
3. Update all HTML files to use `getApiUrl()`
4. Push to GitHub
5. Done! 🎉
