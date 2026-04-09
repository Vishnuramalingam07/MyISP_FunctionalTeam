# STEP-BY-STEP DEPLOYMENT GUIDE
## Deploy MyISP Tools Backend to Render.com (FREE)

### THE PROBLEM
- GitHub Pages hosts your HTML files (frontend)
- GitHub Pages CANNOT run Python/Flask (backend)
- You need to deploy Flask separately

### THE SOLUTION
Deploy Flask backend to Render.com (free tier available)

---

## 📋 STEP 1: Create Render.com Account

1. Go to: **https://render.com**
2. Click **"Get Started for Free"**
3. Click **"Sign up with GitHub"**
4. Authorize Render to access your GitHub account
5. Done! You now have a Render account

---

## 📋 STEP 2: Deploy Your Flask App

### 2.1 Create New Web Service

1. In Render dashboard, click **"New +"** (top right)
2. Select **"Web Service"**

### 2.2 Connect Your Repository

1. If this is your first time:
   - Click **"Configure account"** 
   - Select which repositories Render can access
   - Choose **"Only select repositories"**
   - Select: `MyISP_FunctionalTeam`
   - Click **"Install"**

2. Back in Render:
   - You should see `MyISP_FunctionalTeam` in the list
   - Click **"Connect"** next to it

### 2.3 Configure Service (Render Auto-Detects Everything!)

Render will automatically read your `render.yaml` file and configure:
- ✅ Name: `myisp-tools`
- ✅ Environment: `Python`
- ✅ Region: `Singapore`
- ✅ Build Command: `pip install -r requirements.txt`
- ✅ Start Command: `waitress-serve --host=0.0.0.0 --port=$PORT app:app`

**You don't need to change anything!** Just click **"Create Web Service"**

### 2.4 Wait for Deployment

1. Render will start building your app
2. Watch the logs scroll (this takes 2-4 minutes)
3. Look for: **"Your service is live"** 🎉

### 2.5 Copy Your Deployed URL

Once deployed, you'll see your URL at the top:
```
https://myisp-tools.onrender.com
```

**COPY THIS URL** - you'll need it in the next step!

---

## 📋 STEP 3: Update config.js

1. Open `config.js` in your project
2. Find line 25 (the `API_BASE_URL` line)
3. Change it from:
   ```javascript
   const API_BASE_URL = '';
   ```
   
   To (use YOUR actual Render URL):
   ```javascript
   const API_BASE_URL = 'https://myisp-tools.onrender.com';
   ```

4. Save the file

---

## 📋 STEP 4: Push to GitHub

```powershell
# Navigate to your project folder
cd C:\Users\vishnu.ramalingam\Myisp_Tools_Live\MyISP_FunctionalTeam

# Add changes
git add config.js

# Commit
git commit -m "Configure backend URL for production deployment"

# Push to GitHub
git push origin main
```

GitHub Pages will auto-update in 1-2 minutes.

---

## 📋 STEP 5: Test Your Deployment

1. Wait 2 minutes for GitHub Pages to update
2. Go to: `https://vishnuramalingam07.github.io/MyISP_FunctionalTeam/daily-report.html`
3. Click **"Generate Daily Status Report"**
4. Should work! (No more "Network error: Unexpected token '<'" error)

**First request might be slow (30-60 seconds)** because Render free tier "wakes up" from sleep.

---

## ⚠️ IMPORTANT: Render Free Tier Limitations

- ✅ FREE forever
- ✅ 512 MB RAM
- ✅ Shared CPU
- ⚠️ **Spins down after 15 minutes of inactivity**
- ⚠️ **First request after sleep takes 30-60 seconds**

If you need always-on service: Upgrade to **Starter plan ($7/month)**

---

## 🔧 TROUBLESHOOTING

### Issue: "Repository not found in Render"
**Fix:** 
1. Go to Render dashboard
2. Click your profile → Account Settings
3. Click "GitHub" under Connected Accounts
4. Click "Configure" 
5. Add `MyISP_FunctionalTeam` repository

### Issue: Build fails on Render
**Fix:** Check Render logs for errors. Common issues:
- Missing dependencies in `requirements.txt`
- Python version mismatch
- Port configuration

### Issue: "CORS error" in browser console
**Fix:** Your `app.py` should already have CORS enabled. If not, add:
```python
from flask_cors import CORS
CORS(app)
```

### Issue: Still getting "Network error" on GitHub Pages
**Fix:** 
1. Check that `config.js` has the correct Render URL
2. Clear browser cache (Ctrl+Shift+Delete)
3. Check browser console (F12) for actual error message
4. Verify Render service is running (green status)

---

## 📱 OPTIONAL: Environment Variables (Database)

If you're using Supabase/PostgreSQL:

1. In Render dashboard, go to your service
2. Click **"Environment"** tab
3. Add these variables:
   - `SUPABASE_URL` = your_supabase_url
   - `SUPABASE_ANON_KEY` = your_supabase_key

Otherwise, the app will use CSV files (already working).

---

## ✅ SUCCESS CHECKLIST

- [ ] Created Render.com account
- [ ] Connected GitHub repository to Render
- [ ] Deployed web service (status shows "Live")
- [ ] Copied Render URL
- [ ] Updated `config.js` with Render URL
- [ ] Pushed changes to GitHub
- [ ] Waited 2 minutes for GitHub Pages to update
- [ ] Tested report generation - NO MORE ERRORS! 🎉

---

## 📞 NEED HELP?

Screenshot your Render dashboard or browser console error and share it.

---

**TL;DR:**
1. Go to render.com → Sign up with GitHub
2. Create Web Service → Connect MyISP_FunctionalTeam repo
3. Click "Create Web Service" → Wait 3 mins → Copy URL
4. Update `config.js` with your Render URL
5. Push to GitHub
6. Done! 🎉
