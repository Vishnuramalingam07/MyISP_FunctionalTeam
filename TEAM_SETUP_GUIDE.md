# ============================================================
# Team Setup Guide - PostgreSQL Shared Database Access
# ============================================================

## 🌐 Database Server Information

**Server Details:**
- **Host:** 192.168.1.2
- **Port:** 5432
- **Database:** myisp_tools
- **Username:** postgres
- **Password:** postgres123

**Server Computer:** vishnu.ramalingam's machine
**Network:** Must be on the same local network (192.168.x.x)

---

## 👥 For Team Members - Application Setup

### Option 1: Use the Web Application (Easiest)

Simply access the MyISP Tools web interface in your browser:

**URL:** http://192.168.1.2:8000

That's it! No installation needed. All data is automatically saved to the shared PostgreSQL database.

**Important:** The server computer (192.168.1.2) must be:
- ✓ Powered on
- ✓ Connected to the network
- ✓ Running the MyISP Tools application

---

### Option 2: Run Application Locally with Shared Database

If you want to run the application on your own computer but use the shared database:

#### Step 1: Get the Application Files
Copy the MyISP_Tools folder from the server or download from your shared location.

#### Step 2: Install Python (if not already installed)
Download from: https://www.python.org/downloads/
- Install Python 3.10 or higher
- Check "Add to PATH" during installation

#### Step 3: Set Up the Application

1. Open PowerShell in the MyISP_Tools folder
2. Create virtual environment:
   ```powershell
   python -m venv .venv
   ```

3. Activate virtual environment:
   ```powershell
   .venv\Scripts\Activate.ps1
   ```

4. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```

#### Step 4: Configure Database Connection

Create a `.env` file in the MyISP_Tools folder with:

```env
POSTGRES_HOST=192.168.1.2
POSTGRES_PORT=5432
POSTGRES_DB=myisp_tools
POSTGRES_USER=postgres
POSTGRES_PASSWORD=postgres123
```

#### Step 5: Run the Application

```powershell
python app.py
```

Access at: http://localhost:8000

---

## 🔧 For Database Administrator (Server Owner)

### Starting the Database Server

**Option 1 - Use the startup script:**
```powershell
.\Start-All.ps1
```

**Option 2 - Manual start:**
```powershell
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" start
python app.py
```

### Stopping the Database Server

```powershell
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" stop
```

### Check Database Status

```powershell
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" status
```

---

## 🔥 Firewall Configuration (One-Time Setup)

**IMPORTANT:** The server computer needs to allow incoming connections.

**Run this once (as Administrator):**
1. Right-click `Configure-Firewall.ps1`
2. Select "Run with PowerShell as Administrator"
3. Follow the prompts

This opens port 5432 for PostgreSQL connections from your local network.

---

## 📊 Direct Database Access (For Advanced Users)

### Using pgAdmin (GUI Tool)

1. Download pgAdmin from: https://www.pgadmin.org/download/
2. Install and open pgAdmin
3. Right-click "Servers" → "Register" → "Server"
4. Enter connection details:
   - Name: MyISP Tools Database
   - Host: 192.168.1.2
   - Port: 5432
   - Database: myisp_tools
   - Username: postgres
   - Password: postgres123

### Using psql (Command Line)

```powershell
# Set password environment variable
$env:PGPASSWORD="postgres123"

# Connect to database
psql -h 192.168.1.2 -U postgres -d myisp_tools

# Example queries
SELECT COUNT(*) FROM attendance_records;
SELECT * FROM team_members;
\q  # to exit
```

---

## ⚠️ Important Notes for Team

### ✅ DO:
- Keep the server computer on during work hours
- Ensure you're on the same network as the server
- Use the web interface (http://192.168.1.2:8000) for easiest access
- Report connection issues immediately

### ❌ DON'T:
- Don't manually modify database files
- Don't share the database password outside the team
- Don't try to connect from outside your local network
- Don't accidentally shut down the server computer

---

## 🔍 Troubleshooting

### "Cannot connect to database" error

**Check 1:** Is the server computer on and connected to network?
```powershell
ping 192.168.1.2
```

**Check 2:** Is PostgreSQL running on the server?
Ask the server owner to check status.

**Check 3:** Are you on the same network?
Your IP should be 192.168.1.x or similar.

**Check 4:** Is the web app running?
Try accessing http://192.168.1.2:8000 in your browser.

### "Connection timed out" error

- Check Windows Firewall on server computer
- Run `Configure-Firewall.ps1` as Administrator on server
- Verify port 5432 is open

### "Password authentication failed" error

- Verify you're using the correct password: `postgres123`
- Check your `.env` file has correct credentials

---

## 📞 Support

**Server Administrator:** vishnu.ramalingam
**Server Computer:** 192.168.1.2

For issues:
1. Check this guide first
2. Try restarting the application
3. Contact the server administrator

---

## 🎯 Quick Reference Card

```
┌─────────────────────────────────────────────────┐
│  MyISP Tools - Quick Connection Info            │
├─────────────────────────────────────────────────┤
│  Web App:  http://192.168.1.2:8000             │
│  DB Host:  192.168.1.2                          │
│  DB Port:  5432                                 │
│  Database: myisp_tools                          │
│  User:     postgres                             │
│  Password: postgres123                          │
└─────────────────────────────────────────────────┘
```

**For most users:** Just bookmark http://192.168.1.2:8000 and use the web interface!
