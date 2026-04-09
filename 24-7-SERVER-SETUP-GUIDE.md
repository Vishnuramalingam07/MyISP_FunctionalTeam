# 24/7 Server Setup Guide - MyISP Tools

## Quick Setup (Recommended)

### **Method 1: Automated Setup (Easiest)**
1. Right-click on `Setup-24-7-Server.bat`
2. Select **"Run as administrator"**
3. Press any key to confirm
4. Done! Server will run 24/7

**What this does:**
- Creates a Windows scheduled task that starts on boot
- Server runs in background (no visible window)
- Automatically restarts if it crashes
- Available at: http://192.168.1.2:8000

**To remove:**
- Run `Remove-24-7-Server.bat` as administrator

---

## Alternative Methods

### **Method 2: Windows Service (Most Reliable)**

1. **Install NSSM (Non-Sucking Service Manager)**
   ```powershell
   # Download from: https://nssm.cc/download
   # Or use chocolatey:
   choco install nssm
   ```

2. **Create the service:**
   ```powershell
   cd C:\Users\vishnu.ramalingam\MyISP_Tools
   
   nssm install MyISPTools python.exe
   nssm set MyISPTools AppDirectory "C:\Users\vishnu.ramalingam\MyISP_Tools"
   nssm set MyISPTools AppParameters "app.py"
   nssm set MyISPTools DisplayName "MyISP Internal Tools Server"
   nssm set MyISPTools Description "Flask server for MyISP team tools and reports"
   nssm set MyISPTools Start SERVICE_AUTO_START
   ```

3. **Start the service:**
   ```powershell
   nssm start MyISPTools
   ```

**Benefits:**
- True Windows service
- Automatic restart on failure
- Runs under SYSTEM account
- Better logging and monitoring

**To manage:**
```powershell
nssm stop MyISPTools        # Stop service
nssm restart MyISPTools     # Restart service
nssm remove MyISPTools      # Remove service
nssm edit MyISPTools        # Edit configuration
```

---

### **Method 3: Python Process Manager (PM2 Alternative)**

1. **Install waitress (production WSGI server):**
   ```powershell
   pip install waitress
   ```

2. **Create production server file:**
   ```python
   # Save as: run_production.py
   from waitress import serve
   from app import app
   
   if __name__ == '__main__':
       print("Starting MyISP Tools Server on http://0.0.0.0:8000")
       serve(app, host='0.0.0.0', port=8000, threads=4)
   ```

3. **Use Task Scheduler:**
   - Open Task Scheduler
   - Create Basic Task
   - Name: "MyISP Tools Server"
   - Trigger: "When the computer starts"
   - Action: "Start a program"
   - Program: `C:\Users\vishnu.ramalingam\MyISP_Tools\.venv\Scripts\python.exe`
   - Arguments: `run_production.py`
   - Start in: `C:\Users\vishnu.ramalingam\MyISP_Tools`
   - Check: "Run with highest privileges"

---

### **Method 4: Docker Container (Advanced)**

1. **Create Dockerfile:**
   ```dockerfile
   FROM python:3.13-slim
   
   WORKDIR /app
   
   COPY requirements.txt .
   RUN pip install --no-cache-dir -r requirements.txt
   
   COPY . .
   
   EXPOSE 8000
   
   CMD ["python", "app.py"]
   ```

2. **Create docker-compose.yml:**
   ```yaml
   version: '3.8'
   services:
     myisp-tools:
       build: .
       ports:
         - "8000:8000"
       restart: always
       volumes:
         - ./data:/app/data
   ```

3. **Run:**
   ```powershell
   docker-compose up -d
   ```

---

## Monitoring & Maintenance

### Check if server is running:
```powershell
# Check process
tasklist | findstr python

# Check port
netstat -ano | findstr :8000

# Test server
curl http://localhost:8000
```

### View logs:
- Task Scheduler method: Check Event Viewer → Windows Logs → Application
- NSSM method: Check `C:\Users\vishnu.ramalingam\MyISP_Tools\nssm_logs\`

### Restart server:
```powershell
# Task Scheduler method
schtasks /run /tn "MyISP_Tools_Server"

# NSSM method
nssm restart MyISPTools
```

---

## Firewall Configuration

To allow team access (if needed):
```powershell
# Allow incoming connections on port 8000
netsh advfirewall firewall add rule name="MyISP Tools Server" dir=in action=allow protocol=TCP localport=8000

# Remove rule
netsh advfirewall firewall delete rule name="MyISP Tools Server"
```

---

## Troubleshooting

### Server not starting:
1. Check if port 8000 is already in use:
   ```powershell
   netstat -ano | findstr :8000
   ```

2. Check Python is in PATH:
   ```powershell
   python --version
   ```

3. Verify Flask is installed:
   ```powershell
   pip list | findstr Flask
   ```

### Can't access from other computers:
1. Check Windows Firewall
2. Verify IP address: `ipconfig`
3. Ensure server is listening on 0.0.0.0, not 127.0.0.1

### Server crashes:
1. Check Event Viewer logs
2. Add error logging to app.py
3. Use production WSGI server (waitress or gunicorn)

---

## Recommendations

**For Development:** Use the regular `Start-Server.bat`

**For Production/24x7:**
1. **Best:** Method 2 (Windows Service with NSSM) - Most reliable
2. **Good:** Method 1 (Automated Setup) - Quick and easy
3. **Better Performance:** Method 3 (Waitress + Task Scheduler)

---

## Security Considerations

1. **Don't expose to internet** without proper authentication
2. **Use HTTPS** for sensitive data (requires SSL certificate)
3. **Set up authentication** in Flask app
4. **Regular updates** - Keep Flask and dependencies updated
5. **Backup data** regularly

---

## Current Status

After running `Setup-24-7-Server.bat`:
- ✅ Server starts automatically on system boot
- ✅ Runs in background (no window)
- ✅ Accessible at: http://192.168.1.2:8000
- ✅ Team can access from any browser on the network

To verify it's working:
```powershell
# Check scheduled task
schtasks /query /tn "MyISP_Tools_Server"

# Test server
curl http://localhost:8000
```
