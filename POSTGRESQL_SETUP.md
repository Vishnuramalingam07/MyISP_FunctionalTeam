# ============================================================
# PostgreSQL Setup Guide for MyISP Tools
# ============================================================

## ✅ Installation Complete!

Your MyISP Tools application is now using a local PostgreSQL database.

### 📋 Summary

**Database Details:**
- Host: localhost
- Port: 5432
- Database: myisp_tools
- Username: postgres
- Password: postgres123

**What Was Set Up:**
1. ✓ PostgreSQL 17.9 installed
2. ✓ Database 'myisp_tools' created
3. ✓ Schema tables created (authorized_users, team_members, attendance_records, attendance_logs)
4. ✓ App updated to use PostgreSQL instead of Supabase
5. ✓ Python PostgreSQL adapter (psycopg2) installed

---

## 🚀 Starting Your Application

### Every time you restart your computer:

1. **Start PostgreSQL:**
   ```powershell
   "C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" start
   ```

2. **Start your Flask app:**
   ```powershell
   cd C:\Users\vishnu.ramalingam\MyISP_Tools
   .venv\Scripts\python.exe app.py
   ```

3. **Access your application:**
   - Local: http://localhost:8000
   - Network: http://192.168.1.2:8000

### Quick start script (recommended):

Run this PowerShell script to start everything:
```powershell
.\Start-All.ps1
```

---

## 🛠️ Database Management

### Connect to Database:
```powershell
$env:PGPASSWORD="postgres123"
"C:\Program Files\PostgreSQL\17\bin\psql" -U postgres -d myisp_tools
```

### Useful SQL Commands:
```sql
-- List all tables
\dt

-- View table structure
\d attendance_records

-- Count records
SELECT COUNT(*) FROM attendance_records;

-- View recent logs
SELECT * FROM attendance_logs ORDER BY saved_at DESC LIMIT 10;

-- Exit psql
\q
```

### Backup Database:
```powershell
$env:PGPASSWORD="postgres123"
"C:\Program Files\PostgreSQL\17\bin\pg_dump" -U postgres myisp_tools > backup.sql
```

### Restore Database:
```powershell
$env:PGPASSWORD="postgres123"
"C:\Program Files\PostgreSQL\17\bin\psql" -U postgres myisp_tools < backup.sql
```

---

## 🔧 Configuration

To change database settings, create a `.env` file:

```env
POSTGRES_HOST=localhost
POSTGRES_PORT=5432
POSTGRES_DB=myisp_tools
POSTGRES_USER=postgres
POSTGRES_PASSWORD=postgres123
```

---

## 📊 Database Schema

### Tables Created:

1. **authorized_users** - User access control
2. **team_members** - Team member details
3. **attendance_records** - Daily attendance data
4. **attendance_logs** - Audit trail of changes

---

## ⚠️ Troubleshooting

### App says "Database disabled":
Check if PostgreSQL is running:
```powershell
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" status
```

If not running, start it:
```powershell
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" start
```

### Connection refused error:
1. Verify PostgreSQL is running (see above)
2. Check firewall settings
3. Verify password is correct

### Permission errors:
Run PowerShell as Administrator for database operations.

---

## 🔒 Security Notes

- Default password is `postgres123` - **change this for production use**
- Database is only accessible from localhost by default
- Consider setting up Windows Service for auto-start

---

## 📈 Next Steps

1. Migrate existing CSV/Excel data to PostgreSQL (optional)
2. Set up automatic PostgreSQL startup as Windows Service
3. Configure backups
4. Change default password for production

---

## 💡 Benefits of PostgreSQL vs CSV

✅ Concurrent user access  
✅ ACID compliance (data integrity)  
✅ Better performance for queries  
✅ Built-in backup/restore tools  
✅ Advanced querying capabilities  
✅ Audit logs and transaction history  

---

For questions or issues, refer to:
- PostgreSQL docs: https://www.postgresql.org/docs/
- App source: app.py and postgres_client.py
