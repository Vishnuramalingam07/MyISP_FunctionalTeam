# MyISP Internal Tools

Internal tools application for team management, attendance tracking, and project reporting with PostgreSQL database.

## 🌐 Overview

MyISP Internal Tools is a Flask-based web application that provides centralized access to various reporting tools, attendance management, and team analytics. All data is stored in a PostgreSQL database for reliability and concurrent access.

## 📊 Available Tools

1. **Attendance Tracking** - Daily attendance management with Excel exports
2. **Team Management** - Team member details and lead assignments
3. **Daily Report** - Generate daily activity and progress reports
4. **Regression Report** - View regression test results and analysis
5. **Defect Status Report** - Track defect status and resolution metrics
6. **ADO Test Case Upload** - Upload test cases to Azure DevOps
7. **Missing Field Report** - Track and identify missing field items
8. **Screenshot Tool** - Quick screenshot capture and markup

## 🚀 Getting Started

### Prerequisites

- Python 3.10 or higher
- PostgreSQL 17+ (for database)
- Windows OS (for some features like pywin32)

### Installation

1. **Clone the Repository:**
```bash
git clone <your-repo-url>
cd MyISP_Tools
```

2. **Set Up Virtual Environment:**
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

3. **Install Dependencies:**
```powershell
pip install -r requirements.txt
```

4. **Set Up Database:**

For local PostgreSQL:
```powershell
createdb -U postgres myisp_tools
psql -U postgres -d myisp_tools -f postgres_schema.sql
```

For team network setup, see [TEAM_SETUP_GUIDE.md](TEAM_SETUP_GUIDE.md)

5. **Configure Environment:**

Create `.env` file with your database credentials:
```env
POSTGRES_HOST=localhost
POSTGRES_PORT=5432
POSTGRES_DB=myisp_tools
POSTGRES_USER=postgres
POSTGRES_PASSWORD=your_password_here
```

6. **Run the Application:**
```powershell
python app.py
```

Access at: http://localhost:8000

## 📁 Project Structure

```
MyISP_Tools/
├── app.py                      # Main Flask application
├── postgres_client.py          # Database connection module
├── postgres_schema.sql         # Database schema
├── requirements.txt            # Python dependencies
├── index.html                  # Main landing page
├── styles.css                  # Global stylesheet
├── Attendance/                 # Attendance tracking module
├── Main_Release_Daily_Status_Report/
├── Regression_Report/
├── ADO_Testcase_Upload/
├── Screenshot_Tool/
└── ...
```

## 🔧 Configuration Files

- **`.env`** - Environment variables (database credentials) - NOT committed to Git
- **`.env.example`** - Template for environment variables
- **`requirements.txt`** - Python package dependencies
- **`postgres_schema.sql`** - Database schema definition
- **`.gitignore`** - Files to exclude from Git

## 📖 Documentation

- **[POSTGRESQL_SETUP.md](POSTGRESQL_SETUP.md)** - Database setup and management guide
- **[TEAM_SETUP_GUIDE.md](TEAM_SETUP_GUIDE.md)** - Team collaboration and network access setup
- **[DEPLOYMENT.md](DEPLOYMENT.md)** - Deployment options and cloud hosting guide
- Individual feature documentation in respective folders

## 🔐 Security Considerations

- For production use, implement proper authentication
- Consider HTTPS for secure data transmission
- Validate and sanitize all user inputs
- Implement role-based access control if needed

## 📝 Future Enhancements

- Ba� Deployment Options

### Option 1: Local Network Server (Current Setup)
Run on one computer, team accesses via network.
```powershell
.\Start-All.ps1
```
Access at: http://192.168.1.2:8000

### Option 2: Cloud Deployment
Deploy to clouquestions, contact the development team.

## 📄 License

Internal use only - proprietary software.

---

## 🎯 Quick Command Reference

```powershell
# Start everything
.\Start-All.ps1

# Just start PostgreSQL
"C:\Program Files\PostgreSQL\17\bin\pg_ctl" -D "C:\Program Files\PostgreSQL\17\data" start

# Just start the app
python app.py

# Database backup
pg_dump -U postgres myisp_tools > backup.sql

# Install dependencies
pip install -r requirements.txt
```

---

**Last Updated:** April 10 Scalable solution

See [DEPLOYMENT.md](DEPLOYMENT.md) for detailed deployment instructions.

## 🔒 Security & Best Practices

### Before Committing to Git:
- ✅ Never commit `.env` files (included in `.gitignore`)
- ✅ Never commit database files
- ✅ Change default passwords before production
- ✅ Keep `.venv` folder out of Git
- ✅ Update `requirements.txt` when adding packages

### Production Deployment:
- Use HTTPS/SSL
- Enable database backups
- Configure proper firewall rules
- Implement rate limiting
- Set up monitoring and logging

## 🛡️ Database Backup

**Backup:**
```powershell
pg_dump -U postgres myisp_tools > backup.sql
```

**Restore:**
```powershell
psql -U postgres myisp_tools < backup.sql
```

## 🤝 Contributing

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## ⚠️ Important Notes

- Database server must be running for the app to work
- Team members need network access to the database server
- Keep the server computer on during work hours
- Regular database backups are essential

## 💡 Quick Start for Team

**Using Web Interface (Easiest):**
Just access: http://192.168.1.2:8000

**Running Locally with Shared Database:**
1. Get the code from Git
2. Create `.env` with server connection details
3. Install dependencies: `pip install -r requirements.txt`
4. Run: `python app.py`

## 📞