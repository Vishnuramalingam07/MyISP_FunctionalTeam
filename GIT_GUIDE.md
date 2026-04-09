# Git Setup and Deployment Guide

## 🎯 Quick Start: Push Your Code to Git

### Step 1: Initialize Git Repository (if not already done)

```powershell
cd C:\Users\vishnu.ramalingam\MyISP_Tools

# Initialize Git
git init

# Add all files (respecting .gitignore)
git add .

# Create first commit
git commit -m "Initial commit: MyISP Tools with PostgreSQL database"
```

### Step 2: Create Remote Repository

**Option A: GitHub**
1. Go to https://github.com
2. Click "New Repository"
3. Name it: `MyISP_Tools` (or your preferred name)
4. Choose Private or Public
5. **DO NOT** initialize with README (you already have one)
6. Click "Create Repository"

**Option B: Azure DevOps**
1. Go to https://dev.azure.com
2. Create new project
3. Go to Repos → Files
4. Copy the repository URL

**Option C: GitLab**
1. Go to https://gitlab.com
2. New Project → Create blank project
3. Copy the repository URL

### Step 3: Connect and Push

```powershell
# Add remote (replace with your actual URL)
git remote add origin https://github.com/yourusername/MyISP_Tools.git

# Push to remote
git push -u origin main

# Or if your branch is named master:
git push -u origin master
```

---

## 🔒 IMPORTANT: Security Checklist Before Pushing

### ✅ Files That SHOULD Be in Git:
- ✓ Source code (.py files)
- ✓ HTML/CSS/JavaScript files
- ✓ Database schema (postgres_schema.sql)
- ✓ requirements.txt
- ✓ README.md and documentation
- ✓ .gitignore
- ✓ .env.example (template only, NO real passwords)

### ❌ Files That SHOULD NOT Be in Git:
- ❌ .env (contains real passwords)
- ❌ Database files (.db, data/)
- ❌ Virtual environment (.venv/)
- ❌ __pycache__/
- ❌ *.log files
- ❌ User data (Master_Attendance.xlsx, etc.)
- ❌ Uploaded files

**Verify your .gitignore is correct!**

```powershell
# Check what will be committed
git status

# Check for sensitive files
git status | Select-String ".env"  # Should show nothing or only .env.example
```

---

## 📦 Making Code Changes

```powershell
# Check current status
git status

# Add specific files
git add file1.py file2.py

# Or add all changed files
git add .

# Commit with message
git commit -m "Description of your changes"

# Push to remote
git push

# Pull latest changes from team
git pull
```

---

## 🌿 Branching Strategy

### For Feature Development:

```powershell
# Create feature branch
git checkout -b feature/attendance-improvements

# Make your changes
# ... edit files ...

# Commit changes
git add .
git commit -m "Add new attendance features"

# Push feature branch
git push -u origin feature/attendance-improvements

# Merge back to main (via Pull Request on GitHub/Azure)
```

### For Bug Fixes:

```powershell
# Create bugfix branch
git checkout -b bugfix/fix-login-issue

# Make fixes and commit
git add .
git commit -m "Fix login authentication issue"

# Push
git push -u origin bugfix/fix-login-issue
```

---

## 🚀 Deployment Scenarios

### Scenario 1: Team Member Clones Repository

```powershell
# Clone repository
git clone https://github.com/yourusername/MyISP_Tools.git
cd MyISP_Tools

# Set up virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Create .env file (copy from .env.example)
copy .env.example .env
# Edit .env with actual database credentials

# Run application
python app.py
```

### Scenario 2: Deploy to Cloud

See [DEPLOYMENT.md](DEPLOYMENT.md) for:
- Azure App Service deployment
- Heroku deployment
- Docker deployment
- Other cloud options

---

## 🔄 Keeping Your Code Updated

```powershell
# Before starting work, get latest changes
git pull

# After making changes
git add .
git commit -m "Your change description"
git push

# If there are conflicts
git pull
# Resolve conflicts in files
git add .
git commit -m "Merge conflicts resolved"
git push
```

---

## 📋 Common Git Commands

```powershell
# Check status
git status

# See commit history
git log --oneline

# See what changed
git diff

# Undo uncommitted changes
git checkout -- filename.py

# Undo last commit (keep changes)
git reset --soft HEAD~1

# Create .gitignore
# (already done - see .gitignore file)

# Add remote
git remote add origin <url>

# View remotes
git remote -v

# Clone repository
git clone <url>
```

---

## ⚠️ Database Deployment Considerations

### Important Notes:

1. **Database is NOT in Git**
   - Only schema (postgres_schema.sql) is tracked
   - Actual data stays on the server
   
2. **For Production Deployment:**
   - Set up NEW PostgreSQL database in cloud/production
   - Run postgres_schema.sql to create tables
   - Migrate data if needed (use pg_dump)
   
3. **Team Members Connecting:**
   - They DON'T need their own PostgreSQL
   - They connect to YOUR database (192.168.1.2)
   - Just need correct .env configuration

---

## 🎯 Recommended Workflow

### For You (Database Owner):
```powershell
# Daily workflow
1. .\Start-All.ps1  # Start PostgreSQL and app
2. # Work on code
3. git add .
4. git commit -m "Description"
5. git push
```

### For Team Members:
```powershell
# First time setup
1. git clone <repo-url>
2. pip install -r requirements.txt
3. Create .env with: POSTGRES_HOST=192.168.1.2
4. python app.py

# Daily workflow
1. git pull  # Get latest changes
2. python app.py  # Run application
```

---

## 🔐 Protecting Secrets

### Never Commit These to Git:
- Actual passwords
- API keys
- Database credentials (in .env)
- SSL certificates
- Private keys

### If You Accidentally Committed a Secret:

```powershell
# Remove file from Git (but keep local copy)
git rm --cached .env

# Add to .gitignore
echo ".env" >> .gitignore

# Commit the removal
git add .gitignore
git commit -m "Remove .env from tracking"
git push

# IMPORTANT: Change the password immediately!
# The secret is still in Git history
```

To completely remove from history (advanced):
```powershell
# Use BFG Repo-Cleaner or git filter-branch
# See: https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository
```

---

## 📖 Additional Resources

- **Git Documentation:** https://git-scm.com/doc
- **GitHub Guides:** https://guides.github.com
- **Azure DevOps Repos:** https://docs.microsoft.com/azure/devops/repos/
- **Git Cheat Sheet:** https://training.github.com/downloads/github-git-cheat-sheet.pdf

---

## ✅ Pre-Push Checklist

Before pushing to Git:

- [ ] Checked `git status` for unwanted files
- [ ] Verified .env is NOT in the commit
- [ ] Updated requirements.txt if you added packages
- [ ] Tested code locally
- [ ] Written meaningful commit message
- [ ] Checked .gitignore is working correctly
- [ ] No passwords or secrets in code
- [ ] Documentation updated if needed

---

**Ready to push? Your code is prepared for Git!** 🚀
