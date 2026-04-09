# Deployment Guide for MyISP Tools

## 📋 Overview

This guide covers various deployment options for MyISP Tools, from local network hosting to cloud deployment.

---

## 🏠 Option 1: Local Network Server (Current Setup)

**Best for:** Small teams, local network access

### Current Configuration:
- **Database:** PostgreSQL running on your computer (192.168.1.2)
- **App:** Flask development server
- **Access:** Local network only (192.168.1.x)

### Pros:
✅ No cloud costs
✅ Full data control
✅ Easy setup (already done!)
✅ Fast for local team

### Cons:
❌ Server computer must stay on
❌ Not accessible outside network
❌ Limited to development server performance

### Running 24/7:
```powershell
# Use the startup script
.\Start-All.ps1

# Or run with production server (Waitress)
python run_production.py
```

### Making it Auto-Start:
Create a Windows Task Scheduler job to run `Start-All.ps1` at system startup.

---

## ☁️ Option 2: Azure App Service (Recommended for Enterprise)

**Best for:** Enterprise deployment, scalability, Microsoft ecosystem

### Requirements:
- Azure subscription
- Azure Database for PostgreSQL

### Steps:

1. **Create Azure Database for PostgreSQL:**
   ```bash
   az postgres flexible-server create \
     --resource-group myisp-rg \
     --name myisp-db \
     --location eastus \
     --admin-user dbadmin \
     --admin-password <secure-password> \
     --sku-name Standard_B1ms
   ```

2. **Migrate Database:**
   ```bash
   # Export from local
   pg_dump -U postgres myisp_tools > myisp_backup.sql
   
   # Import to Azure
   psql -h myisp-db.postgres.database.azure.com \
        -U dbadmin \
        -d postgres < myisp_backup.sql
   ```

3. **Create App Service:**
   ```bash
   az webapp up \
     --resource-group myisp-rg \
     --name myisp-tools \
     --runtime "PYTHON:3.11" \
     --sku B1
   ```

4. **Configure Environment Variables:**
   ```bash
   az webapp config appsettings set \
     --resource-group myisp-rg \
     --name myisp-tools \
     --settings \
       POSTGRES_HOST=myisp-db.postgres.database.azure.com \
       POSTGRES_DB=myisp_tools \
       POSTGRES_USER=dbadmin \
       POSTGRES_PASSWORD=<password>
   ```

### Cost Estimate:
- App Service (B1): ~$13/month
- PostgreSQL (B1): ~$25/month
- **Total:** ~$38/month

---

## 🟣 Option 3: Heroku (Easy Cloud Deployment)

**Best for:** Quick cloud deployment, small teams

### Steps:

1. **Install Heroku CLI:**
   ```bash
   # Download from https://devcenter.heroku.com/articles/heroku-cli
   ```

2. **Create Heroku App:**
   ```bash
   heroku create myisp-tools
   heroku addons:create heroku-postgresql:essential-0
   ```

3. **Deploy:**
   ```bash
   git push heroku main
   ```

4. **Run Database Migration:**
   ```bash
   heroku run python -c "from postgres_client import postgres; print('DB Connected')"
   ```

### Cost:
- App: $7/month (Eco dyno)
- Database: $5/month (Essential-0)
- **Total:** ~$12/month

---

## 🐳 Option 4: Docker + Any Cloud Provider

**Best for:** Flexibility, portability

### Create Dockerfile:

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Expose port
EXPOSE 8000

# Run application
CMD ["python", "run_production.py"]
```

### Docker Compose (with PostgreSQL):

```yaml
version: '3.8'

services:
  web:
    build: .
    ports:
      - "8000:8000"
    environment:
      - POSTGRES_HOST=db
      - POSTGRES_DB=myisp_tools
      - POSTGRES_USER=postgres
      - POSTGRES_PASSWORD=postgres123
    depends_on:
      - db
  
  db:
    image: postgres:17
    environment:
      - POSTGRES_DB=myisp_tools
      - POSTGRES_USER=postgres
      - POSTGRES_PASSWORD=postgres123
    volumes:
      - postgres_data:/var/lib/postgresql/data
      - ./postgres_schema.sql:/docker-entrypoint-initdb.d/schema.sql

volumes:
  postgres_data:
```

### Deploy to:
- **DigitalOcean App Platform:** ~$12/month
- **AWS ECS:** Variable pricing
- **Google Cloud Run:** Pay per use
- **Railway:** ~$5/month

---

## 🌐 Option 5: Windows Server (On-Premise)

**Best for:** Corporate environment, existing infrastructure

### Requirements:
- Windows Server 2019/2022
- IIS (Internet Information Services)
- PostgreSQL installed

### Setup:

1. **Install IIS with Python support**
2. **Configure as Windows Service**
3. **Set up reverse proxy in IIS**
4. **Configure SSL certificates**

See Microsoft documentation for detailed IIS + Python setup.

---

## 📊 Deployment Comparison

| Option | Cost/Month | Complexity | Uptime | Scalability |
|--------|-----------|------------|--------|-------------|
| Local Network | Free | Low | Manual | Limited |
| Azure | ~$38 | Medium | 99.9% | High |
| Heroku | ~$12 | Low | 99.9% | Medium |
| Docker (DO) | ~$12 | Medium | 99.9% | Medium |
| Windows Server | Variable | High | Varies | Medium |

---

## 🔒 Security Checklist (Production)

Before deploying to production:

- [ ] Change all default passwords
- [ ] Enable HTTPS/SSL
- [ ] Set up database backups
- [ ] Configure firewall rules
- [ ] Enable database connection pooling
- [ ] Set up monitoring and logging
- [ ] Implement rate limiting
- [ ] Use environment variables for secrets
- [ ] Enable CORS only for trusted domains
- [ ] Set up user authentication (not just IP-based)

---

## 📦 Pre-Deployment Steps

### 1. Update requirements.txt
```powershell
pip freeze > requirements.txt
```

### 2. Test Production Server Locally
```powershell
python run_production.py
```

### 3. Create Database Backup
```powershell
pg_dump -U postgres myisp_tools > backup_before_deploy.sql
```

### 4. Update Configuration
Create production `.env`:
```env
POSTGRES_HOST=<production-host>
POSTGRES_PORT=5432
POSTGRES_DB=myisp_tools
POSTGRES_USER=<production-user>
POSTGRES_PASSWORD=<secure-password>
FLASK_ENV=production
SECRET_KEY=<generate-random-secret>
```

---

## 🔄 Database Migration Strategy

### From Local to Cloud:

1. **Backup Local Database:**
   ```powershell
   pg_dump -U postgres myisp_tools > local_backup.sql
   ```

2. **Create Cloud Database**

3. **Restore to Cloud:**
   ```powershell
   psql -h <cloud-host> -U <user> -d myisp_tools < local_backup.sql
   ```

4. **Update .env in Application**

5. **Test Connection**

6. **Update Team with New Connection Details**

---

## 📈 Monitoring & Maintenance

### Recommended Tools:
- **Database:** pg_stat_statements for query analysis
- **Application:** Flask logging to file/service
- **Uptime:** UptimeRobot (free tier)
- **Errors:** Sentry (error tracking)

### Regular Tasks:
- Weekly database backups
- Monthly dependency updates
- Quarterly security audits
- Monitor disk space and performance

---

## 🆘 Rollback Plan

If deployment fails:

1. Keep local database running
2. Point team back to local server (192.168.1.2:8000)
3. Restore from backup if needed
4. Debug issues in staging environment

---

## 💡 Recommendations

**For Your Current Setup:**

1. **Immediate (Stay Local):**
   - Use `run_production.py` instead of development server
   - Set up automatic database backups
   - Create Windows Task Scheduler for auto-start

2. **Short-term (3-6 months):**
   - Consider Azure if team grows
   - Set up cloud backup solution
   - Implement proper authentication

3. **Long-term:**
   - Move to cloud for reliability
   - Implement CI/CD pipeline
   - Add monitoring and alerting

---

## 📞 Need Help?

- **Azure:** https://docs.microsoft.com/azure/app-service/
- **Heroku:** https://devcenter.heroku.com/
- **Docker:** https://docs.docker.com/
- **PostgreSQL:** https://www.postgresql.org/docs/

---

**Current Status:** Your app is ready to deploy! Choose the option that fits your needs and budget.
