# Azure Deployment - Quick Start Guide

## CRITICAL FIX: The Problem
Your `prestart` script was trying to run `kill-port` which isn't installed on Azure (it's a dev dependency).
This caused immediate startup failure.

## What I Fixed ✅
1. **Removed prestart script** that was breaking Azure deployment
2. **Created start:dev script** for local development (kills port 3000 first)
3. **Simplified .deployment config** (Azure just needs to run `npm install --production`)

---

## Step-by-Step Deployment Instructions

### Part 1: Initialize Git (If Not Already Done)
```powershell
cd C:\Users\batan\OneDrive\Desktop\CCRW_PlainLanguage\CCRW-Annotator

# Initialize git repository
git init

# Add all files
git add .

# Create first commit
git commit -m "Initial commit - Office Add-in with secure token endpoint"
```

### Part 2: Add Azure Remote
```powershell
# Get your Azure Git URL from Portal → Deployment Center → Local Git
# It looks like: https://YOUR_APP_NAME.scm.azurewebsites.net:443/YOUR_APP_NAME.git

git remote add azure https://YOUR_APP_NAME.scm.azurewebsites.net:443/YOUR_APP_NAME.git
```

### Part 3: Configure Azure Portal (MUST DO!)

#### A. Set Startup Command
1. Azure Portal → Your App Service → **Configuration**
2. Tab: **General Settings**
3. **Startup command**: `node server.js`
4. Click **Save** (bottom) → **Continue** → **Save** (top)
5. Click **Restart**

#### B. Set Environment Variables
1. Same Configuration page → Tab: **Application settings**
2. Click **+ New application setting** for each:

| Name | Value |
|------|-------|
| `REACT_APP_DIRECT_LINE_SECRET` | `your_actual_bot_secret_key_here` |
| `SCM_DO_BUILD_DURING_DEPLOYMENT` | `true` |
| `WEBSITE_NODE_DEFAULT_VERSION` | `~20` |
| `NODE_ENV` | `production` |

3. Click **Save** → **Continue** → **Save** (top) → **Restart**

### Part 4: Deploy to Azure
```powershell
# Build the app locally first
npm run build

# Add the built files to git
git add dist/

# Commit everything
git add -A
git commit -m "Add built files and fix Azure startup"

# Push to Azure (you'll be prompted for credentials)
git push azure main

# Or if your branch is named master:
git push azure master
```

### Part 5: Monitor Deployment
After pushing, watch the deployment in real-time:

**Option A: Command Line**
```powershell
# Stream logs
az webapp log tail --name YOUR_APP_NAME --resource-group YOUR_RESOURCE_GROUP
```

**Option B: Azure Portal**
1. Go to **Deployment Center** → **Logs**
2. Click on the latest deployment to see details
3. Look for these success messages:
   ```
   ✅ Downloading deployment script
   ✅ Running 'npm install --production'
   ✅ Deployment successful
   ```

4. Go to **Log stream** (left menu)
5. Look for:
   ```
   ✅ [timestamp] Starting server...
   ✅ [timestamp] Server is running on port 8080
   ```

### Part 6: Test Your Deployment

Test these endpoints (replace YOUR_APP_NAME):

```powershell
# Health check
curl https://YOUR_APP_NAME.azurewebsites.net/health

# Expected: {"status":"OK","timestamp":"2026-03-03T..."}

# Token endpoint
curl https://YOUR_APP_NAME.azurewebsites.net/api/get-token

# Expected: {"token":"eyJhbGc..."}

# Main app page
curl https://YOUR_APP_NAME.azurewebsites.net/

# Expected: HTML with taskpane content
```

---

## Local Development (Use This Instead of npm start)

For local development, use the new command that kills port 3000 first:

```powershell
cd CCRW-Annotator

# Option 1: Use the new dev script
npm run start:dev

# Option 2: Manual cleanup then start
Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id $_ -Force }
npm start
```

---

## Troubleshooting

### If You Still See "Application Error"

#### Issue 1: Check Startup Command
```
Portal → Configuration → General Settings → Startup command
Make sure it says: node server.js
Save → Restart
```

#### Issue 2: Check Environment Variable
```
Portal → Configuration → Application settings
Find: REACT_APP_DIRECT_LINE_SECRET
If missing or wrong, fix it → Save → Restart
```

#### Issue 3: Check Deployment Logs
```
Portal → Deployment Center → Logs
Look for errors like:
- "npm install failed" → SCM_DO_BUILD_DURING_DEPLOYMENT not set
- "Cannot find module" → npm install didn't run or failed
```

#### Issue 4: Check Runtime Logs
```
Portal → Log stream
Look for the server startup messages
If you see "Error: Cannot find module 'express'", it means npm install didn't run
```

#### Issue 5: Manual Verification via Kudu
1. Portal → **Advanced Tools** → **Go** (opens Kudu)
2. Click **Debug console** → **PowerShell**
3. Navigate to `D:\home\site\wwwroot`
4. Run: `ls` to verify files are there
5. Run: `node server.js` to test startup manually
6. Look for the error message

---

## Files Required for Deployment

Make sure these files are committed:
- ✅ `server.js` (entry point)
- ✅ `package.json` (dependencies list)
- ✅ `package-lock.json` (exact versions)
- ✅ `web.config` (IIS configuration)
- ✅ `.deployment` (build configuration)
- ✅ `dist/` folder with built files
- ❌ `.env` (DO NOT commit - use Azure portal settings)
- ❌ `node_modules/` (DO NOT commit - Azure installs these)

---

## Quick Checklist Before Deploy

- [ ] Built app locally: `npm run build`
- [ ] Verified dist/ folder has files
- [ ] Removed prestart script (already done ✅)
- [ ] Committed all changes to git
- [ ] Set Startup Command in Azure: `node server.js`
- [ ] Set `REACT_APP_DIRECT_LINE_SECRET` in Azure Portal
- [ ] Set `SCM_DO_BUILD_DURING_DEPLOYMENT=true` in Azure Portal
- [ ] Pushed to Azure: `git push azure main`
- [ ] Checked deployment logs: Deployment Center → Logs
- [ ] Checked runtime logs: Log stream
- [ ] Tested: `https://YOUR_APP.azurewebsites.net/health`

---

## Success Indicators

After successful deployment, you should see in **Log stream**:
```
[2026-03-03T00:26:10.529Z] Starting server...
[2026-03-03T00:26:10.540Z] PORT=8080
[2026-03-03T00:26:10.550Z] NODE_ENV=production
[2026-03-03T00:26:10.560Z] Has REACT_APP_DIRECT_LINE_SECRET: true
[2026-03-03T00:26:10.580Z] Server is running on port 8080
```

And visiting `https://YOUR_APP.azurewebsites.net/` shows your taskpane! ✅
