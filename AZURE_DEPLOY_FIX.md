# Azure Deployment Guide - CCRW Annotator

## Problem Diagnosis
You're seeing "Application Error" on Azure because:
1. `npm install` wasn't running during deployment
2. Server couldn't start without required dependencies
3. Startup command wasn't configured

## Solution Overview

### Local Verification ✅
Server is running locally and responding to:
- `http://localhost:3000/health` → `{"status":"OK"}`
- `http://localhost:3000/api/get-token` → Returns valid JWT token

### Azure Deployment Checklist

#### 1. **Must Do: Commit These Files**
Before deploying, ensure these files are in your Git repository:
```
✅ server.js                  (main entry point)
✅ package.json               (dependency list)
✅ package-lock.json          (exact versions)
✅ .deployment                (build configuration)
✅ web.config                 (IIS configuration for Azure)
✅ dist/                       (built taskpane files)
✅ .env                        (local development only)
```

**Commit changes:**
```bash
cd CCRW-Annotator
git add .
git commit -m "Fix Azure deployment - add web.config, improve server logging, ensure npm install"
git push origin main
```

#### 2. **Azure Portal Configuration**

**Step A: Set Startup Command**
1. Go to **Azure Portal** → Your App Service → **Configuration**
2. Click on **General Settings** tab
3. Find **Startup command**
4. Enter: `node server.js`
5. Click **Save** → **Continue** → **Save** (at the top)
6. **Restart** the app service

**Step B: Configure Application Settings**
1. Click on **Application settings** tab (same Configuration page)
2. Ensure these environment variables exist:

| Name | Value | Notes |
|------|-------|-------|
| `REACT_APP_DIRECT_LINE_SECRET` | `your_bot_secret_key` | ⚠️ Must match your Bot Framework direct line secret |
| `NODE_ENV` | `production` | Optional but recommended |
| `WEBSITE_NODE_DEFAULT_VERSION` | `~20` | Recommended Node version |
| `SCM_DO_BUILD_DURING_DEPLOYMENT` | `true` | Forces npm install & build |

3. Click **Save**
4. Click **Restart** (at the top of page)

#### 3. **Deploy to Azure**

**Option A: Git Push Deployment (Recommended)**
```bash
# From your repo root
cd CCRW-Annotator

# Make sure everything is committed
git status

# Push to Azure remote
git push azure main
```

**Option B: Manual ZIP Deployment**
```bash
# Create a zip file with everything
Compress-Archive -Path "CCRW-Annotator/*" -DestinationPath "deploy.zip" -Force

# Deploy the zip
az webapp deployment source config-zip `
  --resource-group YOUR_RESOURCE_GROUP `
  --name YOUR_APP_SERVICE_NAME `
  --src deploy.zip
```

#### 4. **Monitor Deployment Logs**

After deploying, check logs in Azure:
1. Go to **Deployment Center** → **Logs**
2. Look for these signs of success:
   ```
   ✅ Downloading deployment script
   ✅ Running npm install
   ✅ Running npm run build
   ✅ Deployment successful
   ```
3. Check **Log stream** (real-time):
   - Go to **App Service logs** in left menu
   - Look for: `"Server is running on port 8080"`

#### 5. **Verify Deployment**

Test these endpoints (replace YOUR_APP_NAME):

```bash
# Health check
curl https://YOUR_APP_NAME.azurewebsites.net/health

# Token endpoint
curl https://YOUR_APP_NAME.azurewebsites.net/api/get-token

# Main app
curl https://YOUR_APP_NAME.azurewebsites.net/
```

Expected responses:
- `GET /health` → `{"status":"OK","timestamp":"..."}`
- `GET /api/get-token` → `{"token":"eyJ..."}`
- `GET /` → HTML taskpane page

#### 6. **Troubleshooting**

If you still see "Application Error":

**Check 1: View Error Details**
1. In App Service, go to **Advanced Tools** → **Go**
2. Click **Debug Console** → **PowerShell**
3. Navigate to `D:\home\site\wwwroot`
4. Run: `node server.js` (to see exact error)

**Check 2: Verify Startup Command**
1. Go to **Configuration** → **General Settings**
2. Confirm **Startup command** is: `node server.js`
3. Save and Restart

**Check 3: Check Logs for Missing Module**
1. Go to **Log stream**
2. Search for: `Cannot find module`
3. If found, it means npm install didn't run
4. Go back to Configuration → Application settings
5. Verify `SCM_DO_BUILD_DURING_DEPLOYMENT=true`
6. Click Save → Restart → Check logs again

**Check 4: Verify Environment Variable**
1. Go to **Configuration** → **Application settings**
2. Scroll to find: `REACT_APP_DIRECT_LINE_SECRET`
3. If missing, add it with your actual secret
4. Click Save → Restart

#### 7. **New Features for Better Diagnostics**

The updated `server.js` now logs:
- ✅ Server startup timestamp
- ✅ PORT being used
- ✅ Whether REACT_APP_DIRECT_LINE_SECRET is configured
- ✅ Health check endpoint for monitoring
- ✅ Detailed error messages if startup fails

### Quick Checklist Before Redeployment

- [ ] All files committed to Git (`git status` is clean)
- [ ] `.deployment` file has `SCM_DO_BUILD_DURING_DEPLOYMENT=true`
- [ ] `package.json` has all required dependencies
- [ ] `web.config` is present (for IIS integration)
- [ ] `REACT_APP_DIRECT_LINE_SECRET` is set in Azure Portal
- [ ] Startup command is set to `node server.js`
- [ ] App service has been restarted after configuration

### Success Indicators

After deployment, you should see in **Log stream**:
```
[2026-03-03T00:26:10.529Z] Starting server...
[2026-03-03T00:26:10.540Z] PORT=8080
[2026-03-03T00:26:10.550Z] NODE_ENV=production
[2026-03-03T00:26:10.560Z] Has REACT_APP_DIRECT_LINE_SECRET: true
[2026-03-03T00:26:10.580Z] Server is running on port 8080
[2026-03-03T00:26:10.590Z] Visit: http://localhost:8080
```

Then visit: `https://YOUR_APP_NAME.azurewebsites.net/` ✅

