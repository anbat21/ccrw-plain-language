# IMMEDIATE FIX STEPS FOR AZURE DEPLOYMENT
## App: ccrw-plain-language-e4eaeyc9c6gdesah

Your app is showing "Application Error" and not responding. Follow these steps IN ORDER:

---

## STEP 1: Check What's Deployed (Kudu Console)

1. Open this URL in your browser:
   ```
   https://ccrw-plain-language-e4eaeyc9c6gdesah.scm.azurewebsites.net
   ```

2. Click **Debug console** → **PowerShell**

3. Navigate to: `D:\home\site\wwwroot`

4. Run command: `ls`

**What you should see:**
```
server.js
package.json
package-lock.json
dist/
web.config
```

**If the folder is EMPTY:**
- Nothing was deployed yet
- Skip to STEP 4 (Deploy section)

**If files exist:**
- Continue to STEP 2

---

## STEP 2: Check Why App Crashed (View Logs)

### Option A: Real-time Log Stream

1. Azure Portal → Your App Service: `ccrw-plain-language-e4eaeyc9c6gdesah`
2. Left menu → **Log stream**
3. Wait for logs to appear
4. Look for error messages (red text)

### Option B: Deployment Logs

1. Azure Portal → Your App Service
2. Left menu → **Deployment Center** → **Logs**
3. Click the latest deployment
4. Look for:
   - ✓ "Deployment successful" (good!)
   - ✗ "npm install failed" (bad!)
   - ✗ "Error: Cannot find module" (bad!)

**Common Errors:**

| Error Message | Cause | Fix |
|---------------|-------|-----|
| `Error: Cannot find module 'express'` | npm didn't run | Set `SCM_DO_BUILD_DURING_DEPLOYMENT=true` |
| `Missing REACT_APP_DIRECT_LINE_SECRET` | Env var not set | Add it in Application settings |
| `Error: listen EADDRINUSE` | Won't happen on Azure | Ignore |
| `Module build failed` | Webpack issue | Files not deployed properly |

---

## STEP 3: Fix Configuration (Azure Portal)

### A. Set Startup Command

1. Portal → Your App Service → **Configuration**
2. Tab: **General settings**
3. Find **Startup command**
4. Enter exactly: `node server.js`
5. Click **Save** (bottom of page)
6. Click **Continue** when prompted
7. Click **Save** (blue button at top)
8. Click **Restart** (at very top of page)

### B. Set Environment Variables

1. Same Configuration page
2. Tab: **Application settings**
3. Click **+ New application setting**
4. Add each of these:

**Setting 1:**
- Name: `REACT_APP_DIRECT_LINE_SECRET`
- Value: `YOUR_ACTUAL_BOT_SECRET_KEY_HERE`

**Setting 2:**
- Name: `SCM_DO_BUILD_DURING_DEPLOYMENT`
- Value: `true`

**Setting 3:**
- Name: `WEBSITE_NODE_DEFAULT_VERSION`
- Value: `~20`

**Setting 4:**
- Name: `NODE_ENV`
- Value: `production`

5. Click **Save** → **Continue** → Click **Restart**

---

## STEP 4: Deploy Code to Azure

### A. Local Build First

```powershell
cd C:\Users\batan\OneDrive\Desktop\CCRW_PlainLanguage\CCRW-Annotator

# Build the app
npm run build

# Verify dist folder exists
ls dist
```

### B. Initialize Git (if not done)

```powershell
# Check if git is initialized
git status

# If you get error "not a git repository", run:
git init
git add .
git commit -m "Initial commit for Azure deployment"
```

### C. Add Azure Remote

```powershell
# Get your deployment credentials from Portal > Deployment Center
# Then add the remote (replace USERNAME and PASSWORD):

git remote add azure https://USERNAME:PASSWORD@ccrw-plain-language-e4eaeyc9c6gdesah.scm.azurewebsites.net:443/ccrw-plain-language-e4eaeyc9c6gdesah.git

# Or if already added, update it:
git remote set-url azure https://USERNAME:PASSWORD@ccrw-plain-language-e4eaeyc9c6gdesah.scm.azurewebsites.net:443/ccrw-plain-language-e4eaeyc9c6gdesah.git
```

### D. Push to Azure

```powershell
# Make sure everything is committed
git add -A
git commit -m "Fix Azure deployment - ready for production"

# Push to Azure (this will deploy)
git push azure main

# If your branch is named master:
git push azure master
```

---

## STEP 5: Verify Deployment Worked

### A. Watch Deployment Progress

After `git push azure main`, you'll see output like:
```
remote: Updating branch 'main'.
remote: Updating submodules.
remote: Running post deployment command(s)...
remote: Running 'npm install --production'
remote: Deployment successful.
```

### B. Check Log Stream

1. Portal → Log stream
2. Wait 30 seconds
3. You should see:
```
[timestamp] Starting server...
[timestamp] PORT=8080
[timestamp] Has REACT_APP_DIRECT_LINE_SECRET: true
[timestamp] Server is running on port 8080
```

### C. Test the App

```powershell
# Test health endpoint
curl https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net/health

# Should return: {"status":"OK","timestamp":"..."}

# Test token endpoint
curl https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net/api/get-token

# Should return: {"token":"eyJ..."}

# Test main page
curl https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net/

# Should return: HTML page
```

---

## STEP 6: If Still Broken - Manual Debug

### In Kudu Console:

```powershell
# Navigate to app folder
cd D:\home\site\wwwroot

# Manually test startup
node server.js

# This will show the EXACT error
```

Common errors you might see:
- `Cannot find module 'express'` → npm install didn't run
- `Missing REACT_APP_DIRECT_LINE_SECRET` → env var not set
- `Cannot find module './dist'` → dist folder missing (run npm run build locally first)

---

## Quick Reference - Your App URLs

- **Main App:** https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net
- **Health Check:** https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net/health
- **Token API:** https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net/api/get-token
- **Kudu Console:** https://ccrw-plain-language-e4eaeyc9c6gdesah.scm.azurewebsites.net

---

## Diagnostic Commands

Run these to check status:

```powershell
# Run the diagnostic script
cd C:\Users\batan\OneDrive\Desktop\CCRW_PlainLanguage\CCRW-Annotator
.\test-azure.ps1
```

---

## Most Likely Issue Right Now

Based on "Application Error" message, the most likely problems are:

1. **Nothing deployed yet**
   - Fix: Run `git push azure main`

2. **npm install didn't run**
   - Fix: Set `SCM_DO_BUILD_DURING_DEPLOYMENT=true` in Portal
   - Then re-deploy

3. **Missing environment variable**
   - Fix: Add `REACT_APP_DIRECT_LINE_SECRET` in Portal
   - Then restart app

4. **Wrong startup command**
   - Fix: Set startup command to `node server.js` in Portal
   - Then restart app

---

## Success Checklist

- [ ] Files visible in Kudu console (D:\home\site\wwwroot)
- [ ] Startup command = `node server.js`
- [ ] REACT_APP_DIRECT_LINE_SECRET is set
- [ ] SCM_DO_BUILD_DURING_DEPLOYMENT = true
- [ ] Log stream shows "Server is running on port 8080"
- [ ] Health endpoint returns 200 OK
- [ ] Token endpoint returns JWT token
- [ ] Main page loads HTML

When ALL items are checked, your app will work! ✅
