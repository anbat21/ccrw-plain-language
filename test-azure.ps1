# Azure Deployment Diagnostic & Fix Script
# App URL: https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Azure App Service Diagnostics" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

$appName = "ccrw-plain-language-e4eaeyc9c6gdesah"
$appUrl = "https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net"

# Test 1: Can we reach the app?
Write-Host "[1/6] Testing connectivity..." -ForegroundColor Yellow
try {
    $response = Invoke-WebRequest -Uri "$appUrl/health" -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    Write-Host "  OK App is responding!" -ForegroundColor Green
    Write-Host "  Status: $($response.StatusCode)" -ForegroundColor Green
    Write-Host "  Response: $($response.Content)" -ForegroundColor Green
} catch {
    Write-Host "  ERROR App is NOT responding" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  This means the app crashed or was not deployed." -ForegroundColor Red
}
Write-Host ""

# Test 2: Check if files were deployed
Write-Host "[2/6] Checking Kudu (Advanced Tools)..." -ForegroundColor Yellow
Write-Host "  Open: https://$appName.scm.azurewebsites.net" -ForegroundColor Cyan
Write-Host "  Then: Debug Console > PowerShell" -ForegroundColor Cyan
Write-Host "  Navigate to: D:\home\site\wwwroot" -ForegroundColor Cyan
Write-Host "  Run: ls" -ForegroundColor Cyan
Write-Host "  You should see: server.js, package.json, dist/" -ForegroundColor Cyan
Write-Host ""

# Test 3: Check logs
Write-Host "[3/6] How to view logs:" -ForegroundColor Yellow
Write-Host "  Option A: Azure Portal > Your App > Log stream" -ForegroundColor Cyan
Write-Host "  Option B: Azure Portal > Your App > Deployment Center > Logs" -ForegroundColor Cyan
Write-Host ""

# Test 4: Configuration check
Write-Host "[4/6] Required Azure Portal Configuration:" -ForegroundColor Yellow
Write-Host "  Configuration > General Settings:" -ForegroundColor Cyan
Write-Host "    - Startup command: node server.js" -ForegroundColor White
Write-Host ""
Write-Host "  Configuration > Application settings (must have):" -ForegroundColor Cyan
Write-Host "    - REACT_APP_DIRECT_LINE_SECRET = (your bot secret)" -ForegroundColor White
Write-Host "    - SCM_DO_BUILD_DURING_DEPLOYMENT = true" -ForegroundColor White
Write-Host "    - WEBSITE_NODE_DEFAULT_VERSION = ~20" -ForegroundColor White
Write-Host ""

# Test 5: Common issues
Write-Host "[5/6] Common Deployment Issues:" -ForegroundColor Yellow
Write-Host "  Issue A: 'Application Error' = App crashed on startup" -ForegroundColor Red
Write-Host "    Fix: Check Log stream for error message" -ForegroundColor Green
Write-Host ""
Write-Host "  Issue B: 'Cannot find module' = npm install didn't run" -ForegroundColor Red
Write-Host "    Fix: Set SCM_DO_BUILD_DURING_DEPLOYMENT=true" -ForegroundColor Green
Write-Host ""
Write-Host "  Issue C: Package.json error" -ForegroundColor Red
Write-Host "    Fix: Remove prestart script (already done)" -ForegroundColor Green
Write-Host ""

# Test 6: Quick deploy check
Write-Host "[6/6] Quick Deploy Checklist:" -ForegroundColor Yellow
Write-Host "  [ ] Files committed to Git?" -ForegroundColor White
Write-Host "  [ ] Pushed to Azure remote?" -ForegroundColor White
Write-Host "  [ ] Startup command set?" -ForegroundColor White
Write-Host "  [ ] REACT_APP_DIRECT_LINE_SECRET set?" -ForegroundColor White
Write-Host "  [ ] Checked deployment logs?" -ForegroundColor White
Write-Host ""

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Next Steps" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "1. Open Kudu console to verify files exist" -ForegroundColor White
Write-Host "2. Check Log stream for real error message" -ForegroundColor White
Write-Host "3. Verify startup command and env vars" -ForegroundColor White
Write-Host "4. If nothing deployed, run: git push azure main" -ForegroundColor White
Write-Host ""
