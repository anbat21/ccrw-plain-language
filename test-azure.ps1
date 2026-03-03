param(
    [string]$AppName = "ccrw-plain-language-e4eaeyc9c6gdesah",
    [string]$AppUrl = "https://ccrw-plain-language-e4eaeyc9c6gdesah.canadacentral-01.azurewebsites.net",
    [string]$LocalUrl = "http://localhost:3000"
)

$scmHost = "$AppName.scm.azurewebsites.net"
$scmUrl = "https://$scmHost"

try {
    $appUri = [Uri]$AppUrl
    $hostParts = $appUri.Host.Split('.')
    if ($hostParts.Length -ge 2) {
        $scmHost = "$($hostParts[0]).scm.$($hostParts[1..($hostParts.Length - 1)] -join '.')"
        $scmUrl = "https://$scmHost"
    }
}
catch {
}

$results = [System.Collections.Generic.List[object]]::new()

function Add-Result {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )
    $results.Add([PSCustomObject]@{
        Check  = $Name
        Result = if ($Passed) { "PASS" } else { "FAIL" }
        Detail = $Detail
    })
}

function Test-Http {
    param(
        [string]$Name,
        [string]$Url,
        [int]$TimeoutSec = 10,
        [scriptblock]$Validator,
        [switch]$AllowAuthChallenge
    )

    try {
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop
        if ($Validator) {
            $isOk = & $Validator $resp
            if ($isOk) {
                Add-Result -Name $Name -Passed $true -Detail "HTTP $($resp.StatusCode)"
            } else {
                Add-Result -Name $Name -Passed $false -Detail "HTTP $($resp.StatusCode), unexpected content"
            }
        } else {
            Add-Result -Name $Name -Passed $true -Detail "HTTP $($resp.StatusCode)"
        }
    }
    catch {
        if ($AllowAuthChallenge -and $_.Exception.Response -and ($_.Exception.Response.StatusCode.value__ -in 401, 403)) {
            Add-Result -Name $Name -Passed $true -Detail "HTTP $($_.Exception.Response.StatusCode.value__) (auth challenge)"
        } else {
            Add-Result -Name $Name -Passed $false -Detail $_.Exception.Message
        }
    }
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Local + Azure Verifier" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Local: $LocalUrl"
Write-Host "Azure: $AppUrl"
Write-Host "SCM:   $scmUrl"
Write-Host ""

Test-Http -Name "Local /health" -Url "$LocalUrl/health" -TimeoutSec 5 -Validator {
    param($resp)
    $resp.Content -match '"status"\s*:\s*"OK"'
}

Test-Http -Name "Local /api/get-token" -Url "$LocalUrl/api/get-token" -TimeoutSec 8 -Validator {
    param($resp)
    $json = $resp.Content | ConvertFrom-Json
    -not [string]::IsNullOrWhiteSpace($json.token)
}

Test-Http -Name "Local /taskpane.html" -Url "$LocalUrl/taskpane.html" -TimeoutSec 5 -Validator {
    param($resp)
    $resp.Content -match '<html|<title|taskpane'
}

try {
    $dns = Resolve-DnsName -Name $scmHost -Type A -ErrorAction Stop | Select-Object -First 1
    Add-Result -Name "SCM DNS resolve" -Passed $true -Detail "Resolved to $($dns.IPAddress)"
}
catch {
    Add-Result -Name "SCM DNS resolve" -Passed $false -Detail $_.Exception.Message
}

Test-Http -Name "Azure /health" -Url "$AppUrl/health" -TimeoutSec 10 -Validator {
    param($resp)
    $resp.Content -match '"status"\s*:\s*"OK"'
}

Test-Http -Name "Azure /" -Url "$AppUrl/" -TimeoutSec 10

Test-Http -Name "SCM root" -Url $scmUrl -TimeoutSec 10 -AllowAuthChallenge

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Results" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

$results | Format-Table -AutoSize

$failed = $results | Where-Object { $_.Result -eq "FAIL" }

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "All checks passed." -ForegroundColor Green
    exit 0
}

Write-Host "One or more checks failed." -ForegroundColor Yellow

if ($failed.Check -contains "Azure /health" -or $failed.Check -contains "Azure /") {
    Write-Host "- Azure app is not serving expected responses. Check App Service status + Log stream." -ForegroundColor Yellow
}

if ($failed.Check -contains "SCM root") {
    Write-Host "- SCM endpoint unreachable. Verify App Service exists/running and network access." -ForegroundColor Yellow
}

if ($failed.Check -contains "Local /health" -or $failed.Check -contains "Local /api/get-token") {
    Write-Host "- Local server may be down. Start it with: npm start" -ForegroundColor Yellow
}

exit 1
