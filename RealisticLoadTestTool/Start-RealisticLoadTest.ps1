#Requires -Version 5.1
# ================================================================
#  Power BI Realistic Load Test - One-Click Start
#  For: Power BI China | Import mode
#  Usage: Double-click "qidong.cmd" (admin elevation handled by CMD)
# ================================================================

# Config
$INSTANCES = 8   # Number of concurrent Chrome windows

Set-ExecutionPolicy Bypass -Scope Process -Force

$workingDir = Split-Path -Parent $PSCommandPath
$testDir    = Join-Path $workingDir (Get-Date -Format "MM-dd-yyyy_HH_mm_ss")
$htmlFile   = "RealisticLoadTest.html"

function Write-Step { param($n,$msg) Write-Host "`n[$n/5] $msg" -ForegroundColor Yellow }
function Write-OK   { param($msg)    Write-Host "      $msg"    -ForegroundColor Green  }
function Write-Info { param($msg)    Write-Host "      $msg"    -ForegroundColor Gray   }

try {

    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host "  Power BI Realistic Load Test - China P1           " -ForegroundColor Cyan
    Write-Host "====================================================" -ForegroundColor Cyan

    # Step 1: Check and install PowerShell module
    Write-Step 1 "Checking MicrosoftPowerBIMgmt module..."
    if (-not (Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable)) {
        Write-Info "First run - installing module (1-3 minutes)..."
        Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop
    Write-OK "Module ready"

    # Step 2: Create test directory and copy JS libs
    Write-Step 2 "Preparing test directory and JS libs..."
    New-Item -Path $testDir -ItemType Directory -Force | Out-Null

    $jquerySrc = Join-Path $workingDir "jquery.min.js"
    $jqueryDst = Join-Path $testDir    "jquery.min.js"
    if (Test-Path $jquerySrc) {
        Copy-Item $jquerySrc $jqueryDst -ErrorAction Stop
    } else {
        Write-Info "Downloading jquery.min.js..."
        Invoke-WebRequest -Uri "https://cdn.bootcdn.net/ajax/libs/jquery/1.11.1/jquery.min.js" -OutFile $jqueryDst -ErrorAction Stop
    }

    $pbiJsSrc = Join-Path $workingDir "powerbi.min.js"
    $pbiJsDst = Join-Path $testDir    "powerbi.min.js"
    if (Test-Path $pbiJsSrc) {
        Copy-Item $pbiJsSrc $pbiJsDst -ErrorAction Stop
    } else {
        Write-Info "Downloading powerbi.min.js..."
        Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/powerbi-client@2.23.1/dist/powerbi.min.js" -OutFile $pbiJsDst -ErrorAction Stop
    }
    Write-OK "JS libs ready"

    # Step 3: Login to Power BI China
    Write-Step 3 "Login to Power BI China (browser auth window will open)..."
    $user     = Login-PowerBI -Environment China
    $rawToken = Get-PowerBIAccessToken -AsString | ForEach-Object { $_.replace("Bearer ", "").Trim() }
    Write-OK "Logged in: $($user.UserName)"

    # Step 4: Select workspace and report
    Write-Step 4 "Select workspace and report..."
    $workspaceList = Get-PowerBIWorkspace
    $idx = 1
    foreach ($ws in $workspaceList) {
        Write-Host ("  [{0,2}] {1}" -f $idx, $ws.Name) -ForegroundColor Green
        $idx++
    }
    $wsIdx = [int](Read-Host "`n      Enter workspace number") - 1

    $reportList = Get-PowerBIReport -WorkspaceId $workspaceList[$wsIdx].Id
    $idx = 1
    foreach ($r in $reportList) {
        Write-Host ("  [{0,2}] {1}" -f $idx, $r.Name) -ForegroundColor Green
        $idx++
    }
    $rIdx     = [int](Read-Host "`n      Enter report number") - 1
    $report   = $reportList[$rIdx]
    $embedUrl = $report.EmbedUrl
    Write-OK "Selected: $($report.Name)"

    # Step 5: Generate config files and launch Chrome windows
    Write-Step 5 "Generating config and launching $INSTANCES Chrome windows..."

    # PBIToken.JSON - write JS variable assignment that HTML loads as <script>
    $tokenContent = "accessToken='" + '{"PBIToken":"' + $rawToken + '"}' + "';"
    [System.IO.File]::WriteAllText(
        (Join-Path $testDir "PBIToken.JSON"),
        $tokenContent,
        [System.Text.Encoding]::UTF8
    )

    # PBIReport.JSON - read master template and inject reportUrl
    $templatePath = Join-Path $workingDir "PBIReport.json"
    if (Test-Path $templatePath) {
        $reportContent = (Get-Content $templatePath -Raw) -replace '"reportUrl":\s*""', ('"reportUrl": "' + $embedUrl + '"')
    } else {
        $reportContent = 'reportParameters= {"reportUrl": "' + $embedUrl + '", "sessionRestart": 100, "thinkTimeSeconds": 5};'
    }
    [System.IO.File]::WriteAllText(
        (Join-Path $testDir "PBIReport.JSON"),
        $reportContent,
        [System.Text.Encoding]::UTF8
    )

    # RealisticLoadTest.html - replace CDN URLs with local JS files
    $masterHtml  = Join-Path $workingDir $htmlFile
    $htmlContent = (Get-Content $masterHtml -Raw)
    $htmlContent = $htmlContent -replace 'https://ajax\.googleapis\.com/ajax/libs/jquery/[^"]+', 'jquery.min.js'
    $htmlContent = $htmlContent -replace 'https://cdn\.rawgit\.com/[^"]+powerbi\.min\.js', 'powerbi.min.js'
    [System.IO.File]::WriteAllText(
        (Join-Path $testDir $htmlFile),
        $htmlContent,
        [System.Text.Encoding]::UTF8
    )

    # Decode token expiry time for display
    try {
        $payload = $rawToken.Split('.')[1]
        $pad = 4 - ($payload.Length % 4)
        if ($pad -ne 4) { $payload += '=' * $pad }
        $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        $expTime = [DateTimeOffset]::FromUnixTimeSeconds($decoded.exp).LocalDateTime
    } catch {
        $expTime = (Get-Date).AddMinutes(60)
    }

    # Find Chrome via registry
    $chromePath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(default)'
    if ([string]::IsNullOrWhiteSpace($chromePath) -or !(Test-Path $chromePath)) {
        $chromePath = (Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(default)'
    }
    if ([string]::IsNullOrWhiteSpace($chromePath) -or !(Test-Path $chromePath)) {
        throw "Chrome not found. Please install Google Chrome and retry."
    }

    # Launch Chrome windows
    $htmlPath = Join-Path $testDir $htmlFile
    for ($i = 0; $i -lt $INSTANCES; $i++) {
        $profileDir = Join-Path $testDir "P$i"
        Start-Process -FilePath $chromePath -ArgumentList (
            "--allow-file-access-from-files",
            "--user-data-dir=$profileDir",
            "--disable-default-apps",
            "--new-window",
            $htmlPath
        )
        Start-Sleep -Seconds 2
    }

    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Green
    Write-Host "  Test started!" -ForegroundColor Green
    Write-Host "  Report:      $($report.Name)" -ForegroundColor White
    Write-Host "  Workspace:   $($workspaceList[$wsIdx].Name)" -ForegroundColor White
    Write-Host "  Concurrency: $INSTANCES Chrome windows" -ForegroundColor White
    Write-Host "  Token exp:   $($expTime.ToString('HH:mm:ss'))" -ForegroundColor Yellow
    Write-Host "  Monitor:     app.powerbi.cn -> Capacity Metrics" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Green
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Red
    Write-Host "  ERROR - check message below" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Red
    Write-Host ""
} finally {
    pause
}
