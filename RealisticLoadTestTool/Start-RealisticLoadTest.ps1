#Requires -Version 5.1
# ================================================================
#  Power BI Realistic Load Test — 一键启动
#  适用于: Power BI China (世纪互联) | Import 模式
#  用法:   双击 "启动测试.cmd" 或以管理员身份运行此脚本
# ================================================================

# ── 配置区（修改这里）────────────────────────────────────────
$INSTANCES = 8        # 并发 Chrome 窗口数（建议 = CPU 核心数）
# ─────────────────────────────────────────────────────────────

# 自动提权到管理员
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Set-ExecutionPolicy Bypass -Scope Process -Force

$workingDir = Split-Path -Parent $PSCommandPath
$testDir    = Join-Path $workingDir (Get-Date -Format "MM-dd-yyyy_HH_mm_ss")
$htmlFile   = "RealisticLoadTest.html"

function Write-Step($n, $msg) { Write-Host "`n[$n/5] $msg" -ForegroundColor Yellow }
function Write-OK($msg)       { Write-Host "      $msg" -ForegroundColor Green }
function Write-Info($msg)     { Write-Host "      $msg" -ForegroundColor Gray }

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   Power BI Realistic Load Test — China P1        ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan

# ── 步骤 1: 检查并安装 PowerShell 模块 ───────────────────────
Write-Step 1 "检查 MicrosoftPowerBIMgmt 模块..."
if (-not (Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable)) {
    Write-Info "首次运行，正在安装模块（约 1-3 分钟）..."
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}
Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop
Write-OK "模块就绪 ✓"

# ── 步骤 2: 创建测试目录并准备 JS 库 ─────────────────────────
Write-Step 2 "准备测试目录和 JS 库..."
New-Item -Path $testDir -ItemType Directory -Force | Out-Null

$jquerySrc = Join-Path $workingDir "jquery.min.js"
$pbiJsSrc  = Join-Path $workingDir "powerbi.min.js"
$jqueryDst = Join-Path $testDir "jquery.min.js"
$pbiJsDst  = Join-Path $testDir "powerbi.min.js"

# 优先从仓库目录复制，失败则尝试网络下载
foreach ($pair in @(($jquerySrc,$jqueryDst,"https://cdn.bootcdn.net/ajax/libs/jquery/1.11.1/jquery.min.js"),
                    ($pbiJsSrc,$pbiJsDst,"https://cdn.jsdelivr.net/npm/powerbi-client@2.23.1/dist/powerbi.min.js"))) {
    $src, $dst, $url = $pair
    if (Test-Path $src) {
        Copy-Item $src $dst
    } else {
        Write-Info "本地未找到，尝试下载: $url"
        Invoke-WebRequest -Uri $url -OutFile $dst -ErrorAction Stop
    }
}
Write-OK "JS 库就绪 ✓"

# ── 步骤 3: 登录 Power BI China ──────────────────────────────
Write-Step 3 "登录 Power BI China（将弹出浏览器认证窗口）..."
$user     = Login-PowerBI -Environment China
$rawToken = Get-PowerBIAccessToken -AsString | ForEach-Object { $_.replace("Bearer ", "").Trim() }
Write-OK "登录成功: $($user.UserName) ✓"

# ── 步骤 4: 选择工作区和报表 ──────────────────────────────────
Write-Step 4 "选择工作区和报表..."
$workspaceList = Get-PowerBIWorkspace
$idx = 1
foreach ($ws in $workspaceList) {
    Write-Host ("  [{0,2}] {1}" -f $idx, $ws.Name) -ForegroundColor Green
    $idx++
}
$wsIdx = [int](Read-Host "`n      输入工作区序号") - 1

$reportList = Get-PowerBIReport -WorkspaceId $workspaceList[$wsIdx].Id
$idx = 1
foreach ($r in $reportList) {
    Write-Host ("  [{0,2}] {1}" -f $idx, $r.Name) -ForegroundColor Green
    $idx++
}
$rIdx     = [int](Read-Host "`n      输入报表序号") - 1
$report   = $reportList[$rIdx]
$embedUrl = $report.EmbedUrl
Write-OK "已选报表: $($report.Name) ✓"

# ── 步骤 5: 生成配置文件 + 修复 HTML + 启动 ──────────────────
Write-Step 5 "生成配置文件并启动 $INSTANCES 个测试窗口..."

# PBIToken.JSON
$tokenContent = "accessToken='{""PBIToken"":""$rawToken""}';"
[System.IO.File]::WriteAllText(
    (Join-Path $testDir "PBIToken.JSON"),
    $tokenContent,
    [System.Text.Encoding]::UTF8
)

# PBIReport.JSON — 读取主模板，注入真实 reportUrl
$templatePath = Join-Path $workingDir "PBIReport.json"
if (Test-Path $templatePath) {
    $reportContent = (Get-Content $templatePath -Raw) -replace '"reportUrl":\s*""', ('"reportUrl": "' + $embedUrl + '"')
} else {
    $reportContent = "reportParameters= {`"reportUrl`": `"$embedUrl`", `"sessionRestart`": 100, `"thinkTimeSeconds`": 5};"
}
[System.IO.File]::WriteAllText(
    (Join-Path $testDir "PBIReport.JSON"),
    $reportContent,
    [System.Text.Encoding]::UTF8
)

# RealisticLoadTest.html — 读取主模板，将 CDN 替换为本地 JS
$masterHtml = Join-Path $workingDir $htmlFile
$htmlContent = (Get-Content $masterHtml -Raw) `
    -replace 'https://ajax\.googleapis\.com/ajax/libs/jquery/[^"]+', 'jquery.min.js' `
    -replace 'https://cdn\.rawgit\.com/[^"]+powerbi\.min\.js',       'powerbi.min.js'
[System.IO.File]::WriteAllText(
    (Join-Path $testDir $htmlFile),
    $htmlContent,
    [System.Text.Encoding]::UTF8
)

# Token 过期时间
try {
    $payload = $rawToken.Split('.')[1]
    $pad = 4 - ($payload.Length % 4); if ($pad -ne 4) { $payload += '=' * $pad }
    $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    $expTime = [DateTimeOffset]::FromUnixTimeSeconds($decoded.exp).LocalDateTime
} catch { $expTime = (Get-Date).AddMinutes(60) }

# 查找 Chrome 路径
$chromePath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(default)'
if ([string]::IsNullOrWhiteSpace($chromePath) -or !(Test-Path $chromePath)) {
    $chromePath = (Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(default)'
}
if ([string]::IsNullOrWhiteSpace($chromePath) -or !(Test-Path $chromePath)) {
    throw "未找到 Chrome，请安装 Google Chrome 后重试。"
}

# 启动 Chrome 窗口
$htmlPath = Join-Path $testDir $htmlFile
for ($i = 0; $i -lt $INSTANCES; $i++) {
    Start-Process -FilePath $chromePath -ArgumentList @(
        "--allow-file-access-from-files",
        "--user-data-dir=`"$testDir\P$i`"",
        "--disable-default-apps",
        "--new-window",
        "`"$htmlPath`""
    )
    Start-Sleep -Seconds 2
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   测试已启动！                                    ║" -ForegroundColor Green
Write-Host "╠══════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host ("║   报表:    {0,-38}║" -f $report.Name)                             -ForegroundColor White
Write-Host ("║   工作区:  {0,-38}║" -f $workspaceList[$wsIdx].Name)              -ForegroundColor White
Write-Host ("║   并发数:  {0,-38}║" -f "$INSTANCES 个 Chrome 窗口")              -ForegroundColor White
Write-Host ("║   Token 过期: {0,-35}║" -f $expTime.ToString("HH:mm:ss"))         -ForegroundColor Yellow
Write-Host "╠══════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host "║  ⚠  请在 Token 过期前记录窗口计数和 Metrics      ║" -ForegroundColor Yellow
Write-Host "║  ⚠  监控: app.powerbi.cn → Capacity Metrics      ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
pause
