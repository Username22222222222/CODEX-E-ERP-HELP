param(
    [Parameter(Mandatory = $true)]
    [string]$Password,
    [string]$BaseUrl = "https://MICRO/",
    [string]$Company = "DEMO-DE",
    [string]$Email = "manager@x-erp.de",
    [string]$ScreenshotAuditCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-audit.csv",
    [string]$ToolDll = "C:\Users\micha\Documents\X-ERP-HELP\tools\ScreenshotToolLocal\bin\Release\net10.0\ScreenshotTool.dll",
    [string]$LogPath = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\capture-missing-screenshots.log"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ScreenshotAuditCsv)) { throw "Screenshot audit CSV not found: $ScreenshotAuditCsv" }
if (-not (Test-Path -LiteralPath $ToolDll)) { throw "Tool DLL not found: $ToolDll" }

$topics = Import-Csv -LiteralPath $ScreenshotAuditCsv -Delimiter ';' |
    Where-Object { $_.Status -eq 'fehlt' } |
    ForEach-Object { $_.Topic } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LogPath) | Out-Null
"Missing screenshot topics: $($topics.Count)" | Set-Content -LiteralPath $LogPath -Encoding UTF8

if ($topics.Count -eq 0) {
    [pscustomobject]@{ Topics = 0; ModesRun = 0; LogPath = $LogPath } | ConvertTo-Json
    return
}

$env:XERP_SCREENSHOT_COMPANY = $Company
$env:XERP_SCREENSHOT_EMAIL = $Email
$env:XERP_SCREENSHOT_BASE_URL = $BaseUrl
$env:XERP_SCREENSHOT_PASSWORD = $Password

$modes = @(
    @(),
    @('--direct'),
    @('--popups'),
    @('--tabs'),
    @('--wizards')
)

$modesRun = 0
foreach ($mode in $modes) {
    $label = if ($mode.Count -eq 0) { 'list-routes' } else { $mode -join ' ' }
    "=== $label ===" | Add-Content -LiteralPath $LogPath -Encoding UTF8
    & dotnet $ToolDll @mode @topics 2>&1 | Tee-Object -FilePath $LogPath -Append
    $modesRun++
}

Remove-Item Env:\XERP_SCREENSHOT_PASSWORD -ErrorAction SilentlyContinue
Remove-Item Env:\XERP_SCREENSHOT_BASE_URL -ErrorAction SilentlyContinue

[pscustomobject]@{
    Topics = $topics.Count
    ModesRun = $modesRun
    LogPath = $LogPath
} | ConvertTo-Json
