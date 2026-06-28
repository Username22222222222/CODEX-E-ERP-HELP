param(
    [Parameter(Mandatory = $true)]
    [string]$Password,
    [string]$BaseUrl = "https://MICRO/",
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ToolArgs
)

$ErrorActionPreference = 'Stop'
$env:XERP_SCREENSHOT_BASE_URL = $BaseUrl
$env:XERP_SCREENSHOT_PASSWORD = $Password
try {
    & dotnet "C:\Users\micha\Documents\X-ERP-HELP\tools\ScreenshotToolLocal\bin\Release\net10.0\ScreenshotTool.dll" @ToolArgs
}
finally {
    Remove-Item Env:\XERP_SCREENSHOT_PASSWORD -ErrorAction SilentlyContinue
    Remove-Item Env:\XERP_SCREENSHOT_BASE_URL -ErrorAction SilentlyContinue
}
