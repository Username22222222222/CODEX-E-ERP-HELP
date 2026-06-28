param(
    [string]$DllPath = "C:\X-ERP\X-ERP\X-ERP.Server\bin\Debug\net10.0\XERP.Server.dll",
    [string]$WorkingDirectory = "C:\X-ERP\X-ERP\X-ERP.Server",
    [string]$SqlServerName = "MICRO\X",
    [string]$SqlUserName = "sa",
    [string]$SqlPasswordEncrypted = "1Z36P1Bu9+MMsMuXYuyECg==",
    [string]$OutLog = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server-dll.out.log",
    [string]$ErrLog = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server-dll.err.log",
    [string]$PidFile = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server-dll.pid"
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $DllPath)) { throw "DLL not found: $DllPath" }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutLog) | Out-Null
if (Test-Path -LiteralPath $OutLog) { Remove-Item -LiteralPath $OutLog -Force }
if (Test-Path -LiteralPath $ErrLog) { Remove-Item -LiteralPath $ErrLog -Force }

$env:ASPNETCORE_ENVIRONMENT = 'Development'
$env:ASPNETCORE_URLS = 'https://localhost:7177'
$env:ConnectionStrings__SQLServerName = $SqlServerName
$env:ConnectionStrings__SQLUserName = $SqlUserName
$env:ConnectionStrings__SQLPassword = $SqlPasswordEncrypted

$process = Start-Process `
    -FilePath 'dotnet' `
    -ArgumentList @($DllPath) `
    -WorkingDirectory $WorkingDirectory `
    -WindowStyle Hidden `
    -PassThru `
    -RedirectStandardOutput $OutLog `
    -RedirectStandardError $ErrLog

Set-Content -LiteralPath $PidFile -Value $process.Id -Encoding ASCII

[pscustomobject]@{
    Pid = $process.Id
    OutLog = $OutLog
    ErrLog = $ErrLog
    PidFile = $PidFile
} | ConvertTo-Json
