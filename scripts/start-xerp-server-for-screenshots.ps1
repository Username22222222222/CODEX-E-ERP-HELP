param(
    [string]$Project = "C:\X-ERP\X-ERP\X-ERP.Server\XERP.Server.csproj",
    [string]$WorkingDirectory = "C:\X-ERP\X-ERP\X-ERP.Server",
    [string]$OutLog = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server.out.log",
    [string]$ErrLog = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server.err.log",
    [string]$PidFile = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\xerp-server.pid"
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutLog) | Out-Null
if (Test-Path -LiteralPath $OutLog) { Remove-Item -LiteralPath $OutLog -Force }
if (Test-Path -LiteralPath $ErrLog) { Remove-Item -LiteralPath $ErrLog -Force }

$process = Start-Process `
    -FilePath 'dotnet' `
    -ArgumentList @('run', '--project', $Project, '--launch-profile', 'https') `
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
