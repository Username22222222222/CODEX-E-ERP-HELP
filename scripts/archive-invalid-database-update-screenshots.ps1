param(
    [string]$ScreenshotRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help\Screenshots",
    [string]$ArchiveDir = "C:\Users\micha\Documents\X-ERP-HELP\ARCHIV\bad-screenshots-database-update-20260625",
    [datetime]$CreatedAfter = "2026-06-25T03:40:00",
    [long]$InvalidFileSize = 16208
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ScreenshotRoot)) { throw "Screenshot root not found: $ScreenshotRoot" }
New-Item -ItemType Directory -Force -Path $ArchiveDir | Out-Null

$bad = Get-ChildItem -LiteralPath $ScreenshotRoot -File -Filter *.webp |
    Where-Object { $_.Length -eq $InvalidFileSize -and $_.LastWriteTime -ge $CreatedAfter }

foreach ($file in $bad) {
    Move-Item -LiteralPath $file.FullName -Destination (Join-Path $ArchiveDir $file.Name) -Force
}

[pscustomobject]@{
    Archived = $bad.Count
    Archive = $ArchiveDir
} | ConvertTo-Json
