param(
    [string]$ScreenshotSyncCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-sync-from-views.csv",
    [string]$ViewsRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help\views",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\missing-screenshot-html-audit.csv"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ScreenshotSyncCsv)) { throw "CSV not found: $ScreenshotSyncCsv" }
if (-not (Test-Path -LiteralPath $ViewsRoot)) { throw "Views root not found: $ViewsRoot" }

$htmlByName = @{}
Get-ChildItem -LiteralPath $ViewsRoot -Recurse -File -Filter *.html |
    ForEach-Object {
        $name = $_.Name.ToLowerInvariant()
        if (-not $htmlByName.ContainsKey($name)) {
            $htmlByName[$name] = New-Object System.Collections.Generic.List[object]
        }
        $htmlByName[$name].Add($_)
    }

$rows = Import-Csv -LiteralPath $ScreenshotSyncCsv -Delimiter ';' |
    Where-Object { $_.Action -eq 'not-found-in-views' } |
    ForEach-Object {
        $htmlName = [IO.Path]::GetFileNameWithoutExtension($_.FileName) + '.html'
        $key = $htmlName.ToLowerInvariant()
        $candidates = if ($htmlByName.ContainsKey($key)) { @($htmlByName.Item($key).ToArray()) } else { @() }
        $first = $candidates | Select-Object -First 1
        [pscustomobject]@{
            Row = $_.Row
            Topic = $_.Topic
            FileName = $_.FileName
            HtmlName = $htmlName
            HtmlFound = $candidates.Count -gt 0
            HtmlCandidateCount = $candidates.Count
            HtmlPath = if ($first) { $first.FullName } else { '' }
            HtmlLength = if ($first) { $first.Length } else { 0 }
        }
    }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
$rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    OutputCsv = $OutputCsv
    MissingScreenshots = @($rows).Count
    HtmlFound = @($rows | Where-Object { $_.HtmlFound }).Count
    HtmlMissing = @($rows | Where-Object { -not $_.HtmlFound }).Count
} | ConvertTo-Json
