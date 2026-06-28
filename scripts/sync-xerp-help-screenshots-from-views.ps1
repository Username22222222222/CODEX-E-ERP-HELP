param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$ViewsRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help\views",
    [string]$ScreenshotRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help\Screenshots",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-sync-from-views.csv",
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path -LiteralPath $ViewsRoot)) { throw "Views root not found: $ViewsRoot" }

function Get-FileNameFromScreenshot([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    $v = $value.Trim().Replace('\', '/')
    return [IO.Path]::GetFileName($v)
}

$imageExtensions = @('.webp', '.png', '.jpg', '.jpeg')
$imagesByName = @{}
Get-ChildItem -LiteralPath $ViewsRoot -Recurse -File |
    Where-Object { $imageExtensions -contains $_.Extension.ToLowerInvariant() } |
    ForEach-Object {
        $name = $_.Name.ToLowerInvariant()
        if (-not $imagesByName.ContainsKey($name)) {
            $imagesByName[$name] = New-Object System.Collections.Generic.List[object]
        }
        $imagesByName[$name].Add($_)
    }

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $headers = @{}
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        $h = [string]$ws.Cells.Item(1, $c).Text
        if (-not [string]::IsNullOrWhiteSpace($h)) { $headers[$h] = $c }
    }

    foreach ($required in @('Screenshot', 'SCREENSHOT_REL_PATH', 'SCREENSHOT_WEB_PATH', 'IMAGE_STATUS')) {
        if (-not $headers.ContainsKey($required)) { throw "Column missing: $required" }
    }

    $results = New-Object System.Collections.Generic.List[object]
    $copied = 0
    $matched = 0
    $missing = 0
    $ambiguous = 0
    $alreadyExisting = 0

if ($Apply) {
    New-Item -ItemType Directory -Force -Path $ScreenshotRoot | Out-Null
    $backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = Join-Path $backupDir "X-ERP-HELP.before-screenshot-sync-from-views-$stamp.xlsx"
    Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force
}

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $status = [string]$ws.Cells.Item($r, $headers['IMAGE_STATUS']).Text
        if ($status -ne 'fehlt') { continue }

        $screenshot = [string]$ws.Cells.Item($r, $headers['Screenshot']).Text
        if ([string]::IsNullOrWhiteSpace($screenshot)) {
            $screenshot = [string]$ws.Cells.Item($r, $headers['SCREENSHOT_REL_PATH']).Text
        }
        $fileName = Get-FileNameFromScreenshot $screenshot
        if ([string]::IsNullOrWhiteSpace($fileName)) { continue }

        $targetPath = Join-Path $ScreenshotRoot $fileName
        if (Test-Path -LiteralPath $targetPath) {
            $alreadyExisting++
            if ($Apply) { $ws.Cells.Item($r, $headers['IMAGE_STATUS']).Value = 'vorhanden' }
            $results.Add([pscustomobject]@{
                Row = $r
                Topic = [string]$ws.Cells.Item($r, 1).Text
                FileName = $fileName
                Action = 'already-existing'
                Source = ''
                Target = $targetPath
                CandidateCount = 0
            })
            continue
        }

        $key = $fileName.ToLowerInvariant()
        if (-not $imagesByName.ContainsKey($key)) {
            $missing++
            $results.Add([pscustomobject]@{
                Row = $r
                Topic = [string]$ws.Cells.Item($r, 1).Text
                FileName = $fileName
                Action = 'not-found-in-views'
                Source = ''
                Target = $targetPath
                CandidateCount = 0
            })
            continue
        }

        $candidateList = $imagesByName.Item($key)
        $candidates = @($candidateList.ToArray())
        $source = $candidates | Sort-Object Length -Descending | Select-Object -First 1
        $matched++
        if ($candidates.Count -gt 1) { $ambiguous++ }

        if ($Apply) {
            Copy-Item -LiteralPath $source.FullName -Destination $targetPath -Force
            $ws.Cells.Item($r, $headers['IMAGE_STATUS']).Value = 'vorhanden'
            $copied++
        }

        $results.Add([pscustomobject]@{
            Row = $r
            Topic = [string]$ws.Cells.Item($r, 1).Text
            FileName = $fileName
            Action = if ($Apply) { 'copied' } else { 'would-copy' }
            Source = $source.FullName
            Target = $targetPath
            CandidateCount = $candidates.Count
        })
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $results | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

    if ($Apply) {
        $pkg.Save()
    }

    [pscustomobject]@{
        Apply = [bool]$Apply
        OutputCsv = $OutputCsv
        MissingRowsProcessed = $results.Count
        MatchedInViews = $matched
        Copied = $copied
        AlreadyExisting = $alreadyExisting
        NotFoundInViews = $missing
        DuplicateNameMatches = $ambiguous
        ViewImageFiles = ($imagesByName.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    } | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
