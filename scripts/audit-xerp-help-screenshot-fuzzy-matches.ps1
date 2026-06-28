param(
    [string]$ScreenshotAuditCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-audit.csv",
    [string]$HelpRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-fuzzy-matches.csv"
)

$ErrorActionPreference = 'Stop'

function Normalize-Name([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    return ([regex]::Replace($value.ToLowerInvariant(), '[^a-z0-9]', ''))
}

if (-not (Test-Path -LiteralPath $ScreenshotAuditCsv)) { throw "CSV not found: $ScreenshotAuditCsv" }
if (-not (Test-Path -LiteralPath $HelpRoot)) { throw "Help root not found: $HelpRoot" }

$imageExtensions = @('.webp', '.png', '.jpg', '.jpeg')
$allImages = Get-ChildItem -LiteralPath $HelpRoot -Recurse -File |
    Where-Object { $imageExtensions -contains $_.Extension.ToLowerInvariant() } |
    ForEach-Object {
        [pscustomobject]@{
            FullName = $_.FullName
            Name = $_.Name
            Stem = [IO.Path]::GetFileNameWithoutExtension($_.Name)
            Normalized = Normalize-Name ([IO.Path]::GetFileNameWithoutExtension($_.Name))
            Length = $_.Length
        }
    }

$rows = Import-Csv -LiteralPath $ScreenshotAuditCsv -Delimiter ';' |
    Where-Object { $_.Status -eq 'fehlt' } |
    ForEach-Object {
        $targetStem = [IO.Path]::GetFileNameWithoutExtension($_.Screenshot)
        $targetNormalized = Normalize-Name $targetStem
        $matches = @(
            $allImages | Where-Object {
                $_.Normalized -eq $targetNormalized -or
                $_.Normalized.Contains($targetNormalized) -or
                $targetNormalized.Contains($_.Normalized)
            } | Sort-Object Length -Descending
        )
        $first = $matches | Select-Object -First 1
        [pscustomobject]@{
            Row = $_.Row
            Topic = $_.Topic
            MissingFile = [IO.Path]::GetFileName($_.Screenshot)
            MatchCount = $matches.Count
            BestMatchName = if ($first) { $first.Name } else { '' }
            BestMatchPath = if ($first) { $first.FullName } else { '' }
            BestMatchBytes = if ($first) { $first.Length } else { 0 }
        }
    }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
$rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    OutputCsv = $OutputCsv
    MissingScreenshots = @($rows).Count
    RowsWithFuzzyMatch = @($rows | Where-Object { [int]$_.MatchCount -gt 0 }).Count
    RowsWithoutFuzzyMatch = @($rows | Where-Object { [int]$_.MatchCount -eq 0 }).Count
    HelpImageFiles = @($allImages).Count
} | ConvertTo-Json
