param(
    [string]$ScriptPath = "C:\X-ERP\X-IMPORT\Company-Script.sql",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\sql-audit\company-script-objects.csv"
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "Script not found: $ScriptPath" }

$lines = Get-Content -LiteralPath $ScriptPath
$objects = New-Object System.Collections.Generic.List[object]

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    if ($line -notmatch '^\s*CREATE\s+(?:OR\s+ALTER\s+)?(?:(UNIQUE|CLUSTERED|NONCLUSTERED)\s+)?(TABLE|PROCEDURE|PROC|FUNCTION|VIEW|TRIGGER|TYPE|SCHEMA|INDEX)\b(.*)$') {
        continue
    }

    $kind = $matches[2].ToUpperInvariant()
    $rest = $matches[3].Trim()
    $nameLine = $line
    $name = $null
    if ($rest -match '^((?:\[?[A-Za-z_#@][\w@#$]*\]?\.)?\[?[A-Za-z_#@][\w@#$]*\]?)(?:\s|\(|$)') {
        $name = $matches[1]
    }
    else {
        for ($j = $i + 1; $j -lt [Math]::Min($lines.Count, $i + 8); $j++) {
            $candidate = $lines[$j].Trim()
            if ($candidate.Length -eq 0 -or $candidate.StartsWith('--')) { continue }
            $nameLine = $lines[$j]
            if ($candidate -match '^((?:\[?[A-Za-z_#@][\w@#$]*\]?\.)?\[?[A-Za-z_#@][\w@#$]*\]?)(?:\s|\(|$)') {
                $name = $matches[1]
            }
            break
        }
    }

    if ($name) {
        $normalized = ($name -replace '^\[dbo\]\.', '' -replace '^dbo\.', '' -replace '\[|\]', '')
        if ($normalized.StartsWith('#') -or $normalized.StartsWith('@')) { continue }
        $objects.Add([pscustomobject]@{
            Line = $i + 1
            Kind = $kind
            Name = $normalized
            Source = $nameLine.Trim()
        })
    }
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
$objects | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

$duplicates = @($objects | Group-Object Kind, Name | Where-Object { $_.Count -gt 1 } | ForEach-Object {
    [pscustomobject]@{
        Object = $_.Name
        Count = $_.Count
        Lines = @($_.Group | Select-Object -ExpandProperty Line)
    }
})

[pscustomobject]@{
    Objects = $objects.Count
    Counts = @($objects | Group-Object Kind | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Kind = $_.Name; Count = $_.Count } })
    DuplicateCount = $duplicates.Count
    Duplicates = $duplicates
    OutputCsv = $OutputCsv
} | ConvertTo-Json -Depth 5
