param(
    [string]$ScriptPath = "C:\X-ERP\X-IMPORT\Company-Script.sql",
    [string]$OutputJson = "C:\Users\micha\Documents\X-ERP-HELP\outputs\sql-audit\company-script-audit.json"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Script not found: $ScriptPath"
}

$lines = Get-Content -LiteralPath $ScriptPath

function Add-Issue {
    param(
        [System.Collections.Generic.List[object]]$Issues,
        [string]$Severity,
        [string]$Code,
        [int]$Line,
        [string]$Message,
        [string]$Snippet = ''
    )
    $Issues.Add([pscustomobject]@{
        Severity = $Severity
        Code = $Code
        Line = $Line
        Message = $Message
        Snippet = $Snippet
    })
}

$issues = New-Object System.Collections.Generic.List[object]

$batches = New-Object System.Collections.Generic.List[object]
$start = 1
$buffer = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^\s*GO\s*(--.*)?$') {
        $batches.Add([pscustomobject]@{
            StartLine = $start
            EndLine = $i
            Text = ($buffer -join "`n")
        })
        $buffer.Clear()
        $start = $i + 2
    }
    else {
        $buffer.Add($lines[$i])
    }
}
if ($buffer.Count -gt 0) {
    $batches.Add([pscustomobject]@{
        StartLine = $start
        EndLine = $lines.Count
        Text = ($buffer -join "`n")
    })
}

$createRegex = [regex]'(?i)^\s*CREATE\s+(?:OR\s+ALTER\s+)?(?:(UNIQUE|CLUSTERED|NONCLUSTERED)\s+)?(TABLE|PROCEDURE|PROC|FUNCTION|VIEW|TRIGGER|TYPE|SCHEMA|INDEX)\s+((?:\[[^\]]+\]|\w+)(?:\s*\.\s*(?:\[[^\]]+\]|\w+))?)'
$objects = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $lines.Count; $i++) {
    $match = $createRegex.Match($lines[$i])
    if ($match.Success) {
        $kind = $match.Groups[2].Value.ToUpperInvariant()
        $name = ($match.Groups[3].Value -replace '\s+', '' -replace '^\[dbo\]\.', '' -replace '^dbo\.', '').Trim()
        $objects.Add([pscustomobject]@{
            Kind = $kind
            Name = $name
            Line = $i + 1
            Raw = $match.Value.Trim()
        })
    }
}

$duplicateCreates = $objects |
    Group-Object Kind, Name |
    Where-Object { $_.Count -gt 1 } |
    ForEach-Object {
        [pscustomobject]@{
            Object = $_.Name
            Count = $_.Count
            Lines = @($_.Group | Select-Object -ExpandProperty Line)
        }
    }

foreach ($dup in $duplicateCreates) {
    Add-Issue $issues 'High' 'DUPLICATE_CREATE' ($dup.Lines[0]) ("Object is created more than once: {0}; lines: {1}" -f $dup.Object, ($dup.Lines -join ', '))
}

foreach ($batch in $batches) {
    $batchText = [string]$batch.Text
    if ([string]::IsNullOrWhiteSpace($batchText)) { continue }

    $createMatches = $createRegex.Matches($batchText)
    if ($createMatches.Count -gt 1) {
        Add-Issue $issues 'Medium' 'MULTIPLE_CREATE_IN_BATCH' ([int]$batch.StartLine) ('Batch contains multiple CREATE statements. SQL Server requires CREATE PROC/FUNCTION/VIEW/TRIGGER to start a batch; verify this batch intentionally contains only compatible CREATE TABLE/INDEX statements.')
    }

    $routineMatches = [regex]::Matches($batchText, '(?im)^\s*CREATE\s+(?:OR\s+ALTER\s+)?(PROCEDURE|PROC|FUNCTION|VIEW|TRIGGER)\b')
    if ($routineMatches.Count -gt 0) {
        $prefix = $batchText.Substring(0, $routineMatches[0].Index)
        if ($prefix.Trim().Length -gt 0) {
            Add-Issue $issues 'High' 'ROUTINE_NOT_FIRST_IN_BATCH' ([int]$batch.StartLine) 'CREATE PROCEDURE/FUNCTION/VIEW/TRIGGER is not the first statement in its batch.'
        }
    }

    $beginTran = ([regex]::Matches($batchText, '(?i)\bBEGIN\s+TRAN(?:SACTION)?\b')).Count
    $commitTran = ([regex]::Matches($batchText, '(?i)\bCOMMIT(?:\s+TRAN(?:SACTION)?)?\b')).Count
    if ($beginTran -gt 0 -and $commitTran -eq 0 -and $batchText -notmatch '(?i)\bROLLBACK(?:\s+TRAN(?:SACTION)?)?\b') {
        Add-Issue $issues 'High' 'TRAN_WITHOUT_END' ([int]$batch.StartLine) 'Batch starts a transaction but no COMMIT or ROLLBACK was found in the same batch.'
    }
}

$linePatterns = @(
    @{ Severity = 'High'; Code = 'USE_SYSTEM_DATABASE'; Pattern = '^\s*USE\s+\[(master|model|msdb|tempdb)\]' },
    @{ Severity = 'High'; Code = 'DROP_DATABASE'; Pattern = '\bDROP\s+DATABASE\b' },
    @{ Severity = 'High'; Code = 'DROP_TABLE'; Pattern = '\bDROP\s+TABLE\b' },
    @{ Severity = 'Medium'; Code = 'SELECT_STAR'; Pattern = '^\s*SELECT\s+\*' },
    @{ Severity = 'Medium'; Code = 'NOLOCK'; Pattern = '\bNOLOCK\b' },
    @{ Severity = 'Medium'; Code = 'PRINT_DEBUG'; Pattern = '^\s*PRINT\s+' },
    @{ Severity = 'Low'; Code = 'TODO'; Pattern = 'TODO|FIXME' }
)

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    foreach ($p in $linePatterns) {
        if ($line -match $p.Pattern) {
            Add-Issue $issues $p.Severity $p.Code ($i + 1) ('Pattern found: ' + $p.Code) $line.Trim()
        }
    }
}

$summary = [pscustomobject]@{
    ScriptPath = $ScriptPath
    LineCount = $lines.Count
    SizeBytes = (Get-Item -LiteralPath $ScriptPath).Length
    BatchCount = $batches.Count
    ObjectCounts = @($objects | Group-Object Kind | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Kind = $_.Name; Count = $_.Count } })
    DuplicateCreates = @($duplicateCreates)
    IssueCounts = @($issues | Group-Object Severity, Code | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Issue = $_.Name; Count = $_.Count } })
    Issues = @($issues | Sort-Object @{Expression='Severity';Descending=$false}, Line)
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputJson) | Out-Null
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
$summary | ConvertTo-Json -Depth 5
