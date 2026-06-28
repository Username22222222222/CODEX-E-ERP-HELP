param(
    [string]$ScriptPath = "C:\X-ERP\X-IMPORT\Company-Script.sql",
    [string]$OutputJson = "C:\Users\micha\Documents\X-ERP-HELP\outputs\sql-audit\company-script-scriptdom-errors.json"
)

$ErrorActionPreference = 'Stop'

$dllCandidates = @(
    "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Extensions\Application\Microsoft.SqlServer.TransactSql.ScriptDom.dll",
    "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\Extensions\Microsoft\SQLCommon\Microsoft.SqlServer.TransactSql.ScriptDom.dll",
    "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\Extensions\Microsoft\SQLDB\DAC\Microsoft.SqlServer.TransactSql.ScriptDom.dll"
)
$dll = $dllCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $dll) { throw "ScriptDom DLL not found." }
if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "Script not found: $ScriptPath" }

Add-Type -Path $dll

$parser = [Microsoft.SqlServer.TransactSql.ScriptDom.TSql170Parser]::new($false)
$errors = $null
$reader = [System.IO.StreamReader]::new($ScriptPath, [System.Text.Encoding]::UTF8, $true)
try {
    $fragment = $parser.Parse($reader, [ref]$errors)
}
finally {
    $reader.Dispose()
}

$result = [pscustomobject]@{
    ScriptPath = $ScriptPath
    ParserDll = $dll
    ParseErrors = @($errors | ForEach-Object {
        [pscustomobject]@{
            Line = $_.Line
            Column = $_.Column
            Offset = $_.Offset
            Number = $_.Number
            Message = $_.Message
        }
    })
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputJson) | Out-Null
$result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
$result | ConvertTo-Json -Depth 4
