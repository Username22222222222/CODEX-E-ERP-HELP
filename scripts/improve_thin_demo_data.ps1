param(
    [string]$Server = "MICRO\X",
    [string]$Database = "DEMO-DE",
    [string]$User = "sa",
    [string]$Password = $env:XERP_SQL_PASSWORD,
    [string]$CompletenessCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\program-audit\demo-data-field-completeness.csv",
    [string]$OutCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\program-audit\demo-data-field-improvement-log.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($Password)) {
    throw "XERP_SQL_PASSWORD is required."
}

Add-Type -AssemblyName System.Data

function New-Connection {
    $cs = "Server=$Server;Database=$Database;User Id=$User;Password=$Password;TrustServerCertificate=True;Encrypt=False"
    $conn = [System.Data.SqlClient.SqlConnection]::new($cs)
    $conn.Open()
    return $conn
}

function Invoke-Table {
    param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql)
    $cmd = $Connection.CreateCommand()
    $cmd.CommandTimeout = 120
    $cmd.CommandText = $Sql
    $dt = [System.Data.DataTable]::new()
    $reader = $cmd.ExecuteReader()
    $dt.Load($reader)
    return ,$dt
}

function Invoke-Scalar {
    param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql)
    $cmd = $Connection.CreateCommand()
    $cmd.CommandTimeout = 120
    $cmd.CommandText = $Sql
    return $cmd.ExecuteScalar()
}

function Quote-Name {
    param([string]$Name)
    return "[" + $Name.Replace("]", "]]") + "]"
}

function Short-Code {
    param([string]$Text, [int]$Max = 28)
    $s = ($Text -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($s.Length -gt $Max) { return $s.Substring(0, $Max) }
    return $s
}

function Get-Demo-Value {
    param([object]$Column, [string]$TableName, [int]$Ordinal)

    $name = [string]$Column.column_name
    $type = ([string]$Column.type_name).ToLowerInvariant()
    $maxLength = [int]$Column.max_length
    $base = "Hilfe-Demo " + (Short-Code $TableName 30)

    if ($type -eq "bit") { return $true }
    if ($type -in @("tinyint","smallint","int","bigint")) {
        if ($name -match "CreatedBy|ModifiedBy|Employee|Author|User") { return 1 }
        if ($name -match "Progress|Percent") { return 50 }
        if ($name -match "Minutes|Duration") { return 30 }
        return 1 + $Ordinal
    }
    if ($type -in @("decimal","numeric","money","smallmoney","float","real")) {
        if ($name -match "Price|Amount|Value|Cost|Rate|Total|Net|Gross|Debit|Credit") { return [decimal]123.45 }
        if ($name -match "Quantity|Weight|Hours|Duration") { return [decimal]2.5 }
        if ($name -match "Percent|Discount|Progress") { return [decimal]10.0 }
        return [decimal]1.0
    }
    if ($type -eq "date") { return [datetime]::Today.AddDays(14) }
    if ($type -in @("datetime","datetime2","smalldatetime")) { return [datetime]::Now.AddDays(1) }
    if ($type -eq "datetimeoffset") { return [datetimeoffset]::Now.AddDays(1) }
    if ($type -eq "uniqueidentifier") { return [guid]::NewGuid() }

    if ($type -in @("nvarchar","varchar","nchar","char","text","ntext","xml")) {
        $value = switch -Regex ($name) {
            "Password|Secret" { "DEMO-NICHT-ECHT"; break }
            "Token|ApiKey|AccessKey|Refresh" { "DEMO-TOKEN-NICHT-ECHT"; break }
            "Email|Mail" { "hilfe-demo@example.invalid"; break }
            "Phone|Mobile|Tel|Fax" { "+49 30 123456-0"; break }
            "Url|Uri|Endpoint|Host|Server" { "https://demo.x-erp.invalid"; break }
            "IBAN" { "DE02120300000000202051"; break }
            "BIC" { "BYLADEM1001"; break }
            "Vat|Tax" { "DE123456789"; break }
            "Zip|Postal" { "10115"; break }
            "City" { "Berlin"; break }
            "Street" { "Demostrasse 1"; break }
            "Country" { "Deutschland"; break }
            "Subject|Betreff|Title|Caption" { "$base Beispielvorgang"; break }
            "Name|Code|Number|No|Matchcode|Reference" { "$base"; break }
            "Xml" { "<Demo><Info>Hilfe-Demo fuer Screenshots</Info></Demo>"; break }
            "Json|Payload" { "{""demo"":true,""source"":""X-ERP Hilfe""}"; break }
            "Description|Info|Content|Text|Comment|Message|Note|Remark|Details|Body|Log" { "Kuenstlicher, fachlich lesbarer Demowert fuer X-ERP Hilfe-Screenshots in der DEMO-DE."; break }
            default { "$base " + $name }
        }

        if ($maxLength -gt 0) {
            $chars = if ($type -in @("nvarchar","nchar","ntext","xml")) { [math]::Floor($maxLength / 2) } else { $maxLength }
            if ($chars -gt 0 -and $value.Length -gt $chars) { $value = $value.Substring(0, $chars) }
        }
        return $value
    }

    return $null
}

$conn = New-Connection
try {
    $thinTables = Import-Csv -Path $CompletenessCsv -Delimiter ';' |
        Where-Object { $_.status -eq "THIN" } |
        Select-Object -ExpandProperty table -Unique |
        Sort-Object

    $targetSql = ($thinTables | ForEach-Object { $_.Replace("'", "''") }) -join "','"

    $columnSql = @"
SELECT
    s.name AS schema_name,
    t.name AS table_name,
    s.name + '.' + t.name AS full_name,
    c.column_id,
    c.name AS column_name,
    ty.name AS type_name,
    c.max_length,
    c.is_nullable,
    c.is_identity,
    c.is_computed,
    c.generated_always_type,
    CASE WHEN pk.column_id IS NULL THEN 0 ELSE 1 END AS is_pk
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
LEFT JOIN (
    SELECT ic.object_id, ic.column_id
    FROM sys.indexes i
    JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
    WHERE i.is_primary_key = 1
) pk ON pk.object_id = c.object_id AND pk.column_id = c.column_id
WHERE s.name + '.' + t.name IN ('$targetSql')
ORDER BY s.name, t.name, c.column_id;
"@
    $columns = Invoke-Table $conn $columnSql

    $fkSql = @"
SELECT
    sch1.name + '.' + tab1.name AS child_full_name,
    col1.name AS child_column,
    sch2.name AS parent_schema,
    tab2.name AS parent_table,
    sch2.name + '.' + tab2.name AS parent_full_name,
    col2.name AS parent_column
FROM sys.foreign_key_columns fkc
JOIN sys.tables tab1 ON tab1.object_id=fkc.parent_object_id
JOIN sys.schemas sch1 ON sch1.schema_id=tab1.schema_id
JOIN sys.columns col1 ON col1.object_id=tab1.object_id AND col1.column_id=fkc.parent_column_id
JOIN sys.tables tab2 ON tab2.object_id=fkc.referenced_object_id
JOIN sys.schemas sch2 ON sch2.schema_id=tab2.schema_id
JOIN sys.columns col2 ON col2.object_id=tab2.object_id AND col2.column_id=fkc.referenced_column_id
WHERE sch1.name + '.' + tab1.name IN ('$targetSql');
"@
    $fks = Invoke-Table $conn $fkSql

    $colsByTable = @{}
    foreach ($row in $columns.Rows) {
        $key = [string]$row.full_name
        if (-not $colsByTable.ContainsKey($key)) { $colsByTable[$key] = @() }
        $colsByTable[$key] += $row
    }

    $fkByTableColumn = @{}
    foreach ($fk in $fks.Rows) {
        $fkByTableColumn[[string]$fk.child_full_name + "|" + [string]$fk.child_column] = $fk
    }

    $log = [System.Collections.Generic.List[object]]::new()

    foreach ($table in $thinTables) {
        if (-not $colsByTable.ContainsKey($table)) { continue }
        $parts = $table.Split(".")
        $schema = $parts[0]
        $name = $parts[1]

        $pkCols = @($colsByTable[$table] | Where-Object { [bool]$_.is_pk })
        $where = $null
        $pkParams = @{}

        if ($pkCols.Count -gt 0) {
            $selectPkCols = ($pkCols | ForEach-Object { Quote-Name $_.column_name }) -join ", "
            $scoreParts = @()
            foreach ($c in $colsByTable[$table]) {
                $type = ([string]$c.type_name).ToLowerInvariant()
                if ([bool]$c.is_identity -or [bool]$c.is_computed -or ([int]$c.generated_always_type -ne 0) -or $type -in @("timestamp","rowversion","varbinary","binary","image")) { continue }
                $qcol = Quote-Name $c.column_name
                if ($type -in @("nvarchar","varchar","nchar","char","text","ntext","xml")) {
                    $scoreParts += "CASE WHEN $qcol IS NOT NULL AND NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(max), $qcol))), N'') IS NOT NULL THEN 1 ELSE 0 END"
                } else {
                    $scoreParts += "CASE WHEN $qcol IS NOT NULL THEN 1 ELSE 0 END"
                }
            }
            $scoreSql = if ($scoreParts.Count -gt 0) { $scoreParts -join " + " } else { "0" }
            $pkTable = Invoke-Table $conn ("SELECT TOP (1) $selectPkCols FROM " + (Quote-Name $schema) + "." + (Quote-Name $name) + " ORDER BY ($scoreSql) DESC")
            if ($pkTable.Rows.Count -eq 0) { continue }
            $whereParts = @()
            foreach ($pk in $pkCols) {
                $param = "@pk_" + $pk.column_name
                $whereParts += (Quote-Name $pk.column_name) + " = " + $param
                $pkParams[$param] = $pkTable.Rows[0].Item([string]$pk.column_name)
            }
            $where = $whereParts -join " AND "
        } else {
            $where = "1 = 1"
        }

        $ordinal = 0
        foreach ($col in $colsByTable[$table]) {
            $type = ([string]$col.type_name).ToLowerInvariant()
            $colName = [string]$col.column_name
            if ([bool]$col.is_identity -or [bool]$col.is_computed -or ([int]$col.generated_always_type -ne 0) -or $type -in @("timestamp","rowversion","varbinary","binary","image","geography","geometry","hierarchyid")) { continue }

            $value = $null
            $fkKey = $table + "|" + $colName
            if ($fkByTableColumn.ContainsKey($fkKey)) {
                $fk = $fkByTableColumn[$fkKey]
                $value = Invoke-Scalar $conn ("SELECT TOP (1) " + (Quote-Name $fk.parent_column) + " FROM " + (Quote-Name $fk.parent_schema) + "." + (Quote-Name $fk.parent_table) + " ORDER BY " + (Quote-Name $fk.parent_column))
                if ($null -eq $value -or $value -is [System.DBNull]) {
                    $log.Add([pscustomobject]@{ Table=$table; Column=$colName; Status="SKIPPED_NO_PARENT"; Message="Kein Parent in " + $fk.parent_full_name })
                    continue
                }
            } else {
                $value = Get-Demo-Value $col $table $ordinal
            }

            if ($null -eq $value) {
                $log.Add([pscustomobject]@{ Table=$table; Column=$colName; Status="SKIPPED_NO_VALUE"; Message="Kein sicherer Demowert ableitbar." })
                continue
            }

            $qcol = Quote-Name $colName
            $needsFill = if ($type -in @("nvarchar","varchar","nchar","char","text","ntext","xml")) {
                "($qcol IS NULL OR NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(max), $qcol))), N'') IS NULL)"
            } else {
                "$qcol IS NULL"
            }
            $topClause = if ($pkCols.Count -gt 0) { "" } else { "TOP (1) " }
            $cmd = $conn.CreateCommand()
            $cmd.CommandTimeout = 120
            $cmd.CommandText = "UPDATE $topClause" + (Quote-Name $schema) + "." + (Quote-Name $name) + " SET $qcol = @value WHERE $where AND $needsFill"
            [void]$cmd.Parameters.AddWithValue("@value", $value)
            foreach ($key in $pkParams.Keys) { [void]$cmd.Parameters.AddWithValue($key, $pkParams[$key]) }

            try {
                $affected = $cmd.ExecuteNonQuery()
                if ($affected -gt 0) {
                    $log.Add([pscustomobject]@{ Table=$table; Column=$colName; Status="UPDATED"; Message="Leeres Feld mit Demowert gefuellt." })
                }
            } catch {
                $log.Add([pscustomobject]@{ Table=$table; Column=$colName; Status="ERROR"; Message=$_.Exception.Message })
            }
            $ordinal++
        }
    }

    $log | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $log | Group-Object Status | Sort-Object Count -Descending | Select-Object Count,Name
} finally {
    $conn.Close()
}
