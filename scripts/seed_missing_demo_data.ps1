param(
    [string]$Server = "MICRO\X",
    [string]$Database = "DEMO-DE",
    [string]$User = "sa",
    [string]$Password = $env:XERP_SQL_PASSWORD,
    [string]$AuditCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\program-audit\editview-demo-data-audit.csv",
    [string]$OutCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\program-audit\demo-data-seed-log.csv"
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
    param([string]$Text, [int]$Max = 24)
    $s = ($Text -replace '[^A-Za-z0-9]', '')
    if ($s.Length -gt $Max) { return $s.Substring(0, $Max) }
    return $s
}

function Get-Demo-Value {
    param(
        [object]$Column,
        [string]$TableName,
        [int]$Ordinal
    )

    $name = [string]$Column.column_name
    $type = ([string]$Column.type_name).ToLowerInvariant()
    $maxLength = [int]$Column.max_length
    $base = "Hilfe-Demo " + (Short-Code $TableName 32)

    if ($type -in @("bit")) { return $false }
    if ($type -in @("tinyint","smallint","int","bigint")) { return [int64](1 + $Ordinal) }
    if ($type -in @("decimal","numeric","money","smallmoney","float","real")) { return [decimal]1.0 }
    if ($type -in @("date")) { return [datetime]::Today }
    if ($type -in @("datetime","datetime2","smalldatetime","datetimeoffset")) { return [datetimeoffset]::UtcNow }
    if ($type -in @("uniqueidentifier")) { return [guid]::NewGuid() }

    if ($type -in @("nvarchar","varchar","nchar","char","text","ntext")) {
        $value = switch -Regex ($name) {
            "^(Id|.*Id)$" { "HELP-" + (Short-Code $TableName 30); break }
            "Email" { "hilfe-demo@example.invalid"; break }
            "Mail" { "hilfe-demo@example.invalid"; break }
            "Phone|Mobile|Tel" { "+49 30 123456-0"; break }
            "Url|Uri|Link" { "https://x-erp.de/de/help/"; break }
            "Subject|Betreff|Title|Caption" { "$base Beispiel"; break }
            "Name|Code|Number|No|Matchcode" { "$base"; break }
            "Description|Info|Content|Text|Comment|Message|Note|Body" { "Kuenstlicher Demodatensatz fuer X-ERP Hilfe-Screenshots in DEMO-DE."; break }
            "Street" { "Demostrasse 1"; break }
            "City" { "Berlin"; break }
            "Zip|Postal" { "10115"; break }
            "Country" { "Germany"; break }
            default { "$base " + $name }
        }

        if ($maxLength -gt 0) {
            $chars = if ($type -in @("nvarchar","nchar","ntext")) { [math]::Floor($maxLength / 2) } else { $maxLength }
            if ($chars -gt 0 -and $value.Length -gt $chars) { $value = $value.Substring(0, $chars) }
        }
        return $value
    }

    return $null
}

function Get-First-Parent-Value {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [object]$Fk
    )
    $sql = "SELECT TOP (1) " + (Quote-Name $Fk.parent_column) + " FROM " + (Quote-Name $Fk.parent_schema) + "." + (Quote-Name $Fk.parent_table) + " ORDER BY " + (Quote-Name $Fk.parent_column)
    return Invoke-Scalar $Connection $sql
}

$conn = New-Connection
try {
    $auditRows = Import-Csv -Path $AuditCsv -Delimiter ';' | Where-Object { $_.status -eq "MISSING_DATA" -and -not [string]::IsNullOrWhiteSpace($_.table) }
    $targetTables = $auditRows | Select-Object -ExpandProperty table -Unique | Sort-Object

    $metaSql = @"
SELECT
    s.name AS schema_name,
    t.name AS table_name,
    s.name + '.' + t.name AS full_name,
    c.column_id,
    c.name AS column_name,
    ty.name AS type_name,
    c.max_length,
    c.precision,
    c.scale,
    c.is_nullable,
    c.is_identity,
    c.is_computed,
    c.generated_always_type,
    CASE WHEN dc.object_id IS NULL THEN 0 ELSE 1 END AS has_default,
    CASE WHEN pk.column_id IS NULL THEN 0 ELSE 1 END AS is_pk
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
LEFT JOIN sys.default_constraints dc ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
LEFT JOIN (
    SELECT ic.object_id, ic.column_id
    FROM sys.indexes i
    JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
    WHERE i.is_primary_key = 1
) pk ON pk.object_id = c.object_id AND pk.column_id = c.column_id
WHERE s.name + '.' + t.name IN ('__TARGETS__')
ORDER BY s.name, t.name, c.column_id;
"@
    $targetSql = ($targetTables | ForEach-Object { $_.Replace("'", "''") }) -join "','"
    $columns = Invoke-Table $conn ($metaSql.Replace("__TARGETS__", $targetSql))

    $fkSql = @"
SELECT
    sch1.name + '.' + tab1.name AS child_full_name,
    sch1.name AS child_schema,
    tab1.name AS child_table,
    col1.name AS child_column,
    sch2.name AS parent_schema,
    tab2.name AS parent_table,
    sch2.name + '.' + tab2.name AS parent_full_name,
    col2.name AS parent_column
FROM sys.foreign_key_columns fkc
JOIN sys.foreign_keys fk ON fk.object_id = fkc.constraint_object_id
JOIN sys.tables tab1 ON tab1.object_id = fkc.parent_object_id
JOIN sys.schemas sch1 ON sch1.schema_id = tab1.schema_id
JOIN sys.columns col1 ON col1.object_id = tab1.object_id AND col1.column_id = fkc.parent_column_id
JOIN sys.tables tab2 ON tab2.object_id = fkc.referenced_object_id
JOIN sys.schemas sch2 ON sch2.schema_id = tab2.schema_id
JOIN sys.columns col2 ON col2.object_id = tab2.object_id AND col2.column_id = fkc.referenced_column_id
WHERE sch1.name + '.' + tab1.name IN ('__TARGETS__')
ORDER BY child_full_name, child_column;
"@
    $fks = Invoke-Table $conn ($fkSql.Replace("__TARGETS__", $targetSql))

    $colsByTable = @{}
    foreach ($row in $columns.Rows) {
        $key = [string]$row.full_name
        if (-not $colsByTable.ContainsKey($key)) { $colsByTable[$key] = @() }
        $colsByTable[$key] += $row
    }

    $fksByTableColumn = @{}
    foreach ($fk in $fks.Rows) {
        $fksByTableColumn[[string]$fk.child_full_name + "|" + [string]$fk.child_column] = $fk
    }

    $skipTables = @(
        "dbo.SystemDatabaseChangeLog",
        "dbo.WebshopActivityLog"
    )

    $log = [System.Collections.Generic.List[object]]::new()
    $pending = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $targetTables) { [void]$pending.Add($t) }

    $progress = $true
    $round = 0
    while ($pending.Count -gt 0 -and $progress -and $round -lt 20) {
        $round++
        $progress = $false
        foreach ($table in @($pending)) {
            if ($skipTables -contains $table) {
                $log.Add([pscustomobject]@{ Table=$table; Status="SKIPPED_TECHNICAL"; Message="Technische Logtabelle wird nicht kuenstlich befuellt."; InsertedId="" })
                [void]$pending.Remove($table)
                $progress = $true
                continue
            }

            if (-not $colsByTable.ContainsKey($table)) {
                $log.Add([pscustomobject]@{ Table=$table; Status="SKIPPED_NO_METADATA"; Message="Keine Spaltenmetadaten gefunden."; InsertedId="" })
                [void]$pending.Remove($table)
                $progress = $true
                continue
            }

            $parts = $table.Split(".")
            $schema = $parts[0]
            $name = $parts[1]
            $count = [int](Invoke-Scalar $conn ("SELECT COUNT_BIG(*) FROM " + (Quote-Name $schema) + "." + (Quote-Name $name)))
            if ($count -gt 0) {
                $log.Add([pscustomobject]@{ Table=$table; Status="SKIPPED_HAS_DATA"; Message="Tabelle hat bereits Daten."; InsertedId="" })
                [void]$pending.Remove($table)
                $progress = $true
                continue
            }

            $insertCols = [System.Collections.Generic.List[object]]::new()
            $values = @{}
            $blocked = $false
            $blockReason = ""
            $ordinal = 0

            foreach ($col in $colsByTable[$table]) {
                $type = ([string]$col.type_name).ToLowerInvariant()
                $colName = [string]$col.column_name
                if ([bool]$col.is_identity -or [bool]$col.is_computed -or ([int]$col.generated_always_type -ne 0) -or $type -in @("timestamp","rowversion")) { continue }

                $mustSet = (-not [bool]$col.is_nullable) -and (-not [bool]$col.has_default)
                $fkKey = $table + "|" + $colName
                if ($fksByTableColumn.ContainsKey($fkKey)) {
                    if (-not $mustSet -and -not [bool]$col.is_pk) { continue }
                    $parentValue = Get-First-Parent-Value $conn $fksByTableColumn[$fkKey]
                    if ($null -eq $parentValue -or $parentValue -is [System.DBNull]) {
                        $blocked = $true
                        $blockReason = "Kein Parent-Datensatz fuer " + $colName + " in " + $fksByTableColumn[$fkKey].parent_full_name
                        break
                    }
                    $insertCols.Add($col)
                    $values[$colName] = $parentValue
                    continue
                }

                if ($mustSet -or [bool]$col.is_pk) {
                    $val = Get-Demo-Value $col $table $ordinal
                    if ($null -eq $val) {
                        $blocked = $true
                        $blockReason = "Kein Wert fuer Pflichtspalte " + $colName + " (" + $type + ") ableitbar."
                        break
                    }
                    $insertCols.Add($col)
                    $values[$colName] = $val
                    $ordinal++
                }
            }

            if ($blocked) {
                $log.Add([pscustomobject]@{ Table=$table; Status="WAITING_OR_SKIPPED"; Message=$blockReason; InsertedId="" })
                continue
            }

            $insertColumnSql = ($insertCols | ForEach-Object { Quote-Name $_.column_name }) -join ", "
            $paramSql = ($insertCols | ForEach-Object { "@" + $_.column_name }) -join ", "
            if ([string]::IsNullOrWhiteSpace($insertColumnSql)) {
                $log.Add([pscustomobject]@{ Table=$table; Status="SKIPPED_NO_REQUIRED_COLUMNS"; Message="Keine Pflichtspalten zu setzen."; InsertedId="" })
                [void]$pending.Remove($table)
                $progress = $true
                continue
            }

            $cmd = $conn.CreateCommand()
            $cmd.CommandTimeout = 120
            $cmd.CommandText = "INSERT INTO " + (Quote-Name $schema) + "." + (Quote-Name $name) + " (" + $insertColumnSql + ") VALUES (" + $paramSql + "); SELECT CONVERT(nvarchar(200), COALESCE(SCOPE_IDENTITY(), 0));"
            foreach ($col in $insertCols) {
                $v = $values[[string]$col.column_name]
                $p = $cmd.Parameters.AddWithValue("@" + $col.column_name, $v)
                [void]$p
            }

            try {
                $insertedId = $cmd.ExecuteScalar()
                $log.Add([pscustomobject]@{ Table=$table; Status="INSERTED"; Message="Hilfe-Demodatensatz angelegt."; InsertedId=$insertedId })
                [void]$pending.Remove($table)
                $progress = $true
            } catch {
                $log.Add([pscustomobject]@{ Table=$table; Status="ERROR"; Message=$_.Exception.Message; InsertedId="" })
                [void]$pending.Remove($table)
                $progress = $true
            }
        }
    }

    foreach ($table in $pending) {
        $log.Add([pscustomobject]@{ Table=$table; Status="SKIPPED_BLOCKED"; Message="Nach mehreren Durchlaeufen nicht sicher befuellbar."; InsertedId="" })
    }

    $log | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $log | Group-Object Status | Sort-Object Count -Descending | Select-Object Count,Name
} finally {
    $conn.Close()
}
