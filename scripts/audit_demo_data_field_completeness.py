import csv
import os
import subprocess
from pathlib import Path

ROOT = Path(r"C:\Users\micha\Documents\X-ERP-HELP")
AUDIT = ROOT / "outputs" / "program-audit" / "editview-demo-data-audit.csv"
OUT_CSV = ROOT / "outputs" / "program-audit" / "demo-data-field-completeness.csv"
OUT_MD = ROOT / "outputs" / "program-audit" / "demo-data-field-completeness-summary.md"


def sqlcmd(query: str) -> str:
    password = os.environ["XERP_SQL_PASSWORD"]
    cp = subprocess.run(
        [
            "sqlcmd",
            "-S",
            os.environ.get("XERP_SQL_SERVER", r"MICRO\X"),
            "-d",
            os.environ.get("XERP_SQL_DATABASE", "DEMO-DE"),
            "-U",
            os.environ.get("XERP_SQL_USER", "sa"),
            "-P",
            password,
            "-W",
            "-s",
            "\t",
            "-h",
            "-1",
            "-Q",
            query,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
    )
    if cp.returncode:
        raise RuntimeError(cp.stderr or cp.stdout)
    return cp.stdout


def main() -> None:
    with AUDIT.open("r", encoding="utf-8-sig", newline="") as f:
        rows = [
            r
            for r in csv.DictReader(f, delimiter=";")
            if r.get("table") and r.get("status") in {"HAS_DATA", "LOW_DATA"}
        ]
    tables = sorted({r["table"] for r in rows})
    targets = "','".join(t.replace("'", "''") for t in tables)
    meta_query = f"""
SET NOCOUNT ON;
SELECT s.name+'.'+t.name AS full_name, c.name, ty.name AS type_name, c.is_nullable, c.is_identity, c.is_computed, c.generated_always_type, c.max_length
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id=t.schema_id
JOIN sys.columns c ON c.object_id=t.object_id
JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE s.name+'.'+t.name IN ('{targets}')
ORDER BY full_name, c.column_id;
"""
    meta = []
    for line in sqlcmd(meta_query).splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) >= 8:
            meta.append(
                {
                    "table": parts[0],
                    "col": parts[1],
                    "type": parts[2].lower(),
                    "identity": parts[4] == "1",
                    "computed": parts[5] == "1",
                    "generated": parts[6] != "0",
                }
            )

    by_table = {}
    skip_types = {
        "timestamp",
        "rowversion",
        "varbinary",
        "binary",
        "image",
        "geography",
        "geometry",
        "hierarchyid",
    }
    skip_columns = {
        "DateCreated",
        "CreatedBy",
        "DateModified",
        "ModifiedBy",
    }
    for item in meta:
        if (
            item["identity"]
            or item["computed"]
            or item["generated"]
            or item["type"] in skip_types
            or item["col"] in skip_columns
        ):
            continue
        by_table.setdefault(item["table"], []).append(item)

    results = []
    for table, cols in by_table.items():
        if not cols:
            continue
        schema, name = table.split(".", 1)
        exprs = []
        for col in cols:
            qcol = "[" + col["col"].replace("]", "]]") + "]"
            if col["type"] in {"nvarchar", "varchar", "nchar", "char", "text", "ntext", "xml"}:
                exprs.append(
                    f"CASE WHEN {qcol} IS NOT NULL AND NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(max), {qcol}))), N'') IS NOT NULL THEN 1 ELSE 0 END"
                )
            else:
                exprs.append(f"CASE WHEN {qcol} IS NOT NULL THEN 1 ELSE 0 END")
        score = "+".join(exprs)
        query = f"""
SET NOCOUNT ON;
SELECT TOP (1) COUNT_BIG(*) OVER() AS row_count, ({score}) AS filled_count
FROM [{schema}].[{name}]
ORDER BY ({score}) DESC;
"""
        try:
            lines = sqlcmd(query).strip().splitlines()
            if lines:
                parts = lines[0].split("\t")
                row_count = int(parts[0])
                filled = int(parts[1])
            else:
                row_count = 0
                filled = 0
        except Exception:
            row_count = -1
            filled = -1
        total = len(cols)
        percent = round((filled / total * 100) if total and filled >= 0 else 0, 1)
        status = "GOOD" if percent >= 80 else ("OK" if percent >= 50 else "THIN")
        results.append(
            {
                "table": table,
                "row_count": row_count,
                "best_filled_fields": filled,
                "fillable_fields": total,
                "filled_percent": percent,
                "status": status,
            }
        )

    results.sort(key=lambda r: (r["status"], r["filled_percent"], r["table"]))
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "table",
                "row_count",
                "best_filled_fields",
                "fillable_fields",
                "filled_percent",
                "status",
            ],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(results)

    thin = [r for r in results if r["status"] == "THIN"]
    ok = [r for r in results if r["status"] == "OK"]
    good = [r for r in results if r["status"] == "GOOD"]
    lines = [
        "# Demodaten-Feldvollständigkeit",
        "",
        "Bewertung: Es wird je Tabelle der am besten gefüllte Datensatz betrachtet. Technische Spalten wie Identity, RowVersion, Computed, Binärfelder sowie reine Auditfelder werden ignoriert.",
        "",
        f"- GOOD: {len(good)}",
        f"- OK: {len(ok)}",
        f"- THIN: {len(thin)}",
        "",
        "## Dünne Tabellen",
        "",
    ]
    for item in sorted(thin, key=lambda r: (r["filled_percent"], r["table"])):
        lines.append(
            f"- {item['table']}: {item['best_filled_fields']}/{item['fillable_fields']} Felder ({item['filled_percent']}%), Zeilen: {item['row_count']}"
        )
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("tables", len(results))
    print("GOOD", len(good), "OK", len(ok), "THIN", len(thin))
    print(OUT_CSV)
    print(OUT_MD)


if __name__ == "__main__":
    main()
