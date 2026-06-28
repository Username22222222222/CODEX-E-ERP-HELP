import csv
import json
import os
import re
import subprocess
from pathlib import Path
from urllib.parse import quote

ROOT = Path(r"C:\Users\micha\Documents\X-ERP-HELP")
AUDIT = ROOT / "outputs" / "program-audit" / "editview-demo-data-audit.csv"
FIELD = ROOT / "outputs" / "program-audit" / "demo-data-field-completeness.csv"
OUT = ROOT / "outputs" / "program-audit" / "editview-screenshot-queue.json"
OUT_CSV = ROOT / "outputs" / "program-audit" / "editview-screenshot-queue.csv"
SCREENSHOT_DIR = Path(r"D:\DATEN\HOMEPAGES\x-erp.de\de\help\Screenshots\AUTO-20260627")


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


def quote_name(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


def route_from_razor(path: str) -> tuple[str | None, str | None]:
    text = Path(path).read_text(encoding="utf-8-sig", errors="replace")
    routes = re.findall(r'@page\s+"([^"]+)"', text)
    route = None
    for candidate in routes:
        if "{Id" in candidate and "{Popup" not in candidate:
            route = candidate
            break
    if route is None:
        for candidate in routes:
            if "{Id" in candidate:
                route = candidate
                break
    default_register = None
    assignments = re.findall(r'(?:string|var)\s+\w*selectedRegisterId\s*=\s*([A-Za-z0-9_]+)\s*;', text, flags=re.I)
    if not assignments:
        assignments = re.findall(r'HighlightNodeOnLoad\s*=\s*@?([A-Za-z0-9_]+)', text)
    consts = dict(re.findall(r'const\s+string\s+([A-Za-z0-9_]+)\s*=\s*"([^"]+)"', text))
    if assignments:
        token = assignments[0]
        default_register = consts.get(token, token)
    if default_register is None:
        m = re.search(r'REGISTER_DEFAULT\s*=\s*([A-Za-z0-9_]+)', text)
        if m:
            default_register = consts.get(m.group(1))
    return route, default_register


def pk_for_table(table: str) -> list[str]:
    schema, name = table.split(".", 1)
    query = f"""
SET NOCOUNT ON;
SELECT c.name
FROM sys.indexes i
JOIN sys.index_columns ic ON ic.object_id=i.object_id AND ic.index_id=i.index_id
JOIN sys.columns c ON c.object_id=ic.object_id AND c.column_id=ic.column_id
WHERE i.is_primary_key=1 AND i.object_id=OBJECT_ID(N'{schema}.{name}')
ORDER BY ic.key_ordinal;
"""
    return [line.strip() for line in sqlcmd(query).splitlines() if line.strip()]


def best_pk_values(table: str, pk_cols: list[str]) -> dict[str, str] | None:
    if not pk_cols:
        return None
    schema, name = table.split(".", 1)
    cols_query = f"""
SET NOCOUNT ON;
SELECT c.name, ty.name
FROM sys.columns c
JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID(N'{schema}.{name}')
  AND c.is_identity=0
  AND c.is_computed=0
  AND c.generated_always_type=0
  AND ty.name NOT IN ('timestamp','rowversion','varbinary','binary','image','geography','geometry','hierarchyid')
ORDER BY c.column_id;
"""
    exprs = []
    for line in sqlcmd(cols_query).splitlines():
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        col, typ = parts[0], parts[1].lower()
        qcol = quote_name(col)
        if typ in {"nvarchar", "varchar", "nchar", "char", "text", "ntext", "xml"}:
            exprs.append(
                f"CASE WHEN {qcol} IS NOT NULL AND NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(max), {qcol}))), N'') IS NOT NULL THEN 1 ELSE 0 END"
            )
        else:
            exprs.append(f"CASE WHEN {qcol} IS NOT NULL THEN 1 ELSE 0 END")
    score = " + ".join(exprs) if exprs else "0"
    select_cols = ", ".join(f"CONVERT(nvarchar(4000), {quote_name(c)})" for c in pk_cols)
    query = f"""
SET NOCOUNT ON;
SELECT TOP (1) {select_cols}
FROM {quote_name(schema)}.{quote_name(name)}
ORDER BY ({score}) DESC;
"""
    lines = [line for line in sqlcmd(query).splitlines() if line.strip()]
    if not lines:
        return None
    values = lines[0].split("\t")
    return {col: values[i] for i, col in enumerate(pk_cols) if i < len(values)}


def build_url(route: str, pk_values: dict[str, str], default_register: str | None) -> str | None:
    value = pk_values.get("Id") or next(iter(pk_values.values()), None)
    if value is None:
        return None
    url = route
    url = re.sub(r"{Id(?::[^}]+)?}", quote(str(value), safe=""), url)
    url = re.sub(r"{SelectedRegisterId(?::[^}]+)?}", quote(default_register or "Overview", safe=""), url)
    url = re.sub(r"{Popup(?::[^}]+)?}", "false", url)
    # Skip routes that still require extra context, such as ArticleTemplateId.
    if "{" in url or "}" in url:
        return None
    return "https://micro" + url


def main() -> None:
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    field_status = {}
    with FIELD.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            field_status[row["table"]] = row["status"]
    records = []
    skips = []
    with AUDIT.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            if row.get("kind") != "View":
                continue
            if row.get("status") not in {"HAS_DATA", "LOW_DATA"}:
                skips.append({**row, "skip_reason": row.get("status")})
                continue
            table = row.get("table", "")
            if field_status.get(table) not in {"GOOD", "OK"}:
                skips.append({**row, "skip_reason": "FIELD_STATUS_" + field_status.get(table, "")})
                continue
            razor = row.get("razor_file", "")
            if not razor or not Path(razor).exists():
                skips.append({**row, "skip_reason": "NO_RAZOR_FILE"})
                continue
            try:
                route, default_register = route_from_razor(razor)
                if not route:
                    skips.append({**row, "skip_reason": "NO_ID_ROUTE"})
                    continue
                pk_cols = pk_for_table(table)
                pk_values = best_pk_values(table, pk_cols)
                if not pk_values:
                    skips.append({**row, "skip_reason": "NO_PK_VALUE"})
                    continue
                url = build_url(route, pk_values, default_register)
                if not url:
                    skips.append({**row, "skip_reason": "ROUTE_NEEDS_EXTRA_PARAMETERS"})
                    continue
                name = re.sub(r"[^A-Za-z0-9_.-]+", "_", row["view"]).strip("_")
                shot = SCREENSHOT_DIR / f"{name}.png"
                records.append(
                    {
                        "view": row["view"],
                        "breadcrumb": row["breadcrumb"],
                        "table": table,
                        "field_status": field_status.get(table),
                        "razor_file": razor,
                        "route": route,
                        "default_register": default_register or "",
                        "pk": pk_values,
                        "url": url,
                        "screenshot": str(shot),
                    }
                )
            except Exception as exc:
                skips.append({**row, "skip_reason": f"ERROR: {exc}"})

    OUT.write_text(json.dumps({"records": records, "skips": skips}, ensure_ascii=False, indent=2), encoding="utf-8")
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["view", "breadcrumb", "table", "field_status", "url", "screenshot", "default_register"],
            delimiter=";",
        )
        writer.writeheader()
        for item in records:
            writer.writerow({k: item.get(k, "") for k in writer.fieldnames})
    print("records", len(records))
    print("skips", len(skips))
    print(OUT)
    print(OUT_CSV)


if __name__ == "__main__":
    main()
