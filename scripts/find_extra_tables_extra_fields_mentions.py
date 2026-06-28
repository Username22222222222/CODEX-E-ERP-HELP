from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


WORKBOOK = Path(r"C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx")
SHEET_XML = "xl/worksheets/sheet1.xml"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def split_ref(ref: str) -> tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", ref)
    if not match:
        raise ValueError(ref)
    return match.group(1), int(match.group(2))


def cell_text(cell: ET.Element) -> str:
    if cell.attrib.get("t") == "inlineStr":
        return "".join(cell.itertext()).strip()
    value = cell.find(NS + "v")
    return "" if value is None or value.text is None else value.text.strip()


def main() -> None:
    with zipfile.ZipFile(WORKBOOK, "r") as zf:
        root = ET.fromstring(zf.read(SHEET_XML))
    sheet_data = root.find(NS + "sheetData")
    if sheet_data is None:
        raise RuntimeError("sheetData not found")

    headers: dict[str, str] = {}
    hits: list[tuple[int, str, str, str]] = []
    pattern = re.compile(
        r"extra[-\s]?tabellen.{0,40}extra[-\s]?felder|extra[-\s]?felder.{0,40}extra[-\s]?tabellen",
        re.IGNORECASE | re.DOTALL,
    )
    for row in sheet_data.findall(NS + "row"):
        row_num = int(row.attrib["r"])
        values: dict[str, str] = {}
        for cell in row.findall(NS + "c"):
            col, _ = split_ref(cell.attrib["r"])
            values[col] = cell_text(cell)
        if row_num == 1:
            headers = {col: value for col, value in values.items()}
            continue
        topic = values.get("A", "")
        for col, value in values.items():
            if pattern.search(value):
                hits.append((row_num, headers.get(col, col), topic, value.replace("\n", "\\n")[:500]))

    print(f"hits={len(hits)}")
    for hit in hits[:200]:
        print("ROW", hit[0], "COL", hit[1], "TOPIC", hit[2])
        print(hit[3])


if __name__ == "__main__":
    main()
