from __future__ import annotations

import copy
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET


WORKBOOK = Path(r"C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx")
ARCHIVE = Path(r"C:\Users\micha\Documents\X-ERP-HELP\ARCHIV")
SHEET_XML = "xl/worksheets/sheet1.xml"
NS_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = f"{{{NS_URI}}}"


ET.register_namespace("", NS_URI)


def split_ref(ref: str) -> tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", ref)
    if not match:
        raise ValueError(ref)
    return match.group(1), int(match.group(2))


def column_number(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def row_number_to_update(row: ET.Element) -> int:
    return int(row.attrib["r"])


def cell_col(cell: ET.Element) -> str:
    return split_ref(cell.attrib["r"])[0]


def cell_text(cell: ET.Element) -> str:
    if cell.attrib.get("t") == "inlineStr":
        return "".join(cell.itertext())
    value = cell.find(NS + "v")
    return "" if value is None or value.text is None else value.text


def set_cell_text(row: ET.Element, col: str, value: str) -> None:
    cells = {cell_col(cell): cell for cell in row.findall(NS + "c")}
    cell = cells.get(col)
    if cell is None:
        cell = ET.Element(NS + "c", {"r": f"{col}{row.attrib['r']}", "t": "inlineStr"})
        row.append(cell)
    cell.attrib["t"] = "inlineStr"
    for child in list(cell):
        cell.remove(child)
    inline = ET.SubElement(cell, NS + "is")
    text = ET.SubElement(inline, NS + "t")
    if value.startswith(" ") or value.endswith(" ") or "\n" in value:
        text.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
    text.text = value


def row_values(row: ET.Element) -> dict[str, str]:
    return {cell_col(cell): cell_text(cell) for cell in row.findall(NS + "c")}


def renumber_row(row: ET.Element, new_number: int) -> None:
    row.attrib["r"] = str(new_number)
    for cell in row.findall(NS + "c"):
        col = cell_col(cell)
        cell.attrib["r"] = f"{col}{new_number}"
    row[:] = sorted(row.findall(NS + "c"), key=lambda cell: column_number(cell_col(cell)))


def replace_in_existing_cells(row: ET.Element, old: str, new: str) -> None:
    for cell in row.findall(NS + "c"):
        value = cell_text(cell)
        if old in value:
            set_cell_text(row, cell_col(cell), value.replace(old, new))


def replace_text_in_existing_cells(row: ET.Element, replacements: dict[str, str]) -> None:
    for cell in row.findall(NS + "c"):
        value = cell_text(cell)
        changed = value
        for old, new in replacements.items():
            changed = changed.replace(old, new)
        if changed != value:
            set_cell_text(row, cell_col(cell), changed)


def set_extra_fields_root(row: ET.Element) -> None:
    values = {
        "A": "Extra-Felder",
        "B": "ExtraModules",
        "H": "Dieses Unterverzeichnis bündelt die EditViews und Wizards im Bereich Extra-Felder.",
        "J": "views-extra-felder",
        "K": "views/extra-felder",
        "L": "Extra-Felder Ansichten in X-ERP",
        "M": "Übersicht der X-ERP EditViews und Wizards im Bereich Extra-Felder.",
        "N": "Extra-Felder Ansichten",
        "R": "/de/help/views/extra-felder/",
        "S": "views-extra-felder",
        "V": "ansichten",
        "W": "extra-felder.md",
        "X": "/de/help/ansichten/extra-felder",
        "Y": "D:/DATEN/HOMEPAGES/x-erp.de/de/help/ansichten/extra-felder.md",
        "Z": "Extra-Felder",
        "AA": "Ansichten > Extra-Felder",
        "AB": "Ansichten",
        "AC": "1",
        "AD": "0001.0014",
        "AE": "ansichten/extra-felder",
    }
    for col, value in values.items():
        set_cell_text(row, col, value)
    for col in ("AF", "AG", "AH", "AI", "AK"):
        set_cell_text(row, col, "")


def set_extra_tables_root(row: ET.Element) -> None:
    values = {
        "A": "Extra-Tabellen",
        "B": "ExtraModules",
        "H": "Dieses Unterverzeichnis bündelt die EditViews und Wizards im Bereich Extra-Tabellen.",
        "J": "views-extra-tabellen",
        "K": "views/extra-tabellen",
        "L": "Extra-Tabellen Ansichten in X-ERP",
        "M": "Übersicht der X-ERP EditViews und Wizards im Bereich Extra-Tabellen.",
        "N": "Extra-Tabellen Ansichten",
        "R": "/de/help/views/extra-tabellen/",
        "S": "views-extra-tabellen",
        "V": "ansichten",
        "W": "extra-tabellen.md",
        "X": "/de/help/ansichten/extra-tabellen",
        "Y": "D:/DATEN/HOMEPAGES/x-erp.de/de/help/ansichten/extra-tabellen.md",
        "Z": "Extra-Tabellen",
        "AA": "Ansichten > Extra-Tabellen",
        "AB": "Ansichten",
        "AC": "1",
        "AD": "0001.0015",
        "AE": "ansichten/extra-tabellen",
    }
    for col, value in values.items():
        set_cell_text(row, col, value)
    for col in ("AF", "AG", "AH", "AI", "AK"):
        set_cell_text(row, col, "")


def adjust_extra_fields_row(row: ET.Element) -> None:
    replace_text_in_existing_cells(
        row,
        {
            "Extra-Tabellen und Extra-Felder": "Extra-Felder",
            "extra-tabellen-und-extra-felder": "extra-felder",
            "0001.0014.0004": "0001.0014.0001",
        },
    )
    values = row_values(row)
    if values.get("AB") == "Extra-Tabellen und Extra-Felder":
        set_cell_text(row, "AB", "Extra-Felder")


def remap_extra_tables_order(value: str) -> str:
    match = re.match(r"^0001\.0014\.(000[1-3]|000[5-8])(\..*)?$", value)
    if not match:
        return value.replace("0001.0014", "0001.0015", 1) if value.startswith("0001.0014.") else value
    old_child = int(match.group(1))
    new_child = old_child if old_child <= 3 else old_child - 1
    suffix = match.group(2) or ""
    return f"0001.0015.{new_child:04d}{suffix}"


def adjust_extra_tables_row(row: ET.Element) -> None:
    replace_text_in_existing_cells(
        row,
        {
            "Extra-Tabellen und Extra-Felder": "Extra-Tabellen",
            "extra-tabellen-und-extra-felder": "extra-tabellen",
        },
    )
    values = row_values(row)
    order = values.get("AD")
    if order:
        set_cell_text(row, "AD", remap_extra_tables_order(order))
    if values.get("AB") == "Extra-Tabellen und Extra-Felder":
        set_cell_text(row, "AB", "Extra-Tabellen")


def increment_later_ansichten_order(row: ET.Element) -> None:
    values = row_values(row)
    order = values.get("AD", "")
    match = re.match(r"^0001\.(\d{4})(.*)$", order)
    if not match:
        return
    top = int(match.group(1))
    if top < 15:
        return
    set_cell_text(row, "AD", f"0001.{top + 1:04d}{match.group(2)}")


def update_dimension(root: ET.Element, new_max_row: int) -> None:
    dimension = root.find(NS + "dimension")
    if dimension is None:
        return
    ref = dimension.attrib.get("ref", "")
    match = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    if match:
        dimension.attrib["ref"] = f"{match.group(1)}{match.group(2)}:{match.group(3)}{new_max_row}"


def main() -> None:
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = ARCHIVE / f"X-ERP-HELP-before-extra-split-{timestamp}.xlsx"
    shutil.copy2(WORKBOOK, backup)

    with zipfile.ZipFile(WORKBOOK, "r") as zf:
        files = {name: zf.read(name) for name in zf.namelist()}

    root = ET.fromstring(files[SHEET_XML])
    sheet_data = root.find(NS + "sheetData")
    if sheet_data is None:
        raise RuntimeError("sheetData not found")

    rows = list(sheet_data.findall(NS + "row"))
    row_by_number = {row_number_to_update(row): row for row in rows}
    for required in (2176, 2177, 2206, 2228, 2241, 2242):
        if required not in row_by_number:
            raise RuntimeError(f"Expected row {required} not found")

    prefix = [copy.deepcopy(row_by_number[row]) for row in sorted(row_by_number) if row < 2176]
    field_root = copy.deepcopy(row_by_number[2176])
    table_root = copy.deepcopy(row_by_number[2176])
    set_extra_fields_root(field_root)
    set_extra_tables_root(table_root)

    field_rows = [copy.deepcopy(row_by_number[row]) for row in range(2206, 2229)]
    table_rows = [copy.deepcopy(row_by_number[row]) for row in list(range(2177, 2206)) + list(range(2229, 2242))]
    suffix = [copy.deepcopy(row_by_number[row]) for row in sorted(row_by_number) if row >= 2242]

    for row in field_rows:
        adjust_extra_fields_row(row)
    for row in table_rows:
        adjust_extra_tables_row(row)
    for row in suffix:
        increment_later_ansichten_order(row)

    new_rows = prefix + [field_root] + field_rows + [table_root] + table_rows + suffix
    for index, row in enumerate(new_rows, start=1):
        renumber_row(row, index)

    for row in list(sheet_data):
        sheet_data.remove(row)
    for row in new_rows:
        sheet_data.append(row)

    update_dimension(root, len(new_rows))
    files[SHEET_XML] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in files.items():
                zf.writestr(name, data)
        shutil.move(str(tmp_path), WORKBOOK)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()

    print(f"backup={backup}")
    print("split=Extra-Felder, Extra-Tabellen")
    print(f"workbook={WORKBOOK}")


if __name__ == "__main__":
    main()
