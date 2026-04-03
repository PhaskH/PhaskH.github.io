from __future__ import annotations

import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

SOURCE = Path("/Users/robert/Desktop/MetaGenMHN/google_source - aggressive excel repair v5.xlsx")
OUTPUT = Path("/Users/robert/Desktop/MetaGenMHN/data/workbook-data.json")
OUTPUT_JS = Path("/Users/robert/Desktop/MetaGenMHN/data/workbook-data.js")
SHEETS = ["Calculator1", "Weapons", "Types", "Riftborne", "Status", "Skills"]


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(num: int) -> str:
    out = []
    while num:
        num, rem = divmod(num - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def parse_ref(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)", ref)
    if not match:
        raise ValueError(f"invalid cell reference: {ref}")
    return col_to_num(match.group(1)), int(match.group(2))


def normalize_scalar(text: str | None, cell_type: str | None, shared_strings: list[str]):
    if text is None:
        return None
    if cell_type == "s":
        return shared_strings[int(text)]
    if cell_type == "b":
        return bool(int(text))
    if cell_type == "str":
        return text
    try:
        value = float(text)
    except ValueError:
        return text
    if value.is_integer():
        return int(value)
    return value


def sanitize_formula(sheet: str, ref: str, formula: str) -> str:
    if sheet == "Calculator1" and ref == "AX86":
        formula = 'SUMIF($AX$68:$AY$74,0.75,$AX$77:$AY$83)'
    elif sheet == "Calculator1" and ref == "AX88":
        formula = 'SUMIF($AX$68:$AY$74,1.25,$AX$77:$AY$83)'
    if sheet == "Calculator1" and ref == "S67":
        formula = (
            'IF($B$70="Raw",INDEX(Riftborne!$B$4:$B$24,MATCH($E$18,Riftborne!$A$4:$A$24,0)),'
            'IF($B$70="Element",INDEX(Riftborne!$C$4:$C$24,MATCH($E$18,Riftborne!$A$4:$A$24,0)),'
            'INDEX(Riftborne!$E$4:$E$24,MATCH($E$18,Riftborne!$A$4:$A$24,0))))'
        )
    elif sheet == "Calculator1" and ref == "S68":
        formula = (
            'COUNTIF($E$19:$E$21,"Attack")*'
            'IF($B$70="Raw",INDEX(Riftborne!$I$3:$I$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'IF($B$70="Element",INDEX(Riftborne!$J$3:$J$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'INDEX(Riftborne!$K$3:$K$6,MATCH("Attack",Riftborne!$H$3:$H$6,0))))'
        )

    formula = re.sub(r"\bTrue\b", "TRUE()", formula)
    formula = re.sub(r"\bFalse\b", "FALSE()", formula)
    return f"={formula}"


def parse_sqref(sqref: str) -> list[str]:
    refs: list[str] = []
    for token in sqref.split():
        if ":" not in token:
            refs.append(token)
            continue
        start, end = token.split(":")
        start_col, start_row = parse_ref(start)
        end_col, end_row = parse_ref(end)
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                refs.append(f"{num_to_col(col)}{row}")
    return refs


def iter_range(range_ref: str) -> list[str]:
    start, end = range_ref.split(":")
    start_col, start_row = parse_ref(start.replace("$", ""))
    end_col, end_row = parse_ref(end.replace("$", ""))
    refs: list[str] = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            refs.append(f"{num_to_col(col)}{row}")
    return refs


def resolve_validation_options(
    formula: str,
    workbook_values: dict[str, dict[str, object]],
) -> list[object]:
    if formula.startswith('"') and formula.endswith('"'):
        return [item.strip() for item in formula[1:-1].split(",")]

    if "!" in formula:
        sheet_name, range_ref = formula.split("!", 1)
    else:
        sheet_name, range_ref = "Calculator1", formula

    sheet_name = sheet_name.strip("'")
    refs = iter_range(range_ref)
    values = workbook_values[sheet_name]
    return [values[ref] for ref in refs if ref in values and values[ref] not in (None, "")]


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(SOURCE) as archive:
        shared_strings_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        shared_strings = [
            "".join(t.text or "" for t in si.iterfind(".//a:t", NS))
            for si in shared_strings_root.findall("a:si", NS)
        ]

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}
        sheet_files = {
            sheet.attrib["name"]: f"xl/{rel_map[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]}"
            for sheet in workbook_root.find("a:sheets", NS)
        }

        sheet_matrices: dict[str, list[list[object | None]]] = {}
        workbook_values: dict[str, dict[str, object]] = {}

        for sheet_name in SHEETS:
            root = ET.fromstring(archive.read(sheet_files[sheet_name]))
            cells = root.findall(".//a:c", NS)
            max_row = 0
            max_col = 0
            workbook_values[sheet_name] = {}
            parsed_cells: list[tuple[int, int, object]] = []

            for cell in cells:
                ref = cell.attrib["r"]
                col, row = parse_ref(ref)
                max_col = max(max_col, col)
                max_row = max(max_row, row)

                formula = cell.find("a:f", NS)
                value = cell.find("a:v", NS)
                if formula is not None and formula.text:
                    content: object = sanitize_formula(sheet_name, ref, formula.text)
                else:
                    content = normalize_scalar(
                        None if value is None else value.text,
                        cell.attrib.get("t"),
                        shared_strings,
                    )

                if value is not None:
                    workbook_values[sheet_name][ref] = normalize_scalar(
                        value.text,
                        cell.attrib.get("t"),
                        shared_strings,
                    )
                elif formula is None:
                    workbook_values[sheet_name][ref] = None

                parsed_cells.append((row, col, content))

            matrix: list[list[object | None]] = [
                [None for _ in range(max_col)] for _ in range(max_row)
            ]
            for row, col, content in parsed_cells:
                matrix[row - 1][col - 1] = content

            sheet_matrices[sheet_name] = matrix

        calculator_root = ET.fromstring(archive.read(sheet_files["Calculator1"]))
        data_validations = calculator_root.find("a:dataValidations", NS)
        validation_map: dict[str, list[object]] = {}
        if data_validations is not None:
            for validation in data_validations.findall("a:dataValidation", NS):
                formula1 = validation.find("a:formula1", NS)
                if formula1 is None or not formula1.text:
                    continue
                options = resolve_validation_options(formula1.text, workbook_values)
                for ref in parse_sqref(validation.attrib["sqref"]):
                    validation_map[ref] = options

    build_fields = [
        {
            "ref": f"B{row}",
            "labelRef": f"A{row}",
            "options": validation_map.get(f"B{row}"),
            "defaultValue": workbook_values["Calculator1"].get(f"B{row}"),
        }
        for row in range(3, 63)
    ]

    weapon_fields = [
        {
            "ref": f"E{row}",
            "labelRef": f"D{row}",
            "options": validation_map.get(f"E{row}"),
            "defaultValue": workbook_values["Calculator1"].get(f"E{row}"),
        }
        for row in range(3, 8)
    ]

    payload = {
        "sheets": sheet_matrices,
        "buildFields": build_fields,
        "weaponFields": weapon_fields,
        "resultCell": "H12",
    }

    serialized = json.dumps(payload, separators=(",", ":"))
    OUTPUT.write_text(serialized)
    OUTPUT_JS.write_text(f"window.WORKBOOK_DATA={serialized};")
    print(f"wrote {OUTPUT}")
    print(f"wrote {OUTPUT_JS}")


if __name__ == "__main__":
    main()
