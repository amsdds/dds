import json
import sys
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_URL_IRI = "https://dds.schiphol.nl/asset/"

SHEET_CONVENTION = "Conventies"
SHEET_VERSION = "Versie toetsingsregel"

def norm(s: str) -> str:
    return (s or "").strip()

def as_bool_ja_nee(s: str) -> bool:
    return norm(s).lower() == "ja"

def main(xlsx_path: str, out_json_path: str) -> None:
    xlsx = Path(xlsx_path)
    if not xlsx.exists():
        raise FileNotFoundError(f"Excel not found: {xlsx}")

    wb = load_workbook(filename=xlsx, data_only=True)

    # 1) Latest rule version from sheet "Versie toetsingsregel", cell B1
    latest = ""
    if SHEET_VERSION in wb.sheetnames:
        ws_v = wb[SHEET_VERSION]
        latest = norm(ws_v.cell(row=1, column=2).value)  # B1
    else:
        latest = ""

    # 2) Conventions
    if SHEET_CONVENTION not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{SHEET_CONVENTION}' not found in {xlsx.name}")

    ws = wb[SHEET_CONVENTION]

    rules = []
    # Your Java loops r = 1..lastRow (skips header row 0). Here: start at row 2.
    # Mapping (1-based column numbers):
    # col 2 -> iri suffix (Java row.getCell(1))
    # col 3 -> objectIdRequired (Java row.getCell(2))
    # col 4 -> aasRegex (Java row.getCell(3))
    # col 5 -> aasOpbouw (Java row.getCell(4))
    # col 6 -> aasVoorbeeld (Java row.getCell(5))
    # col 7 -> omschrijvingTemplate (Java row.getCell(6))
    # col 8 -> omschrijvingUitleg (Java row.getCell(7))
    # col 9 -> omschrijvingVoorbeeld (Java row.getCell(8))
    for r in range(2, ws.max_row + 1):
        iri_suffix = norm(ws.cell(row=r, column=2).value)
        object_id_required = as_bool_ja_nee(ws.cell(row=r, column=3).value)
        aas_regex = norm(ws.cell(row=r, column=4).value)

        if not iri_suffix or not aas_regex:
            continue

        iri = DEFAULT_URL_IRI + iri_suffix

        rule = {
            "iri": iri,
            "objectIdRequired": object_id_required,
            "aasRegex": aas_regex,
            "aasOpbouw": norm(ws.cell(row=r, column=5).value),
            "aasVoorbeeld": norm(ws.cell(row=r, column=6).value),
            "omschrijvingTemplate": norm(ws.cell(row=r, column=7).value),
            "omschrijvingUitleg": norm(ws.cell(row=r, column=8).value),
            "omschrijvingVoorbeeld": norm(ws.cell(row=r, column=9).value),
            "row": r  # handig voor debugging
        }
        rules.append(rule)

    out = {
        "latestRuleVersion": latest,
        "rules": rules
    }

    out_path = Path(out_json_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote {out_path} with {len(rules)} rules (latestRuleVersion='{latest}')")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python tools/export_excel_to_json.py <input.xlsx> <output.json>")
        sys.exit(2)
    main(sys.argv[1], sys.argv[2])
