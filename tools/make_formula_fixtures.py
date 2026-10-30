# /// script
# requires-python = ">=3.11"
# dependencies = ["xlsxwriter"]
# ///
"""One-off generator for the formula-note fixture pair (files are committed;
this script documents their provenance).

xlsxwriter, not openpyxl, because only xlsxwriter writes CACHED formula
results -- the state a real Excel save has, and the only state in which
"formula changed, value unchanged" is a verifiable claim.

Run once from anydiff-web: uv run tools/make_formula_fixtures.py
"""

from pathlib import Path
import re
import zipfile

import xlsxwriter

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures"

# (addr, kind, payload) where payload is (formula, cached) or a literal.
V1 = [
    ("B1", "lit", 7),          # -> edited literal (plain value change)
    ("C2", "form", ("A2*2", 4)),   # A2=2; -> formula edit, SAME result
    ("D3", "form", ("A2+1", 3)),   # -> formula edit, DIFFERENT result
    ("E4", "lit", 1),           # -> replaced by formula with same value
    ("F5", "lit", 5),           # -> replaced by formula, NO cached result
    ("G6", "bool", True),       # -> becomes number 1 (bool-vs-int guard)
    ("H7", "form", ("A2+8", 10)),  # -> deleted (formula cell removed)
]
V2 = [
    ("B1", "lit", 8),
    ("C2", "form", ("A2*3", 4)),   # 2*3 == 4: formula moved, value held
    ("D3", "form", ("A2+2", 4)),   # value 3 -> 4 too
    ("E4", "form", ("1", 1)),
    ("F5", "form", ("A2+3", None)),  # no cached value written
    ("G6", "lit", 1),
    # H7 absent
]


def write(path: str, cells: list) -> None:
    wb = xlsxwriter.Workbook(str(OUT / path))
    ws = wb.add_worksheet("Formulas")
    ws.write("A2", 2)  # shared plain input both versions agree on
    for addr, kind, payload in cells:
        if kind == "form":
            formula, cached = payload
            if cached is None:
                ws.write_formula(addr, "=" + formula)
            else:
                ws.write_formula(addr, "=" + formula, None, cached)
        elif kind == "bool":
            ws.write_boolean(addr, payload)
        else:
            ws.write_number(addr, payload)
    wb.close()


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    write("formula-notes.xlsx", V1)
    write("formula-notes-v2.xlsx", V2)
    # xlsxwriter defaults a missing cached result to <v>0</v>; a real
    # "never-calculated" file has NO <v> at all. Strip it so F5 exercises the
    # engine's cached-values-unavailable rule instead of a fake 5 -> 0 edit.
    path = OUT / "formula-notes-v2.xlsx"
    src = zipfile.ZipFile(path)
    parts = {i.filename: src.read(i.filename) for i in src.infolist()}
    src.close()
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    stripped, n = re.subn(r'(<c r="F5"[^>]*>.*?)(<v>[^<]*</v>)', r"\1", sheet, count=1)
    assert n == 1, "F5 cell not found for <v> stripping"
    parts["xl/worksheets/sheet1.xml"] = stripped.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts.items():
            out.writestr(name, data)
    for p in sorted(OUT.glob("formula-notes*.xlsx")):
        print(p.name, p.stat().st_size, "bytes")
