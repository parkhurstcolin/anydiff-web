"""Dump golden Change records from the Python engine for the web port's tests.

The Python engine is the spec: tests/golden/*.json are its exact output on
fixture pairs whose bytes were verified identical to public/demo/ copies
(sha256). The TS differ must reproduce these field for field.

Run from the anydiff-web root:
  uv run --project ../Builder/projects/anydiff python tools/gen_golden.py
"""

import json
import sys
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[2] / "Builder" / "projects" / "anydiff"
sys.path.insert(0, str(ENGINE / "src"))

from anydiff.diff import diff_docx, diff_xlsx  # noqa: E402
from anydiff.extractors.docx import extract_docx  # noqa: E402
from anydiff.extractors.xlsx import extract_xlsx  # noqa: E402

FIXTURES = ENGINE / "tests" / "fixtures"
LOCAL_FIXTURES = Path(__file__).resolve().parents[1] / "tests" / "fixtures"
OUT = Path(__file__).resolve().parents[1] / "tests" / "golden"


def fixture(name: str) -> Path:
    local = LOCAL_FIXTURES / name
    return local if local.exists() else FIXTURES / name

PAIRS = {
    "q3-forecast.json": ("Q3-forecast.xlsx", "Q3-forecast-v2.xlsx"),
    "q3-identical.json": ("Q3-forecast.xlsx", "Q3-forecast.xlsx"),
    "formula-notes.json": ("formula-notes.xlsx", "formula-notes-v2.xlsx"),
    "meeting-notes.json": ("meeting-notes.docx", "meeting-notes-v2.docx"),
    "meeting-notes-identical.json": ("meeting-notes.docx", "meeting-notes.docx"),
}

if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    for name, (old, new) in PAIRS.items():
        if old.endswith(".docx"):
            changes = diff_docx(extract_docx(fixture(old)), extract_docx(fixture(new)), display=old)
        else:
            changes = diff_xlsx(extract_xlsx(fixture(old)), extract_xlsx(fixture(new)), display=old)
        data = [c.to_dict() for c in changes]
        (OUT / name).write_text(json.dumps(data, indent=1, default=str) + "\n")
        print(f"{name}: {len(data)} changes")
