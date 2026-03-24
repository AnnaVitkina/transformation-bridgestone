"""
Transform JSON from the output folder to a DataFrame and write to Excel.
Column order: Column0, Column1, ... Column9, Column10, ... then other keys alphabetically.
"""

import json
import re
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Please install pandas: pip install pandas openpyxl")
    raise SystemExit(1)

OUTPUT_DIR = Path(__file__).resolve().parent / "output"

_COLUMN_PATTERN = re.compile(r"^Column(\d+)$", re.IGNORECASE)


def _column_sort_key(name):
    """Sort Column0, Column1, ... Column9, Column10, ... then other keys."""
    m = _COLUMN_PATTERN.match(name)
    if m:
        return (0, int(m.group(1)))
    return (1, name)


def get_json_files():
    """List .json files in the output folder."""
    if not OUTPUT_DIR.exists():
        print(f"Output folder not found: {OUTPUT_DIR}")
        return []
    files = sorted(f.name for f in OUTPUT_DIR.iterdir() if f.suffix.lower() == ".json")
    return files


def main():
    files = get_json_files()
    if not files:
        print("No .json files found in the output folder.")
        return

    print("JSON files in 'output' folder:")
    for i, name in enumerate(files, 1):
        print(f"  {i}. {name}")
    choice = input("Which file number to transform? ").strip()
    try:
        idx = int(choice)
        if idx < 1 or idx > len(files):
            raise ValueError("Invalid number")
        filename = files[idx - 1]
    except (ValueError, IndexError):
        print("Invalid choice.")
        return

    path = OUTPUT_DIR / filename
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    df = pd.DataFrame(data)
    # Fix column order: Column0, Column1, ... Column10, ... then other keys
    df = df.reindex(columns=sorted(df.columns, key=_column_sort_key))

    out_path = path.with_suffix(".xlsx")
    df.to_excel(out_path, index=False, engine="openpyxl")

    print(f"\nSaved {len(df)} rows to: {out_path}")


if __name__ == "__main__":
    main()