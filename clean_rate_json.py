"""
Clean rate JSON files in the output folder (standalone script for already-exported JSON).

Uses :func:`rate_to_json.clean_tariff_rows_twice` — **two** ``_clean_data`` passes — same as
``rate_to_json.main()`` and ``rate_creation.run_tariff_pipeline``. A single pass can leave
side-by-side dual-title sheets unexpanded; the second pass matches export → reload behaviour.

**Dual title row (side-by-side):** ``_clean_data`` ends with ``_expand_dual_title_row_side_by_side``.
Disable expansion with ``RATE_DISABLE_DUAL_TITLE_SPLIT=1``.
"""

import json
from pathlib import Path

from rate_to_json import clean_tariff_rows_twice

OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def clean_file(path: Path) -> None:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        print(f"  Skip (not a list): {path.name}")
        return
    original_rows = len(data)
    cleaned = clean_tariff_rows_twice(data)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, indent=2, ensure_ascii=False)
    print(f"  {path.name}: {original_rows} -> {len(cleaned)} rows")


def main():
    if not OUTPUT_DIR.exists():
        print(f"Output folder not found: {OUTPUT_DIR}")
        return
    json_files = list(OUTPUT_DIR.glob("*.json"))
    if not json_files:
        print("No JSON files in output folder.")
        return
    print("Cleaning JSON files...")
    for path in sorted(json_files):
        clean_file(path)
    print("Done.")


if __name__ == "__main__":
    main()
