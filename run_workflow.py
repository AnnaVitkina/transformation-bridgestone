"""
Workflow entrypoint for Google Colab / Drive and for a local machine.

Why ``ModuleNotFoundError: rate_to_json``: with ``exec(open(...).read())``, Python does not add
the repo folder to ``sys.path``. This script prepends it automatically; if the repo is not at
``/content/transformation-bridgestone``, set ``TRANSFORMATION_BRIDGESTONE_REPO`` before running.

Local usage (from this repo, after ``pip install openpyxl pandas``):

  python run_workflow.py
  python run_workflow.py --root /path/to/RMT_Bridgestone

On a local machine the data root defaults to this repo folder (``input/``, ``output/``,
``previous_rate/`` or ``previous rate/``, ``processing/``, ``archive/``). Override with
``--root`` or ``RMT_BRIDGESTONE_ROOT``.

Colab usage (after mounting Drive and cloning the repo under /content/transformation-bridgestone):

  import os
  os.chdir("/content/transformation-bridgestone")
  # ``exec`` does not set ``__name__`` to ``"__main__"`` unless you pass globals:
  exec(
      compile(open("run_workflow.py", encoding="utf-8").read(), "run_workflow.py", "exec"),
      {"__name__": "__main__"},
  )
  # Or: ``from run_workflow import main; main()``

Or with explicit root (Colab defaults to the shared Drive path if that folder exists):

  os.environ["RMT_BRIDGESTONE_ROOT"] = "/path/to/RMT_Bridgestone"
  from run_workflow import main as _run_workflow
  _run_workflow()

Steps:
  1) Point rate_to_json, rate_card_extraction, and rate_creation at the data folders.
  2) Run rate_creation.main() (interactive prompts).
  3) Rename *_rate_matrix.xlsx / .json to match the previous-rate workbook stem with v.N bumped (v.1 -> v.2, etc.).
  4) Move processed tariff xlsx files from input/ to archive/.

Data layout (override with RMT_BRIDGESTONE_ROOT or --root):
  .../input
  .../output
  .../previous_rate   (``previous rate`` is also accepted if that folder already exists)
  .../processing
  .../archive
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


def _repo_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path(
            os.environ.get(
                "TRANSFORMATION_BRIDGESTONE_REPO",
                "/content/transformation-bridgestone",
            )
        ).resolve()


def _ensure_repo_on_sys_path() -> None:
    """Colab ``exec(open(...).read())`` does not put the repo on ``sys.path``; sibling imports fail."""
    s = str(_repo_dir())
    if s not in sys.path:
        sys.path.insert(0, s)


_ensure_repo_on_sys_path()

import json
import re
import shutil
from datetime import datetime

COLAB_DEFAULT_ROOT = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Bridgestone/Fiege"
)
DEFAULT_ROOT = COLAB_DEFAULT_ROOT


def _resolve_root(cli_root: str | None = None) -> Path:
    if cli_root:
        return Path(cli_root).expanduser().resolve()
    env = os.environ.get("RMT_BRIDGESTONE_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()
    colab = Path(COLAB_DEFAULT_ROOT)
    if colab.is_dir():
        return colab.resolve()
    return _repo_dir()


def _previous_rate_dir(root: Path) -> Path:
    """Prefer ``previous_rate``; reuse an existing ``previous rate`` folder (local scripts)."""
    underscore = root / "previous_rate"
    spaced = root / "previous rate"
    if underscore.is_dir():
        return underscore
    if spaced.is_dir():
        return spaced
    return underscore


def _apply_data_paths(root: Path) -> None:
    root = root.resolve()
    inp = root / "input"
    prev = _previous_rate_dir(root)
    out = root / "output"
    proc = root / "processing"
    import rate_to_json as rt

    rt.INPUT_DIR = inp
    import rate_card_extraction as rce

    rce.INPUT_DIR = prev
    rce.OUTPUT_DIR = out
    import rate_creation as rc

    rc.TARIFF_INPUT_DIR = inp
    rc.PREVIOUS_RATE_DIR = prev
    rc.OUTPUT_DIR = out
    rc.PROCESSING_DIR = proc


def bump_version_in_stem(stem: str) -> str:
    """Bump the last `` v.N`` in the stem (e.g. v.1 -> v.2). If none, append `` v.2``."""
    matches = list(re.finditer(r"(\s+v\.(\d+)\b)", stem, flags=re.IGNORECASE))
    if not matches:
        return stem.rstrip() + " v.2"
    m = matches[-1]
    n = int(m.group(2))
    return stem[: m.start(1)] + f" v.{n + 1}" + stem[m.end(1) :]


def _unique_dest(archive_dir: Path, name: str) -> Path:
    dest = archive_dir / name
    if not dest.exists():
        return dest
    stem, suf = Path(name).stem, Path(name).suffix
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return archive_dir / f"{stem}_{ts}{suf}"


def _rename_pair(
    folder: Path,
    old_stem: str,
    new_stem: str,
) -> None:
    for ext in (".xlsx", ".json"):
        old = folder / f"{old_stem}{ext}"
        new = folder / f"{new_stem}{ext}"
        if not old.exists():
            continue
        if new.exists() and new != old:
            raise FileExistsError(f"Target already exists: {new}")
        old.rename(new)


def post_process_after_rate_creation(root: Path) -> None:
    meta_path = root / "processing" / "last_run_workflow.json"
    if not meta_path.is_file():
        print("No last_run_workflow.json — skipping rename/archive (run may have exited early).")
        return
    with open(meta_path, encoding="utf-8") as f:
        meta = json.load(f)

    prev_stem = Path(meta["prev_rate_file"]).stem
    bumped = bump_version_in_stem(prev_stem)
    old_base = Path(meta["output_matrix_basename"]).stem
    new_base = f"{bumped}_rate_matrix"

    if old_base == new_base:
        print(f"Output name already matches bumped previous-rate stem: {new_base}")
    else:
        out_dir = root / "output"
        proc_dir = root / "processing"
        _rename_pair(out_dir, old_base, new_base)
        _rename_pair(proc_dir, old_base, new_base)
        print(f"Renamed matrix files to stem: {new_base}")

    archive_dir = root / "archive"
    archive_dir.mkdir(parents=True, exist_ok=True)
    for p in meta.get("tariff_input_files") or []:
        src = Path(p)
        if not src.is_file():
            print(f"Skip archive (missing): {src}")
            continue
        dest = _unique_dest(archive_dir, src.name)
        shutil.move(str(src), str(dest))
        print(f"Archived: {src.name} -> {dest}")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Run the Bridgestone rate-matrix workflow (local or Colab)."
    )
    parser.add_argument(
        "--root",
        default=None,
        help=(
            "Data folder containing input/, previous_rate/, output/, processing/, archive/. "
            "Defaults to RMT_BRIDGESTONE_ROOT, then the Colab Drive path if it exists, "
            "otherwise this repo directory."
        ),
    )
    args, _unknown = parser.parse_known_args(argv)

    root = _resolve_root(args.root)
    print(f"Data root: {root}")
    prev = _previous_rate_dir(root)
    for folder in (root / "input", root / "output", prev, root / "processing", root / "archive"):
        folder.mkdir(parents=True, exist_ok=True)

    _apply_data_paths(root)

    import rate_creation as rc

    rc.main()
    post_process_after_rate_creation(root)


if __name__ == "__main__":
    main()
