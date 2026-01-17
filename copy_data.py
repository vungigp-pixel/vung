#!/usr/bin/env python3
import argparse
import importlib.util
import os
import re
import shutil
from pathlib import Path

def ensure_openpyxl_available() -> None:
    if importlib.util.find_spec("openpyxl") is None:
        raise SystemExit(
            "Missing dependency: openpyxl. Install it with 'pip install openpyxl'."
        )


def load_mapping(xlsx_path: Path) -> dict[str, str]:
    ensure_openpyxl_available()
    import openpyxl

    workbook = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    sheet = workbook.active
    mapping: dict[str, str] = {}
    for row in sheet.iter_rows(values_only=True):
        if not row:
            continue
        source_name = str(row[0]).strip() if row[0] is not None else ""
        dest_name = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if source_name and dest_name:
            mapping[source_name] = dest_name
    return mapping


def default_paths() -> tuple[str, str, str]:
    if os.name == "nt":
        return (
            r"L:\KonTum_2021_2025",
            r"\\wsl.localhost\Ubuntu-22.04\home\dc\cpi\wintomseed\TKT",
            r"\\wsl.localhost\Ubuntu-22.04\home\dc\cpi\wintomseed\rename.xlsx",
        )
    return (
        r"/mnt/l/KonTum_2021_2025",
        r"/home/dc/cpi/wintomseed/TKT",
        r"/home/dc/cpi/wintomseed/rename.xlsx",
    )


def parse_args() -> argparse.Namespace:
    source_root_default, dest_root_default, mapping_default = default_paths()
    parser = argparse.ArgumentParser(
        description=(
            "Copy data from KonTum_2021_2025 into TKT with rename mapping from Excel."
        )
    )
    parser.add_argument(
        "--source-root",
        default=source_root_default,
        help="Source root directory.",
    )
    parser.add_argument(
        "--dest-root",
        default=dest_root_default,
        help="Destination root directory.",
    )
    parser.add_argument(
        "--mapping-xlsx",
        default=mapping_default,
        help="Excel file with mapping from source folder (col A) to dest folder (col B).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned copy operations without copying files.",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="Skip files that already exist in the destination.",
    )
    return parser.parse_args()


def iter_source_files(source_root: Path):
    level1_pattern = "DataTramSonTayQN*"
    yymmdd_re = re.compile(r"^\d{6}$")
    yymmddhh_re = re.compile(r"^\d{8}$")
    filename_re = re.compile(r"^\d{8}\.\d{2}$")

    for level1_dir in source_root.glob(level1_pattern):
        if not level1_dir.is_dir():
            continue
        short_dir = level1_dir / "SHORT"
        if not short_dir.is_dir():
            continue
        for day_dir in short_dir.iterdir():
            if not day_dir.is_dir() or not yymmdd_re.match(day_dir.name):
                continue
            for hour_dir in day_dir.iterdir():
                if not hour_dir.is_dir() or not yymmddhh_re.match(hour_dir.name):
                    continue
                for data_file in hour_dir.iterdir():
                    if data_file.is_file() and filename_re.match(data_file.name):
                        yield level1_dir.name, day_dir.name, hour_dir.name, data_file


def match_mapping(level1_name: str, mapping: dict[str, str]) -> str | None:
    best_match = None
    best_len = -1
    for source_pattern, dest_name in mapping.items():
        prefix = source_pattern.rstrip("*")
        if not prefix:
            continue
        if level1_name.startswith(prefix) and len(prefix) > best_len:
            best_match = dest_name
            best_len = len(prefix)
    return best_match


def copy_files(
    source_root: Path,
    dest_root: Path,
    mapping: dict[str, str],
    dry_run: bool,
    skip_existing: bool,
) -> int:
    copied_count = 0
    for level1_name, day_name, hour_name, data_file in iter_source_files(source_root):
        mapped_level1 = match_mapping(level1_name, mapping)
        if not mapped_level1:
            print(f"WARN: No mapping for {level1_name}, skipping {data_file}")
            continue
        dest_dir = dest_root / mapped_level1 / day_name / hour_name
        dest_file = dest_dir / data_file.name
        if skip_existing and dest_file.exists():
            print(f"SKIP: Exists {dest_file}")
            continue
        if dry_run:
            print(f"DRY RUN: {data_file} -> {dest_file}")
        else:
            dest_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(data_file, dest_file)
            print(f"COPIED: {data_file} -> {dest_file}")
            copied_count += 1
    return copied_count


def main() -> None:
    args = parse_args()
    source_root = Path(args.source_root)
    dest_root = Path(args.dest_root)
    mapping_path = Path(args.mapping_xlsx)

    if not source_root.exists():
        raise SystemExit(f"Source root not found: {source_root}")
    if not mapping_path.exists():
        raise SystemExit(f"Mapping file not found: {mapping_path}")

    mapping = load_mapping(mapping_path)
    copied = copy_files(
        source_root=source_root,
        dest_root=dest_root,
        mapping=mapping,
        dry_run=args.dry_run,
        skip_existing=args.skip_existing,
    )
    if not args.dry_run:
        print(f"Done. Copied {copied} file(s).")


if __name__ == "__main__":
    main()
