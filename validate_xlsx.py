from __future__ import annotations

import argparse
import posixpath
import sys
import zipfile
from pathlib import Path, PurePosixPath
import xml.etree.ElementTree as ET


XML_NS = {
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def normalize_zip_path(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")

    base_dir = posixpath.dirname(base_part)
    normalized = posixpath.normpath(posixpath.join(base_dir, target))
    return normalized.lstrip("/")


def find_main_workbook(zip_names: set[str]) -> str | None:
    if "xl/workbook.xml" in zip_names:
        return "xl/workbook.xml"

    for name in zip_names:
        if name.endswith("/workbook.xml"):
            return name

    return None


def validate_xml_parts(zf: zipfile.ZipFile, verbose: bool) -> list[str]:
    errors: list[str] = []

    for name in sorted(zf.namelist()):
        if not name.endswith(".xml") and not name.endswith(".rels"):
            continue

        try:
            ET.fromstring(zf.read(name))
            if verbose:
                print(f"OK XML: {name}")
        except ET.ParseError as exc:
            errors.append(f"Invalid XML in {name}: {exc}")

    return errors


def validate_relationship_targets(zf: zipfile.ZipFile, verbose: bool) -> list[str]:
    errors: list[str] = []
    zip_names = set(zf.namelist())

    for rels_name in sorted(name for name in zip_names if name.endswith(".rels")):
        try:
            root = ET.fromstring(zf.read(rels_name))
        except ET.ParseError:
            continue

        rels_parent = str(PurePosixPath(rels_name).parent)
        if rels_parent.endswith("_rels"):
            source_part = str(PurePosixPath(rels_parent).parent)
        else:
            source_part = rels_parent

        source_file = source_part.rstrip("/")
        if source_file == "":
            source_file = ""
        elif not source_file.endswith(".xml"):
            source_file = posixpath.join(source_file, "")

        for rel in root.findall("r:Relationship", XML_NS):
            target = rel.get("Target", "")
            target_mode = rel.get("TargetMode", "")
            rel_id = rel.get("Id", "")

            if target_mode == "External":
                if verbose:
                    print(f"SKIP External rel {rels_name}#{rel_id} -> {target}")
                continue

            if not target:
                errors.append(f"Empty target in {rels_name} (Id={rel_id})")
                continue

            resolved = normalize_zip_path(source_file, target)
            if resolved not in zip_names:
                errors.append(
                    f"Broken rel target in {rels_name} (Id={rel_id}): {target} -> {resolved}"
                )
            elif verbose:
                print(f"OK REL: {rels_name}#{rel_id} -> {resolved}")

    return errors


def validate_required_parts(zf: zipfile.ZipFile) -> list[str]:
    errors: list[str] = []
    zip_names = set(zf.namelist())

    required = ["[Content_Types].xml", "_rels/.rels"]
    for part in required:
        if part not in zip_names:
            errors.append(f"Missing required part: {part}")

    workbook_part = find_main_workbook(zip_names)
    if not workbook_part:
        errors.append("Missing workbook part (expected xl/workbook.xml)")

    return errors


def run_checks(xlsx_path: Path, verbose: bool) -> int:
    if not xlsx_path.exists():
        print(f"ERROR: file not found: {xlsx_path}")
        return 2

    if not xlsx_path.is_file():
        print(f"ERROR: not a file: {xlsx_path}")
        return 2

    errors: list[str] = []

    try:
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            errors.extend(validate_required_parts(zf))
            errors.extend(validate_xml_parts(zf, verbose=verbose))
            errors.extend(validate_relationship_targets(zf, verbose=verbose))
    except zipfile.BadZipFile as exc:
        print(f"ERROR: invalid XLSX zip container: {exc}")
        return 2

    if errors:
        print("XLSX validation found issues:")
        for issue in errors:
            print(f"- {issue}")
        return 1

    print("XLSX validation passed (zip/xml/rels checks).")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate XLSX package structure and XML integrity for debugging Excel repair issues."
    )
    parser.add_argument("xlsx", help="Path to .xlsx file")
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print every checked XML part and relationship",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return run_checks(Path(args.xlsx), verbose=args.verbose)


if __name__ == "__main__":
    sys.exit(main())
