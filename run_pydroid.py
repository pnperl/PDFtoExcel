#!/usr/bin/env python3
"""One-tap interactive runner for Pydroid 3."""

from __future__ import annotations

import os
from pathlib import Path

from app import build_pdf_inputs, extract_rows_from_pdf, load_password_map, write_excel


def ask(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or (default or "")


def list_pdfs(folder: Path) -> list[str]:
    return sorted(str(p) for p in folder.glob("*.pdf") if p.is_file())


def main() -> int:
    print("PDF to Excel (Offline) - Pydroid runner")
    default_folder = "/storage/emulated/0/Download"
    folder = Path(ask("Folder containing PDFs", default_folder)).expanduser()

    if not folder.exists() or not folder.is_dir():
        print(f"Error: folder not found -> {folder}")
        return 1

    pdfs = list_pdfs(folder)
    if not pdfs:
        print(f"No PDF files found in: {folder}")
        return 1

    print("\nFound PDFs:")
    for i, p in enumerate(pdfs, start=1):
        print(f"  {i}. {os.path.basename(p)}")

    pick = ask(
        "Choose files by number (comma separated) or type 'all'",
        "all",
    ).lower()

    selected_paths: list[str] = []
    if pick == "all":
        selected_paths = pdfs
    else:
        try:
            indices = [int(x.strip()) for x in pick.split(",") if x.strip()]
            selected_paths = [pdfs[i - 1] for i in indices]
        except Exception:
            print("Invalid selection. Example: 1,2,5")
            return 1

    output_name = ask("Output Excel file name", "combined_statements.xlsx")
    output_path = str(folder / output_name)

    default_password = ask("Common PDF password (leave blank if none)", "") or None
    password_file_raw = ask("Password mapping file name (optional)", "")
    password_file = str(folder / password_file_raw) if password_file_raw else None

    password_map = load_password_map(password_file)
    inputs = build_pdf_inputs(selected_paths, default_password, password_map)

    all_rows = []
    for pdf in inputs:
        print(f"Processing: {pdf.path}")
        all_rows.extend(extract_rows_from_pdf(pdf))

    write_excel(all_rows, output_path)
    print(f"\nDone ✅\nSaved: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
