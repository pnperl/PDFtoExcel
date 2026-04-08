#!/usr/bin/env python3
"""
Offline PDF-to-Excel combiner for bank/credit-card statements.

Features:
- Combines multiple PDFs into one Excel workbook.
- Supports password-protected PDFs.
- Tries table extraction first; falls back to line-based extraction.
- Adds source metadata (file, page, extraction mode).
- Runs fully offline.
- Designed to be easy to run in Pydroid 3 (single command).
"""

from __future__ import annotations

import argparse
import getpass
import os
import re
import sys
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
import pdfplumber
from pypdf import PdfReader


@dataclass
class PDFInput:
    path: str
    password: str | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert and combine one or more PDF statements to a single Excel file "
            "fully offline."
        )
    )
    parser.add_argument(
        "pdfs",
        nargs="+",
        help="PDF paths. Example: statements/jan.pdf statements/feb.pdf",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="combined_statements.xlsx",
        help="Output Excel file path (default: combined_statements.xlsx)",
    )
    parser.add_argument(
        "--password",
        default=None,
        help=(
            "One password applied to every encrypted PDF. "
            "If omitted, app prompts when needed."
        ),
    )
    parser.add_argument(
        "--password-file",
        default=None,
        help=(
            "Optional file with per-PDF passwords. Format: "
            "filename.pdf=password (one per line)."
        ),
    )
    return parser.parse_args()


def load_password_map(password_file: str | None) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not password_file:
        return mapping

    with open(password_file, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                raise ValueError(f"Invalid password mapping line: {line}")
            filename, password = line.split("=", 1)
            mapping[filename.strip()] = password.strip()

    return mapping


def ensure_pdf_readable(path: str, password: str | None) -> str | None:
    """Validate PDF access using pypdf and return working password (or None)."""
    with open(path, "rb") as f:
        reader = PdfReader(f)
        if not reader.is_encrypted:
            return None

        if password and reader.decrypt(password):
            return password

        entered = getpass.getpass(f"Password required for '{path}': ")
        if reader.decrypt(entered):
            return entered

        raise ValueError(f"Wrong password for encrypted file: {path}")


def normalize_row(cells: list[str], desired_len: int) -> list[str]:
    cells = [c.strip() if isinstance(c, str) else "" for c in cells]
    if len(cells) < desired_len:
        cells.extend([""] * (desired_len - len(cells)))
    return cells[:desired_len]


def to_amount_token(token: str) -> str:
    cleaned = token.replace(",", "").replace("$", "").strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
        return cleaned
    return token


def parse_lines_as_statement_rows(text: str) -> list[list[str]]:
    """
    Fallback parser for statement-like lines.

    Heuristic output columns:
    Date | Description | Amount | Balance | Raw Line
    """
    rows: list[list[str]] = []
    if not text:
        return rows

    date_pat = re.compile(r"^(\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)\b")

    for raw_line in text.splitlines():
        line = " ".join(raw_line.split())
        if not line:
            continue

        date_match = date_pat.match(line)
        if not date_match:
            continue

        date = date_match.group(1)
        remainder = line[date_match.end() :].strip()

        tail_nums = re.findall(r"-?\$?\d[\d,]*(?:\.\d{2})", remainder)
        amount = ""
        balance = ""
        desc = remainder

        if tail_nums:
            amount = to_amount_token(tail_nums[-1])
            desc = remainder[: remainder.rfind(tail_nums[-1])].strip()
            if len(tail_nums) >= 2:
                balance = amount
                amount = to_amount_token(tail_nums[-2])
                desc = remainder[: remainder.rfind(tail_nums[-2])].strip()

        rows.append([date, desc, amount, balance, line])

    return rows


def extract_rows_from_pdf(pdf: PDFInput) -> list[dict[str, str]]:
    extracted: list[dict[str, str]] = []

    with pdfplumber.open(pdf.path, password=pdf.password) as doc:
        for page_num, page in enumerate(doc.pages, start=1):
            tables = page.extract_tables() or []

            if tables:
                for table_idx, table in enumerate(tables, start=1):
                    if not table:
                        continue
                    max_len = max(len(r) for r in table if r) if table else 0
                    if max_len == 0:
                        continue

                    headers = normalize_row(table[0], max_len)
                    if not any(headers):
                        headers = [f"Column_{i}" for i in range(1, max_len + 1)]

                    for row in table[1:]:
                        if not row:
                            continue
                        values = normalize_row(row, max_len)
                        if not any(values):
                            continue

                        item = {
                            "source_file": os.path.basename(pdf.path),
                            "page": str(page_num),
                            "extraction_mode": f"table_{table_idx}",
                        }
                        for i, (h, v) in enumerate(zip(headers, values), start=1):
                            key = h or f"Column_{i}"
                            item[key] = v
                        extracted.append(item)
            else:
                text = page.extract_text() or ""
                parsed_rows = parse_lines_as_statement_rows(text)
                for row in parsed_rows:
                    extracted.append(
                        {
                            "source_file": os.path.basename(pdf.path),
                            "page": str(page_num),
                            "extraction_mode": "line_fallback",
                            "Date": row[0],
                            "Description": row[1],
                            "Amount": row[2],
                            "Balance": row[3],
                            "Raw Line": row[4],
                        }
                    )

    return extracted


def build_pdf_inputs(paths: Iterable[str], default_password: str | None, password_map: dict[str, str]) -> list[PDFInput]:
    inputs: list[PDFInput] = []
    for p in paths:
        if not os.path.isfile(p):
            raise FileNotFoundError(f"PDF not found: {p}")
        filename = os.path.basename(p)
        selected_password = password_map.get(filename, default_password)
        working_password = ensure_pdf_readable(p, selected_password)
        inputs.append(PDFInput(path=p, password=working_password))
    return inputs


def write_excel(all_rows: list[dict[str, str]], output_path: str) -> None:
    if not all_rows:
        raise RuntimeError("No rows extracted. Confirm PDFs contain selectable text or tables.")

    df = pd.DataFrame(all_rows)
    meta_cols = [c for c in ["source_file", "page", "extraction_mode"] if c in df.columns]
    other_cols = sorted([c for c in df.columns if c not in meta_cols])
    df = df[meta_cols + other_cols]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="combined", index=False)


def main() -> int:
    args = parse_args()

    try:
        password_map = load_password_map(args.password_file)
        pdf_inputs = build_pdf_inputs(args.pdfs, args.password, password_map)

        all_rows: list[dict[str, str]] = []
        for pdf in pdf_inputs:
            print(f"Processing: {pdf.path}")
            all_rows.extend(extract_rows_from_pdf(pdf))

        write_excel(all_rows, args.output)
        print(f"Done. Wrote combined Excel file: {args.output}")
        return 0

    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
