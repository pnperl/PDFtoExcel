# PDFtoExcel (Offline)

Convert and combine multiple PDF statements (including password-protected PDFs) into **one Excel file**.

This project is designed for **offline use** and works well in **Pydroid 3**.

## Features
- Combine multiple PDFs into one Excel workbook.
- Supports encrypted/password-protected PDFs.
- Uses table extraction when available.
- Fallback parser for transaction-like lines (date/description/amount/balance).
- Adds metadata for source file, page, and extraction mode.
- 100% local/offline processing.

## Install (Pydroid 3)
Run these once in the Pydroid terminal:

```bash
pip install pandas openpyxl pdfplumber pypdf
```

## One-click run (Pydroid 3)
If your files are in the same folder, run:

```bash
python app.py statement1.pdf statement2.pdf -o combined_statements.xlsx
```

That is the one-command flow.

## Password-protected PDFs
### Option A: same password for all encrypted PDFs
```bash
python app.py jan.pdf feb.pdf --password "1234" -o combined.xlsx
```

### Option B: different password per file
Create `passwords.txt`:

```text
jan.pdf=1111
feb.pdf=2222
```

Run:

```bash
python app.py jan.pdf feb.pdf --password-file passwords.txt -o combined.xlsx
```

If a password is missing or incorrect, the app prompts securely in terminal.

## Notes for credit card statements
- Statement layouts vary by bank.
- Best results happen when PDFs contain selectable text (not scanned image only).
- For scanned image PDFs, OCR would be required (not included here).

## Output
Creates one Excel file with sheet `combined`, including:
- `source_file`
- `page`
- `extraction_mode`
- Extracted transaction columns
