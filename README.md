# PDFtoExcel (Offline)

Convert and combine multiple PDF statements (including password-protected PDFs) into **one Excel file**.

This project is designed for **offline use** and works in **Pydroid 3 / Android**.

## Why this version works better on Pydroid
- Removed heavy `pandas` dependency (common install failure on mobile).
- Core install now needs only:
  - `pypdf`
  - `openpyxl`
- `pdfplumber` is optional (for better table extraction).

## Install (Pydroid 3)
Run in Pydroid terminal:

```bash
pip install pypdf openpyxl
```

Optional (better tables):

```bash
pip install pdfplumber
```

If library installation failed before, run this first and try again:

```bash
python -m pip install --upgrade pip setuptools wheel
```

## How to run app on mobile (Android + Pydroid 3)

### Method 1 (Recommended): one-tap interactive runner
1. Open **Pydroid 3**.
2. Grant storage permission.
3. Put PDFs in one folder (example: `/storage/emulated/0/Download`).
4. Open `run_pydroid.py` and tap **Run ▶**.
5. Enter folder, choose files (`all` or `1,2,3`), and output name.
6. Enter password options if needed.
7. Excel file saves to the same folder.

Command form:

```bash
python run_pydroid.py
```

### Method 2: direct CLI command

```bash
python app.py statement1.pdf statement2.pdf -o combined_statements.xlsx
```

## Password-protected PDFs
### Same password for all
```bash
python app.py jan.pdf feb.pdf --password "1234" -o combined.xlsx
```

### Different passwords per file
Create `passwords.txt`:

```text
jan.pdf=1111
feb.pdf=2222
```

Run:

```bash
python app.py jan.pdf feb.pdf --password-file passwords.txt -o combined.xlsx
```

## Output
Creates one Excel file, sheet name `combined`, with metadata columns:
- `source_file`
- `page`
- `extraction_mode`

## Notes
- If `pdfplumber` is not installed, the app still works using text fallback parsing.
- For scanned image-only PDFs, OCR is required (not included).
