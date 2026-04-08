# PDFtoExcel (Offline)

Convert and combine multiple PDF statements (including password-protected PDFs) into **one Excel file**.

This project is designed for **offline use** and works well in **Pydroid 3 / Android**.

## Features
- Combine multiple PDFs into one Excel workbook.
- Supports encrypted/password-protected PDFs.
- Uses table extraction when available.
- Fallback parser for transaction-like lines (date/description/amount/balance).
- Adds metadata for source file, page, and extraction mode.
- 100% local/offline processing.

## Install (Pydroid 3)
Run these once in Pydroid terminal:

```bash
pip install -r requirements.txt
```

If `pip install -r requirements.txt` fails, run:

```bash
pip install pandas openpyxl pdfplumber pypdf
```

---

## How to run app on mobile (Android + Pydroid 3)

### Method 1 (Recommended): one-tap interactive runner
1. Open **Pydroid 3**.
2. Give storage permissions if prompted.
3. Put your PDF statements in one folder (example: `/storage/emulated/0/Download`).
4. Open `run_pydroid.py` and tap the **Run ▶** button.
5. Follow prompts:
   - PDF folder path
   - file selection (`all` or numbers like `1,2,4`)
   - output Excel name
   - optional password and optional password mapping file
6. Your Excel output is saved in the same folder you selected.

One command equivalent:

```bash
python run_pydroid.py
```

### Method 2: direct CLI command

```bash
python app.py statement1.pdf statement2.pdf -o combined_statements.xlsx
```

---

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

If a password is missing/incorrect, the app prompts securely in terminal.

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
