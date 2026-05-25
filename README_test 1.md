# Advanced Privacy Anonymizer Pro v3

Tools for extracting text/table data, reviewing detected sensitive values, and writing anonymized outputs. The v3 scripts now support Excel workbooks by converting `.xlsx` / `.xlsm` files into JSON-like data before using the existing anonymization workflow.

## Supported Files

`anonymize_data_pro_v3.py`

- `.json`: recursively anonymizes JSON values and saves `_pro_anonymized.json`
- `.pdf`: extracts text and saves `_pro_anonymized.txt`
- `.docx`: anonymizes paragraph text and saves `_pro_anonymized.docx`
- `.txt` and other text files: saves `_pro_anonymized<extension>`
- `.xlsx` / `.xlsm`: converts workbook sheets to JSON and saves `_pro_anonymized.json`

`pdf_docx_to_json_ano_v3.py`

- `.pdf`: extracts pages and tables to JSON, then anonymizes
- `.docx`: extracts paragraphs/tables to JSON, then anonymizes
- `.xlsx` / `.xlsm`: converts workbook sheets to JSON, then anonymizes
- Output: `<input_name>_anonymized.json`

## Installation

Install the Python packages used by the scripts:

```bash
pip install faker PyPDF2 pypdf pdfplumber python-docx pandas openpyxl
```

`pandas` and `openpyxl` are required only for Excel input.

## Excel Conversion Behavior

Excel input follows the same practical rules as `convertexceltojson3_worked_string-convert_books_for_Gen_AI.py`:

- All workbook sheets are loaded.
- Rows are treated as data (`header=None`), so no row is promoted to column headers.
- Completely empty rows are dropped.
- Empty cells become `null` in JSON.
- A one-column sheet becomes a simple list.
- A multi-column sheet becomes a list of row lists.
- Dates and other non-JSON scalar values are converted to JSON-friendly values before saving.

The workbook JSON structure is:

```json
{
  "metadata": {
    "source_type": "EXCEL",
    "sheet_count": 2,
    "sheets": ["Sheet1", "Sheet2"]
  },
  "sheets": {
    "Sheet1": [["Name", "Email"], ["Alice", "alice@example.com"]],
    "Sheet2": ["single column value"]
  }
}
```

## Usage

Run direct anonymization:

```bash
python anonymize_data_pro_v3.py path/to/input.xlsx
python anonymize_data_pro_v3.py path/to/input.json
python anonymize_data_pro_v3.py path/to/input.docx
```

Run extraction-to-JSON anonymization:

```bash
python pdf_docx_to_json_ano_v3.py path/to/input.pdf
python pdf_docx_to_json_ano_v3.py path/to/input.docx
python pdf_docx_to_json_ano_v3.py path/to/input.xlsx
```

## Review Workflow

After scanning, the scripts show detected sensitive data:

- `Accumulated`: values already known from `sensitive_data_list.json`
- `New (Regex)`: values detected by built-in regex patterns
- `Manual`: exact names you add during the prompt

At the approval prompt:

- Type `y` to approve all detections.
- Type `n` or press Enter to cancel anonymization.
- Type IDs such as `1,3,5` to approve only selected detections.

Approved new detections are appended to `sensitive_data_list.json` so future sessions can detect them as accumulated values.

## Outputs and Logs

- `anonymize_data_pro_v3.py` writes replacement details to `anonymization_pro_log.txt`.
- `pdf_docx_to_json_ano_v3.py` writes replacement details to `extraction_anonymization_log.txt`.
- Both scripts preserve the nested JSON/list structure and anonymize string values recursively.

## Manual Sensitive List Entries

You can manually add entries to `sensitive_data_list.json`:

```json
{
  "pattern": "Secret Project X",
  "category": "project",
  "sensitive": false
}
```

Set `"sensitive": true` for case-sensitive matching.
