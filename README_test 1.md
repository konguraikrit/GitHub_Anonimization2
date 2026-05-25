# Advanced Privacy Anonymizer Pro v3

Tools for extracting text/table data, reviewing detected sensitive values, and writing anonymized outputs. The v3 scripts support Excel workbooks, Thai NER-based entity detection, and an expanded set of Thai PII regex patterns.

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

Install the required Python packages:

```bash
pip install faker PyPDF2 pypdf pdfplumber python-docx pandas openpyxl
```

`pandas` and `openpyxl` are required only for Excel input.

To enable Thai NER detection (optional, downloads ~100 MB model on first use):

```bash
pip install pythainlp
```

## Detected PII Categories

Both scripts detect the following sensitive data using regex patterns:

| Category | Description |
|---|---|
| `name` | Thai names with honorifics: นาย, นาง, นางสาว, ดร., ศ., รศ., ผศ., นพ., พญ., ทนาย, คุณ, เด็กชาย, เด็กหญิง |
| `company` | Thai organizations: บริษัท...จำกัด, ห้างหุ้นส่วน, มูลนิธิ, สมาคม, การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย / กฟผ. |
| `thai_id` | Thai national ID (format: 0-0000-00000-00-0) |
| `passport` | Passport numbers (context-gated: requires "Passport" or หนังสือเดินทาง prefix) |
| `line_id` | LINE ID handles (Line ID / LINE ID / ไลน์) |
| `email` | Email addresses |
| `phone` | Thai phone numbers (+66 or 0x format) |
| `address` | Thai addresses with เลขที่, แขวง/ตำบล, เขต/อำเภอ, จังหวัด |
| `date` | ISO dates and Thai Buddhist-era dates |
| `timestamp` | ISO 8601 timestamps |
| `ssn` | US Social Security Numbers |
| `credit_card` | Credit/debit card numbers |
| `coordinates` | GPS coordinates |
| `doc_code` | Internal document codes |
| `project_code` | Project codes (NN/YYYY format) |

## Thai NER Detection (Optional)

When `THAI_NER=1` is set, both scripts additionally run PyThaiNLP NER (`thainer-corpus-v2-base-crf`) to detect Thai entities that may not carry honorific prefixes. NER detects:

| NER Tag | Mapped Category |
|---|---|
| PERSON | name |
| ORGANIZATION | company |
| LOCATION | location |
| DATE | date |
| TIME | timestamp |
| MONEY | money |
| PERCENT | percent |
| EMAIL | email |
| PHONE | phone |
| URL | url |

NER runs on raw string values in the data structure (not on serialized JSON text), avoiding false positives from JSON syntax tokens.

Enable NER:

```bash
# Windows
set THAI_NER=1
python anonymize_data_pro_v3.py path/to/input.json

# Linux / macOS
THAI_NER=1 python anonymize_data_pro_v3.py path/to/input.json
```

The NER model is cached in memory after the first call within a session.

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
- `New (NER)`: values detected by Thai NER (when `THAI_NER=1`)
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
