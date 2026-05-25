# 🛡️ Advanced Privacy Anonymizer Pro

A professional-grade tool for detecting and anonymizing sensitive information across multiple file formats. This tool combines persistent "accumulated" knowledge with dynamic PII (Personally Identifiable Information) detection.

## 🚀 Key Features

*   **Multi-Format Support:** Process `.json`, `.pdf`, `.docx`, and `.txt` files.
*   **Persistent Memory:** Maintains a `sensitive_data_list.json` that "remembers" approved sensitive keywords across sessions.
*   **Dynamic Detection:** Uses Regex to automatically identify SSNs, Credit Cards, Emails, Coordinates, and Project Codes.
*   **Human-in-the-Loop:** A clear review step allows you to approve or reject specific detections before any changes are applied.
*   **Consistent Anonymization:** Uses `Faker` with a session cache to ensure that the same sensitive value (e.g., "EGAT") always receives the same fake replacement within a single file.
*   **Detailed Logging:** Tracks every replacement made in `anonymization_pro_log.txt`.

---

## 🛠️ Installation

Ensure you have the required libraries installed:

```bash
pip install faker PyPDF2 python-docx
```

---

## 📖 Step-by-Step Usage

### Step 1: Prepare your data
Place the file you want to anonymize in the project directory.

### Step 2: Run the script
Execute the script via terminal, providing the path to your file:

```bash
python anonymize_data_pro.py path/to/your/file.json
```

### Step 3: Review Detections
The script will scan the file and display a table of results:
*   **Accumulated:** Items already known from your `sensitive_data_list.json`.
*   **New (Regex):** Freshly detected patterns like emails or credit card numbers.

### Step 4: Approve Replacements
You will be prompted to select which items to anonymize:
*   Type `y` to approve **all** detected items.
*   Type `n` to cancel the operation.
*   Type specific IDs (e.g., `1,3,5`) to approve only those items.

### Step 5: Check Outputs
*   **Anonymized File:** A new file will be created with the suffix `_pro_anonymized` (e.g., `data_pro_anonymized.json`).
*   **Updated List:** Any **newly approved** items are automatically added to `sensitive_data_list.json` for future sessions.
*   **Session Log:** Review `anonymization_pro_log.txt` to see exactly what was replaced and with what value.

---

## 📂 Project Structure

*   **`anonymize_data_pro.py`**: The core execution engine.
*   **`sensitive_data_list.json`**: Your persistent database of sensitive keywords. You can manually edit this file to add or remove patterns.
*   **`anonymization_pro_log.txt`**: A historical record of all anonymization actions.

---

## 💡 Pro Tips

*   **Manual List Updates:** You can manually add entries to `sensitive_data_list.json` using this format:
    ```json
    {
        "pattern": "Secret Project X",
        "category": "project",
        "sensitive": false
    }
    ```
    *(Set `"sensitive": true` for case-sensitive matching).*
*   **PDF Output:** When processing a `.pdf`, the tool extracts the text and saves the anonymized result as a `.txt` file to preserve data integrity.
*   **Recursive JSON:** For `.json` files, the tool recursively scans all values while keeping your keys (structure) intact.
