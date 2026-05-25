# -*- coding: utf-8 -*-
"""
Created on Sun May  4 11:27:40 2025

@author: EGAT
"""

import pandas as pd
import json
import numpy as np

# 📌 Set path to your Excel file
excel_path = r"C:/Users/EGAT/Documents/Claude_AIEE/20260515 new_ground_truth_v3/วสท56.xlsx"

# Load all sheets with NO header (every row treated as data)
sheets_dict = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl', header=None)

# Dictionary to hold all sheet data
combined_data = {}

# Loop through each sheet
for sheet_name, df in sheets_dict.items():
    # Drop completely empty rows
    df = df.dropna(how='all').reset_index(drop=True)

    # Replace NaN with None to make it JSON-friendly
    df = df.replace({np.nan: None})

    # If only one column, store as a list
    if df.shape[1] == 1:
        combined_data[sheet_name] = df.iloc[:, 0].tolist()
    else:
        # Otherwise, store as a list of rows (lists)
        combined_data[sheet_name] = df.values.tolist()

# 📁 Output path
output_path = r"C:/Users/EGAT/Documents/Claude_AIEE/20260515 new_ground_truth_v3/วสท56.json"

# Save JSON
with open(output_path, 'w', encoding='utf-8') as f:
    # ❌ โค้ดเดิมที่ทำให้เกิด error หากเจอ datetime object
    # json.dump(combined_data, f, indent=4, ensure_ascii=False)
    
    # ✅ โค้ดใหม่: เพิ่ม default=str เพื่อแปลง datetime ให้อัตโนมัติ
    json.dump(combined_data, f, indent=4, ensure_ascii=False, default=str)

print(f"✅ Done! JSON saved to: {output_path}")
