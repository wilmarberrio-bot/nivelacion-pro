import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

INPUT_FILE = r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\query_result_2026-02-17T10_27_43.628776613-05_00.xlsx"

try:
    print(f"Reading {INPUT_FILE}...")
    wb = openpyxl.load_workbook(INPUT_FILE, read_only=True, data_only=True)
    sheet = wb.active
    
    print("\nFirst 10 rows:")
    for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
        print(row)

except Exception as e:
    print(f"Error: {e}")
