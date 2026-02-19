import glob
import os
import openpyxl

def inspect_data():
    files = glob.glob("*.xlsx")
    files.sort(key=os.path.getmtime, reverse=True)
    latest = files[0]
    
    wb = openpyxl.load_workbook(latest, read_only=True, data_only=True)
    sheet = wb.active
    
    rows = list(sheet.iter_rows(min_row=1, max_row=5, values_only=True))
    headers = rows[0]
    
    for i, row in enumerate(rows[1:], 1):
        print(f"\nRow {i}:")
        for h, v in zip(headers, row):
            print(f"  {h}: {v}")
    
    wb.close()

if __name__ == "__main__":
    inspect_data()
