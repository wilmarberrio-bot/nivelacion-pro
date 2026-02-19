import glob
import os
import openpyxl

def inspect_latest_excel():
    files = glob.glob("*.xlsx")
    if not files:
        print("No xlsx files found.")
        return

    files.sort(key=os.path.getmtime, reverse=True)
    latest = files[0]
    print(f"Inspecting latest file: {latest}")
    
    try:
        wb = openpyxl.load_workbook(latest, read_only=True, data_only=True)
        sheet = wb.active
        headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        print("Headers found:")
        for h in headers:
            print(f"- {h}")
        wb.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_latest_excel()
