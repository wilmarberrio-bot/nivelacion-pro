
import openpyxl
import os

files = [
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\pre_ruta_2_0_2026-02-17T10_51_45.802611833-05_00.xlsx",
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\query_result_2026-02-17T10_27_43.628776613-05_00.xlsx",
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\seguimiento_de_marcaciones_2026-02-17T10_47_41.588123824-05_00.xlsx"
]

def get_column_values(sheet, col_idx, max_rows=10):
    val_list = []
    for row in sheet.iter_rows(min_row=2, max_row=max_rows, min_col=col_idx, max_col=col_idx, values_only=True):
         if row[0] is not None:
             val_list.append(row[0])
    return val_list

for f in files:
    print(f"\n--- Analysis of {os.path.basename(f)} ---")
    try:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        sheet = wb.active
        
        # Get headers from first row
        headers = []
        for cell in sheet[1]:
            headers.append(cell.value)
        
        print("Columns:", headers)
        
        # Check for potential status columns or type columns
        for idx, h in enumerate(headers, start=1):
            if h and ('estado' in str(h).lower() or 'status' in str(h).lower() or 'tipo' in str(h).lower()):
                print(f"Potential interesting column '{h}':")
                vals = get_column_values(sheet, idx)
                print(f"  Sample values: {vals}")
                
    except Exception as e:
        print(f"Error reading file: {e}")
