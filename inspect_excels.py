
import pandas as pd
import os

files = [
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\pre_ruta_2_0_2026-02-17T10_51_45.802611833-05_00.xlsx",
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\query_result_2026-02-17T10_27_43.628776613-05_00.xlsx",
    r"c:\Users\Usuario\OneDrive\Documentos\nivelacion\seguimiento_de_marcaciones_2026-02-17T10_47_41.588123824-05_00.xlsx"
]

for f in files:
    print(f"\n--- Analysis of {os.path.basename(f)} ---")
    try:
        df = pd.read_excel(f, nrows=5)
        print("Columns:", list(df.columns))
        print("First row values (sample):")
        print(df.iloc[0].to_dict() if not df.empty else "Empty DataFrame")
        
        # Check for potential status columns
        potential_status_cols = [c for c in df.columns if 'estado' in c.lower() or 'status' in c.lower()]
        if potential_status_cols:
            print(f"Potential Status Columns: {potential_status_cols}")
            # Load full column to check unique values
            full_df = pd.read_excel(f, usecols=potential_status_cols)
            for col in potential_status_cols:
                print(f"Unique values in '{col}': {full_df[col].unique()[:10]}") # Show first 10 unique
                
    except Exception as e:
        print(f"Error reading file: {e}")
