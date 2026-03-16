import os
import pandas as pd

directory = r'C:\Users\LENOVO\.gemini\antigravity\scratch\sales_dashboard_exsis'

print(f"Scanning directory: {directory}")

for filename in os.listdir(directory):
    if filename.endswith(".xlsx") and not filename.startswith("~$"):
        filepath = os.path.join(directory, filename)
        print(f"\n--- File: {filename} ---")
        try:
            filename = "Rev Rifai Jan 2026 full.xlsx" 
            filepath = os.path.join(directory, filename)
            
            if not os.path.exists(filepath):
                print(f"File not found: {filepath}")
                exit()

            print(f"--- Inspecting {filename} ---")
            df = pd.read_excel(filepath, header=None)
            
            print("\n--- Last 20 Rows ---")
            print(df.tail(20))
            
            print("\n--- Rows with 'Total' in Col 4 --")
            for i, row in df.iterrows():
                val = str(row[4]).lower()
                if "total" in val or "grand" in val:
                    print(f"Row {i}: {row.tolist()}")

            exit()

            # Find header with 'Customer' and 'Jan'
            for i, row in df.iterrows():
                row_str = " ".join([str(val).lower() for val in row if pd.notna(val)])
                if "cust" in row_str or "shipper" in row_str:
                    print(f"\nPotential Header at Row {i}:")
                    print(row.tolist())
                    
                    # Print Next Row (Sub-headers)
                    if i + 1 < len(df):
                        print(f"\nPotential Sub-Header at Row {i+1}:")
                        print(df.iloc[i+1].tolist())

                if "jan" in row_str and "feb" in row_str:
                    print(f"\nPotential Month Row at Row {i}:")
                    print(row.tolist())

        except Exception as e:
            print(f"Error reading file: {e}")
