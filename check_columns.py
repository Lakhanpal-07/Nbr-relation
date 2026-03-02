import pandas as pd

print("=== YOUR EXCEL FILE STRUCTURE ===")
df = pd.read_excel('Sector-Addition-NBR-List.xlsx')

print("EXACT COLUMN NAMES:")
print(df.columns.tolist())
print("\nFIRST 5 ROWS:")
print(df.head())
print(f"\nTotal rows loaded: {len(df)}")
print("\n=== COPY COLUMN NAMES ABOVE ===")
