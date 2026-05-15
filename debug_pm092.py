import pandas as pd
import os

path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'PM_092.xlsx')
df = pd.read_excel(path, header=None)

print("=== ROWS 216-267 (VTAD2606002 + VTAD2605001 lan 2) ===")
for i in range(216, min(267, len(df))):
    row = df.iloc[i]
    vals = {k: v for k, v in row.items() if pd.notna(v)}
    if vals:
        print(f"Row {i:3d}: {vals}")
