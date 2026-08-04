import pandas as pd
import sys

# Usage: python3 delete_column.py "Column Name"
if len(sys.argv) < 2:
    print("Usage: python3 delete_column.py 'Column Name'")
    print("\nAvailable columns:")
    df = pd.read_csv('Liquidity_Monthly_Calendar_2010_2026.csv')
    print([c for c in df.columns])
    sys.exit(1)

col = sys.argv[1]
df = pd.read_csv('Liquidity_Monthly_Calendar_2010_2026.csv')

if col not in df.columns:
    print(f"Column '{col}' not found. Available columns:")
    print([c for c in df.columns])
    sys.exit(1)

df = df.drop(columns=[col])
df.to_csv('Liquidity_Monthly_Calendar_2010_2026.csv', index=False)
print(f"Deleted '{col}'. Remaining columns: {df.columns.tolist()}")