#!/usr/bin/env python3
import pandas as pd

# 读取Excel文件查看列名
excel_file = "BFG-CO2H-HEX.xlsx"
df = pd.read_excel(excel_file, sheet_name='Sheet1')

print("Excel columns:")
for i, col in enumerate(df.columns):
    print(f"  {i}: '{col}'")

print(f"\nTotal columns: {len(df.columns)}")
print(f"Total rows: {len(df)}")

# 查看前几行数据
print("\nFirst few rows of hot stream and cold stream columns:")
for i, row in df.head(3).iterrows():
    if 'hot stream' in df.columns and 'Cold stream' in df.columns:
        print(f"Row {i+1}: hot='{row['hot stream']}', cold='{row['Cold stream']}'")
    else:
        print("Hot stream or Cold stream columns not found")
