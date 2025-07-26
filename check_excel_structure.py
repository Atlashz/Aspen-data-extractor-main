#!/usr/bin/env python3
"""
Simple script to check BFG-CO2H-HEX.xlsx structure and validate our enhancements
"""

try:
    import pandas as pd
    import os
    
    excel_file = "BFG-CO2H-HEX.xlsx"
    
    if os.path.exists(excel_file):
        print(f"Excel file found: {excel_file}")
        
        # Try to read the Excel file
        try:
            data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl')
            print(f"Successfully loaded Excel data: {data.shape[0]} rows, {data.shape[1]} columns")
            print(f"Columns: {list(data.columns)}")
            
            # Check for hot/cold stream related columns
            columns = [col.lower() for col in data.columns]
            print("\nAnalyzing columns for hot/cold stream data:")
            
            hot_related = [col for col in data.columns if any(phrase in col.lower() for phrase in ['hot', 'shell', 'h_'])]
            cold_related = [col for col in data.columns if any(phrase in col.lower() for phrase in ['cold', 'tube', 'c_'])]
            temp_related = [col for col in data.columns if any(phrase in col.lower() for phrase in ['temp', 'temperature', 'in', 'out'])]
            flow_related = [col for col in data.columns if any(phrase in col.lower() for phrase in ['flow', 'mass', 'molar'])]
            
            print(f"Hot-related columns: {hot_related}")
            print(f"Cold-related columns: {cold_related}")
            print(f"Temperature-related columns: {temp_related}")
            print(f"Flow-related columns: {flow_related}")
            
            # Show first few rows
            print(f"\nFirst 3 rows of data:")
            print(data.head(3).to_string())
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
    else:
        print(f"Excel file not found: {excel_file}")
        
except ImportError as e:
    print(f"Missing required libraries: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")