import pandas as pd

def check_excel_data():
    """Check BFG_Economic_Analysis.xlsx for data content"""
    xl = pd.ExcelFile('BFG_Economic_Analysis.xlsx')
    
    print("Checking BFG_Economic_Analysis.xlsx...")
    print(f"Available sheets: {xl.sheet_names}")
    print()
    
    total_cells = 0
    for sheet in xl.sheet_names:
        df = pd.read_excel('BFG_Economic_Analysis.xlsx', sheet_name=sheet)
        non_empty_cells = df.count().sum()
        total_cells += non_empty_cells
        
        print(f"=== {sheet} ===")
        print(f"Shape: {df.shape}")
        print(f"Non-empty cells: {non_empty_cells}")
        
        if non_empty_cells > 0:
            print("Sample data:")
            print(df.head(3))
        else:
            print("Sheet is empty")
        print()
    
    print(f"Total non-empty cells across all sheets: {total_cells}")
    return total_cells > 0

if __name__ == "__main__":
    has_data = check_excel_data()
    print(f"File has data: {has_data}")