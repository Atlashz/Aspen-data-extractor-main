#!/usr/bin/env python3
"""
检查最终生成的经济分析报告内容
验证是否包含所有必要的成本项目
"""

import pandas as pd

def check_excel_output():
    """检查Excel文件内容"""
    file_name = "BFG_Economic_Analysis.xlsx"
    
    print(f"检查文件: {file_name}")
    print("=" * 50)
    
    try:
        # 检查CAPEX数据
        capex_df = pd.read_excel(file_name, sheet_name='CAPEX Breakdown')
        print("\nCAPEX分析:")
        
        # 寻找具体的成本项目
        found_items = []
        for i, row in capex_df.iterrows():
            col1 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            col2 = row.iloc[1] if len(row) > 1 else None
            
            if col1 and col2 is not None and pd.notna(col2):
                try:
                    value = float(col2)
                    if value > 0:
                        found_items.append((col1, value))
                except:
                    pass
        
        if found_items:
            for item, value in found_items:
                print(f"  + {item}: ${value:,.0f}")
        else:
            print("  WARNING: 未找到具体的CAPEX项目数据")
        
        # 检查OPEX数据
        opex_df = pd.read_excel(file_name, sheet_name='OPEX Analysis')
        print("\nOPEX分析:")
        
        opex_items = []
        for i, row in opex_df.iterrows():
            col1 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            col2 = row.iloc[1] if len(row) > 1 else None
            
            if col1 and col2 is not None and pd.notna(col2):
                try:
                    value = float(col2)
                    if value > 0:
                        opex_items.append((col1, value))
                except:
                    pass
        
        if opex_items:
            for item, value in opex_items:
                print(f"  + {item}: ${value:,.0f}")
        else:
            print("  WARNING: 未找到具体的OPEX项目数据")
        
        # 检查设备详情
        try:
            equipment_df = pd.read_excel(file_name, sheet_name='Equipment Details')
            print(f"\n设备详情: {len(equipment_df)} 行数据")
            
            # 寻找设备成本数据
            equipment_count = 0
            total_equipment_cost = 0
            
            for i, row in equipment_df.iterrows():
                # 检查是否包含成本信息
                for col in range(len(row)):
                    value = row.iloc[col]
                    if pd.notna(value) and isinstance(value, (int, float)) and value > 1000:  # 假设设备成本 > $1000
                        equipment_count += 1
                        total_equipment_cost += value
                        break
            
            if equipment_count > 0:
                print(f"  + 找到 {equipment_count} 个设备，总成本约: ${total_equipment_cost:,.0f}")
            else:
                print("  WARNING: 未找到设备成本明细")
                
        except Exception as e:
            print(f"  ERROR: 无法读取设备详情: {e}")
        
        # 检查财务分析
        try:
            financial_df = pd.read_excel(file_name, sheet_name='Financial Analysis')
            print(f"\n财务分析:")
            
            # 寻找关键财务指标
            key_metrics = ['NPV', 'IRR', 'Payback', 'CAPEX', 'OPEX']
            
            for i, row in financial_df.iterrows():
                col1 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                col2 = row.iloc[1] if len(row) > 1 else None
                
                for metric in key_metrics:
                    if metric.lower() in col1.lower() and col2 is not None and pd.notna(col2):
                        try:
                            value = float(col2)
                            if 'NPV' in metric or 'CAPEX' in metric or 'OPEX' in metric:
                                print(f"  + {col1}: ${value:,.0f}")
                            else:
                                print(f"  + {col1}: {value}")
                        except:
                            print(f"  + {col1}: {col2}")
        
        except Exception as e:
            print(f"  ERROR: 无法读取财务分析: {e}")
        
        print("\n" + "=" * 50)
        print("SUCCESS: 文件检查完成!")
        
    except Exception as e:
        print(f"ERROR: 检查文件时出错: {e}")

if __name__ == "__main__":
    check_excel_output()