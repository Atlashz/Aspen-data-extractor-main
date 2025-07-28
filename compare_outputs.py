#!/usr/bin/env python3
"""
对比主程序输出与测试文件输出
验证修复效果
"""

import pandas as pd

def compare_files():
    """对比两个Excel文件的内容"""
    main_file = "BFG_Economic_Analysis.xlsx"
    test_file = "test_economics_report.xlsx"
    
    print("=== 经济分析文件对比 ===")
    print(f"主程序输出: {main_file}")
    print(f"测试文件:   {test_file}")
    print("=" * 50)
    
    try:
        # 读取CAPEX数据
        main_capex = pd.read_excel(main_file, sheet_name='CAPEX Breakdown')
        test_capex = pd.read_excel(test_file, sheet_name='CAPEX Breakdown')
        
        print(f"CAPEX Breakdown:")
        print(f"  主程序: {main_capex.shape[0]} 行, {main_capex.count().sum()} 个非空单元格")
        print(f"  测试文件: {test_capex.shape[0]} 行, {test_capex.count().sum()} 个非空单元格")
        
        # 读取OPEX数据
        main_opex = pd.read_excel(main_file, sheet_name='OPEX Analysis')
        test_opex = pd.read_excel(test_file, sheet_name='OPEX Analysis')
        
        print(f"\nOPEX Analysis:")
        print(f"  主程序: {main_opex.shape[0]} 行, {main_opex.count().sum()} 个非空单元格")
        print(f"  测试文件: {test_opex.shape[0]} 行, {test_opex.count().sum()} 个非空单元格")
        
        # 提取关键数值
        def extract_key_values(df):
            values = {}
            for i, row in df.iterrows():
                col1 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if len(row) > 1 and pd.notna(row.iloc[1]):
                    try:
                        value = float(row.iloc[1])
                        if 'Total CAPEX' in col1:
                            values['CAPEX'] = value
                        elif 'Total Annual OPEX' in col1:
                            values['OPEX'] = value
                        elif 'NPV' in col1:
                            values['NPV'] = value
                    except:
                        pass
            return values
        
        # 从财务分析中提取值
        try:
            main_financial = pd.read_excel(main_file, sheet_name='Financial Analysis')
            test_financial = pd.read_excel(test_file, sheet_name='Financial Analysis')
            
            main_values = extract_key_values(main_capex) 
            main_values.update(extract_key_values(main_opex))
            main_values.update(extract_key_values(main_financial))
            
            test_values = extract_key_values(test_capex)
            test_values.update(extract_key_values(test_opex))
            test_values.update(extract_key_values(test_financial))
            
            print(f"\n关键指标对比:")
            metrics = ['CAPEX', 'OPEX', 'NPV']
            
            for metric in metrics:
                main_val = main_values.get(metric, 0)
                test_val = test_values.get(metric, 0)
                
                if main_val > 0 and test_val > 0:
                    ratio = main_val / test_val
                    print(f"  {metric:5}: 主程序=${main_val:12,.0f} | 测试=${test_val:12,.0f} | 比率={ratio:.2f}")
                elif main_val > 0:
                    print(f"  {metric:5}: 主程序=${main_val:12,.0f} | 测试=无数据")
                elif test_val > 0:
                    print(f"  {metric:5}: 主程序=无数据 | 测试=${test_val:12,.0f}")
                else:
                    print(f"  {metric:5}: 都无数据")
        
        except Exception as e:
            print(f"  WARNING: 无法对比财务指标: {e}")
        
        print("\n" + "=" * 50)
        
        # 判断修复是否成功
        main_total_cells = main_capex.count().sum() + main_opex.count().sum()
        test_total_cells = test_capex.count().sum() + test_opex.count().sum()
        
        if main_total_cells > test_total_cells:
            print("SUCCESS: 主程序输出的数据比测试文件更完整!")
        elif main_total_cells == test_total_cells:
            print("INFO: 主程序和测试文件的数据量相当")
        else:
            print("WARNING: 主程序输出的数据可能还不够完整")
            
        print(f"数据完整性: 主程序={main_total_cells}单元格, 测试文件={test_total_cells}单元格")
        
    except Exception as e:
        print(f"ERROR: 对比过程出错: {e}")

if __name__ == "__main__":
    compare_files()