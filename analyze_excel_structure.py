#!/usr/bin/env python3
"""
Excel结构深度分析工具
专门用于分析BFG-CO2H-HEX.xlsx文件结构，识别数据提取问题
"""

import pandas as pd
import os
import json
from typing import Dict, List, Any, Optional
from datetime import datetime
import numpy as np

class ExcelStructureAnalyzer:
    """Excel文件结构深度分析器"""
    
    def __init__(self, excel_file: str):
        self.excel_file = excel_file
        self.analysis_results = {}
        self.all_worksheets = {}
        
    def analyze_complete_structure(self) -> Dict[str, Any]:
        """完整分析Excel文件结构"""
        print(f"🔍 开始分析Excel文件: {self.excel_file}")
        print("=" * 80)
        
        if not os.path.exists(self.excel_file):
            return {"error": f"文件不存在: {self.excel_file}"}
        
        results = {
            "file_info": self._get_file_info(),
            "worksheets": {},
            "summary": {},
            "data_patterns": {},
            "column_analysis": {},
            "recommendations": []
        }
        
        try:
            # 1. 获取所有工作表信息
            print("📋 分析工作表结构...")
            worksheets_info = self._analyze_all_worksheets()
            results["worksheets"] = worksheets_info
            
            # 2. 分析数据模式
            print("📊 分析数据模式...")
            data_patterns = self._analyze_data_patterns()
            results["data_patterns"] = data_patterns
            
            # 3. 分析列名结构
            print("🏷️ 分析列名结构...")
            column_analysis = self._analyze_column_structures()
            results["column_analysis"] = column_analysis
            
            # 4. 生成摘要
            print("📈 生成分析摘要...")
            summary = self._generate_summary()
            results["summary"] = summary
            
            # 5. 生成修复建议
            print("💡 生成修复建议...")
            recommendations = self._generate_recommendations()
            results["recommendations"] = recommendations
            
        except Exception as e:
            results["error"] = f"分析过程中出错: {str(e)}"
            print(f"❌ 分析失败: {str(e)}")
        
        self.analysis_results = results
        return results
    
    def _get_file_info(self) -> Dict[str, Any]:
        """获取文件基本信息"""
        stat = os.stat(self.excel_file)
        return {
            "file_path": self.excel_file,
            "file_size_bytes": stat.st_size,
            "file_size_mb": round(stat.st_size / 1024 / 1024, 2),
            "last_modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "analysis_time": datetime.now().isoformat()
        }
    
    def _analyze_all_worksheets(self) -> Dict[str, Any]:
        """分析所有工作表"""
        worksheets_info = {}
        
        try:
            # 获取所有工作表名称
            xl_file = pd.ExcelFile(self.excel_file, engine='openpyxl')
            sheet_names = xl_file.sheet_names
            
            print(f"   发现 {len(sheet_names)} 个工作表: {sheet_names}")
            
            for sheet_name in sheet_names:
                print(f"   分析工作表: {sheet_name}")
                
                try:
                    # 读取工作表数据
                    df = pd.read_excel(self.excel_file, sheet_name=sheet_name, engine='openpyxl')
                    self.all_worksheets[sheet_name] = df
                    
                    # 分析工作表结构
                    sheet_info = {
                        "shape": df.shape,
                        "columns": list(df.columns),
                        "column_count": len(df.columns),
                        "row_count": len(df),
                        "data_types": df.dtypes.astype(str).to_dict(),
                        "null_counts": df.isnull().sum().to_dict(),
                        "non_null_counts": df.count().to_dict(),
                        "sample_data": {},
                        "numeric_columns": [],
                        "text_columns": [],
                        "mixed_columns": []
                    }
                    
                    # 分析列类型
                    for col in df.columns:
                        col_data = df[col].dropna()
                        if len(col_data) > 0:
                            if pd.api.types.is_numeric_dtype(col_data):
                                sheet_info["numeric_columns"].append(col)
                            elif pd.api.types.is_string_dtype(col_data) or col_data.dtype == 'object':
                                # 检查是否为混合类型
                                numeric_count = sum(pd.to_numeric(col_data, errors='coerce').notna())
                                if numeric_count > 0 and numeric_count < len(col_data):
                                    sheet_info["mixed_columns"].append(col)
                                else:
                                    sheet_info["text_columns"].append(col)
                    
                    # 获取样本数据
                    if not df.empty:
                        sheet_info["sample_data"] = {
                            "first_row": df.iloc[0].to_dict() if len(df) > 0 else {},
                            "last_row": df.iloc[-1].to_dict() if len(df) > 0 else {},
                            "sample_rows": df.head(3).to_dict('records') if len(df) > 0 else []
                        }
                    
                    worksheets_info[sheet_name] = sheet_info
                    print(f"     ✅ {sheet_name}: {df.shape[0]} 行 × {df.shape[1]} 列")
                    
                except Exception as e:
                    worksheets_info[sheet_name] = {"error": f"读取失败: {str(e)}"}
                    print(f"     ❌ {sheet_name}: 读取失败 - {str(e)}")
            
        except Exception as e:
            return {"error": f"无法打开Excel文件: {str(e)}"}
        
        return worksheets_info
    
    def _analyze_data_patterns(self) -> Dict[str, Any]:
        """分析数据模式"""
        patterns = {
            "heat_exchanger_indicators": [],
            "temperature_patterns": [],
            "pressure_patterns": [],
            "flow_patterns": [],
            "duty_patterns": [],
            "area_patterns": [],
            "stream_patterns": [],
            "potential_hex_data": {}
        }
        
        # 定义搜索模式
        search_patterns = {
            "heat_exchanger_indicators": [
                'heat', 'exchanger', 'hex', 'hx', '换热', '换热器', 'cooler', 'heater', 
                'condenser', 'reboiler', 'boiler', '冷却器', '加热器', '冷凝器', '再沸器'
            ],
            "temperature_patterns": [
                'temp', 'temperature', '温度', 'inlet', 'outlet', 'in', 'out', 
                'hot', 'cold', '热', '冷', '进口', '出口', '入口', 'shell', 'tube'
            ],
            "pressure_patterns": [
                'press', 'pressure', '压力', '压强', 'bar', 'psi', 'pa', 'mpa', 'kpa'
            ],
            "flow_patterns": [
                'flow', 'mass', 'volume', 'molar', '流量', '质量', '体积', '摩尔', 
                'kg/h', 'kmol/h', 'm3/h', 'rate', '速率'
            ],
            "duty_patterns": [
                'duty', 'load', 'heat', 'power', '负荷', '热负荷', '功率', 'kw', 'mw', 
                'kj/h', 'mj/h', 'btu', 'kcal'
            ],
            "area_patterns": [
                'area', 'surface', '面积', '换热面积', 'm2', 'm²', 'ft2', 'ft²'
            ],
            "stream_patterns": [
                'stream', 'side', 'shell', 'tube', '流股', '壳程', '管程', 
                'hot side', 'cold side', 'process', 'utility'
            ]
        }
        
        # 搜索所有工作表中的模式
        for sheet_name, df in self.all_worksheets.items():
            if isinstance(df, pd.DataFrame):
                sheet_patterns = {}
                
                for pattern_type, keywords in search_patterns.items():
                    matching_columns = []
                    for col in df.columns:
                        col_str = str(col).lower()
                        for keyword in keywords:
                            if keyword.lower() in col_str:
                                matching_columns.append({
                                    "column": col,
                                    "keyword": keyword,
                                    "match_position": col_str.find(keyword.lower())
                                })
                                break
                    
                    if matching_columns:
                        patterns[pattern_type].extend(matching_columns)
                        sheet_patterns[pattern_type] = matching_columns
                
                # 评估该工作表是否包含热交换器数据
                hex_score = 0
                hex_indicators = []
                
                if sheet_patterns.get("heat_exchanger_indicators"):
                    hex_score += 3
                    hex_indicators.append("包含热交换器相关列名")
                
                if sheet_patterns.get("temperature_patterns"):
                    hex_score += 2
                    hex_indicators.append("包含温度相关列名")
                
                if sheet_patterns.get("duty_patterns"):
                    hex_score += 2
                    hex_indicators.append("包含热负荷相关列名")
                
                if sheet_patterns.get("area_patterns"):
                    hex_score += 2
                    hex_indicators.append("包含面积相关列名")
                
                if sheet_patterns.get("stream_patterns"):
                    hex_score += 1
                    hex_indicators.append("包含流股相关列名")
                
                patterns["potential_hex_data"][sheet_name] = {
                    "hex_score": hex_score,
                    "indicators": hex_indicators,
                    "patterns": sheet_patterns,
                    "likely_hex_sheet": hex_score >= 3
                }
        
        return patterns
    
    def _analyze_column_structures(self) -> Dict[str, Any]:
        """分析列名结构"""
        column_analysis = {
            "all_columns": [],
            "column_patterns": {},
            "naming_conventions": {},
            "suggested_mappings": {}
        }
        
        all_columns = []
        for sheet_name, df in self.all_worksheets.items():
            if isinstance(df, pd.DataFrame):
                for col in df.columns:
                    all_columns.append({
                        "sheet": sheet_name,
                        "column": str(col),
                        "column_lower": str(col).lower(),
                        "length": len(str(col)),
                        "has_chinese": any('\u4e00' <= char <= '\u9fff' for char in str(col)),
                        "has_underscore": '_' in str(col),
                        "has_space": ' ' in str(col),
                        "has_number": any(char.isdigit() for char in str(col))
                    })
        
        column_analysis["all_columns"] = all_columns
        
        # 分析命名约定
        naming_conventions = {
            "chinese_columns": len([c for c in all_columns if c["has_chinese"]]),
            "underscore_columns": len([c for c in all_columns if c["has_underscore"]]),
            "space_columns": len([c for c in all_columns if c["has_space"]]),
            "number_columns": len([c for c in all_columns if c["has_number"]]),
            "total_columns": len(all_columns),
            "unique_columns": len(set(c["column_lower"] for c in all_columns))
        }
        
        column_analysis["naming_conventions"] = naming_conventions
        
        # 生成建议的列名映射
        suggested_mappings = self._generate_column_mappings(all_columns)
        column_analysis["suggested_mappings"] = suggested_mappings
        
        return column_analysis
    
    def _generate_column_mappings(self, all_columns: List[Dict]) -> Dict[str, List[str]]:
        """生成建议的列名映射"""
        mappings = {
            "equipment_name": [],
            "duty": [],
            "area": [],
            "hot_stream_name": [],
            "cold_stream_name": [],
            "hot_inlet_temp": [],
            "hot_outlet_temp": [],
            "cold_inlet_temp": [],
            "cold_outlet_temp": [],
            "hot_flow": [],
            "cold_flow": [],
            "pressure": []
        }
        
        # 定义映射规则
        mapping_rules = {
            "equipment_name": ['name', 'id', 'tag', 'equipment', '设备', '名称', 'hex'],
            "duty": ['duty', 'load', 'heat', '负荷', '热负荷', 'kw', 'mw'],
            "area": ['area', '面积', 'm2', 'm²', 'surface'],
            "hot_stream_name": ['hot', 'shell', '热', '壳程', 'hot stream', 'hot side'],
            "cold_stream_name": ['cold', 'tube', '冷', '管程', 'cold stream', 'cold side'],
            "hot_inlet_temp": ['hot', 'inlet', 'in', '进口', '入口', 'shell', 'temp'],
            "hot_outlet_temp": ['hot', 'outlet', 'out', '出口', 'shell', 'temp'],
            "cold_inlet_temp": ['cold', 'inlet', 'in', '进口', '入口', 'tube', 'temp'],
            "cold_outlet_temp": ['cold', 'outlet', 'out', '出口', 'tube', 'temp'],
            "hot_flow": ['hot', 'flow', 'mass', '流量', 'shell', 'kg/h'],
            "cold_flow": ['cold', 'flow', 'mass', '流量', 'tube', 'kg/h'],
            "pressure": ['press', 'pressure', '压力', 'bar', 'psi']
        }
        
        # 为每个类别找到匹配的列
        for category, keywords in mapping_rules.items():
            matching_columns = []
            
            for col_info in all_columns:
                col_lower = col_info["column_lower"]
                
                # 计算匹配分数
                match_score = 0
                matched_keywords = []
                
                for keyword in keywords:
                    if keyword.lower() in col_lower:
                        match_score += 1
                        matched_keywords.append(keyword)
                
                if match_score > 0:
                    matching_columns.append({
                        "column": col_info["column"],
                        "sheet": col_info["sheet"],
                        "match_score": match_score,
                        "matched_keywords": matched_keywords
                    })
            
            # 按匹配分数排序
            matching_columns.sort(key=lambda x: x["match_score"], reverse=True)
            mappings[category] = matching_columns[:5]  # 保留前5个最佳匹配
        
        return mappings
    
    def _generate_summary(self) -> Dict[str, Any]:
        """生成分析摘要"""
        summary = {
            "total_worksheets": len(self.all_worksheets),
            "total_columns": 0,
            "total_rows": 0,
            "worksheets_with_data": 0,
            "likely_hex_worksheets": [],
            "data_quality_issues": [],
            "extraction_readiness": "unknown"
        }
        
        for sheet_name, df in self.all_worksheets.items():
            if isinstance(df, pd.DataFrame):
                summary["total_columns"] += len(df.columns)
                summary["total_rows"] += len(df)
                if not df.empty:
                    summary["worksheets_with_data"] += 1
        
        # 识别可能的热交换器工作表
        for sheet_name, info in self.analysis_results.get("data_patterns", {}).get("potential_hex_data", {}).items():
            if info.get("likely_hex_sheet", False):
                summary["likely_hex_worksheets"].append({
                    "sheet": sheet_name,
                    "score": info.get("hex_score", 0),
                    "indicators": info.get("indicators", [])
                })
        
        # 评估提取准备度
        if len(summary["likely_hex_worksheets"]) > 0:
            summary["extraction_readiness"] = "good"
        elif summary["worksheets_with_data"] > 0:
            summary["extraction_readiness"] = "needs_analysis"
        else:
            summary["extraction_readiness"] = "poor"
        
        return summary
    
    def _generate_recommendations(self) -> List[str]:
        """生成修复建议"""
        recommendations = []
        
        summary = self.analysis_results.get("summary", {})
        
        # 基于工作表数量的建议
        if summary.get("total_worksheets", 0) > 1:
            recommendations.append("建议修改代码支持多工作表读取，当前只读取第一个工作表")
        
        # 基于热交换器数据的建议
        likely_hex_sheets = summary.get("likely_hex_worksheets", [])
        if len(likely_hex_sheets) > 0:
            sheet_names = [sheet["sheet"] for sheet in likely_hex_sheets]
            recommendations.append(f"发现可能包含热交换器数据的工作表: {sheet_names}")
            recommendations.append("建议优先从这些工作表提取数据")
        
        # 基于列名映射的建议
        column_analysis = self.analysis_results.get("column_analysis", {})
        suggested_mappings = column_analysis.get("suggested_mappings", {})
        
        for category, mappings in suggested_mappings.items():
            if len(mappings) > 0:
                best_match = mappings[0]
                recommendations.append(f"建议将列 '{best_match['column']}' 映射为 {category}")
        
        # 基于命名约定的建议
        naming_conventions = column_analysis.get("naming_conventions", {})
        if naming_conventions.get("chinese_columns", 0) > 0:
            recommendations.append("发现中文列名，建议增加中文关键词匹配")
        
        if naming_conventions.get("space_columns", 0) > naming_conventions.get("underscore_columns", 0):
            recommendations.append("列名多使用空格分隔，建议调整匹配模式")
        
        return recommendations
    
    def print_analysis_report(self):
        """打印分析报告"""
        if not self.analysis_results:
            print("❌ 尚未执行分析")
            return
        
        print("\n" + "=" * 80)
        print("📊 EXCEL文件结构分析报告")
        print("=" * 80)
        
        # 文件信息
        file_info = self.analysis_results.get("file_info", {})
        print(f"\n📁 文件信息:")
        print(f"   文件路径: {file_info.get('file_path', 'N/A')}")
        print(f"   文件大小: {file_info.get('file_size_mb', 'N/A')} MB")
        print(f"   最后修改: {file_info.get('last_modified', 'N/A')}")
        
        # 摘要信息
        summary = self.analysis_results.get("summary", {})
        print(f"\n📈 摘要信息:")
        print(f"   工作表总数: {summary.get('total_worksheets', 0)}")
        print(f"   总列数: {summary.get('total_columns', 0)}")
        print(f"   总行数: {summary.get('total_rows', 0)}")
        print(f"   有数据的工作表: {summary.get('worksheets_with_data', 0)}")
        print(f"   提取准备度: {summary.get('extraction_readiness', 'unknown')}")
        
        # 工作表详情
        worksheets = self.analysis_results.get("worksheets", {})
        print(f"\n📋 工作表详情:")
        for sheet_name, sheet_info in worksheets.items():
            if "error" in sheet_info:
                print(f"   ❌ {sheet_name}: {sheet_info['error']}")
            else:
                print(f"   ✅ {sheet_name}: {sheet_info['row_count']} 行 × {sheet_info['column_count']} 列")
                print(f"      列名: {sheet_info['columns'][:5]}{'...' if len(sheet_info['columns']) > 5 else ''}")
        
        # 热交换器数据识别
        likely_hex_sheets = summary.get("likely_hex_worksheets", [])
        if likely_hex_sheets:
            print(f"\n🔥 可能的热交换器数据工作表:")
            for sheet_info in likely_hex_sheets:
                print(f"   🎯 {sheet_info['sheet']} (评分: {sheet_info['score']})")
                for indicator in sheet_info['indicators']:
                    print(f"      • {indicator}")
        
        # 修复建议
        recommendations = self.analysis_results.get("recommendations", [])
        if recommendations:
            print(f"\n💡 修复建议:")
            for i, rec in enumerate(recommendations, 1):
                print(f"   {i}. {rec}")
        
        print("\n" + "=" * 80)
    
    def save_analysis_to_json(self, output_file: str = None) -> str:
        """保存分析结果到JSON文件"""
        if not output_file:
            output_file = f"excel_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            # 转换numpy类型为Python原生类型
            def convert_numpy_types(obj):
                if isinstance(obj, np.integer):
                    return int(obj)
                elif isinstance(obj, np.floating):
                    return float(obj)
                elif isinstance(obj, np.ndarray):
                    return obj.tolist()
                elif isinstance(obj, dict):
                    return {key: convert_numpy_types(value) for key, value in obj.items()}
                elif isinstance(obj, list):
                    return [convert_numpy_types(item) for item in obj]
                else:
                    return obj
            
            converted_results = convert_numpy_types(self.analysis_results)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(converted_results, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 分析结果已保存到: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"❌ 保存失败: {str(e)}")
            return ""


def main():
    """主函数"""
    excel_file = "BFG-CO2H-HEX.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ 文件不存在: {excel_file}")
        return
    
    # 创建分析器
    analyzer = ExcelStructureAnalyzer(excel_file)
    
    # 执行完整分析
    results = analyzer.analyze_complete_structure()
    
    # 打印报告
    analyzer.print_analysis_report()
    
    # 保存结果
    analyzer.save_analysis_to_json()
    
    print(f"\n🏁 分析完成！")

if __name__ == "__main__":
    main()