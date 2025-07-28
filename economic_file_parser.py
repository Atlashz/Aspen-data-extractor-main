#!/usr/bin/env python3
"""
Aspen Economic File Parser

解析Aspen Plus经济分析文件，包括：
- IZP文件 (Aspen Icarus Cost Estimator项目文件)
- SZP文件 (Aspen经济分析数据文件)
- 二进制成本数据提取和结构化

Author: TEA Analysis Framework
Date: 2025-07-27
Version: 1.0
"""

import os
import sys
import struct
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple, Union
from datetime import datetime
import zipfile
import tempfile

# Import local data structures
from data_interfaces import (
    CostItem, CapexData, OpexData, FinancialParameters, 
    EconomicAnalysisResults, CostCategory, CurrencyType, CostBasis
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class EconomicFileParser:
    """
    Aspen经济文件解析器主类
    
    支持解析多种Aspen经济分析文件格式，提取成本数据、
    计算参数和经济模型配置信息。
    """
    
    def __init__(self):
        self.supported_formats = ['.izp', '.szp']
        self.extracted_data = {}
        self.parsing_errors = []
        self.warnings = []
    
    def parse_file(self, file_path: str) -> EconomicAnalysisResults:
        """
        解析经济文件并返回结构化数据
        
        Args:
            file_path: 经济文件路径
            
        Returns:
            EconomicAnalysisResults: 解析后的经济数据
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"Economic file not found: {file_path}")
        
        file_ext = file_path.suffix.lower()
        
        if file_ext not in self.supported_formats:
            raise ValueError(f"Unsupported file format: {file_ext}")
        
        logger.info(f"Parsing economic file: {file_path}")
        
        try:
            if file_ext == '.izp':
                return self._parse_izp_file(file_path)
            elif file_ext == '.szp':
                return self._parse_szp_file(file_path)
        except Exception as e:
            logger.error(f"Error parsing {file_path}: {str(e)}")
            self.parsing_errors.append(str(e))
            raise
    
    def _parse_izp_file(self, file_path: Path) -> EconomicAnalysisResults:
        """
        解析IZP文件 (Aspen Icarus Cost Estimator项目文件)
        
        IZP文件是压缩的二进制格式，包含项目成本数据和计算参数
        """
        logger.info(f"Parsing IZP file: {file_path}")
        
        # 创建经济分析结果容器
        results = EconomicAnalysisResults(
            project_name=file_path.stem,
            timestamp=datetime.now(),
            analysis_version="1.0"
        )
        
        try:
            # 尝试作为ZIP文件解析
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                file_list = zip_file.namelist()
                logger.info(f"IZP contains {len(file_list)} files: {file_list}")
                
                # 查找关键文件
                cost_files = [f for f in file_list if 'cost' in f.lower() or 'econ' in f.lower()]
                data_files = [f for f in file_list if '.dat' in f.lower() or '.txt' in f.lower()]
                
                # 解析成本数据文件
                for cost_file in cost_files:
                    try:
                        with zip_file.open(cost_file) as f:
                            content = f.read()
                            self._extract_cost_data_from_binary(content, results)
                    except Exception as e:
                        logger.warning(f"Could not parse {cost_file}: {str(e)}")
                
                # 解析数据文件
                for data_file in data_files:
                    try:
                        with zip_file.open(data_file) as f:
                            content = f.read()
                            if self._is_text_content(content):
                                text_content = content.decode('utf-8', errors='ignore')
                                self._extract_text_based_data(text_content, results)
                    except Exception as e:
                        logger.warning(f"Could not parse {data_file}: {str(e)}")
                        
        except zipfile.BadZipFile:
            # 如果不是ZIP文件，尝试直接解析二进制内容
            logger.info("IZP is not a ZIP file, attempting binary parsing")
            with open(file_path, 'rb') as f:
                content = f.read()
                self._extract_cost_data_from_binary(content, results)
        
        # 添加数据源信息
        results.data_sources.append(f"IZP file: {file_path}")
        results.estimation_methods.append("Aspen Icarus Cost Estimator")
        
        return results
    
    def _parse_szp_file(self, file_path: Path) -> EconomicAnalysisResults:
        """
        解析SZP文件 (Aspen经济分析数据文件)
        
        SZP文件包含序列化的经济分析数据和计算结果
        """
        logger.info(f"Parsing SZP file: {file_path}")
        
        results = EconomicAnalysisResults(
            project_name=file_path.stem,
            timestamp=datetime.now(),
            analysis_version="1.0"
        )
        
        with open(file_path, 'rb') as f:
            content = f.read()
            
            # 解析文件头信息
            header_info = self._parse_szp_header(content[:1024])
            
            # 提取经济数据
            self._extract_szp_economic_data(content, results, header_info)
        
        # 添加数据源信息
        results.data_sources.append(f"SZP file: {file_path}")
        results.estimation_methods.append("Aspen Process Economic Analyzer")
        
        return results
    
    def _extract_cost_data_from_binary(self, content: bytes, results: EconomicAnalysisResults):
        """
        从二进制内容中提取成本数据
        
        Args:
            content: 二进制文件内容
            results: 经济分析结果容器
        """
        logger.info("Extracting cost data from binary content")
        
        # 查找可能的数字模式和字符串
        cost_keywords = [
            b'CAPEX', b'OPEX', b'EQUIPMENT', b'INSTALLATION', 
            b'MATERIAL', b'LABOR', b'UTILITY', b'MAINTENANCE',
            b'Cost', b'cost', b'Price', b'price', b'Total', b'total'
        ]
        
        # 搜索成本相关的文本和数据
        text_chunks = self._extract_text_chunks(content)
        numeric_data = self._extract_numeric_data(content)
        
        # 初始化CAPEX和OPEX数据
        capex_data = CapexData(project_name=results.project_name)
        opex_data = OpexData(project_name=results.project_name)
        
        # 解析提取的数据
        equipment_costs = self._identify_equipment_costs(text_chunks, numeric_data)
        utility_costs = self._identify_utility_costs(text_chunks, numeric_data)
        
        # 添加成本项目
        for name, cost in equipment_costs.items():
            cost_item = CostItem(
                name=name,
                category=CostCategory.EQUIPMENT,
                base_cost=cost,
                currency=CurrencyType.USD,
                estimation_method="Binary extraction"
            )
            capex_data.add_cost_item(cost_item)
        
        for name, cost in utility_costs.items():
            cost_item = CostItem(
                name=name,
                category=CostCategory.UTILITIES,
                base_cost=cost,
                currency=CurrencyType.USD,
                estimation_method="Binary extraction"
            )
            opex_data.add_opex_item(cost_item)
        
        # 计算总成本
        capex_data.calculate_total_capex()
        opex_data.calculate_annual_opex(capex_data.total_capex)
        
        results.capex_data = capex_data
        results.opex_data = opex_data
        results.total_capex = capex_data.total_capex
        results.annual_opex = opex_data.annual_opex
    
    def _parse_szp_header(self, header_bytes: bytes) -> Dict[str, Any]:
        """
        解析SZP文件头信息
        
        Args:
            header_bytes: 文件头字节数据
            
        Returns:
            Dict containing header information
        """
        header_info = {
            'file_version': None,
            'creation_date': None,
            'project_name': None,
            'currency': None,
            'data_sections': []
        }
        
        try:
            # 查找版本信息
            if b'VERSION' in header_bytes:
                version_pos = header_bytes.find(b'VERSION')
                version_data = header_bytes[version_pos:version_pos+20]
                header_info['file_version'] = self._extract_version_number(version_data)
            
            # 查找项目名称
            text_parts = self._extract_text_chunks(header_bytes)
            if text_parts:
                header_info['project_name'] = text_parts[0][:50]  # 限制长度
            
            logger.info(f"SZP header info: {header_info}")
            
        except Exception as e:
            logger.warning(f"Error parsing SZP header: {str(e)}")
        
        return header_info
    
    def _extract_szp_economic_data(self, content: bytes, results: EconomicAnalysisResults, 
                                  header_info: Dict[str, Any]):
        """
        从SZP内容中提取经济数据
        
        Args:
            content: 完整文件内容
            results: 经济分析结果容器
            header_info: 文件头信息
        """
        logger.info("Extracting economic data from SZP content")
        
        # 查找数据段
        sections = self._identify_data_sections(content)
        
        # 解析各个数据段
        for section_name, section_data in sections.items():
            if 'cost' in section_name.lower():
                self._parse_cost_section(section_data, results)
            elif 'financial' in section_name.lower():
                self._parse_financial_section(section_data, results)
            elif 'equipment' in section_name.lower():
                self._parse_equipment_section(section_data, results)
    
    def _extract_text_chunks(self, content: bytes) -> List[str]:
        """
        从二进制内容中提取文本片段
        
        Args:
            content: 二进制内容
            
        Returns:
            List of text strings found in content
        """
        text_chunks = []
        
        # 尝试不同编码
        encodings = ['utf-8', 'latin-1', 'cp1252', 'ascii']
        
        for encoding in encodings:
            try:
                # 查找连续的可打印字符
                decoded = content.decode(encoding, errors='ignore')
                
                # 提取长度大于3的字符串片段
                words = []
                current_word = ""
                
                for char in decoded:
                    if char.isprintable() and not char.isspace():
                        current_word += char
                    else:
                        if len(current_word) > 3:
                            words.append(current_word)
                        current_word = ""
                
                text_chunks.extend(words)
                break  # 成功解码，退出循环
                
            except Exception:
                continue
        
        # 去重并排序
        unique_chunks = list(set(text_chunks))
        return unique_chunks[:100]  # 限制数量
    
    def _extract_numeric_data(self, content: bytes) -> List[float]:
        """
        从二进制内容中提取数值数据
        
        Args:
            content: 二进制内容
            
        Returns:
            List of numeric values found
        """
        numbers = []
        
        # 尝试解析4字节和8字节浮点数
        for i in range(0, len(content) - 8, 4):
            try:
                # 4字节浮点数
                num_float = struct.unpack('<f', content[i:i+4])[0]
                if self._is_reasonable_cost_value(num_float):
                    numbers.append(float(num_float))
                
                # 8字节浮点数
                if i + 8 <= len(content):
                    num_double = struct.unpack('<d', content[i:i+8])[0]
                    if self._is_reasonable_cost_value(num_double):
                        numbers.append(float(num_double))
                        
            except (struct.error, OverflowError):
                continue
        
        # 去重并排序
        unique_numbers = sorted(list(set(numbers)))
        return unique_numbers[:500]  # 限制数量
    
    def _is_reasonable_cost_value(self, value: float) -> bool:
        """
        判断数值是否为合理的成本值
        
        Args:
            value: 数值
            
        Returns:
            True if value seems like a reasonable cost
        """
        if not isinstance(value, (int, float)):
            return False
        
        # 排除无穷大、NaN等特殊值
        if not (-1e10 < value < 1e10):
            return False
        
        # 排除太小的值（可能是系数或比例）
        if abs(value) < 1.0:
            return False
        
        # 排除明显不合理的大值
        if abs(value) > 1e9:
            return False
        
        return True
    
    def _identify_equipment_costs(self, text_chunks: List[str], 
                                numeric_data: List[float]) -> Dict[str, float]:
        """
        识别设备成本数据
        
        Args:
            text_chunks: 文本片段
            numeric_data: 数值数据
            
        Returns:
            Dictionary of equipment names and costs
        """
        equipment_costs = {}
        
        # 设备类型关键词
        equipment_keywords = [
            'REACTOR', 'PUMP', 'COMPRESSOR', 'HEAT_EXCHANGER', 'COLUMN', 
            'SEPARATOR', 'TANK', 'VESSEL', 'TOWER', 'DISTILLATION'
        ]
        
        # 查找设备相关的文本和对应的数值
        for i, text in enumerate(text_chunks):
            text_upper = text.upper()
            for keyword in equipment_keywords:
                if keyword in text_upper:
                    # 查找附近的数值作为成本
                    if i < len(numeric_data):
                        cost = numeric_data[min(i, len(numeric_data)-1)]
                        equipment_name = f"{keyword}_{i+1}"
                        equipment_costs[equipment_name] = cost
        
        # 如果没找到设备，使用较大的数值作为设备成本
        if not equipment_costs and numeric_data:
            large_values = [v for v in numeric_data if v > 10000]
            for i, cost in enumerate(large_values[:10]):  # 取前10个大值
                equipment_costs[f"EQUIPMENT_{i+1}"] = cost
        
        return equipment_costs
    
    def _identify_utility_costs(self, text_chunks: List[str], 
                               numeric_data: List[float]) -> Dict[str, float]:
        """
        识别公用工程成本数据
        
        Args:
            text_chunks: 文本片段
            numeric_data: 数值数据
            
        Returns:
            Dictionary of utility names and costs
        """
        utility_costs = {}
        
        # 公用工程关键词
        utility_keywords = [
            'STEAM', 'COOLING_WATER', 'ELECTRICITY', 'FUEL_GAS', 
            'COMPRESSED_AIR', 'NITROGEN', 'WATER', 'POWER'
        ]
        
        # 查找公用工程相关的文本和数值
        for i, text in enumerate(text_chunks):
            text_upper = text.upper()
            for keyword in utility_keywords:
                if keyword in text_upper or keyword.replace('_', '') in text_upper:
                    # 查找附近的较小数值作为年消耗成本
                    nearby_values = numeric_data[max(0, i-5):i+5]
                    small_values = [v for v in nearby_values if 100 < v < 50000]
                    if small_values:
                        utility_costs[keyword] = small_values[0]
        
        return utility_costs
    
    def _identify_data_sections(self, content: bytes) -> Dict[str, bytes]:
        """
        识别数据文件中的不同数据段
        
        Args:
            content: 文件内容
            
        Returns:
            Dictionary of section names and their binary content
        """
        sections = {}
        
        # 查找常见的数据段标识符
        section_markers = [
            b'COST_DATA', b'EQUIPMENT_LIST', b'FINANCIAL_PARAMS',
            b'CAPEX', b'OPEX', b'ECONOMICS', b'SUMMARY'
        ]
        
        current_pos = 0
        for marker in section_markers:
            marker_pos = content.find(marker, current_pos)
            if marker_pos != -1:
                # 查找下一个标识符确定段的结束位置
                next_pos = len(content)
                for next_marker in section_markers:
                    next_marker_pos = content.find(next_marker, marker_pos + len(marker))
                    if next_marker_pos != -1:
                        next_pos = min(next_pos, next_marker_pos)
                
                section_data = content[marker_pos:next_pos]
                sections[marker.decode('ascii', errors='ignore')] = section_data
                current_pos = marker_pos + len(marker)
        
        return sections
    
    def _parse_cost_section(self, section_data: bytes, results: EconomicAnalysisResults):
        """解析成本数据段"""
        # 提取成本相关数据并添加到结果中
        numeric_values = self._extract_numeric_data(section_data)
        text_chunks = self._extract_text_chunks(section_data)
        
        # 创建成本项目
        for i, value in enumerate(numeric_values[:20]):  # 限制数量
            if value > 1000:  # 过滤小值
                cost_item = CostItem(
                    name=f"COST_ITEM_{i+1}",
                    category=CostCategory.EQUIPMENT,
                    base_cost=value,
                    estimation_method="SZP binary extraction"
                )
                results.capex_data.add_cost_item(cost_item)
    
    def _parse_financial_section(self, section_data: bytes, results: EconomicAnalysisResults):
        """解析财务数据段"""
        numeric_values = self._extract_numeric_data(section_data)
        
        # 查找可能的折现率、税率等财务参数
        percentage_values = [v for v in numeric_values if 0.01 < v < 1.0]
        
        if percentage_values:
            financial_params = FinancialParameters(project_name=results.project_name)
            
            # 尝试识别不同的财务参数
            if len(percentage_values) >= 1:
                financial_params.discount_rate = percentage_values[0]
            if len(percentage_values) >= 2:
                financial_params.tax_rate = percentage_values[1]
            
            results.financial_params = financial_params
    
    def _parse_equipment_section(self, section_data: bytes, results: EconomicAnalysisResults):
        """解析设备数据段"""
        # 提取设备相关信息
        pass  # 待实现具体逻辑
    
    def _is_text_content(self, content: bytes) -> bool:
        """
        判断内容是否为文本格式
        
        Args:
            content: 字节内容
            
        Returns:
            True if content appears to be text
        """
        try:
            # 尝试解码前1000字节
            sample = content[:1000].decode('utf-8')
            # 检查可打印字符的比例
            printable_ratio = sum(1 for c in sample if c.isprintable()) / len(sample)
            return printable_ratio > 0.7
        except:
            return False
    
    def _extract_text_based_data(self, text_content: str, results: EconomicAnalysisResults):
        """
        从文本内容中提取经济数据
        
        Args:
            text_content: 文本内容
            results: 经济分析结果容器
        """
        # 查找数值和相关的描述
        lines = text_content.split('\n')
        
        for line in lines:
            # 查找包含成本相关关键词和数值的行
            if any(keyword in line.upper() for keyword in ['COST', 'PRICE', 'TOTAL', 'CAPEX', 'OPEX']):
                # 提取数值
                import re
                numbers = re.findall(r'[\d,]+\.?\d*', line)
                if numbers:
                    try:
                        value = float(numbers[0].replace(',', ''))
                        if self._is_reasonable_cost_value(value):
                            # 创建成本项目
                            cost_item = CostItem(
                                name=f"TEXT_ITEM_{len(results.capex_data.equipment_costs)+1}",
                                category=CostCategory.OTHER,
                                base_cost=value,
                                estimation_method="Text extraction",
                                notes=[line.strip()]
                            )
                            results.capex_data.add_cost_item(cost_item)
                    except ValueError:
                        continue
    
    def _extract_version_number(self, version_data: bytes) -> str:
        """
        从版本数据中提取版本号
        
        Args:
            version_data: 包含版本信息的字节数据
            
        Returns:
            Version string if found
        """
        try:
            text = version_data.decode('ascii', errors='ignore')
            import re
            version_match = re.search(r'(\d+\.?\d*)', text)
            if version_match:
                return version_match.group(1)
        except:
            pass
        return "Unknown"
    
    def get_parsing_report(self) -> Dict[str, Any]:
        """
        获取解析报告
        
        Returns:
            Dictionary containing parsing statistics and issues
        """
        return {
            'supported_formats': self.supported_formats,
            'parsing_errors': self.parsing_errors,
            'warnings': self.warnings,
            'extracted_data_keys': list(self.extracted_data.keys())
        }


def test_economic_parser():
    """测试经济文件解析器"""
    parser = EconomicFileParser()
    
    # 查找测试文件
    test_files = [
        "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.izp",
        "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.szp"
    ]
    
    for test_file in test_files:
        if os.path.exists(test_file):
            print(f"\nTesting parser with: {test_file}")
            try:
                results = parser.parse_file(test_file)
                print(f"✅ Successfully parsed {test_file}")
                print(f"   Project: {results.project_name}")
                print(f"   Total CAPEX: ${results.total_capex:,.2f}")
                print(f"   Annual OPEX: ${results.annual_opex:,.2f}")
                print(f"   Equipment items: {len(results.equipment_list)}")
                print(f"   Data sources: {len(results.data_sources)}")
            except Exception as e:
                print(f"❌ Error parsing {test_file}: {str(e)}")
    
    # 打印解析报告
    report = parser.get_parsing_report()
    print(f"\n=== Parsing Report ===")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    test_economic_parser()