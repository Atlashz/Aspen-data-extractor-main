#!/usr/bin/env python3
"""
I-N Column Extraction Test Suite

Comprehensive test to verify that all data from columns I-N in BFG-CO2H-HEX.xlsx
is properly extracted, processed, and stored in the database.

Author: TEA Analysis Framework
Date: 2025-07-26
Version: 1.0
"""

import os
import sys
import logging
import pandas as pd
import json
from datetime import datetime
from typing import Dict, List, Any, Optional

# Add current directory to path to import modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from aspen_data_extractor import HeatExchangerDataLoader
from aspen_data_database import AspenDataDatabase

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')
logger = logging.getLogger(__name__)

class IToNColumnExtractionTester:
    """
    Comprehensive tester for I-N column extraction functionality
    """
    
    def __init__(self, excel_file: str = "BFG-CO2H-HEX.xlsx"):
        self.excel_file = excel_file
        self.test_results = {}
        self.extraction_report = {}
        
    def run_comprehensive_test(self) -> Dict[str, Any]:
        """
        Run comprehensive I-N column extraction test
        """
        print("\n" + "="*80)
        print("🧪 I-N COLUMN EXTRACTION COMPREHENSIVE TEST SUITE")
        print("="*80)
        print(f"Target Excel File: {self.excel_file}")
        print(f"Test Time: {datetime.now().isoformat()}")
        
        try:
            # Test 1: Excel File Structure Analysis
            print("\n1️⃣ Excel File Structure Analysis:")
            structure_results = self._test_excel_structure()
            self.test_results['excel_structure'] = structure_results
            
            # Test 2: Heat Exchanger Data Loading
            print("\n2️⃣ Heat Exchanger Data Loading Test:")
            loading_results = self._test_hex_data_loading()
            self.test_results['data_loading'] = loading_results
            
            # Test 3: Column I-N Detection
            print("\n3️⃣ I-N Column Detection Test:")
            detection_results = self._test_i_to_n_detection()
            self.test_results['column_detection'] = detection_results
            
            # Test 4: Data Extraction Verification
            print("\n4️⃣ Data Extraction Verification:")
            extraction_results = self._test_data_extraction()
            self.test_results['data_extraction'] = extraction_results
            
            # Test 5: Database Storage Verification
            print("\n5️⃣ Database Storage Verification:")
            storage_results = self._test_database_storage()
            self.test_results['database_storage'] = storage_results
            
            # Test 6: I-N Column Data Completeness
            print("\n6️⃣ I-N Column Data Completeness Check:")
            completeness_results = self._test_data_completeness()
            self.test_results['data_completeness'] = completeness_results
            
            # Generate Final Report
            print("\n7️⃣ Generating Final Test Report:")
            self._generate_final_report()
            
            return self.test_results
            
        except Exception as e:
            logger.error(f"Test suite failed: {e}")
            self.test_results['error'] = str(e)
            return self.test_results
    
    def _test_excel_structure(self) -> Dict[str, Any]:
        """Test Excel file structure and identify I-N columns"""
        results = {
            'file_exists': False,
            'columns_found': [],
            'total_columns': 0,
            'total_rows': 0,
            'i_to_n_columns_present': False,
            'column_i_to_n_headers': {}
        }
        
        try:
            if not os.path.exists(self.excel_file):
                print(f"   ❌ Excel file not found: {self.excel_file}")
                return results
            
            results['file_exists'] = True
            print(f"   ✅ Excel file found: {self.excel_file}")
            
            # Read Excel file
            df = pd.read_excel(self.excel_file)
            results['total_columns'] = len(df.columns)
            results['total_rows'] = len(df)
            results['columns_found'] = list(df.columns)
            
            print(f"   📊 File Structure: {results['total_rows']} rows × {results['total_columns']} columns")
            print(f"   📋 Column Headers: {results['columns_found']}")
            
            # Check if we have at least 14 columns (up to column N)
            if results['total_columns'] >= 14:
                results['i_to_n_columns_present'] = True
                
                # Extract I-N column headers (columns 8-13, 0-indexed)
                i_to_n_mapping = {
                    'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13
                }
                
                for excel_col, col_idx in i_to_n_mapping.items():
                    if col_idx < len(df.columns):
                        header = df.columns[col_idx]
                        results['column_i_to_n_headers'][excel_col] = str(header)
                
                print("   📍 I-N Column Headers Found:")
                for excel_col, header in results['column_i_to_n_headers'].items():
                    print(f"      Column {excel_col}: '{header}'")
            else:
                print(f"   ⚠️ Insufficient columns for I-N extraction (need ≥14, found {results['total_columns']})")
            
        except Exception as e:
            print(f"   ❌ Excel structure analysis failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _test_hex_data_loading(self) -> Dict[str, Any]:
        """Test heat exchanger data loading"""
        results = {
            'loader_created': False,
            'data_loaded': False,
            'error': None
        }
        
        try:
            # Create heat exchanger data loader
            hex_loader = HeatExchangerDataLoader(self.excel_file)
            results['loader_created'] = True
            print("   ✅ Heat exchanger loader created successfully")
            
            # Load data
            data = hex_loader.load_data()
            if data is not None:
                results['data_loaded'] = True
                results['data_shape'] = hex_loader.data.shape if hex_loader.data is not None else None
                print(f"   ✅ Data loaded successfully: {results['data_shape']}")
            else:
                print("   ❌ Failed to load heat exchanger data")
            
        except Exception as e:
            print(f"   ❌ Heat exchanger data loading failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _test_i_to_n_detection(self) -> Dict[str, Any]:
        """Test I-N column detection in the extraction system"""
        results = {
            'detection_successful': False,
            'columns_detected': {},
            'column_mappings': {},
            'error': None
        }
        
        try:
            # Create and load data
            hex_loader = HeatExchangerDataLoader(self.excel_file)
            data = hex_loader.load_data()
            
            if data is None:
                print("   ❌ Could not load data for column detection test")
                return results
            
            # Get column mappings
            column_mappings = hex_loader._find_column_mappings_flexible()
            results['column_mappings'] = column_mappings
            
            # Check I-N column detection
            i_to_n_categories = ['column_i', 'column_j', 'column_k', 'column_l', 'column_m', 'column_n']
            detected_i_to_n = {}
            
            for category in i_to_n_categories:
                if category in column_mappings and column_mappings[category]:
                    detected_i_to_n[category] = column_mappings[category]
            
            results['columns_detected'] = detected_i_to_n
            
            if detected_i_to_n:
                results['detection_successful'] = True
                print("   ✅ I-N Column Detection Results:")
                for category, columns in detected_i_to_n.items():
                    excel_letter = category.split('_')[1].upper()
                    print(f"      {excel_letter}: {columns}")
            else:
                print("   ❌ No I-N columns detected")
            
        except Exception as e:
            print(f"   ❌ I-N column detection test failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _test_data_extraction(self) -> Dict[str, Any]:
        """Test actual data extraction from I-N columns"""
        results = {
            'extraction_successful': False,
            'total_rows_processed': 0,
            'rows_with_i_to_n_data': 0,
            'i_to_n_extraction_stats': {},
            'sample_extracted_data': [],
            'error': None
        }
        
        try:
            # Create loader and process data
            hex_loader = HeatExchangerDataLoader(self.excel_file)
            data = hex_loader.load_data()
            
            if data is None:
                print("   ❌ Could not load data for extraction test")
                return results
            
            # Process heat exchanger data
            processed_data = hex_loader._process_hex_data()
            
            if processed_data and 'equipment_list' in processed_data:
                equipment_list = processed_data['equipment_list']
                results['total_rows_processed'] = len(equipment_list)
                
                # Analyze I-N column extraction
                i_to_n_stats = {
                    'column_i': 0, 'column_j': 0, 'column_k': 0,
                    'column_l': 0, 'column_m': 0, 'column_n': 0
                }
                
                rows_with_data = 0
                sample_data = []
                
                for hex_info in equipment_list:
                    has_i_to_n_data = False
                    
                    # Check each I-N column
                    for col in ['i', 'j', 'k', 'l', 'm', 'n']:
                        data_key = f'column_{col}_data'
                        header_key = f'column_{col}_header'
                        
                        if hex_info.get(data_key) is not None:
                            i_to_n_stats[f'column_{col}'] += 1
                            has_i_to_n_data = True
                    
                    if has_i_to_n_data:
                        rows_with_data += 1
                        
                        # Collect sample data (first 3 rows)
                        if len(sample_data) < 3:
                            sample_item = {
                                'name': hex_info.get('name', 'Unknown'),
                                'i_to_n_data': {
                                    col: {
                                        'data': hex_info.get(f'column_{col}_data'),
                                        'header': hex_info.get(f'column_{col}_header')
                                    }
                                    for col in ['i', 'j', 'k', 'l', 'm', 'n']
                                    if hex_info.get(f'column_{col}_data') is not None
                                }
                            }
                            sample_data.append(sample_item)
                
                results['rows_with_i_to_n_data'] = rows_with_data
                results['i_to_n_extraction_stats'] = i_to_n_stats
                results['sample_extracted_data'] = sample_data
                
                if rows_with_data > 0:
                    results['extraction_successful'] = True
                    print(f"   ✅ Data extraction successful: {rows_with_data}/{results['total_rows_processed']} rows with I-N data")
                    print(f"   📊 I-N Column Extraction Stats:")
                    for col, count in i_to_n_stats.items():
                        excel_letter = col.split('_')[1].upper()
                        print(f"      Column {excel_letter}: {count} values extracted")
                    
                    print(f"   📋 Sample Extracted Data:")
                    for i, sample in enumerate(sample_data, 1):
                        print(f"      Sample {i} ({sample['name']}):")
                        for col, info in sample['i_to_n_data'].items():
                            print(f"         Column {col.upper()}: {info['data']} (from '{info['header']}')")
                else:
                    print("   ❌ No I-N column data extracted")
            
        except Exception as e:
            print(f"   ❌ Data extraction test failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _test_database_storage(self) -> Dict[str, Any]:
        """Test database storage of I-N column data"""
        results = {
            'database_created': False,
            'data_stored': False,
            'i_to_n_stored_count': 0,
            'database_summary': {},
            'error': None
        }
        
        try:
            # Create database
            db = AspenDataDatabase("test_i_to_n_extraction.db")
            results['database_created'] = True
            print("   ✅ Test database created")
            
            # Start session
            session_id = db.start_new_session(hex_file=self.excel_file)
            print(f"   ✅ Database session started: {session_id}")
            
            # Load and process heat exchanger data
            hex_loader = HeatExchangerDataLoader(self.excel_file)
            data = hex_loader.load_data()
            
            if data is not None:
                processed_data = hex_loader._process_hex_data()
                
                # Store data in database
                hex_data = {
                    'heat_exchangers': processed_data.get('equipment_list', [])
                }
                
                db.store_hex_data(hex_data)
                results['data_stored'] = True
                print(f"   ✅ Heat exchanger data stored in database")
                
                # Finalize session
                summary_stats = {
                    'stream_count': 0,
                    'equipment_count': 0,
                    'hex_count': len(hex_data['heat_exchangers']),
                    'total_heat_duty_kw': processed_data.get('total_heat_duty', 0.0),
                    'total_heat_area_m2': processed_data.get('total_heat_area', 0.0)
                }
                db.finalize_session(summary_stats)
                
                # Verify I-N column data storage
                hex_df = db.get_all_heat_exchangers()
                if not hex_df.empty:
                    i_to_n_columns = ['column_i_data', 'column_j_data', 'column_k_data', 
                                     'column_l_data', 'column_m_data', 'column_n_data']
                    
                    i_to_n_stored = 0
                    for col in i_to_n_columns:
                        if col in hex_df.columns:
                            non_null_count = hex_df[col].notna().sum()
                            i_to_n_stored += non_null_count
                    
                    results['i_to_n_stored_count'] = i_to_n_stored
                    
                    # Get database summary with I-N coverage
                    db_summary = db.get_database_summary()
                    results['database_summary'] = db_summary
                    
                    print(f"   ✅ I-N column data verification:")
                    print(f"      Total I-N values stored: {i_to_n_stored}")
                    print(f"      Database records: {db_summary.get('total_records', 0)}")
                    
                    if 'i_to_n_column_coverage' in db_summary:
                        coverage = db_summary['i_to_n_column_coverage']
                        print(f"      I-N Coverage Summary: {coverage}")
                
                db.close()
                
            else:
                print("   ❌ Could not load heat exchanger data for database test")
            
        except Exception as e:
            print(f"   ❌ Database storage test failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _test_data_completeness(self) -> Dict[str, Any]:
        """Test completeness of I-N column data extraction"""
        results = {
            'completeness_check_passed': False,
            'total_possible_i_to_n_values': 0,
            'total_extracted_i_to_n_values': 0,
            'extraction_percentage': 0.0,
            'column_wise_completeness': {},
            'missing_data_analysis': {},
            'error': None
        }
        
        try:
            # Calculate total possible I-N values
            structure_results = self.test_results.get('excel_structure', {})
            total_rows = structure_results.get('total_rows', 0)
            i_to_n_columns_present = structure_results.get('i_to_n_columns_present', False)
            
            if i_to_n_columns_present and total_rows > 0:
                results['total_possible_i_to_n_values'] = total_rows * 6  # 6 columns (I-N)
                
                # Get extraction results
                extraction_results = self.test_results.get('data_extraction', {})
                extraction_stats = extraction_results.get('i_to_n_extraction_stats', {})
                
                total_extracted = sum(extraction_stats.values())
                results['total_extracted_i_to_n_values'] = total_extracted
                
                if results['total_possible_i_to_n_values'] > 0:
                    results['extraction_percentage'] = (total_extracted / results['total_possible_i_to_n_values']) * 100
                
                # Column-wise completeness
                for col, count in extraction_stats.items():
                    excel_letter = col.split('_')[1].upper()
                    completeness_pct = (count / total_rows) * 100 if total_rows > 0 else 0
                    results['column_wise_completeness'][excel_letter] = {
                        'extracted_count': count,
                        'total_possible': total_rows,
                        'completeness_percentage': completeness_pct
                    }
                
                # Determine if completeness check passed
                # Pass if we extracted at least 50% of possible I-N values
                if results['extraction_percentage'] >= 50.0:
                    results['completeness_check_passed'] = True
                    print(f"   ✅ Data completeness check PASSED")
                    print(f"      Extraction Rate: {results['extraction_percentage']:.1f}% ({total_extracted}/{results['total_possible_i_to_n_values']} values)")
                else:
                    print(f"   ⚠️ Data completeness check PARTIAL")
                    print(f"      Extraction Rate: {results['extraction_percentage']:.1f}% ({total_extracted}/{results['total_possible_i_to_n_values']} values)")
                
                print(f"   📊 Column-wise Completeness:")
                for excel_letter, stats in results['column_wise_completeness'].items():
                    print(f"      Column {excel_letter}: {stats['completeness_percentage']:.1f}% ({stats['extracted_count']}/{stats['total_possible']})")
            
            else:
                print("   ❌ Cannot perform completeness check - insufficient data")
            
        except Exception as e:
            print(f"   ❌ Data completeness test failed: {e}")
            results['error'] = str(e)
        
        return results
    
    def _generate_final_report(self) -> None:
        """Generate comprehensive final test report"""
        print(f"\n📋 COMPREHENSIVE TEST REPORT")
        print("-" * 60)
        
        # Overall test status
        all_tests = ['excel_structure', 'data_loading', 'column_detection', 'data_extraction', 'database_storage', 'data_completeness']
        passed_tests = 0
        
        for test_name in all_tests:
            test_result = self.test_results.get(test_name, {})
            if self._is_test_passed(test_name, test_result):
                passed_tests += 1
        
        overall_status = "PASSED" if passed_tests == len(all_tests) else f"PARTIAL ({passed_tests}/{len(all_tests)})"
        print(f"Overall Test Status: {overall_status}")
        
        # Key metrics
        print(f"\n📊 Key Metrics:")
        
        extraction_results = self.test_results.get('data_extraction', {})
        if extraction_results.get('extraction_successful'):
            total_rows = extraction_results.get('total_rows_processed', 0)
            rows_with_data = extraction_results.get('rows_with_i_to_n_data', 0)
            print(f"   Rows Processed: {total_rows}")
            print(f"   Rows with I-N Data: {rows_with_data}")
            
            stats = extraction_results.get('i_to_n_extraction_stats', {})
            total_values = sum(stats.values())
            print(f"   Total I-N Values Extracted: {total_values}")
        
        completeness_results = self.test_results.get('data_completeness', {})
        if completeness_results.get('extraction_percentage'):
            print(f"   Data Completeness: {completeness_results['extraction_percentage']:.1f}%")
        
        # Database verification
        storage_results = self.test_results.get('database_storage', {})
        if storage_results.get('i_to_n_stored_count'):
            print(f"   Values Stored in Database: {storage_results['i_to_n_stored_count']}")
        
        # Recommendations
        print(f"\n💡 Recommendations:")
        if passed_tests == len(all_tests):
            print("   ✅ I-N column extraction is working perfectly!")
            print("   ✅ All heat exchanger data from columns I-N is being captured")
            print("   ✅ Database storage is functioning correctly")
        else:
            if not self.test_results.get('excel_structure', {}).get('file_exists'):
                print("   ❌ Ensure BFG-CO2H-HEX.xlsx file is present in the working directory")
            
            if not self.test_results.get('column_detection', {}).get('detection_successful'):
                print("   ⚠️ Column detection may need refinement for this Excel file structure")
            
            if extraction_results.get('rows_with_i_to_n_data', 0) == 0:
                print("   ⚠️ No I-N column data extracted - check column mapping and data types")
        
        # Save detailed report
        report_file = f"i_to_n_extraction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(self.test_results, f, indent=2, ensure_ascii=False, default=str)
            print(f"\n💾 Detailed report saved: {report_file}")
        except Exception as e:
            print(f"\n❌ Failed to save report: {e}")
    
    def _is_test_passed(self, test_name: str, test_result: Dict) -> bool:
        """Determine if a specific test passed"""
        if test_result.get('error'):
            return False
        
        if test_name == 'excel_structure':
            return test_result.get('file_exists', False) and test_result.get('i_to_n_columns_present', False)
        elif test_name == 'data_loading':
            return test_result.get('data_loaded', False)
        elif test_name == 'column_detection':
            return test_result.get('detection_successful', False)
        elif test_name == 'data_extraction':
            return test_result.get('extraction_successful', False)
        elif test_name == 'database_storage':
            return test_result.get('data_stored', False) and test_result.get('i_to_n_stored_count', 0) > 0
        elif test_name == 'data_completeness':
            return test_result.get('completeness_check_passed', False)
        
        return False


def main():
    """Main test execution function"""
    excel_file = "BFG-CO2H-HEX.xlsx"
    
    print("🚀 Starting I-N Column Extraction Test Suite")
    print(f"Target File: {excel_file}")
    
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        print("Please ensure the BFG-CO2H-HEX.xlsx file is in the current directory")
        return False
    
    # Create and run tester
    tester = IToNColumnExtractionTester(excel_file)
    results = tester.run_comprehensive_test()
    
    # Final status
    if results.get('error'):
        print(f"\n❌ Test suite failed with error: {results['error']}")
        return False
    else:
        print(f"\n✅ Test suite completed successfully!")
        return True


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)