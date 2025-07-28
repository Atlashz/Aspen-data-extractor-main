#!/usr/bin/env python3
"""
Aspen Plus数据提取器

专用于从Aspen Plus仿真文件中提取工程数据并构建SQLite数据库的工具。
支持通过COM接口提取流股数据、设备参数和热交换器信息，
同时可处理Excel格式的热交换器数据表。

主要功能:
- Aspen Plus COM接口连接和数据提取
- Excel热交换器数据处理
- SQLite数据库构建和管理
- 数据验证和导出

Author: 数据提取工具
Date: 2025-07-25
Version: 2.0
"""

import os
import sys
import math
import json
import logging
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
from aspen_data_database import AspenDataDatabase

# Windows COM support (conditional import)
try:
    import win32com.client as win32
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    win32 = None
    pythoncom = None

# Import custom data interfaces
from data_interfaces import (
    AspenProcessData, StreamData, UnitOperationData, UtilityData,
    EquipmentSizeData, EquipmentType, MaterialType, PressureLevel,
    CostItem, CapexData, OpexData, FinancialParameters, EconomicAnalysisResults,
    CostCategory, CurrencyType, CostBasis
)

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Import stream classification
try:
    from stream_classifier import StreamClassifier, StreamCategory
    STREAM_CLASSIFICATION_AVAILABLE = True
    logger.info("✅ Stream classification module loaded")
except ImportError:
    STREAM_CLASSIFICATION_AVAILABLE = False
    logger.warning("⚠️ Stream classification module not available")

# Import equipment matching
try:
    import sys
    sys.path.append(os.path.join(os.path.dirname(__file__), 'equipment match'))
    from equipment_model_matcher import EquipmentModelMatcher
    EQUIPMENT_MATCHER_AVAILABLE = True
    logger.info("✅ Equipment matcher module loaded")
except ImportError:
    EQUIPMENT_MATCHER_AVAILABLE = False
    logger.warning("⚠️ Equipment matcher module not available")

# Import enhanced equipment detector
try:
    from enhanced_equipment_detector import EnhancedEquipmentDetector
    ENHANCED_EQUIPMENT_DETECTOR_AVAILABLE = True
    logger.info("✅ Enhanced equipment detector module loaded")
except ImportError:
    ENHANCED_EQUIPMENT_DETECTOR_AVAILABLE = False
    logger.warning("⚠️ Enhanced equipment detector module not available")



class AspenConnectionError(Exception):
    """Custom exception for Aspen connection issues"""
    pass


class AspenCOMInterface:
    """
    Interface to Aspen Plus using COM automation
    
    Handles connection to Aspen Plus and comprehensive data extraction.
    Based on proven methods from bfg_co2h_aspen_analyzer.py.
    Enhanced for Windows COM compatibility.
    """
    
    def __init__(self):
        self.app = None
        self.simulation = None
        self.connected = False
    
    def test_com_availability(self) -> Dict[str, Any]:
        """Test Windows COM setup for Aspen Plus"""
        test_results = {
            'pywin32_available': WIN32COM_AVAILABLE,
            'com_objects_found': [],
            'platform': sys.platform,
            'recommendations': []
        }
        
        # Check pywin32 availability
        if not WIN32COM_AVAILABLE:
            test_results['recommendations'].append("Install pywin32: pip install pywin32")
            logger.error("❌ pywin32 not available")
            return test_results
        
        logger.info("✅ pywin32 is available")
        
        # Test COM object availability
        com_objects = [
            "Apwn.Document",
            "AspenTech.AspenPlus.Document", 
            "Apwn.Document.1"
        ]
        
        pythoncom.CoInitialize()
        try:
            for com_obj in com_objects:
                try:
                    app = win32.Dispatch(com_obj)
                    test_results['com_objects_found'].append(com_obj)
                    app = None  # Release object
                    logger.info(f"✅ Found COM object: {com_obj}")
                except:
                    logger.warning(f"❌ COM object not available: {com_obj}")
        finally:
            pythoncom.CoUninitialize()
        
        # Provide recommendations
        if not test_results['com_objects_found']:
            test_results['recommendations'].extend([
                "Ensure Aspen Plus is installed on this Windows machine",
                "Run 'regsvr32 apwn.exe' as administrator to register COM objects",
                "Check if Aspen Plus is the correct version for your license"
            ])
        
        return test_results
        
    def connect(self, file_path: str = None, visible: bool = False, use_active: bool = False) -> bool:
        """
        Connect to Aspen Plus - either open file or connect to active instance
        
        Args:
            file_path: Path to Aspen Plus .apw or .bkp file (optional if use_active=True)
            visible: Whether to make Aspen Plus visible
            use_active: If True, connect to active instance instead of opening file
            
        Returns:
            bool: True if connection successful
        """
        try:
            # Check COM availability
            if not WIN32COM_AVAILABLE:
                logger.error("❌ win32com not available. This module requires Windows with Aspen Plus.")
                logger.error("   Install pywin32: pip install pywin32")
                return False
            
            logger.info(f"Attempting to connect to Aspen Plus...")
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Connect to COM object
            connection_success = False
            if use_active:
                # Try to connect to active instance first
                try:
                    self.app = win32.GetActiveObject("Apwn.Document")
                    logger.info("✅ Connected to active Aspen Plus instance")
                    connection_success = True
                except Exception as e:
                    logger.warning(f"GetActiveObject failed: {str(e)}, trying Dispatch")
            
            if not connection_success:
                # Try different COM object names
                com_objects = [
                    "Apwn.Document",
                    "AspenTech.AspenPlus.Document", 
                    "Apwn.Document.1"
                ]
                
                for com_obj in com_objects:
                    try:
                        logger.info(f"Trying COM object: {com_obj}")
                        self.app = win32.Dispatch(com_obj)
                        connection_success = True
                        logger.info(f"✅ Successfully created COM object: {com_obj}")
                        break
                    except Exception as e:
                        logger.warning(f"Failed to create COM object {com_obj}: {str(e)}")
                        continue
            
            if not connection_success:
                raise AspenConnectionError("Could not create any Aspen Plus COM object")
            
            # Configure application
            try:
                self.app.Visible = visible
                self.app.SuppressDialogs = True
                logger.info(f"Configured Aspen Plus visibility: {visible}")
            except Exception as e:
                logger.warning(f"Could not configure app settings: {str(e)}")
            
            # Initialize with file if provided and not using active instance
            if file_path and not use_active:
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"Aspen file not found: {file_path}")
                
                abs_path = os.path.abspath(file_path)
                logger.info(f"Attempting to open file: {abs_path}")
                
                initialization_methods = [
                    ("InitFromArchive2", lambda: self.app.InitFromArchive2(abs_path)),
                    ("InitFromArchive", lambda: self.app.InitFromArchive(abs_path)),
                    ("Open", lambda: self.app.Open(abs_path))
                ]
                
                init_success = False
                for method_name, method_func in initialization_methods:
                    try:
                        logger.info(f"Trying initialization method: {method_name}")
                        method_func()
                        init_success = True
                        logger.info(f"✅ Successfully initialized with: {method_name}")
                        break
                    except Exception as e:
                        logger.warning(f"Failed with {method_name}: {str(e)}")
                        continue
                
                if not init_success:
                    raise AspenConnectionError("Could not initialize Aspen Plus simulation")
            
            # Get simulation object
            self._get_simulation_object()
            
            # Test connection
            self._test_simulation_access()
            
            self.connected = True
            logger.info(f"🎉 Successfully connected to Aspen Plus")
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to connect to Aspen: {str(e)}")
            self.connected = False
            return False
    
    def _get_simulation_object(self):
        """Get simulation object from Aspen application"""
        try:
            # Primary: Try Tree object (this is what works!)
            if hasattr(self.app, 'Tree'):
                self.simulation = self.app.Tree
                if self.simulation:
                    logger.info("✅ Using app.Tree as simulation object")
                else:
                    raise Exception("app.Tree is None")
            else:
                # Fallback: Try Simulation object
                self.simulation = self.app.Simulation
                if self.simulation is None:
                    # Last resort: Try Engine.Simulation
                    if hasattr(self.app, 'Engine') and hasattr(self.app.Engine, 'Simulation'):
                        self.simulation = self.app.Engine.Simulation
                    
                    if self.simulation is None:
                        raise AspenConnectionError("Could not access any simulation object")
        except Exception as e:
            logger.error(f"Failed to get simulation object: {str(e)}")
            raise AspenConnectionError("Failed to access simulation object")
    
    def _test_simulation_access(self):
        """Test simulation access by trying to access basic properties"""
        try:
            test_node = self.simulation.FindNode("\\Data")
            if test_node is None:
                logger.warning("Could not find Data node - simulation may not be fully loaded")
            else:
                logger.info("✅ Successfully verified simulation access")
        except Exception as e:
            logger.warning(f"Could not verify simulation access: {str(e)}")
    
    def disconnect(self):
        """Disconnect from Aspen Plus and cleanup COM objects"""
        try:
            if self.app is not None:
                try:
                    # Try different close methods
                    if hasattr(self.app, 'Close'):
                        self.app.Close()
                    elif hasattr(self.app, 'Quit'):
                        self.app.Quit()
                    logger.info("Closed Aspen Plus application")
                except Exception as e:
                    logger.warning(f"Could not close Aspen Plus cleanly: {str(e)}")
                
                # Clear references
                self.app = None
                self.simulation = None
                self.connected = False
                
                # Cleanup COM
                try:
                    if WIN32COM_AVAILABLE:
                        pythoncom.CoUninitialize()
                except Exception as e:
                    logger.warning(f"COM cleanup warning: {str(e)}")
                
                logger.info("✅ Disconnected from Aspen Plus")
        except Exception as e:
            logger.warning(f"Error during disconnect: {str(e)}")
            # Force cleanup
            self.app = None
            self.simulation = None
            self.connected = False
    
    def connect_to_active(self, file_path: str = None) -> bool:
        """
        Connect to active Aspen Plus instance or initialize with file
        
        Args:
            file_path: Optional path to Aspen file if need to initialize
            
        Returns:
            bool: True if connection successful
        """
        try:
            # Check COM availability
            if not WIN32COM_AVAILABLE:
                logger.error("❌ win32com not available. This module requires Windows with Aspen Plus.")
                logger.error("   Install pywin32: pip install pywin32")
                return False
            
            logger.info("Attempting to connect to active Aspen Plus instance")
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Try to connect to active instance first
            try:
                self.app = win32.GetActiveObject("Apwn.Document")
                logger.info("✅ Connected to active Aspen Plus instance")
                connection_success = True
            except Exception as e:
                logger.warning(f"GetActiveObject failed: {e}")
                logger.info("Trying alternative connection method...")
                
                # Try Dispatch method
                try:
                    self.app = win32.Dispatch("Apwn.Document")
                    logger.info("✅ Connected using Dispatch")
                    connection_success = True
                except Exception as e2:
                    logger.error(f"Dispatch also failed: {e2}")
                    return False
            
            # Check if application is initialized
            try:
                # Test if simulation is accessible
                if hasattr(self.app, 'Simulation') and self.app.Simulation is not None:
                    logger.info("✅ Application already initialized")
                    init_success = True
                else:
                    logger.warning("Application not initialized: Apwn.Document.Simulation")
                    init_success = False
            except Exception as e:
                logger.warning(f"Cannot access simulation: {e}")
                init_success = False
            
            # Initialize with file if needed and file_path provided
            if not init_success and file_path:
                if not os.path.exists(file_path):
                    logger.error(f"File not found: {file_path}")
                    return False
                
                abs_path = os.path.abspath(file_path)
                logger.info(f"Attempting to initialize with file: {abs_path}")
                
                # Try different initialization methods
                initialization_methods = [
                    ("InitFromArchive2", lambda: self.app.InitFromArchive2(abs_path)),
                    ("InitFromArchive", lambda: self.app.InitFromArchive(abs_path)),
                    ("Open", lambda: self.app.Open(abs_path))
                ]
                
                for method_name, method_func in initialization_methods:
                    try:
                        logger.info(f"Trying initialization method: {method_name}")
                        method_func()
                        init_success = True
                        logger.info(f"✅ Successfully initialized with: {method_name}")
                        break
                    except Exception as e:
                        logger.warning(f"Failed with {method_name}: {e}")
                        continue
                
                if not init_success:
                    logger.error("Could not initialize Aspen Plus with any method")
                    return False
            
            # Get simulation object
            try:
                # Try Tree object first (proven to work)
                if hasattr(self.app, 'Tree') and self.app.Tree is not None:
                    self.simulation = self.app.Tree
                    logger.info("✅ Using app.Tree as simulation object")
                elif hasattr(self.app, 'Simulation') and self.app.Simulation is not None:
                    self.simulation = self.app.Simulation
                    logger.info("✅ Using app.Simulation as simulation object")
                else:
                    if hasattr(self.app, 'Engine') and hasattr(self.app.Engine, 'Simulation'):
                        self.simulation = self.app.Engine.Simulation
                        logger.info("✅ Using app.Engine.Simulation as simulation object")
                    else:
                        raise AspenConnectionError("Could not access simulation object")
            except Exception as e:
                logger.error(f"Failed to get simulation object: {str(e)}")
                raise AspenConnectionError("Failed to access simulation object")
            
            # Test connection by trying to access a basic property
            try:
                test_node = self.simulation.FindNode("\\Data")
                if test_node is None:
                    logger.warning("Could not find Data node - simulation may not be fully loaded")
                else:
                    logger.info("✅ Successfully verified simulation access")
            except Exception as e:
                logger.warning(f"Could not verify simulation access: {str(e)}")
            
            self.connected = True
            logger.info(f"🎉 Successfully connected to active Aspen Plus instance")
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to connect to Aspen: {str(e)}")
            logger.error("   Ensure Aspen Plus is installed and properly registered on Windows")
            self.connected = False
            return False
    
    
    def get_stream_property(self, stream_name: str, property_name: str) -> Optional[float]:
        """
        Get a specific property from a stream
        
        Args:
            stream_name: Name of the stream in Aspen
            property_name: Property name (e.g., 'TEMP_OUT', 'PRES_OUT', 'MASSFLMX')
            
        Returns:
            Property value or None if not found
        """
        try:
            # Use the FindNode approach that works with app.Tree
            property_path = f"\\Data\\Streams\\{stream_name}\\Output\\{property_name}\\MIXED"
            node = self.simulation.FindNode(property_path)
            
            if node and hasattr(node, 'Value'):
                value = node.Value
                return float(value) if value is not None else None
            else:
                logger.warning(f"Could not find node {property_path}")
                return None
                
        except Exception as e:
            logger.warning(f"Could not get {property_name} for stream {stream_name}: {str(e)}")
            return None
    
    def get_block_property(self, block_name: str, property_name: str) -> Optional[float]:
        """
        Get a specific property from a unit operation block
        
        Args:
            block_name: Name of the block in Aspen
            property_name: Property name (e.g., 'DUTY', 'PRES-DROP')
            
        Returns:
            Property value or None if not found
        """
        try:
            block = self.simulation.Flowsheet.Blocks(block_name)
            value = block.GetValue(property_name)
            return float(value) if value is not None else None
        except Exception as e:
            logger.warning(f"Could not get {property_name} for block {block_name}: {str(e)}")
            return None
    
    def get_stream_composition(self, stream_name: str) -> Dict[str, float]:
        """
        Get mole fraction composition of a stream using Tree node access
        
        Args:
            stream_name: Name of the stream in Aspen
            
        Returns:
            Dictionary of component mole fractions
        """
        composition = {}
        try:
            # Use Tree node approach that works with our connection method
            components_path = f"\\Data\\Streams\\{stream_name}\\Output\\MOLEFRAC\\MIXED"
            comp_node = self.simulation.FindNode(components_path)
            
            if comp_node and hasattr(comp_node, 'Elements'):
                # Get all components
                for i in range(comp_node.Elements.Count):
                    try:
                        comp_element = comp_node.Elements.Item(i)
                        comp_name = comp_element.Name
                        comp_value = comp_element.Value
                        
                        if comp_value is not None and comp_value > 1e-6:
                            composition[comp_name] = float(comp_value)
                    except Exception as e:
                        logger.warning(f"Error getting component {i} for stream {stream_name}: {str(e)}")
            else:
                # Alternative approach: try individual component paths
                common_components = ['H2', 'CO', 'CO2', 'H2O', 'CH4', 'CH3OH', 'N2', 'O2']
                for comp in common_components:
                    try:
                        comp_path = f"\\Data\\Streams\\{stream_name}\\Output\\MOLEFRAC\\MIXED\\{comp}"
                        comp_node = self.simulation.FindNode(comp_path)
                        if comp_node and hasattr(comp_node, 'Value'):
                            value = comp_node.Value
                            if value is not None and value > 1e-6:
                                composition[comp] = float(value)
                    except Exception:
                        continue
                        
        except Exception as e:
            logger.warning(f"Could not get composition for stream {stream_name}: {str(e)}")
            
        return composition
    
    def get_all_streams(self) -> List[str]:
        """Get all stream names from Aspen simulation"""
        try:
            streams_node = self.simulation.FindNode(r"\Data\Streams")
            if streams_node and hasattr(streams_node, 'Elements'):
                stream_names = []
                for i in range(streams_node.Elements.Count):
                    stream_name = streams_node.Elements.Item(i).Name
                    stream_names.append(stream_name)
                return stream_names
            else:
                logger.error("Could not access streams node or no Elements attribute")
                return []
        except Exception as e:
            logger.error(f"Could not get stream names: {str(e)}")
            return []
    
    def get_all_blocks(self) -> List[str]:
        """Get all block names from Aspen simulation"""
        try:
            blocks_node = self.simulation.FindNode(r"\Data\Blocks")
            if blocks_node and hasattr(blocks_node, 'Elements'):
                block_names = []
                for i in range(blocks_node.Elements.Count):
                    block_name = blocks_node.Elements.Item(i).Name
                    block_names.append(block_name)
                return block_names
            else:
                logger.error("Could not access blocks node or no Elements attribute")
                return []
        except Exception as e:
            logger.error(f"Could not get block names: {str(e)}")
            return []
    
    def get_aspen_value(self, path: str):
        """Get value from Aspen Plus tree node using our Tree connection"""
        try:
            node = self.simulation.FindNode(path)
            if node and hasattr(node, 'Value'):
                return node.Value
            return None
        except Exception as e:
            logger.debug(f"Could not get value from path {path}: {str(e)}")
            return None
    
    def get_block_type(self, block_name: str) -> Optional[str]:
        """Get Aspen block type using Tree node access"""
        try:
            # The TYPE parameter is in the Input section
            type_path = f"\\Data\\Blocks\\{block_name}\\Input\\TYPE"
            node = self.simulation.FindNode(type_path)
            
            if node and hasattr(node, 'Value') and node.Value:
                return str(node.Value)
            
            # Fallback: try other possible paths
            fallback_paths = [
                f"\\Data\\Blocks\\{block_name}\\Subobject",
                f"\\Data\\Blocks\\{block_name}\\BlockType",
                f"\\Data\\Blocks\\{block_name}\\Input\\CLASS",
                f"\\Data\\Blocks\\{block_name}\\Input\\MODEL"
            ]
            
            for path in fallback_paths:
                try:
                    node = self.simulation.FindNode(path)
                    if node and hasattr(node, 'Value') and node.Value:
                        return str(node.Value)
                except Exception:
                    continue
            
            return None
            
        except Exception as e:
            logger.debug(f"Could not get block type for {block_name}: {str(e)}")
            return None
    
    def get_equipment_parameters(self, block_name: str) -> Dict[str, Any]:
        """Get equipment parameters"""
        if not self.connected:
            logger.warning("Not connected to Aspen Plus")
            return {}
            
        parameters = {}
        
        # Get specific parameters based on equipment type
        equipment_type = self.get_block_type(block_name)
        
        try:
            # Get common parameters
            common_params = {
                'TEMP': 'temperature',
                'PRES': 'pressure', 
                'DUTY': 'duty',
                'QCALC': 'calculated_duty'
            }
            
            # Add specific parameters based on equipment type
            if equipment_type == 'ISENTROPIC':  # Compressor
                specific_params = {
                    'WNET': 'net_work',
                    'BRAKE_POWER': 'brake_power',
                    'B_PRES': 'inlet_pressure',
                    'B_PRES2': 'outlet_pressure',
                    'B_TEMP': 'inlet_temperature',
                    'B_TEMP2': 'outlet_temperature',
                    'COMP_DUTY': 'compression_duty'
                }
            elif equipment_type == 'T-SPEC':  # Heater
                specific_params = {
                    'B_TEMP': 'inlet_temperature',
                    'B_TEMP2': 'outlet_temperature',
                    'B_PRES': 'inlet_pressure',
                    'B_PRES2': 'outlet_pressure'
                }
            elif equipment_type in ['RADFRAC', 'DSTWU']:  # Distillation column
                specific_params = {
                    'BOTTOM_TEMP': 'bottom_temperature',
                    'BOT_LFLOW': 'bottom_liquid_flow',
                    'BOT_VFLOW': 'bottom_vapor_flow',
                    'BU_RATIO': 'bottoms_up_ratio',
                    'B_PRES': 'bottom_pressure'
                }
            elif equipment_type in ['HEATX', 'HEATER']:  # Heat exchanger
                specific_params = {
                    'B_TEMP': 'hot_inlet_temp',
                    'B_PRES': 'hot_inlet_pressure',
                    'TEMP_OUT': 'outlet_temperature',
                    'PRES_OUT': 'outlet_pressure',
                    'IN_PRES': 'inlet_pressure',
                    'TOT_MASS_ABS': 'total_mass_absorbed'
                }
            else:
                specific_params = {}
            
            # Merge parameter mappings
            all_params = {**common_params, **specific_params}
            
            # Get parameter values
            for aspen_param, friendly_name in all_params.items():
                value = None
                # Try two path formats
                paths_to_try = [
                    f"\\Data\\Blocks\\{block_name}\\Output\\{aspen_param}\\MIXED",
                    f"\\Data\\Blocks\\{block_name}\\Output\\{aspen_param}"
                ]
                
                for path in paths_to_try:
                    try:
                        value = self.get_aspen_value(path)
                        if value is not None:
                            break
                    except Exception:
                        continue
                
                # Only add valid non-zero values
                if value is not None and (isinstance(value, (int, float)) and value != 0):
                    parameters[friendly_name] = value
                    logger.debug(f"Found {friendly_name}: {value}")
                    
        except Exception as e:
            logger.error(f"Error getting equipment parameters for {block_name}: {e}")
            
        return parameters
    
    def get_stream_display_name(self, stream_name: str) -> str:
        """
        Get user-defined display name for a stream from Aspen Plus
        
        Args:
            stream_name: System name of the stream
            
        Returns:
            User-defined display name if found, otherwise returns system name
        """
        # The stream_name from Aspen is already the user-defined name (e.g., "H2IN", "BFG-FEED")
        # Return it directly as the custom name
        return stream_name
    
    def get_equipment_display_name(self, block_name: str) -> str:
        """
        Get user-defined display name for equipment from Aspen Plus
        
        Args:
            block_name: System name of the equipment block
            
        Returns:
            User-defined display name if found, otherwise returns system name
        """
        # The block_name from Aspen is already the user-defined name (e.g., "B1", "COOL2")
        # Return it directly as the custom name
        return block_name
    
    # =================== Economic Data Extraction Methods ===================
    
    def extract_economic_data(self, project_name: str = None) -> EconomicAnalysisResults:
        """
        Extract economic data from Aspen Plus simulation
        
        Args:
            project_name: Optional project name for the analysis
            
        Returns:
            EconomicAnalysisResults: Complete economic analysis data
        """
        if not self.connected:
            raise AspenConnectionError("Not connected to Aspen Plus")
        
        logger.info("🔄 Extracting economic data from Aspen Plus...")
        
        # Create economic analysis results container
        results = EconomicAnalysisResults(
            project_name=project_name or "Aspen_Simulation",
            timestamp=datetime.now(),
            analysis_version="1.0"
        )
        
        try:
            # Extract CAPEX data
            capex_data = self._extract_capex_data(results.project_name)
            results.capex_data = capex_data
            results.total_capex = capex_data.calculate_total_capex()
            
            # Extract OPEX data
            opex_data = self._extract_opex_data(results.project_name, results.total_capex)
            results.opex_data = opex_data
            results.annual_opex = opex_data.calculate_annual_opex(results.total_capex)
            
            # Extract financial parameters
            financial_params = self._extract_financial_parameters(results.project_name)
            results.financial_params = financial_params
            
            # Calculate economic indicators
            self._calculate_economic_indicators(results)
            
            # Extract equipment sizing for costing
            results.equipment_list = self._extract_equipment_for_costing()
            
            # Add metadata
            results.data_sources.extend([
                "Aspen Plus COM Interface",
                "Equipment sizing correlations",
                "Standard cost databases"
            ])
            
            results.estimation_methods.extend([
                "Equipment module costing",
                "Installation factor method",
                "Utilities consumption analysis"
            ])
            
            logger.info("✅ Economic data extraction completed")
            
        except Exception as e:
            logger.error(f"Error extracting economic data: {str(e)}")
            results.assumptions.append(f"Extraction error: {str(e)}")
            raise
        
        return results
    
    def _extract_capex_data(self, project_name: str) -> CapexData:
        """
        Extract capital expenditure data from Aspen simulation
        
        Args:
            project_name: Project name for the analysis
            
        Returns:
            CapexData: Capital cost data structure
        """
        capex_data = CapexData(
            project_name=project_name,
            currency=CurrencyType.USD,
            basis_year=2024
        )
        
        try:
            # Extract equipment data for costing
            equipment_data = self.extract_all_equipment()
            
            # Create EquipmentSizer for cost estimation
            sizer = EquipmentSizer()
            
            for block_name, eq_data in equipment_data.items():
                try:
                    # Get equipment type and parameters
                    eq_type = eq_data.get('type', 'unknown')
                    params = eq_data.get('parameters', {})
                    
                    # Estimate equipment cost based on type and size
                    base_cost = self._estimate_equipment_cost(eq_type, params, block_name)
                    
                    if base_cost > 0:
                        # Create cost item for equipment
                        cost_item = CostItem(
                            name=block_name,
                            category=CostCategory.EQUIPMENT,
                            base_cost=base_cost,
                            currency=CurrencyType.USD,
                            quantity=1.0,
                            unit="each",
                            installation_factor=2.5,  # Typical installation factor
                            material_factor=1.0,
                            location_factor=1.0,
                            escalation_factor=1.05,  # 5% escalation
                            estimation_method="COM interface sizing",
                            cost_basis=CostBasis.INSTALLED
                        )
                        
                        capex_data.add_cost_item(cost_item)
                        logger.debug(f"Added equipment cost: {block_name} - ${base_cost:,.0f}")
                
                except Exception as e:
                    logger.warning(f"Could not estimate cost for {block_name}: {str(e)}")
            
            # Add indirect costs
            self._add_indirect_capex_costs(capex_data)
            
        except Exception as e:
            logger.error(f"Error extracting CAPEX data: {str(e)}")
            raise
        
        return capex_data
    
    def _extract_opex_data(self, project_name: str, total_capex: float) -> OpexData:
        """
        Extract operating expenditure data from Aspen simulation
        
        Args:
            project_name: Project name for the analysis
            total_capex: Total capital expenditure for percentage-based costs
            
        Returns:
            OpexData: Operating cost data structure
        """
        opex_data = OpexData(
            project_name=project_name,
            currency=CurrencyType.USD,
            operating_hours=8760.0  # Full year operation
        )
        
        try:
            # Extract utility consumption from equipment
            self._extract_utility_costs(opex_data)
            
            # Extract raw material costs from streams
            self._extract_raw_material_costs(opex_data)
            
            # Add labor costs (estimated)
            self._add_labor_costs(opex_data, total_capex)
            
            # Add maintenance costs
            self._add_maintenance_costs(opex_data, total_capex)
            
        except Exception as e:
            logger.error(f"Error extracting OPEX data: {str(e)}")
            raise
        
        return opex_data
    
    def _extract_financial_parameters(self, project_name: str) -> FinancialParameters:
        """
        Extract or set default financial parameters
        
        Args:
            project_name: Project name for the analysis
            
        Returns:
            FinancialParameters: Financial analysis parameters
        """
        financial_params = FinancialParameters(
            project_name=project_name,
            project_life=20,
            discount_rate=0.10,
            tax_rate=0.25,
            depreciation_method="straight_line",
            depreciation_life=10
        )
        
        # Try to extract production capacity from major product streams
        try:
            stream_data = self.extract_all_streams()
            product_streams = self._identify_product_streams(stream_data)
            
            if product_streams:
                # Use the largest product stream as main product
                main_product = max(product_streams.items(), key=lambda x: x[1].mass_flow)
                financial_params.annual_production = main_product[1].mass_flow * 8760  # kg/year
                logger.info(f"Estimated annual production: {financial_params.annual_production:,.0f} kg/year")
        
        except Exception as e:
            logger.warning(f"Could not estimate production capacity: {str(e)}")
        
        return financial_params
    
    def _estimate_equipment_cost(self, eq_type: str, params: Dict[str, Any], block_name: str) -> float:
        """
        Estimate equipment cost based on type and parameters
        
        Args:
            eq_type: Equipment type string
            params: Equipment parameters from Aspen
            block_name: Equipment block name
            
        Returns:
            Estimated equipment cost in USD
        """
        base_cost = 0.0
        
        try:
            # Equipment-specific cost correlations
            if 'REACTOR' in eq_type.upper() or 'CSTR' in eq_type.upper():
                # Reactor costing based on volume
                volume = params.get('volume_m3', 10.0)  # Default 10 m3
                base_cost = 50000 * (volume ** 0.6)  # Cost scaling factor
                
            elif 'PUMP' in eq_type.upper():
                # Pump costing based on power
                power = params.get('power_kW', 10.0)  # Default 10 kW
                base_cost = 5000 * (power ** 0.7)
                
            elif 'COMPRESSOR' in eq_type.upper():
                # Compressor costing based on power
                power = params.get('power_kW', 100.0)  # Default 100 kW
                base_cost = 15000 * (power ** 0.7)
                
            elif 'HEAT' in eq_type.upper() or 'HX' in eq_type.upper():
                # Heat exchanger costing based on area
                area = params.get('area_m2', 100.0)  # Default 100 m2
                base_cost = 1000 * (area ** 0.65)
                
            elif 'COLUMN' in eq_type.upper() or 'DISTIL' in eq_type.upper():
                # Distillation column costing based on diameter and height
                diameter = params.get('diameter_m', 2.0)  # Default 2 m
                height = params.get('height_m', 20.0)    # Default 20 m
                base_cost = 25000 * (diameter ** 1.5) * (height ** 0.8)
                
            elif 'SEPARATOR' in eq_type.upper() or 'FLASH' in eq_type.upper():
                # Separator costing based on volume
                volume = params.get('volume_m3', 5.0)  # Default 5 m3
                base_cost = 20000 * (volume ** 0.6)
                
            elif 'TANK' in eq_type.upper() or 'VESSEL' in eq_type.upper():
                # Tank costing based on volume
                volume = params.get('volume_m3', 20.0)  # Default 20 m3
                base_cost = 8000 * (volume ** 0.7)
                
            else:
                # Generic equipment cost
                base_cost = 25000  # Default cost for unknown equipment
            
            # Apply minimum cost
            base_cost = max(base_cost, 5000)  # Minimum $5,000
            
            logger.debug(f"Estimated cost for {block_name} ({eq_type}): ${base_cost:,.0f}")
            
        except Exception as e:
            logger.warning(f"Error estimating cost for {block_name}: {str(e)}")
            base_cost = 25000  # Default fallback cost
        
        return base_cost
    
    def _add_indirect_capex_costs(self, capex_data: CapexData):
        """
        Add indirect CAPEX costs (engineering, construction, etc.)
        
        Args:
            capex_data: CAPEX data structure to modify
        """
        # Calculate equipment subtotal
        equipment_total = sum(item.calculate_installed_cost() 
                            for item in capex_data.equipment_costs.values())
        
        if equipment_total > 0:
            # Engineering costs (12% of equipment)
            engineering_cost = CostItem(
                name="Engineering & Design",
                category=CostCategory.ENGINEERING,
                base_cost=equipment_total * 0.12,
                currency=CurrencyType.USD,
                estimation_method="Percentage of equipment cost"
            )
            capex_data.add_cost_item(engineering_cost)
            
            # Construction management (8% of equipment)
            construction_cost = CostItem(
                name="Construction Management",
                category=CostCategory.CONSTRUCTION,
                base_cost=equipment_total * 0.08,
                currency=CurrencyType.USD,
                estimation_method="Percentage of equipment cost"
            )
            capex_data.add_cost_item(construction_cost)
    
    def _extract_utility_costs(self, opex_data: OpexData):
        """
        Extract utility costs from equipment data
        
        Args:
            opex_data: OPEX data structure to modify
        """
        try:
            equipment_data = self.extract_all_equipment()
            
            # Standard utility rates (USD per unit)
            utility_rates = {
                'electricity': 0.08,      # $/kWh
                'steam': 25.0,           # $/MT
                'cooling_water': 0.05,   # $/m3
                'fuel_gas': 8.0          # $/GJ
            }
            
            total_power = 0.0
            total_heating = 0.0
            total_cooling = 0.0
            
            # Sum up utility consumption from all equipment
            for block_name, eq_data in equipment_data.items():
                params = eq_data.get('parameters', {})
                
                # Power consumption
                power_kw = params.get('power_kW', 0.0)
                if power_kw > 0:
                    total_power += power_kw
                
                # Heating duty
                duty_kw = params.get('duty_kW', 0.0)
                if duty_kw > 0:  # Heating
                    total_heating += duty_kw
                elif duty_kw < 0:  # Cooling
                    total_cooling += abs(duty_kw)
            
            # Calculate annual utility costs
            if total_power > 0:
                annual_electricity_cost = total_power * 8760 * utility_rates['electricity']
                electricity_item = CostItem(
                    name="Electricity",
                    category=CostCategory.UTILITIES,
                    base_cost=annual_electricity_cost,
                    currency=CurrencyType.USD,
                    unit="kWh/year",
                    estimation_method="Equipment power consumption"
                )
                opex_data.add_opex_item(electricity_item)
            
            if total_heating > 0:
                # Convert heating duty to steam consumption (assuming 2000 kJ/kg steam)
                steam_consumption_kg_hr = total_heating * 3600 / 2000000  # kg/hr
                annual_steam_cost = steam_consumption_kg_hr * 8760 * utility_rates['steam'] / 1000
                steam_item = CostItem(
                    name="Process Steam",
                    category=CostCategory.UTILITIES,
                    base_cost=annual_steam_cost,
                    currency=CurrencyType.USD,
                    unit="MT/year",
                    estimation_method="Heating duty conversion"
                )
                opex_data.add_opex_item(steam_item)
            
            if total_cooling > 0:
                # Convert cooling duty to cooling water (assuming 20°C ΔT)
                cooling_water_m3_hr = total_cooling * 3600 / (4186 * 1000 * 20)  # m3/hr
                annual_cooling_cost = cooling_water_m3_hr * 8760 * utility_rates['cooling_water']
                cooling_item = CostItem(
                    name="Cooling Water",
                    category=CostCategory.UTILITIES,
                    base_cost=annual_cooling_cost,
                    currency=CurrencyType.USD,
                    unit="m3/year",
                    estimation_method="Cooling duty conversion"
                )
                opex_data.add_opex_item(cooling_item)
                
        except Exception as e:
            logger.warning(f"Error extracting utility costs: {str(e)}")
    
    def _extract_raw_material_costs(self, opex_data: OpexData):
        """
        Extract raw material costs from feed streams
        
        Args:
            opex_data: OPEX data structure to modify
        """
        try:
            stream_data = self.extract_all_streams()
            feed_streams = self._identify_feed_streams(stream_data)
            
            # Standard raw material prices (USD per MT)
            material_prices = {
                'H2': 2000,      # Hydrogen
                'CO': 500,       # Carbon monoxide
                'CO2': 100,      # Carbon dioxide
                'CH4': 300,      # Methane
                'METHANOL': 400, # Methanol
                'WATER': 1,      # Water
                'default': 500   # Default price for unknown materials
            }
            
            for stream_name, stream in feed_streams.items():
                # Estimate material cost based on main component
                main_component = self._identify_main_component(stream)
                price_per_mt = material_prices.get(main_component.upper(), material_prices['default'])
                
                # Calculate annual cost
                annual_consumption_mt = stream.mass_flow * 8760 / 1000  # Convert kg/hr to MT/year
                annual_cost = annual_consumption_mt * price_per_mt
                
                material_item = CostItem(
                    name=f"Raw Material - {stream_name}",
                    category=CostCategory.RAW_MATERIALS,
                    base_cost=annual_cost,
                    currency=CurrencyType.USD,
                    unit="MT/year",
                    quantity=annual_consumption_mt,
                    estimation_method="Feed stream analysis"
                )
                opex_data.add_opex_item(material_item)
                
        except Exception as e:
            logger.warning(f"Error extracting raw material costs: {str(e)}")
    
    def _add_labor_costs(self, opex_data: OpexData, total_capex: float):
        """
        Add estimated labor costs
        
        Args:
            opex_data: OPEX data structure to modify
            total_capex: Total CAPEX for scaling labor costs
        """
        # Estimate labor requirements based on plant size
        annual_labor_cost = max(total_capex * 0.02, 500000)  # 2% of CAPEX or minimum $500k
        
        labor_item = CostItem(
            name="Operating Labor",
            category=CostCategory.LABOR,
            base_cost=annual_labor_cost,
            currency=CurrencyType.USD,
            unit="$/year",
            estimation_method="Percentage of CAPEX"
        )
        opex_data.add_opex_item(labor_item)
    
    def _add_maintenance_costs(self, opex_data: OpexData, total_capex: float):
        """
        Add estimated maintenance costs
        
        Args:
            opex_data: OPEX data structure to modify
            total_capex: Total CAPEX for calculating maintenance costs
        """
        # Maintenance typically 3-5% of CAPEX annually
        annual_maintenance_cost = total_capex * 0.04  # 4% of CAPEX
        
        maintenance_item = CostItem(
            name="Maintenance & Repairs",
            category=CostCategory.MAINTENANCE,
            base_cost=annual_maintenance_cost,
            currency=CurrencyType.USD,
            unit="$/year",
            estimation_method="Percentage of CAPEX"
        )
        opex_data.add_opex_item(maintenance_item)
    
    def _calculate_economic_indicators(self, results: EconomicAnalysisResults):
        """
        Calculate economic indicators (NPV, IRR, etc.)
        
        Args:
            results: Economic analysis results to update
        """
        try:
            # Calculate NPV
            results.npv = results.financial_params.calculate_npv(
                results.total_capex, 
                results.annual_opex
            )
            
            # Calculate production cost per unit
            results.production_cost = results.calculate_production_cost()
            
            # Simple payback period
            if results.annual_opex > 0:
                annual_profit = results.financial_params.annual_revenue - results.annual_opex
                if annual_profit > 0:
                    results.payback_period = results.total_capex / annual_profit
            
            # Update results
            results.irr = results.financial_params.irr or 0.0
            
        except Exception as e:
            logger.warning(f"Error calculating economic indicators: {str(e)}")
    
    def _extract_equipment_for_costing(self) -> Dict[str, EquipmentSizeData]:
        """
        Extract equipment data formatted for costing analysis
        
        Returns:
            Dictionary of equipment sizing data
        """
        equipment_list = {}
        
        try:
            equipment_data = self.extract_all_equipment()
            sizer = EquipmentSizer()
            
            for block_name, eq_data in equipment_data.items():
                try:
                    eq_type = eq_data.get('type', 'unknown')
                    params = eq_data.get('parameters', {})
                    
                    # Convert to EquipmentType enum
                    equipment_type = self._map_to_equipment_type(eq_type)
                    
                    # Create equipment sizing data
                    sizing_data = EquipmentSizeData(
                        equipment_type=equipment_type,
                        name=block_name,
                        volume=params.get('volume_m3'),
                        area=params.get('area_m2'),
                        diameter=params.get('diameter_m'),
                        length=params.get('length_m'),
                        height=params.get('height_m'),
                        power_rating=params.get('power_kW'),
                        design_pressure=params.get('pressure_bar'),
                        design_temperature=params.get('temperature_C'),
                        estimated_cost=self._estimate_equipment_cost(eq_type, params, block_name),
                        cost_basis="2024 USD, installed"
                    )
                    
                    equipment_list[block_name] = sizing_data
                    
                except Exception as e:
                    logger.warning(f"Error processing equipment {block_name}: {str(e)}")
        
        except Exception as e:
            logger.error(f"Error extracting equipment for costing: {str(e)}")
        
        return equipment_list
    
    def _identify_feed_streams(self, stream_data: Dict[str, StreamData]) -> Dict[str, StreamData]:
        """Identify feed streams from stream data"""
        feed_streams = {}
        
        for name, stream in stream_data.items():
            # Simple heuristic: streams with "FEED", "IN", "INPUT" in name
            if any(keyword in name.upper() for keyword in ['FEED', 'IN', 'INPUT']):
                feed_streams[name] = stream
        
        return feed_streams
    
    def _identify_product_streams(self, stream_data: Dict[str, StreamData]) -> Dict[str, StreamData]:
        """Identify product streams from stream data"""
        product_streams = {}
        
        for name, stream in stream_data.items():
            # Simple heuristic: streams with "PRODUCT", "OUT", "OUTPUT" in name
            if any(keyword in name.upper() for keyword in ['PRODUCT', 'OUT', 'OUTPUT']):
                product_streams[name] = stream
        
        return product_streams
    
    def _identify_main_component(self, stream: StreamData) -> str:
        """Identify main component in a stream"""
        if stream.composition:
            # Find component with highest mole fraction
            main_comp = max(stream.composition.items(), key=lambda x: x[1])
            return main_comp[0]
        return "UNKNOWN"
    
    def _map_to_equipment_type(self, eq_type_str: str) -> EquipmentType:
        """Map equipment type string to EquipmentType enum"""
        eq_type_upper = eq_type_str.upper()
        
        if 'REACTOR' in eq_type_upper:
            return EquipmentType.REACTOR
        elif 'PUMP' in eq_type_upper:
            return EquipmentType.PUMP
        elif 'COMPRESSOR' in eq_type_upper:
            return EquipmentType.COMPRESSOR
        elif 'HEAT' in eq_type_upper or 'HX' in eq_type_upper:
            return EquipmentType.HEAT_EXCHANGER
        elif 'COLUMN' in eq_type_upper or 'DISTIL' in eq_type_upper:
            return EquipmentType.DISTILLATION_COLUMN
        elif 'SEPARATOR' in eq_type_upper or 'FLASH' in eq_type_upper:
            return EquipmentType.SEPARATOR
        elif 'TANK' in eq_type_upper or 'VESSEL' in eq_type_upper:
            return EquipmentType.TANK
        else:
            return EquipmentType.UNKNOWN
    
    # =================== Flowsheet Connection Methods ===================
    
    def load_flowsheet_connections(self, flowsheet_file: str = "aspen_flowsheet.xlsx") -> Dict[str, Dict[str, List[str]]]:
        """
        从Excel文件加载流股连接信息
        
        Args:
            flowsheet_file: Excel文件路径
            
        Returns:
            设备连接信息字典，格式: {设备名: {'inlet_streams': [], 'outlet_streams': []}}
        """
        self.flowsheet_connections = {}
        
        try:
            if not os.path.exists(flowsheet_file):
                logger.debug(f"Flowsheet file not found: {flowsheet_file}")
                return self.flowsheet_connections
                
            logger.debug(f"Loading flowsheet connections from: {flowsheet_file}")
            
            # 读取Excel文件
            df = pd.read_excel(flowsheet_file, sheet_name='Aspen Data Tables')
            
            # 解析连接信息
            stream_connections = self.parse_excel_connections(df)
            
            # 构建设备连接关系
            self.flowsheet_connections = self.build_equipment_connections(stream_connections)
            
            logger.debug(f"Loaded connections for {len(self.flowsheet_connections)} equipment items")
            
        except Exception as e:
            logger.debug(f"Failed to load flowsheet connections: {str(e)}")
            self.flowsheet_connections = {}
        
        return self.flowsheet_connections
    
    def parse_excel_connections(self, df) -> Dict[str, Dict[str, Optional[str]]]:
        """
        解析Excel中的连接信息
        
        Args:
            df: Excel DataFrame
            
        Returns:
            流股连接信息字典，格式: {流股名: {'from': 设备名, 'to': 设备名}}
        """
        stream_connections = {}
        
        try:
            # 找到关键行
            stream_name_row = None
            from_row = None
            to_row = None
            
            for idx, row in df.iterrows():
                first_col_value = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                
                if first_col_value == "Stream Name":
                    stream_name_row = idx
                elif first_col_value == "From":
                    from_row = idx
                elif first_col_value == "To":
                    to_row = idx
            
            if not all([stream_name_row is not None, from_row is not None, to_row is not None]):
                logger.debug("Could not find required rows in flowsheet Excel")
                return stream_connections
            
            # 提取流股连接信息
            for col_idx in range(3, len(df.columns)):
                stream_name = str(df.iloc[stream_name_row, col_idx]) if pd.notna(df.iloc[stream_name_row, col_idx]) else None
                from_eq = str(df.iloc[from_row, col_idx]) if pd.notna(df.iloc[from_row, col_idx]) else None
                to_eq = str(df.iloc[to_row, col_idx]) if pd.notna(df.iloc[to_row, col_idx]) else None
                
                # 过滤掉无效数据
                if stream_name and stream_name != 'nan' and stream_name != '':
                    stream_connections[stream_name] = {
                        'from': from_eq if from_eq and from_eq != 'nan' else None,
                        'to': to_eq if to_eq and to_eq != 'nan' else None
                    }
            
            logger.debug(f"Parsed {len(stream_connections)} stream connections")
            
        except Exception as e:
            logger.debug(f"Error parsing Excel connections: {str(e)}")
        
        return stream_connections
    
    def build_equipment_connections(self, stream_connections: Dict[str, Dict[str, Optional[str]]]) -> Dict[str, Dict[str, List[str]]]:
        """
        构建设备连接关系
        
        Args:
            stream_connections: 流股连接信息
            
        Returns:
            设备连接信息字典
        """
        equipment_connections = {}
        
        try:
            # 收集所有设备
            all_equipment = set()
            for stream_info in stream_connections.values():
                if stream_info['from']:
                    all_equipment.add(stream_info['from'])
                if stream_info['to']:
                    all_equipment.add(stream_info['to'])
            
            # 为每个设备创建连接信息
            for equipment in all_equipment:
                equipment_connections[equipment] = {
                    'inlet_streams': [],
                    'outlet_streams': []
                }
            
            # 填充连接信息
            for stream_name, stream_info in stream_connections.items():
                from_eq = stream_info['from']
                to_eq = stream_info['to']
                
                # 对于源设备，这是出料流股
                if from_eq and from_eq in equipment_connections:
                    equipment_connections[from_eq]['outlet_streams'].append(stream_name)
                
                # 对于目标设备，这是进料流股
                if to_eq and to_eq in equipment_connections:
                    equipment_connections[to_eq]['inlet_streams'].append(stream_name)
            
            logger.debug(f"Built connections for {len(equipment_connections)} equipment items")
            
        except Exception as e:
            logger.debug(f"Error building equipment connections: {str(e)}")
        
        return equipment_connections
    
    def get_equipment_stream_connections_from_excel(self, block_name: str) -> Tuple[List[str], List[str]]:
        """
        从Excel连接信息中获取设备的进出料流股
        
        Args:
            block_name: 设备名称
            
        Returns:
            (inlet_streams, outlet_streams): 进料和出料流股列表
        """
        if not hasattr(self, 'flowsheet_connections') or not self.flowsheet_connections:
            # 尝试加载连接信息
            self.load_flowsheet_connections()
        
        if block_name in self.flowsheet_connections:
            connections = self.flowsheet_connections[block_name]
            return connections['inlet_streams'], connections['outlet_streams']
        
        return [], []


class EquipmentSizer:
    """
    Equipment sizing calculations based on process conditions
    
    Implements industry-standard correlations for estimating equipment
    dimensions from process data extracted from Aspen simulations.
    """
    
    def __init__(self):
        # Material properties and design factors
        self.material_properties = {
            MaterialType.CARBON_STEEL: {'density': 7850, 'allowable_stress': 138},  # kg/m3, MPa
            MaterialType.SS304: {'density': 8000, 'allowable_stress': 138},
            MaterialType.SS316: {'density': 8000, 'allowable_stress': 138},
            MaterialType.HASTELLOY_C: {'density': 8890, 'allowable_stress': 207}
        }
        
        # Design safety factors
        self.safety_factors = {
            'pressure': 1.1,     # Design pressure = 1.1 × operating pressure
            'temperature': 1.1,   # Design temperature factor
            'stress': 4.0        # Safety factor for stress calculations
        }
    
    def size_reactor(self, 
                    volumetric_flow: float,
                    residence_time: float, 
                    pressure: float,
                    temperature: float,
                    material: MaterialType = MaterialType.SS316) -> EquipmentSizeData:
        """
        Size a reactor based on residence time requirement
        
        Args:
            volumetric_flow: Volumetric flow rate in m3/hr
            residence_time: Required residence time in hours
            pressure: Operating pressure in bar
            temperature: Operating temperature in °C
            material: Construction material
            
        Returns:
            EquipmentSizeData object with sizing results
        """
        # Calculate reactor volume
        volume = volumetric_flow * residence_time  # m3
        
        # Assume L/D ratio of 4 for packed bed reactor
        l_d_ratio = 4.0
        diameter = (4 * volume / (math.pi * l_d_ratio)) ** (1/3)
        length = diameter * l_d_ratio
        
        # Design conditions with safety factors
        design_pressure = pressure * self.safety_factors['pressure']
        design_temperature = temperature * self.safety_factors['temperature']
        
        # Calculate wall thickness (ASME pressure vessel code)
        wall_thickness = self._calculate_wall_thickness(
            diameter, design_pressure, material
        )
        
        # Determine pressure level
        pressure_level = self._get_pressure_level(design_pressure)
        
        return EquipmentSizeData(
            equipment_type=EquipmentType.REACTOR,
            name="Main Reactor",
            diameter=diameter,
            length=length,
            volume=volume,
            design_pressure=design_pressure,
            design_temperature=design_temperature,
            material=material,
            pressure_level=pressure_level,
            wall_thickness=wall_thickness,
            sizing_basis={
                'volumetric_flow': volumetric_flow,
                'residence_time': residence_time,
                'l_d_ratio': l_d_ratio
            },
            assumptions=[
                f"L/D ratio = {l_d_ratio}",
                "Packed bed reactor configuration",
                f"Pressure safety factor = {self.safety_factors['pressure']}"
            ]
        )
    
    def size_heat_exchanger(self,
                           duty: float,
                           delta_t_lm: float,
                           pressure: float,
                           temperature: float,
                           material: MaterialType = MaterialType.SS304) -> EquipmentSizeData:
        """
        Size a shell-and-tube heat exchanger
        
        Args:
            duty: Heat duty in kW
            delta_t_lm: Log mean temperature difference in °C
            pressure: Design pressure in bar
            temperature: Design temperature in °C
            material: Construction material
            
        Returns:
            EquipmentSizeData object with sizing results
        """
        # Assume overall heat transfer coefficient
        if pressure > 20:
            u_overall = 300  # W/m2-K for high pressure service
        else:
            u_overall = 500  # W/m2-K for moderate pressure
        
        # Calculate required heat transfer area
        area = (duty * 1000) / (u_overall * delta_t_lm)  # m2
        
        # Estimate tube bundle and shell dimensions
        # Assume 25mm OD tubes, triangular pitch
        tube_od = 0.025  # m
        tube_pitch = 1.25 * tube_od
        tube_length = 6.0  # m (standard length)
        
        # Number of tubes based on area
        tube_count = int(area / (math.pi * tube_od * tube_length))
        
        # Shell diameter from tube count (approximation)
        shell_diameter = math.sqrt(tube_count * tube_pitch**2 / 0.785)
        
        # Design conditions
        design_pressure = pressure * self.safety_factors['pressure']
        design_temperature = temperature * self.safety_factors['temperature']
        pressure_level = self._get_pressure_level(design_pressure)
        
        return EquipmentSizeData(
            equipment_type=EquipmentType.HEAT_EXCHANGER,
            name="Shell-Tube Heat Exchanger",
            diameter=shell_diameter,
            length=tube_length,
            area=area,
            design_pressure=design_pressure,
            design_temperature=design_temperature,
            material=material,
            pressure_level=pressure_level,
            tube_count=tube_count,
            sizing_basis={
                'duty': duty,
                'delta_t_lm': delta_t_lm,
                'u_overall': u_overall,
                'tube_od': tube_od,
                'tube_length': tube_length
            },
            assumptions=[
                f"Overall U = {u_overall} W/m2-K",
                f"Tube OD = {tube_od*1000} mm",
                f"Tube length = {tube_length} m"
            ]
        )
    
    def size_compressor(self,
                       volumetric_flow: float,
                       suction_pressure: float,
                       discharge_pressure: float,
                       temperature: float,
                       efficiency: float = 0.75) -> EquipmentSizeData:
        """
        Size a centrifugal compressor
        
        Args:
            volumetric_flow: Volumetric flow rate at suction in m3/hr
            suction_pressure: Suction pressure in bar
            discharge_pressure: Discharge pressure in bar
            temperature: Suction temperature in °C
            efficiency: Isentropic efficiency
            
        Returns:
            EquipmentSizeData object with sizing results
        """
        # Calculate pressure ratio and power requirement
        pressure_ratio = discharge_pressure / suction_pressure
        
        # Determine number of stages (limit pressure ratio per stage)
        max_ratio_per_stage = 3.5
        stages = max(1, int(math.ceil(math.log(pressure_ratio) / math.log(max_ratio_per_stage))))
        
        # Calculate power requirement (assuming ideal gas)
        gamma = 1.3  # Heat capacity ratio for typical gases
        
        # Isentropic power calculation
        power_isentropic = (gamma / (gamma - 1)) * suction_pressure * 100 * volumetric_flow / 3600 * \
                          ((pressure_ratio ** ((gamma - 1) / gamma)) - 1) / 1000  # kW
        
        # Actual power with efficiency
        power_actual = power_isentropic / efficiency
        
        # Estimate impeller diameter (empirical correlation)
        impeller_diameter = 0.5 * (volumetric_flow / 3600) ** 0.5  # m
        
        design_pressure = discharge_pressure * self.safety_factors['pressure']
        pressure_level = self._get_pressure_level(design_pressure)
        
        return EquipmentSizeData(
            equipment_type=EquipmentType.COMPRESSOR,
            name="Centrifugal Compressor",
            diameter=impeller_diameter,
            design_pressure=design_pressure,
            design_temperature=temperature + 50,  # Estimate discharge temperature
            material=MaterialType.CARBON_STEEL,
            pressure_level=pressure_level,
            stages=stages,
            power_rating=power_actual,
            sizing_basis={
                'volumetric_flow': volumetric_flow,
                'pressure_ratio': pressure_ratio,
                'efficiency': efficiency,
                'gamma': gamma
            },
            assumptions=[
                f"Isentropic efficiency = {efficiency}",
                f"Heat capacity ratio = {gamma}",
                f"Max pressure ratio per stage = {max_ratio_per_stage}"
            ]
        )
    
    def size_distillation_column(self,
                                vapor_flow: float,
                                liquid_flow: float,
                                pressure: float,
                                stages: int = 20) -> EquipmentSizeData:
        """
        Size a distillation column
        
        Args:
            vapor_flow: Vapor flow rate in kmol/hr
            liquid_flow: Liquid flow rate in kmol/hr
            pressure: Operating pressure in bar
            stages: Number of theoretical stages
            
        Returns:
            EquipmentSizeData object with sizing results
        """
        # Estimate vapor density (assuming ideal gas, MW = 30)
        mw_avg = 30  # kg/kmol
        vapor_density = (pressure * 100 * mw_avg) / (8.314 * 298)  # kg/m3
        
        # Calculate vapor velocity (Souders-Brown equation)
        c_sb = 0.05  # m/s, conservative value for packed columns
        vapor_velocity = c_sb * math.sqrt((1000 - vapor_density) / vapor_density)
        
        # Column diameter
        vapor_volumetric = vapor_flow * mw_avg / vapor_density / 3600  # m3/s
        diameter = math.sqrt(4 * vapor_volumetric / (math.pi * vapor_velocity))
        
        # Column height (assuming 0.6 m per theoretical stage for packed column)
        height_per_stage = 0.6  # m
        height = stages * height_per_stage
        
        design_pressure = pressure * self.safety_factors['pressure']
        pressure_level = self._get_pressure_level(design_pressure)
        
        return EquipmentSizeData(
            equipment_type=EquipmentType.DISTILLATION_COLUMN,
            name="Distillation Column",
            diameter=diameter,
            height=height,
            design_pressure=design_pressure,
            design_temperature=150,  # Estimated
            material=MaterialType.SS304,
            pressure_level=pressure_level,
            sizing_basis={
                'vapor_flow': vapor_flow,
                'liquid_flow': liquid_flow,
                'stages': stages,
                'vapor_velocity': vapor_velocity
            },
            assumptions=[
                f"Theoretical stages = {stages}",
                f"Height per stage = {height_per_stage} m",
                f"Souders-Brown constant = {c_sb} m/s"
            ]
        )
    
    def _calculate_wall_thickness(self, diameter: float, pressure: float, 
                                 material: MaterialType) -> float:
        """
        Calculate wall thickness for pressure vessel (ASME code)
        
        Args:
            diameter: Internal diameter in m
            pressure: Design pressure in bar
            material: Construction material
            
        Returns:
            Wall thickness in mm
        """
        # Get material allowable stress
        allowable_stress = self.material_properties[material]['allowable_stress']  # MPa
        
        # Convert pressure to MPa
        pressure_mpa = pressure / 10
        
        # ASME formula: t = P*R / (S*E - 0.6*P)
        # Where: P = pressure, R = radius, S = allowable stress, E = efficiency (0.85)
        efficiency = 0.85
        radius = diameter / 2 * 1000  # Convert to mm
        
        thickness = (pressure_mpa * radius) / (allowable_stress * efficiency - 0.6 * pressure_mpa)
        
        # Add corrosion allowance
        corrosion_allowance = 3.0  # mm
        
        return thickness + corrosion_allowance
    
    def _get_pressure_level(self, pressure: float) -> PressureLevel:
        """Determine pressure level classification"""
        if pressure < 10:
            return PressureLevel.LOW
        elif pressure < 50:
            return PressureLevel.MEDIUM
        else:
            return PressureLevel.HIGH


class HeatExchangerDataLoader:
    """
    Load and process heat exchanger data from Excel file
    Enhanced with multi-worksheet support and flexible data extraction
    """
    
    def __init__(self, excel_file: str):
        self.excel_file = excel_file
        self.data = None
        self.processed_data = None
        self.all_worksheets = {}
        self.extraction_log = []
        
    def load_data(self) -> Optional[pd.DataFrame]:
        """Load heat exchanger data from Excel file with multi-worksheet support"""
        try:
            logger.info(f"Loading heat exchanger data from {self.excel_file}")
            
            if not os.path.exists(self.excel_file):
                raise FileNotFoundError(f"Excel file not found: {self.excel_file}")
            
            # Load all worksheets
            all_data = self._load_all_worksheets()
            
            if not all_data:
                raise Exception("No data could be loaded from any worksheet")
            
            # Find and combine heat exchanger data from all worksheets
            combined_data = self._combine_hex_data_from_worksheets(all_data)
            
            if combined_data is not None and not combined_data.empty:
                self.data = combined_data
                logger.info(f"Successfully combined data with shape: {self.data.shape}")
                logger.info(f"Combined columns: {list(self.data.columns)}")
                
                # Process the combined data
                self.processed_data = self._process_hex_data()
                
                return self.data
            else:
                raise Exception("No valid heat exchanger data found in any worksheet")
            
        except Exception as e:
            logger.error(f"Failed to load heat exchanger data: {e}")
            return None
    
    def _load_all_worksheets(self) -> Dict[str, pd.DataFrame]:
        """Load data from all worksheets in the Excel file"""
        all_data = {}
        
        try:
            # Get all worksheet names
            xl_file = pd.ExcelFile(self.excel_file, engine='openpyxl')
            sheet_names = xl_file.sheet_names
            
            logger.info(f"Found {len(sheet_names)} worksheets: {sheet_names}")
            self.extraction_log.append(f"Found worksheets: {sheet_names}")
            
            for sheet_name in sheet_names:
                try:
                    logger.info(f"Loading worksheet: {sheet_name}")
                    
                    # Try multiple loading methods for each worksheet
                    df = None
                    loading_methods = [
                        ("openpyxl", lambda: pd.read_excel(self.excel_file, sheet_name=sheet_name, engine='openpyxl')),
                        ("xlrd", lambda: pd.read_excel(self.excel_file, sheet_name=sheet_name, engine='xlrd'))
                    ]
                    
                    for method_name, method_func in loading_methods:
                        try:
                            df = method_func()
                            logger.info(f"✅ {sheet_name} loaded with {method_name}: {df.shape}")
                            self.extraction_log.append(f"Sheet {sheet_name}: {df.shape[0]}x{df.shape[1]} - {method_name}")
                            break
                        except Exception as e:
                            logger.warning(f"{method_name} failed for {sheet_name}: {e}")
                            continue
                    
                    if df is not None and not df.empty:
                        # Clean column names
                        df.columns = [str(col).strip() for col in df.columns]
                        all_data[sheet_name] = df
                        self.all_worksheets[sheet_name] = df
                        
                        # Log worksheet analysis
                        hex_score = self._evaluate_hex_worksheet(df, sheet_name)
                        logger.info(f"   HEX relevance score for {sheet_name}: {hex_score}/10")
                        
                    else:
                        logger.warning(f"Worksheet {sheet_name} is empty or could not be loaded")
                        self.extraction_log.append(f"Sheet {sheet_name}: EMPTY or FAILED")
                        
                except Exception as e:
                    logger.error(f"Failed to load worksheet {sheet_name}: {e}")
                    self.extraction_log.append(f"Sheet {sheet_name}: ERROR - {str(e)}")
                    continue
            
        except Exception as e:
            logger.error(f"Could not access Excel file worksheets: {e}")
            
            # Fallback: try to load first sheet only
            try:
                logger.info("Falling back to single worksheet loading...")
                df = pd.read_excel(self.excel_file, sheet_name=0, engine='openpyxl')
                all_data["Sheet1"] = df
                self.all_worksheets["Sheet1"] = df
                logger.info(f"Fallback successful: {df.shape}")
                self.extraction_log.append(f"Fallback to Sheet1: {df.shape[0]}x{df.shape[1]}")
            except Exception as fallback_error:
                logger.error(f"Fallback also failed: {fallback_error}")
        
        return all_data
    
    def _evaluate_hex_worksheet(self, df: pd.DataFrame, sheet_name: str) -> int:
        """Evaluate how likely a worksheet contains heat exchanger data (0-10 score)"""
        if df.empty:
            return 0
        
        score = 0
        columns_lower = [str(col).lower() for col in df.columns]
        
        # Heat exchanger keywords (weight: 3)
        hex_keywords = ['heat', 'exchanger', 'hex', 'hx', '换热', '换热器', 'cooler', 'heater', 'condenser']
        for keyword in hex_keywords:
            if any(keyword in col for col in columns_lower):
                score += 3
                break
        
        # Temperature keywords (weight: 2)
        temp_keywords = ['temp', 'temperature', '温度', 'hot', 'cold', '热', '冷']
        temp_count = sum(1 for keyword in temp_keywords if any(keyword in col for col in columns_lower))
        score += min(temp_count, 2)
        
        # Duty/Load keywords (weight: 2)
        duty_keywords = ['duty', 'load', '负荷', 'power', 'kw', 'mw']
        if any(keyword in col for col in columns_lower for keyword in duty_keywords):
            score += 2
        
        # Area keywords (weight: 2)
        area_keywords = ['area', '面积', 'm2', 'm²', 'surface']
        if any(keyword in col for col in columns_lower for keyword in area_keywords):
            score += 2
        
        # Flow keywords (weight: 1)
        flow_keywords = ['flow', '流量', 'mass', 'kg/h', 'stream']
        if any(keyword in col for col in columns_lower for keyword in flow_keywords):
            score += 1
        
        self.extraction_log.append(f"HEX score for {sheet_name}: {score}/10 - Columns: {columns_lower[:5]}")
        return min(score, 10)
    
    def _combine_hex_data_from_worksheets(self, all_data: Dict[str, pd.DataFrame]) -> Optional[pd.DataFrame]:
        """Combine heat exchanger data from multiple worksheets"""
        if not all_data:
            return None
        
        # Evaluate and rank worksheets by HEX relevance
        worksheet_scores = []
        for sheet_name, df in all_data.items():
            score = self._evaluate_hex_worksheet(df, sheet_name)
            worksheet_scores.append((score, sheet_name, df))
        
        # Sort by score (highest first)
        worksheet_scores.sort(key=lambda x: x[0], reverse=True)
        
        logger.info("Worksheet HEX relevance ranking:")
        for score, sheet_name, df in worksheet_scores:
            logger.info(f"  {sheet_name}: {score}/10 ({df.shape[0]} rows, {df.shape[1]} cols)")
        
        # Strategy 1: Use the highest scoring worksheet
        if worksheet_scores[0][0] >= 3:
            best_sheet = worksheet_scores[0]
            logger.info(f"Using best worksheet: {best_sheet[1]} (score: {best_sheet[0]}/10)")
            self.extraction_log.append(f"Selected worksheet: {best_sheet[1]} (score: {best_sheet[0]}/10)")
            return best_sheet[2]
        
        # Strategy 2: Combine data from multiple worksheets if no clear winner
        logger.info("No single worksheet has high HEX relevance, attempting to combine data...")
        
        combined_data_list = []
        for score, sheet_name, df in worksheet_scores:
            if score > 0 and not df.empty:
                # Add sheet identifier
                df_copy = df.copy()
                df_copy['source_worksheet'] = sheet_name
                combined_data_list.append(df_copy)
                logger.info(f"Including {sheet_name} in combined data ({df.shape[0]} rows)")
        
        if combined_data_list:
            # Try to concatenate DataFrames
            try:
                combined_df = pd.concat(combined_data_list, ignore_index=True, sort=False)
                logger.info(f"Successfully combined data: {combined_df.shape}")
                self.extraction_log.append(f"Combined {len(combined_data_list)} worksheets: {combined_df.shape}")
                return combined_df
            except Exception as e:
                logger.warning(f"Failed to combine worksheets: {e}")
                # Fallback to first available worksheet
                logger.info(f"Falling back to first worksheet: {worksheet_scores[0][1]}")
                return worksheet_scores[0][2]
        
        # Strategy 3: Last resort - use first available worksheet
        if worksheet_scores:
            logger.warning("Using first available worksheet as last resort")
            self.extraction_log.append(f"Last resort: using {worksheet_scores[0][1]}")
            return worksheet_scores[0][2]
        
        return None
    
    def _find_column_mappings_flexible(self) -> Dict[str, List[str]]:
        """Enhanced flexible column mapping with multiple matching strategies"""
        if self.data is None:
            return {}
        
        mappings = {
            'equipment_name': [],
            'duty': [],
            'area': [],
            'temperature': [],
            'pressure': [],
            'hot_stream_name': [],
            'cold_stream_name': [],
            'hot_inlet_temp': [],
            'hot_outlet_temp': [],
            'cold_inlet_temp': [],
            'cold_outlet_temp': [],
            'hot_flow': [],
            'cold_flow': [],
            'hot_composition': [],
            'cold_composition': [],
            'generic_flow': [],
            'generic_stream': []
        }
        
        # Enhanced keyword patterns with Chinese support and variations
        keyword_patterns = {
            'equipment_name': [
                # English
                'name', 'id', 'tag', 'equipment', 'hex', 'exchanger', 'unit', 'no', 'number',
                # Chinese
                '名称', '设备', '换热器', '编号', '序号', 'HEX', 'ID'
            ],
            'duty': [
                # English
                'duty', 'load', 'heat', 'power', 'thermal', 'energy', 'kw', 'mw', 'btu', 'kcal',
                'q', 'q_duty', 'heat_duty', 'thermal_load',
                # Chinese
                '负荷', '热负荷', '功率', '热量', '能量', '热功率'
            ],
            'area': [
                # English
                'area', 'surface', 'heat_area', 'transfer_area', 'm2', 'm²', 'ft2', 'ft²',
                # Chinese  
                '面积', '换热面积', '传热面积', '表面积'
            ],
            'temperature': [
                # English
                'temp', 'temperature', 'deg', 'celsius', 'fahrenheit', '°c', '°f',
                # Chinese
                '温度', '度'
            ],
            'pressure': [
                # English
                'press', 'pressure', 'bar', 'psi', 'pa', 'mpa', 'kpa', 'atm',
                # Chinese
                '压力', '压强'
            ],
            'hot_stream_name': [
                # English
                'hot', 'shell', 'hot_stream', 'hot_side', 'hot_fluid', 'process', 
                'hot_name', 'shell_name', 'hot_stream_name',
                # Chinese
                '热', '热流', '壳程', '热侧', '热介质'
            ],
            'cold_stream_name': [
                # English
                'cold', 'tube', 'cold_stream', 'cold_side', 'cold_fluid', 'utility',
                'cold_name', 'tube_name', 'cold_stream_name',
                # Chinese
                '冷', '冷流', '管程', '冷侧', '冷介质'
            ],
            'hot_inlet_temp': [
                # English
                'hot_in', 'hot_inlet', 'shell_in', 'shell_inlet', 'h_in', 'hot_temp_in',
                'hot_in_temp', 'shell_in_temp', 'hot_inlet_temperature', 'hot t in',
                # Chinese
                '热进', '热入口', '壳程进口', '热侧进口'
            ],
            'hot_outlet_temp': [
                # English
                'hot_out', 'hot_outlet', 'shell_out', 'shell_outlet', 'h_out', 'hot_temp_out',
                'hot_out_temp', 'shell_out_temp', 'hot_outlet_temperature', 'hot t out',
                # Chinese
                '热出', '热出口', '壳程出口', '热侧出口'
            ],
            'cold_inlet_temp': [
                # English
                'cold_in', 'cold_inlet', 'tube_in', 'tube_inlet', 'c_in', 'cold_temp_in',
                'cold_in_temp', 'tube_in_temp', 'cold_inlet_temperature', 'cold t in',
                # Chinese
                '冷进', '冷入口', '管程进口', '冷侧进口'
            ],
            'cold_outlet_temp': [
                # English
                'cold_out', 'cold_outlet', 'tube_out', 'tube_outlet', 'c_out', 'cold_temp_out',
                'cold_out_temp', 'tube_out_temp', 'cold_outlet_temperature', 'cold t out',
                # Chinese
                '冷出', '冷出口', '管程出口', '冷侧出口'
            ],
            'hot_flow': [
                # English
                'hot_flow', 'shell_flow', 'hot_mass', 'hot_mass_flow', 'hot_molar',
                'hot_flow_rate', 'shell_flow_rate', 'process_flow',
                # Chinese
                '热流量', '壳程流量', '热侧流量'
            ],
            'cold_flow': [
                # English  
                'cold_flow', 'tube_flow', 'cold_mass', 'cold_mass_flow', 'cold_molar',
                'cold_flow_rate', 'tube_flow_rate', 'utility_flow',
                # Chinese
                '冷流量', '管程流量', '冷侧流量'
            ],
            'hot_composition': [
                # English
                'hot_comp', 'shell_comp', 'hot_composition', 'hot_components',
                # Chinese
                '热流组分', '壳程组分'
            ],
            'cold_composition': [
                # English
                'cold_comp', 'tube_comp', 'cold_composition', 'cold_components', 
                # Chinese
                '冷流组分', '管程组分'
            ],
            'generic_flow': [
                # English
                'flow', 'mass', 'molar', 'kg/h', 'kmol/h', 'm3/h', 'rate', 'flowrate',
                # Chinese
                '流量', '质量', '摩尔', '速率'
            ],
            'generic_stream': [
                # English
                'stream', 'fluid', 'medium', 'side',
                # Chinese  
                '流股', '介质', '流体', '侧'
            ]
        }
        
        columns = [str(col) for col in self.data.columns]
        
        # Strategy 1: Exact keyword matching (case-insensitive)
        for category, keywords in keyword_patterns.items():
            for col in columns:
                col_lower = col.lower().strip()
                for keyword in keywords:
                    if keyword.lower() in col_lower:
                        if col not in mappings[category]:
                            mappings[category].append(col)
        
        # Strategy 2: Partial matching for complex column names
        for col in columns:
            col_lower = col.lower().strip()
            col_parts = col_lower.replace('_', ' ').replace('-', ' ').split()
            
            # Check multi-word patterns
            for category, keywords in keyword_patterns.items():
                for keyword in keywords:
                    keyword_parts = keyword.lower().split()
                    if len(keyword_parts) > 1:
                        # Multi-word keyword matching
                        if all(part in col_lower for part in keyword_parts):
                            if col not in mappings[category]:
                                mappings[category].append(col)
                    else:
                        # Single word in multi-part column name
                        if keyword.lower() in col_parts:
                            if col not in mappings[category]:
                                mappings[category].append(col)
        
        # Strategy 3: Pattern-based inference for unmatched columns
        unmatched_columns = [col for col in columns if not any(col in col_list for col_list in mappings.values())]
        
        for col in unmatched_columns:
            col_lower = col.lower().strip()
            
            # Infer based on patterns
            if any(char.isdigit() for char in col_lower):
                # Contains numbers - likely temperature, pressure, or flow
                if 'temp' in col_lower or '°' in col_lower or 'deg' in col_lower:
                    mappings['temperature'].append(col)
                elif 'bar' in col_lower or 'psi' in col_lower or 'pa' in col_lower:
                    mappings['pressure'].append(col)
                elif 'kg' in col_lower or 'flow' in col_lower or 'rate' in col_lower:
                    mappings['generic_flow'].append(col)
            
            # Pattern-based categorization
            if len(col_lower) <= 5 and any(char.isdigit() for char in col_lower):
                # Short columns with numbers (likely equipment names)
                mappings['equipment_name'].append(col)
        
        # Strategy 4: Remove duplicates and sort by relevance
        for category in mappings:
            mappings[category] = list(dict.fromkeys(mappings[category]))  # Remove duplicates while preserving order
        
        # Log mapping results
        total_mapped = sum(len(cols) for cols in mappings.values())
        total_columns = len(columns)
        
        self.extraction_log.append(f"Column mapping: {total_mapped}/{total_columns} columns mapped")
        
        # Add unmapped columns to log for debugging
        mapped_columns = set()
        for col_list in mappings.values():
            mapped_columns.update(col_list)
        unmapped = [col for col in columns if col not in mapped_columns]
        if unmapped:
            self.extraction_log.append(f"Unmapped columns: {unmapped}")
            logger.warning(f"⚠️ Unmapped columns: {unmapped[:5]}{'...' if len(unmapped) > 5 else ''}")
        
        return mappings
    
    def _safe_numeric_conversion(self, value, column_name: str) -> Optional[float]:
        """Safely convert value to numeric with enhanced error handling"""
        if pd.isna(value):
            return None
        
        # If already numeric
        if isinstance(value, (int, float)) and not np.isnan(value):
            return float(value)
        
        # If string, try to extract numeric value
        if isinstance(value, str):
            # Clean string
            cleaned = str(value).strip()
            if not cleaned:
                return None
            
            # Try direct conversion first
            try:
                return float(cleaned)
            except ValueError:
                pass
            
            # Extract numeric parts from string with units
            import re
            # Match numbers (including scientific notation)
            numeric_pattern = r'[-+]?(?:\d*\.?\d+)(?:[eE][-+]?\d+)?'
            matches = re.findall(numeric_pattern, cleaned)
            
            if matches:
                try:
                    return float(matches[0])
                except ValueError:
                    pass
        
        # Log conversion failure
        self.extraction_log.append(f"Failed to convert '{value}' in column '{column_name}' to numeric")
        return None
    
    def _convert_duty_to_kw(self, value: float, column_name: str) -> float:
        """Convert duty value to kW based on column name hints"""
        if value == 0:
            return 0.0
        
        column_lower = column_name.lower()
        
        # Unit conversions to kW
        if any(unit in column_lower for unit in ['kj/h', 'kj/hr']):
            return value / 3600  # kJ/h to kW
        elif any(unit in column_lower for unit in ['mj/h', 'mj/hr']):
            return value * 1000 / 3600  # MJ/h to kW
        elif any(unit in column_lower for unit in ['j/h', 'j/hr']):
            return value / 3600000  # J/h to kW
        elif any(unit in column_lower for unit in ['btu/h', 'btu/hr']):
            return value * 0.000293071  # BTU/h to kW
        elif any(unit in column_lower for unit in ['kcal/h', 'kcal/hr']):
            return value * 0.001163  # kcal/h to kW
        elif any(unit in column_lower for unit in ['mw', 'megawatt']):
            return value * 1000  # MW to kW
        elif any(unit in column_lower for unit in ['w', 'watt']) and 'kw' not in column_lower:
            return value / 1000  # W to kW
        else:
            # Default assumption: already in kW
            return abs(value)
    
    def _convert_area_to_m2(self, value: float, column_name: str) -> float:
        """Convert area value to m² based on column name hints"""
        if value == 0:
            return 0.0
        
        column_lower = column_name.lower()
        
        # Unit conversions to m²
        if any(unit in column_lower for unit in ['ft2', 'ft²', 'sq_ft', 'sqft']):
            return value * 0.092903  # ft² to m²
        elif any(unit in column_lower for unit in ['in2', 'in²', 'sq_in', 'sqin']):
            return value * 0.00064516  # in² to m²
        elif any(unit in column_lower for unit in ['cm2', 'cm²']):
            return value / 10000  # cm² to m²
        elif any(unit in column_lower for unit in ['mm2', 'mm²']):
            return value / 1000000  # mm² to m²
        else:
            # Default assumption: already in m²
            return abs(value)
    
    def _process_hex_data(self) -> Dict[str, Any]:
        """Process heat exchanger data for better integration with TEA calculations"""
        if self.data is None:
            return {}
        
        processed = {
            'equipment_list': [],
            'total_heat_duty': 0.0,
            'total_heat_area': 0.0,
            'hex_count': 0,
            'temperature_ranges': {},
            'pressure_levels': {}
        }
        
        try:
            # Enhanced flexible column matching
            column_mappings = self._find_column_mappings_flexible()
            
            # Extract mapped columns for easier access
            duty_cols = column_mappings.get('duty', [])
            area_cols = column_mappings.get('area', [])
            temp_cols = column_mappings.get('temperature', [])
            pres_cols = column_mappings.get('pressure', [])
            name_cols = column_mappings.get('equipment_name', [])
            
            hot_stream_name_cols = column_mappings.get('hot_stream_name', [])
            cold_stream_name_cols = column_mappings.get('cold_stream_name', [])
            
            hot_temp_in_cols = column_mappings.get('hot_inlet_temp', [])
            hot_temp_out_cols = column_mappings.get('hot_outlet_temp', [])
            cold_temp_in_cols = column_mappings.get('cold_inlet_temp', [])
            cold_temp_out_cols = column_mappings.get('cold_outlet_temp', [])
            
            hot_flow_cols = column_mappings.get('hot_flow', [])
            cold_flow_cols = column_mappings.get('cold_flow', [])
            
            hot_flow_cols = column_mappings.get('hot_flow', [])
            cold_flow_cols = column_mappings.get('cold_flow', [])
            
            # Log discovered column mappings
            logger.info("🔍 Enhanced Column Detection Results:")
            for category, columns in column_mappings.items():
                if columns:
                    logger.info(f"   {category}: {columns}")
            
            # Log extraction statistics
            self.extraction_log.append(f"Column mappings found: {sum(len(cols) for cols in column_mappings.values())} total")
            self.extraction_log.append(f"Key mappings - Duty: {len(duty_cols)}, Area: {len(area_cols)}, Temp: {len(temp_cols)}")
            
            logger.info(f"📊 Processing {len(self.data)} rows with enhanced column detection...")
            
            # Process each row as a heat exchanger
            for idx, row in self.data.iterrows():
                hex_info = {
                    'index': idx,
                    'name': f"HEX-{idx:03d}",
                    'duty': 0.0,
                    'area': 0.0,
                    # Simplified: Direct stream mapping
                    'hot_stream_name': None,    # Will be mapped to inlet_streams
                    'cold_stream_name': None,   # Will be mapped to outlet_streams
                    'hot_stream_inlet_temp': None,
                    'hot_stream_outlet_temp': None,
                    'cold_stream_inlet_temp': None,
                    'cold_stream_outlet_temp': None,
                    'temperatures': [],
                    'pressures': []
                }
                
                # Enhanced data extraction with robust conversion
                extraction_success = False
                extraction_warnings = []
                
                # Extract name with multiple column fallback
                name_extracted = False
                for name_col in name_cols:
                    if name_col in row.index and pd.notna(row[name_col]):
                        hex_info['name'] = str(row[name_col]).strip()
                        name_extracted = True
                        break
                
                if not name_extracted and idx < 999:
                    # Use source worksheet info if available
                    worksheet_suffix = f"-{row.get('source_worksheet', 'UNK')}" if 'source_worksheet' in row else ""
                    hex_info['name'] = f"HEX-{idx:03d}{worksheet_suffix}"
                
                # Enhanced duty extraction with unit conversion
                duty_extracted = False
                for duty_col in duty_cols:
                    if duty_col in row.index:
                        duty_val = self._safe_numeric_conversion(row[duty_col], duty_col)
                        if duty_val is not None:
                            # Smart unit conversion
                            duty_kw = self._convert_duty_to_kw(duty_val, duty_col)
                            hex_info['duty'] = duty_kw
                            processed['total_heat_duty'] += abs(duty_kw)
                            duty_extracted = True
                            extraction_success = True
                            break
                
                # Enhanced area extraction
                area_extracted = False
                for area_col in area_cols:
                    if area_col in row.index:
                        area_val = self._safe_numeric_conversion(row[area_col], area_col)
                        if area_val is not None:
                            # Convert area units if needed
                            area_m2 = self._convert_area_to_m2(area_val, area_col)
                            hex_info['area'] = area_m2
                            processed['total_heat_area'] += area_m2
                            area_extracted = True
                            extraction_success = True
                            break
                
                # Extract temperatures
                for temp_col in temp_cols:
                    temp_val = row[temp_col]
                    if pd.notna(temp_val) and isinstance(temp_val, (int, float)):
                        hex_info['temperatures'][temp_col] = float(temp_val)
                
                # Extract pressures
                for pres_col in pres_cols:
                    pres_val = row[pres_col]
                    if pd.notna(pres_val) and isinstance(pres_val, (int, float)):
                        hex_info['pressures'][pres_col] = float(pres_val)
                
                # Enhanced: Extract hot stream data with better column selection
                if hot_stream_name_cols:
                    # 优先选择包含"stream"的列名，然后是包含"hot"的列名
                    best_hot_col = None
                    for col in hot_stream_name_cols:
                        if 'stream' in col.lower():
                            best_hot_col = col
                            break
                    if not best_hot_col:
                        for col in hot_stream_name_cols:
                            if 'hot' in col.lower():
                                best_hot_col = col
                                break
                    if not best_hot_col:
                        best_hot_col = hot_stream_name_cols[0]
                    
                    hot_name_val = row[best_hot_col]
                    if pd.notna(hot_name_val):
                        hex_info['hot_stream_name'] = str(hot_name_val)
                
                if hot_temp_in_cols:
                    hot_temp_in_val = row[hot_temp_in_cols[0]]
                    if pd.notna(hot_temp_in_val) and isinstance(hot_temp_in_val, (int, float)):
                        hex_info['hot_stream_inlet_temp'] = float(hot_temp_in_val)
                
                if hot_temp_out_cols:
                    hot_temp_out_val = row[hot_temp_out_cols[0]]
                    if pd.notna(hot_temp_out_val) and isinstance(hot_temp_out_val, (int, float)):
                        hex_info['hot_stream_outlet_temp'] = float(hot_temp_out_val)
                
                if hot_flow_cols:
                    hot_flow_val = row[hot_flow_cols[0]]
                    if pd.notna(hot_flow_val) and isinstance(hot_flow_val, (int, float)):
                        hex_info['hot_stream_flow_rate'] = float(hot_flow_val)
                
                # Enhanced: Extract cold stream data with better column selection
                if cold_stream_name_cols:
                    # 优先选择包含"stream"的列名，然后是包含"cold"的列名
                    best_cold_col = None
                    for col in cold_stream_name_cols:
                        if 'stream' in col.lower():
                            best_cold_col = col
                            break
                    if not best_cold_col:
                        for col in cold_stream_name_cols:
                            if 'cold' in col.lower():
                                best_cold_col = col
                                break
                    if not best_cold_col:
                        best_cold_col = cold_stream_name_cols[0]
                    
                    cold_name_val = row[best_cold_col]
                    if pd.notna(cold_name_val):
                        hex_info['cold_stream_name'] = str(cold_name_val)
                
                if cold_temp_in_cols:
                    cold_temp_in_val = row[cold_temp_in_cols[0]]
                    if pd.notna(cold_temp_in_val) and isinstance(cold_temp_in_val, (int, float)):
                        hex_info['cold_stream_inlet_temp'] = float(cold_temp_in_val)
                
                if cold_temp_out_cols:
                    cold_temp_out_val = row[cold_temp_out_cols[0]]
                    if pd.notna(cold_temp_out_val) and isinstance(cold_temp_out_val, (int, float)):
                        hex_info['cold_stream_outlet_temp'] = float(cold_temp_out_val)
                
                if cold_flow_cols:
                    cold_flow_val = row[cold_flow_cols[0]]
                    if pd.notna(cold_flow_val) and isinstance(cold_flow_val, (int, float)):
                        hex_info['cold_stream_flow_rate'] = float(cold_flow_val)
                
                # Store stream connections for database
                inlet_streams = []
                outlet_streams = []
                
                if hex_info['hot_stream_name']:
                    inlet_streams.append(hex_info['hot_stream_name'])
                if hex_info['cold_stream_name']:
                    outlet_streams.append(hex_info['cold_stream_name'])
                
                hex_info['inlet_streams'] = inlet_streams
                hex_info['outlet_streams'] = outlet_streams
                
                # Enhanced: Validate temperature and flow consistency (non-blocking) 
                validation_warnings = self._validate_hex_data(hex_info)
                if validation_warnings:
                    hex_info['validation_warnings'] = validation_warnings
                    # Don't log as warning, just info for now
                    logger.debug(f"Heat exchanger {hex_info['name']} validation notes: {validation_warnings}")
                
                # 🔥 RELAXED DATA FILTERING - Include more data rows
                # Count what data we actually extracted
                data_indicators = []
                
                if hex_info['duty'] != 0.0:
                    data_indicators.append(f"duty={hex_info['duty']:.1f}")
                if hex_info['area'] != 0.0:
                    data_indicators.append(f"area={hex_info['area']:.1f}")
                if hex_info['hot_stream_name']:
                    data_indicators.append("hot_stream")
                if hex_info['cold_stream_name']:
                    data_indicators.append("cold_stream")
                if hex_info['hot_stream_inlet_temp'] is not None:
                    data_indicators.append("hot_temp")
                if hex_info['cold_stream_inlet_temp'] is not None:
                    data_indicators.append("cold_temp")
                
                # Enhanced inclusion criteria (much more permissive)
                should_include = (
                    # Any basic heat exchanger data
                    extraction_success or
                    # Any stream names
                    hex_info['hot_stream_name'] or hex_info['cold_stream_name'] or
                    # Any temperature values
                    hex_info['hot_stream_inlet_temp'] is not None or
                    hex_info['cold_stream_inlet_temp'] is not None or
                    # Any non-empty data
                    len(data_indicators) > 0 or
                    # Row has at least some non-null values (fallback)
                    len([v for v in row.values if pd.notna(v) and str(v).strip()]) >= 2
                )
                
                if should_include:
                    hex_info['data_quality'] = 'extracted' if extraction_success else 'partial'
                    hex_info['data_indicators'] = data_indicators
                    hex_info['extraction_warnings'] = extraction_warnings
                    
                    processed['equipment_list'].append(hex_info)
                    processed['hex_count'] += 1
                    
                    # Log what we found
                    indicators_str = ', '.join(data_indicators) if data_indicators else 'basic_row_data'
                    logger.debug(f"✅ Included {hex_info['name']}: {indicators_str}")
                else:
                    # Log what we're skipping and why
                    non_null_count = len([v for v in row.values if pd.notna(v) and str(v).strip()])
                    logger.debug(f"⚠️ Skipped row {idx}: only {non_null_count} non-null values, no recognizable HEX data")
            
            logger.info(f"Processed {processed['hex_count']} heat exchangers")
            logger.info(f"Total heat duty: {processed['total_heat_duty']:,.0f} kW")
            logger.info(f"Total heat area: {processed['total_heat_area']:,.0f} m²")
            
        except Exception as e:
            logger.error(f"Error processing heat exchanger data: {e}")
        
        return processed
    
    def _validate_hex_data(self, hex_info: Dict[str, Any]) -> List[str]:
        """Validate heat exchanger data for temperature and flow consistency"""
        warnings = []
        
        # Temperature consistency checks
        hot_inlet = hex_info.get('hot_stream_inlet_temp')
        hot_outlet = hex_info.get('hot_stream_outlet_temp')
        cold_inlet = hex_info.get('cold_stream_inlet_temp')
        cold_outlet = hex_info.get('cold_stream_outlet_temp')
        
        # Hot stream should cool down (inlet > outlet)
        if hot_inlet is not None and hot_outlet is not None:
            if hot_inlet <= hot_outlet:
                warnings.append(f"Hot stream inlet temp ({hot_inlet}°C) should be > outlet temp ({hot_outlet}°C)")
        
        # Cold stream should heat up (outlet > inlet)
        if cold_inlet is not None and cold_outlet is not None:
            if cold_outlet <= cold_inlet:
                warnings.append(f"Cold stream outlet temp ({cold_outlet}°C) should be > inlet temp ({cold_inlet}°C)")
        
        # Heat transfer feasibility (hot side should be hotter)
        if (hot_inlet is not None and cold_outlet is not None and 
            hot_inlet <= cold_outlet):
            warnings.append(f"Hot inlet ({hot_inlet}°C) should be > cold outlet ({cold_outlet}°C) for heat transfer")
        
        if (hot_outlet is not None and cold_inlet is not None and 
            hot_outlet <= cold_inlet):
            warnings.append(f"Hot outlet ({hot_outlet}°C) should be > cold inlet ({cold_inlet}°C) for heat transfer")
        
        # Temperature range checks
        if hot_inlet is not None and (hot_inlet < -50 or hot_inlet > 1000):
            warnings.append(f"Hot inlet temperature ({hot_inlet}°C) seems unrealistic")
        
        if cold_inlet is not None and (cold_inlet < -100 or cold_inlet > 500):
            warnings.append(f"Cold inlet temperature ({cold_inlet}°C) seems unrealistic")
        
        # Flow rate checks
        hot_flow = hex_info.get('hot_stream_flow_rate')
        cold_flow = hex_info.get('cold_stream_flow_rate')
        
        if hot_flow is not None and hot_flow <= 0:
            warnings.append(f"Hot stream flow rate ({hot_flow}) should be positive")
        
        if cold_flow is not None and cold_flow <= 0:
            warnings.append(f"Cold stream flow rate ({cold_flow}) should be positive")
        
        # Duty and area consistency
        duty = hex_info.get('duty', 0)
        area = hex_info.get('area', 0)
        
        if duty > 0 and area <= 0:
            warnings.append("Heat duty specified but no heat transfer area")
        
        if area > 0 and duty <= 0:
            warnings.append("Heat transfer area specified but no heat duty")
        
        return warnings
    
    def get_summary(self) -> Dict[str, Any]:
        """Get comprehensive summary of heat exchanger data"""
        if self.data is None:
            return {}
            
        summary = {
            'total_heat_exchangers': len(self.data),
            'columns': list(self.data.columns),
            'sample_data': self.data.head().to_dict() if not self.data.empty else {},
            'data_types': self.data.dtypes.to_dict()
        }
        
        # Check for heat exchanger relevant columns
        hex_keywords = ['heat', 'duty', 'area', 'temperature', 'pressure', 'load', 'exchanger', 'hex']
        relevant_cols = []
        for col in self.data.columns:
            if any(keyword.lower() in str(col).lower() for keyword in hex_keywords):
                relevant_cols.append(col)
        
        summary['relevant_columns'] = relevant_cols
        
        # Add processed data summary
        if self.processed_data:
            summary['processed_summary'] = {
                'processed_hex_count': self.processed_data['hex_count'],
                'total_heat_duty_kW': self.processed_data['total_heat_duty'],
                'total_heat_area_m2': self.processed_data['total_heat_area']
            }
        
        return summary
    
    def get_heat_exchanger_data_for_tea(self) -> Dict[str, Any]:
        """Get heat exchanger data formatted for TEA calculations"""
        if self.processed_data is None:
            return {}
        
        return {
            'heat_exchangers': self.processed_data['equipment_list'],
            'total_heat_duty_kW': self.processed_data['total_heat_duty'],
            'total_heat_area_m2': self.processed_data['total_heat_area'],
            'hex_count': self.processed_data['hex_count'],
            'average_hex_size_m2': self.processed_data['total_heat_area'] / max(1, self.processed_data['hex_count'])
        }


class AspenDataExtractor:
    """
    Main class for extracting and processing Aspen simulation data
    
    Combines proven COM interface methods with heat exchanger data integration.
    Based on successful patterns from bfg_co2h_aspen_analyzer.py and bfg_co2h_pure_simulation_analyzer.py.
    """
    
    def __init__(self, config_file: Optional[str] = None, db_path: str = "aspen_data.db"):
        self.com_interface = AspenCOMInterface()
        self.equipment_sizer = EquipmentSizer()
        self.config = self._load_configuration(config_file)
        self.hex_loader = None
        self.stream_data = {}
        self.block_data = {}
        self.equipment_connections = {}  # Store stream connections separately
        
        # Initialize stream classifier if available
        if STREAM_CLASSIFICATION_AVAILABLE:
            self.stream_classifier = StreamClassifier()
            logger.info("✅ Stream classifier initialized")
        else:
            self.stream_classifier = None
            logger.warning("⚠️ Stream classifier not available")
        
        # Initialize equipment matcher if available
        if EQUIPMENT_MATCHER_AVAILABLE:
            try:
                self.equipment_matcher = EquipmentModelMatcher()
                logger.info("✅ Equipment matcher initialized")
            except Exception as e:
                self.equipment_matcher = None
                logger.warning(f"⚠️ Failed to initialize equipment matcher: {e}")
        else:
            self.equipment_matcher = None
            logger.warning("⚠️ Equipment matcher not available")
            
        # Initialize enhanced equipment detector if available
        if ENHANCED_EQUIPMENT_DETECTOR_AVAILABLE:
            try:
                self.equipment_detector = EnhancedEquipmentDetector()
                logger.info("✅ Enhanced equipment detector initialized")
            except Exception as e:
                self.equipment_detector = None
                logger.warning(f"⚠️ Failed to initialize equipment detector: {e}")
        else:
            self.equipment_detector = None
        
        # Initialize database for data persistence
        try:
            self.database = AspenDataDatabase(db_path)
            logger.info(f"✅ Database initialized: {db_path}")
        except Exception as e:
            logger.warning(f"⚠️ Could not initialize database: {e}")
            self.database = None
        
    def _load_configuration(self, config_file: Optional[str]) -> Dict[str, Any]:
        """Load configuration from file or use defaults"""
        default_config = {
            'stream_mappings': {
                'feed': 'FEED',
                'product': 'PRODUCT', 
                'recycle': 'RECYCLE'
            },
            'block_mappings': {
                'reactor': 'R-101',
                'compressor': 'K-101',
                'heat_exchanger': 'E-101'
            },
            'equipment_defaults': {
                'reactor_residence_time': 2.0,  # hours
                'heat_exchanger_delta_t': 25.0,  # °C
                'compressor_efficiency': 0.75
            }
        }
        
        if config_file and os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
                # Merge with defaults
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
            except Exception as e:
                logger.warning(f"Could not load config file {config_file}: {str(e)}")
        
        return default_config
    
    # Equipment type detection mapping - simplified dictionary approach
    EQUIPMENT_NAME_PATTERNS = {
        EquipmentType.REACTOR: ['reactor', 'rxn', 'rct', 'react', '反应器'],
        EquipmentType.COMPRESSOR: ['compressor', 'comp', 'blower', 'fan', '压缩机', '风机'],
        EquipmentType.PUMP: ['pump', 'p-', '泵'],
        EquipmentType.DISTILLATION_COLUMN: ['column', 'tower', 'col', 't-', 'distil', 'absorb', 'strip', '塔', '蒸馏'],
        EquipmentType.HEAT_EXCHANGER: ['exchanger', 'hx', 'cooler', 'heater', 'condenser', 'reboiler', 'hex', 'e-', 'h-', 'c-'],
        EquipmentType.SEPARATOR: ['separator', 'sep', 'drum', 'vessel', 's-', 'd-', 'v-', 'flash', 'flash-', 'flsh', '分离器', '储罐', '容器'],
        EquipmentType.TANK: ['tank', 'storage', 'tk-', '储罐', '存储'],
        EquipmentType.VALVE: ['valve', 'control', 'regulator', 'throttle', 'v-', 'split', 'mix', 'splitter', 'mixer', 'fsplit', 'tee', 'junction']
    }
    
    def _detect_equipment_type_from_name(self, aspen_name: str, logical_name: str = None) -> Optional[EquipmentType]:
        """Simplified equipment type detection using dictionary mapping"""
        search_text = f"{aspen_name.lower()} {logical_name.lower() if logical_name else ''}"
        
        # Check each equipment type's patterns
        for equipment_type, keywords in self.EQUIPMENT_NAME_PATTERNS.items():
            if any(keyword in search_text for keyword in keywords):
                return equipment_type
        
        return None
    
    def extract_complete_data(self, aspen_file: str) -> AspenProcessData:
        """
        Extract complete process data from Aspen simulation
        
        Args:
            aspen_file: Path to Aspen simulation file
            
        Returns:
            AspenProcessData object with complete simulation data
        """
        logger.info(f"Starting data extraction from {aspen_file}")
        
        # Connect to Aspen
        if not self.com_interface.connect(aspen_file):
            raise AspenConnectionError(f"Could not connect to Aspen file: {aspen_file}")
        
        try:
            # Extract stream data
            streams = self._extract_stream_data()
            
            # Extract unit operation data
            units = self._extract_unit_operation_data()
            
            # Extract utility data
            utilities = self._extract_utility_data()
            
            # Package results
            process_data = AspenProcessData(
                simulation_name=Path(aspen_file).stem,
                timestamp=datetime.now(),
                streams=streams,
                units=units,
                utilities=utilities,
                global_parameters=self._extract_global_parameters()
            )
            
            # Integrate heat exchanger data if available
            if self.hex_loader and self.hex_loader.data is not None:
                process_data = self.integrate_hex_with_aspen_data(process_data)
                logger.info("Heat exchanger data integrated with Aspen data")
            
            logger.info("Data extraction completed successfully")
            return process_data
            
        finally:
            self.com_interface.disconnect()
    
    def load_hex_data(self, excel_file: str) -> bool:
        """Load heat exchanger data from Excel file"""
        try:
            self.hex_loader = HeatExchangerDataLoader(excel_file)
            data = self.hex_loader.load_data()
            
            if data is not None:
                logger.info("✅ Heat exchanger data loaded successfully")
                return True
            else:
                logger.error("❌ Failed to load heat exchanger data")
                return False
                
        except Exception as e:
            logger.error(f"Error loading heat exchanger data: {str(e)}")
            return False
    
    def get_hex_summary(self) -> Dict[str, Any]:
        """Get heat exchanger data summary"""
        if self.hex_loader:
            return self.hex_loader.get_summary()
        return {}
    
    def get_hex_data_for_tea(self) -> Dict[str, Any]:
        """Get heat exchanger data formatted for TEA calculations"""
        if self.hex_loader:
            return self.hex_loader.get_heat_exchanger_data_for_tea()
    
    def get_hex_extraction_report(self) -> Dict[str, Any]:
        """Get detailed heat exchanger extraction report for diagnosis"""
        if not self.hex_loader:
            return {"error": "No heat exchanger data loaded"}
        
        report = {
            "extraction_timestamp": datetime.now().isoformat(),
            "file_path": self.hex_loader.excel_file,
            "extraction_log": self.hex_loader.extraction_log.copy(),
            "worksheets_analyzed": len(self.hex_loader.all_worksheets),
            "total_data_extracted": 0,
            "data_quality_breakdown": {},
            "column_mapping_success": {},
            "recommendations": []
        }
        
        if self.hex_loader.processed_data:
            processed = self.hex_loader.processed_data
            report["total_data_extracted"] = processed.get('hex_count', 0)
            report["total_heat_duty_kw"] = processed.get('total_heat_duty', 0.0)
            report["total_heat_area_m2"] = processed.get('total_heat_area', 0.0)
            
            # Analyze data quality
            quality_counts = {}
            for hex_item in processed.get('equipment_list', []):
                quality = hex_item.get('data_quality', 'unknown')
                quality_counts[quality] = quality_counts.get(quality, 0) + 1
            
            report["data_quality_breakdown"] = quality_counts
            
            # Analyze extraction success by data type
            data_type_success = {
                'duty_extracted': 0,
                'area_extracted': 0,
                'temperatures_extracted': 0,
                'stream_names_extracted': 0,
                'hot_temps_extracted': 0,
                'cold_temps_extracted': 0
            }
            
            for hex_item in processed.get('equipment_list', []):
                if hex_item.get('duty', 0) != 0:
                    data_type_success['duty_extracted'] += 1
                if hex_item.get('area', 0) != 0:
                    data_type_success['area_extracted'] += 1
                if hex_item.get('temperatures'):
                    data_type_success['temperatures_extracted'] += 1
                if hex_item.get('hot_stream_name') or hex_item.get('cold_stream_name'):
                    data_type_success['stream_names_extracted'] += 1
                if hex_item.get('hot_stream_inlet_temp') is not None:
                    data_type_success['hot_temps_extracted'] += 1
                if hex_item.get('cold_stream_inlet_temp') is not None:
                    data_type_success['cold_temps_extracted'] += 1
            
            report["extraction_success_by_type"] = data_type_success
        
        # Generate recommendations
        recommendations = []
        
        if report["total_data_extracted"] == 0:
            recommendations.append("No heat exchanger data was extracted. Check column names and data format.")
        elif report["total_data_extracted"] < 5:
            recommendations.append("Very few heat exchangers extracted. Consider relaxing filtering criteria.")
        
        extraction_log = report.get("extraction_log", [])
        if any("Unmapped columns" in log for log in extraction_log):
            recommendations.append("Some columns could not be mapped. Review column naming conventions.")
        
        if report.get("total_heat_duty_kw", 0) == 0:
            recommendations.append("No heat duty data extracted. Check duty column identification and units.")
        
        if report.get("total_heat_area_m2", 0) == 0:
            recommendations.append("No heat transfer area data extracted. Check area column identification.")
        
        report["recommendations"] = recommendations
        
        return report
    
    def print_hex_extraction_report(self):
        """Print detailed heat exchanger extraction report"""
        report = self.get_hex_extraction_report()
        
        if "error" in report:
            print(f"❌ {report['error']}")
            return
        
        print("\n" + "="*80)
        print("🔥 HEAT EXCHANGER DATA EXTRACTION REPORT")
        print("="*80)
        
        print(f"📁 File: {report['file_path']}")
        print(f"📊 Worksheets Analyzed: {report['worksheets_analyzed']}")
        print(f"🎯 Total Heat Exchangers Extracted: {report['total_data_extracted']}")
        
        if report['total_data_extracted'] > 0:
            print(f"⚡ Total Heat Duty: {report.get('total_heat_duty_kw', 0):,.1f} kW")
            print(f"📐 Total Heat Area: {report.get('total_heat_area_m2', 0):,.1f} m²")
        
        # Data quality breakdown
        quality_breakdown = report.get('data_quality_breakdown', {})
        if quality_breakdown:
            print(f"\n📈 Data Quality Breakdown:")
            for quality, count in quality_breakdown.items():
                percentage = (count / report['total_data_extracted']) * 100
                print(f"   {quality}: {count} ({percentage:.1f}%)")
        
        # Extraction success by type
        success_by_type = report.get('extraction_success_by_type', {})
        if success_by_type:
            print(f"\n🎯 Extraction Success by Data Type:")
            for data_type, count in success_by_type.items():
                if report['total_data_extracted'] > 0:
                    percentage = (count / report['total_data_extracted']) * 100
                    print(f"   {data_type.replace('_', ' ').title()}: {count}/{report['total_data_extracted']} ({percentage:.1f}%)")
        
        # Recent extraction log entries
        extraction_log = report.get('extraction_log', [])
        if extraction_log:
            print(f"\n📝 Key Extraction Log Entries:")
            for log_entry in extraction_log[-10:]:  # Show last 10 entries
                print(f"   • {log_entry}")
        
        # Recommendations
        recommendations = report.get('recommendations', [])
        if recommendations:
            print(f"\n💡 Recommendations:")
            for i, rec in enumerate(recommendations, 1):
                print(f"   {i}. {rec}")
        
        print("\n" + "="*80)
        return {}
    
    def integrate_hex_with_aspen_data(self, process_data: AspenProcessData) -> AspenProcessData:
        """Integrate heat exchanger Excel data with Aspen process data"""
        if not self.hex_loader or not self.hex_loader.processed_data:
            logger.warning("No heat exchanger data loaded for integration")
            return process_data
        
        try:
            hex_data = self.hex_loader.get_heat_exchanger_data_for_tea()
            
            # Add heat exchanger data to global parameters
            if not hasattr(process_data, 'global_parameters'):
                process_data.global_parameters = {}
            
            process_data.global_parameters.update({
                'total_hex_count': hex_data.get('hex_count', 0),
                'total_heat_duty_kW': hex_data.get('total_heat_duty_kW', 0.0),
                'total_heat_area_m2': hex_data.get('total_heat_area_m2', 0.0),
                'average_hex_size_m2': hex_data.get('average_hex_size_m2', 0.0),
                'hex_data_source': 'excel_integration'
            })
            
            # Create virtual heat exchanger units if none exist in Aspen data
            hex_units_in_aspen = sum(1 for unit in process_data.units.values() 
                                   if unit.type == EquipmentType.HEAT_EXCHANGER)
            
            if hex_units_in_aspen == 0 and hex_data.get('hex_count', 0) > 0:
                logger.info(f"Adding {hex_data['hex_count']} virtual heat exchangers from Excel data")
                
                for i, hex_info in enumerate(hex_data.get('heat_exchangers', [])):
                    unit_name = hex_info.get('name', f'HEX-EXCEL-{i:03d}')
                    duty = hex_info.get('duty', 0.0) * 1000  # Convert kW to W for UnitOperationData
                    
                    virtual_unit = UnitOperationData(
                        name=unit_name,
                        type=EquipmentType.HEAT_EXCHANGER,
                        duty=duty,
                        pressure_drop=None
                    )
                    
                    process_data.units[unit_name] = virtual_unit
            
            logger.info(f"Successfully integrated heat exchanger data: {hex_data['hex_count']} units, {hex_data['total_heat_duty_kW']:.0f} kW total")
            
        except Exception as e:
            logger.error(f"Error integrating heat exchanger data: {str(e)}")
        
        return process_data
    
    def extract_and_store_all_data(self, aspen_file: str, hex_file: str = None) -> Dict[str, Any]:
        """
        完整的数据提取和数据库存储流程
        
        Args:
            aspen_file: Aspen Plus 文件路径
            hex_file: 换热器Excel文件路径 (可选)
            
        Returns:
            包含提取结果和统计信息的字典
        """
        if not self.database:
            raise ValueError("Database not initialized. Cannot store data.")
        
        logger.info("\n" + "="*80)
        logger.info("🚀 STARTING COMPLETE DATA EXTRACTION AND STORAGE")
        logger.info("="*80)
        
        extraction_results = {
            'success': False,
            'aspen_file': aspen_file,
            'hex_file': hex_file,
            'session_id': None,
            'extraction_time': datetime.now().isoformat(),
            'data_counts': {},
            'errors': []
        }
        
        try:
            # 1. 开始新的数据库会话
            session_id = self.database.start_new_session(aspen_file, hex_file)
            extraction_results['session_id'] = session_id
            logger.info(f"✅ Database session started: {session_id}")
            
            # 2. 加载换热器数据（如果提供）
            if hex_file and os.path.exists(hex_file):
                logger.info(f"📊 Loading heat exchanger data from {hex_file}")
                hex_success = self.load_hex_data(hex_file)
                if hex_success:
                    logger.info("✅ Heat exchanger data loaded successfully")
                    # 存储换热器数据到数据库
                    hex_data = self.get_hex_data_for_tea()
                    if hex_data:
                        self.database.store_hex_data(hex_data)
                        extraction_results['data_counts']['heat_exchangers'] = hex_data.get('hex_count', 0)
                else:
                    extraction_results['errors'].append("Failed to load heat exchanger data")
            else:
                logger.info("⚠️ No heat exchanger file provided or file not found")
                extraction_results['data_counts']['heat_exchangers'] = 0
            
            # 3. 连接到Aspen Plus
            logger.info(f"🔌 Connecting to Aspen Plus: {aspen_file}")
            
            # 确保使用正确的文件路径
            if not os.path.exists(aspen_file):
                # 尝试在aspen_files子目录中查找
                potential_path = os.path.join(os.path.dirname(__file__), "aspen_files", os.path.basename(aspen_file))
                if os.path.exists(potential_path):
                    aspen_file = potential_path
                    logger.info(f"✅ Found file in aspen_files directory: {aspen_file}")
                else:
                    raise FileNotFoundError(f"Aspen file not found: {aspen_file} or {potential_path}")
            
            success = self.com_interface.connect(aspen_file)
            
            if not success:
                # 提供详细的连接失败信息
                logger.error(f"❌ Could not connect to Aspen file: {aspen_file}")
                logger.error("Possible reasons:")
                logger.error("  1. Aspen Plus is not installed")
                logger.error("  2. COM objects are not registered")
                logger.error("  3. File path is incorrect")
                logger.error("  4. Insufficient permissions")
                
                # 尝试COM可用性测试
                com_test = self.com_interface.test_com_availability()
                logger.error(f"COM test results: {com_test}")
                
                raise Exception(f"Could not connect to Aspen file: {aspen_file}")
            
            logger.info("✅ Successfully connected to Aspen Plus")
            
            # 4. 提取流股数据
            logger.info("🌊 Extracting stream data...")
            streams = self.extract_all_streams()
            
            if streams:
                # 转换StreamData对象为字典
                streams_dict = {}
                for name, stream in streams.items():
                    if hasattr(stream, '__dict__'):
                        stream_dict = {
                            'temperature': stream.temperature,
                            'pressure': stream.pressure,
                            'mass_flow': stream.mass_flow,
                            'volume_flow': stream.volume_flow,
                            'molar_flow': stream.molar_flow,
                            'composition': stream.composition
                        }
                        
                        # Add classification and custom name data if available
                        if hasattr(stream, 'category'):
                            stream_dict['stream_category'] = getattr(stream, 'category')
                            stream_dict['stream_sub_category'] = getattr(stream, 'sub_category', '')
                            stream_dict['classification_confidence'] = getattr(stream, 'classification_confidence', 0.0)
                        
                        if hasattr(stream, 'custom_name'):
                            stream_dict['custom_name'] = getattr(stream, 'custom_name')
                        
                        streams_dict[name] = stream_dict
                    else:
                        streams_dict[name] = stream
                
                self.database.store_stream_data(streams_dict)
                extraction_results['data_counts']['streams'] = len(streams)
                logger.info(f"✅ Stored {len(streams)} stream records in database")
            else:
                extraction_results['errors'].append("No stream data extracted")
            
            # 5. 提取设备数据
            logger.info("⚙️ Extracting equipment data...")
            equipment = self.extract_all_equipment()
            
            if equipment:
                self.database.store_equipment_data(equipment)
                extraction_results['data_counts']['equipment'] = len(equipment)
                logger.info(f"✅ Stored {len(equipment)} equipment records in database")
            else:
                extraction_results['errors'].append("No equipment data extracted")
            
            # 6. 计算统计信息
            total_heat_duty = 0.0
            total_heat_area = 0.0
            
            if self.hex_loader and self.hex_loader.processed_data:
                total_heat_duty = self.hex_loader.processed_data.get('total_heat_duty', 0.0)
                total_heat_area = self.hex_loader.processed_data.get('total_heat_area', 0.0)
            
            # 7. 完成数据库会话
            summary_stats = {
                'stream_count': extraction_results['data_counts'].get('streams', 0),
                'equipment_count': extraction_results['data_counts'].get('equipment', 0),
                'hex_count': extraction_results['data_counts'].get('heat_exchangers', 0),
                'total_heat_duty_kw': total_heat_duty,
                'total_heat_area_m2': total_heat_area
            }
            
            self.database.finalize_session(summary_stats)
            extraction_results['summary_stats'] = summary_stats
            
            # 8. 断开Aspen连接
            self.com_interface.disconnect()
            logger.info("🔌 Disconnected from Aspen Plus")
            
            extraction_results['success'] = True
            
            logger.info("\n" + "="*80)
            logger.info("🎉 DATA EXTRACTION AND STORAGE COMPLETED SUCCESSFULLY")
            logger.info("="*80)
            logger.info(f"📊 Summary Statistics:")
            logger.info(f"   • Streams: {summary_stats['stream_count']}")
            logger.info(f"   • Equipment: {summary_stats['equipment_count']}")
            logger.info(f"   • Heat Exchangers: {summary_stats['hex_count']}")
            logger.info(f"   • Total Heat Duty: {summary_stats['total_heat_duty_kw']:.1f} kW")
            logger.info(f"   • Total Heat Area: {summary_stats['total_heat_area_m2']:.1f} m²")
            logger.info(f"   • Session ID: {session_id}")
            logger.info("="*80)
            
            return extraction_results
            
        except Exception as e:
            extraction_results['errors'].append(str(e))
            logger.error(f"❌ Data extraction failed: {str(e)}")
            
            # 确保断开连接
            try:
                self.com_interface.disconnect()
            except:
                pass
            
            return extraction_results
    
    def get_database_summary(self) -> Dict[str, Any]:
        """获取数据库摘要信息"""
        if not self.database:
            return {"error": "Database not initialized"}
        
        return self.database.get_database_summary()
    
    def export_database_to_json(self, output_file: str = None) -> bool:
        """导出数据库到JSON文件"""
        if not self.database:
            logger.error("Database not initialized")
            return False
        
        if not output_file:
            output_file = f"aspen_data_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        return self.database.export_to_json(output_file)
    
    def calculate_equipment_sizes(self, process_data: AspenProcessData) -> Dict[str, EquipmentSizeData]:
        """
        Calculate equipment sizes based on process data
        Note: This method is kept for compatibility but focus is on data extraction
        
        Args:
            process_data: Process data from Aspen simulation
            
        Returns:
            Dictionary of equipment sizing results
        """
        equipment_sizes = {}
        
        logger.info("Equipment sizing is handled by equipment sizing calculations")
        logger.info("Focus of this extractor is on accurate data extraction from Aspen and Excel")
        
        return equipment_sizes
    
    def extract_all_streams(self) -> Dict[str, StreamData]:
        """Extract all stream data from Aspen simulation using proven methods"""
        streams = {}
        
        if not self.com_interface.connected:
            logger.error("Not connected to Aspen Plus")
            return streams
            
        try:
            logger.info("Extracting all stream data from Aspen Plus...")
            
            # Get all stream names
            stream_names = self.com_interface.get_all_streams()
            
            for stream_name in stream_names:
                try:
                    stream_path = f"\\Data\\Streams\\{stream_name}"
                    
                    # Extract stream properties using proven paths
                    temp = self.com_interface.get_aspen_value(f"{stream_path}\\Output\\TEMP_OUT\\MIXED")
                    pres = self.com_interface.get_aspen_value(f"{stream_path}\\Output\\PRES_OUT\\MIXED")
                    mass_flow = self.com_interface.get_aspen_value(f"{stream_path}\\Output\\MASSFLMX\\MIXED")
                    vol_flow = self.com_interface.get_aspen_value(f"{stream_path}\\Output\\VOLFLMX\\MIXED")
                    molar_flow = self.com_interface.get_aspen_value(f"{stream_path}\\Output\\MOLEFLMX\\MIXED")
                    
                    # Get composition
                    composition = self.com_interface.get_stream_composition(stream_name)
                    
                    if temp is not None and pres is not None and mass_flow is not None:
                        # Create basic stream data dict for classification
                        stream_data_dict = {
                            'name': stream_name,
                            'temperature': float(temp),
                            'pressure': float(pres),
                            'mass_flow': float(mass_flow),
                            'volume_flow': float(vol_flow) if vol_flow is not None else 0.0,
                            'molar_flow': float(molar_flow) if molar_flow is not None else 0.0,
                            'composition': composition
                        }
                        
                        # Classify stream if classifier is available
                        stream_category = None
                        stream_sub_category = ""
                        classification_confidence = 0.0
                        
                        if self.stream_classifier:
                            try:
                                classification = self.stream_classifier.classify_stream(stream_data_dict)
                                stream_category = classification.category.value
                                stream_sub_category = classification.sub_category
                                classification_confidence = classification.confidence
                                
                                logger.debug(f"🏷️ {stream_name}: {stream_category} - {stream_sub_category} (置信度: {classification_confidence:.2f})")
                            except Exception as e:
                                logger.warning(f"Failed to classify stream {stream_name}: {str(e)}")
                        
                        # Create StreamData object
                        stream_data = StreamData(
                            name=stream_name,
                            temperature=float(temp),
                            pressure=float(pres),
                            mass_flow=float(mass_flow),
                            volume_flow=float(vol_flow) if vol_flow is not None else 0.0,
                            molar_flow=float(molar_flow) if molar_flow is not None else 0.0,
                            composition=composition
                        )
                        
                        # Add classification attributes if available
                        if stream_category:
                            setattr(stream_data, 'category', stream_category)
                            setattr(stream_data, 'sub_category', stream_sub_category)
                            setattr(stream_data, 'classification_confidence', classification_confidence)
                        
                        # Get user-defined display name from Aspen Plus
                        display_name = self.com_interface.get_stream_display_name(stream_name)
                        setattr(stream_data, 'custom_name', display_name)
                        
                        streams[stream_name] = stream_data
                        logger.info(f"Extracted stream: {stream_name} - T:{temp:.1f}°C, P:{pres:.1f}bar - {stream_category}")
                    
                    
                except Exception as e:
                    logger.warning(f"Could not extract stream {stream_name}: {str(e)}")
            
            self.stream_data = streams
            logger.info(f"✅ Successfully extracted {len(streams)} streams")
            
            # Print classification summary if classifier is available
            if self.stream_classifier and streams:
                self._print_stream_classification_summary(streams)
            
            return streams
            
        except Exception as e:
            logger.error(f"Failed to extract stream data: {str(e)}")
            return streams
    
    def _extract_stream_data(self) -> Dict[str, StreamData]:
        """Legacy method - now calls extract_all_streams"""
        return self.extract_all_streams()
    
    def _print_stream_classification_summary(self, streams: Dict[str, StreamData]):
        """Print stream classification summary"""
        if not streams:
            return
        
        logger.info("\n" + "="*60)
        logger.info("🏷️ STREAM CLASSIFICATION SUMMARY")
        logger.info("="*60)
        
        # Count by category
        category_counts = {}
        total_classified = 0
        
        for stream_name, stream_data in streams.items():
            if hasattr(stream_data, 'category'):
                category = getattr(stream_data, 'category')
                category_counts[category] = category_counts.get(category, 0) + 1
                total_classified += 1
        
        logger.info(f"Total streams: {len(streams)}")
        logger.info(f"Classified streams: {total_classified}")
        
        if category_counts:
            logger.info("\n按分类统计:")
            for category, count in sorted(category_counts.items()):
                percentage = (count / len(streams)) * 100
                logger.info(f"  {category}: {count} ({percentage:.1f}%)")
            
            logger.info("\n详细分类:")
            current_category = None
            for stream_name, stream_data in sorted(streams.items(), key=lambda x: getattr(x[1], 'category', '未分类')):
                if hasattr(stream_data, 'category'):
                    category = getattr(stream_data, 'category')
                    sub_category = getattr(stream_data, 'sub_category', '')
                    confidence = getattr(stream_data, 'classification_confidence', 0.0)
                    
                    if category != current_category:
                        current_category = category
                        logger.info(f"\n{category}:")
                    
                    sub_info = f" - {sub_category}" if sub_category else ""
                    confidence_info = f" (置信度: {confidence:.2f})"
                    logger.info(f"  • {stream_name}{sub_info}{confidence_info}")
        
        logger.info("="*60)
    
    def extract_all_equipment(self) -> Dict[str, Dict[str, Any]]:
        """Extract all equipment data from Aspen simulation using enhanced methods"""
        equipment = {}
        
        if not self.com_interface.connected:
            logger.error("Not connected to Aspen Plus")
            return equipment
            
        try:
            logger.info("Extracting all equipment data from Aspen Plus...")
            
            # Get all block names
            block_names = self.com_interface.get_all_blocks()
            logger.info(f"Found {len(block_names)} equipment blocks")
            
            for block_name in block_names:
                try:
                    # Get Aspen block type
                    block_type = self.com_interface.get_block_type(block_name)
                    
                    # Use strict Excel-based equipment matching if available
                    if self.equipment_matcher:
                        equipment_type = self.equipment_matcher.get_equipment_type(block_name)
                        eq_info = self.equipment_matcher.get_equipment_info(block_name)
                        equipment_function = eq_info.get('function', 'Unknown') if eq_info else 'Unknown'
                        
                        if equipment_type == "Unknown Equipment":
                            logger.info(f"⚠️ Equipment {block_name} not found in Excel specifications, skipping...")
                            continue
                    
                    # Use enhanced equipment detector if available (fallback)
                    elif self.equipment_detector:
                        equipment_info_obj = self.equipment_detector.detect_equipment_type(
                            equipment_name=block_name,
                            aspen_type=block_type
                        )
                        
                        equipment_type = equipment_info_obj.category
                        equipment_function = self.equipment_detector.get_equipment_function(equipment_info_obj)
                        
                        # Check if equipment should be included
                        if not self.equipment_detector.should_include_equipment(equipment_info_obj):
                            logger.info(f"Skipping equipment: {block_name} ({equipment_type})")
                            continue
                        
                        # Get comprehensive parameters using enhanced method
                        parameters = self._extract_comprehensive_parameters(
                            block_name, equipment_info_obj
                        )
                        
                        # Get additional metadata
                        importance = self.equipment_detector.get_equipment_importance(equipment_info_obj)
                        
                    else:
                        # Fallback to original method
                        equipment_type = self._map_aspen_block_type(block_type) if block_type else "Unknown"
                        equipment_function = "Unknown"
                        parameters = self.com_interface.get_equipment_parameters(block_name)
                        importance = "Unknown"
                    
                    # For Excel-matched equipment, get parameters using unified method
                    if self.equipment_matcher and equipment_type != "Unknown Equipment":
                        parameters = self._extract_equipment_parameters_unified(block_name, equipment_type)
                        importance = "High"  # Excel-specified equipment is high priority
                    
                    # Build comprehensive equipment info
                    equipment_info = {
                        "name": block_name,
                        "type": equipment_type,
                        "aspen_type": block_type or "Unknown",
                        "importance": importance if 'importance' in locals() else "Medium",
                        "function": equipment_function,
                        "parameters": parameters,
                        "parameter_count": len(parameters),
                        "excel_specified": self.equipment_matcher is not None and equipment_type != "Unknown Equipment"
                    }
                    
                    # Get user-defined display name from Aspen Plus
                    display_name = self.com_interface.get_equipment_display_name(block_name)
                    equipment_info["custom_name"] = display_name
                    
                    # Add stream connections for database storage
                    try:
                        inlet_streams, outlet_streams = self.com_interface.get_equipment_stream_connections_from_excel(block_name)
                        equipment_info["inlet_streams"] = inlet_streams
                        equipment_info["outlet_streams"] = outlet_streams
                    except Exception as e:
                        logger.debug(f"Could not get stream connections for {block_name}: {str(e)}")
                        equipment_info["inlet_streams"] = []
                        equipment_info["outlet_streams"] = []
                    
                    equipment[block_name] = equipment_info
                    logger.info(f"✅ Extracted {block_name}: {equipment_type} with {len(parameters)} parameters {'(Excel specified)' if equipment_info.get('excel_specified') else ''}")
                    
                except Exception as e:
                    logger.warning(f"Could not extract equipment {block_name}: {str(e)}")
            
            self.block_data = equipment
            logger.info(f"✅ Successfully extracted {len(equipment)} equipment items")
            
            # Collect stream connections separately (doesn't affect database)
            self._collect_equipment_connections(equipment)
            
            # Print equipment summary
            self._print_equipment_summary(equipment)
            
            return equipment
            
        except Exception as e:
            logger.error(f"Failed to extract equipment data: {str(e)}")
            return equipment
    
    def _collect_equipment_connections(self, equipment: Dict[str, Any]):
        """Collect stream connection information separately from main equipment data"""
        self.equipment_connections = {}
        
        for block_name in equipment.keys():
            try:
                inlet_streams, outlet_streams = self.com_interface.get_equipment_stream_connections_from_excel(block_name)
                self.equipment_connections[block_name] = {
                    'inlet_streams': inlet_streams,
                    'outlet_streams': outlet_streams
                }
            except Exception as e:
                logger.debug(f"Could not get stream connections for {block_name}: {str(e)}")
                self.equipment_connections[block_name] = {
                    'inlet_streams': [],
                    'outlet_streams': []
                }
        
        logger.info(f"✅ Collected stream connections for {len(self.equipment_connections)} equipment items")
    
    def get_equipment_stream_connections(self, equipment_name: str = None):
        """Get stream connection information for equipment
        
        Args:
            equipment_name (str, optional): Specific equipment name. If None, returns all connections.
            
        Returns:
            dict: Stream connection information
        """
        if equipment_name:
            return self.equipment_connections.get(equipment_name, {
                'inlet_streams': [],
                'outlet_streams': []
            })
        else:
            return self.equipment_connections
    
    def _map_aspen_block_type(self, aspen_type: str) -> str:
        """Map Aspen block type to equipment type"""
        if not aspen_type:
            return "Unknown"
        
        aspen_type_upper = aspen_type.upper()
        
        # Common Aspen Plus block types mapping
        type_mapping = {
            # Reactors
            "RSTOIC": "Reactor",
            "RPLUG": "Reactor", 
            "RCSTR": "Reactor",
            "RGIBB": "Reactor",
            "RYIELD": "Reactor",
            
            # Separators and Flash
            "FLASH2": "Separator",
            "FLASH3": "Separator", 
            "SEP": "Separator",
            "SEP2": "Separator",
            
            # Distillation
            "RADFRAC": "Distillation Column",
            "DSTWU": "Distillation Column",
            "SHORTCUT": "Distillation Column",
            
            # Heat Exchange
            "HEATX": "Heat Exchanger",
            "HEATER": "Heat Exchanger",
            "COOLER": "Heat Exchanger",
            "MHEATX": "Heat Exchanger",
            
            # Compression and Pumping
            "COMPR": "Compressor",
            "MCOMPR": "Compressor",
            "PUMP": "Pump",
            "ISENTROPIC": "Compressor",  # Added for MC1
            
            # Mixing and Splitting
            "MIXER": "Mixer",
            "FSPLIT": "Splitter",
            "SPLIT": "Splitter",
            
            # Controllers and Specs
            "T-SPEC": "Temperature Controller",  # Added for MEOH
            "P-SPEC": "Pressure Controller",
            "DSGN-SPEC": "Design Spec",
            
            # Others
            "VALVE": "Valve",
            "PIPE": "Pipe"
        }
        
        return type_mapping.get(aspen_type_upper, f"Unknown ({aspen_type})")
    
    # Equipment parameter mapping for unified extraction
    EQUIPMENT_PARAMETER_MAPS = {
        'reactor': {
            'volume_m3': ['\\Input\\VOLUME'],
            'temperature_C': ['\\Output\\B_TEMP', '\\Output\\T'],
            'pressure_bar': ['\\Output\\B_PRES', '\\Output\\P'],
            'duty_kW': ['\\Output\\B_DUTY', '\\Output\\DUTY']
        },
        'compressor': {
            'power_kW': ['\\Output\\WNET', '\\Output\\BRAKE_POWER'],
            'inlet_pressure_bar': ['\\Output\\PIN', '\\Output\\B_PRES'],
            'outlet_pressure_bar': ['\\Output\\POUT', '\\Output\\B_PRES2'],
            'inlet_temperature_C': ['\\Output\\TIN', '\\Output\\B_TEMP'],
            'outlet_temperature_C': ['\\Output\\TOUT', '\\Output\\B_TEMP2'],
            'efficiency': ['\\Input\\EFF']
        },
        'column': {
            'theoretical_stages': ['\\Input\\NSTAGE'],
            'feed_stage': ['\\Input\\FEED_STAGE\\1'],
            'top_pressure_bar': ['\\Output\\TOP_PRES'],
            'bottom_pressure_bar': ['\\Output\\BOT_PRES'],
            'top_temperature_C': ['\\Output\\TOP_TEMP'],
            'bottom_temperature_C': ['\\Output\\BOT_TEMP'],
            'reflux_ratio': ['\\Output\\MOLE_RR'],
            'reboiler_duty_kW': ['\\Output\\REB_DUTY'],
            'condenser_duty_kW': ['\\Output\\COND_DUTY']
        },
        'heat_exchanger': {
            'duty_kW': ['\\Output\\QCALC', '\\Output\\DUTY'],
            'outlet_temperature_C': ['\\Output\\T', '\\Output\\B_TEMP2'],
            'outlet_pressure_bar': ['\\Output\\P', '\\Output\\B_PRES2']
        },
        'separator': {
            'temperature_C': ['\\Output\\B_TEMP', '\\Output\\T'],
            'pressure_bar': ['\\Output\\B_PRES', '\\Output\\P'],
            'duty_kW': ['\\Output\\B_DUTY', '\\Output\\DUTY']
        },
        'general': {
            'temperature_C': ['\\Output\\T'],
            'pressure_bar': ['\\Output\\P'],
            'duty_kW': ['\\Output\\DUTY', '\\Output\\QCALC']
        }
    }
    
    def _extract_equipment_parameters_unified(self, block_name: str, equipment_type: str = 'general') -> Dict[str, Any]:
        """
        Unified method for extracting equipment parameters
        
        Args:
            block_name: Equipment block name
            equipment_type: Type of equipment ('reactor', 'compressor', 'column', etc.)
            
        Returns:
            Dictionary of extracted parameters
        """
        parameters = {}
        block_path = f"\\Data\\Blocks\\{block_name}"
        
        # Get parameter map for equipment type
        param_map = self.EQUIPMENT_PARAMETER_MAPS.get(equipment_type.lower(), 
                                                     self.EQUIPMENT_PARAMETER_MAPS['general'])
        
        try:
            for param_name, path_list in param_map.items():
                value = None
                
                # Try each path until we find a valid value
                for path_suffix in path_list:
                    full_path = block_path + path_suffix
                    try:
                        value = self.com_interface.get_aspen_value(full_path)
                        if self._is_valid_parameter_value(value):
                            break
                    except Exception:
                        continue
                
                # Process and store the value
                if self._is_valid_parameter_value(value):
                    # Convert power/duty values from W to kW if needed
                    if 'duty' in param_name.lower() and isinstance(value, (int, float)):
                        if abs(value) > 10000:  # Assume values > 10kW are in Watts
                            value = value / 1000
                        parameters[param_name] = abs(value)  # Always positive for duty
                    elif 'power' in param_name.lower() and isinstance(value, (int, float)):
                        parameters[param_name] = abs(value)
                    else:
                        parameters[param_name] = value
                    
                    logger.debug(f"Found {param_name}: {value}")
            
            # Calculate derived parameters
            if equipment_type.lower() == 'compressor':
                if 'inlet_pressure_bar' in parameters and 'outlet_pressure_bar' in parameters:
                    if parameters['inlet_pressure_bar'] > 0:
                        parameters['compression_ratio'] = parameters['outlet_pressure_bar'] / parameters['inlet_pressure_bar']
                        
        except Exception as e:
            logger.error(f"Error in unified parameter extraction for {block_name}: {str(e)}")
        
        return parameters

    def _extract_comprehensive_parameters(self, block_name: str, equipment_info_obj) -> Dict[str, Any]:
        """
        Enhanced parameter extraction using unified method
        
        Args:
            block_name: Equipment block name
            equipment_info_obj: Enhanced detector equipment info object
            
        Returns:
            Equipment parameters dictionary
        """
        try:
            if self.equipment_detector:
                # Determine equipment category for parameter mapping
                equipment_category = equipment_info_obj.category.lower()
                
                # Map categories to parameter extraction types
                category_map = {
                    'reactor': 'reactor',
                    'compressor': 'compressor',
                    'distillation_column': 'column',
                    'heat_exchanger': 'heat_exchanger',
                    'separator': 'separator'
                }
                
                equipment_type = category_map.get(equipment_category, 'general')
                parameters = self._extract_equipment_parameters_unified(block_name, equipment_type)
                
                # Add common parameters if not found
                self._add_common_parameters(block_name, parameters)
                
                logger.info(f"Extracted {len(parameters)} parameters for {block_name} ({equipment_type})")
                return parameters
            else:
                # Fallback to original method
                return self.com_interface.get_equipment_parameters(block_name)
                
        except Exception as e:
            logger.error(f"Error in comprehensive parameter extraction for {block_name}: {str(e)}")
            # Final fallback
            return self._extract_equipment_parameters_unified(block_name, 'general')
    
    def _is_valid_parameter_value(self, value) -> bool:
        """Check if parameter value is valid"""
        if value is None:
            return False
        
        # Handle different value types
        if isinstance(value, (int, float)):
            return value != 0 and not (isinstance(value, float) and (
                math.isnan(value) or math.isinf(value)
            ))
        elif isinstance(value, str):
            return len(value.strip()) > 0
        else:
            return True
    
    def _add_common_parameters(self, block_name: str, parameters: Dict[str, Any]):
        """Add common equipment parameters"""
        # Common parameters that most equipment should have
        common_params = [
            ('temperature', f"\\\\Data\\\\Blocks\\\\{block_name}\\\\Output\\\\T"),
            ('pressure', f"\\\\Data\\\\Blocks\\\\{block_name}\\\\Output\\\\P"),
            ('duty', f"\\\\Data\\\\Blocks\\\\{block_name}\\\\Output\\\\DUTY"),
            ('qcalc', f"\\\\Data\\\\Blocks\\\\{block_name}\\\\Output\\\\QCALC")
        ]
        
        for param_name, path in common_params:
            if f"output_{param_name}" not in parameters:
                try:
                    value = self.com_interface.get_aspen_value(path)
                    if self._is_valid_parameter_value(value):
                        parameters[f"common_{param_name}"] = value
                except Exception:
                    continue
    
    def _print_equipment_summary(self, equipment: Dict[str, Dict[str, Any]]):
        """Print equipment extraction summary"""
        if not equipment:
            return
        
        logger.info("\n" + "="*60)
        logger.info("EQUIPMENT EXTRACTION SUMMARY")
        logger.info("="*60)
        
        # Count by type
        type_counts = {}
        total_params = 0
        
        for eq_name, eq_data in equipment.items():
            eq_type = eq_data.get('type', 'Unknown')
            type_counts[eq_type] = type_counts.get(eq_type, 0) + 1
            total_params += eq_data.get('parameter_count', 0)
        
        logger.info(f"Total Equipment: {len(equipment)}")
        logger.info(f"Total Parameters: {total_params}")
        
        logger.info("\nEquipment by Type:")
        for eq_type, count in sorted(type_counts.items()):
            logger.info(f"  {eq_type}: {count}")
        
        logger.info("\nDetailed Equipment List:")
        for eq_name, eq_data in equipment.items():
            # 构建流股连接信息显示
            connections = self.equipment_connections.get(eq_name, {})
            inlet_streams = connections.get('inlet_streams', [])
            outlet_streams = connections.get('outlet_streams', [])
            stream_info = f"[{len(inlet_streams)}→{len(outlet_streams)}]"
            
            logger.info(f"  {eq_name}: {eq_data['type']} {stream_info} "
                       f"({eq_data.get('aspen_type', 'Unknown')}) "
                       f"- {eq_data.get('parameter_count', 0)} params")
        
        logger.info("="*60)
    
    
    def _extract_unit_operation_data(self) -> Dict[str, UnitOperationData]:
        """Legacy method - now calls extract_all_equipment and converts format"""
        units = {}
        equipment_data = self.extract_all_equipment()
        
        for block_name, eq_data in equipment_data.items():
            try:
                # Convert equipment data to UnitOperationData format
                detected_type = self._detect_equipment_type_from_name(block_name)
                
                # Extract duty from parameters if available
                params = eq_data.get("parameters", {})
                duty = params.get("duty_kW", None)
                duty = duty * 1000 if duty else None  # Convert back to watts for UnitOperationData
                
                unit_data = UnitOperationData(
                    name=block_name,
                    type=detected_type if detected_type else EquipmentType.OTHER,
                    duty=duty,
                    pressure_drop=None  # Not extracted in this version
                )
                
                units[block_name] = unit_data
                
            except Exception as e:
                logger.warning(f"Could not convert equipment data for {block_name}: {str(e)}")
        
        return units
    
    def _extract_utility_data(self) -> Dict[str, UtilityData]:
        """Extract utility data - placeholder for future implementation"""
        utilities = {}
        
        # This could be implemented in future to extract:
        # - Steam consumption
        # - Cooling water usage
        # - Electricity consumption
        # - Fuel gas usage
        
        logger.info("Utility data extraction not implemented in this version")
        return utilities
    
    def _extract_global_parameters(self) -> Dict[str, Any]:
        """Extract global simulation parameters"""
        global_params = {}
        
        try:
            # Get simulation title if available
            title_node = self.com_interface.simulation.FindNode("\\Title")
            if title_node and hasattr(title_node, 'Value'):
                global_params['simulation_title'] = str(title_node.Value)
            
            # Add extraction timestamp
            global_params['extraction_timestamp'] = datetime.now().isoformat()
            
            # Add data source information
            global_params['data_source'] = 'aspen_plus_com_interface'
            global_params['extractor_version'] = '2.0'
            
        except Exception as e:
            logger.warning(f"Could not extract global parameters: {str(e)}")
        
        return global_params
    
    def export_data(self, data: AspenProcessData, output_file: str):
        """Export extracted data to file"""
        try:
            # Convert data to dictionary for JSON export
            export_dict = {
                'simulation_name': data.simulation_name,
                'timestamp': data.timestamp.isoformat(),
                'streams': {name: {
                    'name': stream.name,
                    'temperature': stream.temperature,
                    'pressure': stream.pressure,
                    'mass_flow': stream.mass_flow,
                    'volume_flow': stream.volume_flow,
                    'molar_flow': stream.molar_flow,
                    'composition': stream.composition
                } for name, stream in data.streams.items()},
                'equipment': {name: {
                    'name': unit.name,
                    'type': unit.type.value if hasattr(unit.type, 'value') else str(unit.type),
                    'duty': unit.duty,
                    'pressure_drop': unit.pressure_drop
                } for name, unit in data.units.items()},
                'utilities': {
                    'electricity_kW': data.utilities.electricity,
                    'heating_steam_kg_hr': data.utilities.heating_steam,
                    'cooling_water_m3_hr': data.utilities.cooling_water,
                    'fuel_gas_GJ_hr': data.utilities.fuel_gas
                },
                'global_parameters': data.global_parameters
            }
            
            # Export to JSON
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(export_dict, f, indent=2, ensure_ascii=False)
            
            logger.info(f"✅ Data exported to {output_file}")
            
        except Exception as e:
            logger.error(f"Failed to export data: {str(e)}")


def run_extractor_tests(aspen_file: str = None, hex_file: str = None, verbose: bool = True) -> Dict[str, Any]:
    """
    Unified test function for Aspen data extractor
    
    Args:
        aspen_file: Path to Aspen simulation file (optional, will use default if None)
        hex_file: Path to heat exchanger Excel file (optional, will use default if None)  
        verbose: Whether to print detailed output
        
    Returns:
        Dictionary with test results
    """
    # Set default file paths if not provided
    if not aspen_file or not hex_file:
        current_dir = Path(__file__).parent
        aspen_file = aspen_file or str(current_dir / "aspen_files" / "BFG-CO2H-MEOH V2 (purge burning).apw")
        hex_file = hex_file or str(current_dir / "BFG-CO2H-HEX.xlsx")
    
    results = {
        'timestamp': datetime.now().isoformat(),
        'aspen_file': aspen_file,
        'hex_file': hex_file,
        'tests': {},
        'success_count': 0,
        'total_tests': 0
    }
    
    if verbose:
        print("Enhanced Aspen Data Extractor - Unified Test Suite")
        print("=" * 55)
        print(f"Platform: {sys.platform} | Python: {sys.version.split()[0]}")
        print(f"Aspen file: {'✅' if os.path.exists(aspen_file) else '❌'} {Path(aspen_file).name}")
        print(f"HEX file: {'✅' if os.path.exists(hex_file) else '❌'} {Path(hex_file).name}")
    
    # Initialize extractor
    extractor = AspenDataExtractor()
    
    # Test 1: COM diagnostics
    results['total_tests'] += 1
    if verbose:
        print(f"\n1. Windows COM diagnostics...")
    try:
        com_test = extractor.com_interface.test_com_availability()
        results['tests']['com_diagnostics'] = {
            'success': com_test['pywin32_available'] and len(com_test['com_objects_found']) > 0,
            'pywin32_available': com_test['pywin32_available'],
            'com_objects_found': com_test['com_objects_found'],
            'recommendations': com_test['recommendations']
        }
        if results['tests']['com_diagnostics']['success']:
            results['success_count'] += 1
            if verbose:
                print(f"   ✅ COM setup OK - {len(com_test['com_objects_found'])} objects found")
        else:
            if verbose:
                print(f"   ❌ COM setup issues - check recommendations")
    except Exception as e:
        results['tests']['com_diagnostics'] = {'success': False, 'error': str(e)}
        if verbose:
            print(f"   ❌ COM test failed: {str(e)}")
    
    # Test 2: Heat exchanger data loading
    if os.path.exists(hex_file):
        results['total_tests'] += 1
        if verbose:
            print(f"\n2. Heat exchanger data loading...")
        try:
            hex_success = extractor.load_hex_data(hex_file)
            summary = extractor.get_hex_summary() if hex_success else {}
            results['tests']['hex_loading'] = {
                'success': hex_success,
                'hex_count': summary.get('total_heat_exchangers', 0),
                'columns_count': len(summary.get('columns', []))
            }
            if hex_success:
                results['success_count'] += 1
                if verbose:
                    print(f"   ✅ Loaded {summary.get('total_heat_exchangers', 0)} heat exchangers")
            else:
                if verbose:
                    print(f"   ❌ Failed to load heat exchanger data")
        except Exception as e:
            results['tests']['hex_loading'] = {'success': False, 'error': str(e)}
            if verbose:
                print(f"   ❌ HEX loading error: {str(e)}")
    
    # Test 3: Aspen data extraction
    if os.path.exists(aspen_file) and results['tests'].get('com_diagnostics', {}).get('success', False):
        results['total_tests'] += 1
        if verbose:
            print(f"\n3. Aspen Plus data extraction...")
        try:
            if extractor.com_interface.connect(aspen_file, visible=False):
                streams = extractor.extract_all_streams()
                equipment = extractor.extract_all_equipment()
                extractor.com_interface.disconnect()
                
                results['tests']['aspen_extraction'] = {
                    'success': True,
                    'streams_count': len(streams),
                    'equipment_count': len(equipment)
                }
                results['success_count'] += 1
                if verbose:
                    print(f"   ✅ Extracted {len(streams)} streams, {len(equipment)} equipment")
            else:
                results['tests']['aspen_extraction'] = {'success': False, 'error': 'Connection failed'}
                if verbose:
                    print(f"   ❌ Failed to connect to Aspen Plus")
        except Exception as e:
            results['tests']['aspen_extraction'] = {'success': False, 'error': str(e)}
            if verbose:
                print(f"   ❌ Aspen extraction error: {str(e)}")
    
    # Test 4: Equipment sizing
    results['total_tests'] += 1
    if verbose:
        print(f"\n4. Equipment sizing calculations...")
    try:
        test_hex = extractor.equipment_sizer.size_heat_exchanger(
            duty=1000.0, delta_t_lm=25.0, pressure=30.0, temperature=200.0
        )
        results['tests']['equipment_sizing'] = {
            'success': True,
            'test_hex_area_m2': test_hex.area
        }
        results['success_count'] += 1
        if verbose:
            print(f"   ✅ Equipment sizing OK - Test HEX: {test_hex.area:.1f} m²")
    except Exception as e:
        results['tests']['equipment_sizing'] = {'success': False, 'error': str(e)}
        if verbose:
            print(f"   ❌ Equipment sizing error: {str(e)}")
    
    # Summary
    results['success_rate'] = results['success_count'] / results['total_tests'] if results['total_tests'] > 0 else 0
    
    if verbose:
        print(f"\n{'='*55}")
        print(f"Test Results: {results['success_count']}/{results['total_tests']} successful ({results['success_rate']:.1%})")
        
        if results['success_rate'] >= 0.75:
            print(f"🎉 Extractor is working well! Ready for TEA calculations!")
        elif results['success_rate'] >= 0.5:
            print(f"⚠️  Partial functionality - some issues need attention")
        else:
            print(f"❌ Setup required - check diagnostics above")
    
    return results


def main():
    """Main entry point - run tests with default settings"""
    return run_extractor_tests(verbose=True)


if __name__ == "__main__":
    main()