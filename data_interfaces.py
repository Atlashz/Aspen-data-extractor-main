#!/usr/bin/env python3
"""
Data Interfaces for Aspen Plus Data Extraction

Defines data structures and interfaces for process data extracted from
Aspen Plus simulations. These classes provide standardized data formats
for streams, equipment, and utilities.

Author: TEA Analysis Framework
Date: 2025-07-25
Version: 1.0
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any
from datetime import datetime
from enum import Enum


class EquipmentType(Enum):
    """Equipment type enumeration"""
    REACTOR = "reactor"
    COMPRESSOR = "compressor"
    PUMP = "pump"
    HEAT_EXCHANGER = "heat_exchanger"
    DISTILLATION_COLUMN = "distillation_column"
    SEPARATOR = "separator"
    TANK = "tank"
    VALVE = "valve"
    MIXER = "mixer"
    SPLITTER = "splitter"
    UNKNOWN = "unknown"


class MaterialType(Enum):
    """Material type enumeration for equipment construction"""
    CARBON_STEEL = "carbon_steel"
    SS304 = "ss304"
    SS316 = "ss316"
    HASTELLOY_C = "hastelloy_c"
    INCONEL = "inconel"
    TITANIUM = "titanium"


class PressureLevel(Enum):
    """Pressure level classification"""
    LOW = "low"          # < 10 bar
    MEDIUM = "medium"    # 10-50 bar
    HIGH = "high"        # > 50 bar


@dataclass
class StreamData:
    """
    Data structure for process stream information
    
    Contains all relevant thermodynamic and composition data
    for a process stream extracted from Aspen Plus.
    """
    name: str
    temperature: float  # °C
    pressure: float     # bar
    mass_flow: float    # kg/hr
    volume_flow: float = 0.0  # m3/hr
    molar_flow: float = 0.0   # kmol/hr
    composition: Dict[str, float] = field(default_factory=dict)  # Mole fractions
    enthalpy: Optional[float] = None  # kJ/hr
    entropy: Optional[float] = None   # kJ/hr-K
    density: Optional[float] = None   # kg/m3
    phase: Optional[str] = None       # Vapor, Liquid, Mixed
    
    def __post_init__(self):
        """Validate stream data after initialization"""
        if self.temperature < -273.15:
            raise ValueError(f"Invalid temperature: {self.temperature}°C")
        if self.pressure <= 0:
            raise ValueError(f"Invalid pressure: {self.pressure} bar")
        if self.mass_flow < 0:
            raise ValueError(f"Invalid mass flow: {self.mass_flow} kg/hr")


@dataclass
class UnitOperationData:
    """
    Data structure for unit operation/equipment information
    
    Contains operational parameters and design specifications
    for process equipment extracted from Aspen Plus.
    """
    name: str
    type: EquipmentType
    duty: Optional[float] = None        # kW (positive for heating, negative for cooling)
    pressure_drop: Optional[float] = None  # bar
    temperature: Optional[float] = None    # °C
    pressure: Optional[float] = None       # bar
    efficiency: Optional[float] = None     # Fraction (0-1)
    power_consumption: Optional[float] = None  # kW
    
    # Additional parameters stored as dictionary
    parameters: Dict[str, Any] = field(default_factory=dict)
    
    # Metadata
    aspen_block_type: Optional[str] = None
    notes: List[str] = field(default_factory=list)
    
    def add_parameter(self, name: str, value: Any, unit: str = None):
        """Add a parameter with optional unit information"""
        self.parameters[name] = {
            'value': value,
            'unit': unit
        }
    
    def get_parameter(self, name: str, default=None):
        """Get parameter value"""
        param = self.parameters.get(name, {})
        return param.get('value', default)


@dataclass
class UtilityData:
    """
    Data structure for utility consumption information
    
    Tracks utility requirements such as steam, cooling water,
    electricity, and fuel gas consumption.
    """
    equipment_name: str
    utility_type: str      # Steam, Cooling Water, Electricity, Fuel Gas
    consumption: float     # Amount consumed
    unit: str             # kg/hr, kW, m3/hr, etc.
    cost_factor: Optional[float] = None  # $/unit
    
    # Steam-specific parameters
    steam_pressure: Optional[float] = None  # bar
    steam_temperature: Optional[float] = None  # °C
    
    # Cooling water parameters
    inlet_temperature: Optional[float] = None  # °C
    outlet_temperature: Optional[float] = None  # °C
    flow_rate: Optional[float] = None  # m3/hr


@dataclass
class EquipmentSizeData:
    """
    Data structure for equipment sizing results
    
    Contains calculated dimensions, materials, and design parameters
    for process equipment based on sizing correlations.
    """
    equipment_type: EquipmentType
    name: str
    
    # Dimensions
    diameter: Optional[float] = None    # m
    length: Optional[float] = None      # m
    height: Optional[float] = None      # m
    volume: Optional[float] = None      # m3
    area: Optional[float] = None        # m2
    
    # Design conditions
    design_pressure: Optional[float] = None    # bar
    design_temperature: Optional[float] = None # °C
    material: Optional[MaterialType] = None
    pressure_level: Optional[PressureLevel] = None
    
    # Construction details
    wall_thickness: Optional[float] = None     # mm
    tube_count: Optional[int] = None          # For heat exchangers
    stages: Optional[int] = None              # For compressors/columns
    power_rating: Optional[float] = None      # kW
    
    # Sizing basis and assumptions
    sizing_basis: Dict[str, Any] = field(default_factory=dict)
    assumptions: List[str] = field(default_factory=list)
    
    # Cost estimation placeholder
    estimated_cost: Optional[float] = None    # $
    cost_basis: Optional[str] = None          # Basis year, location, etc.


@dataclass
class AspenProcessData:
    """
    Complete process data container
    
    Aggregates all extracted data from an Aspen Plus simulation
    including streams, units, utilities, and global parameters.
    """
    simulation_name: str
    timestamp: datetime
    
    # Process data
    streams: Dict[str, StreamData] = field(default_factory=dict)
    units: Dict[str, UnitOperationData] = field(default_factory=dict)
    utilities: Dict[str, UtilityData] = field(default_factory=dict)
    
    # Global simulation parameters
    global_parameters: Dict[str, Any] = field(default_factory=dict)
    
    # Metadata
    aspen_file_path: Optional[str] = None
    extraction_method: Optional[str] = None
    extraction_duration: Optional[float] = None  # seconds
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    
    def add_stream(self, stream: StreamData):
        """Add a stream to the process data"""
        self.streams[stream.name] = stream
    
    def add_unit(self, unit: UnitOperationData):
        """Add a unit operation to the process data"""
        self.units[unit.name] = unit
    
    def add_utility(self, utility: UtilityData):
        """Add utility data"""
        self.utilities[f"{utility.equipment_name}_{utility.utility_type}"] = utility
    
    def get_stream_by_name(self, name: str) -> Optional[StreamData]:
        """Get stream data by name"""
        return self.streams.get(name)
    
    def get_unit_by_name(self, name: str) -> Optional[UnitOperationData]:
        """Get unit operation data by name"""
        return self.units.get(name)
    
    def get_units_by_type(self, equipment_type: EquipmentType) -> List[UnitOperationData]:
        """Get all units of a specific type"""
        return [unit for unit in self.units.values() if unit.type == equipment_type]
    
    def get_summary(self) -> Dict[str, Any]:
        """Get a summary of the process data"""
        return {
            'simulation_name': self.simulation_name,
            'timestamp': self.timestamp.isoformat(),
            'stream_count': len(self.streams),
            'unit_count': len(self.units),
            'utility_count': len(self.utilities),
            'total_mass_flow': sum(stream.mass_flow for stream in self.streams.values()),
            'equipment_types': list(set(unit.type for unit in self.units.values())),
            'warnings_count': len(self.warnings),
            'errors_count': len(self.errors)
        }


# Utility functions for data validation and processing

def validate_stream_data(stream: StreamData) -> List[str]:
    """
    Validate stream data and return list of warnings
    
    Args:
        stream: StreamData object to validate
        
    Returns:
        List of validation warnings
    """
    warnings = []
    
    # Temperature checks
    if stream.temperature < -50:
        warnings.append(f"Very low temperature: {stream.temperature}°C")
    elif stream.temperature > 1000:
        warnings.append(f"Very high temperature: {stream.temperature}°C")
    
    # Pressure checks
    if stream.pressure > 200:
        warnings.append(f"Very high pressure: {stream.pressure} bar")
    
    # Flow consistency checks
    if stream.mass_flow > 0 and stream.volume_flow <= 0:
        warnings.append("Mass flow exists but volume flow is zero")
    
    # Composition checks
    if stream.composition:
        total_mole_fraction = sum(stream.composition.values())
        if abs(total_mole_fraction - 1.0) > 0.01:
            warnings.append(f"Composition doesn't sum to 1.0: {total_mole_fraction:.3f}")
    
    return warnings


def validate_unit_data(unit: UnitOperationData) -> List[str]:
    """
    Validate unit operation data and return list of warnings
    
    Args:
        unit: UnitOperationData object to validate
        
    Returns:
        List of validation warnings
    """
    warnings = []
    
    # Power consumption checks
    if unit.power_consumption and unit.power_consumption < 0:
        warnings.append(f"Negative power consumption: {unit.power_consumption} kW")
    
    # Efficiency checks
    if unit.efficiency:
        if unit.efficiency > 1.0:
            warnings.append(f"Efficiency > 100%: {unit.efficiency}")
        elif unit.efficiency < 0.1:
            warnings.append(f"Very low efficiency: {unit.efficiency}")
    
    # Pressure drop checks
    if unit.pressure_drop and unit.pressure_drop < 0:
        warnings.append(f"Negative pressure drop: {unit.pressure_drop} bar")
    
    return warnings


# Constants and conversion factors
CONVERSION_FACTORS = {
    'temperature': {
        'C_to_K': 273.15,
        'F_to_C': lambda f: (f - 32) * 5/9,
        'C_to_F': lambda c: c * 9/5 + 32
    },
    'pressure': {
        'bar_to_Pa': 100000,
        'psi_to_bar': 0.0689476,
        'bar_to_psi': 14.5038
    },
    'flow': {
        'kg_hr_to_kg_s': 1/3600,
        'lb_hr_to_kg_hr': 0.453592,
        'm3_hr_to_m3_s': 1/3600
    },
    'energy': {
        'kJ_hr_to_kW': 1/3600,
        'BTU_hr_to_kW': 0.000293071,
        'kcal_hr_to_kW': 0.00116222
    }
}


def convert_units(value: float, from_unit: str, to_unit: str, 
                 conversion_type: str) -> float:
    """
    Convert between different units
    
    Args:
        value: Value to convert
        from_unit: Source unit
        to_unit: Target unit
        conversion_type: Type of conversion (temperature, pressure, etc.)
        
    Returns:
        Converted value
    """
    factors = CONVERSION_FACTORS.get(conversion_type, {})
    conversion_key = f"{from_unit}_to_{to_unit}"
    
    if conversion_key in factors:
        factor = factors[conversion_key]
        if callable(factor):
            return factor(value)
        else:
            return value * factor
    else:
        raise ValueError(f"Unknown conversion: {from_unit} to {to_unit} for {conversion_type}")