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
from typing import Dict, List, Optional, Any, Union
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


class CostCategory(Enum):
    """Cost category enumeration for economic analysis"""
    EQUIPMENT = "equipment"
    INSTALLATION = "installation"
    INSTRUMENTATION = "instrumentation"
    PIPING = "piping"
    ELECTRICAL = "electrical"
    BUILDINGS = "buildings"
    YARD_IMPROVEMENTS = "yard_improvements"
    SERVICE_FACILITIES = "service_facilities"
    ENGINEERING = "engineering"
    CONSTRUCTION = "construction"
    CONTRACTORS_FEE = "contractors_fee"
    CONTINGENCY = "contingency"
    RAW_MATERIALS = "raw_materials"
    UTILITIES = "utilities"
    LABOR = "labor"
    MAINTENANCE = "maintenance"
    INSURANCE = "insurance"
    DEPRECIATION = "depreciation"
    OTHER = "other"


class CurrencyType(Enum):
    """Currency type enumeration"""
    USD = "USD"
    EUR = "EUR"
    CNY = "CNY"
    JPY = "JPY"
    GBP = "GBP"


class CostBasis(Enum):
    """Cost basis enumeration for economic calculations"""
    INSTALLED = "installed"
    BARE_MODULE = "bare_module"
    GRASSROOTS = "grassroots"
    BATTERY_LIMITS = "battery_limits"


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


# Economic Data Structures

@dataclass
class CostItem:
    """
    Individual cost item for economic analysis
    
    Represents a single cost component with detailed breakdown
    and calculation parameters.
    """
    name: str
    category: CostCategory
    base_cost: float                    # Base cost value
    currency: CurrencyType = CurrencyType.USD
    basis_year: int = 2024             # Cost basis year
    quantity: float = 1.0              # Equipment quantity
    unit: str = "each"                 # Cost unit (each, kg, m3, etc.)
    
    # Cost factors and multipliers
    installation_factor: float = 1.0   # Installation cost multiplier
    material_factor: float = 1.0       # Material cost factor
    location_factor: float = 1.0       # Location/region factor
    escalation_factor: float = 1.0     # Time escalation factor
    
    # Calculated costs
    installed_cost: Optional[float] = None      # Final installed cost
    total_cost: Optional[float] = None          # Total project cost
    
    # Supporting data
    vendor_quote: Optional[str] = None          # Vendor information
    estimation_method: Optional[str] = None     # Cost estimation method used
    cost_basis: Optional[CostBasis] = None      # Cost basis type
    notes: List[str] = field(default_factory=list)
    
    def calculate_installed_cost(self) -> float:
        """Calculate total installed cost including all factors"""
        self.installed_cost = (self.base_cost * self.quantity * 
                              self.installation_factor * self.material_factor * 
                              self.location_factor * self.escalation_factor)
        return self.installed_cost


@dataclass
class CapexData:
    """
    Capital expenditure (CAPEX) data structure
    
    Contains all capital investment costs broken down by category
    and equipment type for comprehensive economic analysis.
    """
    project_name: str
    total_capex: float = 0.0
    currency: CurrencyType = CurrencyType.USD
    basis_year: int = 2024
    
    # Major cost categories
    equipment_costs: Dict[str, CostItem] = field(default_factory=dict)
    installation_costs: Dict[str, CostItem] = field(default_factory=dict)
    indirect_costs: Dict[str, CostItem] = field(default_factory=dict)
    
    # Cost breakdown by percentage
    equipment_percentage: float = 0.0
    installation_percentage: float = 0.0
    engineering_percentage: float = 0.0
    contingency_percentage: float = 0.0
    
    # Location and timing factors
    location_factor: float = 1.0
    escalation_factor: float = 1.0
    
    # Contingency and indirect costs
    contingency_rate: float = 0.15        # 15% default contingency
    engineering_rate: float = 0.12        # 12% engineering costs
    construction_rate: float = 0.08       # 8% construction management
    
    def add_cost_item(self, cost_item: CostItem):
        """Add a cost item to appropriate category"""
        if cost_item.category == CostCategory.EQUIPMENT:
            self.equipment_costs[cost_item.name] = cost_item
        elif cost_item.category in [CostCategory.INSTALLATION, CostCategory.PIPING, 
                                   CostCategory.INSTRUMENTATION, CostCategory.ELECTRICAL]:
            self.installation_costs[cost_item.name] = cost_item
        else:
            self.indirect_costs[cost_item.name] = cost_item
    
    def calculate_total_capex(self) -> float:
        """Calculate total CAPEX from all cost items"""
        equipment_total = sum(item.calculate_installed_cost() 
                            for item in self.equipment_costs.values())
        installation_total = sum(item.calculate_installed_cost() 
                               for item in self.installation_costs.values())
        indirect_total = sum(item.calculate_installed_cost() 
                           for item in self.indirect_costs.values())
        
        subtotal = equipment_total + installation_total + indirect_total
        contingency = subtotal * self.contingency_rate
        
        self.total_capex = subtotal + contingency
        return self.total_capex


@dataclass
class OpexData:
    """
    Operating expenditure (OPEX) data structure
    
    Contains all annual operating costs including raw materials,
    utilities, labor, and maintenance for economic analysis.
    """
    project_name: str
    annual_opex: float = 0.0
    currency: CurrencyType = CurrencyType.USD
    operating_hours: float = 8760.0      # Annual operating hours
    
    # Raw material costs
    raw_material_costs: Dict[str, CostItem] = field(default_factory=dict)
    
    # Utility costs
    utility_costs: Dict[str, CostItem] = field(default_factory=dict)
    
    # Labor and overhead
    labor_costs: Dict[str, CostItem] = field(default_factory=dict)
    
    # Maintenance and other fixed costs
    maintenance_costs: Dict[str, CostItem] = field(default_factory=dict)
    
    # Operating cost factors (as percentage of CAPEX)
    maintenance_rate: float = 0.03        # 3% of CAPEX annually
    insurance_rate: float = 0.005         # 0.5% of CAPEX annually
    property_tax_rate: float = 0.02       # 2% of CAPEX annually
    
    def add_opex_item(self, cost_item: CostItem):
        """Add an operating cost item to appropriate category"""
        if cost_item.category == CostCategory.RAW_MATERIALS:
            self.raw_material_costs[cost_item.name] = cost_item
        elif cost_item.category == CostCategory.UTILITIES:
            self.utility_costs[cost_item.name] = cost_item
        elif cost_item.category == CostCategory.LABOR:
            self.labor_costs[cost_item.name] = cost_item
        elif cost_item.category == CostCategory.MAINTENANCE:
            self.maintenance_costs[cost_item.name] = cost_item
    
    def calculate_annual_opex(self, capex_total: float = 0.0) -> float:
        """Calculate total annual OPEX"""
        raw_materials_total = sum(item.calculate_installed_cost() 
                                for item in self.raw_material_costs.values())
        utilities_total = sum(item.calculate_installed_cost() 
                            for item in self.utility_costs.values())
        labor_total = sum(item.calculate_installed_cost() 
                        for item in self.labor_costs.values())
        maintenance_total = sum(item.calculate_installed_cost() 
                              for item in self.maintenance_costs.values())
        
        # Add fixed costs based on CAPEX
        maintenance_fixed = capex_total * self.maintenance_rate
        insurance_fixed = capex_total * self.insurance_rate
        property_tax_fixed = capex_total * self.property_tax_rate
        
        self.annual_opex = (raw_materials_total + utilities_total + labor_total + 
                           maintenance_total + maintenance_fixed + 
                           insurance_fixed + property_tax_fixed)
        return self.annual_opex


@dataclass
class FinancialParameters:
    """
    Financial analysis parameters for economic evaluation
    
    Contains all financial assumptions and calculated metrics
    for project profitability analysis.
    """
    project_name: str
    
    # Time parameters
    project_life: int = 20               # Project life in years
    construction_period: int = 2         # Construction period in years
    startup_period: int = 1              # Startup period in years
    
    # Financial assumptions
    discount_rate: float = 0.10          # Discount rate (10%)
    tax_rate: float = 0.25               # Corporate tax rate (25%)
    depreciation_method: str = "straight_line"
    depreciation_life: int = 10          # Depreciation life in years
    
    # Working capital
    working_capital_rate: float = 0.05   # As percentage of annual sales
    
    # Revenue parameters
    annual_revenue: float = 0.0          # Annual revenue
    product_price: float = 0.0           # Product selling price
    annual_production: float = 0.0       # Annual production capacity
    
    # Calculated financial metrics
    npv: Optional[float] = None          # Net Present Value
    irr: Optional[float] = None          # Internal Rate of Return
    payback_period: Optional[float] = None  # Simple payback period
    discounted_payback: Optional[float] = None  # Discounted payback period
    roi: Optional[float] = None          # Return on Investment
    
    # Cash flow components
    annual_cash_flows: List[float] = field(default_factory=list)
    cumulative_cash_flows: List[float] = field(default_factory=list)
    
    def calculate_npv(self, capex: float, annual_opex: float) -> float:
        """Calculate Net Present Value"""
        cash_flows = []
        
        # Initial investment (negative cash flow)
        cash_flows.append(-capex)
        
        # Annual operating cash flows
        annual_net_cash_flow = self.annual_revenue - annual_opex
        annual_after_tax = annual_net_cash_flow * (1 - self.tax_rate)
        
        for year in range(1, self.project_life + 1):
            discounted_flow = annual_after_tax / ((1 + self.discount_rate) ** year)
            cash_flows.append(discounted_flow)
        
        self.npv = sum(cash_flows)
        self.annual_cash_flows = cash_flows
        return self.npv


@dataclass
class EconomicAnalysisResults:
    """
    Complete economic analysis results container
    
    Aggregates all economic data and analysis results for
    comprehensive TEA (Techno-Economic Analysis) reporting.
    """
    project_name: str
    timestamp: datetime
    analysis_version: str = "1.0"
    
    # Cost data
    capex_data: CapexData = field(default_factory=lambda: CapexData(""))
    opex_data: OpexData = field(default_factory=lambda: OpexData(""))
    financial_params: FinancialParameters = field(default_factory=lambda: FinancialParameters(""))
    
    # Equipment sizing and costing
    equipment_list: Dict[str, EquipmentSizeData] = field(default_factory=dict)
    
    # Summary metrics
    total_capex: float = 0.0
    annual_opex: float = 0.0
    production_cost: float = 0.0         # Cost per unit of product
    break_even_price: float = 0.0        # Break-even selling price
    
    # Economic indicators
    npv: float = 0.0
    irr: float = 0.0
    payback_period: float = 0.0
    
    # Sensitivity analysis results
    sensitivity_parameters: Dict[str, Any] = field(default_factory=dict)
    sensitivity_results: Dict[str, Dict[str, float]] = field(default_factory=dict)
    
    # Data sources and methodology
    data_sources: List[str] = field(default_factory=list)
    estimation_methods: List[str] = field(default_factory=list)
    assumptions: List[str] = field(default_factory=list)
    
    # Quality metrics
    confidence_level: Optional[str] = None  # High, Medium, Low
    accuracy_range: Optional[str] = None    # ±10%, ±25%, ±50%
    
    def calculate_production_cost(self) -> float:
        """Calculate production cost per unit"""
        if self.financial_params.annual_production > 0:
            annual_total_cost = (self.annual_opex + 
                               self.total_capex / self.financial_params.project_life)
            self.production_cost = annual_total_cost / self.financial_params.annual_production
        return self.production_cost
    
    def get_economic_summary(self) -> Dict[str, Any]:
        """Get summary of economic analysis results"""
        return {
            'project_name': self.project_name,
            'analysis_date': self.timestamp.isoformat(),
            'total_capex': self.total_capex,
            'annual_opex': self.annual_opex,
            'production_cost': self.production_cost,
            'npv': self.npv,
            'irr': self.irr,
            'payback_period': self.payback_period,
            'confidence_level': self.confidence_level,
            'accuracy_range': self.accuracy_range,
            'equipment_count': len(self.equipment_list),
            'data_sources_count': len(self.data_sources)
        }