# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Aspen Plus data extraction and TEA (Techno-Economic Analysis) toolkit designed to extract process data from Aspen Plus simulations and Excel heat exchanger data, store it in SQLite databases, and perform economic analysis. The project is primarily written in Chinese with English code comments.

## Core Architecture

### Data Extraction Layer
- **`aspen_data_extractor.py`** - Main data extraction engine that connects to Aspen Plus via COM interface (Windows only) and processes Excel heat exchanger data
- **`aspen_data_database.py`** - Database manager for storing extracted Aspen Plus process data in SQLite format
- **`data_interfaces.py`** - Standardized data structures and enums for streams, equipment, and utilities

### Data Processing Components
- **`stream_classifier.py`** - Classifies process streams by type and function
- **`stream_mapping.py`** - Maps streams between different data sources
- **`equipment_model_matcher.py`** - Matches equipment from Aspen to cost models

### Database Structure
The main SQLite database (`aspen_data.db`) contains:
- `streams` - Process stream data (temperature, pressure, flow, composition)
- `equipment` - Equipment operational data and parameters  
- `heat_exchangers` - Heat exchanger specifications from Excel/Aspen
- `sessions` - Extraction session metadata and statistics

## Development Commands

### Environment Setup
```bash
# Install dependencies
pip install -r requirements.txt

# For Windows with Aspen Plus integration
pip install pywin32
```

### Core Operations
```bash
# Extract data from Aspen Plus simulation
python aspen_data_extractor.py

# View database contents
python view_aspen_database.py  # (if exists)

# Check database completeness
python check_database_completeness.py

# Generate status reports
python final_status_report.py
```

### Testing
```bash
# Run basic functionality tests
python simple_test.py  # (if exists)

# Run test suite
python -m pytest tests/  # (if test directory exists)
```

## Key Dependencies

- **pandas, numpy** - Data processing and analysis
- **openpyxl** - Excel file handling for heat exchanger data
- **pywin32** - Windows COM interface for Aspen Plus integration (Windows only)
- **sqlite3** - Database operations (built-in Python module)
- **pydantic** - Data validation

## Platform Requirements

- **Windows required** for Aspen Plus COM interface integration
- **Aspen Plus V11+** for simulation data extraction
- Python 3.7+ with standard scientific computing stack

## Data Flow

1. **Extraction**: `AspenDataExtractor` connects to Aspen Plus via COM interface
2. **Processing**: Stream and equipment data is classified and validated using `data_interfaces`
3. **Storage**: Data is stored in SQLite database via `AspenDataDatabase`
4. **Analysis**: Extracted data serves as input for TEA calculations

## Important Notes

- The codebase uses conditional imports for Windows-specific COM functionality
- Database operations use session-based tracking for data versioning
- Heat exchanger data can be sourced from both Aspen simulations and Excel files
- All monetary values and engineering units follow standard industrial conventions