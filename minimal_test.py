#!/usr/bin/env python3
"""Minimal test script"""

import sys
import os

print("Testing Python environment...")
print(f"Python version: {sys.version}")
print(f"Current directory: {os.getcwd()}")

try:
    import pandas as pd
    print("✅ pandas imported")
except Exception as e:
    print(f"❌ pandas error: {e}")

try:
    # Comment out the problematic imports temporarily
    print("Trying to import main modules...")
    exec("""
import json
import logging
import re
from typing import Dict, List, Optional, Any, Union, Tuple
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Standard library imports
import sqlite3
import os
import sys
from collections import defaultdict, Counter
from dataclasses import dataclass, field

# Third-party library imports
import pandas as pd
import numpy as np

print("✅ All basic imports successful")
""")
    
except Exception as e:
    print(f"❌ Import error: {e}")
    import traceback
    traceback.print_exc()
    
print("Basic test completed")
