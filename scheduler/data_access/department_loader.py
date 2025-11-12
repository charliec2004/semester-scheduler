"""CSV loading utilities for department requirements.

This module provides functions to read and parse department staffing requirements
from a CSV file. The requirements specify target and maximum hours for each
department (e.g., "Marketing should have 15-20 hours staffed per week").
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Tuple

import pandas as pd

from scheduler.domain.models import DepartmentRequirements
from scheduler.data_access.staff_loader import _coerce_numeric, _normalize_columns


def load_department_requirements(path: Path) -> DepartmentRequirements:
    """Load department staffing requirements from a CSV file.
    
    Expected CSV columns:
    - department: Name of the department (e.g., "marketing", "events")
    - target_hours: Desired weekly hours for this department
    - max_hours: Maximum allowed weekly hours for this department
    
    Args:
        path: Path to the CSV file containing department requirements.
    
    Returns:
        DepartmentRequirements dataclass with targets and max_hours dicts.
    
    Raises:
        FileNotFoundError: If the CSV file doesn't exist.
        ValueError: If required columns are missing, data is invalid, or
                   target_hours exceeds max_hours.
    
    Example:
        >>> reqs = load_department_requirements(Path("cpd-requirements.csv"))
        >>> print(reqs.targets["marketing"])  # 15.0
        >>> print(reqs.max_hours["marketing"])  # 20.0
    """
    # Verify the file exists before attempting to read it
    if not path.exists():
        raise FileNotFoundError(f"Department requirements CSV not found: {path}")

    # Read the CSV file into a pandas DataFrame
    df = pd.read_csv(path)
    
    # Strip whitespace from column names (handles "department ", " target_hours", etc.)
    df.columns = [col.strip() for col in df.columns]
    
    # Create a case-insensitive mapping of column names
    # e.g., {"department": "department", "target_hours": "target_hours"}
    column_map = _normalize_columns(df)

    # Helper function to require a column and return its actual name
    def require_column(name: str) -> str:
        if name not in column_map:
            raise ValueError(f"Required column '{name}' not found in {path}")
        return column_map[name]

    # Get the actual column names from the CSV
    dept_col = require_column("department")
    target_col = require_column("target_hours")
    max_col = require_column("max_hours")

    # Initialize dictionaries to store parsed data
    department_targets: Dict[str, float] = {}
    department_max_hours: Dict[str, float] = {}

    # Iterate through each row in the CSV
    for _, row in df.iterrows():
        # Extract and clean the department name
        department = str(row[dept_col]).strip()
        
        # Validate: department name cannot be empty
        if not department:
            raise ValueError("Department requirements CSV contains an empty department name.")
        
        # Validate: no duplicate departments allowed
        if department in department_targets:
            raise ValueError(f"Duplicate department entry detected: '{department}'")

        # Parse and validate numeric values (coerce to float)
        target_hours = _coerce_numeric(row[target_col], target_col, department)
        max_hours = _coerce_numeric(row[max_col], max_col, department)
        
        # Validate: target cannot exceed maximum
        if max_hours < target_hours:
            raise ValueError(
                f"Department '{department}' has target hours ({target_hours}) exceeding max hours ({max_hours})."
            )
        
        # Store the validated data
        department_targets[department] = target_hours
        department_max_hours[department] = max_hours

    # Return a frozen DepartmentRequirements dataclass
    # (immutable to prevent accidental modification)
    return DepartmentRequirements(targets=department_targets, max_hours=department_max_hours)
