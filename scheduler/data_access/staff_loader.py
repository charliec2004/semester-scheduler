"""CSV loading utilities for staff data."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Set

import pandas as pd

from scheduler.config import AVAILABILITY_COLUMNS, DAY_NAMES, FRONT_DESK_ROLE, TIME_SLOT_STARTS
from scheduler.domain.models import StaffData


def _normalize_columns(df: pd.DataFrame) -> Dict[str, str]:
    """Create mapping from lowercase column names to original names."""
    normalized: Dict[str, str] = {}
    for column in df.columns:
        key = column.strip().lower()
        if key in normalized:
            raise ValueError(f"Duplicate column detected when normalizing headers: '{column}'")
        normalized[key] = column.strip()
    return normalized


def _parse_roles(raw_roles: Optional[str]) -> List[str]:
    if pd.isna(raw_roles):
        return []
    return [role.strip() for role in re.split(r"[;,]", str(raw_roles)) if role.strip()]


def _coerce_numeric(value, column_name: str, record_name: str) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValueError(
            f"Invalid numeric value '{value}' for column '{column_name}' on record '{record_name}'"
        ) from None


def load_staff_data(path: Path) -> StaffData:
    if not path.exists():
        raise FileNotFoundError(f"Staff CSV not found: {path}")

    df = pd.read_csv(path)
    df.columns = [col.strip() for col in df.columns]
    column_map = _normalize_columns(df)

    def require_column(name: str) -> str:
        if name not in column_map:
            raise ValueError(f"Required column '{name}' not found in {path}")
        return column_map[name]

    name_col = require_column("name")
    roles_col = require_column("roles")
    target_col = require_column("target_hours")
    max_col = require_column("max_hours")
    year_col = require_column("year")

    missing_availability = [col for col in AVAILABILITY_COLUMNS if col not in df.columns]
    if missing_availability:
        preview = ", ".join(missing_availability[:5])
        suffix = "..." if len(missing_availability) > 5 else ""
        raise ValueError(f"Missing availability columns in {path}: {preview}{suffix}")

    employees: List[str] = []
    qual: Dict[str, Set[str]] = {}
    weekly_hour_limits: Dict[str, float] = {}
    target_weekly_hours: Dict[str, float] = {}
    employee_year: Dict[str, int] = {}
    unavailable: Dict[str, Dict[str, List[int]]] = {}
    all_roles: Set[str] = set()

    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        if not name:
            raise ValueError("Encountered employee row with empty name.")
        if name in qual:
            raise ValueError(f"Duplicate employee name detected: '{name}'")

        roles = _parse_roles(row[roles_col])
        if not roles:
            raise ValueError(f"Employee '{name}' must have at least one role defined.")
        role_set = set(roles)
        all_roles.update(role_set)
        qual[name] = role_set

        max_hours = _coerce_numeric(row[max_col], max_col, name)
        target_hours = min(_coerce_numeric(row[target_col], target_col, name), max_hours)
        weekly_hour_limits[name] = max_hours
        target_weekly_hours[name] = target_hours

        year_value = _coerce_numeric(row[year_col], year_col, name)
        employee_year[name] = int(year_value)

        availability: Dict[str, List[int]] = {}
        for day in DAY_NAMES:
            unavailable_slots: List[int] = []
            for slot_index, start_time in enumerate(TIME_SLOT_STARTS):
                column = f"{day}_{start_time}"
                value = row[column]
                try:
                    can_work = int(float(value)) == 1
                except (TypeError, ValueError):
                    can_work = False
                if not can_work:
                    unavailable_slots.append(slot_index)
            if unavailable_slots:
                availability[day] = unavailable_slots
        if availability:
            unavailable[name] = availability

        employees.append(name)

    if FRONT_DESK_ROLE not in all_roles:
        raise ValueError(f"No employees qualified for required role '{FRONT_DESK_ROLE}'.")

    return StaffData(
        employees=employees,
        qual=qual,
        weekly_hour_limits=weekly_hour_limits,
        target_weekly_hours=target_weekly_hours,
        employee_year=employee_year,
        unavailable=unavailable,
        roles=sorted(all_roles),
    )
