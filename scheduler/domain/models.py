"""Dataclasses and type definitions shared across the scheduler modules."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set


@dataclass(frozen=True)
class StaffData:
    employees: List[str]
    qual: Dict[str, Set[str]]
    weekly_hour_limits: Dict[str, float]
    target_weekly_hours: Dict[str, float]
    employee_year: Dict[str, int]
    unavailable: Dict[str, Dict[str, List[int]]]
    roles: List[str]


@dataclass(frozen=True)
class DepartmentRequirements:
    targets: Dict[str, float]
    max_hours: Dict[str, float]


@dataclass(frozen=True)
class ScheduleRequest:
    staff_csv: Path
    requirements_csv: Path
    output_path: Path
