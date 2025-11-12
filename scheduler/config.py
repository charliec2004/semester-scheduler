"""Centralized knobs for the scheduler. Tweak values here instead of touching the solver."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

# ---------------------------------------------------------------------------
# Calendar + availability grid configuration
# ---------------------------------------------------------------------------
DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri"]

TIME_SLOT_STARTS = [
    "08:00",
    "08:30",
    "09:00",
    "09:30",
    "10:00",
    "10:30",
    "11:00",
    "11:30",
    "12:00",
    "12:30",
    "13:00",
    "13:30",
    "14:00",
    "14:30",
    "15:00",
    "15:30",
    "16:00",
    "16:30",
]

SLOT_NAMES = [
    "8:00-8:30",
    "8:30-9:00",
    "9:00-9:30",
    "9:30-10:00",
    "10:00-10:30",
    "10:30-11:00",
    "11:00-11:30",
    "11:30-12:00",
    "12:00-12:30",
    "12:30-1:00",
    "1:00-1:30",
    "1:30-2:00",
    "2:00-2:30",
    "2:30-3:00",
    "3:00-3:30",
    "3:30-4:00",
    "4:00-4:30",
    "4:30-5:00",
]

AVAILABILITY_COLUMNS = [f"{day}_{time}" for day in DAY_NAMES for time in TIME_SLOT_STARTS]
T_SLOTS = list(range(len(SLOT_NAMES)))  # 30-minute slot indices

# ---------------------------------------------------------------------------
# Role + shift defaults
# ---------------------------------------------------------------------------
FRONT_DESK_ROLE = "front_desk"
DEPARTMENT_HOUR_THRESHOLD = 4  # Allowable +/- hour wiggle room for departments

MIN_SLOTS = 4  # 2 hours minimum shift (4 × 30-min slots)
MAX_SLOTS = 8  # 4 hours maximum shift
MIN_FRONT_DESK_SLOTS = MIN_SLOTS  # Front desk shifts must meet the same minimum length

# ---------------------------------------------------------------------------
# Solver + objective tuning knobs
# ---------------------------------------------------------------------------
DEFAULT_SOLVER_MAX_TIME = 90  # Seconds

FRONT_DESK_COVERAGE_WEIGHT = 10_000  # Weight applied to every covered slot
SHIFT_LENGTH_DAILY_COST = 6  # Slots subtracted per worked day (encourages longer blocks)
DEPARTMENT_SCARCITY_BASE_WEIGHT = 10.0  # Higher values penalize pulling scarce dept staff to front desk

YEAR_TARGET_MULTIPLIERS = {  # Target-hour adherence weight by academic year
    1: 1.0,
    2: 1.2,
    3: 1.5,
    4: 2.0,
}

LARGE_DEVIATION_SLOT_THRESHOLD = 4  # 4 slots = 2 hours from target
EMPLOYEE_LARGE_DEVIATION_PENALTY = 5000  # Per-employee massive penalty
DEPARTMENT_LARGE_DEVIATION_PENALTY = 4000  # Per-department penalty when missing ±threshold

COLLABORATION_MINIMUM_HOURS: Dict[str, int] = {
    # Expected collaborative hours (2+ people in the same department simultaneously)
    "career_education": 1,
    "marketing": 1,
    "employer_engagement": 2,
    "events": 4,
    "data_systems": 0,  # Single-person team; no collaboration requirement
}


@dataclass(frozen=True)
class ObjectiveWeights:
    """Scalar weights applied to each score/penalty component in the objective."""

    department_target: float = 1000.0
    collaborative_hours: float = 200.0
    office_coverage: float = 150.0
    single_coverage: float = 500.0
    target_adherence: float = 100.0
    department_spread: float = 60.0
    department_day_coverage: float = 30.0
    shift_length: float = 20.0
    department_scarcity: float = 8.0
    underclassmen_front_desk: float = 3.0
    morning_preference: float = 0.5
    department_total: float = 1.0


OBJECTIVE_WEIGHTS = ObjectiveWeights()
