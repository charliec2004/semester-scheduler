"""
Employee Scheduling Optimizer using Google OR-Tools CP-SAT Solver

This program creates an optimal weekly employee schedule that satisfies:
- Continuous shift blocks (no split shifts)
- Employee availability constraints
- Role qualifications
- Coverage requirements with front_desk priority
- Work hour limits per shift

Key Features:
- front_desks MUST be present at all times (hard constraint)
- Department work (career education, marketing, internships, employer engagement, events, data & systems) can only occur when a front_desk is staffed
- Multiple employees can work the same role simultaneously
- Each employee works one continuous block per day (3-6 hours)
"""

import argparse
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple, TypedDict

import importlib.util
import pandas as pd
from ortools.sat.python import cp_model


DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri"]
TIME_SLOT_STARTS = [
    "08:00", "08:30", "09:00", "09:30",
    "10:00", "10:30", "11:00", "11:30",
    "12:00", "12:30", "13:00", "13:30",
    "14:00", "14:30", "15:00", "15:30",
    "16:00", "16:30",
]
SLOT_NAMES = [
    "8:00-8:30", "8:30-9:00", "9:00-9:30", "9:30-10:00",
    "10:00-10:30", "10:30-11:00", "11:00-11:30", "11:30-12:00",
    "12:00-12:30", "12:30-1:00", "1:00-1:30", "1:30-2:00",
    "2:00-2:30", "2:30-3:00", "3:00-3:30", "3:30-4:00",
    "4:00-4:30", "4:30-5:00",
]
AVAILABILITY_COLUMNS = [f"{day}_{time}" for day in DAY_NAMES for time in TIME_SLOT_STARTS]
FRONT_DESK_ROLE = "front_desk"
DEPARTMENT_HOUR_THRESHOLD = 4  # +/- hours acceptable window


class StaffData(TypedDict):
    """Type definition for staff data returned by load_staff_data()"""
    employees: List[str]
    qual: Dict[str, Set[str]]
    weekly_hour_limits: Dict[str, float]
    target_weekly_hours: Dict[str, float]
    employee_year: Dict[str, int]
    unavailable: Dict[str, Dict[str, List[int]]]
    roles: List[str]


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

    return {
        "employees": employees,
        "qual": qual,
        "weekly_hour_limits": weekly_hour_limits,
        "target_weekly_hours": target_weekly_hours,
        "employee_year": employee_year,
        "unavailable": unavailable,
        "roles": sorted(all_roles),
    }


def load_department_requirements(path: Path) -> Tuple[Dict[str, float], Dict[str, float]]:
    if not path.exists():
        raise FileNotFoundError(f"Department requirements CSV not found: {path}")

    df = pd.read_csv(path)
    df.columns = [col.strip() for col in df.columns]
    column_map = _normalize_columns(df)

    def require_column(name: str) -> str:
        if name not in column_map:
            raise ValueError(f"Required column '{name}' not found in {path}")
        return column_map[name]

    dept_col = require_column("department")
    target_col = require_column("target_hours")
    max_col = require_column("max_hours")

    department_targets: Dict[str, float] = {}
    department_max_hours: Dict[str, float] = {}

    for _, row in df.iterrows():
        department = str(row[dept_col]).strip()
        if not department:
            raise ValueError("Department requirements CSV contains an empty department name.")
        if department in department_targets:
            raise ValueError(f"Duplicate department entry detected: '{department}'")

        target_hours = _coerce_numeric(row[target_col], target_col, department)
        max_hours = _coerce_numeric(row[max_col], max_col, department)
        if max_hours < target_hours:
            raise ValueError(
                f"Department '{department}' has target hours ({target_hours}) exceeding max hours ({max_hours})."
            )
        department_targets[department] = target_hours
        department_max_hours[department] = max_hours

    return department_targets, department_max_hours


def main():
    """Main function to build and solve the scheduling model"""
    
    parser = argparse.ArgumentParser(
        description="Generate an optimized weekly schedule for CPD student employees."
    )
    parser.add_argument(
        "staff_csv",
        type=Path,
        help="CSV file containing employee information, roles, hours, and availability.",
    )
    parser.add_argument(
        "requirements_csv",
        type=Path,
        help="CSV file specifying department hour targets and maximums.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("schedule.xlsx"),
        help="Destination path for the exported Excel schedule (default: schedule.xlsx).",
    )
    args = parser.parse_args()

    try:
        staff_data = load_staff_data(args.staff_csv)
        department_hour_targets_raw, department_max_hours_raw = load_department_requirements(
            args.requirements_csv
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
    
    # ============================================================================
    # STEP 1: INITIALIZE THE CONSTRAINT PROGRAMMING MODEL
    # ============================================================================
    
    model = cp_model.CpModel()
    
    
    # ============================================================================
    # STEP 2: DEFINE THE PROBLEM DOMAIN
    # ============================================================================
    
    employees: List[str] = staff_data["employees"]
    qual: Dict[str, Set[str]] = staff_data["qual"]
    weekly_hour_limits = {emp: float(hours) for emp, hours in staff_data["weekly_hour_limits"].items()}
    target_weekly_hours = {emp: float(hours) for emp, hours in staff_data["target_weekly_hours"].items()}
    employee_year = {emp: int(year) for emp, year in staff_data["employee_year"].items()}
    unavailable: Dict[str, Dict[str, List[int]]] = staff_data["unavailable"]
    
    days = DAY_NAMES[:]
    roles = list(staff_data["roles"])
    if FRONT_DESK_ROLE not in roles:
        print(f"ERROR: Role '{FRONT_DESK_ROLE}' is required but missing from staff data.", file=sys.stderr)
        sys.exit(1)
    roles = [FRONT_DESK_ROLE] + [role for role in roles if role != FRONT_DESK_ROLE]
    department_roles = [role for role in roles if role != FRONT_DESK_ROLE]
    
    missing_targets = [role for role in department_roles if role not in department_hour_targets_raw]
    missing_max = [role for role in department_roles if role not in department_max_hours_raw]
    if missing_targets:
        print(f"ERROR: Department targets missing for: {', '.join(missing_targets)}", file=sys.stderr)
        sys.exit(1)
    if missing_max:
        print(f"ERROR: Department max hours missing for: {', '.join(missing_max)}", file=sys.stderr)
        sys.exit(1)
    extra_departments = [dept for dept in department_hour_targets_raw if dept not in roles]
    if extra_departments:
        print(f"WARNING: Ignoring department requirements with no matching role: {', '.join(extra_departments)}", file=sys.stderr)

    department_hour_targets = {
        role: float(department_hour_targets_raw[role])
        for role in department_roles
    }
    department_max_hours = {
        role: float(department_max_hours_raw[role])
        for role in department_roles
    }
    department_hour_threshold = DEPARTMENT_HOUR_THRESHOLD
    
    department_sizes = {
        role: sum(1 for employee in employees if role in qual[employee])
        for role in department_roles
    }
    zero_capacity_departments = [role for role, size in department_sizes.items() if size == 0]
    if zero_capacity_departments:
        print(
            "ERROR: No qualified employees found for departments: "
            + ", ".join(zero_capacity_departments),
            file=sys.stderr,
        )
        sys.exit(1)
    
    ROLE_DISPLAY_NAMES = {
        role: " ".join(word.capitalize() for word in role.split("_"))
        for role in roles
    }
    ROLE_DISPLAY_NAMES[FRONT_DESK_ROLE] = "Front Desk"
    
    # Time slot configuration - 30 MINUTE INCREMENTS
    # T represents 18 half-hour time slots from 8am to 5pm
    # Index 0 = 8:00-8:30, Index 1 = 8:30-9:00, Index 2 = 9:00-9:30, ..., Index 17 = 4:30-5:00
    T = list(range(len(SLOT_NAMES)))  # 0-17 for 18 thirty-minute slots
    
    # Shift length constraints (in 30-minute increments)
    # Each slot = 0.5 hours, so multiply by 2 to get slot counts
    MIN_SLOTS = 4   # Minimum 2 hours = 4 thirty-minute slots
    MAX_SLOTS = 8   # Maximum 4 hours = 8 thirty-minute slots (changed from 12)
    MIN_FRONT_DESK_SLOTS = MIN_SLOTS  # Front desk duty must last at least full shift minimum
    
    
    # ============================================================================
    # STEP 3: DEFINE COVERAGE REQUIREMENTS
    # ============================================================================
    
    # Initialize demand dictionary: demand[role][day][time_slot]
    # Value of 1 means "we need 1 person in this role at this time"
    # Value of 0 means "no requirement for this role at this time"
    demand = {
        role: {                    # Dictionary (outer level)
            day: [0] * len(T)      # Dictionary (middle level) -> List (inner level)
            for day in days
        } for role in roles
    }
    
    # front_desk coverage is CRITICAL - must be present at all times
    for day in days:
        for time_slot in T:
            demand["front_desk"][day][time_slot] = 1
    
    # Note: Department roles have no fixed demand - they're assigned flexibly
    # based on availability and the objective function
    
    
    # ============================================================================
    # STEP 4 & STEP 5: QUALIFICATIONS AND AVAILABILITY
    # ============================================================================
    # Qualifications and availability are loaded from the staff CSV input.

    # ============================================================================
    # STEP 6: CREATE DECISION VARIABLES
    # ============================================================================
    
    # Boolean variable: Is employee 'e' working on day 'd' during time slot 't'?
    # work[e,d,t] = 1 means "yes", 0 means "no"
    work = {
        (e, d, t): model.new_bool_var(f"work[{e},{d},{t}]") 
        for e in employees 
        for d in days 
        for t in T
    }
    
    # Boolean variable: Does employee 'e' START their shift at time slot 't' on day 'd'?
    # Used to enforce continuous shift blocks
    start = {
        (e, d, t): model.new_bool_var(f"start[{e},{d},{t}]") 
        for e in employees 
        for d in days 
        for t in T
    }
    
    # Boolean variable: Does employee 'e' END their shift at time slot 't' on day 'd'?
    # Used to enforce continuous shift blocks
    end = {
        (e, d, t): model.new_bool_var(f"end[{e},{d},{t}]") 
        for e in employees 
        for d in days 
        for t in T
    }
    
    # Boolean variables to track front desk assignment transitions (ensures contiguous front desk duty)
    frontdesk_start = {
        (e, d, t): model.new_bool_var(f"frontdesk_start[{e},{d},{t}]")
        for e in employees
        if "front_desk" in qual[e]
        for d in days
        for t in T
    }
    frontdesk_end = {
        (e, d, t): model.new_bool_var(f"frontdesk_end[{e},{d},{t}]")
        for e in employees
        if "front_desk" in qual[e]
        for d in days
        for t in T
    }
    
    # Boolean variable: Is employee 'e' assigned to role 'r' on day 'd' at time 't'?
    # Only create this variable if the employee is qualified for the role
    assign = {
        (e, d, t, r): model.new_bool_var(f"assign[{e},{d},{t},{r}]")
        for e in employees 
        for d in days 
        for t in T 
        for r in roles 
        if r in qual[e]  # Only if employee is qualified
    }
    
    
    # ============================================================================
    # STEP 7: ADD SHIFT CONTIGUITY CONSTRAINTS
    # ============================================================================
    # These constraints ensure each employee works ONE CONTINUOUS BLOCK per day
    # (no split shifts like 8-10am and 2-4pm)
    # HARD CONSTRAINT: Once someone stops working, they cannot start again that day
    
    for e in employees:
        for d in days:
            
            # Constraint 7.1: EXACTLY one shift start per day (if working at all)
            # An employee can't start working multiple times in one day
            # This is a HARD constraint - no split shifts allowed
            model.add(sum(start[e, d, t] for t in T) <= 1)
            
            # Constraint 7.2: EXACTLY one shift end per day (if working at all)
            # An employee can't end their shift multiple times in one day
            # This is a HARD constraint - no split shifts allowed
            model.add(sum(end[e, d, t] for t in T) <= 1)
            
            # Constraint 7.2b: Number of starts must equal number of ends
            # This ensures if someone starts, they must end (and vice versa)
            model.add(sum(start[e, d, t] for t in T) == sum(end[e, d, t] for t in T))
            
            # Constraint 7.3: First time slot boundary
            # If working at time 0, that must be the start (no previous slot exists)
            model.add(work[e, d, 0] == start[e, d, 0])
            
            # Constraint 7.4: Internal time slot transitions
            # This is the KEY constraint for continuous blocks
            # For each time slot after the first:
            #   - If work changes from 0→1, we started (start=1, end(prev)=0)
            #   - If work changes from 1→0, we ended (start=0, end(prev)=1)
            #   - If work stays same, no transition (start=0, end(prev)=0)
            for t in T[1:]:
                model.add(
                    work[e, d, t] - work[e, d, t-1] == start[e, d, t] - end[e, d, t-1]
                )
            
            # Constraint 7.5: Last time slot boundary
            # If working in the last slot, that must be the end (no next slot exists)
            model.add(end[e, d, T[-1]] == work[e, d, T[-1]])
            
            # Calculate total slots worked this day (in 30-minute increments)
            total_slots_today = sum(work[e, d, t] for t in T)
            
            # Constraint 7.6 & 7.7: HARD minimum shift length constraint
            # CRITICAL: If someone works AT ALL, they MUST work at least MIN_SLOTS (2 hours = 4 slots)
            # This prevents tiny shifts like 30 minutes or 1 hour
            # No matter the role, if you work, it's minimum 2 hours in one continuous block
            
            # CRITICAL HARD CONSTRAINT: Enforce minimum shift length
            # We cannot use a simple indicator variable because it creates weak implications
            # Instead, we use a direct constraint: total_slots must be EITHER 0 OR >= MIN_SLOTS
            # 
            # Create boolean: is employee working today?
            works_today = model.new_bool_var(f"works_today[{e},{d}]")
            
            # Force works_today to be 1 if and only if total_slots_today > 0
            # This creates a tight bidirectional link
            model.add(total_slots_today >= 1).only_enforce_if(works_today)
            model.add(total_slots_today == 0).only_enforce_if(works_today.Not())
            
            # HARD CONSTRAINT: If working (works_today=1), MUST work at least MIN_SLOTS
            # This prevents 1, 2, or 3 slot shifts (0.5, 1.0, or 1.5 hours)
            model.add(total_slots_today >= MIN_SLOTS).only_enforce_if(works_today)
            
            # HARD CONSTRAINT: If not working (works_today=0), total must be exactly 0
            model.add(total_slots_today == 0).only_enforce_if(works_today.Not())
            
            # Maximum shift length (always enforced)
            model.add(total_slots_today <= MAX_SLOTS)
            
            # ADDITIONAL NUCLEAR OPTION: Add explicit constraint that blocks 1, 2, 3 slot shifts
            # This is a redundant safeguard to absolutely prevent short shifts
            # For each forbidden value, add a constraint that total_slots != that value
            model.add(total_slots_today != 1)  # Not 30 minutes
            model.add(total_slots_today != 2)  # Not 1 hour
            model.add(total_slots_today != 3)  # Not 1.5 hours
    
    
    # ============================================================================
    # STEP 7B: ADD WEEKLY HOUR LIMIT CONSTRAINTS
    # ============================================================================
    # Limit total hours per employee per week (prevents overwork)
    # Two levels: individual personal maximum preferences AND universal 19-hour limit
    
    UNIVERSAL_MAXIMUM_HOURS = 19  # Universal limit - no one can exceed this regardless of personal preference
    
    for e in employees:
        # Sum up all SLOTS worked across the entire week
        total_weekly_slots = sum(
            work[e, d, t] 
            for d in days 
            for t in T
        )
        
        # Individual personal preference limit (customized per employee)
        max_weekly_hours = weekly_hour_limits.get(e, 40)  # Default to 40 if not specified
        max_weekly_slots = int(round(max_weekly_hours * 2))  # Convert hours to 30-minute slots
        model.add(total_weekly_slots <= max_weekly_slots)
        
        # Universal maximum (applies to everyone)
        universal_max_slots = UNIVERSAL_MAXIMUM_HOURS * 2
        model.add(total_weekly_slots <= universal_max_slots)
        
        print(f"   └─ {e}: max {max_weekly_hours} hours/week (universal limit: {UNIVERSAL_MAXIMUM_HOURS}h)")
    
    
    # ============================================================================
    # STEP 8: ADD AVAILABILITY CONSTRAINTS
    # ============================================================================
    # Employees cannot work during times they've marked as unavailable
    
    for e in employees:
        for d in days:
            # Check if this employee has any unavailability
            if e in unavailable and d in unavailable[e]:
                # Force work variable to 0 for each unavailable time slot
                for t in unavailable[e][d]:
                    if t in T:  # Validate it's a valid time slot
                        model.add(work[e, d, t] == 0)
    
    
    # ============================================================================
    # STEP 9: ADD ROLE ASSIGNMENT CONSTRAINTS
    # ============================================================================
    
    for e in employees:
        for d in days:
            for t in T:
                
                # Constraint 9.1: Can't do two roles simultaneously
                # An employee can be assigned to at most one role at a time
                model.add(
                    sum(assign.get((e, d, t, r), 0) for r in roles) <= 1
                )
                
                # Constraint 9.2: Must be working to be assigned a role
                # If assigned to a role, the employee must be working that slot
                for r in roles:
                    if (e, d, t, r) in assign:
                        model.add(assign[(e, d, t, r)] <= work[e, d, t])
                
                # Constraint 9.2b: CRITICAL REVERSE CONSTRAINT
                # If working, MUST be assigned to exactly one role
                # This ensures work[e,d,t]=1 implies they have a role assignment
                model.add(
                    sum(assign.get((e, d, t, r), 0) for r in qual[e]) == work[e, d, t]
                )
                
                # Constraint 9.3: CRITICAL - Non-front desk roles need front desk supervision
                # Any departmental assignment can ONLY happen when at least one front_desk is present
                # This prevents scenarios where only departmental work is happening unsupervised
                for r in department_roles:
                    if (e, d, t, r) in assign:
                        model.add(
                            sum(assign.get((emp, d, t, "front_desk"), 0) for emp in employees) >= 1
                        ).only_enforce_if(assign[(e, d, t, r)])

    # ============================================================================
    # STEP 9B: FRONT DESK ASSIGNMENT CONTIGUITY
    # ============================================================================
    # Prevent employees from toggling in and out of front desk duty within the same shift
    
    for e in employees:
        if "front_desk" not in qual[e]:
            continue
        for d in days:
            fd_starts = [frontdesk_start[e, d, t] for t in T]
            fd_ends = [frontdesk_end[e, d, t] for t in T]
            model.add(sum(fd_starts) <= 1)
            model.add(sum(fd_ends) <= 1)
            model.add(sum(fd_starts) == sum(fd_ends))
            
            assign_fd_0 = assign[(e, d, 0, "front_desk")]
            model.add(assign_fd_0 == frontdesk_start[e, d, 0])
            
            for t in T[1:]:
                assign_curr = assign[(e, d, t, "front_desk")]
                assign_prev = assign[(e, d, t-1, "front_desk")]
                model.add(
                    assign_curr - assign_prev == frontdesk_start[e, d, t] - frontdesk_end[e, d, t-1]
                )
            
            model.add(frontdesk_end[e, d, T[-1]] == assign[(e, d, T[-1], "front_desk")])
            
            # HARD CONSTRAINT: Front desk minimum 2 hours (4 slots) if working it at all
            # This prevents short front desk stints like 30min or 1 hour
            total_front_desk_slots = sum(assign[(e, d, t, "front_desk")] for t in T)
            
            # NUCLEAR OPTION: Explicitly forbid 1, 2, or 3 slot front desk shifts
            # Total front desk slots must be EITHER 0 (not working front desk) OR >= 4 (minimum 2 hours)
            model.add(total_front_desk_slots != 1)  # Not 30 minutes
            model.add(total_front_desk_slots != 2)  # Not 1 hour
            model.add(total_front_desk_slots != 3)  # Not 1.5 hours
    
    
    # ============================================================================
    # STEP 9C: MINIMUM ROLE DURATION CONSTRAINT (ALL ROLES)
    # ============================================================================
    # Prevent employees from doing any role for less than 1 hour (2 slots)
    # Example: Can't do front_desk 8am-10am, then marketing 10am-10:30am
    # If you switch to a role, you must do it for at least 1 hour continuously
    
    # Create role start/end tracking variables for ALL roles
    role_start = {
        (e, d, t, r): model.new_bool_var(f"role_start[{e},{d},{t},{r}]")
        for e in employees
        for d in days
        for t in T
        for r in roles
        if r in qual[e]
    }
    role_end = {
        (e, d, t, r): model.new_bool_var(f"role_end[{e},{d},{t},{r}]")
        for e in employees
        for d in days
        for t in T
        for r in roles
        if r in qual[e]
    }
    
    for e in employees:
        for d in days:
            for r in roles:
                if r not in qual[e]:
                    continue
                
                # Enforce contiguous role assignment (can't toggle in and out of a role)
                # At most one start and one end per role per day
                model.add(sum(role_start.get((e, d, t, r), 0) for t in T) <= 1)
                model.add(sum(role_end.get((e, d, t, r), 0) for t in T) <= 1)
                model.add(
                    sum(role_start.get((e, d, t, r), 0) for t in T) == 
                    sum(role_end.get((e, d, t, r), 0) for t in T)
                )
                
                # First slot boundary
                if (e, d, 0, r) in assign:
                    model.add(assign[(e, d, 0, r)] == role_start[(e, d, 0, r)])
                
                # Internal transitions
                for t in T[1:]:
                    if (e, d, t, r) in assign and (e, d, t-1, r) in assign:
                        model.add(
                            assign[(e, d, t, r)] - assign[(e, d, t-1, r)] == 
                            role_start[(e, d, t, r)] - role_end[(e, d, t-1, r)]
                        )
                
                # Last slot boundary
                if (e, d, T[-1], r) in assign:
                    model.add(role_end[(e, d, T[-1], r)] == assign[(e, d, T[-1], r)])
                
                # HARD CONSTRAINT: Minimum 1 hour (2 slots) per role assignment
                total_role_slots = sum(assign.get((e, d, t, r), 0) for t in T)
                
                # Forbid single 30-minute slot for any role
                model.add(total_role_slots != 1)
    
    
    # ============================================================================
    # STEP 10: ADD COVERAGE REQUIREMENTS
    # ============================================================================
    # Ensure minimum staffing levels are met for each role
    
    # Create coverage tracking variables (soft constraints via objective)
    front_desk_coverage_score = 0
    
    for d in days:
        for t in T:
            # Create indicator: is front desk covered at this time?
            has_front_desk = model.new_bool_var(f"has_front_desk[{d},{t}]")
            num_front_desk = sum(assign.get((e, d, t, "front_desk"), 0) for e in employees)
            
            # Link indicator to actual coverage (at least 1 front desk)
            model.add(num_front_desk >= 1).only_enforce_if(has_front_desk)
            model.add(num_front_desk == 0).only_enforce_if(has_front_desk.Not())
            
            # VERY STRONG SOFT CONSTRAINT: Front desk should be covered at all times
            # We use MASSIVE weight (10000) to make this extremely high priority
            # This is NOT a hard constraint - if truly impossible, solver can still find a solution
            # But practically, front desk will only be uncovered if NO front-desk-qualified 
            # employee is available at that time slot
            front_desk_coverage_score += 10000 * has_front_desk
            
            # HARD CONSTRAINT: At most 1 front desk at a time (no overstaffing at front desk)
            model.add(num_front_desk <= 1)
            
            # NOTE: Department roles CAN have multiple people working at the same time
            # We removed the hard cap - instead we'll use soft constraints in the objective
            # to encourage spreading people out throughout the week
    
    
    # ============================================================================
    # STEP 11: DEFINE THE OBJECTIVE FUNCTION
    # ============================================================================
    # What we're trying to optimize (maximize in this case)
    
    # Count total departmental assignments across all employees, days, and times
    department_assignments = {
        role: sum(
            assign.get((e, d, t, role), 0)
            for e in employees
            for d in days
            for t in T
        )
        for role in department_roles
    }
    total_department_assignments = sum(department_assignments.values())
    department_max_slots = {
        role: int(round(department_max_hours[role] * 2))
        for role in department_roles
    }
    for role in department_roles:
        model.add(department_assignments[role] <= department_max_slots[role])
    
    # Calculate "spread" metric for each department: count how many time slots have at least 1 worker
    # This encourages distribution throughout the day rather than clustering
    department_spread_score = 0
    for role in department_roles:
        for d in days:
            for t in T:
                has_role = model.new_bool_var(f"has_{role}[{d},{t}]")
                num_role = sum(assign.get((e, d, t, role), 0) for e in employees)
                
                model.add(num_role >= 1).only_enforce_if(has_role)
                model.add(num_role == 0).only_enforce_if(has_role.Not())
                
                department_spread_score += has_role
    
    # Encourage each department to appear across multiple days
    department_day_coverage_score = 0
    for role in department_roles:
        for d in days:
            has_role_day = model.new_bool_var(f"has_{role}[{d}]")
            total_role_day = sum(assign.get((e, d, t, role), 0) for e in employees for t in T)
            model.add(total_role_day >= 1).only_enforce_if(has_role_day)
            model.add(total_role_day == 0).only_enforce_if(has_role_day.Not())
            department_day_coverage_score += has_role_day

    # Encourage departments to hit target weekly hours (soft constraint)
    department_target_score = 0
    department_large_deviation_penalty = 0
    threshold_slots = department_hour_threshold * 2

    for role in department_roles:
        target_hours = department_hour_targets.get(role)
        if target_hours is None:
            continue
        max_capacity_hours = sum(weekly_hour_limits.get(e, 0) for e in employees if role in qual[e])
        max_requirement_hours = department_max_hours.get(role, max_capacity_hours)
        adjusted_target_hours = min(target_hours, max_capacity_hours, max_requirement_hours)
        target_slots = int(adjusted_target_hours * 2)
        total_role_slots = department_assignments[role]

        over = model.new_int_var(0, 200, f"department_over[{role}]")
        under = model.new_int_var(0, 200, f"department_under[{role}]")
        model.add(total_role_slots == target_slots + over - under)

        department_target_score -= over + under

        if threshold_slots > 0:
            large_over = model.new_bool_var(f"department_large_over[{role}]")
            large_under = model.new_bool_var(f"department_large_under[{role}]")

            model.add(over >= threshold_slots).only_enforce_if(large_over)
            model.add(over < threshold_slots).only_enforce_if(large_over.Not())

            model.add(under >= threshold_slots).only_enforce_if(large_under)
            model.add(under < threshold_slots).only_enforce_if(large_under.Not())

            department_large_deviation_penalty -= 4000 * (large_over + large_under)
    
    # Calculate target hours encouragement (SOFT constraint via objective)
    # We want to encourage employees to work close to their target hours
    # This is NOT a hard constraint - solver will try to get close but won't fail if impossible
    # We'll create penalty variables for being over/under target
    # IMPORTANT: Apply graduated weights based on year - upperclassmen get stronger adherence
    # This counterbalances the front desk preference that favors underclassmen
    target_adherence_score = 0
    large_deviation_penalty = 0  # Steep penalty for being 2+ hours off target
    
    for e in employees:
        # Calculate total slots worked by this employee across the week
        total_slots = sum(work[e, d, t] for d in days for t in T)
        
        # Get this employee's target (in hours, convert to slots)
        target_hours = target_weekly_hours.get(e, 11)  # Default 11 hours
        target_slots = int(target_hours * 2)  # Convert to 30-min slots
        
        # Create variables to track deviation from target
        over_target = model.new_int_var(0, 100, f"over_target[{e}]")
        under_target = model.new_int_var(0, 100, f"under_target[{e}]")
        
        # Deviation equation: total_slots = target_slots + over_target - under_target
        model.add(total_slots == target_slots + over_target - under_target)
        
        # Graduated weighting based on year:
        # Seniors and juniors get higher weight to ensure they hit their hours
        # even if it means putting them at front desk (overriding the underclassmen preference)
        year = employee_year.get(e, 2)
        year_multiplier = {
            1: 1.0,  # Freshman - base weight
            2: 1.2,  # Sophomore - slightly higher
            3: 1.5,  # Junior - significantly higher (ensures they get hours)
            4: 2.0,  # Senior - highest priority (ensures they always get hours)
        }.get(year, 1.0)
        
        # Penalize deviation with graduated weight
        # Upperclassmen deviations are penalized more heavily
        target_adherence_score -= year_multiplier * (over_target + under_target)
        
        # STEEP PENALTY for large deviations (2+ hours = 4+ slots off target)
        # This applies to EVERYONE regardless of year
        # Create indicator variables for "large deviation"
        large_over = model.new_bool_var(f"large_over[{e}]")
        large_under = model.new_bool_var(f"large_under[{e}]")
        
        # Large over = more than 4 slots (2 hours) over target
        model.add(over_target >= 4).only_enforce_if(large_over)
        model.add(over_target < 4).only_enforce_if(large_over.Not())
        
        # Large under = more than 4 slots (2 hours) under target
        model.add(under_target >= 4).only_enforce_if(large_under)
        model.add(under_target < 4).only_enforce_if(large_under.Not())
        
        # Apply MASSIVE penalty: -5000 points for each large deviation
        # This is 5x the front desk coverage weight - makes it nearly impossible to have large deviations
        large_deviation_penalty -= 5000 * (large_over + large_under)
    
    # Calculate shift length preference (encourage longer shifts)
    # Prefer fewer, longer shifts (e.g., three 4-hour shifts) over many short shifts (e.g., five 2-hour shifts)
    # We do this by rewarding the total hours worked while penalizing the number of shifts
    shift_length_bonus = 0
    
    for e in employees:
        for d in days:
            # Count if employee works at all this day (this is a "shift day")
            works_this_day = model.new_bool_var(f"works_this_day[{e},{d}]")
            day_slots = sum(work[e, d, t] for t in T)
            
            # Link indicator: works_this_day = 1 if day_slots > 0
            model.add(day_slots >= 1).only_enforce_if(works_this_day)
            model.add(day_slots == 0).only_enforce_if(works_this_day.Not())
            
            # Reward the shift length (more slots per shift = better)
            # But penalize having many shifts (fewer shifts = better)
            # Net effect: encourages longer, fewer shifts
            shift_length_bonus += day_slots  # Reward hours worked
            shift_length_bonus -= 6 * works_this_day  # Strongly penalize number of shifts (6 slots = 3 hours offset)
    
    # Calculate underclassmen front desk preference (SOFT preference)
    # Prefer to put freshmen and sophomores at front desk over juniors and seniors
    # This is NOT a hard constraint - just a gentle nudge when solver has options
    # Scoring: Higher year = penalty for front desk assignment
    # Freshman (1): -1 penalty = prefer them most
    # Sophomore (2): -2 penalty = still good
    # Junior (3): -3 penalty = prefer to avoid
    # Senior (4): -4 penalty = prefer to avoid most
    underclassmen_preference_score = 0
    
    for e in employees:
        year = employee_year.get(e, 2)  # Default to sophomore if not specified
        
        # For each front desk assignment, apply a penalty based on year
        # Lower year (freshman) = smaller penalty = more preferred
        for d in days:
            for t in T:
                if (e, d, t, "front_desk") in assign:
                    # Subtract the year value: freshmen (1) are least penalized
                    underclassmen_preference_score -= year * assign[(e, d, t, "front_desk")]
    
    # ============================================================================
    # DEPARTMENT SCARCITY PENALTY FOR FRONT DESK
    # ============================================================================
    # Prefer pulling people to front desk from departments with MORE qualified people
    # (more options, more flexible scheduling) rather than departments with FEWER
    # This protects scarce resources in small departments (Marketing=2, Employer Engagement=2)
    # and lets bigger departments (Career Education=3, Events=3) contribute more to front desk
    # 
    # Scarcity score: Inverse of department size - smaller departments get higher penalty
    # This takes precedence over seniority - spreading the wealth is the priority!
    
    department_scarcity_penalty = 0
    
    for e in employees:
        # Find which non-front-desk departments this employee belongs to
        employee_departments = [r for r in qual[e] if r != "front_desk" and r in department_roles]
        
        # Calculate scarcity: average inverse of department sizes for this employee's departments
        # If employee is in multiple departments, use the SMALLEST department (most scarce)
        if employee_departments:
            # Get the smallest department size this employee belongs to
            min_dept_size = min(department_sizes[dept] for dept in employee_departments)
            
            # Scarcity penalty: smaller department = higher penalty for using at front desk
            # Dept size 2 = penalty 5 per slot (very scarce - avoid pulling them)
            # Dept size 3 = penalty 3.33 per slot (moderately scarce)
            # Dept size 4+ = penalty 2.5 or less per slot (plenty of people - okay to pull)
            scarcity_factor = 10.0 / min_dept_size
            
            # Apply penalty for each front desk assignment
            for d in days:
                for t in T:
                    if (e, d, t, "front_desk") in assign:
                        # Penalize pulling scarce resources to front desk
                        department_scarcity_penalty -= scarcity_factor * assign[(e, d, t, "front_desk")]
    
    # ============================================================================
    # COLLABORATIVE HOURS TRACKING
    # ============================================================================
    # Track when multiple people work together in the same department (collaboration)
    # We want to ENCOURAGE collaboration - it's good for teamwork and training!
    # Note: Single 30-minute overlaps don't count - must be at least 1 hour together
    
    # Minimum collaborative hours per department
    # These are SOFT targets - we'll try to hit them but won't fail if impossible
    # Collaborative hours = time slots where 2+ people work the same role simultaneously
    # SET YOUR VALUES HERE (in hours, will be converted to 30-min slots):
    min_collaborative_hours = {
        "career_education": 1,      #Set your value
        "marketing": 1,             #Set your value
        "employer_engagement": 2,   #Set your value
        "events": 4,                #Set your value
        "data_systems": 0,          # 0 because only 1 person (Diana)
    }
    
    # Track collaborative slots per department
    collaborative_slots = {}
    for role in department_roles:
        # Count slots where 2+ people work this role simultaneously
        collab_slot_vars = []
        for d in days:
            for t in T:
                # Count how many people are working this department role at this time
                num_in_role = sum(assign.get((e, d, t, role), 0) for e in employees)
                
                # Create a boolean indicator: are there 2+ people in this role right now?
                has_collaboration = model.new_bool_var(f"collab_{role}[{d},{t}]")
                model.add(num_in_role >= 2).only_enforce_if(has_collaboration)
                model.add(num_in_role <= 1).only_enforce_if(has_collaboration.Not())
                
                collab_slot_vars.append(has_collaboration)
        
        # Sum up total collaborative slots for this department
        collaborative_slots[role] = sum(collab_slot_vars)
    
    # Calculate penalty for not meeting collaborative hour minimums
    # This is a SOFT constraint - encourages collaboration but doesn't require it
    collaborative_hours_score = 0
    
    for role in department_roles:
        if role not in min_collaborative_hours:
            continue
        
        min_slots = int(min_collaborative_hours[role] * 2)  # Convert hours to 30-min slots
        
        if min_slots == 0:
            # No collaboration requirement for this department (e.g., data_systems with 1 person)
            continue
        
        # Calculate how far we are from the minimum
        under_collab = model.new_int_var(0, 200, f"under_collab[{role}]")
        model.add(collaborative_slots[role] + under_collab >= min_slots)
        
        # Penalize being under the collaborative minimum
        # Increased penalty to make collaboration a higher priority
        collaborative_hours_score -= under_collab  # Will be multiplied by 200 in objective function
    
    # ============================================================================
    # OFFICE COVERAGE - Encourage at least 2 people in office at all times
    # ============================================================================
    # Track how many people are working (in ANY role) at each time slot
    # We want front desk (1 person) + at least 1 department worker = 2+ total
    # CRITICAL: Having only 1 person (front desk alone) is risky - no backup if they get sick!
    
    office_coverage_score = 0
    single_coverage_penalty = 0  # NEW: penalty for having only 1 person
    
    for d in days:
        for t in T:
            # Count total people working at this time slot (any role)
            total_people = sum(assign.get((e, d, t, r), 0) 
                             for e in employees 
                             for r in roles 
                             if (e, d, t, r) in assign)
            
            # Encourage having at least 2 people in the office
            # Reward each person beyond 1 (so 2 people = +1 bonus, 3 people = +2 bonus, etc.)
            office_coverage_score += total_people - 1
            
            # NEW: Heavy penalty if only 1 person in office (front desk alone - very risky!)
            # Create a boolean variable for "only 1 person working"
            only_one_person = model.new_bool_var(f"only_one_{d}_{t}")
            
            # If total_people == 1, then only_one_person = 1, otherwise 0
            model.add(total_people == 1).only_enforce_if(only_one_person)
            model.add(total_people != 1).only_enforce_if(only_one_person.Not())
            
            # Apply penalty for single coverage
            single_coverage_penalty -= only_one_person  # Will be multiplied by weight in objective
    
    # ============================================================================
    # TIME OF DAY PREFERENCE - Very slight favor toward morning staffing
    # ============================================================================
    # Slightly prefer having more people working in morning hours (8am-12pm)
    # This is a VERY gentle nudge - only matters when everything else is equal
    # Helps avoid scenarios where afternoons are understaffed relative to mornings
    
    morning_preference_score = 0
    morning_slots = [t for t in T if t < 8]  # Slots 0-7 = 8:00am-12:00pm (4 hours)
    
    for d in days:
        for t in morning_slots:
            # Count people working in morning time slots
            morning_workers = sum(assign.get((e, d, t, r), 0) 
                                for e in employees 
                                for r in roles 
                                if (e, d, t, r) in assign)
            morning_preference_score += morning_workers
    
    # Objective: Maximize coverage with priorities:
    # 1. Front desk coverage (weight 10000) - EXTREMELY HIGH PRIORITY - virtually guarantees coverage
    # 2. Large deviation penalty (weight 1) - MASSIVE penalty for being 2+ hours off target (-5000 per person)
    # 3. Department target hours (weight 1000) - DOUBLED to prioritize department hours
    # 4. Collaborative hours (weight 200) - STRONGLY encourage collaboration
    # 5. Office coverage (weight 150) - Encourage 2+ people in office at all times
    # 5b. Single coverage penalty (weight 500) - HEAVILY discourage only 1 person in office (risky!)
    # 6. Target adherence (weight 100) - STRONGLY encourage hitting target hours (graduated by year)
    # 7. Department spread (weight 60) - Prefer departmental presence across many time slots
    # 8. Department day coverage (weight 30) - Encourage each department to appear throughout the week
    # 9. Shift length preference (weight 20) - Gently prefer longer shifts (reduced to allow flexibility)
    # 10. Department scarcity penalty (weight 8) - Prefer pulling from richer departments to front desk
    # 11. Underclassmen at front desk (weight 3) - Moderate preference for freshmen at front desk
    # 12. Morning preference (weight 0.5) - VERY slight favor toward morning staffing (tiebreaker only)
    # 13. Total department hours (weight 1) - Fill available departmental capacity
    # Note: Front desk weight is 10x larger than before - will only be uncovered if IMPOSSIBLE
    #       (i.e., no front-desk-qualified employee available at that time slot)
    model.maximize(
        front_desk_coverage_score +          # Weight 10000 per slot - EXTREMELY high priority
        large_deviation_penalty +            # MASSIVE penalty for 2+ hour deviations (-5000 per person)
        1000 * department_target_score +     # Weight 1000 - DOUBLED to prioritize department hours! (was 500)
        department_large_deviation_penalty + # Severe penalty for large department deviations (-4000)
        200 * collaborative_hours_score +    # Weight 200 - STRONGLY encourage collaboration
        150 * office_coverage_score +        # Weight 150 - encourage 2+ people in office at all times
        500 * single_coverage_penalty +      # Weight 500 - HEAVILY penalize only 1 person (risky!)
        100 * target_adherence_score +       # Strongly encourage target hour adherence
        60 * department_spread_score +
        30 * department_day_coverage_score +
        20 * shift_length_bonus +            # Reduced to allow more flexibility for hour distribution
        8 * department_scarcity_penalty +    # Weight 8 - Protect small departments from over-use at front desk
        3 * underclassmen_preference_score + # Weight 3 - Moderate preference for freshmen at front desk
        0.5 * morning_preference_score +     # Weight 0.5 - VERY gentle morning preference (tiebreaker)
        total_department_assignments         # Fill available departmental capacity
    )
    
    
    # ============================================================================
    # STEP 12: SOLVE THE MODEL
    # ============================================================================
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120  # Increased to 120 seconds for better optimization
    
    print("Solving the scheduling problem...")
    print(f"   - {len(employees)} employees")
    print(f"   - {len(days)} days")
    print(f"   - {len(T)} time slots per day")
    print(f"   - {len(assign)} assignment variables")
    print()
    
    # Track total execution time
    start_time = time.time()
    status = solver.solve(model)
    end_time = time.time()
    total_time = end_time - start_time
    
    
    # ============================================================================
    # STEP 13: DISPLAY THE RESULTS
    # ============================================================================
    
    print_schedule(
        status,
        solver,
        employees,
        days,
        T,
        SLOT_NAMES,
        qual,
        work,
        assign,
        weekly_hour_limits,
        target_weekly_hours,
        total_time,
        roles,
        department_roles,
        ROLE_DISPLAY_NAMES,
        department_hour_targets,
        department_max_hours,
    )
    export_schedule_to_excel(
        status,
        solver,
        employees,
        days,
        T,
        SLOT_NAMES,
        qual,
        work,
        assign,
        weekly_hour_limits,
        target_weekly_hours,
        roles,
        department_roles,
        ROLE_DISPLAY_NAMES,
        department_hour_targets,
        department_max_hours,
        args.output,
    )


# ============================================================================
# PRETTY PRINTING FUNCTIONS
# ============================================================================

def print_schedule(status, solver, employees, days, T, SLOT_NAMES, qual, work, assign, weekly_hour_limits, target_weekly_hours, total_time, roles, department_roles, role_display_names, department_hour_targets, department_max_hours):
    """
    Display the schedule in a readable format with statistics
    
    Args:
        status: Solution status from the solver
        solver: The CP-SAT solver instance
        employees: List of employee names
        days: List of day names
        T: List of time slot indices
        SLOT_NAMES: Human-readable names for time slots
        qual: Employee qualifications dictionary
        work: Work assignment variables
        assign: Role assignment variables
        weekly_hour_limits: Dictionary of weekly hour limits per employee
        target_weekly_hours: Dictionary of target hours per employee
        total_time: Total wall-clock time for solving
        roles: List of all roles including front desk
        department_roles: List of non-front desk roles
        role_display_names: Friendly names for printing per role
        department_hour_targets: Target hours per department role
        department_max_hours: Maximum hours per department role
    """
    
    print("\n" + "=" * 120)
    print(f"SCHEDULE STATUS: {status}")
    print("=" * 120)
    
    # Check if we found a valid solution
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        print("No solution found!")
        print("\nPossible reasons:")
        print("  - Constraints are too restrictive")
        print("  - Not enough qualified employees")
        print("  - Availability conflicts with coverage requirements")
        return
    
    # Print solver statistics
    print(f"\nSolution found!")
    print(f"\nSolver Statistics:")
    print(f"  - Total execution time: {total_time:.2f} seconds")
    print(f"  - Solver computation time: {solver.wall_time:.2f} seconds")
    print(f"  - Branches explored: {solver.num_branches:,}")
    print(f"  - Conflicts encountered: {solver.num_conflicts:,}")
    
    
    # ========================================================================
    # SECTION 1: DAILY SCHEDULE GRID
    # ========================================================================
    
    for d in days:
        print(f"\n{'─' * 120}")
        print(f"{d.upper()}")
        print(f"{'─' * 120}")
        
        # Header row
        role_columns = ["front_desk"] + department_roles
        column_width = 22
        header = f"\n{'Time':<12}" + "".join(f"{role_display_names[role]:<{column_width}}" for role in role_columns)
        print(header)
        print("─" * (12 + column_width * len(role_columns)))
        
        # Data rows - one per time slot
        for t in T:
            time_slot = SLOT_NAMES[t]
            
            row = f"{time_slot:<12}"
            for role in role_columns:
                workers = [
                    e for e in employees
                    if (e, d, t, role) in assign
                    and solver.value(assign[(e, d, t, role)])
                ]
                
                if role == "front_desk":
                    cell = ", ".join(workers) if workers else "ERROR: UNCOVERED"
                else:
                    cell = ", ".join(workers) if workers else "-"
                
                row += f"{cell:<{column_width}}"
            
            print(row)
    
    
    # ========================================================================
    # SECTION 2: EMPLOYEE SUMMARY
    # ========================================================================
    
    print(f"\n{'=' * 120}")
    print("EMPLOYEE SUMMARY")
    print(f"{'=' * 120}\n")
    
    print(f"{'Employee':<15}{'Qualifications':<35}{'Hours (Target/Max)':<30}{'Days Worked'}")
    print("─" * 120)
    
    for e in employees:
        total_slots = 0
        days_worked = []
        
        # Calculate slots worked each day (each slot = 0.5 hours)
        for d in days:
            day_slots = sum(solver.value(work[e, d, t]) for t in T)
            if day_slots > 0:
                day_hours = day_slots * 0.5  # Convert slots to hours
                days_worked.append(f"{d}({day_hours:.1f}h)")
                total_slots += day_slots
        
        # Format employee info
        quals = ", ".join(sorted(qual[e]))
        days_str = ", ".join(days_worked) if days_worked else "None"
        
        # Get targets and limits
        target_hours = target_weekly_hours.get(e, 11)
        weekly_limit = weekly_hour_limits.get(e, 40)
        total_hours = total_slots * 0.5
        
        # Show actual vs target vs max
        hours_str = f"{total_hours:.1f} (↑{target_hours}/max {weekly_limit})"
        
        # Add indicator if they hit their target
        if abs(total_hours - target_hours) <= 0.5:  # Within 30 minutes
            hours_str = f"✓ {hours_str}"
        
        print(f"{e:<15}{quals:<35}{hours_str:<30}{days_str}")
    
    
    # ========================================================================
    # SECTION 3: ROLE DISTRIBUTION STATISTICS
    # ========================================================================
    
    print(f"\n{'=' * 120}")
    print("ROLE DISTRIBUTION")
    print(f"{'=' * 120}\n")
    
    role_totals = {role: 0 for role in roles}
    
    for d in days:
        role_counts = {role: 0 for role in roles}
        
        # Count assignments for this day (in 30-minute slots)
        for t in T:
            for e in employees:
                for role in roles:
                    if (e, d, t, role) in assign and solver.value(assign[(e, d, t, role)]):
                        role_counts[role] += 1
        
        for role in roles:
            role_totals[role] += role_counts[role]
        
        day_summary = ", ".join(
            f"{role_display_names[role]} {role_counts[role] * 0.5:.1f}h"
            for role in roles
            if role_counts[role] > 0
        ) or "No assignments"
        print(f"{d}: {day_summary}")
    
    print("\nTOTAL HOURS BY ROLE")
    print("─" * 90)
    print(f"{'Role':<25}{'Actual':<12}{'Target':<12}{'Max':<12}{'Delta':<12}{'Status'}")
    print("─" * 90)
    
    for role in roles:
        actual_hours = role_totals[role] * 0.5
        target = department_hour_targets.get(role)
        max_hours = department_max_hours.get(role)
        
        # Format columns
        role_name = role_display_names[role]
        actual_str = f"{actual_hours:.1f}h"
        target_str = f"{target:.1f}h" if target is not None else "-"
        max_str = f"{max_hours:.1f}h" if max_hours is not None else "-"
        
        # Calculate delta and status
        if target is not None:
            delta = actual_hours - target
            delta_str = f"{delta:+.1f}h"
            
            # Determine status with emoji/symbol
            if abs(delta) <= 1.0:  # Within 1 hour
                status = "✓ On Target"
            elif delta > 0:
                status = "↑ Over"
            else:
                status = "↓ Under"
        else:
            delta_str = "-"
            status = "-"
        
        print(f"{role_name:<25}{actual_str:<12}{target_str:<12}{max_str:<12}{delta_str:<12}{status}")
    
    print("─" * 90)
    
    print("=" * 120 + "\n")


def export_schedule_to_excel(
    status,
    solver,
    employees,
    days,
    T,
    SLOT_NAMES,
    qual,
    work,
    assign,
    weekly_hour_limits,
    target_weekly_hours,
    roles,
    department_roles,
    role_display_names,
    department_hour_targets,
    department_max_hours,
    output_path: Path,
):
    """
    Export the generated schedule to an Excel workbook with formatted sheets.
    
    This function runs after console printing to avoid impacting scheduling logic.
    """
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        return
    
    # Resolve writer engine dynamically
    role_columns = [FRONT_DESK_ROLE] + department_roles

    # Daily tables and weekly rollup
    daily_tables = []
    weekly_rows = []
    role_headers = [role_display_names[role] for role in role_columns]
    weekly_columns = ["Day", "Time"] + role_headers

    for day in days:
        day_rows = []
        for t in T:
            cell_values = []
            for role in role_columns:
                workers = [
                    e
                    for e in employees
                    if (e, day, t, role) in assign and solver.value(assign[(e, day, t, role)])
                ]
                cell_values.append(", ".join(workers) if workers else ("UNCOVERED" if role == FRONT_DESK_ROLE else ""))
            day_rows.append([SLOT_NAMES[t], *cell_values])
            weekly_rows.append([day, SLOT_NAMES[t], *cell_values])
        daily_tables.append((f"{day} Schedule", ["Time"] + role_headers, day_rows))
    
    # Employee summary data
    summary_rows = []
    for e in employees:
        total_slots = 0
        days_worked = []
        for d in days:
            day_slots = sum(solver.value(work[e, d, t]) for t in T)
            if day_slots > 0:
                total_slots += day_slots
                days_worked.append(f"{d}({day_slots * 0.5:.1f}h)")
        target_hours = target_weekly_hours.get(e, 0)
        max_hours = weekly_hour_limits.get(e, 0)
        total_hours = total_slots * 0.5
        hit_target = abs(total_hours - target_hours) <= 0.5
        summary_rows.append(
            [
                e,
                ", ".join(sorted(qual[e])),
                total_hours,
                target_hours,
                max_hours,
                "✓" if hit_target else "",
                ", ".join(days_worked) if days_worked else "None",
            ]
        )
    summary_columns = [
        "Employee",
        "Qualifications",
        "Hours Worked",
        "Target Hours",
        "Max Hours",
        "Hit Target",
        "Days Worked",
    ]
    
    # Role distribution data
    distribution_rows = []
    role_totals = {role: 0 for role in roles}
    for d in days:
        row = [d]
        for role in roles:
            slot_count = sum(
                solver.value(assign[(e, d, t, role)]) if (e, d, t, role) in assign else 0
                for e in employees
                for t in T
            )
            role_totals[role] += slot_count
            row.append(slot_count * 0.5)
        distribution_rows.append(row)
    total_row = ["TOTAL"] + [role_totals[r] * 0.5 for r in roles]
    distribution_rows.append(total_row)
    distribution_columns = ["Day"] + [role_display_names[role] for role in roles]
    
    dept_summary_headers = [
        "Department",
        "Actual Hours",
        "Target Hours",
        "Max Hours",
        "Delta (Actual-Target)",
    ]
    dept_summary_rows = []
    for role in department_roles:
        actual_hours = role_totals[role] * 0.5
        target = department_hour_targets.get(role)
        max_hours = department_max_hours.get(role)
        delta = actual_hours - target if target is not None else ""
        dept_summary_rows.append([
            role_display_names[role],
            actual_hours,
            target if target is not None else "",
            max_hours if max_hours is not None else "",
            delta,
        ])
    
    engine = None
    for candidate in ("xlsxwriter", "openpyxl"):
        if importlib.util.find_spec(candidate):
            engine = candidate
            break
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    if engine:
        with pd.ExcelWriter(output_path, engine=engine) as writer:
            if weekly_rows:
                df_weekly = pd.DataFrame(weekly_rows, columns=weekly_columns)
                df_weekly.to_excel(writer, sheet_name="Weekly Schedule", index=False)
                _autosize_columns(writer, "Weekly Schedule", df_weekly)
            for sheet_name, columns, rows in daily_tables:
                df_day = pd.DataFrame(rows, columns=columns)
                df_day.to_excel(writer, sheet_name=sheet_name, index=False)
                _autosize_columns(writer, sheet_name, df_day)
            
            df_summary = pd.DataFrame(summary_rows, columns=summary_columns)
            df_summary.to_excel(writer, sheet_name="Employee Summary", index=False)
            _autosize_columns(writer, "Employee Summary", df_summary)
            
            df_distribution = pd.DataFrame(distribution_rows, columns=distribution_columns)
            df_distribution.to_excel(writer, sheet_name="Role Distribution", index=False)
            _autosize_columns(writer, "Role Distribution", df_distribution)
            if dept_summary_rows:
                df_dept = pd.DataFrame(dept_summary_rows, columns=dept_summary_headers)
                df_dept.to_excel(writer, sheet_name="Department Targets", index=False)
                _autosize_columns(writer, "Department Targets", df_dept)
    else:
        sheets_payload = []
        if weekly_rows:
            sheets_payload.append(("Weekly Schedule", [weekly_columns] + weekly_rows))
        for sheet_name, columns, rows in daily_tables:
            sheets_payload.append((sheet_name, [columns] + rows))
        sheets_payload.append(("Employee Summary", [summary_columns] + summary_rows))
        sheets_payload.append(("Role Distribution", [distribution_columns] + distribution_rows))
        if dept_summary_rows:
            sheets_payload.append(("Department Targets", [dept_summary_headers] + dept_summary_rows))
        _write_minimal_xlsx(output_path, sheets_payload)
    
    print(f"Schedule exported to {output_path}")


def _autosize_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame):
    """Automatic column width helper for Excel export."""
    worksheet = writer.sheets[sheet_name]
    engine = getattr(writer, "engine", "").lower()
    for idx, column in enumerate(dataframe.columns):
        series = dataframe[column].astype(str)
        max_length = max(series.map(len).max(), len(str(column)))
        width = min(max_length + 2, 60)
        if engine == "xlsxwriter":
            worksheet.set_column(idx, idx, width)
        elif engine == "openpyxl":
            from openpyxl.utils import get_column_letter

            worksheet.column_dimensions[get_column_letter(idx + 1)].width = width


def _write_minimal_xlsx(output_path: Path, sheets: List[Tuple[str, List[List]]]):
    """Fallback XLSX writer using only the Python standard library."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from xml.sax.saxutils import escape

    def col_letter(index: int) -> str:
        result = ""
        index += 1
        while index:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def build_sheet_xml(rows: List[List]) -> str:
        cells_xml = []
        for row_idx, row in enumerate(rows, start=1):
            cell_parts = []
            for col_idx, value in enumerate(row):
                if value in (None, ""):
                    continue
                cell_ref = f"{col_letter(col_idx)}{row_idx}"
                if isinstance(value, (int, float)):
                    cell_parts.append(f'<c r="{cell_ref}"><v>{value}</v></c>')
                else:
                    text = escape(str(value))
                    cell_parts.append(
                        f'<c r="{cell_ref}" t="inlineStr"><is><t>{text}</t></is></c>'
                    )
            cell_xml = "".join(cell_parts)
            row_xml = f'<row r="{row_idx}">{cell_xml}</row>' if cell_xml else f'<row r="{row_idx}"/>'
            cells_xml.append(row_xml)
        sheet_body = "".join(cells_xml)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheetData>{sheet_body}</sheetData>"
            "</worksheet>"
        )

    sanitized_sheets = []
    for idx, (name, rows) in enumerate(sheets, start=1):
        sheet_name = name[:31] if name else f"Sheet{idx}"
        sanitized_sheets.append((sheet_name, rows))

    workbook_rels = []
    sheets_entries = []
    content_types_overrides = []
    sheet_files = []

    for idx, (name, rows) in enumerate(sanitized_sheets, start=1):
        sheet_filename = f"sheet{idx}.xml"
        rel_id = f"rId{idx}"
        workbook_rels.append(
            f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{sheet_filename}"/>'
        )
        sheets_entries.append(f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="{rel_id}"/>')
        content_types_overrides.append(
            f'<Override PartName="/xl/worksheets/{sheet_filename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        sheet_files.append((f"xl/worksheets/{sheet_filename}", build_sheet_xml(rows)))

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f'{"".join(content_types_overrides)}'
        "</Types>"
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheets_entries)}</sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(workbook_rels)}'
        "</Relationships>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
        "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "</styleSheet>"
    )

    with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", rels_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/styles.xml", styles_xml)
        for filename, xml in sheet_files:
            archive.writestr(filename, xml)

# ============================================================================
# PROGRAM ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()
