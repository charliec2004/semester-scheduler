from __future__ import annotations

import sys
import time
from pathlib import Path
from typing import Dict, List, Set

from ortools.sat.python import cp_model

from scheduler.config import (
    COLLABORATION_MINIMUM_HOURS,
    DAY_NAMES,
    DEFAULT_SOLVER_MAX_TIME,
    DEPARTMENT_HOUR_THRESHOLD,
    DEPARTMENT_LARGE_DEVIATION_PENALTY,
    DEPARTMENT_SCARCITY_BASE_WEIGHT,
    EMPLOYEE_LARGE_DEVIATION_PENALTY,
    FRONT_DESK_COVERAGE_WEIGHT,
    FRONT_DESK_ROLE,
    LARGE_DEVIATION_SLOT_THRESHOLD,
    MAX_SLOTS,
    MIN_FRONT_DESK_SLOTS,
    MIN_SLOTS,
    OBJECTIVE_WEIGHTS,
    SHIFT_LENGTH_DAILY_COST,
    SLOT_NAMES,
    T_SLOTS,
    YEAR_TARGET_MULTIPLIERS,
)
from scheduler.data_access.department_loader import load_department_requirements
from scheduler.data_access.staff_loader import load_staff_data
from scheduler.reporting.console import print_schedule
from scheduler.reporting.export import export_schedule_to_excel


def solve_schedule(
    staff_csv: Path,
    requirements_csv: Path,
    output_path: Path,
    solver_max_time: int = DEFAULT_SOLVER_MAX_TIME,
):
    """Main function to build and solve the scheduling model"""

    staff_data = load_staff_data(staff_csv)
    department_requirements = load_department_requirements(requirements_csv)
    department_hour_targets_raw = department_requirements.targets
    department_max_hours_raw = department_requirements.max_hours
    
    # ============================================================================
    # STEP 1: INITIALIZE THE CONSTRAINT PROGRAMMING MODEL
    # ============================================================================
    
    model = cp_model.CpModel()
    
    
    # ============================================================================
    # STEP 2: DEFINE THE PROBLEM DOMAIN
    # ============================================================================
    
    employees: List[str] = staff_data.employees
    qual: Dict[str, Set[str]] = staff_data.qual
    weekly_hour_limits = {emp: float(hours) for emp, hours in staff_data.weekly_hour_limits.items()}
    target_weekly_hours = {emp: float(hours) for emp, hours in staff_data.target_weekly_hours.items()}
    employee_year = {emp: int(year) for emp, year in staff_data.employee_year.items()}
    unavailable: Dict[str, Dict[str, List[int]]] = staff_data.unavailable
    
    days = DAY_NAMES[:]
    roles = list(staff_data.roles)
    if FRONT_DESK_ROLE not in roles:
        raise ValueError(f"Role '{FRONT_DESK_ROLE}' is required but missing from staff data.")
    roles = [FRONT_DESK_ROLE] + [role for role in roles if role != FRONT_DESK_ROLE]
    department_roles = [role for role in roles if role != FRONT_DESK_ROLE]
    
    missing_targets = [role for role in department_roles if role not in department_hour_targets_raw]
    missing_max = [role for role in department_roles if role not in department_max_hours_raw]
    if missing_targets:
        raise ValueError(f"Department targets missing for: {', '.join(missing_targets)}")
    if missing_max:
        raise ValueError(f"Department max hours missing for: {', '.join(missing_max)}")
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
        raise ValueError(
            "No qualified employees found for departments: " + ", ".join(zero_capacity_departments)
        )
    
    ROLE_DISPLAY_NAMES = {
        role: " ".join(word.capitalize() for word in role.split("_"))
        for role in roles
    }
    ROLE_DISPLAY_NAMES[FRONT_DESK_ROLE] = "Front Desk"
    
    # Time slot configuration
    T = T_SLOTS
    
    
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
            front_desk_coverage_score += FRONT_DESK_COVERAGE_WEIGHT * has_front_desk
            
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
    front_desk_slots_by_employee = {
        e: sum(assign.get((e, d, t, FRONT_DESK_ROLE), 0) for d in days for t in T)
        for e in employees
    }
    dual_front_desk_slots = {
        role: sum(front_desk_slots_by_employee[e] for e in employees if role in qual[e])
        for role in department_roles
    }
    department_effective_units = {
        role: 2 * department_assignments[role] + dual_front_desk_slots[role]
        for role in department_roles
    }
    total_department_units = sum(department_effective_units.values())
    department_max_units = {
        role: int(round(department_max_hours[role] * 4))
        for role in department_roles
    }
    for role in department_roles:
        model.add(department_effective_units[role] <= department_max_units[role])
    
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
    threshold_units = department_hour_threshold * 4

    for role in department_roles:
        target_hours = department_hour_targets.get(role)
        if target_hours is None:
            continue
        max_capacity_hours = sum(weekly_hour_limits.get(e, 0) for e in employees if role in qual[e])
        max_requirement_hours = department_max_hours.get(role, max_capacity_hours)
        adjusted_target_hours = min(target_hours, max_capacity_hours, max_requirement_hours)
        target_units = int(adjusted_target_hours * 4)
        total_role_units = department_effective_units[role]

        over = model.new_int_var(0, 400, f"department_over[{role}]")
        under = model.new_int_var(0, 400, f"department_under[{role}]")
        model.add(total_role_units == target_units + over - under)

        department_target_score -= over + under

        if threshold_units > 0:
            large_over = model.new_bool_var(f"department_large_over[{role}]")
            large_under = model.new_bool_var(f"department_large_under[{role}]")

            model.add(over >= threshold_units).only_enforce_if(large_over)
            model.add(over < threshold_units).only_enforce_if(large_over.Not())

            model.add(under >= threshold_units).only_enforce_if(large_under)
            model.add(under < threshold_units).only_enforce_if(large_under.Not())

            department_large_deviation_penalty -= DEPARTMENT_LARGE_DEVIATION_PENALTY * (large_over + large_under)
    
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
        year_multiplier = YEAR_TARGET_MULTIPLIERS.get(year, 1.0)
        
        # Penalize deviation with graduated weight
        # Upperclassmen deviations are penalized more heavily
        target_adherence_score -= year_multiplier * (over_target + under_target)
        
        # STEEP PENALTY for large deviations (2+ hours = 4+ slots off target)
        # This applies to EVERYONE regardless of year
        # Create indicator variables for "large deviation"
        large_over = model.new_bool_var(f"large_over[{e}]")
        large_under = model.new_bool_var(f"large_under[{e}]")
        
        # Large over = more than threshold slots (default 4 -> 2 hours)
        model.add(over_target >= LARGE_DEVIATION_SLOT_THRESHOLD).only_enforce_if(large_over)
        model.add(over_target < LARGE_DEVIATION_SLOT_THRESHOLD).only_enforce_if(large_over.Not())
        
        # Large under = more than threshold slots under target
        model.add(under_target >= LARGE_DEVIATION_SLOT_THRESHOLD).only_enforce_if(large_under)
        model.add(under_target < LARGE_DEVIATION_SLOT_THRESHOLD).only_enforce_if(large_under.Not())
        
        # Apply MASSIVE penalty for large deviations
        large_deviation_penalty -= EMPLOYEE_LARGE_DEVIATION_PENALTY * (large_over + large_under)
    
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
            shift_length_bonus -= SHIFT_LENGTH_DAILY_COST * works_this_day  # Penalize number of distinct shifts
    
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
            scarcity_factor = DEPARTMENT_SCARCITY_BASE_WEIGHT / min_dept_size
            
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
        if role not in COLLABORATION_MINIMUM_HOURS:
            continue
        
        min_slots = int(COLLABORATION_MINIMUM_HOURS[role] * 2)  # Convert hours to 30-min slots
        
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
        front_desk_coverage_score +             # Massive weighting baked into coverage score itself
        large_deviation_penalty +               # MASSIVE penalty for 2+ hour deviations (-5000 per person)
        OBJECTIVE_WEIGHTS.department_target * department_target_score +
        department_large_deviation_penalty +    # Severe penalty for large department deviations
        OBJECTIVE_WEIGHTS.collaborative_hours * collaborative_hours_score +
        OBJECTIVE_WEIGHTS.office_coverage * office_coverage_score +
        OBJECTIVE_WEIGHTS.single_coverage * single_coverage_penalty +
        OBJECTIVE_WEIGHTS.target_adherence * target_adherence_score +
        OBJECTIVE_WEIGHTS.department_spread * department_spread_score +
        OBJECTIVE_WEIGHTS.department_day_coverage * department_day_coverage_score +
        OBJECTIVE_WEIGHTS.shift_length * shift_length_bonus +
        OBJECTIVE_WEIGHTS.department_scarcity * department_scarcity_penalty +
        OBJECTIVE_WEIGHTS.underclassmen_front_desk * underclassmen_preference_score +
        OBJECTIVE_WEIGHTS.morning_preference * morning_preference_score +
        OBJECTIVE_WEIGHTS.department_total * total_department_units
    )
    
    
    # ============================================================================
    # STEP 12: SOLVE THE MODEL
    # ============================================================================
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = solver_max_time
    
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
        output_path,
    )
    return status
