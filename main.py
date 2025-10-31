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

from ortools.sat.python import cp_model
from typing import Dict, List, Set, Tuple
from pathlib import Path
import pandas as pd
import importlib.util
import time


def main():
    """Main function to build and solve the scheduling model"""
    
    # ============================================================================
    # STEP 1: INITIALIZE THE CONSTRAINT PROGRAMMING MODEL
    # ============================================================================
    
    model = cp_model.CpModel()
    
    
    # ============================================================================
    # STEP 2: DEFINE THE PROBLEM DOMAIN
    # ============================================================================
    
    # List of all employees in the system
    employees = [
        "Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", 
        "Grace", "Henry", "Iris", "Jack", "Kelly", "Leo",
        "Olivia"
    ]
    
    # Days of the week we're scheduling for
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]  # Weekdays only - no weekend shifts
    
    # Available job roles
    roles = [
        "front_desk",
        "career_education",
        "marketing",
        "internships",
        "employer_engagement",
        "events",
        "data_systems",
    ]
    department_roles = [role for role in roles if role != "front_desk"]
    department_sizes = {}
    ROLE_DISPLAY_NAMES = {
        "front_desk": "Front Desk",
        "career_education": "Career Education",
        "marketing": "Marketing",
        "internships": "Internships",
        "employer_engagement": "Employer Engagement",
        "events": "Events",
        "data_systems": "Data & Systems",
    }
    
    # Time slot configuration - 30 MINUTE INCREMENTS
    # T represents 18 half-hour time slots from 8am to 5pm
    # Index 0 = 8:00-8:30, Index 1 = 8:30-9:00, Index 2 = 9:00-9:30, ..., Index 17 = 4:30-5:00
    T = list(range(18))  # 0-17 for 18 thirty-minute slots
    SLOT_NAMES = [
        "8:00-8:30", "8:30-9:00", "9:00-9:30", "9:30-10:00",
        "10:00-10:30", "10:30-11:00", "11:00-11:30", "11:30-12:00",
        "12:00-12:30", "12:30-1:00", "1:00-1:30", "1:30-2:00",
        "2:00-2:30", "2:30-3:00", "3:00-3:30", "3:30-4:00",
        "4:00-4:30", "4:30-5:00"
    ]
    
    # Shift length constraints (in 30-minute increments)
    # Each slot = 0.5 hours, so multiply by 2 to get slot counts
    MIN_SLOTS = 4   # Minimum 2 hours = 4 thirty-minute slots
    MAX_SLOTS = 8   # Maximum 4 hours = 8 thirty-minute slots (changed from 12)
    MIN_FRONT_DESK_SLOTS = MIN_SLOTS  # Front desk duty must last at least full shift minimum
    
    # Weekly hour limits per employee (customizable)
    # Individual personal maximum preferences (different from universal 19-hour limit)
    # These are HARD constraints - employee cannot exceed this limit
    # Note: These are in HOURS, not slots (will be converted internally)
    weekly_hour_limits = {
        "Alice":   12,  # Personal maximum preference
        "Bob":     11,  # Personal maximum preference
        "Charlie": 13,  # Can work more
        "Diana":   12,  # Personal maximum preference
        "Eve":     11,  # Personal maximum preference
        "Frank":   12,  # Personal maximum preference
        "Grace":   12,  # Personal maximum preference
        "Henry":   11,  # Personal maximum preference
        "Iris":    12,  # Personal maximum preference
        "Jack":    11,  # Personal maximum preference
        "Kelly":   11,  # Personal maximum preference
        "Leo":     12,  # Personal maximum preference
        "Olivia":  12,  # Personal maximum preference
    }
    
    # Employee year classification (1=Freshman, 2=Sophomore, 3=Junior, 4=Senior)
    # Used for graduated preferences (e.g., underclassmen preference for front desk)
    employee_year = {
        "Alice":   2,  # Sophomore
        "Bob":     1,  # Freshman
        "Charlie": 3,  # Junior
        "Diana":   2,  # Sophomore
        "Eve":     3,  # Junior
        "Frank":   2,  # Sophomore
        "Grace":   1,  # Freshman
        "Henry":   4,  # Senior
        "Iris":    2,  # Sophomore
        "Jack":    3,  # Junior
        "Kelly":   1,  # Freshman
        "Leo":     4,  # Senior
        "Olivia":  2,  # Sophomore
    }
    
    # Target weekly hours (preferred hours, may differ from max)
    # Note: 19 hours is the universal maximum for everyone (enforced separately)
    
    # Employee year in school (for front desk preference)
    # 1 = Freshman (highest priority for front desk)
    # 2 = Sophomore (high priority)
    # 3 = Junior (lower priority)
    # 4 = Senior (lowest priority - prefer them for other work)
    # This creates a soft preference: underclassmen preferred at front desk
    employee_year = {
        "Alice":   2,  # Sophomore
        "Bob":     1,  # Freshman
        "Charlie": 1,  # Freshman
        "Diana":   2,  # Sophomore
        "Eve":     3,  # Junior
        "Frank":   2,  # Sophomore
        "Grace":   1,  # Freshman
        "Henry":   4,  # Senior
        "Iris":    2,  # Sophomore
        "Jack":    3,  # Junior
        "Kelly":   1,  # Freshman
        "Leo":     2,  # Sophomore
        "Olivia":  1,  # Freshman
    }
    
    # Target weekly hours per employee (encouraged via objective function)
    # These are "soft" goals - the solver tries to get close to these values
    # Customizable per employee based on their preferences
    # Note: 19 hours is the legal maximum safeguard (enforced separately)
    target_weekly_hours = {
        "Alice":   12,  # Target hours
        "Bob":     11,  # Target hours
        "Charlie": 13,  # Target hours (higher)
        "Diana":   11,  # Target hours
        "Eve":     11,  # Target hours
        "Frank":   12,  # Target hours
        "Grace":   11,  # Target hours
        "Henry":   11,  # Target hours
        "Iris":    12,  # Target hours
        "Jack":    11,  # Target hours
        "Kelly":   11,  # Target hours
        "Leo":     12,  # Target hours
        "Olivia":  11,  # Target hours
    }
    department_hour_targets = {
        "career_education": 28,
        "marketing": 30,
        "internships": 16,
        "employer_engagement": 26,
        "events": 27,
        "data_systems": target_weekly_hours["Diana"],  # Single-person dept matches Diana
    }
    department_hour_threshold = 4  # +/- hours acceptable window
    
    
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
    # STEP 4: DEFINE EMPLOYEE QUALIFICATIONS
    # ============================================================================
    
    # Each employee can only work roles they're qualified for
    # Format: {employee_name: {set of roles they can perform}}
    qual = {
        "Alice":   {"front_desk", "events"},
        "Bob":     {"front_desk", "events"},
        "Charlie": {"front_desk", "events"},
        "Diana":   {"data_systems"},
        "Eve":     {"front_desk", "employer_engagement"},
        "Frank":   {"front_desk", "employer_engagement"},
        "Grace":   {"front_desk", "internships"},
        "Henry":   {"front_desk", "internships"},
        "Iris":    {"front_desk", "marketing"},
        "Jack":    {"front_desk", "marketing"},
        "Kelly":   {"front_desk", "career_education"},
        "Leo":     {"front_desk", "career_education"},
        "Olivia":  {"front_desk", "career_education"},
    }
    department_sizes = {
        role: sum(1 for e in employees if role in qual[e])
        for role in department_roles
    }
    
    
    # ============================================================================
    # STEP 5: DEFINE EMPLOYEE AVAILABILITY CONSTRAINTS
    # ============================================================================
    
    # Format: {employee: {day: [list of unavailable time slots]}}
    # College student schedules with typical patterns:
    # - MWF classes (Monday/Wednesday/Friday same times)
    # - TR classes (Tuesday/Thursday same times)
    # - Lunch breaks (varied times: 10:30-11:00, 11:00-11:30, 12:00-1:00, 1:00-1:30, 1:30-2:00)
    # - More availability on Fridays
    # Each slot is 30 minutes: 0=8:00, 2=9:00, 4=10:00, 5=10:30, 6=11:00, 7=11:30, 8=12:00, 9=12:30, 10=1:00, 11=1:30, 12=2:00, 14=3:00, 16=4:00
    
    unavailable = {
        "Alice": {
            # MWF: Chemistry 9:00-10:30am, English 2:00-3:30pm
            # TR: Math 10:00-11:30am
            # Lunch: 10:30-11:00am daily (early lunch)
            "Mon": [2, 3, 4, 5, 12, 13, 14, 15],  # Chemistry, lunch, English
            "Tue": [4, 5, 6, 7],                   # Math, lunch
            "Wed": [2, 3, 4, 5, 12, 13, 14, 15],  # Chemistry, lunch, English
            "Thu": [4, 5, 6, 7],                   # Math, lunch
            "Fri": [5],                            # Just lunch - more availability!
        },
        
        "Bob": {
            # TR: Biology Lab 8:00-11:00am, History 2:00-3:30pm
            # MWF: Philosophy 11:00-12:00pm
            # Lunch: 1:00-1:30pm (late lunch)
            "Mon": [6, 7, 10, 11],                       # Philosophy, lunch
            "Tue": [0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15],  # Bio lab, lunch, History
            "Wed": [6, 7, 10, 11],                       # Philosophy, lunch
            "Thu": [0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15],  # Bio lab, lunch, History
            "Fri": [10, 11],                             # Lunch
        },
        
        "Charlie": {
            # MWF: Calculus 10:00-11:00am, CS 1:00-2:30pm
            # TR: Physics 9:00-11:00am
            # Lunch: 11:30-12:00pm
            "Mon": [4, 5, 7, 10, 11, 12, 13],      # Calculus, lunch, CS
            "Tue": [2, 3, 4, 5, 6, 7],              # Physics, lunch
            "Wed": [4, 5, 7, 10, 11, 12, 13],      # Calculus, lunch, CS
            "Thu": [2, 3, 4, 5, 6, 7],              # Physics, lunch
            "Fri": [7],                             # Just lunch
        },
        
        "Diana": {
            # TR: Art Studio 8:00-12:00pm, Stats 3:00-4:30pm
            # MWF: Sociology 9:30-11:00am
            # Lunch: 12:00-12:30pm
            "Mon": [3, 4, 5, 6, 7, 8, 9],               # Sociology, lunch
            "Tue": [0, 1, 2, 3, 4, 5, 6, 7, 14, 15, 16, 17],  # Art, Stats
            "Wed": [3, 4, 5, 6, 7, 8, 9],               # Sociology, lunch
            "Thu": [0, 1, 2, 3, 4, 5, 6, 7, 14, 15, 16, 17],  # Art, Stats
            "Fri": [8, 9],                               # Lunch
        },
        
        "Eve": {
            # MWF: Business 8:30-10:00am, Marketing 3:00-4:00pm
            # TR: Accounting 11:00-1:00pm (includes lunch)
            # Lunch: 1:30-2:00pm MWF
            "Mon": [1, 2, 3, 4, 11, 12, 14, 15],    # Business, lunch, Marketing
            "Tue": [6, 7, 8, 9, 10, 11],            # Accounting (no separate lunch)
            "Wed": [1, 2, 3, 4, 11, 12, 14, 15],    # Business, lunch, Marketing
            "Thu": [6, 7, 8, 9, 10, 11],            # Accounting
            "Fri": [11, 12],                        # Lunch
        },
        
        "Frank": {
            # TR: Engineering 8:00-10:30am, Lab 2:00-4:30pm
            # MWF: Economics 10:00-11:30am
            # Lunch: 11:00-11:30am
            "Mon": [4, 5, 6, 7],                    # Economics, lunch
            "Tue": [0, 1, 2, 3, 4, 5, 6, 7, 12, 13, 14, 15, 16, 17],  # Engineering, lunch, Lab
            "Wed": [4, 5, 6, 7],                    # Economics, lunch
            "Thu": [0, 1, 2, 3, 4, 5, 6, 7, 12, 13, 14, 15, 16, 17],  # Engineering, lunch, Lab
            "Fri": [6, 7],                          # Lunch
        },
        
        "Grace": {
            # MWF: Psychology 9:00-10:30am
            # TR: Spanish 10:30-12:00pm
            # Lunch: 12:30-1:00pm
            "Mon": [2, 3, 4, 5, 9, 10],             # Psychology, lunch
            "Tue": [5, 6, 7, 9, 10],                # Spanish, lunch
            "Wed": [2, 3, 4, 5, 9, 10],             # Psychology, lunch
            "Thu": [5, 6, 7, 9, 10],                # Spanish, lunch
            "Fri": [9, 10],                         # Lunch
        },
        
        "Henry": {
            # TR: Computer Science 9:30-11:30am, Seminar 1:00-2:00pm
            # MWF: Writing 11:00-12:30pm
            # Lunch: 12:00-12:30pm MWF, 12:30-1:00pm TR
            "Mon": [6, 7, 8, 9],                    # Writing, lunch
            "Tue": [3, 4, 5, 6, 7, 9, 10, 11],      # CS, lunch, Seminar
            "Wed": [6, 7, 8, 9],                    # Writing, lunch
            "Thu": [3, 4, 5, 6, 7, 9, 10, 11],      # CS, lunch, Seminar
            "Fri": [8, 9],                          # Lunch
        },
        
        "Iris": {
            # MWF: Dance 8:00-9:30am, Music 3:30-5:00pm
            # TR: Theater 10:00-12:00pm
            # Lunch: 1:00-1:30pm
            "Mon": [0, 1, 2, 3, 10, 11, 15, 16, 17],  # Dance, lunch, Music
            "Tue": [4, 5, 6, 7, 10, 11],              # Theater, lunch
            "Wed": [0, 1, 2, 3, 10, 11, 15, 16, 17],  # Dance, lunch, Music
            "Thu": [4, 5, 6, 7, 10, 11],              # Theater, lunch
            "Fri": [10, 11],                          # Lunch
        },
        
        "Jack": {
            # TR: Biology 8:30-10:30am, Chemistry 2:30-4:30pm
            # MWF: Geology 1:00-2:30pm
            # Lunch: 11:00-11:30am
            "Mon": [6, 7, 10, 11, 12, 13],          # Lunch, Geology
            "Tue": [1, 2, 3, 4, 5, 6, 7, 13, 14, 15, 16, 17],  # Biology, lunch, Chemistry
            "Wed": [6, 7, 10, 11, 12, 13],          # Lunch, Geology
            "Thu": [1, 2, 3, 4, 5, 6, 7, 13, 14, 15, 16, 17],  # Biology, lunch, Chemistry
            "Fri": [6, 7],                          # Lunch
        },
        
        "Kelly": {
            # MWF: Literature 10:30-12:00pm
            # TR: Anthropology 9:00-10:30am, Political Science 2:00-3:30pm
            # Lunch: 1:30-2:00pm
            "Mon": [5, 6, 7, 11, 12],               # Literature, lunch
            "Tue": [2, 3, 4, 5, 11, 12, 12, 13, 14, 15],  # Anthro, lunch, PoliSci
            "Wed": [5, 6, 7, 11, 12],               # Literature, lunch
            "Thu": [2, 3, 4, 5, 11, 12, 12, 13, 14, 15],  # Anthro, lunch, PoliSci
            "Fri": [11, 12],                        # Lunch
        },
        
        "Leo": {
            # TR: Statistics 8:00-9:30am, Data Science 3:00-5:00pm
            # MWF: Programming 11:30-1:00pm (includes lunch)
            # Lunch: 12:00-12:30pm (within class time MWF)
            "Mon": [7, 8, 9, 10, 11],               # Programming (includes lunch)
            "Tue": [0, 1, 2, 3, 8, 9, 14, 15, 16, 17],  # Stats, lunch, Data Science
            "Wed": [7, 8, 9, 10, 11],               # Programming (includes lunch)
            "Thu": [0, 1, 2, 3, 8, 9, 14, 15, 16, 17],  # Stats, lunch, Data Science
            "Fri": [8, 9],                          # Lunch
        },
        
        "Olivia": {
            # MWF: Environmental Science 9:00-10:00am, Lab 3:00-4:30pm
            # TR: Ecology 11:00-12:30pm
            # Lunch: 1:00-1:30pm
            "Mon": [2, 3, 10, 11, 14, 15, 16, 17],  # Env Sci, lunch, Lab
            "Tue": [6, 7, 8, 9, 10, 11],            # Ecology, lunch
            "Wed": [2, 3, 10, 11, 14, 15, 16, 17],  # Env Sci, lunch, Lab
            "Thu": [6, 7, 8, 9, 10, 11],            # Ecology, lunch
            "Fri": [10, 11],                        # Lunch
        },
    }
    
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
        max_weekly_slots = max_weekly_hours * 2  # Convert hours to 30-minute slots
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
            
            total_front_desk_slots = sum(assign[(e, d, t, "front_desk")] for t in T)
            works_front_desk_today = model.new_bool_var(f"works_front_desk_today[{e},{d}]")
            model.add(total_front_desk_slots >= 1).only_enforce_if(works_front_desk_today)
            model.add(total_front_desk_slots == 0).only_enforce_if(works_front_desk_today.Not())
            model.add(total_front_desk_slots >= MIN_FRONT_DESK_SLOTS).only_enforce_if(works_front_desk_today)
    
    
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
            
            # IMPORTANT: Weight coverage based on time of day
            # Earlier hours are MORE important than later hours
            # This way if coverage must be dropped, it happens at end of day first
            time_weight = 18 - t  # Earlier slots get higher weight (18, 17, 16, ... 1)
            front_desk_coverage_score += time_weight * has_front_desk
            
            # HARD CONSTRAINT: At most 1 front desk at a time (no overstaffing)
            model.add(num_front_desk <= 1)
            
            # OPTIONAL CAP: keep departmental staffing reasonable relative to membership size
            for role in department_roles:
                model.add(
                    sum(assign.get((e, d, t, role), 0) for e in employees) <= department_sizes[role]
                )
    
    
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
        adjusted_target_hours = min(target_hours, max_capacity_hours)
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
    
    # Objective: Maximize coverage with priorities:
    # 1. Front desk coverage (weight 1000) - CRITICAL but soft, prioritizes early hours
    # 2. Large deviation penalty (weight 1) - MASSIVE penalty for being 2+ hours off target (-5000 per person)
    # 3. Target adherence (weight 100) - STRONGLY encourage hitting target hours (graduated by year)
    # 4. Department spread (weight 60) - Prefer departmental presence across many time slots
    # 5. Department day coverage (weight 30) - Encourage each department to appear throughout the week
    # 6. Department hour targets (weight 100) - Encourage departments to hit target hours
    # 7. Shift length preference (weight 20) - Gently prefer longer shifts (reduced to allow flexibility)
    # 8. Underclassmen at front desk (weight 0.5) - VERY gentle nudge when all else equal
    # 9. Total department hours (weight 1) - Fill available departmental capacity
    # Note: Front desk coverage is heavily weighted but NOT a hard constraint
    #       If impossible to cover all hours, later hours drop first (due to time_weight)
    model.maximize(
        1000 * front_desk_coverage_score +   # Prioritize front desk coverage with time weighting
        large_deviation_penalty +            # MASSIVE penalty for 2+ hour deviations (-5000 per person)
        100 * target_adherence_score +       # Strongly encourage target hour adherence
        60 * department_spread_score +
        30 * department_day_coverage_score +
        100 * department_target_score +
        20 * shift_length_bonus +            # Reduced to allow more flexibility for hour distribution
        0.5 * underclassmen_preference_score + # VERY gentle - only matters when everything else equal
        total_department_assignments +
        department_large_deviation_penalty
    )
    
    
    # ============================================================================
    # STEP 12: SOLVE THE MODEL
    # ============================================================================
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60  # Increased to 60 seconds for better optimization
    
    print("🔄 Solving the scheduling problem...")
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
    )


# ============================================================================
# PRETTY PRINTING FUNCTIONS
# ============================================================================

def print_schedule(status, solver, employees, days, T, SLOT_NAMES, qual, work, assign, weekly_hour_limits, target_weekly_hours, total_time, roles, department_roles, role_display_names):
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
    """
    
    print("\n" + "=" * 120)
    print(f"SCHEDULE STATUS: {status}")
    print("=" * 120)
    
    # Check if we found a valid solution
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        print("❌ No solution found!")
        print("\nPossible reasons:")
        print("  - Constraints are too restrictive")
        print("  - Not enough qualified employees")
        print("  - Availability conflicts with coverage requirements")
        return
    
    # Print solver statistics
    print(f"\n✅ Solution found!")
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
        print(f"📅 {d.upper()}")
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
                    cell = ", ".join(workers) if workers else "❌ UNCOVERED"
                else:
                    cell = ", ".join(workers) if workers else "-"
                
                row += f"{cell:<{column_width}}"
            
            print(row)
    
    
    # ========================================================================
    # SECTION 2: EMPLOYEE SUMMARY
    # ========================================================================
    
    print(f"\n{'=' * 120}")
    print("👥 EMPLOYEE SUMMARY")
    print(f"{'=' * 120}\n")
    
    print(f"{'Employee':<12}{'Qualifications':<25}{'Hours (Target/Max)':<25}{'Days Worked'}")
    print("─" * 105)
    
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
        
        print(f"{e:<12}{quals:<25}{hours_str:<25}{days_str}")
    
    
    # ========================================================================
    # SECTION 3: ROLE DISTRIBUTION STATISTICS
    # ========================================================================
    
    print(f"\n{'=' * 120}")
    print("📊 ROLE DISTRIBUTION")
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
    for role in roles:
        print(f" - {role_display_names[role]}: {role_totals[role] * 0.5:.1f} hours")
    
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
    output_path: Path = Path("schedule.xlsx"),
):
    """
    Export the generated schedule to an Excel workbook with formatted sheets.
    
    This function runs after console printing to avoid impacting scheduling logic.
    """
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        return
    
    # Resolve writer engine dynamically
    role_columns = ["front_desk"] + department_roles
    
    # Daily tables
    daily_tables = []
    for day in days:
        columns = ["Time"] + [role_display_names[role] for role in role_columns]
        rows = []
        for t in T:
            row = [SLOT_NAMES[t]]
            for role in role_columns:
                workers = [
                    e
                    for e in employees
                    if (e, day, t, role) in assign and solver.value(assign[(e, day, t, role)])
                ]
                cell_value = ", ".join(workers) if workers else ("UNCOVERED" if role == "front_desk" else "")
                row.append(cell_value)
            rows.append(row)
        daily_tables.append((f"{day} Schedule", columns, rows))
    
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
    
    engine = None
    for candidate in ("xlsxwriter", "openpyxl"):
        if importlib.util.find_spec(candidate):
            engine = candidate
            break
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    if engine:
        with pd.ExcelWriter(output_path, engine=engine) as writer:
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
    else:
        sheets_payload = []
        for sheet_name, columns, rows in daily_tables:
            sheets_payload.append((sheet_name, [columns] + rows))
        sheets_payload.append(("Employee Summary", [summary_columns] + summary_rows))
        sheets_payload.append(("Role Distribution", [distribution_columns] + distribution_rows))
        _write_minimal_xlsx(output_path, sheets_payload)
    
    print(f"📁 Schedule exported to {output_path}")


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
