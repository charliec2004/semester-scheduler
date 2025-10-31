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
- Restockers can only work when a front_desk is present
- Multiple employees can work the same role simultaneously
- Each employee works one continuous block per day (3-6 hours)
"""

from ortools.sat.python import cp_model
from typing import Dict, List, Set, Tuple


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
    roles = ["front_desk", "Restocker"]
    
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
    # 4 = Senior (lowest priority - prefer them for restocking)
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
    
    
    # ============================================================================
    # STEP 3: DEFINE COVERAGE REQUIREMENTS
    # ============================================================================
    
    # Initialize demand dictionary: demand[role][day][time_slot]
    # Value of 1 means "we need 1 person in this role at this time"
    # Value of 0 means "no requirement for this role at this time"
    demand = {
        role: {
            day: [0] * len(T) for day in days
        } for role in roles
    }
    
    # front_desk coverage is CRITICAL - must be present at all times
    for day in days:
        for time_slot in T:
            demand["front_desk"][day][time_slot] = 1
    
    # Note: Restockers have no fixed demand - they're assigned flexibly
    # based on availability and the objective function
    
    
    # ============================================================================
    # STEP 4: DEFINE EMPLOYEE QUALIFICATIONS
    # ============================================================================
    
    # Each employee can only work roles they're qualified for
    # Format: {employee_name: {set of roles they can perform}}
    qual = {
        "Alice":   {"front_desk"},                    # front_desk specialist
        "Bob":     {"front_desk", "Restocker"},       # Cross-trained
        "Charlie": {"front_desk", "Restocker"},       # Cross-trained
        "Diana":   {"Restocker"},                  # Restocker specialist
        "Eve":     {"front_desk"},                    # front_desk specialist
        "Frank":   {"Restocker"},                  # Restocker specialist
        "Grace":   {"front_desk", "Restocker"},       # Cross-trained
        "Henry":   {"front_desk"},                    # front_desk specialist
        "Iris":    {"Restocker"},                  # Restocker specialist
        "Jack":    {"front_desk", "Restocker"},       # Cross-trained
        "Kelly":   {"front_desk"},                    # front_desk specialist
        "Leo":     {"Restocker"},                  # Restocker specialist
        "Olivia":  {"Restocker"},                  # Restocker specialist
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
                
                # Constraint 9.3: CRITICAL - Restockers need front_desk supervision
                # A restocker can ONLY work when at least one front_desk is present
                # This prevents scenarios where only restockers are working
                if (e, d, t, "Restocker") in assign:
                    model.add(
                        sum(assign.get((emp, d, t, "front_desk"), 0) for emp in employees) >= 1
                    ).only_enforce_if(assign[(e, d, t, "Restocker")])
    
    
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
            
            # RESTOCKER COVERAGE: At most 3 restockers at any time
            # This prevents overcrowding while allowing flexibility
            model.add(
                sum(assign.get((e, d, t, "Restocker"), 0) for e in employees) <= 3
            )
    
    
    # ============================================================================
    # STEP 11: DEFINE THE OBJECTIVE FUNCTION
    # ============================================================================
    # What we're trying to optimize (maximize in this case)
    
    # Count total restocker assignments across all employees, days, and times
    restocker_assignments = sum(
        assign.get((e, d, t, "Restocker"), 0) 
        for e in employees 
        for d in days 
        for t in T
    )
    
    # Calculate "spread" metric: count how many time slots have at least 1 restocker
    # This encourages distribution throughout the day rather than clustering
    restocker_spread_score = 0
    for d in days:
        for t in T:
            # Create indicator: does this time slot have any restockers?
            has_restocker = model.new_bool_var(f"has_restocker[{d},{t}]")
            num_restockers = sum(assign.get((e, d, t, "Restocker"), 0) for e in employees)
            
            # Link indicator to actual restocker count
            model.add(num_restockers >= 1).only_enforce_if(has_restocker)
            model.add(num_restockers == 0).only_enforce_if(has_restocker.Not())
            
            restocker_spread_score += has_restocker
    
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
    # 4. Restocker spread (weight 50) - Prefer many time slots with restockers
    # 5. Shift length preference (weight 20) - Gently prefer longer shifts (reduced to allow flexibility)
    # 6. Underclassmen at front desk (weight 0.5) - VERY gentle nudge when all else equal
    # 7. Total restockers (weight 1) - Fill available capacity
    # Note: Front desk coverage is heavily weighted but NOT a hard constraint
    #       If impossible to cover all hours, later hours drop first (due to time_weight)
    model.maximize(
        1000 * front_desk_coverage_score +   # Prioritize front desk coverage with time weighting
        large_deviation_penalty +            # MASSIVE penalty for 2+ hour deviations (-5000 per person)
        100 * target_adherence_score +       # Strongly encourage target hour adherence
        50 * restocker_spread_score +
        20 * shift_length_bonus +            # Reduced to allow more flexibility for hour distribution
        0.5 * underclassmen_preference_score + # VERY gentle - only matters when everything else equal
        restocker_assignments
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
    
    status = solver.solve(model)
    
    
    # ============================================================================
    # STEP 13: DISPLAY THE RESULTS
    # ============================================================================
    
    print_schedule(
        status, solver, employees, days, T, SLOT_NAMES, 
        qual, work, assign, weekly_hour_limits, target_weekly_hours
    )


# ============================================================================
# PRETTY PRINTING FUNCTIONS
# ============================================================================

def print_schedule(status, solver, employees, days, T, SLOT_NAMES, qual, work, assign, weekly_hour_limits, target_weekly_hours):
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
    print(f"  - Solve time: {solver.wall_time:.2f} seconds")
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
        print(f"\n{'Time':<12}", end="")
        print(f"{'front_desks':<40}{'Restockers':<40}")
        print("─" * 92)
        
        # Data rows - one per time slot
        for t in T:
            time_slot = SLOT_NAMES[t]
            
            # Find all front_desks working this slot
            front_desks = [
                e for e in employees
                if (e, d, t, "front_desk") in assign 
                and solver.value(assign[(e, d, t, "front_desk")])
            ]
            
            # Find all restockers working this slot
            restockers = [
                e for e in employees
                if (e, d, t, "Restocker") in assign 
                and solver.value(assign[(e, d, t, "Restocker")])
            ]
            
            # Format the output
            front_desk_str = ", ".join(front_desks) if front_desks else "❌ UNCOVERED"
            restocker_str = ", ".join(restockers) if restockers else "-"
            
            print(f"{time_slot:<12}{front_desk_str:<40}{restocker_str:<40}")
    
    
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
    
    total_front_desk_slots = 0
    total_restocker_slots = 0
    
    for d in days:
        front_desk_count = 0
        restocker_count = 0
        
        # Count assignments for this day (in 30-minute slots)
        for t in T:
            for e in employees:
                if (e, d, t, "front_desk") in assign and solver.value(assign[(e, d, t, "front_desk")]):
                    front_desk_count += 1
                if (e, d, t, "Restocker") in assign and solver.value(assign[(e, d, t, "Restocker")]):
                    restocker_count += 1
        
        total_front_desk_slots += front_desk_count
        total_restocker_slots += restocker_count
        
        # Convert to hours for display
        front_desk_hours = front_desk_count * 0.5
        restocker_hours = restocker_count * 0.5
        
        print(f"{d}: {front_desk_hours:.1f} front_desk-hours, {restocker_hours:.1f} restocker-hours")
    
    # Convert totals to hours
    total_front_desk_hours = total_front_desk_slots * 0.5
    total_restocker_hours = total_restocker_slots * 0.5
    
    print(f"\n{'TOTAL:':<4} {total_front_desk_hours:.1f} front_desk-hours, {total_restocker_hours:.1f} restocker-hours across the week")
    print("=" * 120 + "\n")

# ============================================================================
# PROGRAM ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()