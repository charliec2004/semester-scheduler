# CPD Semester Scheduler - Model Documentation

## Overview

This scheduler uses **Google OR-Tools CP-SAT Solver** (Constraint Programming) to generate optimal weekly schedules for Career and Professional Development student employees. The model balances multiple competing priorities through a weighted objective function while enforcing strict operational constraints.

---

## Time Structure

- **Time slots**: 30-minute increments (8:00 AM - 5:00 PM)
- **Slots per day**: 18 slots
- **Days**: Monday through Friday
- **Hours calculation**: Each slot = 0.5 hours

---

## Hard Constraints

These are **absolute requirements** that must be satisfied for any valid schedule:

### 1. Front Desk Maximum Capacity

- **Maximum 1 person** at front desk at any time
- Cannot have 2+ people (prevents resource waste)
- **Note**: Minimum coverage (at least 1 person) is handled as a soft constraint - see Priority #1 below

### 2. Employee Availability

- Employees can only work during their available time slots
- Set via availability matrix in `employees.csv`

### 3. Role Qualifications

- Employees can only be assigned to roles they're qualified for
- Qualifications defined in `employees.csv`

### 4. Hour Limits

- **Universal maximum**: 19 hours/week (institutional policy)
- **Personal maximum**: Individual preference (e.g., 11-13 hours)
- Cannot exceed the lesser of the two

### 5. Shift Contiguity

- **One continuous block per day** - no split shifts
- If working at all, must work minimum **2 hours** (4 slots)
- Maximum **4 hours** (8 slots) per shift
- Once an employee stops working, they cannot start again that day

### 6. Role Duration Minimum

- Each role assignment must last **at least 1 hour** (2 consecutive slots)
- Prevents toggling between roles (e.g., can't do marketing for 30 min, switch to events for 30 min, back to marketing)
- Ensures meaningful work blocks

### 7. Single Assignment

- Each employee can only work **one role at a time**
- Cannot simultaneously work multiple departments

---

## Soft Constraints (Objective Function)

These are **preferences** optimized through weighted scoring. Higher weight = higher priority.

### Priority Hierarchy (Highest to Lowest)

#### **1. Front Desk Coverage**

**Weight: 10,000 per slot**

- **Why**: Critical service - office must have someone at front desk
- **How**: Massive bonus for covering each time slot
- **Note**: While technically "soft," the 10,000 weight virtually guarantees coverage unless physically impossible (no qualified employee available)

#### **2. Large Hour Deviations**

**Weight: -5,000 per person**

- **Why**: Severe penalty for missing individual target hours by 2+ hours
- **Trigger**: `|actual_hours - target_hours| ≥ 2.0`
- **How**: Binary penalty - either you're within 2 hours or you get hit hard
- **Purpose**: Prevent extreme under/over-scheduling

#### **3. Department Large Deviations**

**Weight: -4,000 per department**

- **Why**: Departments need consistent staffing to function
- **Trigger**: Department is 4+ hours under target
- **How**: Severe penalty for big departmental shortfalls
- **Purpose**: Ensure departments aren't critically understaffed

#### 4. Department Target Hours

Weight: 1000

- **Why**: Departments need adequate coverage to meet operational needs
- **How**: Bonus points for getting departments closer to target hours
- **Purpose**: Prioritize departmental needs over individual preferences
- **Context**: Increased from 100 → 500 → **1000** to address chronic department shortfalls and ensure no department is more than -4 hours under target

#### 5. Collaborative Hours

Weight: 200

- **Why**: Encourage teamwork and training opportunities
- **How**: Penalty for each slot under minimum collaborative target
- **Requirements**:
  - Career Education: 1 hour minimum of 2+ people working together in same role
  - Marketing: 1 hour minimum
  - Internships: 1 hour minimum
  - Employer Engagement: 2 hours minimum
  - Events: 4 hours minimum
  - Data Systems: 0 (only Diana qualified)
- **Note**: Must be sustained overlaps (single 30-min slot doesn't count)
- **Context**: Increased from 50 → **200** to make collaboration a higher priority

#### 6. Office Coverage

Weight: 150

- **Why**: Prevent lonely mornings with only front desk coverage
- **How**: Rewards having 2+ people in the office at any given time
  - 1 person (just front desk) = 0 bonus
  - 2 people (front desk + 1 dept) = +150 points
  - 3 people (front desk + 2 dept) = +300 points
  - etc.
- **Purpose**: Ensure office feels active and staffed throughout the day, not just minimal coverage
- **Context**: Added to address empty mornings where only front desk is present

#### 7. Individual Target Adherence

Weight: 100 (graduated by seniority)

- **Why**: Students want to work their requested hours for income
- **How**: Year-based multiplier encourages hitting targets:
  - **Freshman (year 1)**: 100 × 1.5 = 150 points
  - **Sophomore (year 2)**: 100 × 1.25 = 125 points
  - **Junior (year 3)**: 100 × 1.0 = 100 points
  - **Senior (year 4)**: 100 × 0.8 = 80 points
- **Purpose**: Slight preference to help younger students who may need income more

#### 8. Department Spread

Weight: 60

- **Why**: Better to have departments active throughout the day
- **How**: Rewards departments appearing in many different time slots
- **Example**: Marketing active 8-10am, 12-2pm, 3-5pm is better than just 8am-2pm
- **Purpose**: Ensures departments accessible to students/employers throughout day

#### 9. Department Day Coverage

Weight: 30

- **Why**: Better for departments to be available multiple days/week
- **How**: Rewards departments working across different days
- **Example**: Career Ed on Mon/Wed/Fri is better than Mon/Tue/Wed
- **Purpose**: Improves service accessibility throughout week

#### 10. Shift Length Preference

Weight: 20

- **Why**: Longer, fewer shifts are more efficient than many short shifts
- **How**:
  - Rewards each slot worked (+1 per slot)
  - Penalizes each shift day (-6 per day worked)
  - Net effect: 4-hour shift (8 slots - 6 penalty = +2) beats two 2-hour shifts (8 slots - 12 penalty = -4)
- **Purpose**: Reduce context switching and commute inefficiency

#### 11. Department Scarcity Penalty

Weight: 2

- **Why**: Protect scarce resources in understaffed departments
- **How**: Penalty for pulling employees to front desk based on their department size
  - **2-person dept** (Marketing, Employer Engagement): 10/2 = **5 penalty/slot** → avoid pulling them
  - **3-person dept** (Career Ed, Events): 10/3 = **3.3 penalty/slot** → okay to pull
  - Uses employee's **smallest** department if qualified for multiple
- **Purpose**: Preferentially use employees from "richer" departments for front desk, protecting departments with limited options
- **Priority**: Takes precedence over seniority - spreading the wealth matters more

#### 12. Underclassmen Front Desk Preference

Weight: 0.5

- **Why**: Gentle preference for younger students at front desk
- **How**: Year-based penalty per front desk slot:
  - **Freshman (1)**: -1 penalty = most preferred
  - **Sophomore (2)**: -2 penalty
  - **Junior (3)**: -3 penalty
  - **Senior (4)**: -4 penalty = least preferred
- **Purpose**: Very mild nudge when solver has equivalent options
- **Note**: Lowest priority - only matters when everything else is equal

#### 13. Total Department Assignments

Weight: 1

- **Why**: Fill available capacity rather than leave it unused
- **How**: Small bonus for each department hour worked
- **Purpose**: Tiebreaker - use available hours when possible

---

## How Conflicts Are Resolved

When objectives compete, the model prioritizes by weight:

### Example Scenario

**Eve** is qualified for Employer Engagement (2-person dept) and Front Desk.

**Competing forces:**

- Front desk coverage: **+10,000** for using Eve
- Department scarcity: **-10** penalty (2 × 5 per slot = -10 for 2 slots)
- Employer Engagement target: **+1000** for keeping her in EE (increased from 500)
- Individual target: **+100** for hitting her hours

**Result:** Front desk wins (10,000 >> 1000+100-10), but the -10 penalty makes the solver prefer using someone from Career Education if available.

---

## Current Status & Improvements

### Recent Optimizations

1. **Department Target Weight**: Increased from 100 → 500 → **1000**
   - **Impact**: Departments now prioritized much higher, reducing shortfalls
   - **Goal**: Keep all departments within -4 hours of target

2. **Collaborative Hours Weight**: Increased from 50 → **200**
   - **Impact**: More instances of 2+ people working same role simultaneously
   - **Goal**: Hit minimum collaborative hour targets per department

3. **Office Coverage Addition**: New constraint with weight **150**
   - **Impact**: Prevents empty mornings with only front desk present
   - **Goal**: Always have 2+ people in office (front desk + at least 1 department worker)

4. **Department Scarcity Penalty**: Weight **2**
   - **Impact**: Protects small departments (Marketing, EE) from excessive front desk duty
   - **Goal**: Spread front desk coverage to employees from larger departments

### Known Challenges

1. **Employer Engagement Expansion**
   - Grace added to EE roster (now 3 people: Eve, Frank, Grace)
   - Increased capacity from 23h → 35h
   - Improved coverage but created Events shortfall

2. **Events Staffing**
   - Only 3 qualified people (Alice, Bob, Charlie)
   - All also work front desk, limiting department hours
   - Target: 27h | Capacity: 36h | Actual: varies based on front desk needs

---

## Model Statistics

- **Employees**: 13
- **Departments**: 6 + Front Desk
- **Assignment variables**: ~2,340
- **Solve time**: ~60 seconds
- **Solver**: CP-SAT (Constraint Programming)

---

## Configuration Files

### `employees.csv`

- Employee names, qualifications, target/max hours, year
- 90 availability columns (5 days × 18 slots)
- `1` = available, `0` = unavailable

### `cpd-requirements.csv`

- Department target and maximum hours
- Used to set departmental staffing goals

### `main.py`

- Constraint programming model
- Configurable weights in objective function (lines ~1010-1035)
- Minimum collaborative hours dictionary (lines ~933-940)
- Office coverage scoring (lines ~984-1002)

---

## Interpreting Output

### Schedule Grid

- Rows = time slots
- Columns = roles
- Cells = employee names working that role
- `-` = no one assigned
- Multiple names = collaboration (multiple people working same role)

### Employee Summary

- **✓** = hit target hours exactly
- **↑** = target hours achieved
- **Days Worked** = hours per day breakdown

### Role Distribution

- Shows hours per role per day
- **TOTAL HOURS BY ROLE** section shows:
  - **Actual**: Hours scheduled
  - **Target**: Goal hours
  - **Delta**: Difference (negative = under target)
  - **Status**: ✓ On Target | ↓ Under | ↑ Over

---

## Tuning the Model

To adjust behavior, modify weights in the objective function (line ~1010):

```python
model.maximize(
    10000 * front_desk_coverage_score +    # Extremely high - virtually guarantees coverage
    1000 * department_target_score +       # Increased to 1000 to keep depts within -4h of target
    200 * collaborative_hours_score +      # Increased to 200 to encourage teamwork
    150 * office_coverage_score +          # New - encourages 2+ people in office at all times
    100 * target_adherence_score +         # Individual hour targets
    # ... etc
)
```

**Rule of thumb**: Weight ratios matter more than absolute values. Front desk at 10,000 is 10× more important than dept targets at 1000.

### Current Weight Hierarchy

1. **Front Desk Coverage**: 10,000 (dominant priority)
2. **Large Penalties**: -5,000 (individual), -4,000 (department) for major deviations
3. **Department Targets**: 1,000 (doubled to reduce shortfalls)
4. **Collaborative Hours**: 200 (quadrupled to encourage teamwork)
5. **Office Coverage**: 150 (new - prevent empty mornings)
6. **Individual Targets**: 100 (base level)
7. **Department Spread**: 60
8. **Department Day Coverage**: 30
9. **Shift Length**: 20
10. **Department Scarcity**: 2
11. **Underclassmen Preference**: 0.5
12. **Total Assignments**: 1 (tiebreaker)
