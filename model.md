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

#### **4. Department Target Hours**

**Weight: 500**

- **Why**: Departments need adequate coverage to meet operational needs
- **How**: Bonus points for getting departments closer to target hours
- **Purpose**: Prioritize departmental needs over individual preferences
- **Context**: Increased from 100 to 500 to address chronic department shortfalls

#### **5. Individual Target Adherence**

**Weight: 100 (graduated by seniority)**

- **Why**: Students want to work their requested hours for income
- **How**: Year-based multiplier encourages hitting targets:
  - **Freshman (year 1)**: 100 × 1.5 = 150 points
  - **Sophomore (year 2)**: 100 × 1.25 = 125 points
  - **Junior (year 3)**: 100 × 1.0 = 100 points
  - **Senior (year 4)**: 100 × 0.8 = 80 points
- **Purpose**: Slight preference to help younger students who may need income more

#### **6. Department Spread**

**Weight: 60**

- **Why**: Better to have departments active throughout the day
- **How**: Rewards departments appearing in many different time slots
- **Example**: Marketing active 8-10am, 12-2pm, 3-5pm is better than just 8am-2pm
- **Purpose**: Ensures departments accessible to students/employers throughout day

#### **7. Collaborative Hours**

**Weight: 50**

- **Why**: Encourage teamwork and training opportunities
- **How**: Penalty of 50 points for each hour under minimum collaborative target

- **Requirements**: 
  - Career Education: 1 hour minimum of 2+ people working together
  - Marketing: 1 hour minimum
  - Internships: 1 hour minimum
  - Employer Engagement: 2 hours minimum
  - Events: 4 hours minimum
  - Data Systems: 0 (only Diana qualified)
- **Note**: Must be sustained overlaps (single 30-min slot doesn't count)

#### **8. Department Day Coverage**

**Weight: 30**

- **Why**: Better for departments to be available multiple days/week
- **How**: Rewards departments working across different days
- **Example**: Career Ed on Mon/Wed/Fri is better than Mon/Tue/Wed
- **Purpose**: Improves service accessibility throughout week

#### **9. Shift Length Preference**

**Weight: 20**

- **Why**: Longer, fewer shifts are more efficient than many short shifts
- **How**:

  - Rewards each slot worked (+1 per slot)
  - Penalizes each shift day (-6 per day worked)
  - Net effect: 4-hour shift (8 slots - 6 penalty = +2) beats two 2-hour shifts (8 slots - 12 penalty = -4)
- **Purpose**: Reduce context switching and commute inefficiency

#### **10. Department Scarcity Penalty**

**Weight: 2**

- **Why**: Protect scarce resources in understaffed departments
- **How**: Penalty for pulling employees to front desk based on their department size
  - **2-person dept** (Marketing, Employer Engagement): 10/2 = **5 penalty/slot** → avoid pulling them
  - **3-person dept** (Career Ed, Events): 10/3 = **3.3 penalty/slot** → okay to pull
  - Uses employee's **smallest** department if qualified for multiple
- **Purpose**: Preferentially use employees from "richer" departments for front desk, protecting departments with limited options
- **Priority**: Takes precedence over seniority - spreading the wealth matters more

#### **11. Underclassmen Front Desk Preference**

**Weight: 0.5**

- **Why**: Gentle preference for younger students at front desk
- **How**: Year-based penalty per front desk slot:
  - **Freshman (1)**: -1 penalty = most preferred
  - **Sophomore (2)**: -2 penalty
  - **Junior (3)**: -3 penalty
  - **Senior (4)**: -4 penalty = least preferred
- **Purpose**: Very mild nudge when solver has equivalent options
- **Note**: Lowest priority - only matters when everything else is equal

#### **12. Total Department Assignments**

**Weight: 1**

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
- Employer Engagement target: **+500** for keeping her in EE
- Individual target: **+100** for hitting her hours

**Result:** Front desk wins (10,000 >> 500+100-10), but the -10 penalty makes the solver prefer using someone from Career Education if available.

---

## Current Challenges

### 1. Employer Engagement Shortfall

- **Target**: 23 hours
- **Capacity**: 23 hours (Eve 11h + Frank 12h)
- **Actual**: ~15.5 hours
- **Why**: Only 2 people qualified, both also needed for front desk coverage
- **Solution**: This is a **structural capacity issue**. Target may need reduction or more people need to be cross-trained.

### 2. Marketing Improvement

- **Before scarcity penalty**: 15h (8h under target)
- **After scarcity penalty**: 19.5h (3.5h under target)
- **Why improved**: Scarcity penalty protects Iris and Jack from excessive front desk duty

### 3. Front Desk vs. Department Trade-off

- Front desk weight (10,000) dominates department targets (500)
- This is intentional - service coverage is critical
- Result: Dual-qualified employees lean toward front desk

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
- Configurable weights in objective function (lines ~985-1010)
- Minimum collaborative hours dictionary (lines ~897-904)

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

To adjust behavior, modify weights in the objective function (line ~997):

```python
model.maximize(
    10000 * front_desk_coverage_score +    # Decrease if front desk steals too many people
    500 * department_target_score +        # Increase to prioritize dept targets more
    2 * department_scarcity_penalty +      # Increase to stronger protect small depts
    # ... etc
)
```

**Rule of thumb**: Weight ratios matter more than absolute values. Front desk at 10,000 is 20× more important than dept targets at 500.
