# Semester Scheduler

## Overview

Constraint-programming workflow that builds optimal weekly rosters for Chapman University’s Career & Professional Development student employees. The model uses Google OR-Tools CP-SAT to balance front desk coverage, departmental workload targets, individual availability, and hour preferences.

## Key Features

- Guarantees continuous front desk coverage with single-employee hand-offs.
- Enforces contiguous daily shifts within configurable min/max lengths.
- Respects detailed employee availability, qualifications, and weekly hour limits.
- Encourages each department (Events, Marketing, Employer Engagement, Internships, Career Education, Data & Systems) to hit desired weekly hour totals.
- Produces human-readable console output and an auto-generated `schedule.xlsx` workbook (daily tabs, weekly rollup, employee summary, role distribution).

## Requirements

- Python 3.12 (see `.python-version`)
- Packages listed in `requirements.txt` (includes `ortools`, `pandas`, `numpy`, etc.)

## Usage

### This uses dummy data included

```bash
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows
pip install -r requirements.txt
python main.py
```

The solver prints statistics and per-day grids; once finished, open `schedule.xlsx` for the formatted weekly schedule.
