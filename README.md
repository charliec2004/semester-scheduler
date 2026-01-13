# Semester Scheduler

[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![OR-Tools](https://img.shields.io/badge/OR--Tools-9.0+-green.svg)](https://developers.google.com/optimization)

## Overview

Automated scheduling system that builds optimal weekly rosters for Chapman University's Career & Professional Development student employees using constraint programming (Google OR-Tools CP-SAT).

## Why This Project?

**Problem**: Manually scheduling 13+ employees across 6 departments with varying availability took days of planning each semester and produced suboptimal coverage.

**Solution**: Constraint programming optimizer that:

- Reduces scheduling time from hours to **under 2 minutes**
- Guarantees **100% front desk coverage** (8am-5pm, Mon-Fri) with backups
- Optimizes departmental staffing within target goals
- Balances 2,340+ decision variables across 13 competing priorities

**Technical Highlights**: Multi-objective optimization, sophisticated constraint satisfaction, handles complex edge cases (minimum shift lengths, role transitions, resource scarcity), exports professional Excel schedules.

## Project Structure

``` structure
semester-scheduler/
├── main.py                   # Core scheduling engine (CP-SAT model)
├── requirements.txt          # Dependencies (ortools, pandas, openpyxl)
├── LICENSE                   # MIT License
├── model.md                  # Detailed constraint documentation
├── employees.csv             # Employee data & availability
├── cpd-requirements.csv      # Department targets
├── tests/                    # Test suite (pytest)
│   ├── test_data_loading.py
│   └── test_constraints.py
└── schedule.xlsx             # Generated output
```

## How It Works

1. **Input**: CSV files with employee availability (90 time slots/week) and department targets
2. **Model**: CP-SAT solver with 2,340+ variables, 15+ hard constraints, 13 weighted objectives
3. **Optimize**: Maximizes weighted objective (front desk coverage weight: 10,000) in 60-120 sec
4. **Output**: Excel workbook with daily/weekly schedules, employee summaries, role distribution

**Key Constraints**: Continuous 2-4 hour shifts, no split shifts, 19hr/week max, role qualifications, availability windows, single front desk coverage

## Quick Start

```bash
# Setup
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Run
python main.py employees.csv cpd-requirements.csv --output schedule.xlsx

# Test
pytest tests/ -v
```

## CLI flags

- `--favor <EMPLOYEE>`: Soft preference to hit an employee's target hours. Repeatable.
- `--training <DEPT,PERSON1,PERSON2>`: Soft requirement for two people to co-work in a department. Repeatable. Quote the value in shells that glob brackets (e.g., zsh): `--training"[marketing,Alice,Bob]"` or `--training 'marketing,Alice,Bob'`.
- `--favor-dept <DEPT[:MULT]>`: Softly favor a department's focused hours and target adherence. Optional multiplier (default 1.0) to strengthen the bias. Repeatable.
- `--favor-frontdesk-dept <DEPT[:MULT]>`: Softly favor members of a department for front desk duty. Optional multiplier (default 1.0). Repeatable.

## Input Format

**employees.csv**: `name`, `roles` (semicolon/comma-separated), `target_hours`, `max_hours`, `year`, + 90 availability columns (`Mon_08:00` through `Fri_16:30`, 1=available, 0=unavailable)

**cpd-requirements.csv**: `department`, `target_hours`, `max_hours`

## Key Features

- **Guaranteed Coverage**: Front desk staffed 100% of operating hours
- **Smart Constraints**: Continuous shifts (2-4 hours), no split shifts, respects availability
- **Multi-Objective**: Balances 13 priorities (coverage, targets, collaboration, preferences)
- **Professional Output**: Excel with daily grids, employee summaries, departmental analysis
- **Fast**: Solves complex scheduling problem in ~2 minutes

## Testing

Test suite covers data validation, constraint logic, and edge cases:
```bash
pytest tests/ -v                    # Run all tests
pytest tests/test_data_loading.py  # Data parsing tests
pytest tests/test_constraints.py   # Constraint validation tests
```

## License

MIT License - see [LICENSE](LICENSE) file.

## Technical Stack

- **OR-Tools**: Google's constraint programming solver (CP-SAT)
- **Python 3.12+**: Core language
- **Pandas**: Data manipulation and Excel export
- **Pytest**: Testing framework
