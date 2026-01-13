"""Command-line interface for the semester scheduler."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from scheduler.config import DAY_NAMES, DEFAULT_SOLVER_MAX_TIME, TIME_SLOT_STARTS
from scheduler.domain.models import TimesetRequest, TrainingRequest
from scheduler.engine.solver import solve_schedule


def build_parser() -> argparse.ArgumentParser:
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
    parser.add_argument(
        "--max-solve-seconds",
        type=int,
        default=None,
        help="Optional override for the solver time limit in seconds.",
    )
    parser.add_argument(
        "--favor",
        "-f",
        action="append",
        default=[],
        metavar="EMPLOYEE",
        help=(
            "Employee name to prioritize hitting target hours. "
            "Use multiple times to favor more than one person."
        ),
    )
    parser.add_argument(
        "--training",
        action="append",
        default=[],
        metavar="DEPT,PERSON1,PERSON2",
        help=(
            "Specify a training trio: department,trainee1,trainee2 (brackets optional). "
            "Use multiple times for multiple training pairs."
        ),
    )
    parser.add_argument(
        "--favor-dept",
        action="append",
        default=[],
        metavar="DEPT[:MULT]",
        help=(
            "Softly favor a department: DEPT or DEPT:multiplier to boost focused hours and target adherence. "
            "Repeatable."
        ),
    )
    parser.add_argument(
        "--favor-frontdesk-dept",
        action="append",
        default=[],
        metavar="DEPT[:MULT]",
        help=(
            "Softly favor a department's members for front desk duty. Optional multiplier (default 1.0). "
            "Repeatable."
        ),
    )
    parser.add_argument(
        "--progress",
        action="store_true",
        help="Show a simple progress timer toward the max solve time.",
    )
    parser.add_argument(
        "--timeset",
        action="append",
        nargs=5,
        metavar=("NAME", "DAY", "DEPT", "START", "END"),
        default=[],
        help=(
            "Strongly enforce assigning NAME to DEPT on DAY from START (inclusive) to END (exclusive) "
            "in 30-minute increments. Repeatable."
        ),
    )
    return parser


def _parse_training_args(raw_training: list[str]) -> list[TrainingRequest]:
    """Parse raw --training arguments into structured requests."""
    requests: list[TrainingRequest] = []
    for raw in raw_training:
        value = raw.strip()
        if value.startswith("[") and value.endswith("]"):
            value = value[1:-1]
        parts = [part.strip() for part in value.split(",") if part.strip()]
        if len(parts) != 3:
            raise ValueError(
                f"Invalid --training value '{raw}'. Expected format: DEPT,PERSON1,PERSON2"
            )
        dept, person_one, person_two = parts
        if person_one.lower() == person_two.lower():
            raise ValueError(f"Invalid --training value '{raw}': trainees must be different people.")
        requests.append(
            TrainingRequest(department=dept, trainee_one=person_one, trainee_two=person_two)
        )
    return requests


def _parse_favored_departments(raw: list[str]) -> dict[str, float]:
    favored: dict[str, float] = {}
    for entry in raw:
        value = entry.strip()
        if not value:
            continue
        if ":" in value:
            dept, mult = value.split(":", 1)
            dept = dept.strip()
            try:
                multiplier = float(mult.strip())
            except ValueError:
                raise ValueError(f"Invalid --favor-dept multiplier in '{entry}'. Expected a number after ':'.") from None
        else:
            dept = value
            multiplier = None
        if not dept:
            raise ValueError(f"Invalid --favor-dept value '{entry}'. Department name is required.")
        favored[dept] = multiplier if multiplier is not None else 1.0
    return favored


def _parse_favored_fd_departments(raw: list[str]) -> dict[str, float]:
    favored: dict[str, float] = {}
    for entry in raw:
        value = entry.strip()
        if not value:
            continue
        if ":" in value:
            dept, mult = value.split(":", 1)
            dept = dept.strip()
            try:
                multiplier = float(mult.strip())
            except ValueError:
                raise ValueError(f"Invalid --favor-frontdesk-dept multiplier in '{entry}'. Expected a number after ':'.") from None
        else:
            dept = value
            multiplier = None
        if not dept:
            raise ValueError(f"Invalid --favor-frontdesk-dept value '{entry}'. Department name is required.")
        favored[dept] = multiplier if multiplier is not None else 1.0
    return favored


def _parse_timesets(raw_timesets: list[list[str]]) -> list[TimesetRequest]:
    """Parse --timeset entries into structured requests."""
    time_to_slot = {t: idx for idx, t in enumerate(TIME_SLOT_STARTS)}
    day_lookup = {d.lower(): d for d in DAY_NAMES}
    last_start_minutes = int(TIME_SLOT_STARTS[-1].split(":")[0]) * 60 + int(TIME_SLOT_STARTS[-1].split(":")[1])
    final_edge_minutes = last_start_minutes + 30
    final_edge_label = f"{final_edge_minutes // 60:02d}:{final_edge_minutes % 60:02d}"

    def _normalize_day(day: str) -> str:
        key = day.strip().lower()
        for name in DAY_NAMES:
            lower = name.lower()
            if key == lower or key == lower[:3] or key.startswith(lower):
                return name
        if key in day_lookup:
            return day_lookup[key]
        raise ValueError(f"Invalid day '{day}' for --timeset. Use one of: {', '.join(DAY_NAMES)}.")

    def _normalize_time(value: str, *, is_end: bool = False) -> int:
        text = value.strip()
        if len(text) == 4 and text[1] == ":":
            text = f"0{text}"
        if is_end and text == final_edge_label:
            return len(TIME_SLOT_STARTS)
        if text not in time_to_slot:
            raise ValueError(
                f"Invalid time '{value}' for --timeset. Expected HH:MM on 30-minute increments "
                f"from {TIME_SLOT_STARTS[0]} to {TIME_SLOT_STARTS[-1]}."
            )
        return time_to_slot[text]

    requests: list[TimesetRequest] = []
    for entry in raw_timesets:
        if len(entry) != 5:
            raise ValueError(
                f"Invalid --timeset entry '{entry}'. Expected: NAME DAY DEPT START END (30-minute aligned)."
            )
        name, day, dept, start, end = entry
        start_slot = _normalize_time(start)
        end_slot = _normalize_time(end, is_end=True)
        if end_slot <= start_slot:
            raise ValueError(f"--timeset end time must be after start time (got {start} to {end}).")
        requests.append(
            TimesetRequest(
                employee=name.strip(),
                day=_normalize_day(day),
                department=dept.strip(),
                start_slot=start_slot,
                end_slot=end_slot,
            )
        )
    return requests


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        time_limit = args.max_solve_seconds if args.max_solve_seconds is not None else DEFAULT_SOLVER_MAX_TIME
        training_requests = _parse_training_args(args.training)
        favored_departments = _parse_favored_departments(args.favor_dept)
        favored_frontdesk_departments = _parse_favored_fd_departments(args.favor_frontdesk_dept)
        timeset_requests = _parse_timesets(args.timeset)
        output_path = args.output
        if not str(output_path).lower().endswith(".xlsx"):
            output_path = output_path.with_name(output_path.name + ".xlsx")
        solve_schedule(
            staff_csv=args.staff_csv,
            requirements_csv=args.requirements_csv,
            output_path=output_path,
            solver_max_time=time_limit,
            favored_employees=args.favor,
            training_requests=training_requests,
            favored_departments=favored_departments,
            favored_frontdesk_departments=favored_frontdesk_departments,
            timeset_requests=timeset_requests,
            show_progress=args.progress,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
