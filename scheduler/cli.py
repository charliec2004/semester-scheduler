"""Command-line interface for the semester scheduler."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from scheduler.config import DEFAULT_SOLVER_MAX_TIME
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
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        time_limit = args.max_solve_seconds if args.max_solve_seconds is not None else DEFAULT_SOLVER_MAX_TIME
        solve_schedule(
            staff_csv=args.staff_csv,
            requirements_csv=args.requirements_csv,
            output_path=args.output,
            solver_max_time=time_limit,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
