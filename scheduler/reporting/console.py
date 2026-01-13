"""Console output helpers for solved schedules."""

from __future__ import annotations

from ortools.sat.python import cp_model

from scheduler.config import FRONT_DESK_ROLE
from scheduler.reporting.stats import aggregate_department_hours


def print_schedule(
    status,
    solver,
    employees,
    days,
    time_slots,
    slot_names,
    qual,
    work,
    assign,
    weekly_hour_limits,
    target_weekly_hours,
    total_time,
    roles,
    department_roles,
    role_display_names,
    department_hour_targets,
    department_max_hours,
    primary_frontdesk_department,
):
    """
    Display the schedule in a readable format with statistics.
    """

    print("\n" + "=" * 120)
    print(f"SCHEDULE STATUS: {status}")
    print("=" * 120)

    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        print("No solution found!")
        print("\nPossible reasons:")
        print("  - Constraints are too restrictive")
        print("  - Not enough qualified employees")
        print("  - Availability conflicts with coverage requirements")
        return

    print(f"\nSolution found!")
    print(f"\nSolver Statistics:")
    print(f"  - Total execution time: {total_time:.2f} seconds")
    print(f"  - Solver computation time: {solver.wall_time:.2f} seconds")
    print(f"  - Branches explored: {solver.num_branches:,}")
    print(f"  - Conflicts encountered: {solver.num_conflicts:,}")

    for d in days:
        print(f"\n{'─' * 120}")
        print(f"{d.upper()}")
        print(f"{'─' * 120}")

        role_columns = [FRONT_DESK_ROLE] + department_roles
        column_width = 22
        header = f"\n{'Time':<12}" + "".join(f"{role_display_names[role]:<{column_width}}" for role in role_columns)
        print(header)
        print("─" * (12 + column_width * len(role_columns)))

        for t in time_slots:
            time_slot = slot_names[t]

            row = f"{time_slot:<12}"
            for role in role_columns:
                workers = [
                    e for e in employees if (e, d, t, role) in assign and solver.value(assign[(e, d, t, role)])
                ]

                if role == FRONT_DESK_ROLE:
                    cell = ", ".join(workers) if workers else "ERROR: UNCOVERED"
                else:
                    cell = ", ".join(workers) if workers else "-"

                row += f"{cell:<{column_width}}"

            print(row)

    print(f"\n{'=' * 120}")
    print("EMPLOYEE SUMMARY")
    print(f"{'=' * 120}\n")

    print(f"{'Employee':<15}{'Qualifications':<35}{'Hours (Target/Max)':<30}{'Days Worked'}")
    print("─" * 120)

    for e in employees:
        total_slots = 0
        days_worked = []

        for d in days:
            day_slots = sum(solver.value(work[e, d, t]) for t in time_slots)
            if day_slots > 0:
                day_hours = day_slots * 0.5
                days_worked.append(f"{d}({day_hours:.1f}h)")
                total_slots += day_slots

        quals = ", ".join(sorted(qual[e]))
        days_str = ", ".join(days_worked) if days_worked else "None"

        target_hours = target_weekly_hours.get(e, 11)
        weekly_limit = weekly_hour_limits.get(e, 40)
        total_hours = total_slots * 0.5

        hours_str = f"{total_hours:.1f} (↑{target_hours}/max {weekly_limit})"
        if abs(total_hours - target_hours) <= 0.5:
            hours_str = f"✓ {hours_str}"

        print(f"{e:<15}{quals:<35}{hours_str:<30}{days_str}")

    print(f"\n{'=' * 120}")
    print("ROLE DISTRIBUTION")
    print(f"{'=' * 120}\n")

    role_totals = {role: 0 for role in roles}

    for d in days:
        role_counts = {role: 0 for role in roles}

        for t in time_slots:
            for e in employees:
                for role in roles:
                    if (e, d, t, role) in assign and solver.value(assign[(e, d, t, role)]):
                        role_counts[role] += 1

        for role in roles:
            role_totals[role] += role_counts[role]

        day_summary = (
            ", ".join(
                f"{role_display_names[role]} {role_counts[role] * 0.5:.1f}h" for role in roles if role_counts[role] > 0
            )
            or "No assignments"
        )
        print(f"{d}: {day_summary}")

    print("\nTOTAL HOURS BY ROLE")
    print("─" * 140)
    print(
        f"{'Role':<25}"
        f"{'Actual':<12}"
        f"{'Target':<12}"
        f"{'Max':<12}"
        f"{'Delta':<12}"
        f"{'Dual Hours':<14}"
        f"{'Dual Counted':<15}"
        f"{'Focused':<12}"
        f"{'Status'}"
    )
    print("─" * 140)

    role_direct_slots, _, department_breakdown = aggregate_department_hours(
        solver, employees, days, time_slots, assign, department_roles, qual, primary_frontdesk_department
    )

    for role in roles:
        role_name = role_display_names[role]

        if role == FRONT_DESK_ROLE:
            actual_hours = role_direct_slots[role] * 0.5
            target = department_hour_targets.get(role)
            max_hours = department_max_hours.get(role)
            dual_hours = "-"
            dual_counted = "-"
            focused_hours = f"{actual_hours:.1f}h"
        else:
            stats = department_breakdown[role]
            actual_hours = stats["actual_hours"]
            target = department_hour_targets.get(role)
            max_hours = department_max_hours.get(role)
            dual_hours = f"{stats['dual_hours_total']:.1f}h"
            dual_counted = f"{stats['dual_hours_counted']:.1f}h"
            focused_hours = f"{stats['focused_hours']:.1f}h"

        actual_str = f"{actual_hours:.1f}h"
        target_str = f"{target:.1f}h" if target is not None else "-"
        max_str = f"{max_hours:.1f}h" if max_hours is not None else "-"

        if target is not None:
            delta = actual_hours - target
            delta_str = f"{delta:+.1f}h"

            if abs(delta) <= 1.0:
                status = "✓ On Target"
            elif delta > 0:
                status = "↑ Over"
            else:
                status = "↓ Under"
        else:
            delta_str = "-"
            status = "-"

        print(
            f"{role_name:<25}"
            f"{actual_str:<12}"
            f"{target_str:<12}"
            f"{max_str:<12}"
            f"{delta_str:<12}"
            f"{dual_hours:<14}"
            f"{dual_counted:<15}"
            f"{focused_hours:<12}"
            f"{status}"
        )
