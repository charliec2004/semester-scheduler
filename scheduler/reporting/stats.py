"""Shared helpers for summarizing solved schedules."""

from __future__ import annotations

from typing import Dict, Iterable, List, Tuple

from scheduler.config import FRONT_DESK_ROLE


def aggregate_department_hours(
    solver,
    employees: Iterable[str],
    days: Iterable[str],
    time_slots: Iterable[int],
    assign,
    department_roles: List[str],
    qual: Dict[str, set],
    primary_frontdesk_department: Dict[str, str] | None = None,
) -> Tuple[Dict[str, int], Dict[str, int], Dict[str, Dict[str, float]]]:
    """
    Aggregate focused/dual hours for each department based on a solved schedule.

    Returns:
        role_direct_slots: Direct slot count per role (includes front desk).
        front_desk_slots_by_employee: Slot count per employee when working front desk.
        department_breakdown: Per-department hour metrics (focused, dual, actual).
    """
    role_direct_slots: Dict[str, int] = {FRONT_DESK_ROLE: 0, **{role: 0 for role in department_roles}}
    front_desk_slots_by_employee: Dict[str, int] = {e: 0 for e in employees}

    for e in employees:
        for d in days:
            for t in time_slots:
                fd_key = (e, d, t, FRONT_DESK_ROLE)
                if fd_key in assign and solver.value(assign[fd_key]):
                    role_direct_slots[FRONT_DESK_ROLE] += 1
                    front_desk_slots_by_employee[e] += 1

                for role in department_roles:
                    key = (e, d, t, role)
                    if key in assign and solver.value(assign[key]):
                        role_direct_slots[role] += 1

    dual_slots_by_role: Dict[str, int] = {role: 0 for role in department_roles}
    for e, fd_slots in front_desk_slots_by_employee.items():
        primary = (primary_frontdesk_department or {}).get(e)
        if primary and primary in department_roles:
            dual_slots_by_role[primary] += fd_slots
        else:
            # Fallback: credit to all qualified departments if no primary provided
            for role in department_roles:
                if role in qual[e]:
                    dual_slots_by_role[role] += fd_slots

    department_breakdown: Dict[str, Dict[str, float]] = {}
    for role in department_roles:
        focused_slots = role_direct_slots[role]
        dual_slots = dual_slots_by_role[role]
        focused_hours = focused_slots * 0.5
        dual_hours_total = dual_slots * 0.5
        dual_hours_counted = dual_hours_total * 0.5
        department_breakdown[role] = {
            "focused_slots": focused_slots,
            "focused_hours": focused_hours,
            "dual_slots": dual_slots,
            "dual_hours_total": dual_hours_total,
            "dual_hours_counted": dual_hours_counted,
            "actual_hours": focused_hours + dual_hours_counted,
        }

    return role_direct_slots, front_desk_slots_by_employee, department_breakdown
