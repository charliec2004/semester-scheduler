"""Excel export helpers for solved schedules."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
from ortools.sat.python import cp_model

from scheduler.config import FRONT_DESK_ROLE
from scheduler.reporting.stats import aggregate_department_hours


def export_schedule_to_excel(
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
    roles,
    department_roles,
    role_display_names,
    department_hour_targets,
    department_max_hours,
    output_path: Path,
):
    """Export the generated schedule to an Excel workbook with formatted sheets."""
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        return

    role_columns = [FRONT_DESK_ROLE] + department_roles

    daily_tables = []
    weekly_rows = []
    role_headers = [role_display_names[role] for role in role_columns]
    weekly_columns = ["Day", "Time"] + role_headers

    for day in days:
        day_rows = []
        for t in time_slots:
            cell_values = []
            for role in role_columns:
                workers = [
                    e for e in employees if (e, day, t, role) in assign and solver.value(assign[(e, day, t, role)])
                ]
                cell_values.append(", ".join(workers) if workers else ("UNCOVERED" if role == FRONT_DESK_ROLE else ""))
            day_rows.append([slot_names[t], *cell_values])
            weekly_rows.append([day, slot_names[t], *cell_values])
        daily_tables.append((f"{day} Schedule", ["Time"] + role_headers, day_rows))

    summary_rows = []
    for e in employees:
        total_slots = 0
        days_worked = []
        for d in days:
            day_slots = sum(solver.value(work[e, d, t]) for t in time_slots)
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

    distribution_rows = []
    role_totals = {role: 0 for role in roles}
    for d in days:
        row = [d]
        for role in roles:
            slot_count = sum(
                solver.value(assign[(e, d, t, role)]) if (e, d, t, role) in assign else 0 for e in employees for t in time_slots
            )
            role_totals[role] += slot_count
            row.append(slot_count * 0.5)
        distribution_rows.append(row)
    total_row = ["TOTAL"] + [role_totals[r] * 0.5 for r in roles]
    distribution_rows.append(total_row)
    distribution_columns = ["Day"] + [role_display_names[role] for role in roles]

    _, _, department_breakdown = aggregate_department_hours(
        solver, employees, days, time_slots, assign, department_roles, qual
    )

    dept_summary_headers = [
        "Department",
        "Actual Hours",
        "Target Hours",
        "Max Hours",
        "Delta (Actual-Target)",
        "Focused Hours",
        "Dual Hours Total",
        "Dual Hours Counted",
    ]
    dept_summary_rows = []
    for role in department_roles:
        stats = department_breakdown[role]
        actual_hours = stats["actual_hours"]
        target = department_hour_targets.get(role)
        max_hours = department_max_hours.get(role)
        delta = actual_hours - target if target is not None else ""
        dept_summary_rows.append(
            [
                role_display_names[role],
                actual_hours,
                target if target is not None else "",
                max_hours if max_hours is not None else "",
                delta,
                stats["focused_hours"],
                stats["dual_hours_total"],
                stats["dual_hours_counted"],
            ]
        )

    engine = None
    for candidate in ("xlsxwriter", "openpyxl"):
        if importlib.util.find_spec(candidate):
            engine = candidate
            break

    if engine is None:
        _write_minimal_xlsx(output_path, weekly_columns, weekly_rows)
        return

    with pd.ExcelWriter(output_path, engine=engine) as writer:
        df_weekly = pd.DataFrame(weekly_rows, columns=weekly_columns)
        df_weekly.to_excel(writer, sheet_name="Weekly Grid", index=False)
        _autosize_columns(writer, "Weekly Grid", df_weekly)

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
        if dept_summary_rows:
            df_dept = pd.DataFrame(dept_summary_rows, columns=dept_summary_headers)
            df_dept.to_excel(writer, sheet_name="Department Summary", index=False)
            _autosize_columns(writer, "Department Summary", df_dept)


def _autosize_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame):
    worksheet = writer.sheets[sheet_name]
    for idx, column in enumerate(dataframe.columns):
        max_len = max([len(str(column))] + [len(str(cell)) for cell in dataframe[column]])
        worksheet.set_column(idx, idx, max_len + 2)


def _write_minimal_xlsx(output_path: Path, weekly_columns: List[str], weekly_rows: List[List]):
    """Fallback XLSX writer if no Excel engines are installed."""
    sheets = [
        (
            "xl/worksheets/sheet1.xml",
            _create_sheet_xml("Weekly Grid", weekly_columns, weekly_rows),
        )
    ]
    _write_minimal_xlsx_archive(output_path, sheets)


def _create_sheet_xml(name: str, columns: List[str], rows: List[List]) -> str:
    header = "".join(f'<c t="inlineStr"><is><t>{col}</t></is></c>' for col in columns)
    body = ""
    for row in rows:
        body += "<row>" + "".join(f'<c t="inlineStr"><is><t>{cell}</t></is></c>' for cell in row) + "</row>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        f"<row>{header}</row>"
        f"{body}"
        "</sheetData>"
        "</worksheet>"
    )


def _write_minimal_xlsx_archive(output_path: Path, sheet_files: List[Tuple[str, str]]):
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(
            f'<Override PartName="/{filename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for filename, _ in sheet_files
        )
        + "</Types>"
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    sheets_entries = "".join(
        f'<sheet name="Sheet{idx + 1}" sheetId="{idx + 1}" r:id="rId{idx + 1}"/>' for idx, _ in enumerate(sheet_files)
    )
    workbook_rels = [
        f'<Relationship Id="rId{idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx + 1}.xml"/>'
        for idx, _ in enumerate(sheet_files)
    ]

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_entries}</sheets>"
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
        '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
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
