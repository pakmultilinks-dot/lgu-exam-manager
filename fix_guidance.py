"""Exact fix guidance for export-blocking workbook errors.

Every line here is derived only from the imported workbook itself. Nothing
is invented, nothing is changed in the source file, and no validation or
calculation behaviour is touched. Guidance is plain words: where the error
is, and the exact cell change that fixes it.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation

from academic_engine import PLACEHOLDERS, _key, _value, derive_term
from academic_identity import program_key, registration_key
from validation_scope import issue_source
from workbook_schema import ROLE_LABELS

_COMPACT_TERM = re.compile(r"(SU|S|F)[\s_/-]*(\d{2}|\d{4})", re.I)
_CORE_ROLES = ("registration", "name", "course_code", "course_name", "credits", "grade")
_GUIDED_CODES = ("STUDENT_NAME_CONFLICT", "DEGREE_CONFLICT", "INVALID_CREDITS", "MISSING_REQUIRED")
_MAX_ROWS_LISTED = 10


def _missing(value: str) -> bool:
    return (value or "").strip().upper() in PLACEHOLDERS


def _row_list(rows) -> str:
    """Excel row numbers as plain words, capped so a long list stays readable."""
    ordered = sorted(set(rows))
    if len(ordered) == 1:
        return str(ordered[0])
    if len(ordered) <= _MAX_ROWS_LISTED:
        return ", ".join(str(row) for row in ordered[:-1]) + " and " + str(ordered[-1])
    shown = ", ".join(str(row) for row in ordered[:_MAX_ROWS_LISTED])
    remaining = len(ordered) - _MAX_ROWS_LISTED
    return f"{shown} and {remaining} more {'row' if remaining == 1 else 'rows'}"


class _Evidence:
    """File-derived evidence for fix guidance, built once per dataset."""

    def __init__(self, dataset):
        self.dataset = dataset
        self.layouts = {layout.sheet: layout for layout in dataset.layouts}
        self.names_by_registration = defaultdict(lambda: defaultdict(list))
        self.degrees_by_registration = defaultdict(lambda: defaultdict(list))
        self.credits_by_course = defaultdict(Counter)
        self.course_names = defaultdict(Counter)
        self.program_term_courses = defaultdict(dict)
        self.student_term_courses = defaultdict(set)
        self.registrations_by_name = defaultdict(set)
        self._build()

    def _layout_for(self, row):
        return self.layouts.get(row.get("sheet", ""))

    def _cell(self, row, role):
        layout = self._layout_for(row)
        if layout is None:
            return ""
        column = layout.columns.get(role)
        if column is None:
            return str(layout.defaults.get(role, ""))
        return _value(row["cells"].get(column))

    def _row_term(self, row):
        raw = self._cell(row, "term")
        match = _COMPACT_TERM.fullmatch(raw.strip()) if raw else None
        if match:
            return match[1].upper() + match[2][-2:]
        return derive_term(self._cell(row, "semester"))

    def _build(self):
        for row in self.dataset.source_rows:
            reg = registration_key(self._cell(row, "registration"))
            name = self._cell(row, "name")
            degree = self._cell(row, "degree")
            code = self._cell(row, "course_code")
            credits = self._cell(row, "credits")
            term = self._row_term(row)
            excel_row = row.get("excel_row", 0)
            if reg and name and not _missing(name):
                self.names_by_registration[reg][_key(name)].append((name, excel_row))
                self.registrations_by_name[_key(name)].add(reg)
            if reg and degree and not _missing(degree):
                self.degrees_by_registration[reg][_key(degree)].append((degree, excel_row))
            if code and not _missing(code):
                self.credits_by_course[code.upper()][credits] += 1
                course_name = self._cell(row, "course_name")
                if course_name and not _missing(course_name):
                    self.course_names[code.upper()][course_name] += 1
                if reg and term != "Unspecified term":
                    self.student_term_courses[(reg, term)].add(code.upper())
                program = program_key(degree) if degree and not _missing(degree) else ""
                if program and term != "Unspecified term":
                    self.program_term_courses[(program, term)].setdefault(
                        code.upper(), course_name or "")


def _evidence(dataset):
    cached = getattr(dataset, "_fix_evidence", None)
    if cached is None:
        cached = _Evidence(dataset)
        dataset._fix_evidence = cached
    return cached


def _variant_groups(groups):
    """Rank name/degree variant groups by row count, most common first."""
    ranked = []
    for key, entries in groups.items():
        displays = Counter(display for display, _ in entries)
        ranked.append({
            "key": key,
            "display": displays.most_common(1)[0][0],
            "variants": ", ".join(
                f"{display} ({count} {'row' if count == 1 else 'rows'})"
                for display, count in displays.most_common()),
            "rows": [excel_row for _, excel_row in entries],
        })
    ranked.sort(key=lambda group: (-len(group["rows"]), group["display"]))
    return ranked


def _conflict_fix(evidence, dataset, issue, *, role, label, column_label):
    source, fields = issue_source(dataset, issue)
    reg = fields.get("registration", "")
    if _missing(reg):
        return None
    groups = evidence.names_by_registration.get(registration_key(reg)) if role == "name" \
        else evidence.degrees_by_registration.get(registration_key(reg))
    if not groups or len(groups) < 2:
        return None
    ranked = _variant_groups(groups)
    majority = ranked[0]
    summary = ", ".join(group["variants"] for group in ranked)
    lines = [f"Registration {reg} appears with these {label}: {summary}."]
    if len(majority["rows"]) == len(ranked[1]["rows"]):
        lines.append(
            f"No one spelling is used most often, so check the student's official record, "
            f"pick the correct spelling, and use that spelling in the {column_label} cell "
            f"of every row for this registration.")
    else:
        for group in ranked[1:]:
            lines.append(
                f"Change the {column_label} cell to {majority['display']} "
                f"in Excel row {_row_list(group['rows'])}.")
            if role == "name":
                others = sorted(evidence.registrations_by_name.get(group["key"], set())
                                - {registration_key(reg)})
                if others:
                    lines.append(
                        f"Note: {group['display']} is also the name on registration "
                        f"{', '.join(others[:3])}. Check that the row does not belong to "
                        f"that student before changing it.")
    lines.append("After the change, save the workbook and reimport it.")
    return " ".join(lines)


def _invalid_credits_fix(evidence, dataset, issue):
    source, fields = issue_source(dataset, issue)
    code = fields.get("course_code", "")
    if _missing(code):
        return None
    distribution = evidence.credits_by_course.get(code.upper())
    if not distribution:
        return None
    course_name = fields.get("course_name", "")
    if _missing(course_name):
        names = evidence.course_names.get(code.upper())
        course_name = names.most_common(1)[0][0] if names else ""
    title = f"course {code}" + (f" ({course_name})" if course_name and not _missing(course_name) else "")
    total = sum(distribution.values())
    valid = defaultdict(Counter)
    zero_rows = 0
    blank_rows = 0
    other_rows = 0
    for raw, count in distribution.items():
        if _missing(raw):
            blank_rows += count
            continue
        try:
            value = Decimal(str(raw).strip())
        except (InvalidOperation, ValueError):
            other_rows += count
            continue
        if not value.is_finite():
            other_rows += count
        elif value == 0:
            # Zero credit hours is not a usable course load, so it is never
            # recommended; it is only reported as evidence.
            zero_rows += count
        elif Decimal(0) < value <= Decimal(50):
            valid[value][raw] += count
        else:
            other_rows += count
    if not valid:
        if zero_rows == 0 and other_rows == 0:
            return (
                f"No credit hours for {title} appear anywhere in this file. Every row for "
                f"this course is blank or '-'. Enter the credit hours from the program's "
                f"course catalog in the CRH cell of Excel row {issue.excel_row}.")
        detail = "; ".join(
            part for part in (
                f"{zero_rows} rows say 0" if zero_rows else "",
                f"{blank_rows} rows are blank or '-'" if blank_rows else "",
                f"{other_rows} rows have a value that is not a number" if other_rows else "",
            ) if part)
        return (
            f"No usable credit hours for {title} appear anywhere in this file ({detail}). "
            f"Enter the credit hours from the program's course catalog in the CRH "
            f"cell of Excel row {issue.excel_row}.")
    totals = sorted(((sum(counter.values()), value, counter) for value, counter in valid.items()),
                    key=lambda item: (-item[0], str(item[1])))
    top_count, top_value, top_spellings = totals[0]
    display = top_spellings.most_common(1)[0][0]
    dominant = len(totals) == 1 or (
        top_count / total >= 0.8 and top_count >= 2 * totals[1][0])
    if dominant:
        return (
            f"Course {code}" + (f" ({course_name})" if course_name and not _missing(course_name) else "")
            + f" has {display} credit hours in {top_count} of {total} "
            f"rows in this file. Set the CRH cell to {display} in Excel row {issue.excel_row}.")
    second_count, second_value, second_spellings = totals[1]
    return (
        f"Rows for {title} disagree about credit hours in this file "
        f"({top_count} rows say {top_spellings.most_common(1)[0][0]}, "
        f"{second_count} rows say {second_spellings.most_common(1)[0][0]}). Check the "
        f"program's course catalog and set the CRH cell in Excel row {issue.excel_row} "
        f"to the correct value.")


def _header_label(dataset, issue, role):
    layout = next((item for item in dataset.layouts if item.sheet == issue.sheet), None)
    if layout is not None and role in layout.columns:
        return layout.columns[role]
    return ROLE_LABELS.get(role, role)


def _missing_required_fix(evidence, dataset, issue):
    source, fields = issue_source(dataset, issue)
    row = issue.excel_row
    missing = [role for role in ("registration", "name", "course_code", "course_name")
               if _missing(fields.get(role, ""))]
    present = [role for role in _CORE_ROLES if not _missing(fields.get(role, ""))]
    student = fields.get("name", "")
    reg = fields.get("registration", "")
    semester = fields.get("semester", "")
    term = fields.get("term", "")
    if not term or _missing(term) or term == "Unspecified term":
        term = derive_term(semester) if semester and not _missing(semester) else ""
    if not missing:
        return None
    if all(role in missing for role in ("registration", "name", "course_code", "course_name")):
        extra = ""
        if "grade" in present:
            extra = f", only a grade {fields.get('grade', '')}"
        elif "credits" in present:
            extra = f", only {fields.get('credits', '')} credit hours"
        return (
            f"Excel row {row} has no registration, student name or course{extra}. It cannot "
            f"belong to any student, so delete the row from the source workbook.")
    if ({"course_code", "course_name"} <= set(missing)
            and "grade" not in present and "credits" not in present
            and "registration" in present and "name" in present):
        return (
            f"Excel row {row} has a student and a semester but no course, grade or credit "
            f"hours, so it carries no result. Delete the row from the source workbook.")
    if {"course_code", "course_name"} <= set(missing) and "registration" in present:
        program = program_key(fields.get("degree", "")) if not _missing(fields.get("degree", "")) else ""
        if program and term and term != "Unspecified term":
            offered = evidence.program_term_courses.get((program, term), {})
            taken = evidence.student_term_courses.get((registration_key(reg), term), set())
            candidates = sorted(set(offered) - taken)
            if len(candidates) == 1:
                code = candidates[0]
                course_name = offered[code] or ""
                when = semester if semester and not _missing(semester) else term
                who = student if student and not _missing(student) else reg
                return (
                    f"The Course Code and Course Name cells are empty in Excel row {row}. "
                    f"The only course in {fields.get('degree', '')} in {when} that {who} "
                    f"has no row for is {code}"
                    + (f" ({course_name})" if course_name else "")
                    + f". Check the source records and, if that is the missing course, "
                    f"fill in its code and name in this row.")
        when = semester if semester and not _missing(semester) else term
        who = student if student and not _missing(student) else reg
        return (
            f"The Course Code and Course Name cells are empty in Excel row {row}. Fill "
            f"in the course code and course name from the source records for {who}"
            + (f" in {when}" if when else "")
            + ".")
    labels = [_header_label(dataset, issue, role) for role in missing]
    if len(labels) == 1:
        return (
            f"The {labels[0]} cell is empty in Excel row {row}. Fill it in from the "
            f"source records.")
    joined = ", ".join(labels[:-1]) + " and " + labels[-1]
    return (
        f"The {joined} cells are empty in Excel row {row}. Fill them in from the "
        f"source records.")


def fix_guidance(dataset, issue):
    """Plain-worded fix for one export-blocking error, or None.

    Only ERROR issues with a concrete, file-derived fix get guidance.
    Warnings and information notices are left as-is, and no number or name
    is ever invented.
    """
    if issue is None or issue.level != "ERROR":
        return None
    if not getattr(dataset, "source_rows", None) or not getattr(dataset, "layouts", None):
        return None
    if issue.code not in _GUIDED_CODES:
        return None
    evidence = _evidence(dataset)
    if issue.code == "STUDENT_NAME_CONFLICT":
        return _conflict_fix(evidence, dataset, issue, role="name",
                             label="names", column_label="Student Name")
    if issue.code == "DEGREE_CONFLICT":
        return _conflict_fix(evidence, dataset, issue, role="degree",
                             label="programs", column_label="Degree Title")
    if issue.code == "INVALID_CREDITS":
        return _invalid_credits_fix(evidence, dataset, issue)
    if issue.code == "MISSING_REQUIRED":
        return _missing_required_fix(evidence, dataset, issue)
    return None


def fix_lines_for_errors(dataset, errors):
    """One guidance line per error that has a concrete fix, in issue order."""
    lines = []
    for issue in errors or []:
        fix = fix_guidance(dataset, issue)
        if fix:
            lines.append(fix)
    return lines
