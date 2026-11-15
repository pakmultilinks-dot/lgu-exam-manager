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
_LETTER_GRADES = {"A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D"}


def _grade_group(grade):
    """Letter grades (A to D) carry grade points; PASS, F and the rest do not."""
    return "letter" if (grade or "").strip().upper() in _LETTER_GRADES else "other"


def _plural(count, one, many):
    """Pick the singular or plural wording for a count."""
    return one if count == 1 else many


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
        self.credits_by_course_program = defaultdict(Counter)
        self.credits_by_course_program_grade = defaultdict(Counter)
        self.programs_by_course = defaultdict(set)
        self.program_display_names = defaultdict(Counter)
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

    def program_display(self, key):
        """The program name as written in the file, not the normalized key."""
        names = self.program_display_names.get(key)
        return names.most_common(1)[0][0] if names else key

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
                program = program_key(degree) if degree and not _missing(degree) else ""
                if program:
                    self.programs_by_course[code.upper()].add(program)
                    self.program_display_names[program][degree.strip()] += 1
                    self.credits_by_course_program[(code.upper(), program)][credits] += 1
                    self.credits_by_course_program_grade[
                        (code.upper(), program, _grade_group(self._cell(row, "grade")))][credits] += 1
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


def _split_credits(distribution):
    """Split a credit-hour distribution into usable values, zeros, blanks, other."""
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
            zero_rows += count
        elif Decimal(0) < value <= Decimal(50):
            valid[value][raw] += count
        else:
            other_rows += count
    return valid, zero_rows, blank_rows, other_rows


def _dominant(valid, total_rows):
    """The dominant positive value in a split distribution, or None.

    One value counts as dominant when it is the only positive value, or
    when it covers at least 80 percent of the rows in scope and at least
    twice the next value.
    """
    if not valid:
        return None
    totals = sorted(((sum(counter.values()), value, counter) for value, counter in valid.items()),
                    key=lambda item: (-item[0], str(item[1])))
    top_count, top_value, top_spellings = totals[0]
    if len(totals) == 1 or (top_count / total_rows >= 0.8 and top_count >= 2 * totals[1][0]):
        return top_count, top_value, top_spellings
    return None


def _invalid_credits_fix(evidence, dataset, issue):
    source, fields = issue_source(dataset, issue)
    code = fields.get("course_code", "")
    if _missing(code):
        return None
    course_name = fields.get("course_name", "")
    if _missing(course_name):
        names = evidence.course_names.get(code.upper())
        course_name = names.most_common(1)[0][0] if names else ""
    title = f"course {code}" + (f" ({course_name})" if course_name and not _missing(course_name) else "")
    program = program_key(fields.get("degree", "")) if not _missing(fields.get("degree", "")) else ""
    display_program = evidence.program_display(program) if program else ""
    group = _grade_group(fields.get("grade", ""))
    # Narrowest scope first: same course, program and grade group, then the
    # program, then the whole file. A course can be worth different credits
    # in different programs, so the whole-file value is only a fallback.
    scopes = []
    if program:
        scopes.append((f"for {display_program}", evidence.credits_by_course_program_grade.get((code.upper(), program, group), {})))
        scopes.append((f"for {display_program}", evidence.credits_by_course_program.get((code.upper(), program), {})))
    scopes.append(("", evidence.credits_by_course.get(code.upper(), {})))
    for scope_label, distribution in scopes:
        if not distribution:
            continue
        valid, zero_rows, blank_rows, other_rows = _split_credits(distribution)
        total = sum(distribution.values())
        dom = _dominant(valid, total)
        if dom is not None:
            top_count, _, top_spellings = dom
            display = top_spellings.most_common(1)[0][0]
            return (
                f"Course {code}" + (f" ({course_name})" if course_name and not _missing(course_name) else "")
                + f" has {display} credit hours in {top_count} of {total} rows in this file"
                + (f" {scope_label}" if scope_label else "")
                + f". Set the CRH cell to {display} in Excel row {issue.excel_row}.")
        if not valid and zero_rows > 0:
            if scope_label:
                # Zero is the only recorded value in this scope, so it is
                # the recorded value for the group, not a missing one.
                detail = "; ".join(
                    part for part in (
                        f"{zero_rows} {_plural(zero_rows, 'row says', 'rows say')} 0" if zero_rows else "",
                        f"{blank_rows} {_plural(blank_rows, 'row is', 'rows are')} blank or '-'" if blank_rows else "",
                    ) if part)
                contrast = ""
                others = sorted(
                    other for other in evidence.programs_by_course.get(code.upper(), set())
                    if other != program)
                other_values = []
                for other in others[:2]:
                    other_dist = evidence.credits_by_course_program.get((code.upper(), other), {})
                    other_valid, _, _, _ = _split_credits(other_dist)
                    other_dom = _dominant(other_valid, sum(other_dist.values()))
                    if other_dom is not None:
                        other_values.append(f"{evidence.program_display(other)} rows say {other_dom[2].most_common(1)[0][0]}")
                if other_values:
                    contrast = " The " + "; ".join(other_values) + "."
                return (
                    f"Course {code}" + (f" ({course_name})" if course_name and not _missing(course_name) else "")
                    + f" is recorded as 0 credit hours {scope_label} in this file ({detail}).{contrast} "
                    f"Set the CRH cell to 0 in Excel row {issue.excel_row}. If the course catalog "
                    f"says otherwise, use the catalog value.")
            detail = "; ".join(
                part for part in (
                    f"{zero_rows} {_plural(zero_rows, 'row says', 'rows say')} 0" if zero_rows else "",
                    f"{blank_rows} {_plural(blank_rows, 'row is', 'rows are')} blank or '-'" if blank_rows else "",
                    f"{other_rows} {_plural(other_rows, 'row has', 'rows have')} a value that is not a number" if other_rows else "",
                ) if part)
            return (
                f"No usable credit hours for {title} appear anywhere in this file ({detail}). "
                f"Enter the credit hours from the program's course catalog in the CRH "
                f"cell of Excel row {issue.excel_row}.")
        if valid:
            # Values disagree and none dominates; the answer is the catalog,
            # not a broader scope that would hide the disagreement.
            totals = sorted(((sum(counter.values()), value, counter) for value, counter in valid.items()),
                            key=lambda item: (-item[0], str(item[1])))
            top_count, _, top_spellings = totals[0]
            second_count, _, second_spellings = totals[1]
            return (
                f"Rows for {title} disagree about credit hours in this file "
                f"({top_count} rows say {top_spellings.most_common(1)[0][0]}, "
                f"{second_count} rows say {second_spellings.most_common(1)[0][0]}). Check the "
                f"program's course catalog and set the CRH cell in Excel row {issue.excel_row} "
                f"to the correct value.")
        # This scope has no evidence at all; try the next, broader scope.
    overall = evidence.credits_by_course.get(code.upper(), {})
    if not overall:
        return None
    valid, zero_rows, blank_rows, other_rows = _split_credits(overall)
    if zero_rows == 0 and other_rows == 0:
        return (
            f"No credit hours for {title} appear anywhere in this file. Every row for "
            f"this course is blank or '-'. Enter the credit hours from the program's "
            f"course catalog in the CRH cell of Excel row {issue.excel_row}.")
    detail = "; ".join(
        part for part in (
            f"{zero_rows} {_plural(zero_rows, 'row says', 'rows say')} 0" if zero_rows else "",
            f"{blank_rows} {_plural(blank_rows, 'row is', 'rows are')} blank or '-'" if blank_rows else "",
            f"{other_rows} {_plural(other_rows, 'row has', 'rows have')} a value that is not a number" if other_rows else "",
        ) if part)
    return (
        f"No usable credit hours for {title} appear anywhere in this file ({detail}). "
        f"Enter the credit hours from the program's course catalog in the CRH "
        f"cell of Excel row {issue.excel_row}.")


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
