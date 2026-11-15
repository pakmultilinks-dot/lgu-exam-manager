"""Fix guidance: every export-blocking error names the row and the exact fix.

Synthetic workbook only. No real student personal data is checked in.
"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
import pytest
from academic_engine import load_course_workbook
from fix_guidance import fix_guidance
from test_academic_engine import _df


def _error_frame():
    base = _df()
    extra = pd.DataFrame([
        # Row 6: same registration, different name (name conflict).
        ["SPRING 2026", "Fa-2025/BSCS/001", "SOMEONE ELSE", 0, "CS100", "Course One", 3, 90, "A", "BSCS", 9.0, "S26"],
        # Row 7: missing credit hours for CS101, which other rows say is 4.
        ["SPRING 2026", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS101", "Course Two", "-", 60, "C", "BSCS", 0, "S26"],
        # Row 8: missing course for STUDENT BETA, who already has CS100 and CS101.
        ["SPRING 2026", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "-", "-", 3, 70, "B", "BSCS", 0, "S26"],
        # Row 9: STUDENT GAMMA has only CS100 in S26.
        ["SPRING 2026", "Fa-2025/BSCS/003", "STUDENT GAMMA", 0, "CS100", "Course One", 3, 65, "B", "BSCS", 9.0, "S26"],
        # Row 10: missing course for GAMMA; CS101 is the only course left.
        ["SPRING 2026", "Fa-2025/BSCS/003", "STUDENT GAMMA", 0, "-", "-", 3, 88, "A", "BSCS", 0, "S26"],
        # Rows 11-12: CS999 has no credit hours anywhere in the file.
        ["SPRING 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS999", "New Course", "-", 80, "A", "BSCS", 0, "S25"],
        ["SPRING 2025", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS999", "New Course", "-", 75, "B", "BSCS", 0, "S25"],
        # Rows 13-14: a clean second program, so the pack still exports.
        ["SPRING 2026", "Fa-2025/BS IT/001", "IT STUDENT", 0, "IT100", "IT Course One", 3, 85, "A", "BS IT", 9.0, "S26"],
        ["SPRING 2026", "Fa-2025/BS IT/001", "IT STUDENT", 0, "IT101", "IT Course Two", 3, 78, "B", "BS IT", 9.0, "S26"],
        # Row 15: student and semester only, no course data at all.
        ["SPRING 2026", "Fa-2025/BSCS/003", "STUDENT GAMMA", 0, "-", "-", "-", "-", "-", "BSCS", 0, "S26"],
        # Row 16: nothing but a grade, no student and no course.
        ["SPRING 2026", "-", "-", 0, "-", "-", "-", "-", "A", "-", 0, "S26"],
    ], columns=list(base.columns))
    frame = pd.concat([base, extra], ignore_index=True)
    frame["CRH"] = frame["CRH"].astype(object)
    return frame


@pytest.fixture
def error_dataset(tmp_path):
    source = tmp_path / "errors.xlsx"
    _error_frame().to_excel(source, index=False)
    return load_course_workbook(source)


def _issue(dataset, code, row):
    matches = [i for i in dataset.issues if i.code == code and i.excel_row == row]
    assert len(matches) == 1, f"expected one {code} at row {row}, found {len(matches)}"
    return matches[0]


def test_name_conflict_guidance_names_majority_and_row(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "STUDENT_NAME_CONFLICT", 2))
    assert "STUDENT ALPHA" in fix
    assert "SOMEONE ELSE" in fix
    assert "Change the Student Name cell to STUDENT ALPHA in Excel row 6." in fix


def test_missing_credits_guidance_names_dominant_value(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "INVALID_CREDITS", 7))
    assert "Set the CRH cell to 4 in Excel row 7." in fix
    assert "CS101" in fix


def test_missing_course_guidance_identifies_unique_candidate(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "MISSING_REQUIRED", 10))
    assert "CS101" in fix and "Course Two" in fix
    assert "Check the source records" in fix


def test_missing_course_guidance_falls_back_when_not_identifiable(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "MISSING_REQUIRED", 8))
    assert "Fill in the course code and course name from the source records" in fix
    assert "CS101" not in fix and "CS100" not in fix


def test_no_valid_credit_hours_anywhere_invents_nothing(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "INVALID_CREDITS", 11))
    assert "No credit hours for course CS999 (New Course) appear anywhere in this file" in fix
    assert "course catalog" in fix
    assert "Set the CRH cell to" not in fix


def test_empty_rows_guidance_names_the_row_to_delete(error_dataset):
    fix = fix_guidance(error_dataset, _issue(error_dataset, "MISSING_REQUIRED", 15))
    assert "Excel row 15 has a student and a semester but no course" in fix
    assert "Delete the row from the source workbook." in fix
    fix = fix_guidance(error_dataset, _issue(error_dataset, "MISSING_REQUIRED", 16))
    assert "Excel row 16 has no registration, student name or course, only a grade A." in fix
    assert "so delete the row from the source workbook." in fix


def test_warning_issues_get_no_guidance(error_dataset):
    warnings = [i for i in error_dataset.issues if i.level == "WARNING"]
    assert warnings
    assert all(fix_guidance(error_dataset, issue) is None for issue in warnings)


def test_skip_note_includes_fix_lines(error_dataset, tmp_path):
    from course_reports import export_gazette_pack
    folder = tmp_path / "pack"
    written = export_gazette_pack(error_dataset, error_dataset.courses, folder,
                                  settings={"gazette_term": "S26"})
    assert written
    note = folder / "skipped-programs.txt"
    assert note.exists()
    text = note.read_text(encoding="utf-8")
    assert "BSCS" in text
    assert "How to fix:" in text
    assert "STUDENT ALPHA" in text
    assert "Set the CRH cell to 4 in Excel row 7." in text
    assert "No credit hours for course CS999 (New Course) appear anywhere in this file" in text
    assert "CS101" in text
    assert "\u2014" not in text
