"""Synthetic regression tests. No real student personal data is checked in."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
import pytest
from academic_engine import (admission_session, derive_term, detect_layouts,
                             identify_columns, load_course_workbook)
from academic_reports import generate_draft, _excel_safe


def _df():
    return pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS101", "Course Two", 4, 64.0, "B+", "BSCS", 13.2, "SU26"],
        ["SPRING 2026", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS100", "Course One", 3, 0, "F", "BSCS", 0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])


def test_header_roles_not_using_fuzzy_numeric_guess():
    roles = identify_columns(_df().columns)
    assert roles["registration"] == "REGISTRATION NUMBERS"
    assert roles["quality_points"] == "GPs"
    assert "cgpa" not in roles


def test_header_in_row_six_and_multiple_sheets(tmp_path):
    path = tmp_path / "multi.xlsx"
    with pd.ExcelWriter(path) as w:
        pd.DataFrame(["Title", "University", "Year", "Other", "Information"]).to_excel(w, sheet_name="Cover", index=False, header=False)
        _df().to_excel(w, sheet_name="Results1", index=False, startrow=5)
        _df().iloc[:1].to_excel(w, sheet_name="Results2", index=False)
    layouts = detect_layouts(path)
    assert {(i.sheet, i.header_row) for i in layouts} == {("Results1", 5), ("Results2", 0)}
    dataset = load_course_workbook(path)
    assert len(dataset.courses) == 5
    assert len(dataset.students) == 2
    assert dataset.counts()["warnings"] == 2  # Same student / term / course repeated across sheets


def test_distinct_summer_attempts_and_no_gpa(tmp_path):
    p = tmp_path / "grades.xlsx"
    _df().to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert ds.counts()["students"] == 2
    assert ds.counts()["course_attempts"] == 4
    assert ds.terms == ["S26", "SU26"]
    assert len(ds.filtered(term="SU26")) == 1
    assert ds.filtered(term="S26")[1].quality_points == "10.8"
    assert all("gpa" not in c.__dict__ and "cgpa" not in c.__dict__ for c in ds.courses)


def test_ambiguous_name_conflict_blocks_export(tmp_path):
    df = _df()
    df.loc[2, "STUDENT NAME"] = "MISMATCHED STUDENT"
    p = tmp_path / "conflict.xlsx"
    df.to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert any(i.code == "STUDENT_NAME_CONFLICT" for i in ds.issues)
    with pytest.raises(ValueError, match="blocking"):
        generate_draft("excel", ds, ds.courses, str(tmp_path / "block.xlsx"))


def test_invalid_credit_hour_quarantined(tmp_path):
    df = _df()
    df["CRH"] = df["CRH"].astype(object)
    df.loc[0, "CRH"] = "WRONG"
    p = tmp_path / "bad.xlsx"
    df.to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert ds.counts()["quarantined"] == 1
    assert ds.counts()["errors"] == 1
    assert len(ds.courses) == 3


def test_duplicate_exact_term_and_course_flagged(tmp_path):
    df = pd.concat([_df(), _df().iloc[[0]]], ignore_index=True)
    p = tmp_path / "duplicated.xlsx"
    df.to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert len([i for i in ds.issues if i.code == "POSSIBLE_DUPLICATE"]) == 2


def test_placeholder_ignored(tmp_path):
    df = pd.concat([_df(), pd.DataFrame([{col: "-" for col in _df().columns}])], ignore_index=True)
    p = tmp_path / "blank.xlsx"
    df.to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert len(ds.courses) == 4
    assert ds.counts()["placeholders"] == 1


def test_admission_vs_attempt_semester():
    assert admission_session("Fa-2020/BSCS/030") == "FALL 2020"
    assert admission_session("Sp-2025/BBA/010") == "SPRING 2025"
    assert derive_term("SPRING 2022") == "S22"
    assert derive_term("SUMMER 2026") == "SU26"


def test_export_formats_and_no_formulas(tmp_path):
    p = tmp_path / "input.xlsx"
    _df().to_excel(p, index=False)
    ds = load_course_workbook(p)
    assert _excel_safe("=1+2") == "'=1+2"
    from openpyxl import load_workbook
    import fitz
    from docx import Document
    for fmt, ext in (("pdf", ".pdf"), ("word", ".docx"), ("excel", ".xlsx")):
        destination = tmp_path / ("report" + ext)
        generate_draft(fmt, ds, ds.courses, str(destination))
        assert destination.stat().st_size > 500
        if fmt == "pdf":
            with fitz.open(destination) as pdf:
                assert pdf.page_count >= 2
                assert "DRAFT" not in pdf[0].get_text()
                assert "DRAFT - GP / CH" not in " ".join(page.get_text() for page in pdf)
                assert pdf[0].rect.width == 612 and pdf[0].rect.height == 792
                assert pdf[2].rect.width == 792 and pdf[2].rect.height == 612
        elif fmt == "word":
            doc = Document(destination)
            assert "DRAFT" not in " ".join(p.text for p in doc.paragraphs)
        else:
            wb = load_workbook(destination, read_only=True)
            assert wb.sheetnames == ["READ FIRST", "Course Results", "Calculated Totals", "Repeat Selection Audit", "Point Derivation Audit", "Applied Grade Scale"]
            assert wb["Course Results"].max_row == 5


def test_no_bad_file_path(tmp_path):
    with pytest.raises(ValueError, match="xlsx"):
        detect_layouts(tmp_path / "bad.xls")
