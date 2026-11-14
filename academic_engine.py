"""Reliable LGU course-level Excel reader (separate from legacy wide importer).

A row is a course attempt, not a student. Parsing preserves source values and
provenance; calculation policies live separately in calculation_policy.py.
"""
from __future__ import annotations
from excel_io import excel_engine

from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from functools import lru_cache
import re
import unicodedata

import pandas as pd
from academic_identity import program_key,registration_key
from grading_profiles import default_grade_scale,default_grade_reference


ALIASES = {
    "registration": ("registration", "registration numbers", "registration number", "registration no", "reg no", "reg number", "student id", "roll number", "roll no", "reg id", "registration id", "regd no", "enrollment number"),
    "name": ("student name", "candidate name", "name of student", "full name"),
    "course_code": ("course code", "subject code", "course id"),
    "course_name": ("course name", "subject name", "course title", "subject title"),
    "term": ("term", "short semester", "term code", "semester code", "exam term"),
    "semester": ("semester", "academic semester", "academic term"),
    "credits": ("crh", "credit hours", "credit hour", "chr", "ch", "credits", "cr hrs", "credit hrs", "cr hours"),
    "grade": ("result grade", "grade", "letter grade", "final grade"),
    "marks": ("total marks", "marks obtained", "marks", "score"),
    "grade_point_rate": ("grade point rate", "points per credit", "gp per credit", "grade points per credit", "grade point value"),
    "quality_points": ("gps", "total grade points", "total gp", "quality points", "grade points earned"),
    "degree": ("degree title", "degree", "degree id", "program", "programme"),
    "department": ("department", "department name", "faculty"),
    "award_id": ("award id", "award number", "award no"),
    "admission_session": ("admission session", "admission cohort", "intake", "batch"),
    "semester_number": ("semester number", "semester no", "study semester", "academic semester number"),
    "fee_status": ("fee status", "fee defaulter", "defaulter", "fee clearance status"),
    "fee_comment": ("fee comment", "fee remarks"),
}
NEEDED = frozenset(("registration", "name", "course_code", "course_name", "grade", "credits"))
ALLOWED_GRADES = frozenset(("A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "F", "IP", "I", "W", "PASS", "P", "R"))
PLACEHOLDERS = frozenset(("", "-", "--", "NONE", "NAN", "N/A", "NULL"))
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_COURSE_ROWS = 200_000


def _key(v: object) -> str:
    return _text_key(str(v) if v is not None else '')

@lru_cache(maxsize=8192)
def _text_key(value):
    s = unicodedata.normalize("NFKC", value)
    return re.sub(r"[^a-z0-9]+", " ", s.casefold()).strip()


def _value(v: object) -> str:
    if v is None or pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _empty(v: object) -> bool:
    return _value(v).upper() in PLACEHOLDERS


def identify_columns(labels) -> dict[str, str]:
    """Exact normalized aliases only; do not guess CGPA or marks from ranges."""
    normalized = {_key(c): str(c) for c in labels if _key(c)}
    found = {}
    used = set()
    for role, aliases in ALIASES.items():
        for alias in aliases:
            candidate = normalized.get(_key(alias))
            if candidate and candidate not in used:
                found[role] = candidate
                used.add(candidate)
                break
    return found


@dataclass(frozen=True)
class SheetLayout:
    sheet: str
    header_row: int  # 0-based
    columns: dict[str, str]
    row_count_estimate: int
    source_file: str = ""
    header_rows: int = 1
    defaults: dict[str,str] = field(default_factory=dict)
    course_blocks: list[dict] = field(default_factory=list)
    confirmed: bool = False


@dataclass(frozen=True)
class FeeRecord:
    registration: str
    name: str
    degree: str
    term: str
    status: str
    comment: str = ""
    source_file: str = ""
    source_sheet: str = ""
    source_excel_row: int = 0


@dataclass(frozen=True)
class DataIssue:
    level: str
    sheet: str
    excel_row: int
    code: str
    description: str
    source_file: str = ""


@dataclass(frozen=True)
class CourseAttempt:
    registration: str
    name: str
    degree: str
    admission_session: str
    semester: str
    term: str
    course_code: str
    course_name: str
    credits: str
    marks: str
    grade: str
    quality_points: str
    award_id: str
    source_sheet: str
    source_excel_row: int
    department: str = ""
    source_file: str = ""
    semester_number: str = ""
    grade_point_rate: str = ""


@dataclass
class ResultDataset:
    courses: list[CourseAttempt]
    issues: list[DataIssue]
    layouts: list[SheetLayout]
    source_filename: str
    selected_sheets: list[str] = field(default_factory=list)
    skipped_placeholder_rows: int = 0
    quarantined_rows: int = 0
    source_rows: list[dict] = field(default_factory=list)
    fee_records: list[FeeRecord] = field(default_factory=list)
    grade_scale: dict[str,str] = field(default_factory=default_grade_scale)
    grade_scale_reference: str = field(default_factory=default_grade_reference)
    grade_scale_user_set: bool = False
    source_paths: list[str] = field(default_factory=list)

    @property
    def students(self):
        labels={}
        for course in self.courses:labels.setdefault(registration_key(course.registration),course.registration)
        return set(labels.values())

    @property
    def terms(self):
        return sorted({c.term for c in self.courses if c.term})

    @property
    def degrees(self):
        labels={}
        for course in self.courses:
            if course.degree:labels.setdefault(program_key(course.degree),course.degree)
        return sorted(labels.values(),key=str.casefold)

    @property
    def cohorts(self):
        return sorted({c.admission_session for c in self.courses if c.admission_session})

    def filtered(self, *, term="All Terms", degree="All Degrees", cohort="All Cohorts"):
        return [c for c in self.courses
                if (term == "All Terms" or c.term == term)
                and (degree == "All Degrees" or program_key(c.degree) == program_key(degree))
                and (cohort == "All Cohorts" or c.admission_session == cohort)]

    def counts(self):
        return {"students": len(self.students), "course_attempts": len(self.courses),
                "terms": len(self.terms), "sheets": len(self.layouts),
                "issues": len(self.issues),
                "placeholders": self.skipped_placeholder_rows,
                "quarantined": self.quarantined_rows,
                "errors": sum(i.level == "ERROR" for i in self.issues),
                "warnings": sum(i.level == "WARNING" for i in self.issues)}


def detect_layouts(path: str | Path, *, selected_sheet: str | None = None) -> list[SheetLayout]:
    p = Path(path)
    if not p.is_file() or p.suffix.lower() not in (".xlsx", ".xlsm", ".xls", ".xlsb"):
        raise ValueError("Select an Excel workbook (.xlsx, .xlsm, .xls or .xlsb).")
    if p.stat().st_size > MAX_FILE_BYTES:
        raise ValueError("Workbook exceeds the 50 MB safety limit; split it first.")
    layouts = []
    with pd.ExcelFile(p, engine=excel_engine(path)) as xl:
        if selected_sheet is not None and selected_sheet not in xl.sheet_names:
            raise ValueError(f"Worksheet not found: {selected_sheet}")
        sheets = [selected_sheet] if selected_sheet is not None else xl.sheet_names
        for sheet_name in sheets:
            # Titles and institutional cover material can precede the table.
            preview = pd.read_excel(xl, sheet_name=sheet_name, header=None,
                                    nrows=200, dtype=object)
            from workbook_schema import header_labels,load_profiles,profile_key,default_profile_path
            profiles=load_profiles(default_profile_path());candidates=[]
            for hr in range(len(preview)):
                for depth in (1,2):
                    if hr+depth>len(preview):continue
                    if depth==2 and (not any(_key(v) for v in preview.iloc[hr].tolist()) or NEEDED.issubset(identify_columns(preview.iloc[hr].tolist())) or NEEDED.issubset(identify_columns(preview.iloc[hr+1].tolist()))):continue
                    vals=preview.iloc[hr:hr+depth].values.tolist()
                    labels=header_labels(vals);saved=profiles.get(profile_key(labels,depth))
                    if saved and not saved.get('requires_review') and not saved.get('defaults'):
                        layout=SheetLayout(sheet_name,hr,saved['columns'],0,Path(path).name,depth,saved.get('defaults',{}),saved.get('course_blocks',[]),True)
                        candidates.append((100,-hr,-depth,layout));continue
                    raw=preview.iloc[hr].tolist() if depth==1 else labels
                    cols=identify_columns(raw)
                    if not NEEDED.issubset(cols):continue
                    keys=[_key(v) for v in raw if _key(v)]
                    ambiguous=[role for role,column in cols.items() if keys.count(_key(column))>1]
                    if ambiguous:raise ValueError('Ambiguous duplicate headings: '+', '.join(ambiguous)+'. Choose the correct columns in import review.')
                    candidates.append((len(cols),-hr,-depth,SheetLayout(sheet_name,hr,cols,0,Path(path).name,depth)))
            if candidates:layouts.append(max(candidates,key=lambda item:item[:3])[3])
    return layouts


def admission_session(registration: str) -> str:
    match = re.match(r"\s*(Fa|Sp|Su|Fall|Spring|Summer)\s*[-_/ ]\s*(\d{2,4})", registration, re.I)
    if not match:
        return "Unknown admission session"
    kind = match.group(1).casefold()
    term = {"fa":"FALL", "fall":"FALL", "sp":"SPRING", "spring":"SPRING",
            "su":"SUMMER", "summer":"SUMMER"}[kind]
    year = match.group(2)
    year = "20" + year if len(year) == 2 else year
    return f"{term} {year}"

def registration_program(registration: str) -> str:
    """Read the explicit program segment from a roll number when available."""
    parts=[part.strip() for part in re.split(r'[/\\]',str(registration)) if part.strip()]
    return parts[1] if len(parts)>=3 else ""


def derive_term(semester: str) -> str:
    m = re.search(r"\b(FALL|SPRING|SUMMER)\s+((?:19|20)\d{2})\b", semester.upper())
    if not m:
        return "Unspecified term"
    return {"FALL":"F", "SPRING":"S", "SUMMER":"SU"}[m.group(1)] + m.group(2)[-2:]


def _numeric_text(v: object) -> str:
    s = _value(v)
    if not s or s.upper() in PLACEHOLDERS:
        return ""
    # Keep all source digits and decimals; do not round or change academic values.
    return s


def _validate_decimal(raw: str, *, minimum: Decimal, maximum: Decimal) -> bool:
    try:
        value = Decimal(raw)
        return value.is_finite() and minimum <= value <= maximum
    except (InvalidOperation, ValueError):
        return False


def load_course_workbook(path: str | Path, *, selected_sheet: str | None = None, layouts_override: list[SheetLayout] | None = None) -> ResultDataset:
    """Load recognized course-level sheets; quarantine invalid/ambiguous records.

    No student is identified by name alone. No numeric GPA/CGPA values are made.
    Duplicate course attempts and name conflicts are flagged for human review.
    """
    layouts = layouts_override if layouts_override is not None else detect_layouts(path, selected_sheet=selected_sheet)
    if not layouts:
        raise ValueError("No course-level worksheet detected. Required headings: Registration Numbers, Student Name, Course Code, Course Name, CRH and Grade.")
    records = []
    fee_records = []
    source_rows = []
    issues = []
    names_by_roll = defaultdict(set)
    degrees_by_roll = defaultdict(set)
    occurrences = defaultdict(list)
    skipped_placeholder_rows = 0
    quarantined_rows = 0
    with pd.ExcelFile(path, engine=excel_engine(path)) as xl:
        for layout in layouts:
            from workbook_schema import read_layout_table
            df,c,audit_rows,excel_rows=read_layout_table(xl,layout)
            if len(df)>MAX_COURSE_ROWS:raise ValueError(f"Sheet '{layout.sheet}' exceeds the course-attempt processing limit.")
            for item in audit_rows:item['file']=Path(path).name
            source_rows.extend(audit_rows)
            labels=list(df.columns)
            positions={role:labels.index(column) for role,column in c.items()}
            registration_headings={_key(alias) for alias in ALIASES['registration']}
            filename=Path(path).name
            if 'quality_points' not in c and 'grade_point_rate' not in c:
                issues.append(DataIssue('INFO',layout.sheet,layout.header_row+1,'GRADE_POINTS_NOT_SUPPLIED','Source grade-point column is not supplied; grades and marks are retained without inventing a grading scale.',filename))
            for index, row in enumerate(df.itertuples(index=False,name=None)):
                excel_row = excel_rows[index]
                mapped={role:_value(row[position]) for role,position in positions.items()}
                get = lambda role: mapped.get(role,'')
                if any(_key(value) in registration_headings for value in row) and NEEDED.issubset(identify_columns(row)):
                    issues.append(DataIssue('WARNING',layout.sheet,excel_row,'REPEATED_HEADER','Repeated table heading retained in source audit and excluded from course attempts.'))
                    continue
                reg, name = get("registration"), get("name")
                course_code, course_name = get("course_code"), get("course_name")
                credits = _numeric_text(row[positions["credits"]])
                raw_term = get("term") if not _empty(get("term")) else get("semester")
                compact_term = re.fullmatch(r'(SU|S|F)[\s_/-]*(\d{2}|\d{4})',raw_term,re.I)
                term = compact_term[1].upper()+compact_term[2][-2:] if compact_term else derive_term(raw_term)
                degree = get("degree") if not _empty(get("degree")) else registration_program(reg)
                degree = degree or "Unspecified degree"
                if get("grade").strip().upper() == "FEE" and not _empty(reg) and not _empty(name):
                    fee_records.append(FeeRecord(
                        reg, name, degree, term, "DEFAULTER", "FEE",
                        Path(path).name, layout.sheet, excel_row))
                    continue
                if all(_empty(x) for x in (reg, name, course_code, course_name, credits, get("grade"))):
                    # Real sheets may have 100s of formatting-only or dash rows.
                    skipped_placeholder_rows += 1
                    continue
                if any(_empty(x) for x in (reg, name, course_code, course_name)):
                    issues.append(DataIssue("ERROR", layout.sheet, excel_row,
                                            "MISSING_REQUIRED", "Course row is missing a valid registration, student name, course code or course name."))
                    quarantined_rows += 1
                    continue
                creditless_status=_empty(credits) and get('grade').upper() in ('IP','I','W','PASS','P','U')
                if creditless_status:
                    issues.append(DataIssue('INFO',layout.sheet,excel_row,'NON_GPA_CREDITS_UNAVAILABLE','Status-only course attempt retained with source credit hours unavailable; no GPA inclusion assumed.'))
                elif not _validate_decimal(credits, minimum=Decimal(0), maximum=Decimal(50)):
                    issues.append(DataIssue("ERROR", layout.sheet, excel_row,
                                            "INVALID_CREDITS", "Credit hours are missing or not within the accepted data range of 0-50; no assumption made."))
                    quarantined_rows += 1
                    continue
                grade = get("grade")
                if _empty(grade):
                    issues.append(DataIssue("WARNING", layout.sheet, excel_row,
                                            "MISSING_GRADE", "Grade is blank; retained without assigning a grade."))
                    grade = ""
                elif grade.upper() not in ALLOWED_GRADES:
                    issues.append(DataIssue("WARNING", layout.sheet, excel_row,
                                            "UNRECOGNIZED_GRADE", "Grade label is not in the software's reference list; preserved exactly and requires review."))
                sem = get("semester")
                if term == "Unspecified term":
                    issues.append(DataIssue("WARNING", layout.sheet, excel_row,
                                            "MISSING_TERM", "Course-attempt semester could not be determined."))
                semester_number=get('semester_number')
                if not semester_number:
                    ordinal=re.fullmatch(r'(?:SEMESTER\s*)?(\d+)',sem,re.I)
                    if ordinal:semester_number=ordinal[1]
                if semester_number and (not semester_number.isdigit() or not 1<=int(semester_number)<=14):
                    issues.append(DataIssue('ERROR',layout.sheet,excel_row,'INVALID_SEMESTER_NUMBER','Academic semester number must be an integer from 1 to 14.'))
                if not _empty(get('fee_status')):
                    try:
                        status=normalize_fee_status(get('fee_status'))
                        fee_records.append(FeeRecord(reg,name,degree,term,status,get('fee_comment'),Path(path).name,layout.sheet,excel_row))
                    except ValueError:
                        issues.append(DataIssue('ERROR',layout.sheet,excel_row,'UNKNOWN_FEE_STATUS','Fee status is not recognized; map it explicitly before gazette export.'))
                qp = _numeric_text(row[positions["quality_points"]]) if "quality_points" in c else ""
                if qp and not _validate_decimal(qp, minimum=Decimal(-999999999), maximum=Decimal(999999999)):
                    issues.append(DataIssue("WARNING", layout.sheet, excel_row,
                                            "UNRECOGNIZED_QUALITY_POINTS", "Source grade-point value is nonnumeric; retained for review."))
                rec = CourseAttempt(
                    registration=reg, name=name, degree=degree,
                    admission_session=get('admission_session') if not _empty(get('admission_session')) else admission_session(reg), semester=sem,
                    term=term, course_code=course_code, course_name=course_name,
                    credits=credits, marks=_numeric_text(row[positions["marks"]]) if "marks" in c else "",
                    grade=grade, quality_points=qp, award_id=get("award_id"),
                    source_sheet=layout.sheet, source_excel_row=excel_row,
                    department=get("department"), source_file=Path(path).name,
                    semester_number=semester_number,grade_point_rate=get("grade_point_rate"),
                )
                records.append(rec)
                names_by_roll[registration_key(reg)].add(_key(name))
                degrees_by_roll[registration_key(reg)].add(program_key(degree))
                occurrences[(registration_key(reg), term.casefold(), course_code.casefold())].append((layout.sheet, excel_row))
    for key, locations in occurrences.items():
        if len(locations) > 1:
            for sheet, row in locations:
                issues.append(DataIssue("WARNING", sheet, row, "POSSIBLE_DUPLICATE",
                                        "Registration + term + course appears more than once; attempts retained pending review."))
    first_appearance = {}
    for rec in records:
        first_appearance.setdefault(registration_key(rec.registration), (rec.source_sheet, rec.source_excel_row))
    for reg, found in names_by_roll.items():
        if len(found) > 1:
            sheet, row = first_appearance[reg]
            issues.append(DataIssue("ERROR", sheet, row, "STUDENT_NAME_CONFLICT",
                                    "Same registration number appears with different names; manual verification required."))
    for reg, found in degrees_by_roll.items():
        if len(found) > 1:
            sheet, row = first_appearance[reg]
            issues.append(DataIssue("ERROR", sheet, row, "DEGREE_CONFLICT",
                                    "Same registration number appears with multiple degree titles; requires review."))
    if not records:
        raise ValueError("No valid course records found; see the source workbook.")
    return ResultDataset(courses=records, issues=issues, layouts=layouts,
                         source_filename=Path(path).name,
                         selected_sheets=[x.sheet for x in layouts],
                         skipped_placeholder_rows=skipped_placeholder_rows,
                         quarantined_rows=quarantined_rows, source_rows=source_rows,fee_records=fee_records,source_paths=[str(Path(path).resolve())])


def load_course_workbooks(paths,layouts_by_path=None):
    """Combine source workbooks without discarding duplicates or review rows."""
    from dataclasses import replace
    datasets=[load_course_workbook(p,layouts_override=(layouts_by_path or {}).get(str(p))) for p in paths]
    courses=[c for ds in datasets for c in ds.courses]
    issues=[replace(i,source_file=ds.source_filename) for ds in datasets for i in ds.issues]
    by_identity=defaultdict(list)
    for c in courses:by_identity[registration_key(c.registration)].append(c)
    for attempts in by_identity.values():
        if len({_key(c.name) for c in attempts})>1:
            c=attempts[0];issues.append(DataIssue('ERROR',c.source_sheet,c.source_excel_row,'CROSS_FILE_NAME_CONFLICT','Same registration has conflicting names across workbooks.',c.source_file))
        if len({program_key(c.degree) for c in attempts})>1:
            c=attempts[0];issues.append(DataIssue('ERROR',c.source_sheet,c.source_excel_row,'CROSS_FILE_DEGREE_CONFLICT','Same registration has conflicting degree titles across workbooks.',c.source_file))
    duplicates=defaultdict(list)
    for c in courses:duplicates[(registration_key(c.registration),c.term,c.course_code.casefold())].append(c)
    for attempts in duplicates.values():
        if len({c.source_file for c in attempts})>1:
            for c in attempts:issues.append(DataIssue('WARNING',c.source_sheet,c.source_excel_row,'CROSS_FILE_DUPLICATE','Repeated identity, term and course across workbooks; every row retained.',c.source_file))
    return ResultDataset(courses,issues,[replace(l,source_file=ds.source_filename) for ds in datasets for l in ds.layouts],
        f'Multiple workbooks ({len(datasets)})',[f'{ds.source_filename}: {sh}' for ds in datasets for sh in ds.selected_sheets],
        sum(ds.skipped_placeholder_rows for ds in datasets),sum(ds.quarantined_rows for ds in datasets),
        [r for ds in datasets for r in ds.source_rows], [r for ds in datasets for r in ds.fee_records],
        source_paths=[p for ds in datasets for p in ds.source_paths])


def normalize_fee_status(value):
    status=_key(value)
    if status in ('yes','y','true','1','defaulter','fee defaulter','unpaid','outstanding','overdue'):return 'DEFAULTER'
    if status in ('no','n','false','0','paid','clear','cleared','not defaulter'):return 'CLEAR'
    raise ValueError('Unrecognized fee status.')


def load_fee_workbook(path,dataset):
    """Explicit fee-list import; absent status means the selected file is a defaulter list."""
    records=[]
    if Path(path).stat().st_size>MAX_FILE_BYTES:raise ValueError('Fee workbook exceeds the 50 MB import limit.')
    identities={registration_key(c.registration):c for c in dataset.courses}
    known_terms={c.term.upper() for c in dataset.courses}
    with pd.ExcelFile(path,engine=excel_engine(path)) as workbook:
        for sheet in workbook.sheet_names:
            preview=pd.read_excel(workbook,sheet_name=sheet,header=None,nrows=200,dtype=object)
            headers=[(i,identify_columns(row.tolist())) for i,row in preview.iterrows()]
            layout=next(((i,mapping) for i,mapping in headers if 'registration' in mapping),None)
            if layout is None:continue
            i,mapping=layout
            frame=pd.read_excel(workbook,sheet_name=sheet,header=i,dtype=object)
            if len(frame)>MAX_COURSE_ROWS:raise ValueError('Fee sheet exceeds the 200,000-row import limit.')
            for row_index,row in frame.iterrows():
                get=lambda role:_value(row[mapping[role]]) if role in mapping else ''
                registration=get('registration')
                if _empty(registration):continue
                if registration_key(registration) not in identities:raise ValueError('Fee list contains an unknown registration; reconcile it with the imported results.')
                student=identities[registration_key(registration)]
                if get('name') and _key(get('name'))!=_key(student.name):raise ValueError('Fee-list name conflicts with the imported registration.')
                status=normalize_fee_status(get('fee_status')) if 'fee_status' in mapping else 'DEFAULTER'
                term=get('term')
                if term and not re.fullmatch(r'(SU|S|F)\d{2}',term,re.I):term=derive_term(term)
                if term and term.upper() not in known_terms:raise ValueError('Fee-list exam term does not match the imported results; reconcile it before export.')
                records.append(FeeRecord(registration,student.name,student.degree,term.upper(),status,get('fee_comment'),Path(path).name,sheet,int(row_index)+i+2))
    if not records:raise ValueError('No fee-list registrations found. Include a Registration Number column.')
    return records
