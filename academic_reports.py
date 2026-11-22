"""Non-official DRAFT gazette exports for verified course-level imports.

Uses explicit user-confirmed GPA rules with source provenance. Academic
standing is printed in COMMENTS from the CGPA and the probation cutoff
(operator's probation_cgpa setting, else 2.00); merit and pass/fail
decisions are not inferred.
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from academic_engine import CourseAttempt, ResultDataset
from calculation_policy import calculation_attempts
from academic_identity import program_key,registration_key,student_count


COLS = ("Term", "Course Code", "Course Name", "CRH", "Marks", "Grade", "Source GPs")


def _group(courses: list[CourseAttempt]):
    groups = defaultdict(list)
    labels={}
    for c in courses:
        key=(c.admission_session,program_key(c.degree),registration_key(c.registration))
        labels.setdefault(key,(c.admission_session,c.degree,c.registration,c.name))
        groups[key].append(c)
    return sorted([(labels[key],values) for key,values in groups.items()], key=lambda g: (g[0][0], g[0][1], g[0][2]))


def _row(course: CourseAttempt):
    return (course.term, course.course_code, course.course_name,
            course.credits, course.marks, course.grade, course.quality_points)


def _san(v):
    return str(v if v is not None else "")


def _display_source_number(value):
    """Remove float storage noise for DISPLAY ONLY; original data unchanged.

    This is not an SGPA/CGPA calculation or rounding policy. Only near-exact
    binary floating point tails are shortened (e.g. 8.100000000000001 -> 8.1).
    """
    s = str(value)
    try:
        number = Decimal(s)
        if not number.is_finite():
            return s
        for places in range(7):
            precision = Decimal(1).scaleb(-places)
            rounded = number.quantize(precision)
            if abs(number - rounded) <= Decimal("0.000000000001"):
                return format(rounded, "f").rstrip("0").rstrip(".") if "." in format(rounded, "f") else format(rounded, "f")
    except (InvalidOperation, ValueError):
        pass
    return s


def _excel_safe(value):
    """Do not turn source student text into spreadsheet formulas on export."""
    v = str(value) if value is not None else ""
    if v.lstrip().startswith(("=", "+", "-", "@", "\t", "\r", "\n")):
        return "'" + v
    return v


def _write_gazette_excel_sheet(workbook, dataset, courses, template_key, settings):
    """The Gazette sheet, in the official LGU layout (the same table as the PDF).

    The title (LAHORE GARRISON UNIVERSITY / RESULT STATEMENT / DEPARTMENT OF
    ... / SESSION - ...) appears once at the top only, never repeated on the
    sheet or on printed pages. Columns are the official ones: SR, DEGREE ID,
    REGISTRATION NUMBERS, STUDENT NAME, SEM, COURSE CODE, COURSE NAME, GRADE,
    GP (GP, Adj., T.GP), Chr (CH, Adj., T.CH), SGPA, Adj. GPA, COMMENTS.
    Each student is one block: identity once, one row per course, totals on
    the last course row, Adj. columns blank as in the official sheet, blocks
    padded to 10 rows with '-' rows. Numbers are clean 2-decimal text, never
    raw floats. The audit sheets keep the exact values and per-course points.
    """
    from collections import defaultdict
    from fractions import Fraction
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from academic_identity import program_key
    from calculation_policy import (AcademicIndex, CalculationReviewRequired,
                                    calculation_attempts, display_ratio,
                                    standing_comment)
    from fee_register import selected_defaulters
    from gazette_layout import _cohort_short, _full_term, _plain_number, _profile, _window, term_order
    from gazette_options import current_term, resolve_template_groups
    from semester_reports import roll_order

    # The 17 official columns (the wide layout: GP and Chr each split in 3).
    columns = [
        'SR', 'DEGREE ID', 'REGISTRATION NUMBERS', 'STUDENT NAME',
        'SEM', 'COURSE CODE', 'COURSE NAME', 'GRADE',
        'GP', 'Adj.', 'T.GP', 'CH', 'Adj.', 'T.CH', 'SGPA', 'Adj. GPA', 'COMMENTS',
    ]
    width = len(columns)
    dark_green = '0D542B'
    header_fill = PatternFill('solid', fgColor=dark_green)
    thin = Side(style='thin', color='000000')
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    worksheet = workbook.create_sheet('Gazette', 0)
    workbook.active = 0
    worksheet.sheet_view.showGridLines = False
    worksheet.page_setup.orientation = 'landscape'
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_margins.left = worksheet.page_margins.right = 0.25
    worksheet.page_margins.top = worksheet.page_margins.bottom = 0.4
    # Title only at the top: no print_title_rows, nothing repeated.

    groups = resolve_template_groups(template_key, courses)
    departments = set()
    cohorts = set()
    for key, group_rows in groups:
        for cohort, degree in {(c.admission_session, c.degree) for c in group_rows}:
            profile = _profile(key, cohort, degree)
            departments.update(item['text'].strip() for item in profile['header']
                               if 'DEPARTMENT OF' in item['text'])
            cohorts.add(cohort)
    department_label = (next(iter(departments)) if len(departments) == 1
                        else 'DEPARTMENTAL GAZETTE')
    # Collapse the double spaces the official header text carries.
    department_label = ' '.join(department_label.split())
    term = current_term(dataset, settings or {})
    if len(cohorts) == 1:
        session_line = f'SESSION - {next(iter(cohorts))} - {{ SEMESTER - {_full_term(term)} }}'
    else:
        session_line = f'SESSION - {{ SEMESTER - {_full_term(term)} }}'

    navy = '17365D'
    title = worksheet.cell(1, 1, 'LAHORE GARRISON UNIVERSITY')
    title.font = Font(name='Cambria', size=18, bold=True, color=navy)
    title.alignment = Alignment(horizontal='center', vertical='center')
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    worksheet.cell(2, 1, 'RESULT STATEMENT')
    worksheet.cell(2, 1).font = Font(name='Cambria', size=15, bold=True)
    worksheet.cell(2, 1).alignment = Alignment(horizontal='center', vertical='center')
    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    worksheet.cell(3, 1, department_label)
    worksheet.cell(3, 1).font = Font(name='Cambria', size=13, bold=True, color=navy)
    worksheet.cell(3, 1).alignment = Alignment(horizontal='center', vertical='center')
    worksheet.merge_cells(start_row=3, start_column=1, end_row=3, end_column=width)
    worksheet.cell(4, 1, session_line)
    worksheet.cell(4, 1).font = Font(name='Arial', size=10)
    worksheet.cell(4, 1).alignment = Alignment(horizontal='center', vertical='center')
    worksheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=width)

    # Group header row (GP over its 3 columns, Chr over its 3) and label row.
    for start, end, label in ((9, 11, 'GP'), (12, 14, 'Chr')):
        worksheet.merge_cells(start_row=5, start_column=start, end_row=5, end_column=end)
        cell = worksheet.cell(5, start, label)
        cell.font = Font(name='Arial', size=9, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for column, value in enumerate(columns, 1):
        cell = worksheet.cell(6, column, value)
        cell.font = Font(name='Arial', size=9, bold=True, color='000000')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True,
                                   text_rotation=90 if value == 'GRADE' else 0)
    worksheet.row_dimensions[6].height = 36

    calculations = AcademicIndex(calculation_attempts(dataset))
    fees = selected_defaulters(dataset, courses, term)
    withheld = {registration_key(record.registration) for record in fees}
    contents = []
    gazette_rows = []
    blocks = []
    serial = 0

    for key, department_rows in groups:
        cohort_groups = defaultdict(list)
        for course in department_rows:
            if registration_key(course.registration) not in withheld:
                cohort_groups[(course.admission_session, program_key(course.degree))].append(course)

        for (cohort, degree_id), cohort_rows in sorted(
                cohort_groups.items(),
                key=lambda item: (term_order(_cohort_short(item[0][0])), item[0][1])):
            degree = cohort_rows[0].degree
            reg_groups = defaultdict(list)
            for course in cohort_rows:
                reg_groups[registration_key(course.registration)].append(course)
            start_index = len(gazette_rows)
            # The gazette is for one publication term only, like the PDF:
            # a student's full history is not printed here. Summer folds
            # into its Spring (SUyy -> Syy), so an F21 gazette excludes
            # the SU21 courses that belong to S21, and an S25 gazette
            # shows S25 + SU25 only. Students with no course in the
            # window do not appear in this gazette at all.
            available_terms = {c.term.upper() for c in cohort_rows}
            publication_window = _window(term, available_terms)
            for registration_id, attempts in sorted(
                    reg_groups.items(), key=lambda item: roll_order(item[1][0].registration)):
                attempts = [c for c in attempts if _window(c.term, available_terms) == publication_window]
                if not attempts:
                    continue
                serial += 1
                student_name = attempts[0].name
                registration = attempts[0].registration
                terms = defaultdict(list)
                for course in attempts:
                    terms[course.term].append(course)
                block_start = len(gazette_rows)
                for attempt_term, term_courses in sorted(terms.items(),
                                                        key=lambda item: term_order(item[0])):
                    term_courses.sort(key=lambda c: (c.source_excel_row, c.course_code.casefold()))
                    totals = None
                    comment = ''
                    try:
                        semester = calculations.summary(registration, degree, attempt_term)
                        cumulative = calculations.summary(registration, degree, attempt_term, True)
                        totals = {
                            'gp': display_ratio(Fraction(semester.total_grade_points)),
                            'ch': _plain_number(semester.total_credit_hours),
                            'tgp': display_ratio(Fraction(cumulative.total_grade_points)),
                            'tch': _plain_number(cumulative.total_credit_hours),
                            'sgpa': display_ratio(semester.exact_ratio),
                            'adj_gpa': display_ratio(cumulative.exact_ratio),
                        }
                        # The official gazettes print the standing in
                        # COMMENTS (G.A.ST, or PROBATION below the cutoff).
                        # The cutoff is the operator's probation_cgpa
                        # setting, else 2.00; it is stated on the READ
                        # FIRST sheet so a different line is a settings
                        # change, not a guess.
                        cutoff = (settings or {}).get('probation_cgpa')
                        comment = standing_comment(cumulative.exact_ratio, cutoff)
                    except CalculationReviewRequired as error:
                        totals = {'gp': 'REVIEW', 'ch': 'REVIEW', 'tgp': 'REVIEW', 'tch': 'REVIEW', 'sgpa': 'REVIEW', 'adj_gpa': 'REVIEW'}
                        comment = f'GPA/CGPA review: {error}'
                    for course_index, course in enumerate(term_courses):
                        values = [''] * width
                        if len(gazette_rows) == block_start:
                            values[0] = serial
                            values[1] = course.degree
                            values[2] = registration
                            values[3] = student_name
                        values[4] = course.term.upper()
                        values[5] = course.course_code
                        values[6] = course.course_name.upper()
                        values[7] = course.grade
                        if course_index == len(term_courses) - 1 and totals is not None:
                            values[8] = totals['gp']
                            values[10] = totals['tgp']
                            values[11] = totals['ch']
                            values[13] = totals['tch']
                            values[14] = totals['sgpa']
                            values[15] = totals['adj_gpa']
                            values[16] = comment
                        gazette_rows.append([
                            _excel_safe(value) if isinstance(value, str) else value
                            for value in values
                        ])
                # Pad the block to 10 rows with '-' rows, like the official sheet.
                while len(gazette_rows) - block_start < 10:
                    gazette_rows.append(['', '', '', '', '', '-', '-', '-', '', '', '', '', '', '', '', '', ''])
                blocks.append((block_start, len(gazette_rows) - block_start))
            first_result_row = 7 + start_index
            contents.append((key.replace('-', ' ').title(), cohort, degree,
                             first_result_row, len(gazette_rows) - start_index,
                             len(reg_groups)))

    for row in gazette_rows:
        worksheet.append(row)
    result_end_row = 6 + len(gazette_rows)

    # Identity spans the first 4 rows of each block; NOTE ( If any ) fills
    # the SR column below it, rotated, as in the official sheet.
    for start_index, row_count in blocks:
        first = 7 + start_index
        for col in range(1, 5):
            worksheet.merge_cells(start_row=first, start_column=col, end_row=first + 3, end_column=col)
            worksheet.cell(row=first, column=col).alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True)
        if row_count > 4:
            worksheet.merge_cells(start_row=first + 4, start_column=1,
                                  end_row=first + row_count - 1, end_column=1)
            note = worksheet.cell(row=first + 4, column=1, value='NOTE ( If any )')
            note.font = Font(name='Arial', size=9, bold=True)
            note.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)

    # Plain black grid over the whole table (no Excel Table: it cannot hold
    # merged cells, and the official sheet is a plain bordered grid).
    for r in range(5, result_end_row + 1):
        for c in range(1, width + 1):
            worksheet.cell(row=r, column=c).border = cell_border

    current_row = result_end_row + 2
    fee_title = worksheet.cell(current_row, 1, 'FEE DEFAULTERS')
    fee_title.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    fee_title.fill = PatternFill('solid', fgColor='333333')
    worksheet.merge_cells(start_row=current_row, start_column=1,
                          end_row=current_row, end_column=width)
    current_row += 1
    fee_columns = ['SR', 'DEGREE ID', 'REGISTRATION NUMBERS', 'STUDENT NAME', 'COMMENTS']
    fee_positions = [2, 3, 4, 5, 17]
    for col, value in zip(fee_positions, fee_columns, strict=True):
        cell = worksheet.cell(current_row, col, value)
        cell.font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
        cell.fill = header_fill
    current_row += 1
    if fees:
        for index, record in enumerate(fees, 1):
            values = [index, record.degree, record.registration, record.name, record.comment]
            for col, value in zip(fee_positions, values, strict=True):
                cell = worksheet.cell(current_row, col, _excel_safe(value))
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            current_row += 1
    else:
        message = ('Fee status data was not supplied.' if not dataset.fee_records
                   else 'No selected registrations are marked as fee defaulters.')
        worksheet.cell(current_row, 2, message)
        worksheet.merge_cells(start_row=current_row, start_column=2,
                              end_row=current_row, end_column=width)

    contents_sheet = workbook.create_sheet('Gazette Contents', 1)
    contents_sheet.sheet_view.showGridLines = False
    contents_sheet['A1'] = 'GAZETTE TABLE OF CONTENTS'
    contents_sheet['A1'].font = Font(name='Arial', size=15, bold=True, color=dark_green)
    contents_headers = ['SR', 'GAZETTE', 'SESSION', 'DEGREE', 'FIRST RESULT ROW',
                        'COURSE ROWS', 'STUDENTS']
    for col, value in enumerate(contents_headers, 1):
        cell = contents_sheet.cell(3, col, value)
        cell.font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for index, (department, cohort, degree, first_row, course_rows, students) in enumerate(contents, 1):
        values = [index, department, _cohort_short(cohort), degree, first_row, course_rows, students]
        row_number = 3 + index
        for col, value in enumerate(values, 1):
            contents_sheet.cell(row_number, col, value)
        link = contents_sheet.cell(row_number, 5)
        link.hyperlink = f"#'Gazette'!A{first_row}"
        link.style = 'Hyperlink'
    for column, column_width in enumerate([8, 28, 15, 24, 18, 16, 12], 1):
        contents_sheet.column_dimensions[get_column_letter(column)].width = column_width
    contents_sheet.freeze_panes = 'A4'

    widths = [8, 13, 20, 24, 8, 14, 42, 8, 10, 9, 10, 9, 9, 10, 10, 11, 16]
    for col, column_width in enumerate(widths, 1):
        worksheet.column_dimensions[get_column_letter(col)].width = column_width
    worksheet.freeze_panes = 'F7'
    worksheet.sheet_properties.outlinePr.summaryBelow = True


def _guard(dataset: ResultDataset, courses):
    if not courses:
        raise ValueError("No matching records. Choose a different filter.")
    from validation_scope import relevant_errors
    serious = relevant_errors(dataset,courses,history=True)
    if serious:
        raise ValueError(f"Import has {len(serious)} blocking validation errors. Review these before export.")


def write_draft_pdf(dataset, courses, path, template_key='computer-science',settings=None):
    _guard(dataset,courses)
    from gazette_layout import write_gazette
    return write_gazette(dataset,courses,path,template_key,settings=settings)


def write_draft_word(dataset: ResultDataset, courses: list[CourseAttempt], path: str,
                     settings=None):
    from docx import Document
    from docx.shared import Inches, Pt
    _guard(dataset, courses)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(11.7)
    sec.page_height = Inches(8.3)
    sec.left_margin = sec.right_margin = Inches(0.7)
    sec.top_margin = sec.bottom_margin = Inches(0.65)
    styles = doc.styles["Normal"]
    styles.font.name = "Arial"
    styles.font.size = Pt(9)
    doc.add_heading("Lahore Garrison University", level=1)
    doc.add_heading("RESULT REGISTER", level=2)
    doc.add_paragraph(f"Source: {dataset.source_filename}")
    doc.add_paragraph(f"Students: {student_count(courses)} | Course attempts: {len(courses)}")
    from calculation_policy import AcademicIndex,display_ratio,CalculationReviewRequired,standing_comment
    calculations=AcademicIndex(calculation_attempts(dataset))
    from fee_register import selected_defaulters
    publication_term=None
    if settings is not None:
        from gazette_options import current_term
        publication_term=current_term(dataset,settings)
    fees=selected_defaulters(dataset,courses,publication_term)
    withheld={registration_key(record.registration) for record in fees}
    academic_courses=[course for course in courses if registration_key(course.registration) not in withheld]
    for (cohort, degree, reg, name), attempts in _group(academic_courses):
        doc.add_page_break()
        doc.add_heading("RESULT STATEMENT", level=2)
        doc.add_paragraph(f"Admission session: {cohort}    Degree: {degree}")
        doc.add_paragraph(f"Registration number: {reg}    Student: {name}")
        tab = doc.add_table(rows=1, cols=len(COLS))
        tab.style = "Table Grid"
        from docx.oxml import OxmlElement
        tab.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for cell, label in zip(tab.rows[0].cells, COLS, strict=True):
            cell.text = label
        for c in sorted(attempts, key=lambda x: (x.term, x.course_code, x.source_excel_row)):
            for cell, value in zip(tab.add_row().cells, _row(c), strict=True):
                cell.text = str(value)
        for term in sorted({c.term for c in attempts}):
            try:
                semester=calculations.summary(reg,degree,term)
                cumulative=calculations.summary(reg,degree,term,True)
                doc.add_paragraph(f'{term}: SGPA {display_ratio(semester.exact_ratio)} | CGPA {display_ratio(cumulative.exact_ratio)} | {standing_comment(cumulative.exact_ratio, (settings or {}).get("probation_cgpa"))}')
            except CalculationReviewRequired as error:
                doc.add_paragraph(f'{term}: REVIEW - {error}')
    doc.add_page_break()
    doc.add_heading('FEE DEFAULTERS',level=2)
    if fees:
        fee_table=doc.add_table(rows=1,cols=4);fee_table.style='Table Grid'
        for cell,label in zip(fee_table.rows[0].cells,['Registration','Student','Degree','Comments'],strict=True):cell.text=label
        for record in fees:
            for cell,value in zip(fee_table.add_row().cells,[record.registration,record.name,record.degree,record.comment],strict=True):cell.text=str(value)
    elif not dataset.fee_records:doc.add_paragraph('Fee status was not supplied.')
    else:doc.add_paragraph('No selected registrations are marked as defaulters in the supplied fee records.')
    doc.save(str(path))


def write_draft_excel(dataset: ResultDataset, courses: list[CourseAttempt], path: str,
                      template_key='auto', settings=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from calculation_policy import display_ratio, standing_cutoff
    _guard(dataset, courses)
    w = Workbook()
    summary = w.active
    summary.title = "READ FIRST"
    desc = [
        ("STATUS", "Calculated result audit"),
        ("Source workbook", dataset.source_filename),
        ("Unique students", student_count(courses)),
        ("Course attempts", len(courses)),
        ("Formula", "Total source grade points / total credit hours; highest points per credit repeat once"),
        ("Scope", "SGPA per term; CGPA imported history through term, including history outside display filter"),
        ("Display", "Two decimals, rounded half-up; exact numerator / denominator retained in calculation audit"),
        ("Academic standing",
         f"G.A.ST at CGPA {display_ratio(standing_cutoff((settings or {}).get('probation_cgpa')))} or above; PROBATION below (cutoff from Settings, default 2.00)"),
        ("Grade scale source", dataset.grade_scale_reference),
    ]
    for line in desc:
        summary.append([_excel_safe(v) if isinstance(v, str) else v for v in line])
    summary.column_dimensions["A"].width = 26
    summary.column_dimensions["B"].width = 85
    data = w.create_sheet("Course Results")
    headings = ["Admission Session", "Degree", "Registration Number", "Student Name",
                "Source Semester", "Attempt Term", "Course Code", "Course Name", "Credit Hours",
                "Total Marks", "Grade", "Source GPs", "Source Sheet", "Excel Row", "Source points per credit"]
    data.append(headings)
    for c in sorted(courses, key=lambda x: (x.admission_session, x.degree, x.registration, x.term, x.course_code)):
        vals = [c.admission_session, c.degree, c.registration, c.name, c.semester,
                c.term, c.course_code, c.course_name, c.credits, c.marks, c.grade,
                c.quality_points, c.source_sheet, c.source_excel_row, c.grade_point_rate]
        data.append([_excel_safe(x) if isinstance(x, str) else x for x in vals])
    data.freeze_panes = "E2"
    data.auto_filter.ref = data.dimensions
    for row in (summary[1], data[1]):
        for cell in row:
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="0D542B")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
    for idx, width in enumerate((21, 19, 23, 28, 19, 14, 17, 49, 15, 15, 10, 15, 18, 12), 1):
        data.column_dimensions[get_column_letter(idx)].width = width
    data.row_dimensions[1].height = 28
    from calculation_policy import AcademicIndex,display_ratio,CalculationReviewRequired,USER_CONFIRMED_POLICY
    calculations=AcademicIndex(calculation_attempts(dataset))
    totals=w.create_sheet('Calculated Totals')
    totals.append(['Registration','Degree','Term','Scope','CH','Grade points','Preview GPA','Exact numerator','Exact denominator','Policy','Review'])
    audit=w.create_sheet('Repeat Selection Audit')
    audit.append(['Registration','Degree','As of term','Scope','Course','Decision','Attempt term','CH','Points','Source file','Sheet','Row'])
    scope_labels={}
    for course in courses:
        scope_labels.setdefault((registration_key(course.registration),program_key(course.degree),course.term),(course.registration,course.degree))
    for (reg_id,degree_id,term),(reg,degree) in sorted(scope_labels.items()):
        for cumulative in (False,True):
            scope='CGPA' if cumulative else 'SGPA'
            try:
                summary_value=calculations.summary(reg,degree,term,cumulative)
                totals.append([_excel_safe(reg),_excel_safe(degree),_excel_safe(term),scope,str(summary_value.total_credit_hours),str(summary_value.total_grade_points),display_ratio(summary_value.exact_ratio),str(summary_value.exact_ratio.numerator),str(summary_value.exact_ratio.denominator),USER_CONFIRMED_POLICY.version,''])
                for decision in summary_value.decisions:
                    for status,items in [('KEPT',(decision.selected,)),('EXCLUDED',decision.excluded)]:
                        for c in items:
                            audit.append([_excel_safe(v) for v in [reg,degree,term,scope,c.course_code,status,c.term,c.credits,c.quality_points,c.source_file,c.source_sheet,str(c.source_excel_row)]])
            except CalculationReviewRequired as error:
                totals.append([_excel_safe(reg),_excel_safe(degree),_excel_safe(term),scope,'','','','','',USER_CONFIRMED_POLICY.version,str(error)])
    for sheet in (totals,audit):
        sheet.freeze_panes='D2';sheet.auto_filter.ref=sheet.dimensions
        for column in sheet.columns:sheet.column_dimensions[column[0].column_letter].width=22
    if dataset.fee_records:
        fees=w.create_sheet('Fee Source Records')
        fees.append(['Registration','Name','Degree','Term','Status','Comment','Source file','Sheet','Row'])
        for record in dataset.fee_records:
            fees.append([_excel_safe(v) for v in [record.registration,record.name,record.degree,record.term,record.status,record.comment,record.source_file,record.source_sheet,str(record.source_excel_row)]])
        fees.freeze_panes='A2'
    if dataset.grade_scale or any(c.grade_point_rate for c in dataset.courses):
        derivation=w.create_sheet('Point Derivation Audit')
        derivation.append(['Registration','Degree','Term','Course','Grade','CH','Source GP','Applied rate','Calculated GP','Origin','Source file','Sheet','Row'])
        selected_ids={(registration_key(c.registration),program_key(c.degree)) for c in courses}
        effective=calculation_attempts(dataset)
        for raw,calculated in zip(dataset.courses,effective,strict=True):
            if (registration_key(raw.registration),program_key(raw.degree)) not in selected_ids:continue
            origin='Source' if raw.quality_points else 'Source grade-point rate' if raw.grade_point_rate and calculated.quality_points else 'Applied grade scale' if calculated.quality_points else 'Unresolved'
            derivation.append([_excel_safe(v) for v in [raw.registration,raw.degree,raw.term,raw.course_code,raw.grade,raw.credits,raw.quality_points,raw.grade_point_rate or dataset.grade_scale.get(raw.grade.strip().upper(),''),calculated.quality_points,origin,raw.source_file,raw.source_sheet,str(raw.source_excel_row)]])
        derivation.freeze_panes='E2';derivation.auto_filter.ref=derivation.dimensions
        for column in derivation.columns:derivation.column_dimensions[column[0].column_letter].width=22
        scale=w.create_sheet('Applied Grade Scale')
        scale.append(['Reference',_excel_safe(dataset.grade_scale_reference)])
        scale.append(['Grade','Points per credit'])
        for grade,rate in sorted(dataset.grade_scale.items()):scale.append([_excel_safe(grade),rate])
        scale.append(['Method','Missing source points only = source credits x entered rate. Original course rows remain unchanged.'])
    _write_gazette_excel_sheet(w,dataset,courses,template_key,settings)
    w.save(str(path))


def generate_draft(fmt: str, dataset: ResultDataset, courses: list[CourseAttempt], path: str, template_key="computer-science",settings=None):
    if settings is not None:
        from gazette_options import gazette_rows
        courses=gazette_rows(dataset,courses,settings)
    p = Path(path)
    from export_io import protect_source,atomic_output
    protect_source(dataset,p)
    expected = {"pdf": (".pdf", write_draft_pdf),
                "word": (".docx", write_draft_word),
                "excel": (".xlsx", write_draft_excel)}
    if fmt not in expected:
        raise ValueError(f"Unknown format: {fmt}")
    ext, fn = expected[fmt]
    if p.suffix.lower() != ext:
        raise ValueError(f"Expected a {ext} output file")
    with atomic_output(p) as temporary:
        if fmt == "pdf":fn(dataset, courses, str(temporary), template_key,settings=settings)
        elif fmt == "excel":fn(dataset,courses,str(temporary),template_key,settings=settings)
        else:fn(dataset,courses,str(temporary),settings=settings)
