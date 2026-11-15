from academic_identity import student_count,registration_key,program_key
"""Canonical course report routes: no academic policy is inferred."""
from collections import Counter, defaultdict
from pathlib import Path
from academic_reports import generate_draft, _excel_safe

ROUTES = {
 'register': 'Result gazette',
 'term': 'Term course register',
 'cohort': 'Admission cohort course register',
 'degree': 'Degree course register',
 'department': 'Department course register',
 'subjects': 'Course summary (source counts)',
 'repeats': 'Repeated course audit (all attempts retained)',
 'summer': 'Summer and regular attempt audit',
 'validation': 'Validation and review log',
 'mapping': 'Import mapping and provenance',
 'cgpa_list': 'Roll-number-wise GPA / CGPA by semester',
}

def _publication_block_reason(dataset, rows, errors):
    from collections import Counter
    from pathlib import Path
    from validation_scope import issue_source
    labels={
        'MISSING_REQUIRED':'missing required student/course information',
        'INVALID_CREDITS':'missing or invalid credit hours',
        'STUDENT_NAME_CONFLICT':'conflicting student names for the same roll number',
        'DEGREE_CONFLICT':'conflicting programs for the same roll number',
    }
    counts=Counter(labels.get(issue.code,issue.code.replace('_',' ').lower()) for issue in errors)
    summary=', '.join(f'{number} {label}' for label,number in counts.items())
    locations=[]
    for issue in errors:
        source,_=issue_source(dataset,issue)
        location=(
            Path(issue.source_file or dataset.source_filename).name,
            issue.sheet or 'Unknown sheet',
            issue.excel_row,
        )
        if source and source.get('file'):
            location=(Path(source['file']).name,location[1],location[2])
        if location not in locations:locations.append(location)
        if len(locations)==3:break
    location_text='; '.join(
        f'{filename}, sheet {sheet}, Excel row {row}'
        for filename,sheet,row in locations
    )
    remaining=max(0,len(errors)-len(locations))
    if remaining:location_text+=f'; and {remaining} more location(s)'
    return (
        f'Publication is paused for this selection because {len(errors)} source '
        f'{"issue" if len(errors)==1 else "issues"} need review: {summary}. '
        f'Source location(s): {location_text}. Open Review data to see the original '
        'cells and edit the source workbook, then reimport it. No grade, credit, '
        'name, or CGPA was guessed.'
    )

def capability(key, ds, rows):
    if ds is None:
        return False, 'Needs mapping: import a course workbook first.'
    if key not in ROUTES:
        reasons={
            'lgu_official':'Needs policy and template: signed official gazette format and calculation rules.',
            'gazette':'Needs policy and template: official publication remains locked; choose draft register.',
            'notification':'Needs policy: official result declaration and approved student standing.',
            'cgpa_analysis':'Needs policy: verified student CGPA; source quality points are course values.',
            'merit':'Needs policy: approved ranking, ties, repeat handling and CGPA.',
            'certificates':'Needs policy: verified merit and approved certificate template.',
            'probation':'Needs policy: approved standing rules and student CGPA.',
            'subject_toppers':'Needs policy: comparable source marks and approved ranking rules.',
            'subject_pass_pct':'Needs policy: approved pass/fail rules for each subject.',
            'pass_fail':'Needs source fields: verified student standing or an approved calculation policy.',
            'failed':'Needs source fields: verified failed-student standing.',
            'failing_subjects':'Needs policy: approved failed-course classification.',
            'supplementary':'Needs source fields: verified student standing; course grades are retained separately.',
            'cards':'Needs template: approved result-card fields; choose the draft student register.',
            'transcript':'Needs policy and template: official transcript fields and approved calculated values.',
            'notice':'Needs policy: approved summary for public notice; choose course summary draft.',
            'stats_breakdown':'Needs policy: student-level outcome classifications; use source course summaries.',
            'subject_difficulty':'Needs policy: comparable marking scales; use source course summaries.',
            'student_profile':'Needs policy: approved ranking and student totals; use student course register.',
            'session_comparison':'Not implemented for course mode: use term register and summer/regular audit.',
            'data_validation':'Available alternative: choose Validation and review log.',
        }
        return False,reasons.get(key,'Not implemented for course mode; choose a draft register.')
    if key == 'department' and not any(c.department for c in rows):
        return False, 'Needs mapping: no department column supplied. Degree does not establish department.'
    if key in ('validation', 'mapping'):
        return True, 'Available'
    if not rows:
        return False, 'No records match the current filters.'
    from validation_scope import relevant_errors
    if key=='cgpa_list':rows=expanded_history(ds,rows)
    errors=relevant_errors(ds,rows,history=key in ('register','cgpa_list'),all_terms=key=='cgpa_list')
    if errors:
        return False, _publication_block_reason(ds,rows,errors)
    if key=='register':return True, 'Ready to export. GPA uses 2 decimals, rounded half-up; exact values remain in the Excel audit.'
    return True, 'Ready to export selected source records.'

def table_data(key, ds, rows):
    if key=='cgpa_list':
        from semester_reports import semester_summary_table
        return semester_summary_table(ds,rows)
    if key == 'validation':
        head = ['Level', 'Source file / sheet', 'Excel row', 'Code', 'Finding']
        data = [[i.level, (i.source_file+' / ' if i.source_file else '')+i.sheet, i.excel_row, i.code, i.description] for i in ds.issues]
        return head, data or [['INFO', '', '', 'NO_FINDINGS', 'No automated findings; manual approval still required.']]
    if key == 'mapping':
        return ['Sheet', 'Header Excel row', 'Role', 'Source column'], [[(l.source_file+' / ' if l.source_file else '')+l.sheet, l.header_row+1, k, v] for l in ds.layouts for k,v in l.columns.items()]
    if key == 'subjects':
        groups=defaultdict(list)
        for c in rows: groups[(c.course_code,c.course_name,c.term)].append(c)
        return ['Code','Course','Term','Attempts','Distinct students','Source grade labels'], [[*k,len(v),student_count(v), ', '.join(f'{g}: {n}' for g,n in sorted(Counter(c.grade for c in v).items()))] for k,v in sorted(groups.items())]
    if key == 'repeats':
        groups=defaultdict(list)
        for c in rows: groups[(registration_key(c.registration),program_key(c.degree),c.course_code.casefold())].append(c)
        rows=[c for v in groups.values() if len(v)>1 for c in v]
    if key == 'summer':
        return ['Term','Session type from term label','Attempts','Distinct students'], [[t, 'Summer' if t.upper().startswith('SU') else 'Regular' if t.upper().startswith(('F','S')) else 'Unspecified',len(v),student_count(v)] for t,v in sorted(_terms(rows).items())]
    if not rows:
        return ['Finding'], [['No repeated course attempts in this selection.']]
    order={'term':lambda c:(c.term,c.registration,c.course_code), 'cohort':lambda c:(c.admission_session,c.registration,c.term), 'degree':lambda c:(c.degree,c.registration,c.term), 'department':lambda c:(c.department,c.registration,c.term)}.get(key,lambda c:(c.registration,c.term,c.course_code))
    return ['Registration','Student','Term','Degree','Course code','Course','CH','Grade','Source GP','Sheet / row'], [[c.registration,c.name,c.term,c.degree,c.course_code,c.course_name,c.credits,c.grade,c.quality_points,f'{c.source_file} / {c.source_sheet} / {c.source_excel_row}'] for c in sorted(rows,key=order)]

def _terms(rows):
    out=defaultdict(list)
    for c in rows: out[c.term].append(c)
    return out

def expanded_history(dataset,rows):
    identities={(registration_key(c.registration),program_key(c.degree)) for c in rows}
    return [c for c in dataset.courses if (registration_key(c.registration),program_key(c.degree)) in identities]


def export_report(key, fmt, ds, rows, path, template_key="auto",settings=None):
    from export_io import protect_source,atomic_output
    if ds is None:raise ValueError('Import a course workbook first.')
    protect_source(ds,path)
    extensions={'pdf':'.pdf','word':'.docx','excel':'.xlsx'}
    if fmt not in extensions or Path(path).suffix.lower()!=extensions[fmt]:
        raise ValueError('Choose the filename extension matching the export format.')
    if key=='register' and settings is not None:
        from gazette_options import gazette_rows
        rows=gazette_rows(ds,rows,settings)
    if key=='cgpa_list':rows=expanded_history(ds,rows)
    ok,reason=capability(key,ds,rows)
    if not ok:raise ValueError(reason)
    with atomic_output(path) as temporary:
        return _write_report(key,fmt,ds,rows,str(temporary),template_key,settings)

def export_gazette_pack(ds, rows, folder, *, settings=None):
    """Write one PDF/Word/Excel gazette per automatically routed department.

    A program whose source data needs review blocks only itself: its rows
    are left out of its department's gazette and listed in the pack's
    skipped-programs note, so the ready programs still export instead of
    one blocked program failing the whole pack.
    """
    from collections import Counter
    from gazette_options import gazette_rows, resolve_template_groups
    from academic_identity import program_key
    if ds is None:
        raise ValueError('Import a course workbook first.')
    selected=gazette_rows(ds,ds.courses,settings or {})
    groups=resolve_template_groups('auto',selected)
    output=Path(folder)
    output.mkdir(parents=True,exist_ok=True)
    written=[]
    skipped=[]
    for template,department_rows in groups:
        ready_rows=[]
        for program in sorted({program_key(c.degree) for c in department_rows}):
            program_rows=[c for c in department_rows if program_key(c.degree)==program]
            ok,reason=capability('register',ds,program_rows)
            if ok:
                ready_rows.extend(program_rows)
            else:
                display=Counter(c.degree for c in program_rows).most_common(1)[0][0]
                skipped.append((display,reason))
        if not ready_rows:
            continue
        safe=''.join(char if char.isalnum() else '_' for char in template).strip('_')
        for fmt,ext in (('pdf','.pdf'),('word','.docx'),('excel','.xlsx')):
            target=output/f'LGU_Gazette_{safe}{ext}'
            export_report('register',fmt,ds,ready_rows,str(target),
                          template_key=template,settings=settings)
            written.append(target)
    if skipped:
        note=output/'skipped-programs.txt'
        note.write_text('\n'.join(
            ['The following programs were not included in this pack; their source data needs review before they can be published:', '']
            + [f'- {program}: {reason}' for program,reason in skipped]
        ),encoding='utf-8')
    if not written:
        if skipped:
            raise ValueError('No department gazette could be exported: ' + '; '.join(f'{program}: {reason}' for program,reason in skipped))
        raise ValueError('No department gazette could be exported.')
    return written


def _write_report(key, fmt, ds, rows, path, template_key="computer-science",settings=None):
    if ds and any(Path(c.source_file).name == Path(path).name for c in ds.courses):
        raise ValueError("Output must not overwrite the original workbook.")
    ok, reason=capability(key,ds,rows)
    if not ok: raise ValueError(reason)
    if key=='register': return generate_draft(fmt,ds,rows,path,template_key=template_key,settings=settings)
    head,data=table_data(key,ds,rows)
    title='DRAFT - '+ROUTES[key]
    note=f'{len(rows)} selected course attempts; {student_count(rows)} distinct students. No academic calculations or official approval.'
    if key=='cgpa_list':note='Per-semester GPA and cumulative CGPA through each term. Best repeat counted once. Two decimals, rounded half-up; full history retained.'
    if key=='cgpa_list' and ds.grade_scale:note+=' Applied scale: '+ds.grade_scale_reference
    if fmt=='excel':
        from openpyxl import Workbook
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
        w=Workbook(); s=w.active; s.title='DRAFT report'; s.append([title]); s.append([note]); s.append(head)
        for row in data: s.append([_excel_safe(v) if isinstance(v,str) else v for v in row])
        s.freeze_panes='A4'; s.auto_filter.ref=f'A3:{get_column_letter(len(head))}{s.max_row}'
        for i,h in enumerate(head,1): s.column_dimensions[get_column_letter(i)].width=min(50,max(14,len(h)+3))
        for c in s[3]: c.font=Font(bold=True)
        w.save(path)
    elif fmt=='word':
        from docx import Document
        from docx.shared import Inches, Pt
        d=Document(); sec=d.sections[0]; sec.page_width=Inches(11.7); sec.page_height=Inches(8.3); sec.left_margin=sec.right_margin=Inches(.45)
        d.styles['Normal'].font.size=Pt(8); d.add_heading(title,0); d.add_paragraph(note)
        t=d.add_table(rows=1,cols=len(head)); t.style='Table Grid'
        for c,v in zip(t.rows[0].cells,head,strict=True): c.text=str(v)
        from docx.oxml import OxmlElement
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
        for row in data:
            for c,v in zip(t.add_row().cells,row,strict=True): c.text=str(v)
        d.save(path)
    elif fmt=='pdf':
        from fpdf import FPDF
        class PDF(FPDF):
            def header(self):
                self.set_font('LGU','B',11); self.multi_cell(0,6,title,new_x='LMARGIN',new_y='NEXT'); self.ln(2)
            def footer(self):
                self.set_y(-12); self.set_font('LGU','',8); self.cell(0,5,f'DRAFT / NOT FOR PUBLICATION | Page {self.page_no()}',align='C')
        p=PDF('L'); from pdf_fonts import add_fonts; add_fonts(p); p.set_auto_page_break(True,16); p.add_page(); p.set_font('LGU','',8); p.multi_cell(0,5,note,new_x='LMARGIN',new_y='NEXT')
        safe=lambda v:str(v)
        widths=[31,28,17,12,14,15,15,12,16,12,12,130,36] if key=='cgpa_list' else [36,34,12,15,22,78,10,12,22,36] if len(head)==10 else None
        with p.table(line_height=4, text_align='LEFT',first_row_as_headings=True, col_widths=widths) as t:
            for vals in [head]+data:
                r=t.row()
                for v in vals:r.cell(safe(v))
        p.output(path)
    else: raise ValueError('Unsupported format')
