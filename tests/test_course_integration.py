"""Regression coverage for canonical routes, provenance and actual Tk callbacks."""
from pathlib import Path
import os,sys,time
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import pytest
needs_display=pytest.mark.skipif(not os.environ.get('DISPLAY'),reason='Needs a display; run under xvfb-run or on Windows')
from academic_engine import load_course_workbook
from course_reports import ROUTES,capability,export_report,table_data
from session_store import save,restore
from test_academic_engine import _df

@pytest.fixture
def dataset(tmp_path):
    source=tmp_path/'fixture.xlsx';_df().to_excel(source,index=False)
    return load_course_workbook(source),source

def test_source_values_and_review_rows_retained(tmp_path):
    df=_df();df.loc[0,'Grade']='b';df['CRH']=df['CRH'].astype(object);df.loc[1,'CRH']='invalid'
    source=tmp_path/'source.xlsx';df.to_excel(source,index=False);ds=load_course_workbook(source)
    assert len(ds.source_rows)==4 and ds.courses[0].grade=='b'
    assert ds.source_rows[1]['cells']['CRH']=='invalid'
    assert ds.counts()['quarantined']==1
    assert capability('validation',ds,ds.courses)[0]
    assert not capability('register',ds,ds.courses)[0]

def test_all_routes_and_legacy_reasons(dataset,tmp_path):
    ds,source=dataset
    import reports
    for key in reports.REPORTS:
        ok,reason=capability(key,ds,ds.courses)
        assert isinstance(ok,bool) and reason
        if key not in ROUTES:assert not ok
    assert not capability('department',ds,ds.courses)[0]
    for key in ROUTES:
        if key=='department':continue
        head,rows=table_data(key,ds,ds.courses)
        assert head and rows
        for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
            path=tmp_path/(key+ext);export_report(key,fmt,ds,ds.courses,str(path));assert path.stat().st_size>500

def test_session_roundtrip(dataset,tmp_path):
    ds,source=dataset;path=tmp_path/'session.lguexam'
    save(path,'course',str(source),None,{},{});meta,loaded=restore(path)
    assert loaded.counts()==ds.counts() and loaded.courses[0].registration==ds.courses[0].registration
    # Saving again after restore must preserve original embedded workbook.
    save(path,'course',None,None,{},{});assert restore(path)[1].counts()==ds.counts()

def test_filtered_export_content(dataset,tmp_path):
    ds,source=dataset;rows=ds.filtered(term='SU26');path=tmp_path/'filtered.xlsx'
    export_report('register','excel',ds,rows,str(path))
    from openpyxl import load_workbook
    w=load_workbook(path);assert w['Course Results'].max_row==2
    assert w['Course Results']['F2'].value=='SU26'
    with pytest.raises(ValueError,match='original'):export_report('register','excel',ds,rows,str(source))

@needs_display
def test_tk_dashboard_report_buttons_and_worker(dataset,monkeypatch):
    from app import App
    monkeypatch.setattr(App,'_restore_autosave',lambda self:None)
    app=App();ds,source=dataset
    try:
        app.set_academic_data(ds,str(source));app.show_frame('dashboard');app.update()
        assert app.raw_df is None and app.df is None
        text=[]
        def collect(w):
            try:text.append(str(w.cget('text')))
            except Exception:pass
            for child in w.winfo_children():collect(child)
        collect(app.frames['dashboard'])
        canvas=app.frames['dashboard'].metric_canvas
        metric_text=[canvas.itemcget(i,'text') for i in canvas.find_all() if canvas.type(i)=='text']
        assert 'Distinct students' in metric_text and 'Course attempts' in metric_text
        app.show_frame('reports');app.update();panel=app.frames['reports'].course_panel
        assert panel.format_buttons[0].cget('state')=='normal'
        from course_reports import ROUTES
        panel.route_menu.set(ROUTES['department']);panel.refresh()
        assert panel.format_buttons[0].cget('state')=='disabled'
        academic=app.frames['academic'];academic.terms.set('SU26');academic.refresh_rows()
        assert len(academic.chosen_courses())==1
        called=[];ticks=[]
        app.after(20,lambda:ticks.append(True))
        app.run_job('test responsiveness',lambda:time.sleep(.25),lambda _:called.append(True))
        deadline=time.monotonic()+3
        while not called and time.monotonic()<deadline:app.update();time.sleep(.01)
        assert called and ticks
        app.show_frame('data');app.update();assert academic.winfo_ismapped()
    finally:
        # Cancel scheduled Tk callbacks to avoid callback warnings after destruction.
        for job in app.tk.call('after','info'):app.after_cancel(job)
        app.destroy()


def test_multi_workbook_snapshot_preserves_mapping_and_raw_rows(tmp_path):
    from academic_engine import load_course_workbooks
    a=tmp_path/'one.xlsx';b=tmp_path/'two.xlsx'
    _df().to_excel(a,index=False);_df().iloc[:1].to_excel(b,index=False)
    ds=load_course_workbooks([a,b])
    assert len(ds.courses)==5 and len(ds.students)==2
    assert any(i.code=='CROSS_FILE_DUPLICATE' for i in ds.issues)
    path=tmp_path/'multi.lguexam'
    save(path,'course',None,None,{}, {}, ds)
    loaded=restore(path)[1]
    assert loaded.layouts==ds.layouts and loaded.source_rows==ds.source_rows
    assert loaded.courses==ds.courses


def test_decimal_policy_boundary():
    from calculation_policy import calculate_academic_values,PolicyRequiredError
    with pytest.raises(PolicyRequiredError):calculate_academic_values([])


@needs_display
def test_lazy_report_mode_switches(dataset,monkeypatch):
    # The application has one Tk interpreter per process. Match that lifecycle
    # rather than creating a second interpreter after a destroyed CTk root.
    import os,subprocess
    if os.environ.get('LGU_GUI_ISOLATED')!='1':
        env=dict(os.environ,LGU_GUI_ISOLATED='1')
        result=subprocess.run([sys.executable,'-m','pytest',str(Path(__file__).resolve())+'::test_lazy_report_mode_switches','-q'],env=env,capture_output=True,text=True,timeout=90)
        assert result.returncode==0,result.stdout+result.stderr
        return
    from app import App
    from sample_data import generate
    import importer
    monkeypatch.setattr(App,'_restore_autosave',lambda self:None)
    app=App();ds,source=dataset
    try:
        assert not app.frames['reports'].legacy_built
        app.set_academic_data(ds,str(source));app.show_frame('reports');app.update()
        assert app.frames['reports'].course_panel.winfo_ismapped()
        panel=app.frames['reports'].course_panel
        assert 'no GPA' not in panel.reason.cget('text')
        app.frames['academic'].terms.set('S26')
        assert all(c.term=='S26' for c in app.frames['academic'].chosen_courses())
        # v4.5.2: the register panel returns the full history through the
        # S26 cycle cutoff (S26 + its Summer), not just the Summer rows.
        assert {c.term for c in panel.rows()}=={'S26','SU26'}
        app.frames['academic'].terms.set('All Terms')
        app.show_frame('settings');app.update()
        settings=app.frames['settings']
        assert settings.course_panel.winfo_ismapped() and not settings.legacy_panel.winfo_ismapped()
        assert 'S26' in settings.course_panel.detected.cget('text')
        # The Settings panel never had a template control; templates are
        # chosen in Reports, and only for one program at a time (All
        # programs always routes automatically). Pick the program, then
        # its layout, and confirm the menu shows it.
        panel.program_menu.set('BSCS');panel.refresh()
        panel.set_template('Computer Science gazette')
        app.show_frame('reports');app.update()
        assert panel.template_menu.get()=='Computer Science gazette'
        raw=generate(3,seed=2);app.set_data(raw,importer.suggest_mapping(raw),None)
        app.show_frame('reports');app.update()
        assert app.frames['reports'].legacy_built and not app.frames['reports'].course_panel.winfo_ismapped()
        app.set_academic_data(ds,str(source));app.show_frame('reports');app.update()
        deadline=time.monotonic()+5
        while not app.frames['reports'].course_panel.winfo_ismapped() and time.monotonic()<deadline:
            app.update();time.sleep(.02)
        assert app.frames['reports'].course_panel.winfo_ismapped()
    finally:
        for job in app.tk.call('after','info'):app.after_cancel(job)
        app.destroy()


def test_pack_skips_blocked_program_but_exports_ready_ones(tmp_path):
    from course_reports import export_gazette_pack
    import pandas as pd
    frame=_df().copy()
    blocked=frame.iloc[[0]].copy()
    blocked['DEGREE TITLE']='BS SE'
    blocked['REGISTRATION NUMBERS']='Fa-2025/BSSE/001'
    blocked['STUDENT NAME']='BLOCKED STUDENT'
    extra=blocked.copy();extra['STUDENT NAME']='DIFFERENT NAME'
    data=pd.concat([frame,blocked,extra],ignore_index=True)
    source=tmp_path/'pack.xlsx';data.to_excel(source,index=False)
    ds=load_course_workbook(source)
    folder=tmp_path/'pack'
    written=export_gazette_pack(ds,ds.courses,folder,settings={'gazette_term':'S26'})
    assert len(written)==3
    assert {p.name for p in written}=={'LGU_Gazette_computer_science.pdf','LGU_Gazette_computer_science.docx','LGU_Gazette_computer_science.xlsx'}
    note=folder/'skipped-programs.txt'
    assert note.exists() and 'BS SE' in note.read_text()


def test_pack_with_every_program_skipped_reports_cleanly(tmp_path):
    # Regression: export_gazette_pack crashed with ValueError: too many
    # values to unpack (expected 2) whenever every program in a pack was
    # skipped, because the skipped entries carry 3 values (program, reason,
    # fixes) and the error message unpacked only 2. It must report the
    # blocked program instead of crashing.
    from course_reports import export_gazette_pack
    frame = _df()
    frame.loc[2, 'STUDENT NAME'] = 'MISMATCHED STUDENT'
    source = tmp_path / 'all-skipped.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    folder = tmp_path / 'pack'
    with pytest.raises(ValueError, match='No department gazette could be exported'):
        export_gazette_pack(ds, ds.courses, folder, settings={'gazette_term': 'S26'})
    note = folder / 'skipped-programs.txt'
    assert note.exists() and 'BSCS' in note.read_text()


def test_batch_respects_term_filters_but_keeps_cumulative_history(dataset,tmp_path,monkeypatch):
    import course_ui
    import threading
    from types import SimpleNamespace
    ds,_=dataset
    selected=ds.filtered(term='SU26')
    calls={}
    academic=SimpleNamespace(chosen_courses=lambda ignore_term=False:ds.courses if ignore_term else selected)
    app=SimpleNamespace(academic_dataset=ds,frames={'academic':academic},report_settings={'gazette_term':'SU26'})
    app.run_job=lambda label,work,done:done(work())
    # _ready_program_keys mirrors CourseReportPanel: None means the ready
    # programs were not computed, so batch exports without filtering.
    panel=SimpleNamespace(app=app,stop=threading.Event(),status=SimpleNamespace(configure=lambda **kw:None),_ready_program_keys=lambda:None)
    monkeypatch.setattr(course_ui.filedialog,'askdirectory',lambda **kw:str(tmp_path))
    monkeypatch.setattr(course_ui,'ROUTES',{'register':'Gazette','cgpa_list':'CGPA','term':'Term'})
    monkeypatch.setattr(course_ui,'export_report',lambda key,fmt,ds,rows,path,**kw:calls.setdefault((key,fmt),list(rows)))
    from gazette_options import gazette_rows
    # The gazette takes the full history through the cycle cutoff: the
    # layout needs it for CGPA, and a Spring-cycle gazette must print the
    # Spring courses, not just the Summer rows the filter selected.
    expected_register=gazette_rows(ds,ds.courses,{'gazette_term':'SU26'})
    course_ui.CourseReportPanel.batch(panel)
    for fmt in ('pdf','word','excel'):
        assert calls['term',fmt]==selected
        assert calls['register',fmt]==expected_register
        assert calls['cgpa_list',fmt]==ds.courses


def _two_term_df():
    import pandas as pd
    return pd.DataFrame([
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS102", "Course Three", 3, 90.0, "A", "BSCS", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS103", "Course Four", 3, 85.0, "A-", "BSCS", 11.1, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 87.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM312", "Ecology I", 3, 89.0, "A", "BS BCH", 12.0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])


def test_all_terms_pack_writes_one_pack_per_term(tmp_path):
    from course_reports import export_gazette_pack_all_terms
    source = tmp_path / 'two-term.xlsx'
    _two_term_df().to_excel(source, index=False)
    ds = load_course_workbook(source)
    folder = tmp_path / 'all-terms-pack'
    written = export_gazette_pack_all_terms(ds, ds.courses, folder, settings={})
    names = {p.relative_to(folder).as_posix() for p in written}
    assert names == {
        'F25/LGU_Gazette_computer_science.pdf',
        'F25/LGU_Gazette_computer_science.docx',
        'F25/LGU_Gazette_computer_science.xlsx',
        'S26/LGU_Gazette_computer_science.pdf',
        'S26/LGU_Gazette_computer_science.docx',
        'S26/LGU_Gazette_computer_science.xlsx',
        'S26/LGU_Gazette_allied_health_biology.pdf',
        'S26/LGU_Gazette_allied_health_biology.docx',
        'S26/LGU_Gazette_allied_health_biology.xlsx',
    }
    summary = (folder / 'pack-summary.txt').read_text(encoding='utf-8')
    assert 'F25' in summary and 'S26' in summary
    assert '9 files in total' in summary
    # The F25 pack holds only F25 rows: the S26 Biology student is not in it.
    from openpyxl import load_workbook
    f25 = load_workbook(folder / 'F25' / 'LGU_Gazette_computer_science.xlsx')
    assert 'STUDENT BETA' not in str([c.value for row in f25['Course Results'].iter_rows() for c in row])
    s26 = load_workbook(folder / 'S26' / 'LGU_Gazette_allied_health_biology.xlsx')
    assert 'STUDENT BETA' in str([c.value for row in s26['Course Results'].iter_rows() for c in row])


def test_default_term_is_the_term_with_most_results(tmp_path):
    from gazette_options import current_term
    import pandas as pd
    frame = _two_term_df().iloc[[0, 1, 2]].copy()
    frame.loc[frame['SHORT SEMESTER'] == 'S26', 'SHORT SEMESTER'] = 'S26'
    source = tmp_path / 'default-term.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    # F25 has 2 courses, S26 has 1: the default term is F25 even though S26 is later.
    assert current_term(ds, {}) == 'F25'
    # An explicit term choice still wins.
    assert current_term(ds, {'gazette_term': 'S26'}) == 'S26'


def test_excel_gazette_sheet_matches_official_layout(tmp_path):
    # Regression (21 Nov 2026): the Excel Gazette sheet used a different table
    # than the official gazette (RESULT GAZETTE, PUBLICATION CYCLE line, GP
    # RATE / COURSE GP columns, identity repeated on every row, raw floats
    # like 11.100000000000001). It must match the official sheet: RESULT
    # STATEMENT, SESSION line, the 17 official columns, identity once per
    # student block, totals on the last course row, clean 2-decimal numbers,
    # title only at the top (never repeated on printed pages).
    frame = _two_term_df().iloc[[4, 5]].copy()
    source = tmp_path / 'biology.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'S26'})
    from openpyxl import load_workbook
    ws = load_workbook(path)['Gazette']
    assert ws['A1'].value == 'LAHORE GARRISON UNIVERSITY'
    assert ws['A2'].value == 'RESULT STATEMENT'
    assert 'DEPARTMENT OF' in ws['A3'].value and 'BIOLOGY' in ws['A3'].value
    assert ws['A4'].value == 'SESSION - FALL 2025 - { SEMESTER - SPRING 2026 }'
    assert ws['I5'].value == 'GP' and ws['L5'].value == 'Chr'
    # Non-group labels live in row 5 and merge vertically 5-6; row 6
    # holds only the GP / Chr sub-labels.
    assert [ws.cell(row=5, column=c).value for c in range(1, 18)] == [
        'SR', 'DEGREE ID', 'REGISTRATION NUMBERS', 'STUDENT NAME', 'SEM',
        'COURSE CODE', 'COURSE NAME', 'GRADE', 'GP', None, None, 'Chr',
        None, None, 'SGPA', 'Adj. GPA', 'COMMENTS']
    assert [ws.cell(row=6, column=c).value for c in range(1, 18)] == [
        None, None, None, None, None, None, None, None, 'GP', 'Adj.',
        'T.GP', 'CH', 'Adj.', 'T.CH', None, None, None]
    assert ws.print_title_rows is None
    # Identity once, on the first row of the block only.
    assert ws['A7'].value == 1 and ws['D7'].value == 'STUDENT BETA'
    assert ws['A8'].value is None and ws['D8'].value is None
    # Totals sit at the top of the merged summary block (row 7):
    # 12.00 + 12.00 = 24.00, 6 credits.
    assert ws['I7'].value == '24.00' and ws['K7'].value == '24.00'
    assert ws['L7'].value == '6' and ws['N7'].value == '6'
    assert ws['O7'].value == '4.00' and ws['P7'].value == '4.00'
    # Block padded to 10 rows with '-' rows; NOTE ( If any ) in the SR column.
    assert ws['F9'].value == '-' and ws['G9'].value == '-' and ws['H9'].value == '-'
    assert ws['A11'].value == 'NOTE ( If any )'
    assert ws.max_row >= 16
    # No float storage noise anywhere on the Gazette sheet.
    for row in ws.iter_rows():
        for cell in row:
            assert not (isinstance(cell.value, str) and '000000000' in cell.value), cell.coordinate


def test_excel_gazette_sheet_prints_standing_in_comments(tmp_path):
    # Regression (22 Nov 2026): the COMMENTS column was empty, but the
    # official gazettes print G.A.ST (good standing) or PROBATION there.
    # The cutoff is the probation_cgpa setting, else 2.00 CGPA.
    import pandas as pd
    frame = pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 87.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM312", "Ecology I", 3, 89.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/002", "STUDENT GAMMA", 0, "BCHM311", "Biochemistry I", 3, 20.0, "F", "BS BCH", 0.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/002", "STUDENT GAMMA", 0, "BCHM312", "Ecology I", 3, 15.0, "F", "BS BCH", 0.0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'standing.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'S26'})
    from openpyxl import load_workbook
    ws = load_workbook(path)['Gazette']
    # BETA (CGPA 4.00) gets G.A.ST on the totals row; GAMMA (CGPA 0.00,
    # all F) gets PROBATION on its totals row, one 10-row block later.
    assert ws['Q7'].value == 'G.A.ST'
    assert ws['Q17'].value == 'PROBATION'
    # The READ FIRST sheet states the cutoff that was used.
    first = load_workbook(path)['READ FIRST']
    standing = [row for row in first.iter_rows(values_only=True) if row and row[0] == 'Academic standing']
    assert standing and '2.00' in str(standing[0][1]) and 'G.A.ST' in str(standing[0][1])


def test_pdf_gazette_prints_standing_in_comments(tmp_path):
    # The PDF gazette prints the same standing in COMMENTS as the Excel.
    frame = _two_term_df().iloc[[4, 5]].copy()
    source = tmp_path / 'biology.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.pdf'
    export_report('register', 'pdf', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'S26'})
    import fitz
    text = ''.join(page.get_text() for page in fitz.open(path))
    assert 'G.A.ST' in text


def test_word_register_prints_standing_per_term(tmp_path):
    # Regression (22 Nov 2026): the Word register showed SGPA/CGPA per
    # term but no standing, while the PDF and Excel gazettes print
    # G.A.ST / PROBATION. The Word per-term line now carries it too.
    import pandas as pd
    frame = pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 87.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM312", "Ecology I", 3, 89.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/002", "STUDENT GAMMA", 0, "BCHM311", "Biochemistry I", 3, 20.0, "F", "BS BCH", 0.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/002", "STUDENT GAMMA", 0, "BCHM312", "Ecology I", 3, 15.0, "F", "BS BCH", 0.0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'standing.docx.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'register.docx'
    export_report('register', 'word', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'S26'})
    from docx import Document
    text = '\n'.join(p.text for p in Document(path).paragraphs)
    assert 'S26: SGPA 4.00 | CGPA 4.00 | G.A.ST' in text
    assert 'S26: SGPA 0.00 | CGPA 0.00 | PROBATION' in text


def test_app_title_uses_the_current_version():
    # Regression (22 Nov 2026): the window title was hardcoded to v4.5.2
    # even after v4.6.2 shipped, so the user could not tell which version
    # he was running. The title now comes from one APP_VERSION constant.
    from app import APP_VERSION
    assert APP_VERSION == 'v4.6.8'


def test_float_noise_does_not_flip_half_up_rounding(tmp_path):
    # Regression v4.6.4 (B2): source grade points arrive as binary floats,
    # so a value the university meant as 6.90 can read 6.8999999999999995.
    # That noise flipped half-up rounding when the true ratio sat exactly
    # on a .xx5 boundary: 32 SGPA/CGPA values in F20.xlsx printed 0.01
    # low. Here the clean total is 16.35 over 6 credits (2.725), which
    # must display 2.73, not 2.72.
    import pandas as pd
    frame = pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 69.0, "C+", "BS BCH", 6.8999999999999995, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM312", "Ecology I", 3, 94.5, "B", "BS BCH", 9.45, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'float-noise.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    from calculation_policy import AcademicIndex, display_ratio
    index = AcademicIndex(ds.courses)
    summary = index.summary('Fa-2025/BSBCH/001', 'BS BCH', 'S26')
    assert display_ratio(summary.exact_ratio) == '2.73'


def test_excel_tgp_tch_are_cumulative_not_semester(tmp_path):
    # Regression v4.6.4 (B1): the Excel T.GP / T.CH columns showed the
    # semester totals, not the cumulative totals. ALPHA's S26 section must
    # show GP 23.10 (S26 only) against T.GP 42.90 (F25 + S26), and CH 6
    # against T.CH 13. The PDF already showed cumulative; the Excel must
    # match it.
    frame = _two_term_df().iloc[[0, 1, 2, 3]].copy()
    source = tmp_path / 'two-term.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='computer-science', settings={'gazette_term': 'S26'})
    from openpyxl import load_workbook
    ws = load_workbook(path)['Gazette']
    # ALPHA's S26 section: 2 courses, totals at the top of the merged
    # summary block (row 7).
    assert ws['I7'].value == '23.10'
    assert ws['K7'].value == '42.90'
    assert ws['L7'].value == '6'
    assert ws['N7'].value == '13'
    assert ws['O7'].value == '3.85'
    assert ws['P7'].value == '3.30'


def test_excel_gazette_shows_only_the_publication_term(tmp_path):
    # Regression v4.6.4 (B5/B6): the Excel Gazette sheet printed the
    # student's full history (every term), and a Fall gazette also printed
    # the Summer sections that belong to the Spring gazette. It must show
    # only the publication term window, like the PDF: S26 here, never F25.
    frame = _two_term_df().iloc[[0, 1, 2, 3]].copy()
    source = tmp_path / 'two-term.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='computer-science', settings={'gazette_term': 'S26'})
    from openpyxl import load_workbook
    import re
    ws = load_workbook(path)['Gazette']
    # The SEM column also carries the header label and the fee-section
    # header, so only count values that look like a term code (S26, SU26,
    # F25). Those must be only S26: no F25 history, no other term.
    terms = {ws.cell(row=r, column=5).value for r in range(7, ws.max_row + 1)}
    terms = {t for t in terms if t and isinstance(t, str) and re.fullmatch(r'(S|SU|F)\d+', t.strip())}
    assert terms == {'S26'}


def test_ip_course_does_not_block_gpa(tmp_path):
    # Regression v4.6.4 (B4): one in-progress course used to poison the
    # student's whole SGPA and every later CGPA (SARA FAIZA F25, SAWERA
    # F24). A non-GPA grade (IP, W, I, P, R, blank) is now skipped, so the
    # GPA is computed from the graded courses.
    import pandas as pd
    frame = pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 87.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM312", "Ecology I", 3, 89.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "MATH383", "Mathematics III", "", 0, "IP", "BS BCH", 0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'ip.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    from calculation_policy import AcademicIndex, display_ratio
    index = AcademicIndex(ds.courses)
    summary = index.summary('Fa-2025/BSBCH/001', 'BS BCH', 'S26')
    assert display_ratio(summary.exact_ratio) == '4.00'


def test_adj_gp_shows_replaced_prior_attempt(tmp_path):
    # The official gazettes leave Adj. blank on a first-time course and
    # fill it only on a repeated course, with the replaced prior
    # attempt's GP and CH. This supersedes the v4.6.4 test, which
    # asserted the old wrong rule (every row showing its own points).
    # The same course twice in S26 (A then B): the A row is first-time
    # so Adj. is blank, and the B row is a repeat so Adj. shows the
    # replaced A attempt, 12.00 with 3 CH, not the B row's own 9.00.
    import pandas as pd
    frame = pd.DataFrame([
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 87.0, "A", "BS BCH", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSBCH/001", "STUDENT BETA", 0, "BCHM311", "Biochemistry I", 3, 75.0, "B", "BS BCH", 9.0, "S26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'repeat.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    from calculation_policy import AcademicIndex, calculation_attempts
    from gazette_layout import _adj_values
    calculations = AcademicIndex(calculation_attempts(ds))
    rows = sorted(ds.courses, key=lambda c: c.source_excel_row)
    assert _adj_values(rows[0], 'regular', calculations, rows[0].registration, rows[0].degree) == ('', '')
    assert _adj_values(rows[1], 'regular', calculations, rows[1].registration, rows[1].degree) == ('12.00', '3')


def test_adj_gp_blank_for_first_time_course(tmp_path):
    # Adj. blank rule, cross-term repeat (F20.xlsx case: MUHAMMAD AYHAN
    # FAROOQ failed BTCH311 at 0.00 in F20 and retook it in F22 for
    # 6.00). The F20 row is first-time so Adj. is blank, the F22 row
    # is a repeat so Adj. shows the replaced 0.00 with 3 CH, and a
    # first-time F22 course is blank too.
    import pandas as pd
    frame = pd.DataFrame([
        ["FALL 2020", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 20.0, "F", "BS BT", 0.0, "F20"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 60.0, "C", "BS BT", 6.0, "F22"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH341", "Biotechnology III", 3, 75.0, "B", "BS BT", 9.0, "F22"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'adj-blank.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    from calculation_policy import AcademicIndex, calculation_attempts
    from gazette_layout import _adj_values
    calculations = AcademicIndex(calculation_attempts(ds))
    rows = sorted(ds.courses, key=lambda c: c.source_excel_row)
    assert _adj_values(rows[0], 'regular', calculations, rows[0].registration, rows[0].degree) == ('', '')
    assert _adj_values(rows[1], 'regular', calculations, rows[1].registration, rows[1].degree) == ('0.00', '3')
    assert _adj_values(rows[2], 'regular', calculations, rows[2].registration, rows[2].degree) == ('', '')


def test_default_term_counts_spring_cycle_together(tmp_path):
    # Regression v4.6.4 (R3): the default term counted raw term rows, so a
    # Spring cycle could lose to a smaller Fall term once its Summer rows
    # were split off. S26 has 2 rows and SU26 has 3 (cycle 5), F25 has 4,
    # so the default must be S26, not F25.
    import pandas as pd
    frame = pd.DataFrame([
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS102", "Course Three", 3, 90.0, "A", "BSCS", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS103", "Course Four", 3, 85.0, "A-", "BSCS", 11.1, "S26"],
        ["SUMMER 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS104", "Course Five", 3, 80.0, "B+", "BSCS", 9.9, "SU26"],
        ["SUMMER 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS105", "Course Six", 3, 80.0, "B+", "BSCS", 9.9, "SU26"],
        ["SUMMER 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS106", "Course Seven", 3, 80.0, "B+", "BSCS", 9.9, "SU26"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'cycle.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    from gazette_options import current_term
    assert current_term(ds, {}) == 'S26'

    from pathlib import Path
    assert 'v4.5.2' not in Path('app.py').read_text(encoding='utf-8')


@needs_display
def test_dashboard_metric_click_stays_on_dashboard(dataset, monkeypatch):
    # Regression: the Dashboard metric cards used to carry a hidden
    # <Button-1> binding that jumped to Academic Records, so clicking
    # anywhere on the Dashboard looked like a navigation error.
    import os, subprocess, sys
    if os.environ.get('LGU_GUI_ISOLATED') != '1':
        env = dict(os.environ, LGU_GUI_ISOLATED='1')
        result = subprocess.run([sys.executable, '-m', 'pytest',
                                 str(Path(__file__).resolve()) + '::test_dashboard_metric_click_stays_on_dashboard',
                                 '-q'], env=env, capture_output=True, text=True, timeout=120)
        assert result.returncode == 0, result.stdout + result.stderr
        return
    from app import App
    monkeypatch.setattr(App, '_restore_autosave', lambda self: None)
    app = App()
    ds, source = dataset
    try:
        app.set_academic_data(ds, str(source))
        app.show_frame('dashboard')
        app.update()
        panel = app.frames['dashboard']
        assert panel.winfo_ismapped()
        # No click binding may remain on the metric cards.
        assert panel.metric_canvas.tag_bind('metric', '<Button-1>') == ''
        # Click on the first metric card: the Dashboard must stay shown.
        canvas = panel.metric_canvas
        canvas.update_idletasks()
        canvas.event_generate('<Button-1>', x=60, y=60)
        app.update()
        assert panel.winfo_ismapped()
        assert not app.frames['academic'].winfo_ismapped()
    finally:
        for job in app.tk.call('after', 'info'):
            app.after_cancel(job)
        app.destroy()
def test_gazette_contents_counts_only_active_term_students(tmp_path):
    # Regression (22 Nov 2026): Gazette Contents counted every student in
    # the cohort (len(reg_groups) across all terms), so F24 Maths reported
    # 15 students when only 2 had results in F24. It must count only the
    # students printed in the active publication term.
    import pandas as pd
    frame = pd.DataFrame([
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS102", "Course Three", 3, 90.0, "A", "BSCS", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/001", "STUDENT ALPHA", 0, "CS103", "Course Four", 3, 85.0, "A-", "BSCS", 11.1, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS102", "Course Three", 3, 90.0, "A", "BSCS", 12.0, "S26"],
        ["SPRING 2026", "Fa-2025/BSCS/002", "STUDENT BETA", 0, "CS103", "Course Four", 3, 85.0, "A-", "BSCS", 11.1, "S26"],
        ["FALL 2025", "Fa-2025/BSCS/003", "STUDENT GAMMA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/003", "STUDENT GAMMA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/004", "STUDENT DELTA", 0, "CS100", "Course One", 3, 74.5, "B", "BSCS", 9.0, "F25"],
        ["FALL 2025", "Fa-2025/BSCS/004", "STUDENT DELTA", 0, "CS101", "Course Two", 4, 55.0, "B-", "BSCS", 10.8, "F25"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'contents-count.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='computer-science', settings={'gazette_term': 'S26'})
    from openpyxl import load_workbook
    contents = load_workbook(path)['Gazette Contents']
    students = [row[6] for row in contents.iter_rows(min_row=4, values_only=True) if row[0] is not None]
    assert students == [2]


def test_excel_gazette_adj_matches_pdf_on_repeated_courses(tmp_path):
    # Regression (22 Nov 2026, v4.6.7): the Excel gazette left Adj. GP
    # and Adj. CH blank on repeated courses while the PDF showed the
    # replaced prior attempt. AYHAN failed BTCH311 at 0.00 in F20 and
    # retook it in F22 for 6.00, so the F22 Excel row must show Adj.
    # 0.00 and 3, and a first-time F22 course must stay blank.
    import pandas as pd
    frame = pd.DataFrame([
        ["FALL 2020", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 20.0, "F", "BS BT", 0.0, "F20"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 60.0, "C", "BS BT", 6.0, "F22"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH341", "Biotechnology III", 3, 75.0, "B", "BS BT", 9.0, "F22"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'excel-adj.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.xlsx'
    export_report('register', 'excel', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'F22'})
    from openpyxl import load_workbook
    ws = load_workbook(path)['Gazette']
    # Row 7 is the F22 BTCH311 repeat, row 8 the first-time BTCH341.
    # Adj. GP is col J, Adj. CH is col M.
    assert ws['J7'].value == '0.00' and ws['M7'].value == '3'
    assert ws['J8'].value in (None, '') and ws['M8'].value in (None, '')
    # Regression (22 Nov 2026, v4.6.8): the Adj. GP and Adj. CH cells
    # on repeated courses must be center-aligned like the other
    # numeric columns.
    assert ws['J7'].alignment.horizontal == 'center'
    assert ws['M7'].alignment.horizontal == 'center'


def test_tch_header_clears_column_right_border(tmp_path):
    # Regression (22 Nov 2026, v4.6.7): T.CH touched the right vertical
    # border of its narrow column on 17-column pages. It is now shrunk
    # by 0.5 pt and centred in its column, so it must end with at least
    # 1 pt of clear space before the column's right border.
    import pandas as pd
    frame = pd.DataFrame([
        ["FALL 2020", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 20.0, "F", "BS BT", 0.0, "F20"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH311", "Biotechnology I", 3, 60.0, "C", "BS BT", 6.0, "F22"],
        ["FALL 2022", "Fa-2020/BSBT/002", "STUDENT AYHAN", 0, "BTCH341", "Biotechnology III", 3, 75.0, "B", "BS BT", 9.0, "F22"],
    ], columns=["SEMESTER", "REGISTRATION NUMBERS", "STUDENT NAME", "AWARD ID", "COURSE CODE", "COURSE NAME", "CRH", "Total Marks", "Grade", "DEGREE TITLE", "GPs", "SHORT SEMESTER"])
    source = tmp_path / 'tch.xlsx'
    frame.to_excel(source, index=False)
    ds = load_course_workbook(source)
    path = tmp_path / 'gazette.pdf'
    export_report('register', 'pdf', ds, ds.courses, str(path),
                  template_key='allied-health-biology', settings={'gazette_term': 'F22'})
    import fitz
    doc = fitz.open(path)
    tch_right = None
    for page in doc:
        for block in page.get_text('dict')['blocks']:
            for line in block.get('lines', []):
                for span in line['spans']:
                    if span['text'].strip() == 'T.CH':
                        tch_right = span['bbox'][2]
                        break
                if tch_right is not None:
                    break
            if tch_right is not None:
                break
        if tch_right is not None:
            break
    assert tch_right is not None, 'T.CH header not found in the PDF'
    from gazette_layout import _profile
    profile = _profile('allied-health-biology', 'FALL 2020', 'BS BT', require_cgpa=True)
    edges = profile['edges']
    origin = next(item['origin'][0] for item in profile['header']
                  if item['text'].strip() == 'T.CH')
    right = next(edges[i + 1] for i in range(len(edges) - 1)
                 if edges[i] <= origin < edges[i + 1])
    assert tch_right <= right - 1.0, f'T.CH ends {tch_right:.2f}, column right border {right:.2f}'
