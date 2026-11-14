"""Regression coverage for canonical routes, provenance and actual Tk callbacks."""
from pathlib import Path
import sys,time
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import pytest
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
        assert all(c.term=='SU26' for c in panel.rows())
        app.frames['academic'].terms.set('All Terms')
        app.show_frame('settings');app.update()
        settings=app.frames['settings']
        assert settings.course_panel.winfo_ismapped() and not settings.legacy_panel.winfo_ismapped()
        assert 'S26' in settings.course_panel.detected.cget('text')
        settings.course_panel.select('Automatic from imported data')
        app.show_frame('reports');app.update()
        assert panel.template_menu.get()=='Automatic from imported data'
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
    panel=SimpleNamespace(app=app,stop=threading.Event(),status=SimpleNamespace(configure=lambda **kw:None))
    monkeypatch.setattr(course_ui.filedialog,'askdirectory',lambda **kw:str(tmp_path))
    monkeypatch.setattr(course_ui,'ROUTES',{'register':'Gazette','cgpa_list':'CGPA','term':'Term'})
    monkeypatch.setattr(course_ui,'export_report',lambda key,fmt,ds,rows,path,**kw:calls.setdefault((key,fmt),list(rows)))
    course_ui.CourseReportPanel.batch(panel)
    for fmt in ('pdf','word','excel'):
        assert calls['term',fmt]==selected
        assert calls['register',fmt]==selected
        assert calls['cgpa_list',fmt]==ds.courses
