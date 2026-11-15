from pathlib import Path
from dataclasses import replace
from types import SimpleNamespace
import os
import threading
import pandas as pd
import pytest
needs_display=pytest.mark.skipif(not os.environ.get('DISPLAY'),reason='Needs a display; run under xvfb-run or on Windows')
from academic_engine import load_course_workbook
from course_reports import export_report,capability
from gazette_options import current_term
from calculation_policy import parse_grade_scale
from test_academic_engine import _df


def test_failed_export_preserves_existing_file(tmp_path,monkeypatch):
    import course_reports
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    target=tmp_path/'result.pdf';target.write_bytes(b'PREVIOUS COMPLETE OUTPUT')
    def fail(*args,**kwargs):
        Path(args[4]).write_bytes(b'PARTIAL')
        raise RuntimeError('Simulated writer failure')
    monkeypatch.setattr(course_reports,'_write_report',fail)
    with pytest.raises(RuntimeError):export_report('term','pdf',ds,ds.courses,str(target))
    assert target.read_bytes()==b'PREVIOUS COMPLETE OUTPUT'
    assert not list(tmp_path.glob('.lgu-export-*'))


def test_case_changed_source_name_and_wrong_extension_are_rejected(tmp_path):
    source=tmp_path/'Source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    with pytest.raises(ValueError,match='original'):export_report('term','excel',ds,ds.courses,str(tmp_path/'SOURCE.XLSX'))
    with pytest.raises(ValueError,match='extension'):export_report('term','pdf',ds,ds.courses,str(tmp_path/'wrong.xlsx'))


def test_cgpa_list_checks_all_terms_it_will_print(tmp_path):
    frame=_df().iloc[:3].copy();frame['CRH']=frame['CRH'].astype(object);frame.loc[2,'CRH']='bad'
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    earlier=[c for c in ds.courses if c.term=='S26']
    assert capability('register',ds,earlier)[0]
    assert not capability('cgpa_list',ds,earlier)[0]


def test_current_gazette_not_blocked_by_later_term_error(tmp_path):
    frame=_df().iloc[:3].copy();frame['CRH']=frame['CRH'].astype(object);frame.loc[2,'CRH']='bad'
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    export_report('register','excel',ds,ds.courses,str(tmp_path/'current.xlsx'),settings={'gazette_term':'S26'})
    from openpyxl import load_workbook
    book=load_workbook(tmp_path/'current.xlsx');assert book['Course Results'].max_row==3;book.close()


def test_term_selection_does_not_silently_change(tmp_path):
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    ds.courses.append(replace(ds.courses[0],term='Unspecified term'))
    # v4.5.2: Summer folds into the Spring publication cycle, so the current term is S26, not SU26.
    assert current_term(ds,{})=='S26'
    with pytest.raises(ValueError,match='Settings'):current_term(ds,{'gazette_term':'F20'})


def test_busy_batch_cannot_clear_running_cancellation(monkeypatch):
    import course_ui
    stop=threading.Event();stop.set()
    panel=SimpleNamespace(app=SimpleNamespace(job_busy=True),stop=stop)
    monkeypatch.setattr(course_ui.messagebox,'showinfo',lambda *a,**kw:None)
    monkeypatch.setattr(course_ui.filedialog,'askdirectory',lambda **kw:pytest.fail('Busy batch opened a dialog'))
    course_ui.CourseReportPanel.batch(panel)
    assert stop.is_set()


def test_extreme_grade_scale_input_is_rejected():
    with pytest.raises(ValueError):parse_grade_scale('A=1e99999999','TEST')
    with pytest.raises(ValueError):parse_grade_scale('A=0.1234567890123456789012345','TEST')


def test_export_snapshot_keeps_scale_and_fees_stable(tmp_path):
    from course_ui import dataset_snapshot
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    ds.grade_scale={'B':'3'};frozen=dataset_snapshot(ds)
    ds.grade_scale['B']='4'
    assert frozen.grade_scale['B']=='3' and frozen.courses is ds.courses


@needs_display
def test_settings_paste_and_stale_term_review_routes(tmp_path,monkeypatch):
    from app import App
    from course_reports import ROUTES
    monkeypatch.setattr(App,'_restore_autosave',lambda self:None)
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    app=App()
    try:
        app.set_academic_data(ds,str(source));app.show_frame('settings');app.update()
        settings=app.frames['settings'].course_panel
        entry=settings.cover_entries['cover_department']
        entry.delete(0,'end');entry.insert(0,'PASTED DEPARTMENT')
        assert app.report_settings['cover_department']=='PASTED DEPARTMENT'
        settings.refresh();assert entry.get()=='PASTED DEPARTMENT'
        app.show_frame('dashboard');app.update()
        canvas=app.frames['dashboard'].metric_canvas
        assert 'Errors requiring review' in [canvas.itemcget(i,'text') for i in canvas.find_all() if canvas.type(i)=='text']
        app.report_settings['gazette_term']='F20'
        app.show_frame('reports');app.update();panel=app.frames['reports'].course_panel
        assert panel.format_buttons[0].cget('state')=='disabled'
        panel.route_menu.set(ROUTES['validation']);panel.refresh()
        assert panel.format_buttons[0].cget('state')=='normal'
    finally:
        for job in app.tk.call('after','info'):app.after_cancel(job)
        app.destroy()
