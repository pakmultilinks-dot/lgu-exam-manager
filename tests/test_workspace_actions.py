import os
import time
import pytest
needs_display=pytest.mark.skipif(not os.environ.get('DISPLAY'),reason='Needs a display; run under xvfb-run or on Windows')


@needs_display
def test_explicit_open_and_clear_buttons_preserve_saved_file(tmp_path,monkeypatch):
    from app import App
    from academic_engine import ResultDataset
    from test_cross_program_matrix import attempt
    import session_store
    saved=ResultDataset([attempt()],[],[],'unseen-source.xlsx')
    path=tmp_path/'operator-workspace.lguexam'
    session_store.save(path,'course',None,None,{}, {},saved)
    original=path.read_bytes()
    app=App()
    try:
        app.withdraw();app.update()
        assert app.academic_dataset is None
        panel=app.frames['import'];panel.path='obsolete.xlsx'
        app.report_settings['cover_department']='OLD COVER'
        app.report_settings['cover_period']='OLD PERIOD'
        panel.open_saved_button.invoke()
        deadline=time.monotonic()+30
        while app.job_busy and time.monotonic()<deadline:app.update();time.sleep(.01)
        assert not app.job_busy and app.academic_dataset.courses==saved.courses
        assert panel.path is None
        assert 'cover_department' not in app.report_settings and 'cover_period' not in app.report_settings
        app.compare_df='old comparison';app.compare_mapping={'old':'old'}
        panel.clear_button.invoke();app.update()
        assert app.academic_dataset is None and app.raw_df is None and app.df is None
        assert app.compare_df is None and app.compare_mapping is None
        assert panel.path is None and panel.file_lbl.cget('text')=='No file selected yet'
        assert 'No file selected' in app.header_status.cget('text')
        assert path.read_bytes()==original
        panel.path='active.xlsx';app.job_busy=True
        notices=[]
        import app as module
        monkeypatch.setattr(module.messagebox,'showinfo',lambda *a,**k:notices.append(a))
        panel.clear_button.invoke()
        assert panel.path=='active.xlsx' and notices
    finally:
        for job in app.tk.call('after','info'):app.tk.call('after','cancel',job)
        app.destroy()


def test_source_protection_allows_same_basename_in_another_directory(tmp_path):
    from test_academic_engine import _df
    from academic_engine import load_course_workbook
    from export_io import protect_source
    import session_store
    source=tmp_path/'result.xlsx';_df().to_excel(source,index=False)
    ds=load_course_workbook(source);original=source.read_bytes()
    with pytest.raises(ValueError,match='original workbook'):protect_source(ds,source)
    output=tmp_path/'exports'/'result.xlsx';output.parent.mkdir()
    protect_source(ds,output)
    from course_reports import export_report
    from openpyxl import load_workbook
    with pytest.raises(ValueError,match='original workbook'):
        export_report('register','excel',ds,ds.courses,str(source),template_key='auto')
    export_report('register','excel',ds,ds.courses,str(output),template_key='auto')
    with_workbook=load_workbook(output,read_only=True)
    assert with_workbook['Course Results'].max_row==5
    with_workbook.close()
    snapshot=tmp_path/'snapshot.lguexam';session_store.save(snapshot,'course',source,None,{}, {},ds)
    restored=session_store.restore(snapshot)[1]
    assert restored.source_paths==ds.source_paths
    with pytest.raises(ValueError,match='original workbook'):protect_source(restored,source)
    protect_source(restored,output)
    assert source.read_bytes()==original


def test_gazette_frontmatter_has_no_added_draft_label(tmp_path):
    from test_academic_engine import _df
    from academic_engine import load_course_workbook
    from academic_reports import generate_draft
    import pymupdf
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False)
    dataset=load_course_workbook(source);target=tmp_path/'gazette.pdf'
    generate_draft('pdf',dataset,dataset.courses,target)
    with pymupdf.open(target) as pdf:
        assert 'DRAFT' not in pdf[0].get_text()
        assert 'DRAFT' not in pdf[1].get_text()
        assert 'TABLE OF CONTENTS' in pdf[1].get_text()
        assert 'NOTE ( If any )' in pdf[2].get_text()
        for word in pdf[2].get_text('words'):
            assert 0<=word[0]<word[2]<=pdf[2].rect.width
            assert 0<=word[1]<word[3]<=pdf[2].rect.height
