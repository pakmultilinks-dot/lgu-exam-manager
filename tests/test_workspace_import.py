from dataclasses import replace
from pathlib import Path
import pytest
from workspace_state import import_settings,persist_import
from gazette_options import resolve_template_groups
from academic_engine import ResultDataset
from test_cross_program_matrix import attempt


def test_new_workbook_settings_drop_old_layout_term_and_cover_scope():
    ds=ResultDataset([attempt(term='F25'),attempt(term='S26')],[],[],'new.xlsx')
    settings=import_settings(ds,{'gazette_template':'computer-science','gazette_term':'F25','cover_department':'OLD','cover_period':'OLD','cover_title':'RESULT'})
    assert settings['gazette_template']=='auto' and settings['gazette_term']=='S26'
    assert 'cover_department' not in settings and 'cover_period' not in settings
    assert settings['cover_title']=='RESULT'


def test_replacing_workspace_keeps_previous_data_and_reloads_new_data(tmp_path):
    import session_store
    old=ResultDataset([attempt()],[],[],'old.xlsx')
    new=ResultDataset([attempt(degree='BBA'),attempt(registration='TEST/002')],[],[],'new.xlsx')
    target=tmp_path/'workspace.lguexam'
    persist_import(target,None,old,import_settings(old,{}))
    persist_import(target,None,new,import_settings(new,{}))
    assert session_store.restore(target)[1].courses==new.courses
    assert session_store.restore(tmp_path/'workspace.previous.lguexam')[1].courses==old.courses


def test_wrong_known_department_layout_cannot_be_published(tmp_path):
    from academic_reports import generate_draft
    rows=[attempt(),attempt(degree='BBA',registration='TEST/002')]
    ds=ResultDataset(rows,[],[],'mixed.xlsx')
    with pytest.raises(ValueError,match='does not match'):
        resolve_template_groups('computer-science',rows)
    output=tmp_path/'output.pdf';output.write_bytes(b'PRIOR COMPLETE OUTPUT')
    with pytest.raises(ValueError,match='does not match'):
        generate_draft('pdf',ds,rows,output,'computer-science')
    assert output.read_bytes()==b'PRIOR COMPLETE OUTPUT'
    assert len(resolve_template_groups('auto',rows))==2


def test_calculation_review_points_to_unresolved_historical_source_rows():
    from calculation_review import calculation_review_issues
    # v4.6.4 (B4): a non-GPA grade (IP, W, I, P, R, blank) in history is now
    # skipped, so it no longer blocks the CGPA. The review still points to
    # a genuinely unresolved historical row: here a grade with no rate and
    # no points, so the CGPA cannot be computed.
    past=attempt(term='F25',grade='X',quality_points='',source_excel_row=19)
    current=attempt(term='S26',course_code='NEW101',source_excel_row=24)
    ds=ResultDataset([past,current],[],[],'source.xlsx')
    issues=calculation_review_issues(ds,[current])
    assert len(issues)==1 and issues[0].code=='CGPA_CALCULATION_REVIEW'
    assert issues[0].excel_row==19 and 'X' in issues[0].description
    ds.grade_scale={'A':'3.50'}
    assert 'X' in calculation_review_issues(ds,[current])[0].description



def test_explicit_layout_accepts_a_program_present_in_multiple_reference_catalogs():
    rows=[attempt(degree='BS IT')]
    assert resolve_template_groups('auto',rows)[0][0]=='information-technology'
    assert resolve_template_groups('information-technology',rows)[0][0]=='information-technology'
    assert resolve_template_groups('software-engineering',rows)[0][0]=='software-engineering'
    with pytest.raises(ValueError,match='does not match'):
        resolve_template_groups('computer-science',rows)
