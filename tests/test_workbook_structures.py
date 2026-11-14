"""Unseen structures; fixtures are never application defaults."""
import json
import pandas as pd
import pytest
from academic_engine import SheetLayout, detect_layouts, load_course_workbook
from workbook_schema import inspect_workbook, needs_review, save_profile, load_profiles
from calculation_policy import calculation_attempts


def table():
    return pd.DataFrame([['ENR-893','NA','ZZQ9','Orbital Methods',5,'A','F29',3.1]],columns=['Registration','Student Name','Course Code','Course Name','Credit Hours','Grade','Term','Grade Points Per Credit'])


@pytest.mark.parametrize('suffix',['.xlsx','.xlsm','.xls'])
def test_excel_containers_and_literal_na(tmp_path,suffix):
    path=tmp_path/('unseen'+suffix);frame=table()
    if suffix=='.xls':
        import xlwt
        book=xlwt.Workbook();sheet=book.add_sheet('Unseen')
        for i,row in enumerate([frame.columns.tolist()]+frame.values.tolist()):
            for j,value in enumerate(row):sheet.write(i,j,value)
        book.save(str(path))
    else:frame.to_excel(path,index=False,engine='openpyxl')
    ds=load_course_workbook(path)
    assert len(ds.courses)==1 and ds.courses[0].name=='NA'
    assert calculation_attempts(ds)[0].quality_points=='15.5'


def test_wide_mapping_keeps_original_rows_and_skips_empty_course(tmp_path):
    path=tmp_path/'unfamiliar.xlsx'
    pd.DataFrame([['R-91','Person Z',2,'B',4,'A'],['R-92','Person Q',3,'F','','']],columns=['Identity','Person','Units x','Outcome x','Units y','Outcome y']).to_excel(path,index=False)
    layout=SheetLayout('Sheet1',0,{'registration':'Identity','name':'Person'},0,path.name,defaults={'term':'S29','degree':'New Degree'},course_blocks=[{'columns':{'credits':'Units x','grade':'Outcome x'},'defaults':{'course_code':'Q91','course_name':'First Unseen'}},{'columns':{'credits':'Units y','grade':'Outcome y'},'defaults':{'course_code':'Q92','course_name':'Second Unseen'}}],confirmed=True)
    ds=load_course_workbook(path,layouts_override=[layout])
    assert len(ds.courses)==3 and len(ds.source_rows)==2
    assert [c.source_excel_row for c in ds.courses]==[2,2,3]
    assert [c.course_code for c in ds.courses]==['Q91','Q92','Q91']
    assert ds.courses[2].grade=='F'


def test_mapping_profiles_never_reuse_course_or_metadata_values(tmp_path):
    path=tmp_path/'profiles.json'
    layout=SheetLayout('Private Sheet',0,{'registration':'ID','name':'Person'},1,'Private.xlsx',defaults={'term':'F29','degree':'Private Degree'},course_blocks=[{'columns':{'grade':'Outcome'},'defaults':{'course_code':'SECRET91','course_name':'Private Title','credits':'7'}}])
    save_profile(path,['ID','Person','Outcome'],layout)
    text=path.read_text();profile=next(iter(load_profiles(path).values()))
    assert profile['requires_review'] and profile['defaults']=={}
    assert profile['course_blocks'][0]['defaults']=={}
    assert all(secret not in text for secret in ['F29','Private','SECRET91'])


def test_unknown_extra_sheet_requires_review_and_ambiguous_gp(tmp_path):
    path=tmp_path/'two.xlsx'
    with pd.ExcelWriter(path) as writer:
        table().drop(columns='Grade Points Per Credit').assign(GP=3).to_excel(writer,index=False,sheet_name='Results')
        pd.DataFrame([['a','b','c'],['d','e','f']],columns=['X','Y','Z']).to_excel(writer,index=False,sheet_name='Another table')
    layouts=detect_layouts(path);inspection=inspect_workbook(path,layouts)
    assert needs_review(inspection,layouts)
    assert len(inspection)==2


def test_two_row_headings_and_arbitrary_header_depth(tmp_path):
    path=tmp_path/'deep.xlsx'
    rows=[['Information']]*73+[['Registration','Student','Course','Course','Credit','Result','Exam'],['Number','Name','Code','Name','Hours','Grade','Term'],['ID-78','Person R','K92','Unseen Topic',3,'C+','SU29']]
    pd.DataFrame(rows).to_excel(path,index=False,header=False)
    ds=load_course_workbook(path)
    assert len(ds.courses)==1 and ds.courses[0].source_excel_row==76
    assert ds.courses[0].registration=='ID-78'


def test_remembered_column_mapping_accepts_new_values_only(tmp_path):
    import workbook_schema
    first=tmp_path/'first.xlsx';second=tmp_path/'unrelated.xlsx'
    headers=['Identifier','Person','Catalog','Topic','Units','Outcome','Period']
    pd.DataFrame([['R-91','First','T-1','Topic One',2,'A','F28']],columns=headers).to_excel(first,index=False)
    mapping=dict(zip(['registration','name','course_code','course_name','credits','grade','term'],headers))
    layout=SheetLayout('Sheet1',0,mapping,1,first.name,confirmed=True)
    save_profile(workbook_schema.default_profile_path(),headers,layout)
    pd.DataFrame([['NEW-62','Second','Z-99','Different Topic',7,'F','S29']],columns=headers).to_excel(second,index=False,sheet_name='Another sheet')
    ds=load_course_workbook(second)
    assert len(ds.courses)==1
    course=ds.courses[0]
    assert (course.registration,course.credits,course.grade,course.term)==('NEW-62','7','F','S29')
    assert course.course_name=='Different Topic'


def test_weighted_points_override_explicit_rate_and_scale(tmp_path):
    path=tmp_path/'source-priority.xlsx'
    table().assign(**{'Quality Points':9}).to_excel(path,index=False)
    ds=load_course_workbook(path)
    assert calculation_attempts(ds)[0].quality_points=='9'


def test_xls_session_container_roundtrip(tmp_path):
    import session_store
    from pathlib import Path
    source=Path(__file__).resolve().parents[1]/'diagnostic_assets/FICTIONAL-reader-check.xls'
    ds=load_course_workbook(source)
    target=tmp_path/'old-format.lguexam'
    session_store.save(target,'course',source,None,{}, {},ds)
    assert session_store.restore(target)[1].courses==ds.courses


def test_legacy_reports_do_not_invent_cover_metadata_or_certificate_year():
    from reports import _dmc_settings,_gazette_settings,_dmc_cert_no
    frame=pd.DataFrame()
    for settings in (_dmc_settings(frame),_gazette_settings(frame)):
        assert all(settings[key]=='' for key in ('department','semester','session','declaration_date'))
    assert _dmc_cert_no(frame,1)=='DMC-00001'
    frame.attrs['settings']={'declaration_date':'2031-04-12'}
    assert _dmc_cert_no(frame,7)=='DMC-2031-00007'
