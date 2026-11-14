from dataclasses import replace
from decimal import Decimal
import pandas as pd
import pytest
from academic_engine import load_course_workbook
from course_reports import capability,export_report
from validation_scope import relevant_errors
from calculation_policy import parse_grade_scale,calculation_attempts,AcademicIndex,display_ratio
from test_academic_engine import _df


def test_unrelated_errors_do_not_block_selected_students(tmp_path):
    frame=_df();frame['CRH']=frame['CRH'].astype(object);frame.loc[3,'CRH']='bad'
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False)
    ds=load_course_workbook(source)
    assert ds.counts()['errors']==1
    rows=[c for c in ds.courses if c.registration.endswith('001')]
    assert not relevant_errors(ds,rows)
    assert capability('register',ds,rows)[0]
    export_report('register','excel',ds,rows,str(tmp_path/'selected.xlsx'))


def test_missing_historical_row_blocks_affected_cgpa(tmp_path):
    frame=_df();frame['CRH']=frame['CRH'].astype(object);frame.loc[0,'CRH']='bad'
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False)
    ds=load_course_workbook(source)
    current=[c for c in ds.courses if c.term=='SU26']
    assert len(relevant_errors(ds,current))==1
    assert not relevant_errors(ds,current,history=False)
    assert not capability('register',ds,current)[0]
    assert capability('term',ds,current)[0]


def test_scale_is_explicit_preserves_source_and_survives_session(tmp_path):
    import session_store
    frame=_df();frame['GPs']=''
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False)
    ds=load_course_workbook(source)
    with pytest.raises(ValueError):parse_grade_scale('B=3','')
    with pytest.raises(ValueError):parse_grade_scale('B=NaN','TEST')
    ds.grade_scale=parse_grade_scale('B=3\nB-=2.7\nB+=3.3\nF=0','TEST APPROVED SCALE')
    ds.grade_scale_reference='TEST APPROVED SCALE'
    original=list(ds.courses);index=AcademicIndex(calculation_attempts(ds))
    result=index.summary(original[0].registration,original[0].degree,'SU26',True)
    assert result.total_grade_points==Decimal('22.2')
    assert display_ratio(result.exact_ratio)=='3.17'
    assert ds.courses==original and all(not c.quality_points for c in ds.courses)
    target=tmp_path/'session.lguexam';session_store.save(target,'course',source,None,{}, {},ds)
    restored=session_store.restore(target)[1]
    assert restored.grade_scale==ds.grade_scale and restored.source_rows==ds.source_rows
    export_report('register','excel',restored,restored.courses,str(tmp_path/'audit.xlsx'))
    from openpyxl import load_workbook
    book=load_workbook(tmp_path/'audit.xlsx');assert 'Applied Grade Scale' in book.sheetnames
    book.close()
