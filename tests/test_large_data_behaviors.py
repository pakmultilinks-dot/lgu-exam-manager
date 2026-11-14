from decimal import Decimal
import pandas as pd
from academic_engine import load_course_workbook
from calculation_policy import AcademicIndex,display_ratio
from semester_reports import semester_summary_table
from test_academic_engine import _df

def test_program_typography_does_not_split_cumulative_history(tmp_path):
    frame=_df();frame['DEGREE TITLE']='BS Maths';frame.loc[2,'DEGREE TITLE']='BS MATHS'
    source=tmp_path/'mixed-case.xlsx';frame.to_excel(source,index=False)
    dataset=load_course_workbook(source)
    assert len(dataset.degrees)==1
    assert len(dataset.filtered(degree='bs maths'))==4
    assert dataset.courses[2].degree=='BS MATHS'  # original source spelling stays intact
    student=dataset.courses[0]
    result=AcademicIndex(dataset.courses).summary(student.registration,'bs maths','SU26',True)
    assert result.total_credit_hours==Decimal(7)
    assert result.total_grade_points==Decimal('22.2')
    assert display_ratio(result.exact_ratio)=='3.17'
    head,rows=semester_summary_table(dataset,dataset.courses)
    assert len(rows)==3
    assert rows[1][head.index('CGPA')]=='3.17'

def test_status_only_attempts_are_retained_without_fabricating_credits(tmp_path):
    base=_df().iloc[[0]]
    statuses=pd.concat([base]*4,ignore_index=True)
    statuses['COURSE CODE']=['IP1','W1','PASS1','FINAL1']
    statuses['Grade']=['IP','W','PASS','A']
    statuses['CRH']=['','-','','']
    statuses['GPs']=None
    source=tmp_path/'statuses.xlsx';statuses.to_excel(source,index=False)
    dataset=load_course_workbook(source)
    assert len(dataset.courses)==3 and dataset.quarantined_rows==1
    assert [c.grade for c in dataset.courses]==['IP','W','PASS']
    assert [c.credits for c in dataset.courses]==['','','']
    assert dataset.source_rows[1]['cells']['CRH']=='-'
    assert sum(i.code=='NON_GPA_CREDITS_UNAVAILABLE' for i in dataset.issues)==3
    assert any(i.code=='INVALID_CREDITS' and i.level=='ERROR' for i in dataset.issues)
    assert len(dataset.source_rows)==4


def test_registration_case_does_not_split_history_or_fee_identity(tmp_path):
    from academic_engine import load_fee_workbook
    from fee_register import selected_defaulters
    frame=_df();frame.loc[2,'REGISTRATION NUMBERS']='fa-2025/bscs/001'
    source=tmp_path/'mixed-registration.xlsx';frame.to_excel(source,index=False)
    dataset=load_course_workbook(source)
    assert len(dataset.students)==2
    assert dataset.courses[2].registration=='fa-2025/bscs/001'
    result=AcademicIndex(dataset.courses).summary('FA-2025/BSCS/001','BSCS','SU26',True)
    assert result.total_credit_hours==Decimal(7)
    assert display_ratio(result.exact_ratio)=='3.17'
    _,rows=semester_summary_table(dataset,dataset.courses)
    assert len(rows)==3
    fee=tmp_path/'fees.xlsx'
    pd.DataFrame({'Registration Number':['FA-2025/BSCS/001']}).to_excel(fee,index=False)
    dataset.fee_records=load_fee_workbook(fee,dataset)
    assert len(selected_defaulters(dataset,dataset.courses))==1

