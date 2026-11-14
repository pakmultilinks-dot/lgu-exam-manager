from fractions import Fraction
from decimal import Decimal
import pandas as pd
from openpyxl import load_workbook
from academic_engine import load_course_workbook
from academic_reports import generate_draft
from calculation_policy import AcademicIndex
from test_academic_engine import _df

def test_varied_headers_full_terms_and_repeated_headings(tmp_path):
    frame=_df().rename(columns={'REGISTRATION NUMBERS':'Enrollment Number','CRH':'Credit Hrs','SHORT SEMESTER':'Term Code'})
    frame['Term Code']=['Spring 2026','S-2026','Summer 2026','S26']
    frame['Intake']='FALL 2025'
    mixed=pd.concat([frame.iloc[:2],pd.DataFrame([dict(zip(frame.columns,frame.columns))]),frame.iloc[2:]],ignore_index=True)
    path=tmp_path/'unseen.xlsx'
    with pd.ExcelWriter(path) as writer:
        mixed.to_excel(writer,index=False,startrow=63,sheet_name='Arbitrary Department')
    dataset=load_course_workbook(path)
    assert len(dataset.courses)==4
    assert dataset.terms==['S26','SU26']
    assert len(dataset.source_rows)==5
    assert any(i.code=='REPEATED_HEADER' for i in dataset.issues)

def test_history_is_not_lost_by_display_filter_and_excel_audit(tmp_path):
    source=tmp_path/'history.xlsx';_df().to_excel(source,index=False)
    ds=load_course_workbook(source)
    index=AcademicIndex(ds.courses)
    regular=index.summary(ds.courses[0].registration,'BSCS','S26',True)
    summer=index.summary(ds.courses[0].registration,'BSCS','SU26',True)
    assert regular.exact_ratio==Fraction(Decimal('19.8'))/7
    assert summer.total_credit_hours==7
    assert summer.total_grade_points==Decimal('22.2')
    out=tmp_path/'audit.xlsx'
    generate_draft('excel',ds,ds.filtered(term='SU26'),str(out))
    workbook=load_workbook(out,data_only=True)
    totals=list(workbook['Calculated Totals'].values)
    assert totals[2][4:6]==('7','22.2')
    assert any(row[5]=='EXCLUDED' for row in workbook['Repeat Selection Audit'].values)

def test_ten_thousand_unseen_rows(tmp_path):
    frame=_df().iloc[[0]].copy()
    frame=pd.concat([frame]*10000,ignore_index=True)
    frame['REGISTRATION NUMBERS']=[f'Sp-2024/BSCS/{i:06d}' for i in range(len(frame))]
    source=tmp_path/'large.xlsx';frame.to_excel(source,index=False)
    ds=load_course_workbook(source)
    assert len(ds.courses)==10000 and len(ds.students)==10000
    assert not ds.issues
