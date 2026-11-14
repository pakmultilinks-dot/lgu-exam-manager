from dataclasses import replace
from decimal import Decimal
from fractions import Fraction
import json,zipfile
import pytest
from academic_engine import load_course_workbook
from grading_profiles import default_grade_scale,scale_coverage
from calculation_policy import calculation_attempts,AcademicIndex,display_ratio
from test_academic_engine import _df


@pytest.mark.parametrize('grade,rate',[('F','0.00'),('D','1.00'),('C-','1.70'),('C','2.00'),('C+','2.30'),('B-','2.70'),('B','3.00'),('B+','3.30'),('A-','3.70'),('A','4.00')])
def test_supplied_rates_and_weighted_points(tmp_path,grade,rate):
    frame=_df().iloc[:1].copy();frame['GPs']='';frame['Grade']=grade
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    assert ds.grade_scale[grade]==rate
    attempts=calculation_attempts(ds)
    assert Decimal(attempts[0].quality_points)==Decimal(rate)*3
    assert ds.courses[0].quality_points==''
    index=AcademicIndex(attempts);first=ds.courses[0]
    assert index.summary(first.registration,first.degree,'S26').exact_ratio==Fraction(Decimal(rate))


def test_profile_is_copied_and_source_points_take_priority(tmp_path):
    frame=_df();source=tmp_path/'source.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    ds.grade_scale['B']='3.25'
    assert default_grade_scale()['B']=='3.00'
    assert calculation_attempts(ds)==ds.courses
    assert default_grade_scale()['F']=='0.00'


def test_historical_repeats_use_supplied_scale_without_rounding(tmp_path):
    frame=_df().iloc[:3].copy();frame['GPs']=''
    source=tmp_path/'source.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    original=list(ds.courses);first=original[0]
    result=AcademicIndex(calculation_attempts(ds)).summary(first.registration,first.degree,'SU26',True)
    assert result.total_grade_points==Decimal('22.20') and result.total_credit_hours==7
    assert display_ratio(result.exact_ratio)=='3.17'
    assert ds.courses==original and scale_coverage(ds)['derived']==3
    from academic_reports import generate_draft
    from openpyxl import load_workbook
    output=tmp_path/'audit.xlsx';generate_draft('excel',ds,ds.courses,str(output))
    book=load_workbook(output);audit=book['Point Derivation Audit']
    assert audit['G2'].value is None and audit['H2'].value=='3.00' and audit['I2'].value=='9.00'
    assert audit['J2'].value=='Applied grade scale'
    book.close()


def test_legacy_session_gets_scale_but_saved_manual_clear_is_respected(tmp_path):
    import session_store
    source=tmp_path/'source.xlsx';frame=_df();frame['GPs']='';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    target=tmp_path/'legacy.lguexam';session_store.save(target,'course',source,None,{}, {},ds)
    with zipfile.ZipFile(target) as z:members={name:z.read(name) for name in z.namelist()}
    snapshot=json.loads(members['dataset.json']);snapshot.pop('grade_scale_user_set');snapshot['grade_scale']={};snapshot['grade_scale_reference']=''
    members['dataset.json']=json.dumps(snapshot).encode()
    with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in members.items():z.writestr(name,data)
    restored=session_store.restore(target)[1];assert restored.grade_scale==default_grade_scale()
    restored.grade_scale={};restored.grade_scale_user_set=True
    session_store.save(target,'course',source,None,{}, {},restored)
    assert session_store.restore(target)[1].grade_scale=={}


def test_auto_profiles_receive_new_f_rate_but_legacy_custom_scales_stay_custom(tmp_path):
    import session_store
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    target=tmp_path/'session.lguexam';ds.grade_scale.pop('F');ds.grade_scale_user_set=False
    session_store.save(target,'course',source,None,{}, {},ds)
    assert session_store.restore(target)[1].grade_scale['F']=='0.00'
    with zipfile.ZipFile(target) as z:members={name:z.read(name) for name in z.namelist()}
    snapshot=json.loads(members['dataset.json']);snapshot.pop('grade_scale_user_set');snapshot['grade_scale']={'A':'3.99'}
    members['dataset.json']=json.dumps(snapshot).encode()
    with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in members.items():z.writestr(name,data)
    restored=session_store.restore(target)[1]
    assert restored.grade_scale=={'A':'3.99'} and restored.grade_scale_user_set
    session_store.save(target,'course',source,None,{}, {},restored)
    assert session_store.restore(target)[1].grade_scale=={'A':'3.99'}
