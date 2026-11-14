"""Independent arithmetic oracles across the complete reference-program catalog."""
from dataclasses import replace
from decimal import Decimal
from fractions import Fraction
import random
import pytest
from academic_engine import CourseAttempt,ResultDataset
from calculation_policy import AcademicIndex,calculation_attempts,CalculationReviewRequired,display_ratio
from gazette_options import _programs,_routing_rules,resolve_template_groups

RATES={'F':'0.00','D':'1.00','C-':'1.70','C':'2.00','C+':'2.30','B-':'2.70','B':'3.00','B+':'3.30','A-':'3.70','A':'4.00'}
CREDITS=('0.5','1','1.5','2','3','4','6')

def attempt(**changes):
    fields=dict(registration='TEST/001',name='FICTIONAL STUDENT',degree='BSCS',admission_session='FALL 2025',semester='FALL 2025',term='F25',course_code='TEST101',course_name='Synthetic course',credits='3',marks='',grade='A',quality_points='',award_id='',source_sheet='Synthetic',source_excel_row=2)
    fields.update(changes)
    return CourseAttempt(**fields)


def test_every_reference_program_grade_and_credit_combination():
    programs=sorted(set().union(*_programs().values()) | set(_routing_rules().get('program_templates',{})))
    rows=[];expected={}
    for program in programs:
        assert sum(len(group) for _,group in resolve_template_groups('auto',[attempt(degree=program)]))==1
        for credit in CREDITS:
            for grade,rate in RATES.items():
                registration=f'TEST/{len(rows):06}'
                rows.append(attempt(degree=program,registration=registration,credits=credit,grade=grade))
                expected[registration]=(Fraction(Decimal(credit)),Fraction(Decimal(rate)))
    ds=ResultDataset(rows,[],[],'fictional.xlsx');original=tuple(rows)
    effective=calculation_attempts(ds);index=AcademicIndex(effective)
    for row in effective:
        credit,rate=expected[row.registration]
        assert Fraction(Decimal(row.quality_points))==credit*rate
        summary=index.summary(row.registration,row.degree,'F25')
        assert Fraction(summary.total_credit_hours)==credit
        assert Fraction(summary.total_grade_points)==credit*rate
        assert summary.exact_ratio==rate
        scaled=rate.numerator*100//rate.denominator
        assert display_ratio(summary.exact_ratio)==f'{scaled//100}.{scaled%100:02}'
    assert tuple(ds.courses)==original and all(not row.quality_points for row in rows)


def test_seeded_repeat_and_history_oracle():
    rng=random.Random(73519);rows=[];oracles={}
    programs=sorted(set().union(*_programs().values()) | set(_routing_rules().get('program_templates',{})))
    for student in range(500):
        degree=programs[student%len(programs)];registration=f'ORACLE/{student:05}'
        chosen={}
        for term in ('F25','S26','SU26'):
            current=[]
            for course in range(5):
                credits=Fraction(Decimal(CREDITS[(student+course)%len(CREDITS)]))
                grade=rng.choice(list(RATES));rate=Fraction(Decimal(RATES[grade]))
                row=attempt(registration=registration,degree=degree,term=term,course_code=f'COURSE{course}',credits=str(Decimal(credits.numerator)/Decimal(credits.denominator)),grade=grade)
                rows.append(row);current.append((credits,rate))
                if course not in chosen or rate>chosen[course][1]:chosen[course]=(credits,rate)
            for cumulative,values in ((False,current),(True,list(chosen.values()))):
                total_ch=sum(c for c,r in values);total_gp=sum(c*r for c,r in values)
                oracles[registration,degree,term,cumulative]=(total_ch,total_gp,total_gp/total_ch)
    ds=ResultDataset(rows,[],[],'fictional.xlsx');index=AcademicIndex(calculation_attempts(ds))
    for key,(ch,gp,ratio) in oracles.items():
        summary=index.summary(*key)
        assert (Fraction(summary.total_credit_hours),Fraction(summary.total_grade_points),summary.exact_ratio)==(ch,gp,ratio)


@pytest.mark.parametrize('credits',['0','-1','','bad','NaN','Infinity'])
def test_invalid_credit_hours_never_become_numeric_gpa(credits):
    ds=ResultDataset([attempt(credits=credits)],[],[],'fictional.xlsx')
    with pytest.raises(CalculationReviewRequired):AcademicIndex(calculation_attempts(ds)).summary('TEST/001','BSCS','F25')


@pytest.mark.parametrize('grade',['W','I','IP','PASS','P','R','','U','A+','D+'])
def test_unapproved_or_nonfinal_grades_never_get_invented_points(grade):
    ds=ResultDataset([attempt(grade=grade)],[],[],'fictional.xlsx')
    with pytest.raises(CalculationReviewRequired):AcademicIndex(calculation_attempts(ds)).summary('TEST/001','BSCS','F25')


def test_repeat_compares_rate_and_keeps_the_winning_attempt_credit_hours():
    rows=[attempt(term='F25',credits='6',grade='C'),attempt(term='S26',credits='1',grade='A')]
    ds=ResultDataset(rows,[],[],'fictional.xlsx');index=AcademicIndex(calculation_attempts(ds))
    earlier=index.summary('TEST/001','BSCS','F25',True)
    current=index.summary('TEST/001','BSCS','S26',True)
    assert (earlier.total_credit_hours,earlier.total_grade_points,earlier.exact_ratio)==(6,12,2)
    assert (current.total_credit_hours,current.total_grade_points,current.exact_ratio)==(1,4,4)
    assert len(current.decisions)==1 and current.decisions[0].selected.term=='S26'


def test_identical_names_and_registration_across_programs_do_not_merge_points():
    rows=[attempt(degree='BSCS',grade='A'),attempt(degree='BS IT',grade='D'),attempt(registration='TEST/002',grade='F')]
    index=AcademicIndex(calculation_attempts(ResultDataset(rows,[],[],'fictional.xlsx')))
    assert index.summary('TEST/001','BSCS','F25').exact_ratio==4
    assert index.summary('TEST/001','BS IT','F25').exact_ratio==1
    assert index.summary('TEST/002','BSCS','F25').exact_ratio==0


@pytest.mark.parametrize('program,expected',[('BS BCH','allied-health-biology'),('BS ZOO','allied-health-biology'),('CERUF','art-and-humanities'),('M.Phil Urdu','art-and-humanities'),('MBA (2 Years)','management-science'),('MSBA','management-science')])
def test_historical_workbook_programs_have_confirmed_department_routes(program,expected,tmp_path):
    from academic_reports import generate_draft
    import pymupdf
    row=attempt(degree=program)
    ds=ResultDataset([row],[],[],'fictional.xlsx')
    assert resolve_template_groups('auto',ds.courses)[0][0]==expected
    path=tmp_path/'historical-program.pdf'
    generate_draft('pdf',ds,ds.courses,path,'auto',settings={'gazette_term':'F25'})
    with pymupdf.open(path) as pdf:
        text=''.join(page.get_text() for page in pdf)
        assert ''.join(program.split()) in ''.join(text.split())
        assert '4.00' in text and 'FEE DEFAULTERS' in pdf[-1].get_text()
    assert ds.courses[0] is row and not row.quality_points


@pytest.mark.parametrize('semester',range(1,15))
def test_all_supported_semester_numbers_survive_import(semester,tmp_path):
    from openpyxl import Workbook
    from academic_engine import load_course_workbook
    path=tmp_path/'semester.xlsx';book=Workbook();sheet=book.active
    sheet.append(['Registration Number','Student Name','Degree','Term Code','Semester Number','Course Code','Course Name','Credit Hours','Grade'])
    sheet.append(['TEST/001','FICTIONAL STUDENT','BSCS','S26',semester,'TEST101','Synthetic course',3,'A']);book.save(path);book.close()
    ds=load_course_workbook(path)
    assert len(ds.courses)==1 and ds.courses[0].semester_number==str(semester)
    assert not any(issue.level=='ERROR' for issue in ds.issues)


@pytest.mark.parametrize('semester',[0,-1,15,'1.5','unknown'])
def test_invalid_semester_numbers_are_flagged_before_publication(semester,tmp_path):
    from openpyxl import Workbook
    from academic_engine import load_course_workbook
    from course_reports import capability
    path=tmp_path/'semester.xlsx';book=Workbook();sheet=book.active
    sheet.append(['Registration Number','Student Name','Degree','Term Code','Semester Number','Course Code','Course Name','Credit Hours','Grade'])
    sheet.append(['TEST/001','FICTIONAL STUDENT','BSCS','S26',semester,'TEST101','Synthetic course',3,'A']);book.save(path);book.close()
    ds=load_course_workbook(path)
    assert any(issue.code=='INVALID_SEMESTER_NUMBER' for issue in ds.issues)
    assert not capability('register',ds,ds.courses)[0]


@pytest.mark.parametrize('program',['BS Urdu','Ph. D URDU','M.Phil Urdu','CERUF'])
def test_urdu_identity_is_preserved_when_table_geometry_is_shared(program,tmp_path):
    from academic_reports import generate_draft
    import pymupdf
    row=attempt(degree=program)
    ds=ResultDataset([row],[],[],'fictional.xlsx')
    target=tmp_path/'urdu.pdf';generate_draft('pdf',ds,ds.courses,target,'auto',settings={'gazette_term':'F25'})
    with pymupdf.open(target) as pdf:
        assert 'DEPARTMENT OF URDU' in pdf[0].get_text()
        for page in list(pdf)[2:]:
            text=page.get_text()
            assert 'DEPARTMENT OF URDU' in text
            assert 'INTERNATIONAL RELATION' not in text
            assert 'LAHORE GARRISON UNIVERSITY' in text


def test_shared_arts_fee_appendix_does_not_mislabel_urdu_as_ir(tmp_path):
    from academic_engine import FeeRecord
    from academic_reports import generate_draft
    import pymupdf
    rows=[attempt(registration=f'TEST/{i}',degree=degree) for i,degree in enumerate(['BS IR','BS Urdu'])]
    fees=[FeeRecord(c.registration,c.name,c.degree,c.term,'DEFAULTER') for c in rows]
    ds=ResultDataset(rows,[],[],'fictional.xlsx',fee_records=fees)
    target=tmp_path/'mixed-arts.pdf';generate_draft('pdf',ds,rows,target,'auto',settings={'gazette_term':'F25'})
    with pymupdf.open(target) as pdf:
        text=pdf[-1].get_text()
        assert 'DEPARTMENT OF ART AND HUMANITIES' in text
        assert all(c.degree in text for c in rows)
        assert 'INTERNATIONAL RELATION' not in text


def test_every_measured_body_profile_contains_the_reference_university_heading():
    from gazette_layout import _catalog
    for name,profile in _catalog()['profiles'].items():
        assert any('LAHORE GARRISON UNIVERSITY' in item['text'] for item in profile['header']),name
