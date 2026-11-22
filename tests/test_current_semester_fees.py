from dataclasses import replace
from decimal import Decimal
from fractions import Fraction
from pathlib import Path
import pandas as pd
import pymupdf
import pytest
from academic_engine import load_course_workbook,load_fee_workbook,FeeRecord
from academic_reports import generate_draft
from calculation_policy import AcademicIndex
from course_reports import export_report,table_data
from gazette_options import gazette_rows
from test_academic_engine import _df

def test_current_semester_only_and_earlier_history(tmp_path):
    source=tmp_path/'history.xlsx';_df().to_excel(source,index=False);ds=load_course_workbook(source)
    settings={'gazette_term':'SU26','cover_department':'TEST FACULTY','cover_period':'SUMMER 2026'}
    # v4.5.2: gazette_rows returns the full history through the term cutoff, not just the current term.
    assert {c.term for c in gazette_rows(ds,ds.courses,settings)}=={'S26','SU26'}
    first=ds.courses[0]
    assert AcademicIndex(ds.courses).summary(first.registration,first.degree,'SU26',True).exact_ratio==Fraction(Decimal('22.2'))/7
    output=tmp_path/'current.pdf';generate_draft('pdf',ds,ds.courses,str(output),settings=settings)
    with pymupdf.open(output) as doc:
        body='\n'.join(p.get_text() for p in list(doc)[2:])
        # Official Spring format (v4.6.10): the block lists the Spring
        # courses and the Summer retake, and the totals row combines the
        # Spring + Summer window, matching the Excel gazette. CS101 was
        # retaken in Summer (B+, 13.2 GP) over Spring (B-, 10.8 GP), so
        # the repeat rule keeps the Summer attempt and the block totals
        # are 22.20 / 7 / 3.17, not the Spring-only 19.80 / 7 / 2.83.
        # The cumulative CGPA 3.17 is asserted on the CGPA list below too.
        assert 'CS101' in body and 'CS100' in body
        assert '3.17' in body
        assert 'TEST FACULTY' in doc[0].get_text()
        assert 'FEE DEFAULTERS' in doc[-1].get_text()
        from gazette_layout import _profile
        profile=_profile('computer-science',first.admission_session,first.degree,True)
        edges=profile['edges'];top=profile['body_top']
        for column,expected in {8:'22.20',10:'22.20',11:'7',13:'7',14:'3.17',15:'3.17'}.items():
            text=doc[2].get_textbox(pymupdf.Rect(edges[column]+1,top,edges[column+1]-1,565))
            assert expected in text,(column,text)
    head,records=table_data('cgpa_list',ds,ds.filtered(term='SU26'))
    assert len(records)==2  # same selected student, both earlier and current terms
    assert records[1][head.index('CGPA')]=='3.17'
    for fmt,extension in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
        export_report('cgpa_list',fmt,ds,ds.courses,str(tmp_path/('summary'+extension)))

def test_three_defaulters_are_only_in_final_appendix(tmp_path):
    frame=_df().iloc[[0]].copy()
    frame=pd.concat([frame]*4,ignore_index=True)
    frame['REGISTRATION NUMBERS']=[f'Fa-2025/BSCS/{i:03}' for i in range(1,5)]
    frame['STUDENT NAME']=[f'FICTIONAL STUDENT {i}' for i in range(1,5)]
    frame['Fee status']=['YES','YES','YES','NO'];frame['Semester Number']=14
    source=tmp_path/'fees.xlsx';frame.to_excel(source,index=False);ds=load_course_workbook(source)
    assert len(ds.fee_records)==4 and not ds.issues
    output=tmp_path/'fees.pdf';generate_draft('pdf',ds,ds.courses,str(output),settings={'gazette_term':'S26'})
    with pymupdf.open(output) as doc:
        appendix=doc[-1].get_text()
        previous='\n'.join(p.get_text() for p in list(doc)[2:-1])
        for i in range(1,4):
            assert f'FICTIONAL STUDENT {i}' in appendix
            assert f'FICTIONAL STUDENT {i}' not in previous
        assert 'FICTIONALSTUDENT4' in ''.join(previous.split()) and 'FICTIONAL STUDENT 4' not in appendix
    import session_store
    saved=tmp_path/'session.lguexam';session_store.save(saved,'course',source,None,{}, {},ds)
    assert session_store.restore(saved)[1].fee_records==ds.fee_records
    conflict=replace(ds.fee_records[0],status='CLEAR');ds.fee_records.append(conflict)
    with pytest.raises(ValueError,match='Conflicting fee'):
        generate_draft('pdf',ds,ds.courses,str(tmp_path/'conflict.pdf'))

def test_separate_fee_import_and_semester_limit(tmp_path):
    source=tmp_path/'results.xlsx';frame=_df();frame['Semester Number']=15;frame.to_excel(source,index=False)
    dataset=load_course_workbook(source)
    assert any(i.code=='INVALID_SEMESTER_NUMBER' for i in dataset.issues)
    frame['Semester Number']=1;frame.to_excel(source,index=False);dataset=load_course_workbook(source)
    fee_file=tmp_path/'separate.xlsx'
    pd.DataFrame({'Registration Number':[dataset.courses[0].registration]}).to_excel(fee_file,index=False)
    fees=load_fee_workbook(fee_file,dataset)
    assert len(fees)==1 and fees[0].status=='DEFAULTER' and fees[0].name==dataset.courses[0].name
    pd.DataFrame({'Registration Number':['UNKNOWN']}).to_excel(fee_file,index=False)
    with pytest.raises(ValueError,match='unknown registration'):load_fee_workbook(fee_file,dataset)

def test_fourteen_semesters_and_cumulative_progression(tmp_path):
    frame=pd.concat([_df().iloc[[0]]]*14,ignore_index=True)
    frame['REGISTRATION NUMBERS']='Fa-2019/BSCS/001'
    frame['Semester Number']=list(range(1,15))
    frame['COURSE CODE']=[f'TEST{i:03}' for i in range(1,15)]
    frame['SHORT SEMESTER']=[('F'+str(19+i//2) if i%2==0 else 'S'+str(20+i//2)) for i in range(14)]
    frame['GPs']=[9 if i%2==0 else 12 for i in range(14)]
    frame['Grade']=['B' if i%2==0 else 'A' for i in range(14)]
    source=tmp_path/'fourteen.xlsx';frame.to_excel(source,index=False)
    dataset=load_course_workbook(source);assert not dataset.issues
    head,rows=table_data('cgpa_list',dataset,dataset.courses)
    assert len(rows)==14 and rows[-1][head.index('Semester number')]=='14'
    assert rows[-1][head.index('Cumulative CH')]=='42'
    assert rows[-1][head.index('CGPA')]=='3.50'
    current=gazette_rows(dataset,dataset.courses,{'gazette_term':'S26'})
    # v4.5.2: gazette_rows returns the full history through the term cutoff, not just the current term.
    assert len(current)==14 and current[-1].course_code=='TEST014'
    missing=replace(dataset,courses=[c for c in dataset.courses if c.semester_number!='7'])
    from calculation_policy import CalculationReviewRequired
    with pytest.raises(CalculationReviewRequired,match='Earlier academic semesters'):
        AcademicIndex(missing.courses).summary(current[-1].registration,current[-1].degree,'S26',True)

def test_adjusted_gpa_reference_is_not_mislabeled_as_cgpa(tmp_path):
    from gazette_layout import _profile
    profile=_profile('criminology','SPRING 2026','BSCRCSI')
    assert profile['last_ratio_role']=='adjusted_gpa' and not profile['cgpa']
    source=tmp_path/'criminology.xlsx';frame=_df().iloc[[0]].copy();frame['DEGREE TITLE']='BSCRCSI';frame.to_excel(source,index=False)
    dataset=load_course_workbook(source);output=tmp_path/'criminology.pdf'
    generate_draft('pdf',dataset,dataset.courses,str(output),template_key='criminology',settings={'gazette_term':'S26'})
    with pymupdf.open(output) as document:
        # The fee block closes its own section; the GPA / CGPA summary
        # pages come after all body and fee pages, so they are last.
        assert 'GPA / CGPA SUMMARY' in document[-1].get_text()
        assert '3.00' in document[-1].get_text()
        assert 'FEE DEFAULTERS' in document[-2].get_text()
