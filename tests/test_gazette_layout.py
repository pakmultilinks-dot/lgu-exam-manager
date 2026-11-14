"""Reference-layout regression coverage: complete source audit and distinct cohort schemas."""
from pathlib import Path
from dataclasses import replace
import json
import pymupdf as fitz
from academic_engine import ResultDataset
from academic_reports import generate_draft
from gazette_layout import _profile,_plan
from test_academic_engine import _df
from academic_engine import load_course_workbook

def test_automatic_template_uses_program_not_sample(tmp_path):
    from gazette_options import resolve_template
    import pytest
    source=tmp_path/'program.xlsx';_df().to_excel(source,index=False)
    dataset=load_course_workbook(source)
    assert resolve_template('auto',dataset.courses)=='computer-science'
    management=[replace(c,degree='BBA') for c in dataset.courses]
    assert resolve_template('auto',management)=='management-science'
    with pytest.raises(ValueError,match='spans department'):
        resolve_template('auto',dataset.courses+management)
    with pytest.raises(ValueError,match='not uniquely'):
        resolve_template('auto',[replace(dataset.courses[0],degree='UNKNOWN')])

def test_complete_reference_catalog():
    catalog=json.loads((Path(__file__).resolve().parents[1]/'gazette_assets/reference_layouts.json').read_text())
    assert catalog['audit']['computer-science']['pages']==964
    assert catalog['audit']['management-science']['pages']==571
    for audit in catalog['audit'].values():
        assert audit['body_pages_profiled']+len(audit.get('fee_only_pages',[]))==audit['pages']-2
        assert not audit['special_pages']
    assert any(p.get('fee_only_source_pages') for p in catalog['profiles'].values())
    assert _profile('computer-science','FALL 2020','BSCS')['cgpa'] is False
    assert _profile('computer-science','FALL 2024','BSCS')['cgpa'] is True
    assert _profile('management-science','FALL 2024','BBA')['cgpa'] is True
    assert all('MANAGEMENT' not in item['text'] for item in _profile('computer-science','FALL 2024','BSCS')['header'])
    assert any('MANAGEMENT' in item['text'] for item in _profile('management-science','FALL 2024','BBA')['header'])

def test_exam_windows_do_not_mix_all_history(tmp_path):
    source=tmp_path/'source.xlsx';df=_df();df.loc[0,'SHORT SEMESTER']='F25';df.to_excel(source,index=False)
    ds=load_course_workbook(source);out=tmp_path/'gazette.pdf'
    generate_draft('pdf',ds,ds.courses,str(out))
    with fitz.open(out) as pdf:
        body=[p.get_text() for p in list(pdf)[2:]]
        # One block per student, current semester only: no FALL 2025
        # result section, and the F25 course never appears on a Spring
        # 2026 page.
        assert all('SEMESTER -  SPRING 2026' in t for t in body)
        assert all('F25' not in t for t in body if 'SEMESTER -  SPRING 2026' in t)
        assert any('CGPA' in t for t in body)
        # ALPHA shows CS101 twice (S26 + SU26); CS100 (F25) not at all.
        # BETA shows CS100 once (S26).
        assert sum(t.count('CS100') for t in body)==1
        assert sum(t.count('CS101') for t in body)==2


def test_automatic_combined_programs_and_aliases(tmp_path):
    from gazette_options import resolve_template_groups
    source=tmp_path/'programs.xlsx';_df().to_excel(source,index=False)
    dataset=load_course_workbook(source)
    base=dataset.courses
    rows=[replace(c,degree=degree,registration=c.registration+'-'+str(i))
          for i,degree in enumerate(('BSCS','BS IT','BS SE','BBA (Hons)')) for c in base]
    groups=resolve_template_groups('auto',rows)
    assert [key for key,_ in groups]==['computer-science','information-technology','software-engineering','management-science']
    assert sum(len(group) for _,group in groups)==len(rows)
    dataset.courses=rows
    target=tmp_path/'combined.pdf'
    generate_draft('pdf',dataset,rows,target,'auto')
    with fitz.open(target) as pdf:
        texts=[page.get_text() for page in pdf]
        fees=[i for i,text in enumerate(texts) if 'FEE DEFAULTERS' in text]
        assert len(fees)==4
        # One fee block per department group, closing that group's own
        # pages (not a single appendix after all groups), so the fee
        # pages are spread through the document and the last one is the
        # document's last page.
        assert fees[-1]==len(texts)-1
        assert fees!=list(range(len(texts)-4,len(texts)))
        for degree,fee_page in zip(('BSCS','BS IT','BS SE','BBA (Hons)'),fees):
            assert any(degree in text for text in texts[:fee_page])
