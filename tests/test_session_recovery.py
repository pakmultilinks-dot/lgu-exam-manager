import zipfile,json
import pytest
from academic_engine import load_course_workbook
from session_store import save,restore
from test_academic_engine import _df


def test_dataset_only_session_can_be_saved_repeatedly(tmp_path):
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);dataset=load_course_workbook(source)
    target=tmp_path/'combined.lguexam'
    save(target,'course',None,None,{}, {},dataset)
    with zipfile.ZipFile(target) as archive:assert 'data.xlsx' not in archive.namelist()
    for _ in range(3):
        loaded=restore(target)[1]
        save(target,'course',None,None,{}, {},loaded)
        assert restore(target)[1].courses==dataset.courses


def test_failed_session_write_preserves_previous_archive(tmp_path,monkeypatch):
    import session_store
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False);dataset=load_course_workbook(source)
    target=tmp_path/'session.lguexam';save(target,'course',source,None,{}, {},dataset)
    previous=target.read_bytes()
    def fail(*a,**kw):raise TypeError('Simulated serialization failure')
    monkeypatch.setattr(session_store.json,'dumps',fail)
    with pytest.raises(TypeError):save(target,'course',source,None,{}, {},dataset)
    assert target.read_bytes()==previous and not list(tmp_path.glob('.lgu-export-*'))


def test_snapshot_limit_preserves_previous_save(tmp_path,monkeypatch):
    import session_store
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False)
    dataset=load_course_workbook(source);target=tmp_path/'session.lguexam'
    save(target,'course',source,None,{}, {},dataset)
    previous=target.read_bytes()
    monkeypatch.setattr(session_store,'MAX_DATASET_BYTES',1)
    with pytest.raises(ValueError,match='exceeds limits'):
        save(target,'course',source,None,{}, {},dataset)
    assert target.read_bytes()==previous
    with pytest.raises(ValueError,match='exceeds limits'):restore(target)


@pytest.mark.parametrize('contents',[b'not a zip',None])
def test_damaged_session_has_actionable_error_and_is_untouched(tmp_path,contents):
    target=tmp_path/'bad.lguexam'
    if contents is None:
        with zipfile.ZipFile(target,'w') as archive:archive.writestr('mapping.json','{}')
    else:target.write_bytes(contents)
    original=target.read_bytes()
    with pytest.raises(ValueError,match='incomplete or damaged'):restore(target)
    assert target.read_bytes()==original


def test_word_defaulters_are_withheld_and_listed_at_end(tmp_path):
    import pandas as pd
    from academic_reports import generate_draft
    from docx import Document
    frame=pd.concat([_df().iloc[:1]]*4,ignore_index=True)
    frame['REGISTRATION NUMBERS']=[f'Fa-2025/BSCS/{number:03}' for number in range(1,5)]
    frame['STUDENT NAME']=[f'FICTIONAL STUDENT {number}' for number in range(1,5)]
    frame['Fee Status']=['YES','YES','YES','NO']
    source=tmp_path/'fees.xlsx';frame.to_excel(source,index=False);dataset=load_course_workbook(source)
    target=tmp_path/'fees.docx';generate_draft('word',dataset,dataset.courses,str(target))
    doc=Document(target);paragraphs='\n'.join(p.text for p in doc.paragraphs)
    assert 'FICTIONAL STUDENT 4' in paragraphs
    assert all(f'FICTIONAL STUDENT {n}' not in paragraphs for n in range(1,4))
    assert 'FEE DEFAULTERS' in paragraphs
    table=' '.join(cell.text for row in doc.tables[-1].rows for cell in row.cells)
    assert all(f'FICTIONAL STUDENT {n}' in table for n in range(1,4))
    assert 'FICTIONAL STUDENT 4' not in table
