"""Opt-in desktop matrix, fictional records only; exercises real Tk button callbacks."""
from pathlib import Path
import sys,time,json,hashlib
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
from docx import Document
import pymupdf
from app import App
from tkinter import messagebox,filedialog
from academic_identity import program_key
from course_reports import ROUTES
from gazette_options import resolve_template_groups

PROGRAMS={'computer-science':'BSCS','management-science':'BBA (Hons)','allied-health-biology':'BSBT','art-and-humanities':'BSIR','chemistry':'BS CHEM.','criminology':'BSCRCSI','english':'BS Eng.','information-technology':'BS IT','mathematics':'BS Maths','media-and-communication':'BS MC','physics':'BS Phys','psychology':'BS AP','software-engineering':'BS SE'}

PROGRAMS.update({'historical-accounting':'BS AF','historical-biochemistry':'BS BCH','historical-dfcs':'BS DFCS','historical-microbiology':'BS MB','historical-zoology':'BS ZOO','historical-urdu-certificate':'CERUF','historical-mphil-chemistry':'M.Phil CHEM.','historical-mphil-microbiology':'M.Phil MB','historical-mphil-urdu':'M.Phil Urdu','historical-mphil-zoology':'M.Phil ZOO','historical-mba':'MBA (2 Years)','historical-msba':'MSBA','historical-mscp':'MSCP','historical-phd-urdu':'Ph. D URDU'})
COUNT=len(PROGRAMS)

def main():
    root=Path('.private/cross-program-ui');root.mkdir(parents=True,exist_ok=True)
    source=root/'FICTIONAL-all-departments.xlsx';book=Workbook();sheet=book.active
    sheet.append(['Registration Number','Student Name','Degree','Department','Term Code','Semester Number','Course Code','Course Name','Credit Hours','Grade','Fee Status'])
    for number,(department,program) in enumerate(PROGRAMS.items(),1):
        for student in range(1,5):
            for term,semester,courses in [('F25',1,[('REP101','Repeated course',3,'D'),('HIST102','Earlier course',2,'B'),('HIST103','Earlier elective',1,'C')]),('S26',2,[('REP101','Repeated course',3,'A'),('CUR104','Current course',4,'B+'),('CUR105','Fractional credit course',0.5,'F')])]:
                for code,title,ch,grade in courses:
                    sheet.append([f'Fa-2025/TEST{number:02}/{student:03}',f'FICTIONAL P{number:02} STUDENT {student}',program,department.upper(),term,semester,code,title,ch,grade,'YES' if student<4 else 'NO'])
    book.save(source);book.close();source_hash=hashlib.sha256(source.read_bytes()).hexdigest()
    app=App(restore_autosave=False);app._autosave_path=lambda:str(root/'autosave.lguexam');errors=[];warnings=[];checks=[];beats=[]
    messagebox.showerror=lambda *a,**kw:errors.append(str(a))
    messagebox.showwarning=lambda *a,**kw:warnings.append(str(a))
    messagebox.showinfo=lambda *a,**kw:None
    def beat():beats.append(time.monotonic());app.after(25,beat)
    app.after(25,beat)
    def settle():
        deadline=time.monotonic()+180
        while app.job_busy and time.monotonic()<deadline:app.update();time.sleep(.005)
        app.update();assert not app.job_busy and not errors and not warnings,(errors,warnings)
    def export(panel,fmt,path):
        filedialog.asksaveasfilename=lambda **kw:str(path)
        panel.format_buttons[['pdf','word','excel'].index(fmt)].invoke();settle()
        assert path.stat().st_size>500
    try:
        app.show_frame('reports');app.update();assert app.academic_dataset is None
        checks.append('reports navigation before import')
        start=time.monotonic();app.frames['import']._smart_load(str(source.resolve()));settle();import_time=time.monotonic()-start
        ds=app.academic_dataset
        assert len(ds.courses)==COUNT*24 and len(ds.students)==COUNT*4 and len(ds.degrees)==COUNT and ds.counts()['errors']==0 and ds.counts()['warnings']==0,ds.counts()
        original=tuple(ds.courses)
        for key in app.frames:app.show_frame(key);app.update()
        checks.append(f'background import; {COUNT} programs; {COUNT*4} students; {COUNT*24} attempts; every navigation page')
        academic=app.frames['academic'];app.show_frame('academic')
        pages=(len(ds.courses)+149)//150
        for expected_page in range(1,pages):
            academic.next_button.invoke();app.update();assert academic.page==expected_page
        assert len(academic.tree.get_children())==(len(ds.courses)-150*(pages-1)) and academic.next_button.cget('state')=='disabled'
        academic.prev_button.invoke();assert academic.page==pages-2
        academic.degrees.set(ds.degrees[0]);academic.reset_page();assert academic.page==0
        academic.degrees.set('All Degrees');academic.refresh_rows()
        checks.append('paging through every page; next/previous boundaries; filter resets page')
        settings=app.frames['settings'].course_panel;settings.select_term('S26');app.report_settings['gazette_template']='auto'
        app.show_frame('reports');panel=app.frames['reports'].course_panel;panel.refresh()
        assert len(panel.rows())==COUNT*12 and len(resolve_template_groups('auto',panel.rows()))==13
        outputs=0
        for department,program in PROGRAMS.items():
            panel.program_menu.set(program);panel.refresh();app.update()
            assert len(panel.rows())==12 and panel.format_buttons[0].cget('state')=='normal',(program,panel.reason.cget('text'))
            for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
                target=root/(department+ext);export(panel,fmt,target);outputs+=1
                if fmt=='pdf':
                    with pymupdf.open(target) as document:
                        text='\n'.join(page.get_text() for page in document)
                        assert '3.36' in text and '3.16' in text,(program,text)
                        assert 'HIST102' not in text and 'CUR104' in text and 'CUR105' in text
                        assert 'FEE DEFAULTERS' in document[-1].get_text()
                        if program in ('CERUF','M.Phil Urdu','Ph. D URDU'):
                            assert 'DEPARTMENT OF URDU' in text and 'INTERNATIONAL RELATION' not in text
                        assert 'LAHORE GARRISON UNIVERSITY' in document[2].get_text()
                elif fmt=='excel':
                    w=load_workbook(target);assert w['Course Results'].max_row==13;assert w['Repeat Selection Audit'].max_row>13;w.close()
                else:
                    d=Document(target);text='\n'.join(p.text for p in d.paragraphs)+'\n'+'\n'.join(cell.text for table in d.tables for row in table.rows for cell in row.cells)
                    assert '3.16' in text and 'FEE DEFAULTERS' in text
        checks.append(f'{COUNT} program filters; {COUNT*3} visible PDF/Word/Excel button exports; current-only courses; independent GPA 3.36 and CGPA 3.16 checks')
        panel.program_menu.set('All programs');panel.refresh()
        for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
            export(panel,fmt,root/('all-programs'+ext));outputs+=1
        with pymupdf.open(root/'all-programs.pdf') as document:
            texts=[page.get_text() for page in document];fee_pages=[i for i,text in enumerate(texts) if 'FEE DEFAULTERS' in text]
            assert len(fee_pages)>=13 and fee_pages==list(range(fee_pages[0],len(texts)))
            before='\n'.join(texts[:fee_pages[0]]);after='\n'.join(texts[fee_pages[0]:])
            for number in range(1,COUNT+1):
                for student in range(1,4):
                    name=f'FICTIONAL P{number:02} STUDENT {student}'
                    assert name not in before and name in after,name
                assert f'FICTIONALP{number:02}STUDENT4' in ''.join(before.split())
        checks.append(f'automatic combined department export; all {COUNT*3} defaulters withheld and placed after every academic section')
        for term in ('F25','S26'):
            settings.select_term(term);panel.refresh();assert len(panel.rows())==COUNT*12
            assert all(c.term==term for c in panel.rows())
        settings.select_term('S26')
        academic.terms.set('F25');academic.refresh_rows();panel.refresh();assert {c.term for c in panel.rows()}=={'S26'}
        academic.terms.set('All Terms');academic.refresh_rows()
        checks.append('Settings term is authoritative; academic attempt-term filter does not remove current gazette or earlier CGPA history')
        panel.route_menu.set(ROUTES['cgpa_list']);panel.refresh()
        export(panel,'excel',root/'cgpa-list.xlsx');outputs+=1
        w=load_workbook(root/'cgpa-list.xlsx');data=list(w.active.values);w.close()
        assert len(data)>COUNT*4
        checks.append('roll-number-wise CGPA list includes both semesters')
        panel.route_menu.set(ROUTES['register']);panel.refresh()
        settings.select_term('F99');panel.refresh();assert panel.format_buttons[0].cget('state')=='disabled'
        settings.select_term('S26');panel.refresh();assert panel.format_buttons[0].cget('state')=='normal'
        academic.search_var.set('NO-SUCH-STUDENT');deadline=time.monotonic()+.35
        while time.monotonic()<deadline:app.update();time.sleep(.005)
        panel.refresh();assert panel.format_buttons[0].cget('state')=='disabled'
        academic.search_var.set('');deadline=time.monotonic()+.35
        while time.monotonic()<deadline:app.update();time.sleep(.005)
        panel.refresh();assert panel.format_buttons[0].cget('state')=='normal'
        checks.append('unavailable term and empty search block export; restoring valid selection re-enables buttons')
        import course_ui
        previews=[];course_ui.os.startfile=lambda path:previews.append(path)
        panel.preview_button.invoke();settle();assert len(previews)==1 and Path(previews[0]).exists()
        checks.append('combined PDF preview button')
        app.save_session();settle()
        import session_store
        loaded=session_store.restore(root/'autosave.lguexam')[1]
        assert loaded.courses==ds.courses and loaded.fee_records==ds.fee_records
        assert tuple(ds.courses)==original and hashlib.sha256(source.read_bytes()).hexdigest()==source_hash
        checks.append('session save/reload; raw academic records and source workbook unchanged')
        result={'status':'PASS','departments':13,'programs':COUNT,'students':COUNT*4,'course_attempts':COUNT*24,'fee_defaulters':COUNT*3,'export_files':outputs,'checks':checks,'import_seconds':round(import_time,3),'max_ui_heartbeat_gap_seconds':round(max((b-a for a,b in zip(beats,beats[1:])),default=0),3),'errors':errors,'warnings':warnings}
        (root/'results.json').write_text(json.dumps(result,indent=2));print(json.dumps(result,indent=2))
    finally:
        for job in app.tk.call('after','info'):app.after_cancel(job)
        app.destroy()
if __name__=='__main__':main()
