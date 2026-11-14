"""Opt-in real workbook Tk command acceptance; never logs identities."""
from pathlib import Path
import sys,time,json,argparse
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app import App
from tkinter import messagebox,filedialog
from openpyxl import load_workbook
from docx import Document
import pymupdf as fitz

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--workbook',required=True);args=parser.parse_args()
    root=Path('.private/gui_acceptance');root.mkdir(parents=True,exist_ok=True)
    App._restore_autosave=lambda self:None
    app=App(restore_autosave=False);errors=[];messages=[];timings={};beats=[];worker_gaps=[]
    messagebox.showerror=lambda *a,**k:errors.append(str(a[0]))
    messagebox.showwarning=lambda *a,**k:errors.append(str(a[0]))
    messagebox.showinfo=lambda *a,**k:messages.append(str(a[0]))
    def beat():beats.append(time.perf_counter());app.after(25,beat)
    app.after(25,beat)
    def settle(timeout=180):
        deadline=time.monotonic()+timeout
        last=[time.perf_counter()];running=[True]
        def worker_beat():
            if not running[0]:return
            now=time.perf_counter();worker_gaps.append(now-last[0]);last[0]=now
            if running[0]:app.after(25,worker_beat)
        worker_beat()
        while app.job_busy and time.monotonic()<deadline:app.update();time.sleep(.005)
        app.update();worker_gaps.append(time.perf_counter()-last[0]);running[0]=False
        assert not app.job_busy,'Operation timed out';assert not errors,errors
    try:
        t=time.perf_counter();app.frames['import']._smart_load(args.workbook);settle();timings['import_seconds']=time.perf_counter()-t
        ds=app.academic_dataset;assert ds is not None and app.df is None
        for key in app.frames:app.show_frame(key);app.update()
        academic=app.frames['academic'];academic.next_button.invoke();app.update();assert academic.page==1
        academic.prev_button.invoke();assert academic.page==0
        academic.terms.set('SU22');academic.refresh_rows();assert all(c.term=='SU22' for c in academic.chosen_courses())
        academic.degrees.set(ds.degrees[0]);academic.cohorts.set(ds.cohorts[0]);academic.refresh_rows()
        academic.terms.set('All Terms');reg=ds.courses[0].registration;academic.search_var.set(reg)
        deadline=time.monotonic()+.4
        while time.monotonic()<deadline:app.update();time.sleep(.005)
        rows=academic.chosen_courses();assert rows and all(reg in c.registration for c in rows)
        from gazette_layout import term_order
        current=max({c.term for c in rows},key=term_order)
        app.frames['settings'].course_panel.select_term(current)
        app.show_frame('reports');panel=app.frames['reports'].course_panel
        for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
            target=root/('filtered'+ext);filedialog.asksaveasfilename=lambda **k:str(target)
            if target.exists():target.unlink()
            panel.format_buttons[['pdf','word','excel'].index(fmt)].invoke();settle()
            assert target.exists(),'No fresh export produced'
            if fmt=='pdf':
                with fitz.open(target) as d:
                    text=''.join(p.get_text() for p in d);assert ''.join(reg.split()) in ''.join(text.split()) and 'DRAFT' in text and 'CONTENTS' in text
            elif fmt=='word':
                d=Document(target);assert reg in '\n'.join(p.text for p in d.paragraphs)
            else:
                w=load_workbook(target);assert w['Course Results'].max_row==len(panel.rows())+1;w.close()
        app._autosave_path=lambda:str(root/'saved.lguexam')
        app.save_session();settle()
        import session_store
        restored=session_store.restore(root/'saved.lguexam')[1];assert restored.courses==ds.courses and restored.layouts==ds.layouts
        filedialog.askdirectory=lambda **k:str(root)
        t=time.perf_counter();panel.batch();settle();timings['filtered_batch_seconds']=time.perf_counter()-t
        folders=list(root.glob('LGU_DRAFT_*'));assert folders
        log=(sorted(folders)[-1]/'export-log.txt').read_text();assert 'FAILED' not in log and '30 files generated' in log
        max_gap=max((b-a for a,b in zip(beats,beats[1:])),default=0)
        result={'status':'PASS','dataset_counts':ds.counts(),'filtered_attempts':len(rows),'navigation_sections':len(app.frames),'gui_commands':['background import','all navigation','next/previous page','term/degree/cohort filters','debounced search','visible report PDF/Word/Excel button invocation','session save/reload','filtered batch (30 outputs)'],'timings':timings,'maximum_event_loop_heartbeat_gap_seconds':round(max_gap,3),'maximum_worker_heartbeat_gap_seconds':round(max(worker_gaps,default=0),3),'errors':errors}
        (root/'results.json').write_text(json.dumps(result,indent=2));print(json.dumps(result,indent=2))
    finally:
        for job in app.tk.call('after','info'):app.after_cancel(job)
        app.destroy()
if __name__=='__main__':main()
