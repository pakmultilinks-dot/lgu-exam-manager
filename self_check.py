"""Opt-in packaged-app diagnostics using fictional records and temporary outputs only."""
from pathlib import Path
import json,time,tempfile,zipfile


def run_self_check(app_class,output):
    from tkinter import messagebox,filedialog
    from openpyxl import Workbook,load_workbook
    from course_reports import ROUTES,export_report
    from calculation_policy import AcademicIndex,calculation_attempts,display_ratio
    import session_store
    result={'status':'FAIL','build':'4.3','checks':[]};app=None;errors=[]
    originals=[(messagebox,key,getattr(messagebox,key)) for key in ('showinfo','showerror','showwarning')]
    originals += [(filedialog,key,getattr(filedialog,key)) for key in ('asksaveasfilename','askdirectory')]
    def passed(label):result['checks'].append(label)
    try:
        messagebox.showinfo=lambda *a,**kw:None
        messagebox.showerror=lambda *a,**kw:errors.append(str(a))
        messagebox.showwarning=lambda *a,**kw:errors.append(str(a))
        with tempfile.TemporaryDirectory(prefix='lgu-self-check-') as temporary:
            root=Path(temporary);source=root/'FICTIONAL-self-check.xlsx'
            book=Workbook();sheet=book.active
            sheet.append(['Registration Number','Student Name','Degree','Department','Term Code','Semester Number','Course Code','Course Name','Credit Hours','Grade','Fee Status'])
            for student in range(1,5):
                for term,number,courses in [('F25',1,[('CHK101','Foundations','B'),('CHK102','Communication','C')]),('S26',2,[('CHK101','Foundations','A'),('CHK201','Programming','B')])]:
                    for code,name,grade in courses:
                        sheet.append([f'Fa-2025/BSCS/TEST{student:03}',f'SELF CHECK STUDENT {student}','BSCS','Computer Science',term,number,code,name,3,grade,'YES' if term=='S26' and student<4 else 'NO'])
            book.save(source);book.close()
            app=app_class(restore_autosave=False)
            app._autosave_path=lambda:str(root/'DIAGNOSTIC-autosave.lguexam')
            app.auto_save_imports=True
            app.withdraw()
            def settle():
                deadline=time.monotonic()+180
                while app.job_busy and time.monotonic()<deadline:app.update();time.sleep(.005)
                app.update()
                if app.job_busy:raise RuntimeError('Background operation timed out.')
                if errors:raise RuntimeError('A desktop operation reported an error: '+errors[-1])
            app.frames['import']._smart_load(str(source));settle();dataset=app.academic_dataset
            assert len(dataset.courses)==16 and len(dataset.students)==4 and dataset.counts()['errors']==0
            assert dataset.grade_scale['F']=='0.00'
            from dataclasses import replace
            failed=replace(dataset,courses=[replace(dataset.courses[0],grade='F',quality_points='')])
            failed_points=calculation_attempts(failed)
            assert failed_points[0].quality_points=='0.00'
            failed_index=AcademicIndex(failed_points)
            assert display_ratio(failed_index.summary(failed_points[0].registration,failed_points[0].degree,'F25').exact_ratio)=='0.00'
            passed('background Excel import and supplied grade scale')
            assert session_store.restore(app._autosave_path())[1].courses==dataset.courses
            passed('successful import persists the active workspace automatically')
            for key in app.frames:app.show_frame(key);app.update()
            passed('all navigation sections')
            app.report_settings.update(gazette_term='S26',gazette_template='computer-science')
            first=dataset.courses[0];index=AcademicIndex(calculation_attempts(dataset))
            assert display_ratio(index.summary(first.registration,first.degree,'S26').exact_ratio)=='3.50'
            assert display_ratio(index.summary(first.registration,first.degree,'S26',True).exact_ratio)=='3.00'
            assert all(not c.quality_points for c in dataset.courses)
            passed('repeat selection; GPA 3.50; CGPA 3.00; raw cells unchanged')
            app.show_frame('reports');app.update();panel=app.frames['reports'].course_panel
            for fmt,extension in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
                target=root/('register'+extension)
                filedialog.asksaveasfilename=lambda **kw:str(target)
                panel.export('register',fmt);settle()
                assert target.stat().st_size>500
                if fmt=='pdf':assert target.read_bytes().startswith(b'%PDF-') and b'%%EOF' in target.read_bytes()[-100:]
                elif fmt=='word':
                    with zipfile.ZipFile(target) as archive:
                        xml=archive.read('word/document.xml')
                        assert archive.testzip() is None and b'3.00' in xml and b'FEE DEFAULTERS' in xml
                else:
                    saved=load_workbook(target);assert 'Point Derivation Audit' in saved.sheetnames;saved.close()
            passed('PDF, Word and Excel export commands and file readback')
            import course_ui
            original_open=course_ui.os.startfile;previews=[]
            try:
                course_ui.os.startfile=lambda path:previews.append(path)
                panel.preview();settle()
                assert len(previews)==1 and Path(previews[0]).read_bytes().startswith(b'%PDF-')
            finally:course_ui.os.startfile=original_open
            passed('PDF preview generation and viewer dispatch')
            from gazette_options import resolve_template_groups
            programs=('BSCS','BS IT','BS SE','BBA (Hons)')
            mixed=replace(dataset,courses=[replace(c,degree=programs[int(c.registration[-3:])-1]) for c in dataset.courses])
            groups=resolve_template_groups('auto',mixed.courses)
            assert len(groups)==4 and sum(len(group) for _,group in groups)==16
            for fmt,extension in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
                target=root/('automatic-mixed'+extension)
                export_report('register',fmt,mixed,mixed.courses,str(target),template_key='auto',settings=dict(app.report_settings))
                assert target.stat().st_size>500
            passed('automatic mixed-program routing and PDF, Word, Excel export')
            for program in ('BS BCH','BS ZOO','CERUF','M.Phil Urdu','MBA (2 Years)','MSBA'):
                historical=replace(dataset,courses=[replace(dataset.courses[-1],degree=program)],fee_records=[])
                assert len(resolve_template_groups('auto',historical.courses))==1
                target=root/('historical-'+program.replace(' ','').replace('.','')+'.pdf')
                export_report('register','pdf',historical,historical.courses,str(target),template_key='auto',settings=dict(app.report_settings))
                assert target.stat().st_size>500
            passed('six historical degree routes; source degree labels retained')
            from gazette_layout import _profile
            for program in ('BS Urdu','Ph. D URDU','M.Phil Urdu','CERUF'):
                profile=_profile('art-and-humanities','FALL 2025',program,True)
                labels=[item['text'] for item in profile['header']]
                assert 'DEPARTMENT OF URDU' in labels
                assert any('LAHORE GARRISON UNIVERSITY' in text for text in labels)
            passed('program-specific Urdu headings and reference university heading')
            count=0
            for key in ROUTES:
                for fmt,extension in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
                    target=root/(key+'-route'+extension)
                    export_report(key,fmt,dataset,dataset.courses,str(target),settings=dict(app.report_settings))
                    assert target.stat().st_size>100;count+=1
            passed(f'all {count} report dispatcher outputs')
            filedialog.askdirectory=lambda **kw:str(root)
            panel.batch();settle()
            logs=list(root.glob('LGU_DRAFT_*/export-log.txt'));assert len(logs)==1
            log=logs[0].read_text();assert 'FAILED' not in log and 'failed=0' in log and f'{count} files generated' in log
            passed('background batch export; zero failures')
            session=root/'snapshot.lguexam'
            session_store.save(session,'course',None,None,{},dict(app.report_settings),dataset)
            loaded=session_store.restore(session)[1]
            session_store.save(session,'course',None,None,{},dict(app.report_settings),loaded)
            assert session_store.restore(session)[1].courses==dataset.courses
            passed('dataset-only session save, reload and resave')
            from diagnostic_workbooks import check_unseen_workbooks
            check_unseen_workbooks(app,root,settle,passed)
            saved_courses=list(app.academic_dataset.courses)
            saved_bytes=Path(app._autosave_path()).read_bytes()
            app.frames['import'].clear_button.invoke();app.update()
            assert app.academic_dataset is None and app.raw_df is None and app.source_path is None
            assert Path(app._autosave_path()).read_bytes()==saved_bytes
            passed('Clear current data empties records without deleting the saved workspace')
            app.frames['import'].path='obsolete-source.xlsx'
            app.frames['import'].open_saved_button.invoke();settle()
            assert app.academic_dataset.courses==saved_courses and app.frames['import'].path is None
            passed('Open saved workspace requires explicit action and restores the saved records')
            # Even with a saved workspace on disk, a normal launch must stay empty.
            previous_path=app_class._autosave_path
            app_class._autosave_path=lambda self:str(root/'DIAGNOSTIC-autosave.lguexam')
            fresh=None
            try:
                fresh=app_class();fresh.withdraw();fresh.update()
                assert fresh.academic_dataset is None and fresh.raw_df is None and fresh.df is None
                assert fresh.source_path is None and not fresh.job_busy
                passed('normal startup remains empty with an existing saved workspace')
            finally:
                if fresh is not None:
                    for job in fresh.tk.call('after','info'):fresh.tk.call('after','cancel',job)
                    fresh.destroy()
                app_class._autosave_path=previous_path
            result['status']='PASS'
    except Exception as exc:
        result['error_type']=type(exc).__name__;result['error']=str(exc) or 'No exception message supplied.'
    finally:
        if app is not None:
            try:
                for job in app.tk.call('after','info'):app.after_cancel(job)
                app.destroy()
            except Exception as exc:
                result['status']='FAIL';result['cleanup_error']=type(exc).__name__+': '+str(exc)
        for module,key,value in originals:setattr(module,key,value)
        path=Path(output);path.parent.mkdir(parents=True,exist_ok=True)
        path.write_text(json.dumps(result,indent=2),encoding='utf-8')
    return 0 if result['status']=='PASS' else 1
