from academic_identity import program_key,student_count
"""Course-aware Report Center, sharing Academic Records filters."""
import os
import time
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, Text
import customtkinter as ctk
from course_reports import ROUTES, capability, export_gazette_pack, export_gazette_pack_all_terms, export_report
from gazette_options import (TEMPLATE_LABELS,available_programs,available_template_labels,
                             gazette_program_statuses,
                             gazette_route_summary,selection_description,current_term,gazette_rows,
                             publication_terms,publication_term_label)

def open_pdf_preview(path):
    if not Path(path).is_file():raise ValueError('The generated PDF preview is missing.')
    # os.startfile exists only on Windows; a missing PDF association (or a
    # missing startfile) must show the notice, never fail the preview.
    try:os.startfile(path)
    except (OSError,AttributeError):
        messagebox.showinfo('PDF preview saved',f'The PDF was generated, but Windows could not open a PDF reader.\n\nOpen this file in a PDF reader or browser:\n{path}')


def operation_available(app):
    if getattr(app,'job_busy',False):
        messagebox.showinfo('Operation in progress','Wait for the current operation to finish.')
        return False
    return True


def dataset_snapshot(dataset):
    from dataclasses import replace
    return replace(dataset,grade_scale=dict(dataset.grade_scale),fee_records=list(dataset.fee_records)) if dataset else None


class GazetteSettingsPanel(ctk.CTkScrollableFrame):
    def __init__(self,parent,app):
        super().__init__(parent,fg_color='white',corner_radius=12)
        self.app=app
        ctk.CTkLabel(self,text='Gazette settings',font=('Arial',18,'bold')).pack(anchor='w',padx=24,pady=(20,10))
        self.layout_info=ctk.CTkLabel(
            self,
            text='Gazette layouts are matched in Reports to export-ready programs in the selected cycle. Spring includes same-year Summer; Summer is not a separate publication cycle.',
            wraplength=850,justify='left',
        )
        self.layout_info.pack(anchor='w',padx=24,pady=10)
        ctk.CTkLabel(self,text='Gazette publication cycle (Spring includes Summer)',font=('Arial',13,'bold')).pack(anchor='w',padx=24)
        self.term_menu=ctk.CTkOptionMenu(self,values=['Import results first'],width=360,command=self.select_term)
        self.term_menu.pack(anchor='w',padx=24,pady=10)
        self.cover_entries={}
        self.cover_vars={}
        for key,label in [('cover_department','Cover department (blank = detected)'),('cover_period','Cover period (blank = selected publication cycle)'),('cover_title','Cover title (blank = reference title)'),('cover_prepared_by','Prepared by (blank = reference wording)')]:
            ctk.CTkLabel(self,text=label).pack(anchor='w',padx=24,pady=(4,0))
            variable=ctk.StringVar(value=app.report_settings.get(key,''))
            self.cover_vars[key]=variable
            variable.trace_add('write',lambda *_args,k=key,v=variable:app.report_settings.update({k:v.get().strip()}))
            entry=ctk.CTkEntry(self,width=580,textvariable=variable)
            entry.pack(anchor='w',padx=24,pady=3)
            self.cover_entries[key]=entry
        # CGPA probation cutoff. This is the line the gazette prints in
        # COMMENTS: at or above it is G.A.ST, below it is PROBATION. Blank
        # means 2.00, the usual line. It was unreachable with a dataset
        # loaded (the field lived only on the legacy Settings panel).
        ctk.CTkLabel(self,text='CGPA probation cutoff (blank = 2.00)',font=('Arial',13,'bold')).pack(anchor='w',padx=24,pady=(12,0))
        ctk.CTkLabel(self,text='CGPA at or above this line is G.A.ST, below it is PROBATION in the gazette. Blank uses 2.00.',wraplength=850,justify='left').pack(anchor='w',padx=24)
        self.probation_var=ctk.StringVar(value=str(app.report_settings.get('probation_cgpa','')))
        self.probation_var.trace_add('write',lambda *_args:app.report_settings.update({'probation_cgpa':self.probation_var.get().strip()}))
        self.probation_entry=ctk.CTkEntry(self,width=180,textvariable=self.probation_var)
        self.probation_entry.pack(anchor='w',padx=24,pady=3)
        self.detected=ctk.CTkLabel(self,text='',wraplength=850,justify='left')
        self.detected.pack(anchor='w',padx=24,pady=18)
        ctk.CTkLabel(self,text='Grade points for files with no numeric GP column',font=('Arial',14,'bold')).pack(anchor='w',padx=24,pady=(15,4))
        ctk.CTkLabel(self,text='Your supplied LGU scale is loaded automatically. You can edit GRADE=value entries below. Source GP values take priority. Grades absent from the scale and non-final statuses need an explicit rule.',wraplength=800,justify='left').pack(anchor='w',padx=24)
        self.scale_text=Text(self,width=72,height=7,font=("Arial",11),relief="solid",borderwidth=1)
        self.scale_text.pack(anchor='w',padx=24,pady=5)
        self.scale_reference=ctk.CTkEntry(self,width=580,placeholder_text='Scale source / approval reference')
        self.scale_reference.pack(anchor='w',padx=24,pady=5)
        ctk.CTkButton(self,text='Apply entered scale',command=self.apply_scale).pack(anchor='w',padx=24,pady=6)
        ctk.CTkButton(self,text='Restore supplied LGU scale',command=self.restore_scale).pack(anchor='w',padx=24,pady=6)
        self.coverage_label=ctk.CTkLabel(self,text='',wraplength=800,justify='left')
        self.coverage_label.pack(anchor='w',padx=24,pady=8)
        ctk.CTkButton(self,text='Import fee-defaulter workbook',command=self.import_fees).pack(anchor='w',padx=24,pady=6)
        ctk.CTkLabel(self,text='GPA display: 2 decimal places, rounded half-up.\nCalculations keep their full precision; Excel retains exact totals and repeat selection.\nGazette pages follow the selected department and publication cycle.',wraplength=850,justify='left').pack(anchor='w',padx=24,pady=18)
        ctk.CTkButton(self,text='Open gazette reports',command=lambda:app.show_frame('reports')).pack(anchor='w',padx=24,pady=20)
    def select_term(self,value):
        ds=self.app.academic_dataset
        if ds:
            value=next(
                (term for term in publication_terms(ds)
                 if publication_term_label(term,ds)==value),
                value,
            )
        self.app.report_settings['gazette_term']=value
        self.refresh()
    def refresh(self):
        self.detected.configure(text=selection_description(
            self.app.academic_dataset,self.app.report_settings))
        ds=self.app.academic_dataset
        self.load_scale(ds)
        if ds:
            terms=publication_terms(ds)
            self.term_menu.configure(values=[
                publication_term_label(term,ds) for term in terms
            ] or ['No Fall or Spring cycles found'])
            try:term=current_term(ds,self.app.report_settings)
            except ValueError:self.term_menu.set('Choose a valid cycle')
            else:self.term_menu.set(publication_term_label(term,ds))
        for key,variable in self.cover_vars.items():
            variable.set(self.app.report_settings.get(key,''))
    def apply_scale(self):
        if not operation_available(self.app):return
        from calculation_policy import parse_grade_scale
        ds=self.app.academic_dataset
        if ds is None:return
        try:scale=parse_grade_scale(self.scale_text.get('1.0','end'),self.scale_reference.get())
        except ValueError as exc:messagebox.showwarning('Grade scale',str(exc));return
        ds.grade_scale=scale;ds.grade_scale_reference=self.scale_reference.get().strip();ds.grade_scale_user_set=True
        self.load_scale(ds)
        messagebox.showinfo('Grade scale saved','Applied to calculations with missing source GP only. Save the session to retain this scale. Original source cells remain unchanged.')
    def restore_scale(self):
        if not operation_available(self.app):return
        ds=self.app.academic_dataset
        if ds is None:return
        from grading_profiles import default_grade_scale,default_grade_reference
        ds.grade_scale=default_grade_scale();ds.grade_scale_reference=default_grade_reference();ds.grade_scale_user_set=False
        self.load_scale(ds)
    def load_scale(self,ds):
        self.scale_text.delete('1.0','end')
        self.scale_text.insert('1.0','\n'.join(f'{grade}={rate}' for grade,rate in (ds.grade_scale if ds else {}).items()))
        self.scale_reference.delete(0,'end');self.scale_reference.insert(0,ds.grade_scale_reference if ds else '')
        if ds:
            from grading_profiles import scale_coverage
            coverage=scale_coverage(ds)
            self.coverage_label.configure(text=f"{coverage['source']} attempts have source points; {coverage['source_rate']} have source points per credit; {coverage['derived']-coverage['source_rate']} use the grade scale. Missing grade rates: {', '.join(coverage['unmapped']) or 'none'}. Non-final/incomplete attempts: {sum(coverage['nonfinal'].values())}.")
    def import_fees(self):
        if not operation_available(self.app):return
        ds=self.app.academic_dataset
        if ds is None:return
        path=filedialog.askopenfilename(title='Import fee status / defaulter list',filetypes=[('Excel','*.xlsx *.xlsm *.xls *.xlsb')])
        if not path:return
        from academic_engine import load_fee_workbook
        def done(records):
            ds.fee_records=list(dict.fromkeys(ds.fee_records+records))
            messagebox.showinfo('Fee list imported',f'{student_count(records)} fee registrations imported. Save the session to retain them.')
        self.app.run_job('Reading fee list',lambda:load_fee_workbook(path,ds),done)

class CourseReportPanel(ctk.CTkFrame):
    def __init__(self,parent,app):
        super().__init__(parent,fg_color='#f5f7fa'); self.app=app
        ctk.CTkLabel(self,text='Gazette & Reports',font=('Arial',24,'bold'),text_color='#0d542b').pack(anchor='w',padx=24,pady=14)
        ctk.CTkLabel(self,text='Choose a publication term and a program with results in that term. The selected program’s gazette includes its full history through the selected term.',wraplength=850).pack(anchor='w',padx=24)
        self.selection=ctk.CTkLabel(self,text='',font=('Arial',13,'bold'));self.selection.pack(anchor='w',padx=24,pady=12)
        period=ctk.CTkFrame(self,fg_color='transparent');period.pack(anchor='w',padx=24)
        ctk.CTkLabel(period,text='Gazette cycle (Spring includes Summer)').pack(side='left',padx=(0,10))
        self.term_menu=ctk.CTkOptionMenu(period,values=['Import results first'],command=self.select_term,width=250);self.term_menu.pack(side='left')
        ctk.CTkButton(period,text='Show all imported records',command=self.show_all_records,width=200).pack(side='left',padx=12)
        card=ctk.CTkFrame(self,fg_color='white',corner_radius=12);card.pack(fill='x',padx=24,pady=12)
        ctk.CTkLabel(card,text='Choose a report',font=('Arial',16,'bold')).pack(anchor='w',padx=20,pady=(18,8))
        self.route_menu=ctk.CTkOptionMenu(card,values=list(ROUTES.values()),width=580,fg_color='#0d542b',command=lambda _:self.refresh())
        self.route_menu.pack(anchor='w',padx=20,pady=4)
        self.template_label=ctk.CTkLabel(card,text='Matched gazette layout',text_color='#526174')
        self.template_label.pack(anchor='w',padx=20,pady=(5,0))
        self.template_menu=ctk.CTkOptionMenu(card,values=['No matched gazette'],width=360,command=self.set_template)
        self.template_menu.pack(anchor='w',padx=20,pady=6)
        self.program_menu=ctk.CTkOptionMenu(card,values=['All programs / degrees'],width=360,command=lambda _:self.refresh())
        self.program_menu.pack(anchor='w',padx=20,pady=6)
        self.route_info=ctk.CTkLabel(card,text='',wraplength=760,justify='left',font=('Arial',12,'bold'),text_color='#0d542b')
        self.route_info.pack(anchor='w',padx=20,pady=(2,4))
        self.blocked_info=ctk.CTkLabel(
            card,text='',wraplength=760,justify='left',text_color='#8a4b08')
        self.reason=ctk.CTkLabel(card,text='',wraplength=760,justify='left');self.reason.pack(anchor='w',padx=20,pady=10)
        self.review_blocked_button=ctk.CTkButton(
            card,text='Review blocked programs',command=self.review_blocked,
            width=190,fg_color='#8a4b08',hover_color='#6e3b06',
        )
        bar=ctk.CTkFrame(card,fg_color='transparent');bar.pack(anchor='w',padx=16,pady=(0,20))
        self.format_buttons=[]
        for fmt,color in [('pdf','#a33333'),('word','#245f98'),('excel','#237742')]:
            b=ctk.CTkButton(bar,text='Export '+fmt.title(),width=135,fg_color=color,command=lambda f=fmt:self.export(self.selected_key(),f));b.pack(side='left',padx=4);self.format_buttons.append(b)
        self.preview_button=ctk.CTkButton(bar,text='PDF preview',width=135,command=self.preview);self.preview_button.pack(side='left',padx=4)
        actions=ctk.CTkFrame(self,fg_color='transparent');actions.pack(fill='x',padx=20,pady=10)
        ctk.CTkButton(actions,text='Change filters / search',command=lambda:app.show_frame('academic')).pack(side='left',padx=4)
        ctk.CTkButton(actions,text='Review data',width=110,command=self.review_selection).pack(side='left',padx=4)
        ctk.CTkButton(actions,text='Batch available reports',command=self.batch).pack(side='left',padx=4)
        ctk.CTkButton(actions,text='Department gazette pack',command=self.department_pack).pack(side='left',padx=4)
        ctk.CTkButton(actions,text='All terms gazette pack',command=self.department_pack_all_terms).pack(side='left',padx=4)
        ctk.CTkButton(actions,text='Cancel batch',command=self.cancel).pack(side='left',padx=4)
        self.status=ctk.CTkLabel(self,text='',wraplength=800);self.status.pack(anchor='w',padx=24,pady=8)
        self.details=ctk.CTkLabel(self,text='',wraplength=800,text_color='#526174',justify='left')
        self.details.pack(anchor='w',padx=24,pady=12)
        self.stop=threading.Event()
    def select_term(self,term):
        ds=self.app.academic_dataset
        if ds:
            term=next(
                (code for code in publication_terms(ds)
                 if publication_term_label(code,ds)==term),
                term,
            )
        self.app.report_settings['gazette_term']=term;self.refresh()
    def show_all_records(self):
        academic=self.app.frames['academic']
        academic.terms.set('All Terms');academic.degrees.set('All Degrees');academic.cohorts.set('All Cohorts');academic.search_var.set('');academic.reset_page()
        self.app.show_frame('academic')
    def review_selection(self):
        from review_ui import open_review
        try:rows=self.rows()
        except ValueError:rows=[]
        open_review(self,self.app.academic_dataset,rows)
    def review_blocked(self):
        from review_ui import open_review
        ds=self.app.academic_dataset
        if ds is None:return
        statuses=getattr(self,'_blocked_program_statuses',{})
        blocked_rows=[
            course
            for _,(_,program_rows) in statuses.items()
            for course in program_rows
        ]
        if not blocked_rows:
            blocked_rows=list(ds.courses)
        open_review(
            self,ds,blocked_rows,
            blocked_programs={
                name:reason
                for name,reason in getattr(self,'_blocked_programs',{}).items()
            },
        )
    def update_template_choices(self):
        if self.selected_key()!='register':
            self.template_label.pack_forget()
            self.template_menu.pack_forget()
            return
        ds=self.app.academic_dataset
        program=self.program_menu.get()
        if program not in ('All programs','All programs / degrees'):
            try:term=current_term(ds,self.app.report_settings)
            except ValueError:rows=[]
            else:
                cycle_terms={term.upper()}
                if term.upper().startswith('S') and not term.upper().startswith('SU'):
                    cycle_terms.add('SU'+term.upper()[1:])
                rows=[course for course in ds.courses
                      if course.term.upper() in cycle_terms
                      and program_key(course.degree)==program_key(program)] if ds else []
            labels=available_template_labels(rows)
            if not self.template_label.winfo_manager():
                self.template_label.pack(anchor='w',padx=20,pady=(5,0),before=self.program_menu)
                self.template_menu.pack(anchor='w',padx=20,pady=6,before=self.program_menu)
            if not labels:
                self.app.report_settings['gazette_template']='auto'
                self.template_menu.configure(values=['No matching gazette layout'],state='disabled')
                self.template_menu.set('No matching gazette layout')
                return
            current=self.app.report_settings.get('gazette_template','auto')
            selected=next((label for label,key in TEMPLATE_LABELS.items()
                           if key==current and label in labels),None)
            if selected is None and len(labels)==1:
                selected=labels[0]
                self.app.report_settings['gazette_template']=TEMPLATE_LABELS[selected]
            if selected is None:
                placeholder='Choose a matched gazette layout'
                self.template_menu.configure(values=[placeholder]+labels,state='normal')
                self.template_menu.set(placeholder)
            else:
                self.template_menu.configure(values=labels,state='disabled' if len(labels)==1 else 'normal')
                self.template_menu.set(selected)
            return
        self.template_label.pack_forget()
        self.template_menu.pack_forget()
        self.app.report_settings['gazette_template']='auto'
    def set_template(self,label):
        if label not in TEMPLATE_LABELS or label=='Automatic from imported data':
            return
        self.app.report_settings['gazette_template']=TEMPLATE_LABELS[label]
        self.refresh()
    def selected_key(self):
        label=self.route_menu.get()
        return next(key for key,value in ROUTES.items() if value==label)
    def _ready_program_keys(self):
        """Program keys that are export-ready in the current cycle, or None if not computed yet."""
        ready=getattr(self,'_ready_programs',None)
        if ready is None:
            return None
        return {program_key(name) for name in ready}

    def rows(self):
        if self.selected_key()=='register' and self.app.academic_dataset is not None:
            rows=list(self.app.academic_dataset.courses)
        else:
            rows=list(self.app.frames['academic'].chosen_courses(ignore_term=self.selected_key() in ('register','cgpa_list')))
        program=self.program_menu.get()
        if program not in ('All programs','All programs / degrees'):
            rows=[course for course in rows
                  if program_key(course.degree)==program_key(program)]
        elif self.selected_key()=='register' and self.app.academic_dataset is not None:
            # "All programs" exports only the programs that are ready; blocked
            # programs are never exported and never block the rest.
            ready=self._ready_program_keys()
            if ready is not None:
                rows=[course for course in rows if program_key(course.degree) in ready]
        if self.selected_key()=='register' and self.app.academic_dataset is not None:
            return gazette_rows(self.app.academic_dataset,rows,self.app.report_settings)
        return rows
    def refresh(self):
        ds=self.app.academic_dataset
        if ds:
            terms=publication_terms(ds)
            self.term_menu.configure(values=[
                publication_term_label(term,ds) for term in terms
            ] or ['No Fall or Spring cycles found'])
            try:
                term=current_term(ds,self.app.report_settings)
                self.term_menu.set(publication_term_label(term,ds))
            except ValueError:
                self.term_menu.set('Choose a valid cycle')
            try:
                if self.selected_key()=='register':
                    statuses=gazette_program_statuses(ds,self.app.report_settings)
                    self._ready_programs=[
                        name for name,(ready,_,_) in statuses.items() if ready
                    ]
                    self._blocked_programs={
                        name:reason for name,(ready,reason,_) in statuses.items()
                        if not ready
                    }
                    self._blocked_program_statuses={
                        name:(reason,program_rows)
                        for name,(ready,reason,program_rows) in statuses.items()
                        if not ready
                    }
                    # Only export-ready programs appear in the list; blocked
                    # programs stay hidden and never block an "All" export.
                    values=['All programs / degrees']+self._ready_programs
                    if not statuses:values=['No programs in this publication cycle']
                else:
                    self._ready_programs=None
                    self._blocked_programs={}
                    self._blocked_program_statuses={}
                    values=['All programs / degrees']+available_programs(
                        ds,self.app.report_settings)
            except ValueError:
                self._ready_programs=None
                self._blocked_programs={}
                self._blocked_program_statuses={}
                values=['All programs / degrees']
            old=self.program_menu.get();self.program_menu.configure(values=values)
            self.program_menu.configure(
                state='disabled' if values==['No programs in this publication cycle'] else 'normal')
            default=('All programs / degrees' if 'All programs / degrees' in values else
                     next((value for value in values
                           if value!='No programs in this publication cycle'),
                          'No programs in this publication cycle'))
            self.program_menu.set(old if old in values else default)
            blocked=self._blocked_programs
            if blocked:
                self.blocked_info.configure(
                    text=f'{len(statuses)} program(s) detected in this publication cycle. '
                         f'{len(self._ready_programs or [])} ready to export; '
                         f'{len(blocked)} {"needs" if len(blocked)==1 else "need"} source-data or routing review: '
                         f'{", ".join(blocked)}. Blocked programs are hidden from this list; '
                         'choose Review blocked programs to see the exact workbook rows.'
                )
                self.blocked_info.pack(anchor='w',padx=20,pady=(2,6),before=self.reason)
                self.review_blocked_button.configure(
                    text=f'Review blocked programs ({len(blocked)})')
                self.review_blocked_button.pack(anchor='w',padx=20,pady=(0,8),before=self.reason)
            else:
                self.blocked_info.pack_forget()
                self.review_blocked_button.pack_forget()
        try:rows=self.rows()
        except ValueError as exc:
            self.update_template_choices()
            self.route_info.configure(text='')
            self.reason.configure(text=str(exc));self.selection.configure(text='No records in the selected publication cycle')
            for button in self.format_buttons+[self.preview_button]:button.configure(state='disabled')
            return
        self.update_template_choices()
        term=current_term(ds,self.app.report_settings) if ds and self.selected_key()=='register' else 'Selected records'
        if self.selected_key()=='register':
            from gazette_layout import term_order
            scope_terms=sorted({course.term for course in rows},key=term_order)
            scope=f'{len(scope_terms)} semesters' if len(scope_terms)>1 else (scope_terms[0] if scope_terms else term)
            summary=f'{ds.source_filename if ds else "No workbook"} | {len(ds.courses) if ds else 0} imported attempts\nGazette cycle {publication_term_label(term,ds) if ds else term}: {len(rows)} course rows | {student_count(rows)} students | {scope}'
        else:
            summary=f'{ds.source_filename if ds else "No workbook"} | {len(ds.courses) if ds else 0} imported attempts\nSelected records: {len(rows)} attempts | {student_count(rows)} students'
        self.selection.configure(text=summary)
        if self.selected_key()=='cgpa_list':
            self.selection.configure(text=f'{student_count(rows)} selected students | all imported semesters')
        ok,reason=capability(self.selected_key(),ds,rows)
        if self.selected_key()=='register':
            try:
                route_text=gazette_route_summary(
                    rows,self.app.report_settings.get('gazette_template','auto'))
                cycle_terms={term.upper()}
                if term.upper().startswith('S') and not term.upper().startswith('SU'):
                    cycle_terms.add('SU'+term.upper()[1:])
                ready_keys=self._ready_program_keys()
                all_programs=self.program_menu.get() in ('All programs','All programs / degrees')
                cycle_rows=[
                    course for course in ds.courses
                    if course.term.upper() in cycle_terms
                    and (
                        (all_programs and (ready_keys is None or program_key(course.degree) in ready_keys))
                        or program_key(course.degree)==program_key(self.program_menu.get())
                    )
                ]
                current_rows=[course for course in cycle_rows if course.term.upper()==term.upper()]
                summer_rows=[course for course in cycle_rows
                             if course.term.upper()=='SU'+term.upper()[1:]]
                program=self.program_menu.get()
                cycle_counts=f'{len(current_rows)} course attempts in {term}'
                if summer_rows:
                    cycle_counts+=f' + {len(summer_rows)} in SU{term[1:]}'
                if program in ('All programs','All programs / degrees'):
                    route_text+=(
                        f'\n{cycle_counts} across '
                        f'{len({program_key(course.degree) for course in cycle_rows})} programs; '
                        f'{len(rows)} full-history attempts included through {term}.'
                    )
                else:
                    route_text+=(
                        f'\n{program}: {cycle_counts}; '
                        f'{len(rows)} full-history attempts included through {term}.'
                    )
                self.route_info.configure(text=route_text)
            except ValueError as exc:
                self.route_info.configure(text='Gazette route needs review: '+str(exc))
                if ok:ok=False;reason=str(exc)
            if not ok:
                reason+=(
                    f' The publication includes {term} results and their earlier history; '
                    'some review findings may be from those earlier terms.'
                )
        else:
            self.route_info.configure(text='')
        if ok and self.selected_key() in ('register','cgpa_list'):
            from grading_profiles import scale_coverage
            from course_reports import expanded_history
            coverage=scale_coverage(ds,expanded_history(ds,rows))
            coverage_text=(f"Grade points: {coverage['source']} from the workbook; "
                           f"{coverage['derived']-coverage['source_rate']} from the LGU grade scale.")
            if coverage['unmapped']:
                coverage_text+=' Unmapped grades need review: '+', '.join(sorted(coverage['unmapped']))+'.'
            self.details.configure(text=coverage_text+' CGPA uses imported history through the selected term.')
        if ok and ds and self.selected_key()=='register':
            from calculation_review import calculation_review_issues
            reviews=calculation_review_issues(ds,rows)
            if reviews:
                self.details.configure(text=self.details.cget('text')+f' {len(reviews)} GPA/CGPA results need review; use Review data for affected rows.')
        self.reason.configure(text=reason)
        if not (ok and self.selected_key() in ('register','cgpa_list')):
            self.details.configure(text='')
        for b in self.format_buttons+[self.preview_button]:b.configure(state='normal' if ok else 'disabled')
    def export(self,key,fmt):
        if not operation_available(self.app):return
        ds=dataset_snapshot(self.app.academic_dataset)
        try:rows=self.rows()
        except ValueError as exc:messagebox.showwarning('Selection',str(exc));return
        ok,reason=capability(key,ds,rows)
        if not ok:messagebox.showwarning('Unavailable',reason);return
        ext={'pdf':'.pdf','word':'.docx','excel':'.xlsx'}[fmt]
        path=filedialog.asksaveasfilename(defaultextension=ext,initialfile='LGU_Gazette'+ext if key=='register' else 'LGU_'+key+ext,filetypes=[(fmt,'*'+ext)])
        if not path:return
        template_key=self.app.report_settings.get('gazette_template','auto')
        settings=dict(self.app.report_settings)
        self.status.configure(text='Generating '+key+' in background...')
        def done(_):
            self.status.configure(text='Saved '+Path(path).name)
            messagebox.showinfo('Export saved',path)
        self.app.run_job('Exporting '+key,lambda:export_report(key,fmt,ds,rows,path,template_key=template_key,settings=settings),done)
    def preview(self):
        if not operation_available(self.app):return
        ds=dataset_snapshot(self.app.academic_dataset)
        try:rows=self.rows()
        except ValueError as exc:messagebox.showwarning('Selection',str(exc));return
        selected_key=self.selected_key()
        import tempfile
        path=str(Path(tempfile.mkdtemp(prefix='lgu-preview-'))/'DRAFT-preview.pdf')
        template_key=self.app.report_settings.get('gazette_template','auto')
        settings=dict(self.app.report_settings)
        self.app.run_job('Preparing PDF preview',lambda:export_report(selected_key,'pdf',ds,rows,path,template_key=template_key,settings=settings),lambda _:open_pdf_preview(path))
    def cancel(self):
        self.stop.set();self.status.configure(text='Cancellation requested; current file will finish safely.')
    def department_pack(self):
        if not operation_available(self.app):return
        base=filedialog.askdirectory(title='Choose department gazette destination')
        if not base:return
        ds=dataset_snapshot(self.app.academic_dataset)
        if ds is None:
            messagebox.showwarning('Gazette pack','Import a course workbook first.')
            return
        rows=list(self.app.academic_dataset.courses)
        program=self.program_menu.get()
        if program not in ('All programs','All programs / degrees'):
            rows=[course for course in rows
                  if program_key(course.degree)==program_key(program)]
        else:
            ready=self._ready_program_keys()
            if ready is not None:
                rows=[course for course in rows if program_key(course.degree) in ready]
        folder=Path(base)/(time.strftime('LGU_DEPARTMENT_GAZETTES_%Y%m%d_%H%M%S'))
        settings=dict(self.app.report_settings)
        self.status.configure(text='Generating department gazettes in background...')
        def work():
            return export_gazette_pack(ds,rows,folder,settings=settings)
        def done(written):
            skipped_note=folder/'skipped-programs.txt'
            suffix='\n\nSome programs were skipped; see skipped-programs.txt in the folder.' if skipped_note.exists() else ''
            self.status.configure(text=f'{len(written)} department gazette files saved in {folder}'+(' (some programs skipped)' if skipped_note.exists() else ''))
            messagebox.showinfo('Department gazette pack',f'{len(written)} files saved in:\n{folder}{suffix}')
        self.app.run_job('Department gazette pack',work,done)
    def department_pack_all_terms(self):
        if not operation_available(self.app):return
        base=filedialog.askdirectory(title='Choose all-terms gazette destination')
        if not base:return
        ds=dataset_snapshot(self.app.academic_dataset)
        if ds is None:
            messagebox.showwarning('Gazette pack','Import a course workbook first.')
            return
        rows=list(self.app.academic_dataset.courses)
        folder=Path(base)/(time.strftime('LGU_ALL_TERMS_GAZETTES_%Y%m%d_%H%M%S'))
        settings=dict(self.app.report_settings)
        self.status.configure(text='Generating a gazette for every term in background...')
        def work():
            return export_gazette_pack_all_terms(ds,rows,folder,settings=settings)
        def done(written):
            skipped_note=folder/'skipped-programs.txt'
            suffix='\n\nSome programs were skipped; see skipped-programs.txt in the folder.' if skipped_note.exists() else ''
            self.status.configure(text=f'{len(written)} gazette files saved in {folder} (one folder per term)'+(' (some programs skipped)' if skipped_note.exists() else ''))
            messagebox.showinfo('All terms gazette pack',f'{len(written)} files saved in:\n{folder}\n\nOne folder per term; see pack-summary.txt in the folder.{suffix}')
        self.app.run_job('All terms gazette pack',work,done)
    def batch(self):
        if not operation_available(self.app):return
        base=filedialog.askdirectory(title='Choose private report destination')
        if not base:return
        ds=dataset_snapshot(self.app.academic_dataset)
        academic=self.app.frames['academic']
        filtered_rows=list(academic.chosen_courses())
        history_rows=list(ds.courses)
        program=self.program_menu.get() if hasattr(self,'program_menu') else 'All programs'
        if program not in ('All programs','All programs / degrees'):
            filtered_rows=[course for course in filtered_rows
                           if program_key(course.degree)==program_key(program)]
            history_rows=[course for course in history_rows
                          if program_key(course.degree)==program_key(program)]
        self.stop.clear()
        template_key=self.app.report_settings.get('gazette_template','auto')
        settings=dict(self.app.report_settings)
        import uuid
        folder=Path(base)/(time.strftime('LGU_DRAFT_%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:8])
        def work():
            folder.mkdir(exist_ok=False);log=[];ok_n=0;failed_n=0;skipped_n=0
            for key in ROUTES:
                rows=history_rows if key in ('register','cgpa_list') else filtered_rows
                if key=='register':
                    if program in ('All programs','All programs / degrees'):
                        ready=self._ready_program_keys()
                        if ready is not None:
                            rows=[course for course in rows if program_key(course.degree) in ready]
                    try:rows=gazette_rows(ds,rows,settings)
                    except ValueError as exc:
                        skipped_n+=1;log.append(f'SKIPPED {key}: {exc}');continue
                ok,reason=capability(key,ds,rows)
                if not ok:
                    skipped_n+=1;log.append(f'SKIPPED {key}: {reason}');continue
                for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
                    if self.stop.is_set():break
                    try:
                        export_report(key,fmt,ds,rows,str(folder/(key+ext)),template_key=template_key,settings=settings);ok_n+=1;log.append(f'OK {key}{ext}')
                    except Exception as exc:
                        failed_n+=1;log.append(f'FAILED {key}{ext}: {type(exc).__name__}: {exc}')
                if self.stop.is_set():break
            log.append(f'{ok_n} files generated; failed={failed_n}; skipped={skipped_n}; cancelled={self.stop.is_set()}')
            (folder/'export-log.txt').write_text('\n'.join(log),encoding='utf-8')
            return log[-1]+'; '+str(folder)
        self.status.configure(text='Batch running in background; navigation stays available.')
        self.app.run_job('Batch export',work,lambda result:self.status.configure(text=result))
