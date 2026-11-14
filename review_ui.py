"""Scrollable source-aware review without changing imported cells."""
import os
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, Text
import customtkinter as ctk
from validation_scope import issue_source,relevant_errors


def open_review(parent,dataset,rows,blocked_programs=None):
    if dataset is None:return
    dialog=ctk.CTkToplevel(parent);dialog.title('Review data — source rows');dialog.geometry('1150x700')
    affected=set(relevant_errors(dataset,rows))
    count=dataset.counts()
    blocked_programs=blocked_programs or {}
    ctk.CTkLabel(dialog,text=f"{count['errors']} workbook errors | {len(affected)} affect selected students/history | "
                f"{count['warnings']} warnings | {sum(i.level=='INFO' for i in dataset.issues)} information notices",
                wraplength=1080).pack(padx=15,pady=10)
    if blocked_programs:
        summary='\n'.join(
            f'{program}: {reason}'
            for program, reason in blocked_programs.items())
        ctk.CTkLabel(dialog,text='Blocked programs and reasons:\n'+summary,
                    wraplength=1080,anchor='left',text_color='#8a4b08').pack(anchor='w',padx=15,pady=5)
    ctk.CTkLabel(dialog,text='Select a finding to see its original cells and source location. Open the source workbook, correct the indicated row, save it, and reimport. The application does not modify the source workbook. Information notices do not block exports.',wraplength=1080).pack(padx=15,pady=5)
    from calculation_review import calculation_review_issues
    math_issues=calculation_review_issues(dataset,rows)
    affected.update(math_issues)
    selected_mode=['Selection errors' if affected and blocked_programs else 'Calculation issues' if math_issues and not relevant_errors(dataset,rows) else 'Errors']
    holder=ttk.Frame(dialog);holder.pack(fill='both',expand=True,padx=15,pady=8)
    tree=ttk.Treeview(holder,columns=('scope','level','file','sheet','row','code','student'),show='headings',height=12)
    for key,title,width in [('scope','Affects selection',110),('level','Level',70),('file','Source file',150),('sheet','Sheet',80),('row','Excel row',75),('code','Finding',240),('student','Registration',210)]:
        tree.heading(key,text=title);tree.column(key,width=width)
    bar=ttk.Scrollbar(holder,orient='vertical',command=tree.yview);tree.configure(yscrollcommand=bar.set)
    bar.pack(side='right',fill='y');tree.pack(fill='both',expand=True)
    details=Text(dialog,height=10,font=("Consolas",10),wrap="word");details.pack(fill='x',padx=15,pady=8)
    ordered=sorted(dataset.issues+math_issues,key=lambda i:(i not in affected,{'ERROR':0,'WARNING':1,'INFO':2,'REVIEW':1}.get(i.level,3),i.source_file,i.sheet,i.excel_row))
    def refresh(value=None):
        if value is not None:selected_mode[0]=value
        tree.delete(*tree.get_children())
        for number,issue in enumerate(ordered):
            if selected_mode[0]=='Calculation issues' and issue not in math_issues:continue
            if selected_mode[0]=='Errors' and issue.level!='ERROR':continue
            if selected_mode[0]=='Selection errors' and issue not in affected:continue
            source,fields=issue_source(dataset,issue)
            tree.insert('', 'end',iid=str(number),values=('Yes' if issue in affected else '',issue.level,issue.source_file or dataset.source_filename,issue.sheet,issue.excel_row,issue.code.replace('_',' ').title(),fields.get('registration','')))
    def show(_=None):
        selected=tree.selection()
        if not selected:return
        issue=ordered[int(selected[0])];source,fields=issue_source(dataset,issue)
        text=f'Source workbook: {issue.source_file or dataset.source_filename}\nSheet: {issue.sheet} | Excel row: {issue.excel_row}\n\n'+issue.description+'\n\n'+'\n'.join(f'{key}: {value}' for key,value in (source or {}).get('cells',{}).items())
        details.delete('1.0','end');details.insert('1.0',text)
    def open_source():
        selected=tree.selection()
        if not selected:return
        issue=ordered[int(selected[0])]
        requested=Path(issue.source_file or dataset.source_filename).name.casefold()
        paths=[Path(path) for path in dataset.source_paths]
        source=next((path for path in paths if path.name.casefold()==requested),None)
        if source is None and len(paths)==1:
            source=paths[0]
        if source is None or not source.is_file():
            messagebox.showerror('Source workbook unavailable','The workbook associated with this finding is not available at its imported path. Use the source filename, sheet, and Excel row shown above to locate it.',parent=dialog)
            return
        try:
            os.startfile(str(source))
        except OSError as exc:
            messagebox.showerror('Could not open workbook',str(exc),parent=dialog)
    def export():
        path=filedialog.asksaveasfilename(parent=dialog,defaultextension='.xlsx',initialfile='LGU_Data_Review.xlsx')
        if not path:return
        from course_reports import export_report
        try:export_report('validation','excel',dataset,rows,path)
        except Exception as exc:messagebox.showerror('Review export',str(exc),parent=dialog)
    tree.bind('<<TreeviewSelect>>',show)
    actions=ctk.CTkFrame(dialog);actions.pack(fill='x',padx=15,pady=8)
    mode_menu=ctk.CTkOptionMenu(actions,values=['Errors','Selection errors','Calculation issues','All findings'],command=refresh)
    mode_menu.set(selected_mode[0])
    mode_menu.pack(side='left',padx=5)
    ctk.CTkButton(actions,text='Open source workbook to edit',command=open_source).pack(side='left',padx=5)
    ctk.CTkButton(actions,text='Export full review log',command=export).pack(side='left',padx=5)
    refresh()
    return dialog
