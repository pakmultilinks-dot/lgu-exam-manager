"""Scrollable source-aware review without changing imported cells."""
from tkinter import ttk, filedialog, messagebox, Text
import customtkinter as ctk
from validation_scope import issue_source,relevant_errors


def open_review(parent,dataset,rows):
    if dataset is None:return
    dialog=ctk.CTkToplevel(parent);dialog.title('Review data — source rows');dialog.geometry('1150x700')
    affected=set(relevant_errors(dataset,rows))
    count=dataset.counts()
    ctk.CTkLabel(dialog,text=f"{count['errors']} workbook errors | {len(affected)} affect selected students/history | "
                f"{count['warnings']} warnings | {sum(i.level=='INFO' for i in dataset.issues)} information notices",
                wraplength=1080).pack(padx=15,pady=10)
    ctk.CTkLabel(dialog,text='Select a finding to see its original cells. Correct the indicated row in your source workbook and reimport. '
                'Information notices do not block exports.',wraplength=1080).pack(padx=15,pady=5)
    from calculation_review import calculation_review_issues
    math_issues=calculation_review_issues(dataset,rows)
    affected.update(math_issues)
    selected_mode=['Calculation issues' if math_issues and not relevant_errors(dataset,rows) else 'Errors']
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
        text=issue.description+'\n\n'+ '\n'.join(f'{key}: {value}' for key,value in (source or {}).get('cells',{}).items())
        details.delete('1.0','end');details.insert('1.0',text)
    def export():
        path=filedialog.asksaveasfilename(parent=dialog,defaultextension='.xlsx',initialfile='LGU_Data_Review.xlsx')
        if not path:return
        from course_reports import export_report
        try:export_report('validation','excel',dataset,rows,path)
        except Exception as exc:messagebox.showerror('Review export',str(exc),parent=dialog)
    tree.bind('<<TreeviewSelect>>',show)
    actions=ctk.CTkFrame(dialog);actions.pack(fill='x',padx=15,pady=8)
    ctk.CTkOptionMenu(actions,values=['Errors','Selection errors','Calculation issues','All findings'],command=refresh).pack(side='left',padx=5)
    ctk.CTkButton(actions,text='Export full review log',command=export).pack(side='left',padx=5)
    refresh()
    return dialog
