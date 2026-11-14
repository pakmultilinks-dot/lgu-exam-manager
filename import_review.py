from excel_io import excel_engine
"""Workbook-first import review, with no filename or cohort assumptions."""
from pathlib import Path
from tkinter import ttk,messagebox
import customtkinter as ctk
import pandas as pd
from workbook_schema import ROLE_LABELS,read_raw_table,save_profile,default_profile_path
from academic_engine import SheetLayout,NEEDED,identify_columns

EMPTY='Not supplied'
FIXED={'degree','department','admission_session','term','semester_number','course_code','course_name','credits'}
BLOCK_ROLES=('credits','grade','marks','quality_points','grade_point_rate')

class WorkbookImportReview(ctk.CTkToplevel):
    def __init__(self,parent,path,inspection,apply):
        super().__init__(parent);self.title('Review workbook import');self.geometry('1080x820');self.parent=parent;self.path=str(path);self.apply_callback=apply
        self.states={item['sheet']:dict(item,include=item['recognized'] or item['table_like'],mode='Courses across columns' if item.get('course_blocks') else 'One course per row') for item in inspection}
        self.current=None;self.controls={};self.fixed={};self.blocks=[]
        ctk.CTkLabel(self,text=Path(path).name,font=('Arial',20,'bold')).pack(anchor='w',padx=18,pady=10)
        ctk.CTkLabel(self,text='Choose the result sheets and map their actual columns. Required fields are marked *. Total weighted points and points per credit are different fields.',wraplength=1000,justify='left').pack(anchor='w',padx=18)
        bar=ctk.CTkFrame(self);bar.pack(fill='x',padx=18,pady=8)
        self.sheet=ctk.CTkOptionMenu(bar,values=list(self.states),command=self.select_sheet,width=260);self.sheet.pack(side='left',padx=8)
        self.include=ctk.BooleanVar(value=True);ctk.CTkCheckBox(bar,text='Include this sheet',variable=self.include).pack(side='left',padx=8)
        ctk.CTkLabel(bar,text='Heading row').pack(side='left',padx=5)
        self.header=ctk.CTkEntry(bar,width=65);self.header.pack(side='left')
        self.depth=ctk.CTkOptionMenu(bar,values=['1','2'],width=65);self.depth.pack(side='left',padx=6)
        ctk.CTkButton(bar,text='Read headings',command=self.read_headings,width=120).pack(side='left',padx=8)
        preview=ttk.Frame(self);preview.pack(fill='x',padx=18,pady=6)
        self.tree=ttk.Treeview(preview,show='headings',height=5);self.tree.pack(fill='x')
        scroll=ttk.Scrollbar(preview,orient='horizontal',command=self.tree.xview);scroll.pack(fill='x');self.tree.configure(xscrollcommand=scroll.set)
        self.mode=ctk.CTkOptionMenu(self,values=['One course per row','Courses across columns'],command=self.change_mode,width=280);self.mode.pack(anchor='w',padx=18,pady=6)
        self.editor=ctk.CTkScrollableFrame(self);self.editor.pack(fill='both',expand=True,padx=18,pady=5)
        bottom=ctk.CTkFrame(self);bottom.pack(fill='x',padx=18,pady=10)
        self.remember=ctk.BooleanVar(value=True);ctk.CTkCheckBox(bottom,text='Remember this column structure for future workbooks',variable=self.remember).pack(side='left',padx=6)
        ctk.CTkButton(bottom,text='Import selected sheets',command=self.apply,width=180).pack(side='right',padx=6)
        ctk.CTkButton(bottom,text='Cancel',command=self.destroy,width=90).pack(side='right',padx=6)
        self.select_sheet(next(iter(self.states)))

    def store(self):
        if self.current is None:return
        state=self.states[self.current];state['include']=self.include.get();state['mode']=self.mode.get()
        state['mapping']={role:menu.get() for role,menu in self.controls.items() if menu.get()!=EMPTY}
        state['defaults']={role:entry.get().strip() for role,entry in self.fixed.items() if entry.get().strip() and role not in state['mapping']}
        state['course_blocks']=[{'columns':{role:menu.get() for role,menu in menus.items() if menu.get()!=EMPTY},'defaults':{'course_code':code.get().strip(),'course_name':title.get().strip(),**({'credits':credit.get().strip()} if credit.get().strip() and menus['credits'].get()==EMPTY else {})}} for code,title,menus,credit in self.blocks]

    def select_sheet(self,name):
        self.store();self.current=name;state=self.states[name];self.sheet.set(name);self.include.set(state['include']);self.mode.set(state['mode'])
        self.header.delete(0,'end');self.header.insert(0,str(state['header_row']+1));self.depth.set(str(state['header_rows']))
        self.tree.delete(*self.tree.get_children());self.tree.configure(columns=[f'c{i}' for i in range(len(state['columns']))])
        for i,column in enumerate(state['columns']):self.tree.heading(f'c{i}',text=column);self.tree.column(f'c{i}',width=150)
        for row in state['preview']:self.tree.insert('','end',values=row)
        self.render_editor()

    def render_editor(self):
        for child in self.editor.winfo_children():child.destroy()
        self.controls={};self.fixed={};self.blocks=[];state=self.states[self.current];wide=state['mode']=='Courses across columns'
        values=[EMPTY]+state['columns']
        for role,label in ROLE_LABELS.items():
            if wide and role in ('course_code','course_name',*BLOCK_ROLES):continue
            row=ctk.CTkFrame(self.editor,fg_color='transparent');row.pack(fill='x',pady=3)
            required=role in ({'registration','name'} if wide else NEEDED)
            ctk.CTkLabel(row,text=label+(' *' if required else ''),width=230,anchor='w').pack(side='left')
            menu=ctk.CTkOptionMenu(row,values=values,width=350);menu.set(state['mapping'].get(role,EMPTY));menu.pack(side='left',padx=5);self.controls[role]=menu
            if role in FIXED:
                entry=ctk.CTkEntry(row,width=270,placeholder_text='Or enter one value for this sheet');entry.insert(0,state.get('defaults',{}).get(role,''));entry.pack(side='left',padx=5);self.fixed[role]=entry
        if wide:
            ctk.CTkLabel(self.editor,text='Add each course group. Enter its code and title, then select the matching grade and credit columns.',wraplength=900).pack(anchor='w',pady=10)
            self.block_holder=ctk.CTkFrame(self.editor);self.block_holder.pack(fill='x')
            for block in state.get('course_blocks',[]):self.add_block(block)
            ctk.CTkButton(self.editor,text='Add course group',command=self.add_block).pack(anchor='w',pady=8)

    def add_block(self,block=None):
        block=block or {};frame=ctk.CTkFrame(self.block_holder);frame.pack(fill='x',pady=8)
        top=ctk.CTkFrame(frame,fg_color='transparent');top.pack(fill='x')
        code=ctk.CTkEntry(top,width=200,placeholder_text='Course code');code.insert(0,block.get('defaults',{}).get('course_code',''));code.pack(side='left',padx=5,pady=5)
        title=ctk.CTkEntry(top,width=500,placeholder_text='Course title');title.insert(0,block.get('defaults',{}).get('course_name',''));title.pack(side='left',padx=5)
        menus={}
        for role in BLOCK_ROLES:
            row=ctk.CTkFrame(frame,fg_color='transparent');row.pack(fill='x',pady=2)
            ctk.CTkLabel(row,text=ROLE_LABELS[role]+(' *' if role in ('grade','credits') else ''),width=230,anchor='w').pack(side='left',padx=5)
            menu=ctk.CTkOptionMenu(row,values=[EMPTY]+self.states[self.current]['columns'],width=430);menu.set(block.get('columns',{}).get(role,EMPTY));menu.pack(side='left',padx=5);menus[role]=menu
        credit=ctk.CTkEntry(frame,width=300,placeholder_text='Or enter the credit hours for this course');credit.insert(0,block.get('defaults',{}).get('credits',''));credit.pack(anchor='w',padx=240,pady=4)
        record=(code,title,menus,credit);self.blocks.append(record)
        ctk.CTkButton(top,text='Remove',width=80,command=lambda:(self.blocks.remove(record),frame.destroy())).pack(side='right',padx=5)

    def change_mode(self,value):
        self.store();self.states[self.current]['mode']=value;self.render_editor()

    def read_headings(self):
        self.store()
        try:row=int(self.header.get())-1;depth=int(self.depth.get());assert row>=0
        except (ValueError,AssertionError):messagebox.showwarning('Heading row','Enter the Excel row number containing the headings.',parent=self);return
        sheet=self.current
        def work():
            with pd.ExcelFile(self.path,engine=excel_engine(self.path)) as excel:return read_raw_table(excel,sheet,row,depth,preview=8)
        def done(frame):
            state=self.states[sheet];state.update(header_row=row,header_rows=depth,columns=list(frame.columns),preview=frame.astype(str).values.tolist(),mapping=identify_columns(frame.columns),defaults={},course_blocks=[])
            if self.current==sheet:
                self.current=None;self.select_sheet(sheet)
        self.parent.run_job('Reading worksheet preview',work,done)

    def apply(self):
        if self.parent.job_busy:return
        self.store();layouts=[]
        try:
            for name,state in self.states.items():
                if not state['include']:continue
                columns=state['mapping'];defaults=state['defaults'];blocks=state['course_blocks'] if state['mode']=='Courses across columns' else []
                required={'registration','name'} if state['mode']=='Courses across columns' else NEEDED
                if not required.issubset(set(columns)|set(defaults)):raise ValueError(f'{name}: map every required field or supply the sheet-level course details.')
                if len(set(columns.values()))!=len(columns):raise ValueError(f'{name}: each shared field must use a distinct column.')
                if state['mode']=='Courses across columns' and not blocks:raise ValueError(f'{name}: add at least one course group.')
                for block in blocks:
                    if not block['defaults']['course_code'] or not block['defaults']['course_name'] or 'grade' not in block['columns'] or ('credits' not in block['columns'] and not block['defaults'].get('credits')):raise ValueError(f'{name}: each course group needs a code, title, grade column and credit column.')
                    if len(set(block['columns'].values()))!=len(block['columns']):raise ValueError(f'{name}: the course fields must use distinct columns.')
                layouts.append(SheetLayout(name,state['header_row'],columns,0,Path(self.path).name,state['header_rows'],defaults,blocks,True))
            if not layouts:raise ValueError('Include at least one result sheet.')
        except ValueError as exc:messagebox.showwarning('Complete import mapping',str(exc),parent=self);return
        if self.remember.get():
            try:
                for layout in layouts:save_profile(default_profile_path(),self.states[layout.sheet]['columns'],layout)
            except OSError:messagebox.showwarning('Mapping not remembered','This import can continue, but its column mapping could not be saved.',parent=self)
        self.apply_callback(layouts);self.destroy()
