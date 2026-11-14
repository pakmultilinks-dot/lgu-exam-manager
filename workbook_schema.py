from excel_io import excel_engine
"""Structural workbook reading and explicit long/wide course mappings."""
from collections import Counter
from pathlib import Path
import hashlib,json
import pandas as pd
from openpyxl.utils import get_column_letter

ROLE_LABELS={'registration':'Registration / roll number','name':'Student name','degree':'Program / degree','department':'Department','admission_session':'Admission intake','semester':'Semester description','term':'Exam term','semester_number':'Academic semester number','course_code':'Course code','course_name':'Course title','credits':'Credit hours','grade':'Letter grade / status','marks':'Marks','quality_points':'Total weighted grade points','grade_point_rate':'Grade points per credit','award_id':'Award identifier','fee_status':'Fee status','fee_comment':'Fee comment'}


def header_labels(rows):
    values=[];parent=''
    for column in range(max(map(len,rows),default=0)):
        parts=[]
        for depth,row in enumerate(rows):
            value=row[column] if column<len(row) else None
            text='' if pd.isna(value) else str(value).strip()
            if len(rows)>1 and depth==0:
                parent=text or parent;text=parent
            if text and text not in parts:parts.append(text)
        values.append(' '.join(parts) or f'Column {get_column_letter(column+1)}')
    counts=Counter(values)
    return [value if counts[value]==1 else f'{value} [column {get_column_letter(i+1)}]' for i,value in enumerate(values)]


def read_raw_table(excel, sheet, header_row, header_rows=1, preview=None):
    if header_row<0 or header_rows not in (1,2):raise ValueError('Choose a valid heading row and one or two heading rows.')
    frame=pd.read_excel(excel,sheet_name=sheet,header=None,skiprows=header_row,nrows=None if preview is None else preview+header_rows,dtype=object,keep_default_na=False)
    if len(frame)<header_rows:raise ValueError('The selected heading rows are outside this worksheet.')
    labels=header_labels(frame.iloc[:header_rows].values.tolist())
    body=frame.iloc[header_rows:].copy().reset_index(drop=True);body.columns=labels
    return body


def read_layout_table(excel,layout):
    from academic_engine import NEEDED,PLACEHOLDERS
    frame=read_raw_table(excel,layout.sheet,layout.header_row,layout.header_rows)
    raw=[{'file':Path(layout.source_file).name,'sheet':layout.sheet,'excel_row':layout.header_row+layout.header_rows+i+1,'cells':{str(k):(None if pd.isna(v) else v) for k,v in zip(frame.columns,row)}} for i,row in enumerate(frame.itertuples(index=False,name=None))]
    columns=dict(layout.columns)
    for role,column in columns.items():
        if column not in frame.columns:raise ValueError(f'Choose an existing column for {ROLE_LABELS.get(role,role)} on {layout.sheet}.')
    for role,value in layout.defaults.items():
        if role not in columns:
            column='@Entered '+role;frame[column]=value;columns[role]=column
    if not layout.course_blocks:
        if not NEEDED.issubset(columns):raise ValueError('Map all required course fields before importing this worksheet.')
        return frame,columns,raw,[item['excel_row'] for item in raw]
    if not {'registration','name'}.issubset(columns):raise ValueError('Map registration and student name for the course-column layout.')
    output=[];row_numbers=[]
    block_roles=('course_code','course_name','credits','grade','marks','quality_points','grade_point_rate')
    for i,record in enumerate(frame.to_dict('records')):
        for block in layout.course_blocks:
            mapped=block.get('columns',{});fixed=block.get('defaults',{})
            if not {'course_code','course_name','credits','grade'}.issubset(set(mapped)|set(fixed)):
                raise ValueError('Each course group needs a code, title, credit hours and grade column.')
            if any(column not in frame.columns for column in mapped.values()):raise ValueError('A course group refers to a column that is not in this worksheet.')
            evidence=[record[column] for role,column in mapped.items() if role in ('grade','marks','quality_points','grade_point_rate')]
            if evidence and all(pd.isna(value) or str(value).strip().upper() in PLACEHOLDERS for value in evidence):continue
            row=dict(record)
            for role in block_roles:row['@Course '+role]=record[mapped[role]] if role in mapped else fixed.get(role,'')
            output.append(row);row_numbers.append(raw[i]['excel_row'])
            if len(output)>200000:raise ValueError('The expanded course table exceeds 200,000 attempts. Import smaller selections.')
    for role in block_roles:columns[role]='@Course '+role
    if not output:raise ValueError('The mapped course groups contain no results.')
    return pd.DataFrame(output),columns,raw,row_numbers


def inspect_workbook(path,known=()):
    from academic_engine import identify_columns
    known={layout.sheet:layout for layout in known};result=[]
    with pd.ExcelFile(path,engine=excel_engine(path)) as excel:
        for sheet in excel.sheet_names:
            preview=pd.read_excel(excel,sheet_name=sheet,header=None,nrows=200,dtype=object)
            nonempty=preview.dropna(how='all').dropna(axis=1,how='all')
            if nonempty.empty:continue
            candidates=[]
            for row in range(len(preview)):
                for depth in (1,2):
                    if row+depth>len(preview):continue
                    labels=header_labels(preview.iloc[row:row+depth].values.tolist());roles=identify_columns(labels)
                    candidates.append((len(roles),-depth,-row,row,depth,labels,roles))
            best=max(candidates) if candidates else (0,-1,0,0,1,[],{})
            layout=known.get(sheet)
            row=layout.header_row if layout else best[3];depth=layout.header_rows if layout else best[4]
            body=read_raw_table(excel,sheet,row,depth,preview=8)
            remembered=load_profiles(default_profile_path()).get(profile_key(body.columns,depth),{})
            result.append({'sheet':sheet,'header_row':row,'header_rows':depth,'columns':list(body.columns),'mapping':dict(layout.columns) if layout else remembered.get('columns',identify_columns(body.columns)),'defaults':dict(layout.defaults) if layout else {},'course_blocks':list(layout.course_blocks) if layout else remembered.get('course_blocks',[]),'preview':body.astype(str).values.tolist(),'recognized':layout is not None,'table_like':nonempty.shape[0]>=3 and nonempty.shape[1]>=3})
    return result


def profile_key(labels,depth):
    return hashlib.sha256(json.dumps([list(labels),depth],ensure_ascii=False).encode('utf-8')).hexdigest()


def load_profiles(path):
    path=Path(path)
    if not path.is_file():return {}
    try:
        data=json.loads(path.read_text(encoding='utf-8'))
        return data if isinstance(data,dict) else {}
    except (ValueError,OSError):return {}


def save_profile(path,labels,layout):
    from dataclasses import asdict
    from export_io import atomic_output
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True);data=load_profiles(path)
    record=asdict(layout)
    for field in ('sheet','source_file','row_count_estimate','header_row'):record.pop(field,None)
    record['requires_review']=bool(record.get('defaults') or any(block.get('defaults') for block in record.get('course_blocks',[])))
    record['defaults']={}
    for block in record.get('course_blocks',[]):block['defaults']={}
    data[profile_key(labels,layout.header_rows)]=record
    with atomic_output(path) as staging:Path(staging).write_text(json.dumps(data,indent=2),encoding='utf-8')


def default_profile_path():
    return Path.home()/".lgu-exam-manager/import_profiles.json"


def needs_review(inspection,layouts):
    from academic_engine import _key
    mapped={layout.sheet:layout for layout in layouts}
    for item in inspection:
        layout=mapped.get(item['sheet'])
        if item['table_like'] and layout is None:return True
        if layout is not None and not layout.confirmed:
            supplied={layout.columns.get('quality_points'),layout.columns.get('grade_point_rate')}
            if any(_key(column) in {'gp','grade point','grade points','points','point value'} and column not in supplied for column in item['columns']):return True
    return not layouts
