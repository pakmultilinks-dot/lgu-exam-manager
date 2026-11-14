"""Non-executable session archive; atomic saves and bounded reads."""
from dataclasses import fields, is_dataclass
import datetime
from decimal import Decimal
import json
import os
from pathlib import Path
import shutil
import tempfile
import zipfile
import pandas as pd
from academic_engine import load_course_workbook

MAX_DATASET_BYTES = 512 * 1024 * 1024

def save(path, mode, source, raw, mapping, settings, dataset=None):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        suffix=Path(source).suffix.lower() if mode=='course' and source and Path(source).is_file() else '.xlsx'
        data=Path(tmp)/('data'+suffix)
        if mode=='course':
            if source and Path(source).is_file():shutil.copyfile(source,data)
            elif path.is_file():
                with zipfile.ZipFile(path) as z:
                    same_dataset=True
                    if dataset is not None and 'dataset.json' in z.namelist():
                        previous=json.loads(z.read('dataset.json'))
                        same_dataset=previous.get('source_filename')==dataset.source_filename
                    member=next((name for name in ('data.xlsx','data.xlsm','data.xls','data.xlsb') if name in z.namelist()),None)
                    if member and same_dataset:
                        if z.getinfo(member).file_size>50*1024*1024:raise ValueError('Saved workbook exceeds limits')
                        data=Path(tmp)/member;data.write_bytes(z.read(member))
                    elif dataset is None:raise ValueError('Original workbook and course snapshot are unavailable.')
            elif dataset is None:raise ValueError('Original workbook unavailable; cannot save course session.')
        else:raw.to_excel(data,index=False)
        meta={'version':3,'mode':mode,'mapping':mapping,'settings':settings,'source_filename':Path(source).name if source else None}
        snapshot=Path(tmp)/'dataset.json'
        if dataset is not None:
            # Serialize records directly: asdict duplicates the entire source audit in memory.
            with snapshot.open('w',encoding='utf-8') as handle:
                json.dump(dataset,handle,default=_encode_cell,separators=(',',':'))
            if snapshot.stat().st_size>MAX_DATASET_BYTES:
                raise ValueError('Session dataset exceeds limits')
        from export_io import atomic_output
        with atomic_output(path) as staging:
            with zipfile.ZipFile(staging,'w',zipfile.ZIP_DEFLATED) as z:
                if data.exists():z.write(data,data.name)
                if dataset is not None:z.write(snapshot,'dataset.json')
                z.writestr('mapping.json',json.dumps(meta))

def restore(path):
    try:return _restore(path)
    except (zipfile.BadZipFile,KeyError,TypeError,json.JSONDecodeError,UnicodeDecodeError) as exc:
        raise ValueError("Saved session is incomplete or damaged. Import the original workbook or choose another saved session. The saved file was not changed.") from exc


def _restore(path):
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(path) as z:
            if 'dataset.json' in z.namelist():
                if z.getinfo('dataset.json').file_size>MAX_DATASET_BYTES:raise ValueError('Session dataset exceeds limits')
                if z.getinfo('mapping.json').file_size>1000000:raise ValueError('Session metadata exceeds limits')
                meta=json.loads(z.read('mapping.json'))
                snapshot=json.loads(z.read('dataset.json'),object_hook=_decode_cell)
                from academic_engine import ResultDataset,CourseAttempt,SheetLayout,DataIssue,FeeRecord
                snapshot['courses']=[CourseAttempt(**v) for v in snapshot['courses']]
                snapshot['layouts']=[SheetLayout(**v) for v in snapshot['layouts']]
                snapshot['issues']=[DataIssue(**v) for v in snapshot['issues']]
                snapshot['fee_records']=[FeeRecord(**v) for v in snapshot.get('fee_records',[])]
                from grading_profiles import restore_scale_profile
                restore_scale_profile(snapshot)
                return meta,ResultDataset(**snapshot)
            member=next((name for name in ('data.xlsx','data.xlsm','data.xls','data.xlsb') if name in z.namelist()),'data.xlsx')
            for name,limit in [(member,50*1024*1024),('mapping.json',1000000)]:
                if z.getinfo(name).file_size>limit:raise ValueError('Session member exceeds size limit')
            meta=json.loads(z.read('mapping.json'));data=Path(tmp)/member;data.write_bytes(z.read(member))
        if meta.get('mode')=='course':
            ds=load_course_workbook(data)
            if meta.get('source_filename'):ds.source_filename=Path(meta['source_filename']).name
            return meta,ds
        return meta,pd.read_excel(data,engine='openpyxl')


def _encode_cell(value):
    if is_dataclass(value) and not isinstance(value,type):
        return {field.name:getattr(value,field.name) for field in fields(value)}
    if isinstance(value,(datetime.datetime,datetime.date,datetime.time,Decimal)):
        return {'__lgu_cell_type__':type(value).__name__,'value':str(value)}
    if hasattr(value,'item'):return value.item()
    raise TypeError('Unsupported source cell type: '+type(value).__name__)

def _decode_cell(value):
    kind=value.get('__lgu_cell_type__')
    converters={'datetime':datetime.datetime.fromisoformat,'date':datetime.date.fromisoformat,'time':datetime.time.fromisoformat,'Decimal':Decimal}
    if kind in converters and set(value)=={'__lgu_cell_type__','value'}:return converters[kind](value['value'])
    return value
