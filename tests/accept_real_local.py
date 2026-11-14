"""Authorized real-file acceptance: summaries contain no student identities."""
from pathlib import Path
import sys,time,json
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from academic_engine import load_course_workbook
from course_reports import ROUTES,capability,export_report
from openpyxl import load_workbook
from docx import Document
import pymupdf as fitz
out=Path('.private/acceptance');out.mkdir(parents=True,exist_ok=True)
import argparse
parser=argparse.ArgumentParser();parser.add_argument('--workbook',required=True);args=parser.parse_args()
t=time.perf_counter();ds=load_course_workbook(args.workbook);load_seconds=time.perf_counter()-t
results=[]
for key in ROUTES:
    ok,reason=capability(key,ds,ds.courses)
    if not ok:results.append({'route':key,'blocked':reason});continue
    for fmt,ext in [('pdf','.pdf'),('word','.docx'),('excel','.xlsx')]:
        p=out/(key+ext);t=time.perf_counter()
        try:
            export_report(key,fmt,ds,ds.courses,str(p))
            if fmt=='pdf':
                d=fitz.open(p);assert len(d)>0;assert 'DRAFT' in d[0].get_text();pages=len(d)
                if key=='register':assert ds.courses[0].registration in ''.join(pg.get_text() for pg in d)
            elif fmt=='word':
                d=Document(p);assert len(d.paragraphs)>0;pages=None
            else:
                w=load_workbook(p,read_only=True);assert w.active.max_row>1
                if key=='register':assert w['Course Results'].max_row==len(ds.courses)+1
                w.close();pages=None
            results.append({'route':key,'format':fmt,'status':'PASS','bytes':p.stat().st_size,'seconds':round(time.perf_counter()-t,3),'pages':pages})
        except Exception as e:results.append({'route':key,'format':fmt,'status':'FAIL','error':type(e).__name__+': '+str(e)})
summary={'counts':ds.counts(),'degrees':ds.degrees,'cohorts':ds.cohorts,'terms':ds.terms,'load_seconds':load_seconds,'results':results}
(out/'results.json').write_text(json.dumps(summary,indent=2));print(json.dumps(summary,indent=2))
