"""Editable grading profiles read from bundled data, never inferred from marks."""
from pathlib import Path
from functools import lru_cache
import json

@lru_cache(maxsize=1)
def reference_profile():
    profile=json.loads((Path(__file__).with_name('gazette_assets')/'grade_scale.json').read_text(encoding='utf-8'))
    from calculation_policy import parse_grade_scale
    rates=parse_grade_scale('\n'.join(f'{grade}={rate}' for grade,rate in profile['rates'].items()),profile['reference'])
    return {'id':profile['id'],'reference':profile['reference'],'rates':rates,'previous_references':tuple(profile.get('previous_references',[]))}

def default_grade_scale():return dict(reference_profile()['rates'])
def default_grade_reference():return reference_profile()['reference']

def scale_coverage(dataset,rows=None):
    from decimal import Decimal,InvalidOperation
    from collections import Counter
    rows=dataset.courses if rows is None else rows
    source=derived=source_rate=0;unmapped=Counter();nonfinal=Counter();invalid=0
    for course in rows:
        grade=course.grade.strip().upper()
        if course.quality_points.strip():source+=1;continue
        if grade in ('W','I','IP','P','PASS','R',''):
            nonfinal[grade or '(blank)']+=1;continue
        supplied=getattr(course,'grade_point_rate','').strip()
        rate=supplied or dataset.grade_scale.get(grade)
        if rate is None:unmapped[grade]+=1;continue
        try:
            credits=Decimal(course.credits)
            point_rate=Decimal(rate)
            if not credits.is_finite() or credits<=0 or not point_rate.is_finite() or point_rate<0:invalid+=1;continue
        except InvalidOperation:invalid+=1;continue
        derived+=1
        if supplied:source_rate+=1
    return {'source':source,'source_rate':source_rate,'derived':derived,'unmapped':dict(unmapped),'nonfinal':dict(nonfinal),'invalid_credits':invalid}


def restore_scale_profile(snapshot):
    """Refresh managed defaults; preserve both UI and legacy/programmatic custom scales."""
    from decimal import Decimal
    if 'grade_scale_user_set' not in snapshot:
        snapshot['grade_scale_user_set']=bool(snapshot.get('grade_scale'))
    if snapshot['grade_scale_user_set']:return
    profile=reference_profile();saved=snapshot.get('grade_scale') or {}
    reference=snapshot.get('grade_scale_reference','')
    known_references={profile['reference'],*profile['previous_references']}
    if saved and (reference not in known_references or any(grade not in profile['rates'] or Decimal(str(rate))!=Decimal(profile['rates'][grade]) for grade,rate in saved.items())):
        snapshot['grade_scale_user_set']=True
        return
    snapshot['grade_scale']=default_grade_scale()
    snapshot['grade_scale_reference']=default_grade_reference()
