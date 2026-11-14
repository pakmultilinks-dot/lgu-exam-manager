"""Explain calculation failures using the exact imported attempts, without changing policy."""
from calculation_policy import AcademicIndex,calculation_attempts,CalculationReviewRequired,_values,is_status_only_non_gpa
from academic_identity import registration_key,program_key
from gazette_layout import term_order
from academic_engine import DataIssue


def calculation_review_issues(dataset,rows):
    if dataset is None:return []
    key=(id(dataset.courses),len(dataset.courses),tuple(sorted(dataset.grade_scale.items())))
    if getattr(dataset,'_math_review_key',None)!=key:
        dataset._math_review_index=AcademicIndex(calculation_attempts(dataset));dataset._math_review_key=key
    index=dataset._math_review_index;selected={}
    for c in rows:selected.setdefault((registration_key(c.registration),program_key(c.degree),c.term),c)
    issues=[]
    for (reg,degree,term),current in selected.items():
        for cumulative in (False,True):
            try:index.summary(reg,degree,term,cumulative)
            except CalculationReviewRequired as exc:
                scope=[c for c in index.students.get((reg,degree),[]) if term_order(c.term)<=term_order(term)] if cumulative else [c for c in index.students.get((reg,degree),[]) if c.term==term]
                problem=None;grades=set()
                for c in scope:
                    if is_status_only_non_gpa(c):continue
                    try:_values(c)
                    except CalculationReviewRequired:
                        grades.add(c.grade or '(blank)')
                        if problem is None:problem=c
                problem=problem or current
                label='CGPA' if cumulative else 'GPA'
                details=f'{current.registration} | {current.degree} | through {term}: {label}: {exc}'
                if grades:details+=' Source grades needing a rule or value: '+', '.join(sorted(grades))+'.'
                issues.append(DataIssue('REVIEW',problem.source_sheet,problem.source_excel_row,label+'_CALCULATION_REVIEW',details,problem.source_file))
    return issues
