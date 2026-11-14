"""Roll-number-wise semester summaries built from complete source history."""
import re
from calculation_policy import AcademicIndex,display_ratio,CalculationReviewRequired
from calculation_policy import calculation_attempts
from academic_identity import program_key,registration_key

def roll_order(value):
    return tuple((1,int(part)) if part.isdigit() else (0,part.casefold()) for part in re.split(r'(\d+)',value))

def semester_summary_table(dataset,rows):
    from gazette_layout import term_order
    index=AcademicIndex(calculation_attempts(dataset))
    selected={(registration_key(c.registration),program_key(c.degree)) for c in rows}
    rows=[c for c in dataset.courses if (registration_key(c.registration),program_key(c.degree)) in selected]
    groups={}
    for course in rows:groups.setdefault((registration_key(course.registration),program_key(course.degree),course.term),[]).append(course)
    output=[]
    for (registration,degree,term),attempts in sorted(groups.items(),key=lambda item:(roll_order(item[0][0]),term_order(item[0][2]),item[0][1])):
        full=[c for c in index.students[(registration,degree)] if c.term==term]
        semester_numbers={c.semester_number for c in full if c.semester_number}
        subjects={c.course_code:c.course_name for c in full}
        values=['','','','',''];reviews=[]
        try:
            result=index.summary(registration,degree,term)
            values[:3]=[str(result.total_credit_hours),str(result.total_grade_points),display_ratio(result.exact_ratio)]
        except CalculationReviewRequired as exc:reviews.append('GPA: '+str(exc))
        try:
            cumulative=index.summary(registration,degree,term,True)
            values[3:]=[str(cumulative.total_credit_hours),display_ratio(cumulative.exact_ratio)]
        except CalculationReviewRequired as exc:reviews.append('CGPA: '+str(exc))
        output.append([attempts[0].registration,attempts[0].name,attempts[0].degree,', '.join(sorted(semester_numbers)),term,*values,len(subjects),'; '.join(f'{code}: {name}' for code,name in sorted(subjects.items())),'; '.join(reviews)])
    return ['Registration / roll','Student name','Degree','Semester number','Exam term','Semester CH','Semester points','GPA','Cumulative CH','CGPA','Total subjects','Subject names','Review'],output

def write_gazette_cgpa_summary(pdf,rows,index,body_pages):
    """Companion totals where a reference layout has no CGPA column."""
    from gazette_layout import _cell,_text,term_order
    keys=sorted({(registration_key(c.registration),program_key(c.degree),c.term) for c in rows},key=lambda key:(roll_order(key[0]),term_order(key[2])))
    names={registration_key(c.registration):c.name for c in rows}
    labels={}
    for course in rows:labels.setdefault(registration_key(course.registration),course.registration)
    edges=(26,174,406,470,560,664,766);page_count=0;y=600
    for registration,degree,term in keys:
        if y+25>575:
            pdf.add_page(orientation='L',format=(612,792));page_count+=1
            _text(pdf,'GPA / CGPA SUMMARY',26,40,740,35,size=15,style='B');y=87
            for i,label in enumerate(('REGISTRATION','STUDENT NAME','TERM','GPA','CUMULATIVE CH','CGPA')):
                _cell(pdf,label,edges[i],y,edges[i+1]-edges[i],28,size=8,style='B')
            y+=28
            _text(pdf,str(body_pages+page_count),756,6,30,15,size=10)
        gpa=credits=cgpa='REVIEW'
        try:gpa=display_ratio(index.summary(registration,degree,term).exact_ratio)
        except CalculationReviewRequired:pass
        try:
            cumulative=index.summary(registration,degree,term,True)
            credits=str(cumulative.total_credit_hours);cgpa=display_ratio(cumulative.exact_ratio)
        except CalculationReviewRequired:pass
        for i,value in enumerate((labels[registration],names[registration],term,gpa,credits,cgpa)):
            _cell(pdf,value,edges[i],y,edges[i+1]-edges[i],25,size=8)
        y+=25
    return page_count
