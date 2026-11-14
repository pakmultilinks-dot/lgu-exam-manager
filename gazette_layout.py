"""LGU reference-format gazette renderer. Source attempts only; no invented totals."""
from pathlib import Path
from collections import defaultdict
from functools import lru_cache
import re
from fpdf import FPDF
from calculation_policy import calculation_attempts
from academic_identity import program_key,registration_key

X=(50,68,125,188,244,266,305,526,548,577,601,630,648,663,685,742)
TOP=155
BOTTOM=565
ROW=20


def term_order(value):
    m=re.fullmatch(r'(SU|S|F)(\d{2,4})',value.upper())
    if not m:return (9999,9,value)
    year=int(m[2]);year=year+2000 if year<100 else year
    return (year,{'S':0,'SU':1,'F':2}[m[1]],value)


def _fonts(pdf):
    from pdf_fonts import add_fonts, add_font_family
    add_fonts(pdf)
    root=Path('C:/Windows/Fonts')
    if (root/'calibri.ttf').is_file():
        for style,file in [('', 'calibri.ttf'),('B','calibrib.ttf'),('I','calibrii.ttf')]:pdf.add_font('Body',style,str(root/file))
    else:
        # Bundled DejaVu fonts (gazette_assets/fonts); matplotlib is only a fallback.
        add_font_family(pdf, 'Body', {'': 'DejaVuSans.ttf', 'B': 'DejaVuSans-Bold.ttf', 'I': 'DejaVuSans-Oblique.ttf'})
    root=Path('C:/Windows/Fonts')
    if (root/'cambria.ttc').is_file():
        pdf.add_font('Title','',str(root/'cambria.ttc'))
        for style,file in [('B','cambriab.ttf'),('I','cambriai.ttf'),('BI','cambriaz.ttf')]:pdf.add_font('Title',style,str(root/file))
    else:
        add_font_family(pdf, 'Title', {'': 'DejaVuSerif.ttf', 'B': 'DejaVuSerif-Bold.ttf', 'I': 'DejaVuSerif-Italic.ttf', 'BI': 'DejaVuSerif-BoldItalic.ttf'})


def _wrap(pdf,text,width,size=7.3,family='Body',style=''):
    if '\n' in str(text):
        return [line for part in str(text).split('\n') for line in _wrap(pdf,part,width,size,family,style)]
    pdf.set_font(family,style,size)
    lines=[];line=''
    for word in str(text).split():
        if pdf.get_string_width(word)>width:
            pieces=[];piece=''
            for char in word:
                if piece and pdf.get_string_width(piece+char)>width:pieces.append(piece);piece=''
                piece+=char
            if piece:pieces.append(piece)
        else:pieces=[word]
        for piece in pieces:
            candidate=(line+' '+piece).strip()
            if line and pdf.get_string_width(candidate)>width:lines.append(line);line=piece
            else:line=candidate
    if line:lines.append(line)
    return lines or ['']


def _text(pdf,text,x,y,w,h,size=7.3,style='',family='Body',align='C'):
    lines=_wrap(pdf,text,w-4,size,family,style)
    while len(lines)*size*1.13>h-3 and size>5.5:
        size-=.3;lines=_wrap(pdf,text,w-4,size,family,style)
    pdf.set_font(family,style,size)
    step=size*1.13;start=y+(h-len(lines)*step)/2
    for i,line in enumerate(lines):
        pdf.set_xy(x+2,start+i*step)
        pdf.cell(w-4,step,line,align=align)


def _cell(pdf,text,x,y,w,h,**kw):
    pdf.rect(x,y,w,h)
    _text(pdf,text,x,y,w,h,**kw)


def _cohort_short(value):
    m=re.fullmatch(r'(FALL|SPRING|SUMMER)\s+(\d{4})',value)
    return {'FALL':'F','SPRING':'S','SUMMER':'SU'}[m[1]]+m[2][-2:] if m else value


@lru_cache(maxsize=1)
def _catalog():
    import json
    return json.loads((Path(__file__).with_name('gazette_assets')/'reference_layouts.json').read_text())


def _profile(key,cohort,degree,require_cgpa=False):
    data=_catalog();routes=data['routing'][key]
    options=routes.get(cohort)
    if not options:
        # A formatting fallback only; never an academic rule.
        options=next(iter(routes.values()))
    from gazette_options import canonical_program
    program=canonical_program(degree)
    ident=options.get(program,options['*'])
    profile=data['profiles'][ident]
    if require_cgpa and not profile['cgpa']:
        candidates=[]
        for session,programs in routes.items():
            candidate=data['profiles'][programs.get(program,programs['*'])]
            if candidate['cgpa']:candidates.append((term_order(_cohort_short(session)),candidate))
        if candidates:profile=max(candidates,key=lambda pair:pair[0])[1]
    from gazette_options import _routing_rules
    label=_routing_rules().get('department_headers',{}).get(key,{}).get(program)
    return _with_department_header(profile,label) if label else profile


def _with_department_header(profile,label):
    # Geometry may be shared; department identity must follow the actual program.
    return dict(profile,header=[dict(item,text=label,center=True) if 'DEPARTMENT OF' in item['text'] else item for item in profile['header']])


def _window(term,available):
    term=term.upper()
    if term.startswith('SU') and 'S'+term[2:] in available:return 'S'+term[2:]
    return term


def _full_term(term):
    m=re.fullmatch(r'(SU|S|F)(\d{2,4})',term)
    if not m:return term
    year=m[2] if len(m[2])==4 else '20'+m[2]
    return {'SU':'SUMMER','S':'SPRING','F':'FALL'}[m[1]]+' '+year


def _plan(pdf,rows,key,require_cgpa=False):
    available=defaultdict(set)
    labels={}
    for c in rows:
        identity=(c.admission_session,program_key(c.degree))
        available[identity].add(c.term.upper())
        labels.setdefault(identity,c.degree)
    students=defaultdict(list)
    registration_labels={}
    for c in rows:
        degree=program_key(c.degree)
        window=_window(c.term,available[(c.admission_session,degree)])
        student=registration_key(c.registration)
        registration_labels.setdefault((c.admission_session,degree,student),c.registration)
        students[(c.admission_session,degree,window,student)].append(c)
    pages=[];page=None;serial={}
    for identity,attempts in sorted(students.items(),key=lambda item:(term_order(_cohort_short(item[0][0])),term_order(item[0][2]),item[0][1],item[0][3])):
        cohort,degree_id,window,student=identity
        registration=registration_labels[(cohort,degree_id,student)]
        degree=labels[(cohort,degree_id)]
        name=attempts[0].name
        serial.setdefault(student,len(serial)+1)
        profile=_profile(key,cohort,degree,require_cgpa);x=profile['edges'];row_h=profile['row_height']
        regular=[c for c in attempts if not c.term.upper().startswith('SU')]
        summer=[c for c in attempts if c.term.upper().startswith('SU')]
        sections=[]
        for section,items in [('regular',regular),('summer',summer)]:
            if not items:continue
            values=[]
            for c in sorted(items,key=lambda c:c.source_excel_row):
                lines=len(_wrap(pdf,c.course_name,x[7]-x[6]-4))
                h=max(row_h,lines*8.3+3)
                values.append((c,h,section))
            minimum=6 if regular and summer and section=='regular' else 4 if regular and summer else 10
            while len(values)<minimum:values.append((None,row_h,section))
            sections.extend(values)
        chunks=[];chunk=[];height=0;capacity=20*row_h
        for entry in sections:
            if chunk and height+entry[1]>capacity+.01:chunks.append(chunk);chunk=[];height=0
            chunk.append(entry);height+=entry[1]
        if chunk:chunks.append(chunk)
        for part,chunk in enumerate(chunks):
            while sum(v[1] for v in chunk)<10*row_h-.01:chunk.append((None,row_h,chunk[-1][2]))
            height=sum(v[1] for v in chunk)
            group=(cohort,degree,window)
            if page is None or page['group']!=group or page['used']+height>capacity+.01:
                page={'group':group,'blocks':[],'used':0,'profile':profile};pages.append(page)
            page['blocks'].append((serial[student],(cohort,degree,registration,name),chunk,height,part));page['used']+=height
    return pages


def _cover(pdf,rows,key,settings=None):
    settings=settings or {}
    pdf.add_page(orientation='P',format=(612,792))
    design=_catalog()['cover_design']
    asset=Path(__file__).with_name('gazette_assets')/design['asset']
    if asset.is_file():pdf.image(str(asset),0,0,612,792)
    else:raise ValueError('Gazette cover artwork missing: '+key)
    # Replace the sample publication badge with source term coverage.
    bx,by,bw,bh=design['period_box']
    pdf.set_fill_color(194,183,122);pdf.rect(bx,by,bw,bh,'F')
    terms=sorted({c.term for c in rows},key=term_order)
    years=sorted({term_order(t)[0] for t in terms if term_order(t)[0]!=9999})
    year_label=str(years[0]) if len(years)==1 else f'{years[0]}-{years[-1]}' if years else ''
    has_summer=any(t.upper().startswith('SU') for t in terms)
    regular={t[0] for t in terms if not t.upper().startswith('SU')}
    if has_summer and regular:
        label=('SPRING &' if regular=={'S'} else 'FALL &' if regular=={'F'} else 'REGULAR &')+'\nSUMMER\n'+year_label
    else:label=('SUMMER' if has_summer else 'SPRING' if regular=={'S'} else 'FALL' if regular=={'F'} else 'REGULAR')+'\n'+year_label
    label=settings.get('cover_period','').strip() or label
    pdf.set_text_color(0,0,0)
    _text(pdf,label,bx+2,by+7,bw-4,bh-19,size=20,style='B',family='Title')
    headers={s['text'].strip() for cohort,degree in {(c.admission_session,c.degree) for c in rows} for s in _profile(key,cohort,degree)['header'] if 'DEPARTMENT OF' in s['text']}
    department=next(iter(headers)) if len(headers)==1 else key.replace('-',' ').upper()+' DEPARTMENT'
    department=settings.get('cover_department','').strip() or department
    dx,dy,dw,dh=design['department_box']
    pdf.set_fill_color(0,60,42);pdf.rect(dx,dy,dw,dh,'F');pdf.set_text_color(193,164,65)
    _text(pdf,department,dx+2,dy+1,dw-4,dh-2,size=20,style='B',family='Title')
    title=settings.get('cover_title','').strip()
    if title:
        tx,ty,tw,th=design['title_box']
        pdf.set_fill_color(250,251,249);pdf.rect(tx,ty,tw,th,'F');pdf.set_text_color(0,60,42)
        _text(pdf,title,tx+2,ty+1,tw-4,th-2,size=65,style='B',family='Title')
    prepared=settings.get('cover_prepared_by','').strip()
    if prepared:
        px,py,pw,ph=design['prepared_box']
        pdf.set_fill_color(250,251,249);pdf.rect(px,py,pw,ph,'F');pdf.set_text_color(0,60,42)
        _text(pdf,'Prepared by:\n'+prepared,px+2,py+1,pw-4,ph-2,size=19,style='B',family='Title')


def _contents(pdf,plan):
    entries={}
    for number,page in enumerate(plan,1):
        cohort,degree,window=page['group']
        entry=entries.setdefault(cohort,{'degrees':[],'page':number})
        if degree not in entry['degrees']:entry['degrees'].append(degree)
    entries=list(entries.items())
    for offset in range(0,len(entries),25):
        pdf.add_page(orientation='P',format=(612,792));pdf.set_text_color(0,0,0)
        _text(pdf,'TABLE OF CONTENTS',110,37,420,22,size=16,style='B')
        edges=(111,149,232,472,529);y=60
        for i,head in enumerate(('SR','SESSION','DEGREE','PAGE NUMBER')):_cell(pdf,head,edges[i],y,edges[i+1]-edges[i],29,size=8,style='B')
        y+=29
        for idx,(cohort,entry) in enumerate(entries[offset:offset+25],offset+1):
            vals=(idx,_cohort_short(cohort),', '.join(entry['degrees']),entry['page'])
            for i,v in enumerate(vals):_cell(pdf,str(v),edges[i],y,edges[i+1]-edges[i],22,size=8)
            y+=22


def _source_font(font):
    family='Title' if 'Cambria' in font else 'Body'
    style=('B' if 'Bold' in font else '')+('I' if 'Italic' in font else '')
    return family,style


def _header(pdf,page,number):
    profile=page['profile'];cohort,degree,window=page['group']
    pdf.set_text_color(0,0,0);pdf.set_fill_color(0,0,0)
    for x0,y0,x1,y1 in profile['header_rects']:pdf.rect(x0,y0,x1-x0,y1-y0,'F')
    for item in profile['header']:
        text=item['text'];family,style=_source_font(item['font']);pdf.set_font(family,style,item['size']);x,y=item['origin']
        if 'SESSION' in text:
            text=f'SESSION -  {cohort}  - {{ SEMESTER -  {_full_term(window)} }}'
            x=(792-pdf.get_string_width(text))/2
        if item.get('center'):x=(792-pdf.get_string_width(text))/2
        color=item.get('color',0)
        pdf.set_text_color((color>>16)&255,(color>>8)&255,color&255)
        if text.strip()=='GRADE':
            with pdf.rotation(90,x=x,y=y):pdf.text(x,y,text)
        else:pdf.text(x,y,text)
    pdf.set_text_color(0,0,0);_text(pdf,str(number),760,4,28,15,size=12)


def _block(pdf,sr,identity,chunk,height,part,y,profile,calculations=None):
    x=profile['edges'];row_h=profile['row_height'];cohort,degree,registration,name=identity
    identity_h=min(4*row_h,height)
    for i,value in enumerate((sr,degree,registration,name+(' (cont.)' if part else ''))):
        _cell(pdf,str(value),x[i],y,x[i+1]-x[i],identity_h,size=7.32,style='B' if i==0 else '')
    if height>identity_h:
        note_y=y+identity_h;note_h=height-identity_h
        pdf.rect(x[0],note_y,x[1]-x[0],note_h);pdf.rect(x[1],note_y,x[4]-x[1],note_h)
        center_x=(x[0]+x[1])/2;center_y=note_y+note_h/2
        pdf.set_font('Body','B',7.32)
        label='NOTE ( If any )'
        with pdf.rotation(90,x=center_x,y=center_y):
            # set_xy treats negative x as an offset from the right margin.
            # Place the short label directly instead of a wide rotated cell.
            pdf.text(center_x-pdf.get_string_width(label)/2,center_y+2.4,label)
    current=y;segments=[]
    for c,h,section in chunk:
        if not segments or segments[-1][0]!=section:segments.append([section,current,h])
        else:segments[-1][2]+=h
        values={4:c.term if c else '',5:c.course_code if c else '-',6:c.course_name if c else '-',7:c.grade if c else '-'}
        for i,value in values.items():_cell(pdf,value,x[i],current,x[i+1]-x[i],h,size=7.32,align='L' if i==6 else 'C')
        current+=h
        adjustment={9,12}
    # Reference adjustment columns have per-row lines; totals are merged per semester segment.
    for i in range(8,len(x)-1):
        if i in adjustment:
            current=y
            for c,h,section in chunk:pdf.rect(x[i],current,x[i+1]-x[i],h);current+=h
        else:
            for section,start,h in segments:
                value=''
                courses=[c for c,_,s in chunk if s==section and c]
                if calculations and courses:
                    from calculation_policy import display_ratio, CalculationReviewRequired
                    term=courses[0].term
                    try:
                        cumulative_column=i==10 or (len(x)==18 and i==13) or (i==15 and profile.get('last_ratio_role','cgpa')=='cgpa')
                        summary=calculations.summary(registration,degree,term,cumulative_column)
                        from fractions import Fraction
                        points=display_ratio(Fraction(summary.total_grade_points)).rstrip('0').rstrip('.')
                        credits=display_ratio(Fraction(summary.total_credit_hours)).rstrip('0').rstrip('.')
                        vals={8:points,10:points,11:credits}
                        if len(x)==16:vals[13]=display_ratio(summary.exact_ratio)
                        else:
                            vals[13]=credits
                            vals[14]=display_ratio(summary.exact_ratio)
                            if i==15 and profile.get('last_ratio_role','cgpa')=='cgpa':
                                vals[15]=display_ratio(summary.exact_ratio)
                        value=vals.get(i,'')
                    except CalculationReviewRequired:
                        if i in ((10,13) if len(x)==16 else (10,13,14,15)):value='\u2014'
                if i==len(x)-2 and calculations and courses:
                    pending=[]
                    for cumulative,label in ((False,'GPA'),(True,'CGPA')):
                        try:calculations.summary(registration,degree,courses[0].term,cumulative)
                        except CalculationReviewRequired:pending.append(label+' REVIEW')
                    value='; '.join(pending)
                _cell(pdf,value,x[i],start,x[i+1]-x[i],h,size=6.5)


def _fee_section(pdf,page,y,known=False):
    x=page['profile']['edges'];edges=(x[0],x[1],x[2],x[6],x[7],x[-1])
    _cell(pdf,'FEE DEFAULTERS',edges[0],y,edges[-1]-edges[0],19,size=9.2,style='B');y+=19
    for i,label in enumerate(('SR','DEGREE ID','REGISTRATION NUMBERS','STUDENT NAME','COMMENTS')):
        _cell(pdf,label,edges[i],y,edges[i+1]-edges[i],19,size=7.32,style='B')
    _cell(pdf,'Fee-status data not supplied',edges[0],y+19,edges[-1]-edges[0],20,size=7.32)


def write_gazette(dataset,rows,path,template_key='computer-science',settings=None):
    from gazette_options import resolve_template_groups
    groups=resolve_template_groups(template_key,rows)
    pdf=FPDF(unit='pt',format=(612,792));pdf.set_auto_page_break(False);pdf.set_margins(0,0,0);_fonts(pdf);pdf.set_line_width(.5)
    from calculation_policy import AcademicIndex
    calculations=AcademicIndex(calculation_attempts(dataset))
    from fee_register import selected_defaulters,write_fee_appendix
    appendices=[]
    publication_term=None
    if settings is not None:
        from gazette_options import current_term
        publication_term=current_term(dataset,settings)
    for key,department_rows in groups:
        fees=selected_defaulters(dataset,department_rows,publication_term)
        withheld={registration_key(r.registration) for r in fees}
        body_rows=[c for c in department_rows if registration_key(c.registration) not in withheld]
        plan=_plan(pdf,body_rows,key,require_cgpa=settings is not None)
        _cover(pdf,department_rows,key,settings);_contents(pdf,plan)
        for number,page in enumerate(plan,1):
            pdf.add_page(orientation='L',format=(612,792));pdf.set_line_width(.84)
            _header(pdf,page,number)
            y=page['profile']['body_top']
            for sr,identity,chunk,height,part in page['blocks']:
                pdf.start_section(identity[2]+(' (cont.)' if part else ''),level=0)
                _block(pdf,sr,identity,chunk,height,part,y,page['profile'],calculations);y+=height
        summary_pages=0
        if any(not page['profile']['cgpa'] for page in plan):
            from semester_reports import write_gazette_cgpa_summary
            summary_pages=write_gazette_cgpa_summary(pdf,body_rows,calculations,len(plan))
        appendices.append((fees,department_rows,key,len(plan)+summary_pages))
    # Every academic section precedes every fee appendix in a combined gazette.
    for fees,department_rows,key,body_pages in appendices:
        write_fee_appendix(pdf,fees,dataset,department_rows,key,body_pages)
    pdf.output(str(path))
