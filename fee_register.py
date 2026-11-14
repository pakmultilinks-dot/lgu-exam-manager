"""Verified fee records are appended after all academic result pages."""
from semester_reports import roll_order
from academic_identity import registration_key

def selected_defaulters(dataset,rows):
    registrations={registration_key(c.registration) for c in rows};terms={c.term for c in rows}
    grouped={}
    for record in dataset.fee_records:
        if registration_key(record.registration) in registrations and (not record.term or record.term in terms):
            grouped.setdefault(registration_key(record.registration),[]).append(record)
    selected=[]
    for registration,records in grouped.items():
        if len({r.status for r in records})!=1:
            raise ValueError('Conflicting fee statuses for a selected registration; reconcile source fee records before export.')
        if records[0].status=='DEFAULTER':selected.append(records[0])
    return sorted(selected,key=lambda r:roll_order(r.registration))

def write_fee_appendix(pdf,fees,dataset,rows,template_key,body_pages):
    from gazette_layout import _profile,_cell,_header,_wrap,_with_department_header
    # No invented defaulters or "all clear" assertion when fee data is absent.
    example=rows[0];profile=_profile(template_key,example.admission_session,example.degree)
    headings={item['text'].strip() for cohort,degree in {(course.admission_session,course.degree) for course in rows} for item in _profile(template_key,cohort,degree)['header'] if 'DEPARTMENT OF' in item['text']}
    if len(headings)>1:
        profile=_with_department_header(profile,'DEPARTMENT OF '+template_key.replace('-',' ').upper())
    start=profile['edges'][0];end=profile['edges'][-1];width=end-start
    x=profile['edges'];edges=(x[0],x[1],x[2],x[6],x[7],x[-1])
    entries=fees or [None]
    capacity=20*profile['row_height']-50
    pages=[];chunk=[];height=0
    for number,record in enumerate(entries,1):
        values=(number,record.degree,record.registration,record.name,record.comment) if record else None
        row_height=max(22,max((len(_wrap(pdf,str(value),edges[i+1]-edges[i]-4,size=8))*9.1+4 for i,value in enumerate(values or [])),default=22))
        if row_height>capacity:raise ValueError('A fee-list entry is too long for one page. Shorten its publication comment; the full source remains in the audit.')
        if chunk and height+row_height>capacity:
            pages.append(chunk);chunk=[];height=0
        chunk.append((values,row_height));height+=row_height
    if chunk:pages.append(chunk)
    for page_number,chunk in enumerate(pages,1):
        pdf.add_page(orientation='L',format=(612,792))
        page={'profile':profile,'group':(example.admission_session,example.degree,example.term)}
        _header(pdf,page,body_pages+page_number)
        y=profile['body_top']
        _cell(pdf,'FEE DEFAULTERS',start,y,width,25,size=11,style='B');y+=25
        for i,label in enumerate(('SR','DEGREE ID','REGISTRATION NUMBERS','STUDENT NAME','COMMENTS')):
            _cell(pdf,label,edges[i],y,edges[i+1]-edges[i],25,size=8,style='B')
        y+=25
        for values,row_height in chunk:
            if values is None:
                _cell(pdf,'No fee-status data supplied for this selection' if not dataset.fee_records else 'No defaulter records supplied for this selection',start,y,width,24,size=8)
            else:
                for column,value in enumerate(values):
                    _cell(pdf,str(value),edges[column],y,edges[column+1]-edges[column],row_height,size=8)
            y+=row_height
