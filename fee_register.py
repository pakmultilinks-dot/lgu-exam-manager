"""Fee blocks close each session section, as in the official gazettes."""
from semester_reports import roll_order
from academic_identity import registration_key

def selected_defaulters(dataset,rows,publication_term=None):
    # v4.5.2: the v4.3 copy took only (dataset, rows) and crashed the Excel/Word
    # register export, which passes the publication term. A fee record counts
    # only if its term is in the publication cycle (Spring includes same-year
    # Summer); when the cycle has no fee record, a term-less record applies.
    from gazette_layout import term_order
    registrations={registration_key(c.registration) for c in rows}
    available_terms={c.term for c in rows}
    if publication_term is None:
        publication_term=max(available_terms,key=term_order) if available_terms else None
    cycle_terms=set()
    if publication_term:
        cycle_terms.add(publication_term)
        upper=publication_term.upper()
        if upper.startswith('S') and not upper.startswith('SU'):
            paired='SU'+upper[1:]
            if paired in available_terms:
                cycle_terms.add(paired)
        elif upper.startswith('SU'):
            paired='S'+upper[2:]
            if paired in available_terms:
                cycle_terms.add(paired)
    by_registration={}
    for record in dataset.fee_records:
        registration=registration_key(record.registration)
        if registration in registrations:
            by_registration.setdefault(registration,[]).append(record)
    selected=[]
    for registration,records in by_registration.items():
        cycle_records=[record for record in records if record.term in cycle_terms]
        if cycle_records:
            latest=max((record.term for record in cycle_records),key=term_order)
            applicable=[record for record in cycle_records if record.term==latest]
        else:
            applicable=[record for record in records if not record.term]
        if len({record.status for record in applicable})>1:
            raise ValueError('Conflicting fee statuses for a selected registration; reconcile source fee records before export.')
        if applicable and applicable[0].status=='DEFAULTER':
            selected.append(applicable[0])
    return sorted(selected,key=lambda r:roll_order(r.registration))

def _fee_chunks(pdf,fees,profile):
    # An empty fee list still yields one chunk: the official gazettes print
    # the title and column header with no entries and no message row.
    from gazette_layout import _wrap
    x=profile['edges'];edges=(x[0],x[1],x[2],x[6],x[7],x[-1])
    capacity=20*profile['row_height']-50
    pages=[];chunk=[];height=0
    for number,record in enumerate(fees,1):
        values=(number,record.degree,record.registration,record.name,record.comment)
        row_height=max(22,max((len(_wrap(pdf,str(value),edges[i+1]-edges[i]-4,size=8))*9.1+4 for i,value in enumerate(values)),default=22))
        if row_height>capacity:raise ValueError('A fee-list entry is too long for one page. Shorten its publication comment; the full source remains in the audit.')
        if chunk and height+row_height>capacity:
            pages.append(chunk);chunk=[];height=0
        chunk.append((values,row_height));height+=row_height
    if chunk:pages.append(chunk)
    if not pages:pages=[[]]
    return pages

def fee_page_count(pdf,fees,template_key,cohort,degree):
    # How many pages the fee block for one (cohort, degree) section takes,
    # without rendering it. Used to number the table of contents before the
    # body pages are drawn.
    from gazette_layout import _profile
    return len(_fee_chunks(pdf,fees,_profile(template_key,cohort,degree,require_cgpa=True)))

def write_fee_appendix(pdf,fees,template_key,body_pages,cohort,degree,publication_term=None):
    # Render the fee block for one (cohort, degree) section, at the end of
    # that section's result pages. Returns the number of pages written.
    # The header shows that section's session. Withholding is unchanged:
    # defaulters get no result block (the caller removes them from the
    # body rows before planning).
    from gazette_layout import _profile,_cell,_header
    # The fee page closes the section, so it uses the same 17-column
    # CGPA header (T.CH + CGPA) as the body pages, not the 15-column
    # header the plain profile would give for early cohorts.
    profile=_profile(template_key,cohort,degree,require_cgpa=True)
    start=profile['edges'][0];end=profile['edges'][-1];width=end-start
    x=profile['edges'];edges=(x[0],x[1],x[2],x[6],x[7],x[-1])
    window=publication_term or (fees[0].term if fees else '')
    pages=_fee_chunks(pdf,fees,profile)
    for page_number,chunk in enumerate(pages,1):
        pdf.add_page(orientation='L',format=(612,792))
        page={'profile':profile,'group':(cohort,degree,window)}
        _header(pdf,page,body_pages+page_number)
        y=profile['body_top']
        _cell(pdf,'FEE DEFAULTERS',start,y,width,25,size=11,style='B');y+=25
        for i,label in enumerate(('SR','DEGREE ID','REGISTRATION NUMBERS','STUDENT NAME','COMMENTS')):
            _cell(pdf,label,edges[i],y,edges[i+1]-edges[i],25,size=8,style='B')
        y+=25
        for values,row_height in chunk:
            for column,value in enumerate(values):
                _cell(pdf,str(value),edges[column],y,edges[column+1]-edges[column],row_height,size=8)
            y+=row_height
    return len(pages)
