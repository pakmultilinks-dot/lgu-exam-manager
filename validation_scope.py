"""Report validation follows selected identities, including cumulative history."""
from academic_identity import registration_key, program_key


def issue_source(dataset, issue):
    key=(issue.source_file or dataset.source_filename,issue.sheet,issue.excel_row)
    cache=getattr(dataset,'_review_source_index',None)
    if cache is None:
        cache={(row.get('file',dataset.source_filename),row['sheet'],row['excel_row']):row for row in dataset.source_rows}
        dataset._review_source_index=cache
    source=cache.get(key)
    if source is None and not issue.source_file:
        matches=[row for (_,sheet,number),row in cache.items() if sheet==issue.sheet and number==issue.excel_row]
        if len(matches)==1:source=matches[0]
    layout=next((item for item in dataset.layouts if item.sheet==issue.sheet and
                 (not item.source_file or not source or item.source_file==source.get('file'))),None)
    fields={role:str(source['cells'].get(column) if source['cells'].get(column) is not None else '').strip() for role,column in layout.columns.items()} if source and layout else {}
    if layout:
        for role,value in layout.defaults.items():
            if role not in layout.columns:fields[role]=str(value)
    return source,fields


def relevant_errors(dataset, rows, *, history=True, all_terms=False):
    """Return errors linked to selected registrations or identifiable programs.

    Incomplete source rows may have no registration number; when their mapped
    program is available, scope them to that program. Unidentifiable errors fail
    closed instead of silently allowing an export.
    """
    from gazette_layout import term_order
    from academic_engine import derive_term
    from academic_engine import registration_program
    selected={}
    selected_by_program={}
    for course in rows:
        registration=registration_key(course.registration)
        degree=program_key(course.degree)
        selected.setdefault((registration,degree),set()).add(course.term)
        selected_by_program.setdefault(degree,set()).add(course.term)
    selected_programs=set(selected_by_program)-{''}
    result=[]
    for issue in dataset.issues:
        if issue.level!='ERROR':continue
        source,fields=issue_source(dataset,issue)
        identity=registration_key(fields.get('registration',''))
        issue_program=program_key(fields.get('degree',''))
        if not issue_program:
            issue_program=program_key(registration_program(fields.get('registration','')))
        program_terms=selected_by_program.get(issue_program,set())
        registration_terms=set().union(*(terms for (registration,degree),terms in selected.items()
                                        if identity and registration==identity
                                        and (not issue_program or degree==issue_program))) if identity else set()
        if source is None:
            result.append(issue)
            continue
        if identity:
            if not registration_terms:continue
            applicable_terms=registration_terms
        elif issue_program:
            if issue_program not in selected_programs:continue
            applicable_terms=program_terms
        else:
            result.append(issue)
            continue
        raw=fields.get('term') or fields.get('semester','')
        term=raw.upper() if term_order(raw.upper())[0]!=9999 else derive_term(raw)
        exact_term=term in applicable_terms
        before_cutoff=history and applicable_terms and term_order(term)<=max(map(term_order,applicable_terms))
        if all_terms or term_order(term)[0]==9999 or exact_term or before_cutoff:
            result.append(issue)
    return result
