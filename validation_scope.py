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
    from gazette_layout import term_order
    from academic_engine import derive_term
    selected={}
    for course in rows:
        selected.setdefault((registration_key(course.registration),program_key(course.degree)),set()).add(course.term)
    result=[]
    for issue in dataset.issues:
        if issue.level!='ERROR':continue
        source,fields=issue_source(dataset,issue)
        identity=registration_key(fields.get('registration',''))
        candidates=[terms for (reg,degree),terms in selected.items() if reg==identity and
                    (not fields.get('degree') or degree==program_key(fields['degree']) or 'CONFLICT' in issue.code)]
        if not candidates:
            # Synthetic/legacy sessions may have no raw row audit. Fail closed there.
            if source is None:result.append(issue)
            continue
        if 'CONFLICT' in issue.code:result.append(issue);continue
        raw=fields.get('term') or fields.get('semester','')
        term=raw.upper() if term_order(raw.upper())[0]!=9999 else derive_term(raw)
        terms=set().union(*candidates)
        if all_terms or term_order(term)[0]==9999 or term in terms or (history and term_order(term)<=max(map(term_order,terms))):
            result.append(issue)
    return result
