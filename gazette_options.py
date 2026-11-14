"""Gazette choices inferred from imported schema and reference program catalog.

Reconstructed v4.5.2 module. The source copy shipped in the Drive zip was the
older v4.3 file and did not define the functions v4.5.2 course_ui.py imports,
so this module was rebuilt from the frozen v4.5.2 build (build-v452 PYZ):
helpers that are bytecode-identical to the v4.3 source are kept verbatim,
and the v4.5.2 functions were recovered by disassembling the frozen module
and verifying the result against it with a differential test battery.
"""
from functools import lru_cache
from pathlib import Path
import json

def _template_labels():
    catalog=json.loads((Path(__file__).with_name('gazette_assets')/'reference_layouts.json').read_text(encoding='utf-8'))
    labels={'Automatic from imported data':'auto'}
    for key in catalog['routing']:
        label=key.replace('-',' ').title()+' gazette'
        if key=='computer-science':label='Computer Science gazette'
        if key=='management-science':label='Management Sciences gazette'
        labels[label]=key
    return labels

TEMPLATE_LABELS = _template_labels()

@lru_cache(maxsize=1)
def _programs():
    from gazette_layout import _catalog
    return {key: {program for cohort in routes.values() for program in cohort if program != '*'}
            for key, routes in _catalog()['routing'].items()}

@lru_cache(maxsize=1)
def _routing_rules():
    return json.loads((Path(__file__).with_name('gazette_assets')/'program_routing.json').read_text(encoding='utf-8'))


def canonical_program(degree):
    from academic_identity import program_key
    program=program_key(degree)
    return _routing_rules()['aliases'].get(program,program)


@lru_cache(maxsize=1)
def _department_templates():
    from academic_identity import program_key
    from gazette_layout import _catalog
    labels={}
    for key,routes in _catalog()['routing'].items():
        names={key,key.replace('-',' ')}
        for values in routes.values():
            for ident in values.values():
                names.update(item['text'].strip() for item in _catalog()['profiles'][ident]['header'] if 'DEPARTMENT OF' in item['text'])
        for name in names:
            cleaned=program_key(name.removeprefix('DEPARTMENT OF').strip())
            labels.setdefault(cleaned,set()).add(key)
    return labels


def _department_matches(row):
    from academic_identity import program_key
    value=row.department.strip().upper()
    return _department_templates().get(program_key(value.removeprefix('DEPARTMENT OF').strip()),set())


def resolve_template_groups(choice, rows):
    """Plan all department sections without modifying imported program identities."""
    if choice != 'auto':
        if choice not in TEMPLATE_LABELS.values():raise ValueError('Choose a supported gazette layout.')
        rows=list(rows)
        for degree in {row.degree for row in rows}:
            program=canonical_program(degree)
            allowed={key for key,programs in _programs().items() if program in programs}
            configured=_routing_rules().get('program_templates',{}).get(program)
            if configured:allowed.add(configured)
            if not allowed:allowed.update(_department_matches(next(row for row in rows if row.degree==degree)))
            if allowed and choice not in allowed:
                raise ValueError(f'The selected layout does not match {degree}. Choose Automatic from imported data or a matching department layout.')
        return [(choice,rows)]
    groups={}
    for row in rows:
        program=canonical_program(row.degree)
        matches={key for key, programs in _programs().items() if program in programs}
        configured=_routing_rules().get('program_templates',{}).get(program)
        if configured:
            if configured not in _programs():raise ValueError(f'Invalid configured gazette layout for {row.degree!r}: {configured}')
            matches={configured}
        if not matches:matches=set(_department_matches(row))
        preferred=_routing_rules()['preferred_templates'].get(program)
        if preferred in matches:matches={preferred}
        if len(matches)!=1:
            raise ValueError(f'Program {row.degree!r} is not uniquely matched to a department. Choose its department layout in Settings or supply a department column.')
        key=next(iter(matches));groups.setdefault(key,[]).append(row)
    if not groups:raise ValueError('No current-semester records selected.')
    return list(groups.items())


def resolve_template(choice, rows):
    groups=resolve_template_groups(choice,rows)
    if len(groups)!=1:
        raise ValueError('Selection spans department layouts; use the automatic combined gazette export.')
    return groups[0][0]


def publication_terms(dataset):
    """Return Fall and Spring publication cycles; Summer belongs to Spring."""
    from gazette_layout import term_order
    available={course.term.upper() for course in dataset.courses}
    cycles={term for term in available if term.startswith(('F','S')) and not term.startswith('SU')}
    cycles.update('S'+term[2:] for term in available if term.startswith('SU') and len(term)>2)
    return sorted((term for term in cycles if term_order(term)[0]!=9999), key=term_order)


def publication_term_label(term, dataset):
    import re
    match=re.fullmatch(r'(F|S)(\d{2,4})', term.upper())
    if not match:
        return term
    year=int(match[2])
    if year<100:
        year+=2000
    label=('Fall' if match[1]=='F' else 'Spring')+f' {year} ({term.upper()}'
    if match[1]=='S':
        summer='SU'+match[2]
        if any(course.term.upper()==summer for course in dataset.courses):
            label+=f' + {summer}'
    label+=')'
    return label


def publication_cycle_rows(dataset, term):
    """Rows in the selected cycle; Spring includes its same-year Summer."""
    upper=term.upper()
    terms={upper}
    if upper.startswith('S') and not upper.startswith('SU'):
        terms.add('SU'+upper[1:])
    return [course for course in dataset.courses if course.term.upper() in terms]


def current_term(dataset, settings):
    selected=settings.get('gazette_term')
    if selected:
        upper=selected.upper()
        if upper.startswith('SU'):
            upper='S'+upper[2:]
        if upper in publication_terms(dataset):
            return upper
        raise ValueError('The selected exam term is not available. Choose a term from Settings.')
    valid=publication_terms(dataset)
    if not valid:
        raise ValueError('Map a valid exam term before selecting the current semester.')
    return valid[-1]


def gazette_rows(dataset, rows, settings):
    from gazette_layout import term_order
    term=current_term(dataset, settings)
    cutoff=term_order(term)
    from academic_identity import program_key
    active_terms={term.upper()}
    if term.upper().startswith('S') and not term.upper().startswith('SU'):
        active_terms.add('SU'+term.upper()[1:])
    active_programs={program_key(course.degree) for course in dataset.courses if course.term.upper() in active_terms}
    if not active_programs:
        raise ValueError(f'No programs have results in the selected exam term {term}.')
    selected=[]
    for course in rows:
        if program_key(course.degree) not in active_programs:
            continue
        if term_order(course.term)>cutoff:
            continue
        selected.append(course)
    if term.startswith('S') and not term.startswith('SU'):
        summer='SU'+term[1:]
        selected.extend(c for c in rows if program_key(c.degree) in active_programs and c.term.upper()==summer)
    if not selected:
        raise ValueError('The selected publication cycle has no records in this selection. Change Settings or Academic Records filters.')
    return selected


def available_programs(dataset, settings, *, current_term_only=False):
    """Return imported program values, optionally limited to current-term results."""
    from academic_identity import program_key
    rows=dataset.courses
    if current_term_only:
        rows=publication_cycle_rows(dataset, current_term(dataset, settings))
    labels={}
    for course in rows:
        key=program_key(course.degree)
        if not key:
            continue
        if key not in labels:
            labels[key]={}
        labels[key][course.degree]=labels[key].get(course.degree,0)+1
    programs=[max(options, key=options.get) for options in labels.values()]
    return sorted(programs, key=str.casefold)


def gazette_program_statuses(dataset, settings):
    """Separate programs ready for gazette export from those needing review."""
    from academic_identity import program_key
    from course_reports import capability
    current_term(dataset, settings)
    programs=available_programs(dataset, settings, current_term_only=True)
    statuses={}
    for program in programs:
        identity=program_key(program)
        program_rows=[course for course in dataset.courses if program_key(course.degree)==identity]
        rows=program_rows
        try:
            rows=gazette_rows(dataset, program_rows, settings)
            ready, reason=capability('register', dataset, rows)
            if ready:
                gazette_route_summary(rows)
        except ValueError as exc:
            ready=False
            reason=str(exc)
        statuses[program]=(ready, reason, rows)
    return statuses


def available_gazette_programs(dataset, settings):
    return [program for program, (ready, _, _) in gazette_program_statuses(dataset, settings).items() if ready]


def blocked_gazette_programs(dataset, settings):
    return {program: reason for program, (ready, reason, _) in gazette_program_statuses(dataset, settings).items() if not ready}


def available_template_labels(rows):
    """Offer only actual gazette layouts matched to the imported programs."""
    available=set()
    for degree in {course.degree for course in rows}:
        program=canonical_program(degree)
        matches={key for key, programs in _programs().items() if program in programs}
        configured=_routing_rules().get('program_templates',{}).get(program)
        if configured:
            if configured not in _programs():
                raise ValueError(f'Invalid configured gazette layout for {degree!r}: {configured}')
            matches.add(configured)
        if not matches:
            representative=next(course for course in rows if course.degree==degree)
            matches=set(_department_matches(representative))
        preferred=_routing_rules()['preferred_templates'].get(program)
        if preferred in matches:
            matches={preferred}
        available.update(matches)
    return [label for label, template in TEMPLATE_LABELS.items() if template!='auto' and template in available]


def gazette_route_summary(rows, choice='auto'):
    """Describe the automatic program-to-gazette match without internal jargon."""
    groups=resolve_template_groups(choice, rows)
    labels={key: label.removesuffix(' gazette') for label, key in TEMPLATE_LABELS.items()}
    if len(groups)==1:
        template, matched=groups[0]
        from academic_identity import program_key
        programs=sorted({program_key(course.degree) for course in matched})
        if len(programs)==1:
            program=next(course.degree for course in matched if program_key(course.degree)==programs[0])
            return f"{program} is routed to the {labels.get(template, template.replace('-',' ').title())} Gazette"
        return f"{len(programs)} programs use the {labels.get(template, template.replace('-',' ').title())} Gazette"
    departments=sorted({labels.get(key, key.replace('-',' ').title()) for key, _ in groups})
    return f"{len({course.degree for course in rows})} programs matched to {len(groups)} gazettes: "+', '.join(departments)


def selection_description(dataset, settings=None):
    if not dataset:
        return 'Import a workbook to detect its terms, cohorts and programs.'
    settings=settings or {}
    try:
        term=current_term(dataset, settings)
        programs=available_programs(dataset, settings, current_term_only=True)
    except ValueError:
        programs=[]
        term='unavailable'
    return '\n'.join([
        f"Programs in {publication_term_label(term, dataset)}: "+(', '.join(programs) if programs else 'none'),
        'Admission sessions: '+', '.join(sorted({c.admission_session for c in dataset.courses})),
        'Publication cycles: '+', '.join(publication_term_label(value, dataset) for value in publication_terms(dataset)),
        'The program selector uses programs with results in the selected cycle. Spring includes same-year Summer; Summer is not a separate gazette cycle. No sample semester or declaration date is inserted.',
    ])
