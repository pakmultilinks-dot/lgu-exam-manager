"""Gazette choices inferred from imported schema and reference program catalog."""
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


def selection_description(dataset):
    if not dataset:return 'Import a workbook to detect its terms, cohorts and programs.'
    return '\n'.join([
        'Programs: '+', '.join(dataset.degrees),
        'Admission sessions: '+', '.join(sorted({c.admission_session for c in dataset.courses})),
        'Exam terms: '+', '.join(dataset.terms),
        'Headers and cover period follow the selected records. No sample semester or declaration date is inserted.',
    ])

def current_term(dataset,settings):
    from gazette_layout import term_order
    terms=sorted({c.term for c in dataset.courses},key=term_order)
    selected=settings.get('gazette_term')
    if selected:
        if selected in terms and term_order(selected)[0]!=9999:return selected
        raise ValueError('The selected exam term is not available. Choose a term from Settings.')
    valid=[term for term in terms if term_order(term)[0]!=9999]
    if not valid:raise ValueError('Map a valid exam term before selecting the current semester.')
    return valid[-1]

def gazette_rows(dataset,rows,settings):
    term=current_term(dataset,settings)
    selected=[c for c in rows if c.term==term]
    if not selected:raise ValueError('The current semester has no records in this selection. Change Settings or Academic Records filters.')
    return selected
