"""Explicit source-based calculation policies; originals are never modified."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from fractions import Fraction
from academic_identity import program_key,registration_key

@dataclass(frozen=True)
class ApprovedPolicy:
    version: str
    approval_reference: str
    golden_cases_verified: bool
    formula: str = ''
    repeat_comparison: str = ''
    grade_order: tuple[str, ...] = ()

class PolicyRequiredError(ValueError): pass
class CalculationReviewRequired(ValueError): pass

_NON_GPA_GRADES = frozenset(('W', 'I', 'IP', 'P', 'PASS', 'R', ''))

def is_status_only_non_gpa(attempt):
    """A recognized non-final status with no source credits or points is not GPA-bearing."""
    return (
        attempt.grade.strip().upper() in _NON_GPA_GRADES
        and not attempt.credits.strip()
        and not attempt.quality_points.strip()
    )

USER_CONFIRMED_POLICY = ApprovedPolicy(
    'LGU-user-2026-10-22-v1',
    'User confirmed points / credits and highest points per credit repeat selection, 2026-10-22',
    False, 'points_over_credits', 'points_per_credit')

def display_ratio(value):
    """Two decimal places, truncated toward zero using exact integer arithmetic."""
    scaled = abs(value.numerator) * 100 // value.denominator
    sign = '-' if value.numerator < 0 and scaled else ''
    return f'{sign}{scaled // 100}.{scaled % 100:02d}'

class AcademicIndex:
    """Use full imported student history, never just filtered display rows."""
    def __init__(self, attempts):
        self.students = {}
        self.cache = {}
        for attempt in attempts:
            self.students.setdefault((registration_key(attempt.registration), program_key(attempt.degree)), []).append(attempt)

    def summary(self, registration, degree, term, cumulative=False):
        from gazette_layout import term_order
        registration=registration_key(registration)
        degree=program_key(degree)
        key = (registration, degree, term, cumulative)
        if key not in self.cache:
            scope = self.students.get((registration, degree), [])
            if cumulative:
                if term_order(term)[0] == 9999 or any(term_order(c.term)[0] == 9999 for c in scope):
                    raise CalculationReviewRequired('Unknown term order; cumulative calculation needs review.')
                scope = [c for c in scope if term_order(c.term) <= term_order(term)]
                current_numbers={int(c.semester_number) for c in scope if c.term==term and c.semester_number.isdigit()}
                if len(current_numbers)>1:
                    raise CalculationReviewRequired('Conflicting academic semester numbers in the current exam term.')
                if current_numbers:
                    through=next(iter(current_numbers))
                    supplied={int(c.semester_number) for c in scope if c.semester_number.isdigit()}
                    if set(range(1,through+1))-supplied:
                        raise CalculationReviewRequired('Earlier academic semesters are missing from imported history.')
            else:
                scope = [c for c in scope if c.term == term]
            values = calculate_academic_values(scope, USER_CONFIRMED_POLICY)
            if (registration, degree) not in values:
                raise CalculationReviewRequired('No attempts in calculation scope.')
            self.cache[key] = values[registration, degree]
        return self.cache[key]

@dataclass(frozen=True)
class RepeatDecision:
    key: tuple[str, str, str]
    selected: object
    excluded: tuple

@dataclass(frozen=True)
class AcademicSummary:
    registration: str
    degree: str
    total_credit_hours: Decimal
    total_grade_points: Decimal
    exact_ratio: Fraction
    decisions: tuple[RepeatDecision, ...]
    policy_version: str

def _values(attempt):
    if attempt.grade.strip().upper() in _NON_GPA_GRADES:
        raise CalculationReviewRequired('Non-GPA or incomplete grade requires an explicit inclusion rule.')
    try:
        credits, points = (Decimal(str(v).strip()) for v in (attempt.credits, attempt.quality_points))
    except InvalidOperation:
        raise CalculationReviewRequired(f'Missing or invalid credit hours / grade points for grade {attempt.grade.strip() or "(blank)"}. Supply numeric GP or an explicit matching scale rate.') from None
    if not all(v.is_finite() and v >= 0 for v in (credits, points)) or credits == 0:
        raise CalculationReviewRequired('Invalid or zero-credit attempt requires review.')
    return credits, points

def _sum_exact(values):
    values = list(values)
    exponent = min((v.as_tuple().exponent for v in values), default=0)
    integers = []
    for value in values:
        parts = value.as_tuple()
        coefficient = int(''.join(map(str, parts.digits)))
        integers.append((-1 if parts.sign else 1) * coefficient * 10 ** (parts.exponent - exponent))
    integer = sum(integers)
    return Decimal((int(integer < 0), tuple(map(int, str(abs(integer)))), exponent))

def calculate_academic_values(attempts, policy=None):
    """Scope is supplied explicitly; exact_ratio has no invented rounding rule."""
    if (policy is None or not policy.version or not policy.approval_reference
            or policy.formula not in ('points_over_credits', 'credits_over_points')
            or policy.repeat_comparison not in ('points_per_credit', 'letter_grade')):
        raise PolicyRequiredError('Confirm formula direction and repeat comparison before calculating.')
    if policy.repeat_comparison == 'letter_grade' and (not policy.grade_order or len(set(policy.grade_order)) != len(policy.grade_order)):
        raise PolicyRequiredError('A unique ascending approved grade scale is required.')
    groups = {}
    for attempt in attempts:
        if is_status_only_non_gpa(attempt):
            continue
        if not attempt.registration.strip() or not attempt.course_code.strip():
            raise CalculationReviewRequired('Registration and course code are required.')
        key = (registration_key(attempt.registration), program_key(attempt.degree), attempt.course_code.strip().casefold())
        groups.setdefault(key, []).append(attempt)
    students = {}
    for key, candidates in groups.items():
        def rank(attempt):
            credits, points = _values(attempt)
            if policy.repeat_comparison == 'points_per_credit':
                return Fraction(points) / Fraction(credits)
            try:
                return policy.grade_order.index(attempt.grade.strip())
            except ValueError:
                raise CalculationReviewRequired('Source grade absent from approved scale.') from None
        ranks = [rank(c) for c in candidates]
        winners = [i for i, r in enumerate(ranks) if r == max(ranks)]
        if len({_values(candidates[i]) for i in winners}) != 1:
            raise CalculationReviewRequired('Equal best grades have conflicting credits or points.')
        selected = winners[0]
        decision = RepeatDecision(key, candidates[selected], tuple(c for i, c in enumerate(candidates) if i != selected))
        students.setdefault(key[:2], []).append(decision)
    output = {}
    for key, decisions in students.items():
        values = [_values(d.selected) for d in decisions]
        credits = _sum_exact(c for c, _ in values)
        points = _sum_exact(p for _, p in values)
        numerator, denominator = (points, credits) if policy.formula == 'points_over_credits' else (credits, points)
        if denominator == 0:
            raise CalculationReviewRequired('Selected formula has a zero denominator.')
        output[key] = AcademicSummary(*key, credits, points, Fraction(numerator) / Fraction(denominator), tuple(decisions), policy.version)
    return output


def parse_grade_scale(text,reference):
    if text.strip() and not reference.strip():
        raise ValueError('Enter the source/reference for the grade scale you are applying.')
    values={}
    for line in text.splitlines():
        if not line.strip():continue
        parts=line.split('=')
        if len(parts)!=2 or not parts[0].strip():raise ValueError('Use one GRADE=points-per-credit entry per line.')
        grade=parts[0].strip().upper()
        try:value=Decimal(parts[1].strip())
        except InvalidOperation:raise ValueError('Every grade point rate must be a number.') from None
        if not value.is_finite() or value<0:raise ValueError('Grade point rates must be finite and nonnegative.')
        if len(value.as_tuple().digits)>20 or not -12<=value.as_tuple().exponent<=12:
            raise ValueError('Grade point rates support at most 20 digits and 12 decimal/exponent places.')
        if grade in values:raise ValueError('Each grade may occur only once.')
        values[grade]=str(value)
    return values


def calculation_attempts(dataset):
    """Fill absent numeric points only from an explicitly entered scale; source stays intact."""
    from dataclasses import replace
    from decimal import localcontext
    scale=getattr(dataset,'grade_scale',{})
    output=[]
    for course in dataset.courses:
        rate=getattr(course,'grade_point_rate','').strip() or scale.get(course.grade.strip().upper())
        if not course.quality_points.strip() and rate is not None:
            try:
                credits=Decimal(course.credits);points=Decimal(rate)
                if credits.is_finite() and credits>0 and points.is_finite() and points>=0:
                    with localcontext() as context:
                        context.prec=max(50,len(credits.as_tuple().digits)+len(points.as_tuple().digits)+2)
                        course=replace(course,quality_points=str(credits*points))
            except InvalidOperation:pass
        output.append(course)
    return output
