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

USER_CONFIRMED_POLICY = ApprovedPolicy(
    'LGU-user-2026-10-22-v1',
    'User confirmed points / credits and highest points per credit repeat selection, 2026-10-22',
    False, 'points_over_credits', 'points_per_credit')

def display_ratio(value):
    """Two decimal places, rounded half-up using exact integer arithmetic.

    The official LGU gazettes round half-up: all 2,923 totals rows in the
    three sample PDFs match rounding and none match truncation, so the
    display follows the official format.
    """
    scaled = abs(value) * 100
    whole, remainder = divmod(scaled.numerator, scaled.denominator)
    if remainder * 2 >= scaled.denominator:
        whole += 1
    sign = '-' if value < 0 and whole else ''
    return f'{sign}{whole // 100}.{whole % 100:02d}'

def standing_cutoff(cutoff_text=None):
    """The CGPA line for academic standing, as an exact fraction.

    The operator's probation_cgpa setting wins when it is a number. When
    it is empty (the default), the line is 2.00, the usual university
    probation line: CGPA at or above 2.00 is good standing, below it is
    probation. The cutoff is always stated where the standing is shown,
    so a different university line is a settings change, not a guess.
    """
    if cutoff_text is not None and str(cutoff_text).strip() != '':
        try:
            return Fraction(Decimal(str(cutoff_text).strip()))
        except (InvalidOperation, ValueError, ArithmeticError):
            pass
    return Fraction(2, 1)


def standing_comment(cgpa_exact_ratio, cutoff_text=None):
    """The gazette COMMENTS standing for one student.

    Returns 'G.A.ST' (good academic standing) when the CGPA is at or above
    the probation cutoff, else 'PROBATION'. This is the value the official
    LGU gazettes print in COMMENTS (G.A.ST for most students, PROBATION
    for some); the app fills it instead of leaving COMMENTS empty.
    """
    return 'G.A.ST' if cgpa_exact_ratio >= standing_cutoff(cutoff_text) else 'PROBATION'


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

def _clean_number(value):
    """Recover the intended 2-decimal value from binary float noise.

    Source grade points come from the workbook as floats; a value the
    university meant as 11.10 can arrive as 11.100000000000001, and
    9.90 as 9.899999999999999. That noise is far below a real hundredth,
    but it flips half-up rounding when the true ratio sits exactly on a
    .xx5 boundary (32 SGPA/CGPA values in F20.xlsx printed 0.01 low).
    Only a value within 1e-9 of a clean 2-decimal number is normalised;
    a genuine 3-decimal value is left untouched. The source workbook is
    never changed; only the in-memory value used for the ratio is cleaned.
    """
    from decimal import ROUND_HALF_UP, localcontext
    if not value.is_finite():
        return value
    if value.as_tuple().exponent < -2:
        try:
            with localcontext() as context:
                context.prec = max(28, len(value.as_tuple().digits) + 2)
                clean = value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            return value
        if abs(value - clean) < Decimal('0.000000001'):
            return clean
    return value


def _values(attempt):
    if attempt.grade.strip().upper() in ('W', 'I', 'IP', 'P', 'PASS', 'R', 'U', ''):
        raise CalculationReviewRequired('Non-GPA or incomplete grade requires an explicit inclusion rule.')
    try:
        credits, points = (_clean_number(Decimal(str(v).strip())) for v in (attempt.credits, attempt.quality_points))
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
        if not attempt.registration.strip() or not attempt.course_code.strip():
            raise CalculationReviewRequired('Registration and course code are required.')
        # A non-GPA grade (in progress, incomplete, withdrawn, pass,
        # unsatisfactory, or blank) has no grade points and no credits to
        # count, so it is
        # skipped instead of blocking the student's whole GPA. One IP
        # course used to poison the SGPA and every later CGPA (SARA
        # FAIZA F25, SAWERA F24). A student with only non-GPA attempts
        # still gets a review, via 'No attempts in calculation scope'.
        if attempt.grade.strip().upper() in ('W', 'I', 'IP', 'P', 'PASS', 'R', 'U', ''):
            continue
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
        # Failed (F) attempts keep their grade point of 0 and their credit
        # hours stay in the printed totals, but their credit hours are
        # excluded from the GPA/CGPA denominator, matching the official LGU
        # gazettes (e.g. DANISH WAKEEL: 27.00/12 prints 3.00 = 27.00/9).
        # A student whose selected attempts are all F has a zero
        # denominator; the GPA is 0.00 (no grade points were earned), not a
        # review, since a student who failed everything is a real outcome
        # and must not block the gazette.
        paired = [(_values(d.selected), d.selected) for d in decisions]
        credits = _sum_exact(c for (c, _), _ in paired)
        points = _sum_exact(p for (_, p), _ in paired)
        gpa_credits = _sum_exact(c for (c, _), attempt in paired if attempt.grade.strip().upper() != 'F')
        numerator, denominator = (points, gpa_credits) if policy.formula == 'points_over_credits' else (gpa_credits, points)
        ratio = Fraction(0) if denominator == 0 else Fraction(numerator) / Fraction(denominator)
        output[key] = AcademicSummary(*key, credits, points, ratio, tuple(decisions), policy.version)
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
