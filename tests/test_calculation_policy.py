from types import SimpleNamespace
from decimal import Decimal, localcontext
from fractions import Fraction
import pytest
from calculation_policy import ApprovedPolicy, calculate_academic_values, CalculationReviewRequired

def test_two_places_round_half_up():
    # Official LGU gazettes round GPAs half-up at 2 decimals (all 2,923
    # totals rows in the three sample PDFs match rounding, none match
    # truncation), so display_ratio rounds instead of truncating.
    from calculation_policy import display_ratio
    with localcontext() as context:
        context.prec=2
        for raw, expected in [('1.8875','1.89'),('1.6000','1.60'),('1.999999999999999999999999999999999','2.00'),('0','0.00'),('4','4.00')]:
            assert display_ratio(Fraction(Decimal(raw))) == expected
        assert display_ratio(Fraction(1,3))=='0.33'

def row(code, credits, points, grade='B', registration='R1'):
    return SimpleNamespace(registration=registration, degree='BSCS', course_code=code, credits=credits, quality_points=points, grade=grade)

def policy(formula='points_over_credits'):
    return ApprovedPolicy('test-1', 'synthetic verification', False, formula, 'points_per_credit')

def test_repeats_keep_matching_credits_and_points_and_other_courses():
    attempts = [row('CS1','3','6'), row('CS1','3','12'), row('CS2','3','9'), row('CS1','3','3',registration='R2')]
    result = calculate_academic_values(attempts, policy())
    first = result['R1','BSCS']
    assert (first.total_credit_hours,first.total_grade_points,first.exact_ratio) == (Decimal(6),Decimal(21),Fraction(7,2))
    assert first.decisions[0].selected is attempts[1]
    assert first.decisions[0].excluded == (attempts[0],)
    assert len(attempts) == 4
    assert result['R2','BSCS'].exact_ratio == 1

def test_direction_and_ambiguous_tie():
    assert calculate_academic_values([row('CS1','3','12')],policy('credits_over_points'))['R1','BSCS'].exact_ratio == Fraction(1,4)
    with pytest.raises(CalculationReviewRequired):
        calculate_academic_values([row('CS1','3','12'),row('CS1','4','16')],policy())

def test_exact_sums_ignore_decimal_context():
    with localcontext() as context:
        context.prec = 2
        result = calculate_academic_values([row('A','3.12345','9.12345'),row('B','2.12345','6.12345')],policy())['R1','BSCS']
        assert result.total_credit_hours == Decimal('5.24690')
        assert result.total_grade_points == Decimal('15.24690')

def test_explicit_letter_scale_and_invalid_values():
    approved = ApprovedPolicy('test-2','synthetic verification',False,'points_over_credits','letter_grade',('F','B','A'))
    attempts = [row('A','3','9','B'),row('A','3','12','A')]
    assert calculate_academic_values(attempts,approved)['R1','BSCS'].decisions[0].selected is attempts[1]
    for credits, points in [('0','0'),('3','NaN'),('3',''),('-3','9')]:
        with pytest.raises(CalculationReviewRequired):
            calculate_academic_values([row('A',credits,points)],policy())
    # v4.6.4 (B4): a lone non-GPA grade is skipped, so the student has no
    # attempts and no GPA is produced (the review surfaces at the summary
    # level as 'No attempts in calculation scope'). A non-GPA grade mixed
    # with graded courses still computes from the graded ones.
    for grade in ('W','I','PASS',''):
        assert calculate_academic_values([row('A','3','0',grade)],policy()) == {}
    # Mixed: the non-GPA attempt is skipped and the GPA comes from the
    # graded course alone.
    mixed = calculate_academic_values([row('A','3','12','A'), row('B','3','0','IP')], policy())
    assert mixed['R1','BSCS'].total_grade_points == Decimal('12')
