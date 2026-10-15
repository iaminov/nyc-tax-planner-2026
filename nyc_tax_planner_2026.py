#!/usr/bin/env python3
"""2026 NYC sole-proprietor tax / Solo 401(k) / HSA / ACA planner.

Run: python nyc_tax_planner_2026.py
Test: python nyc_tax_planner_2026.py --test

STANDARD LIBRARY ONLY (Tkinter). Planning model, not a substitute for tax software.
See the Method tab for supported cases, assumptions, and official sources.
"""
from __future__ import annotations

import json
import math
import os
import shutil
import sys
import tkinter as tk
from dataclasses import asdict, dataclass, field
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import Any

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

YEAR = 2026
# Single filer (no dependents); 2025 FPL is used for 2026 ACA tax-credit purposes.
FPL = 15650.0
FED_STANDARD = 16100.0
NY_STANDARD = 8000.0
SS_WAGE_BASE = 184500.0
SOLO_EMPLOYEE_LIMIT = 24500.0
SOLO_TOTAL_LIMIT = 72000.0
SOLO_CATCHUP = 8000.0
SOLO_CATCHUP_60_63 = 11250.0
HSA_SELF = 4400.0
HSA_FAMILY = 8750.0
HSA_CATCHUP = 1000.0
# 2026 IRS single-filer tax rates, expressed as cumulative endpoints.
FED_BRACKETS = [(12400, .10), (50400, .12), (105700, .22),
                (201775, .24), (256225, .32), (640600, .35), (math.inf, .37)]
# NYS 2026 enacted rate reductions; the progressive base omits NY's special
# high-income benefit recapture/tax computations. Flag high-income estimates.
NY_BRACKETS = [(8500, .039), (11700, .044), (13900, .0515),
               (80650, .054), (215400, .059), (1077550, .0685),
               (5000000, .0965), (25000000, .103), (math.inf, .109)]
NYC_BRACKETS = [(12000, .03078), (25000, .03762),
                (50000, .03819), (math.inf, .03876)]
LTCG_BRACKETS = [(49450, .0), (545500, .15), (math.inf, .20)]
PTC_BANDS = [(0, 133, 2.10, 2.10), (133, 150, 3.14, 4.19),
             (150, 200, 4.19, 6.60), (200, 250, 6.60, 8.44),
             (250, 300, 8.44, 9.96), (300, 400, 9.96, 9.96)]
SOURCES = [
    ('2026 federal brackets', 'https://www.irs.gov/filing/federal-income-tax-rates-and-brackets'),
    ('Solo 401(k) contribution rules', 'https://www.irs.gov/retirement-plans/one-participant-401k-plans'),
    ('2026 401(k) limits', 'https://www.irs.gov/retirement-plans/plan-participant-employee/retirement-topics-401k-and-profit-sharing-plan-contribution-limits'),
    ('2026 premium tax credit percentages', 'https://www.irs.gov/irb/2025-32_IRB'),
    ('2026 PTC repayment caps removed', 'https://www.irs.gov/tax-professionals/eitc-central/repayment-caps-eliminated-for-excess-advance-payments-of-the-premium-tax-credit'),
    ('ACA + self-employed health insurance, Pub. 974', 'https://www.irs.gov/publications/p974'),
    ('Self-employed health insurance, Form 7206', 'https://www.irs.gov/pub/irs-pdf/f7206.pdf'),
    ('2026 HSA Bronze rules', 'https://www.irs.gov/irb/2026-02_IRB'),
    ('HSA partial year/last-month rule', 'https://www.irs.gov/publications/p969'),
    ('2026 MCTMT rates', 'https://www.tax.ny.gov/legal/2025/pit-corp-changes.htm'),
    ('2026 MCTMT threshold', 'https://www.tax.ny.gov/pit/file/nyc_yonkers_residents.htm'),
    ('2026 NY tax and standard deduction', 'https://www.tax.ny.gov/forms/current-forms/it/it2104i.htm'),
    ('NYC resident tax brackets', 'https://www.tax.ny.gov/forms/current-forms/it/it201i.htm'),
    ('NYC UBT overview', 'https://www.nyc.gov/site/finance/business/business-unincorporated-business-tax-ubt.page'),
    ('NYC UBT 2026 estimated-tax form', 'https://www.nyc.gov/assets/finance/downloads/pdf/26pdf/business_tax_forms/nyc-5ubti_2026.pdf'),
    ('NYC unincorporated business credit (2026)', 'https://codelibrary.amlegal.com/codes/newyorkcity/latest/NYCadmin/0-0-0-13608'),
    ('NY estimated tax deadlines', 'https://www.tax.ny.gov/pit/estimated_tax/estimated_tax_payment_due_dates.htm'),
    ('NY Essential Plan transition', 'https://ag.ny.gov/press-release/2026/attorney-general-james-issues-guidance-new-yorkers-facing-health-insurance'),
]


def money(x: float) -> str:
    return f'${x:,.2f}'


def progressive(amount: float, brackets: list[tuple[float, float]]) -> float:
    amount = max(0., amount)
    tax = 0.
    previous = 0.
    for ceiling, rate in brackets:
        take = max(0., min(amount, ceiling) - previous)
        tax += take * rate
        previous = ceiling
        if amount <= ceiling:
            break
    return tax


def ptc_percentage(magi: float) -> float:
    """Applicable percentage expressed as a FRACTION, not percent."""
    fpl_pct = 100.0 * magi / FPL
    if fpl_pct > 400.0 + 1e-10:
        return math.inf
    if fpl_pct < 133:
        return .0210
    for start, stop, low, high in PTC_BANDS[1:]:
        if fpl_pct <= stop:
            return (low + (high-low) * (fpl_pct-start)/(stop-start))/100.0
    return .0996


def ltcg_tax(ordinary_taxable: float, preferred: float) -> float:
    """Net long-term gains and qualified dividends stack on ordinary taxable income."""
    if preferred <= 0:
        return 0.
    return progressive(ordinary_taxable+preferred, LTCG_BRACKETS) - progressive(ordinary_taxable, LTCG_BRACKETS)


@dataclass
class Inputs:
    # Monetary entries are ANNUAL unless indicated as monthly.
    gross_business: float = 100000.0
    business_expenses: float = 0.0
    w2_box1: float = 0.0
    w2_social_security_wages: float = 0.0
    w2_medicare_wages: float = 0.0
    w2_additional_medicare_withheld: float = 0.0
    w2_prior_401k_deferrals: float = 0.0
    short_gains: float = 0.0
    long_gains: float = 0.0
    qualified_dividends: float = 0.0
    interest_other_ordinary: float = 0.0
    aca_magi_addbacks: float = 0.0
    other_agi_adjustments: float = 0.0
    age: int = 37
    solo_actual: float = 0.0
    hsa_actual: float = 0.0
    employer_hsa: float = 0.0  # includes pre-tax payroll salary reduction HSA
    hsa_family: bool = False
    hsa_first_month: int = 1
    hsa_last_month: int = 12
    use_hsa_last_month_rule: bool = False
    aca_first_month: int = 1
    aca_last_month: int = 12
    benchmark_monthly: float = 868.20
    enrollment_monthly: float = 798.57
    aptc_monthly: float = 0.0
    self_emp_coverage_eligible: bool = True
    qhp_ptc_eligible: bool = True
    qbi_enabled: bool = True
    mct_zone1_percent: float = 100.0
    mct_zone2_percent: float = 0.0
    ubt_auto_estimate: bool = True
    ubt_city_net_income_adjustment: float = 0.0
    nyc_ubt_annual_manual: float = 0.0
    nyc_other_tax_credits: float = 0.0
    fed_withholding: float = 0.0
    nys_withholding: float = 0.0
    nyc_withholding: float = 0.0
    federal_estimates_already_paid: float = 0.0
    installments_remaining: int = 1  # October 2026: only Jan 15 2027 still due
    ny_estimates_already_paid: float = 0.0
    ubt_estimates_already_paid: float = 0.0
    federal_standard_override: float = FED_STANDARD
    ny_standard_override: float = NY_STANDARD
    tax_year: int = 2026
    mct_threshold_override: float = 0.0


@dataclass
class Result:
    values: dict[str, float] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    hsa_schedule: list[tuple[int, str, float]] = field(default_factory=list)
    ubt_form: str = "NYC-202S (Simplified)"


def validate(i: Inputs) -> None:
    for field_name, value in asdict(i).items():
        if isinstance(value, (int,float)) and not isinstance(value, bool):
            if not math.isfinite(value):
                raise ValueError(f'{field_name}: enter a finite number.')
            if field_name not in ('short_gains', 'long_gains', 'interest_other_ordinary', 'ubt_city_net_income_adjustment') and value < 0:
                raise ValueError(f'{field_name}: negative amounts are unsupported.')
    if not 18 <= i.age <= 110:
        raise ValueError('Age must be between 18 and 110.')
    if i.tax_year not in (2025, 2026):
        raise ValueError('Tax year must be 2025 or 2026.')
    if i.mct_threshold_override < 0:
        raise ValueError('MCTMT threshold override cannot be negative.')
    if i.installments_remaining not in (1, 2, 3, 4):
        raise ValueError('Remaining estimated-tax installments must be an integer from 1 through 4.')
    for name in ('hsa_first_month', 'hsa_last_month', 'aca_first_month', 'aca_last_month'):
        if not (1 <= getattr(i, name) <= 12):
            raise ValueError(f'{name} must be between 1 and 12.')
    if i.hsa_last_month < i.hsa_first_month or i.aca_last_month < i.aca_first_month:
        raise ValueError('Ending month must be on or after starting month.')
    if i.mct_zone1_percent + i.mct_zone2_percent > 100.000001:
        raise ValueError('MCT Zone 1 + Zone 2 business allocations cannot exceed 100%.')
    if i.short_gains < 0 or i.long_gains < 0:
        raise ValueError('Negative capital gains/loss netting needs Schedule D and is unsupported in this simplified planner. Do not enter a negative capital gain.')
    if i.hsa_actual < 0 or i.solo_actual < 0:
        raise ValueError('Contributions cannot be negative.')


def hsa_allowance(i: Inputs) -> tuple[float, list[tuple[int, str, float]], list[str]]:
    annual = HSA_FAMILY if i.hsa_family else HSA_SELF
    if i.age >= 55:
        annual += HSA_CATCHUP
    monthly = annual / 12.
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    schedule = []
    notes = []
    count = 0
    for idx in range(1, 13):
        eligible = i.hsa_first_month <= idx <= i.hsa_last_month
        count += bool(eligible)
        schedule.append((idx, months[idx-1]+(' ✓' if eligible else ' —'), monthly if eligible else 0.0))
    prorated = count * monthly
    if i.use_hsa_last_month_rule:
        if i.hsa_first_month <= 12 <= i.hsa_last_month:
            notes.append('HSA last-month rule used. To keep the additional deduction, remain HSA-eligible through December 31, 2027 (testing period).')
            return annual, schedule, notes
        notes.append('Last-month rule ignored: you were not HSA-eligible on December 1.')
    return prorated, schedule, notes


def se_tax_and_half(profit: float, w2_ss: float) -> tuple[float, float, float]:
    net_se = max(0., profit * .9235)
    if net_se < 400.:
        return 0., 0., net_se  # Schedule SE general $400 net-earnings filing threshold
    ss_remaining = max(0., SS_WAGE_BASE - w2_ss)
    ss_tax = min(net_se, ss_remaining) * .124
    medicare = net_se * .029
    total = ss_tax + medicare
    return total, .5*total, net_se


def solo_limits(i: Inputs, earned: float) -> tuple[float,float,float]:
    earned = max(0., earned)
    catchup = (SOLO_CATCHUP_60_63 if 60 <= i.age <= 63 else SOLO_CATCHUP if i.age>=50 else 0.)
    employee = min(earned, max(0., SOLO_EMPLOYEE_LIMIT + catchup - i.w2_prior_401k_deferrals))
    # Non-elective employer rate for an unincorporated sole proprietor: 20% of adjusted earnings.
    employer = .2 * earned
    # Total additions excluding catch-up not above 72k; total amount also <= earnings.
    # Split basic deferral and additional catch-up for the 415(c) test.
    regular_employee = min(employee, max(0., SOLO_EMPLOYEE_LIMIT-i.w2_prior_401k_deferrals))
    available_employer = max(0., min(employer, SOLO_TOTAL_LIMIT-regular_employee, earned-employee))
    maximum = employee + available_employer
    return maximum, employee, available_employer


def ptc_for_magi(i: Inputs, magi: float) -> float:
    if not i.qhp_ptc_eligible or magi < FPL or magi > 4*FPL:
        return 0.
    n = i.aca_last_month-i.aca_first_month+1
    annual_expected_contribution = max(0., magi)*ptc_percentage(magi)
    monthly_credit = max(0., i.benchmark_monthly - annual_expected_contribution/12.)
    return n*min(i.enrollment_monthly, monthly_credit)


def solve_health_deduction(i: Inputs, agi_before_hi: float, cap: float) -> tuple[float, float, str]:
    """Find max feasible deduction D: D <= earnings limit and D+PTC(MAGI(D))<=premiums.

    Not simply fixed-point iteration: discontinuities at 400% FPL may break it.
    Search subintervals separated by 100%, 133%, and 400% FPL and keep feasible endpoints.
    The solver is for one self-employed Marketplace policy and independent HSA.
    It does not implement every multiple-business/coverage or filing situation in Pub.974.
    """
    months = i.aca_last_month-i.aca_first_month+1
    premiums = i.enrollment_monthly * months
    if not i.self_emp_coverage_eligible or cap <= 0 or premiums <= 0:
        return 0., ptc_for_magi(i, agi_before_hi+i.aca_magi_addbacks), 'No eligible SE health deduction'
    hi_max = min(cap, premiums)
    def credit(d: float) -> float:
        return ptc_for_magi(i, agi_before_hi-d+i.aca_magi_addbacks)
    def room(d: float) -> float:
        return premiums - d - credit(d)
    # All continuous regions are monotone (D + C(D) non-decreasing).
    # At the 400% threshold the credit may jump upward. At 100%, it can drop;
    # check both sides of these boundaries. Carry exact-feasible positions only.
    cuts = [0., hi_max]
    base_magi = agi_before_hi+i.aca_magi_addbacks
    for threshold in [FPL, 1.33*FPL, 4*FPL]:
        cut = base_magi-threshold
        if 0 < cut < hi_max:
            cuts.extend([max(0.,cut-0.0001), cut, min(hi_max,cut+0.0001)])
    cuts = sorted(set(cuts))
    feasible = [0.]
    for left,right in zip(cuts[:-1],cuts[1:]):
        if room(left) >= -1e-7:
            feasible.append(left)
        if room(right) >= -1e-7:
            feasible.append(right)
            continue
        if room(left) < -1e-7:
            continue
        lo,hi = left,right
        for _ in range(75):
            mid = (lo+hi)/2
            if room(mid)>=0:
                lo=mid
            else:
                hi=mid
        feasible.append(lo)
    deduction = max(feasible)
    subsidy = credit(deduction)
    # Reserve a fraction of a cent against crossing the constraint from floating-point.
    if deduction+subsidy > premiums + 1e-7:
        deduction = max(0., deduction-0.01)
        subsidy = credit(deduction)
    method = 'Constraint-based piecewise bisection; D + final PTC ≤ gross enrollment premiums'
    if not i.qhp_ptc_eligible:
        method = 'PTC disabled by user; health insurance capped by earned-income capacity'
    return deduction, subsidy, method


def evaluate(i: Inputs) -> Result:
    validate(i)
    r = Result()
    V = r.values
    warn = r.warnings
    profit = i.gross_business-i.business_expenses
    se_tax,half_se,net_se = se_tax_and_half(profit, i.w2_social_security_wages)
    earned = max(0.,profit-half_se)
    max_solo, max_employee, max_employer = solo_limits(i,earned)
    V.update(net_business=profit, net_se_earnings=net_se, se_tax=se_tax,
             half_se_deduction=half_se, solo_earned_income=earned,
             solo_max=max_solo, solo_employee_limit=max_employee,
             solo_employer_limit=max_employer)
    if i.solo_actual > max_solo+.02:
        raise ValueError(f'Solo 401(k) actual {money(i.solo_actual)} exceeds modeled maximum {money(max_solo)}. '
                         'Adjust contribution or review plan/other employer plan rules.')
    if i.w2_prior_401k_deferrals>SOLO_EMPLOYEE_LIMIT+(SOLO_CATCHUP if i.age>=50 else 0):
        warn.append('Other W-2 plan deferrals exceed the modeled personal deferral limit; possible excess or special-plan rules. Review with tax preparer.')
    # Decompose total Solo input into employee first, then employer; the combined earned-income cap still applies.
    solo_employee = min(i.solo_actual,max_employee)
    solo_employer = max(0.,i.solo_actual-solo_employee)
    if solo_employer > max_employer+.01:
        raise ValueError('Solo employer portion exceeds 20% of adjusted net self-employment earnings.')
    V.update(solo_employee_actual=solo_employee,solo_employer_actual=solo_employer)
    hsa_max,hsamonths,notes = hsa_allowance(i)
    r.hsa_schedule = hsamonths
    r.notes.extend(notes)
    V['hsa_max'] = hsa_max
    V['hsa_monthly_full'] = (HSA_FAMILY if i.hsa_family else HSA_SELF)/12.+(HSA_CATCHUP/12. if i.age>=55 else 0)
    if i.hsa_actual + i.employer_hsa > hsa_max + .01:
        raise ValueError(f'Your HSA + employer contributions {money(i.hsa_actual+i.employer_hsa)} '
                         f'exceed modeled HSA limit {money(hsa_max)}.')
    # Pretax HSA funding from employer/payroll already excluded from W2 box 1; DO NOT double deduct.
    interest = i.interest_other_ordinary
    ordinary = (i.w2_box1 + profit + i.short_gains + interest)
    pref = i.long_gains+i.qualified_dividends
    if pref < -1e-7:
        raise ValueError('Combined qualified dividends and long-term gain is negative; capital-loss netting is beyond this planner.')
    base_agi = ordinary + pref - half_se-i.solo_actual-i.hsa_actual-i.other_agi_adjustments
    hi_earned_cap = max(0.,earned-i.solo_actual) if i.self_emp_coverage_eligible else 0.
    insurance_deduction,ptc,method = solve_health_deduction(i,base_agi,hi_earned_cap)
    agi = base_agi-insurance_deduction
    magi = agi+i.aca_magi_addbacks
    fpl_pct=100*magi/FPL
    enrollment_months = i.aca_last_month-i.aca_first_month+1
    gross_premium = enrollment_months*i.enrollment_monthly
    aptc = enrollment_months*i.aptc_monthly
    V.update(insurance_deduction=insurance_deduction, insurance_deduction_cap=hi_earned_cap,
             gross_premium=gross_premium,ptc=ptc,aptc=aptc,
             ptc_reconciliation=ptc-aptc, agi=agi, aca_magi=magi,fpl_pct=fpl_pct,
             insurance_net_after_ptc=gross_premium-ptc,
             expected_monthly_ptc=ptc/enrollment_months,
             expected_monthly_net_premium=i.enrollment_monthly-ptc/enrollment_months,
             deductible_hsa=i.hsa_actual)
    if not i.self_emp_coverage_eligible:
        warn.append('Self-employed insurance deduction disabled: employer-sponsored plan eligibility or other exclusion may apply.')
    if 0 < fpl_pct < 200:
        warn.append('ACA MAGI is below 200% FPL. NY Medicaid/Essential Plan eligibility may prevent QHP PTC; the QHP eligibility checkbox must be verified month by month.')
    if 200<=fpl_pct<250:
        warn.append('NY Essential Plan cutoff changed July 1, 2026. Verify months eligible for qualified Marketplace coverage independently.')
    if abs(magi-4*FPL)<500:
        warn.append('Near the 400% FPL PTC cliff: minor changes to deductions can change credit sharply; verify with Form 8962 and Pub.974.')
    if i.aptc_monthly>0:
        warn.append('APTC reconciliation is included. From tax year 2026, any excess APTC must be fully repaid; 2025 repayment caps do NOT apply.')
    if i.w2_box1>0 and i.w2_social_security_wages==0:
        warn.append('W-2 income entered without Social Security wages. If W-2 box 3 is positive, enter it to avoid overstating SE Social Security tax.')
    if i.solo_actual:
        r.notes.append('Solo total is allocated to elective deferral first, then employer profit sharing. W-2 box 1 already excludes any W-2 pretax deferrals.')
    r.notes.append(method)
    # QBI reduces TAXABLE income, NOT ACA MAGI or self-employment tax.
    qbi = max(0.,profit-half_se-i.solo_actual-insurance_deduction)
    before_qbi=max(0.,agi-i.federal_standard_override)
    qbi_limit=max(0.,before_qbi-max(0.,pref))*.20
    qbi_ded=min(qbi*.20,qbi_limit) if i.qbi_enabled else 0.
    if i.qbi_enabled and qbi >=1000 and qbi_ded<400 and before_qbi>=400:
        qbi_ded=400. # 2026 active-business minimum deduction, subject to eligibility.
    if before_qbi>201750 and i.qbi_enabled:
        warn.append('Federal taxable income above 2026 simplified QBI threshold: wage/property/SSTB phase-ins not modeled; QBI result may be inaccurate.')
    taxable=max(0.,before_qbi-qbi_ded)
    preferred_taxable=min(max(0.,pref),taxable)
    ordinary_taxable=max(0.,taxable-preferred_taxable)
    federal_inc=progressive(ordinary_taxable,FED_BRACKETS)+ltcg_tax(ordinary_taxable,preferred_taxable)
    # Total Additional Medicare tax includes W-2 Medicare wages (box 5) + SE Medicare base.
    w2_medicare=i.w2_medicare_wages if i.w2_medicare_wages else i.w2_box1
    additional_medicare=.009*max(0.,w2_medicare+net_se-200000.)
    if w2_medicare+net_se > 200000 and i.w2_box1 and not i.w2_medicare_wages:
        warn.append('Additional Medicare Tax: W-2 box 5 omitted; box 1 used as a proxy. Enter box 5 and any Additional Medicare already withheld from Form W-2.')
    niit=max(0.,min(max(0.,i.long_gains+i.qualified_dividends+i.short_gains+i.interest_other_ordinary),agi-200000.))*.038
    if niit:
        warn.append('Net investment income tax uses gross net investment income; investment-expense adjustments and Form 8960 are not modeled.')
    fed_total = federal_inc+se_tax+additional_medicare+niit
    fed_net = fed_total-ptc+aptc-i.fed_withholding-i.w2_additional_medicare_withheld
    # NY taxable income usually follows federal AGI, not federal QBI deduction.
    ny_taxable=max(0.,agi-i.ny_standard_override)
    ny_tax=progressive(ny_taxable,NY_BRACKETS)
    nyc_gross=progressive(ny_taxable,NYC_BRACKETS)
    # 2026 NYC school credits: 2025 resident IT-201 table is the last published
    # return instruction. 2026 changes should be verified against the finalized form.
    school_fixed=63. if agi<=250000 else 0.
    school_rate=(min(12000.,ny_taxable)*.00171 if ny_taxable<=12000
                 else 21.+(ny_taxable-12000.)*.00228) if agi<=500000 and ny_taxable<=500000 else 0.
    # NYC-202 / NYC-202S UBT worksheet (NYC Administrative Code § 11-508 & § 11-503):
    # 1. Gross business receipts: evaluated for the $95,000 filing threshold.
    # 2. Net business profit from Schedule C (+/- user NYC net income adjustment).
    # 3. Personal services deduction: 20% of net income, capped at $10,000 for an individual proprietor.
    # 4. Specific business exemption: $5,000 full-year exemption.
    # 5. Taxable UBT base: max(0, net business income - personal services deduction - $5,000).
    # 6. 4% city UBT rate.
    # 7. Small business tax credit:
    #    - If gross UBT <= $3,400: credit is 100% of tax, eliminating UBT completely to $0.
    #    - If gross UBT between $3,400 and $5,400: credit = tax * (5400 - tax) / 2000.
    #    - If gross UBT >= $5,400: credit is $0.
    # 8. $95,000 gross business receipts filing threshold:
    #    A return (NYC-202 or NYC-202S) MUST generally be filed with the NYC Department of Finance
    #    if gross business receipts exceed $95,000, even if actual tax liability is $0!
    ubt_gross_receipts = max(0., i.gross_business)
    ubt_net_business = max(0., profit + i.ubt_city_net_income_adjustment)
    ubt_personal_services = min(10000., max(0., ubt_net_business * 0.20))
    ubt_taxable = max(0., ubt_net_business - ubt_personal_services - 5000.)
    ubt_precredit = .04 * ubt_taxable

    if ubt_precredit <= 3400:
        ubt_business_credit = ubt_precredit
    elif ubt_precredit < 5400:
        ubt_business_credit = ubt_precredit * (5400 - ubt_precredit) / 2000.
    else:
        ubt_business_credit = 0.

    ubt_computed = max(0., ubt_precredit - ubt_business_credit)
    ubt_due = ubt_computed if i.ubt_auto_estimate else i.nyc_ubt_annual_manual
    ubt_filing_required = (ubt_gross_receipts > 95000.) or (ubt_precredit > 3400.)
    r.ubt_form = "NYC-202 (Full)" if (i.solo_actual > 0 or ubt_net_business > 100000.) else "NYC-202S (Simplified)"

    if ny_taxable<=42000:
        ubt_personal_credit_rate=1.
    elif ny_taxable<142000:
        ubt_personal_credit_rate=1.-(ny_taxable-42000)/100000*.77
    elif ny_taxable<1000000:
        ubt_personal_credit_rate=.23
    elif ny_taxable<1250000:
        ubt_personal_credit_rate=.23-.08*(ny_taxable-1000000)/250000
    else:
        ubt_personal_credit_rate=.15
    ubt_resident_credit=min(max(0.,nyc_gross-i.nyc_other_tax_credits),ubt_due*ubt_personal_credit_rate)
    nyc_tax=max(0.,nyc_gross-i.nyc_other_tax_credits-ubt_resident_credit)
    # NYC school credits are refundable, not reductions of resident tax rate.
    nyc_net_with_school=nyc_tax-school_fixed-school_rate
    if agi>107650:
        warn.append('NY AGI above $107,650: New York special tax-computation/benefit-recapture provisions are NOT implemented. NYS tax is only a bracket-based approximation.')
    if ny_taxable<65000:
        warn.append('NYS and NYC tax tables, small-income credits, and rounding may differ from the continuous-rate calculation shown.')
    # MCTMT is measured from net earnings from self-employment as defined in IRC § 1402(a).
    # For a Schedule C sole proprietor using the regular Schedule SE method:
    # Net SE earnings ≈ Schedule C profit * 0.9235.
    # W-2 wages and capital gains do NOT count toward this threshold.
    # Solo 401(k) contributions do NOT reduce this net earnings figure either.
    mct_threshold = 50000.0 if i.tax_year == 2025 else 150000.0
    if i.mct_threshold_override > 0:
        mct_threshold = i.mct_threshold_override

    zone1 = net_se * i.mct_zone1_percent / 100.
    zone2 = net_se * i.mct_zone2_percent / 100.

    # The threshold is evaluated separately for Zone 1 (including NYC/Brooklyn) and Zone 2.
    # Cliff threshold: If net SE earnings in a zone cross the threshold, the tax applies to the
    # FULL allocated net earnings in that zone, not merely the amount above the threshold.
    zone1_subject = zone1 > mct_threshold
    zone2_subject = zone2 > mct_threshold
    mct1 = zone1 * .006 if zone1_subject else 0.
    mct2 = zone2 * .0034 if zone2_subject else 0.
    mct = mct1 + mct2

    thresh_label = f"${mct_threshold:,.0f} ({i.tax_year} threshold)"
    if zone1 > 0 or zone2 > 0:
        if not zone1_subject and not zone2_subject:
            r.notes.append(
                f"MCTMT: Net SE earnings (Zone 1: {money(zone1)}, Zone 2: {money(zone2)}) are at or below "
                f"the {thresh_label} per-zone threshold. No MCTMT is due. Note: W-2 wages and capital gains "
                f"do not count toward this threshold, and Solo 401(k) contributions do not reduce it."
            )
        else:
            if zone1_subject:
                r.notes.append(
                    f"MCTMT Zone 1: Net SE earnings {money(zone1)} exceed the {thresh_label} threshold. "
                    f"By NY law, the 0.60% tax applies to the FULL allocated Zone 1 earnings ({money(zone1)}), "
                    f"not just the excess over the threshold, totaling {money(mct1)}."
                )
            if zone2_subject:
                r.notes.append(
                    f"MCTMT Zone 2: Net SE earnings {money(zone2)} exceed the {thresh_label} threshold. "
                    f"By NY law, the 0.34% tax applies to the FULL allocated Zone 2 earnings ({money(zone2)}), "
                    f"not just the excess over the threshold, totaling {money(mct2)}."
                )

    if ubt_filing_required:
        if ubt_gross_receipts > 95000:
            warn.append(
                f"NYC UBT Filing: Annual gross business receipts ({money(ubt_gross_receipts)}) exceed the $95,000 filing threshold. "
                f"You are generally required to file Form {r.ubt_form} with the NYC Department of Finance, "
                f"even though your net UBT tax owed is {money(ubt_due)}."
            )
        else:
            warn.append(
                f"NYC UBT Filing: Calculated gross UBT exceeds $3,400. You are required to file Form {r.ubt_form} "
                f"with the NYC Department of Finance."
            )
        if i.solo_actual > 0:
            r.notes.append(
                "NYC UBT Form: Because traditional Solo 401(k) deductions are claimed, NYC Form NYC-202 instructions "
                "generally require filing the full Form NYC-202 rather than the simplified Form NYC-202S."
            )
    elif profit > 0 and i.ubt_auto_estimate:
        r.notes.append(
            f"NYC UBT: Gross receipts ({money(ubt_gross_receipts)}) are at or below $95,000 and gross tax is <= $3,400. "
            "No UBT return or payment is required."
        )

    if school_fixed or school_rate:
        warn.append('NYC school credits estimated from latest published full-year resident rules; verify 2026 IT-201 instructions when released.')
    V.update(qbi=qbi,qbi_deduction=qbi_ded, fed_taxable=taxable,
             federal_regular_income_tax=federal_inc, additional_medicare=additional_medicare,niit=niit,
             federal_total_before_ptc=fed_total,federal_balance=fed_net,
             ny_taxable=ny_taxable,nys_tax=ny_tax,nyc_gross=nyc_gross,
             nyc_tax=nyc_net_with_school,nyc_school_fixed=school_fixed,nyc_school_rate=school_rate,
             ubt_gross_receipts=ubt_gross_receipts,
             ubt_business_net=ubt_net_business,
             ubt_personal_services_deduction=ubt_personal_services,
             ubt_exemption=5000.0,
             ubt_taxable=ubt_taxable,
             ubt_precredit=ubt_precredit,
             ubt_business_credit=ubt_business_credit,
             ubt_resident_credit=ubt_resident_credit,
             ubt_filing_required=1.0 if ubt_filing_required else 0.0,
             mct_threshold=mct_threshold,mct_zone1=zone1,mct_zone2=zone2,
             mct1_tax=mct1,mct2_tax=mct2,mct_tax=mct,
             nyc_ubt=ubt_due,
             nys_balance=ny_tax-i.nys_withholding,
             nyc_balance=nyc_net_with_school-i.nyc_withholding,
             mct_balance=mct,
             ubt_balance=ubt_due)
    # Deduct APTC over/under on federal return. Withholdings entered as annual estimates.
    # Equal installments are a CASH-FLOW GUIDE, not proof of safe-harbor compliance.
    fed_remaining=max(0.,fed_net-i.federal_estimates_already_paid)
    ny_remaining=max(0.,ny_tax+nyc_net_with_school+mct-i.nys_withholding-i.nyc_withholding-i.ny_estimates_already_paid)
    ubt_remaining=max(0.,ubt_due-i.ubt_estimates_already_paid)
    V.update(federal_remaining=fed_remaining,ny_combined_remaining=ny_remaining,
             ubt_remaining=ubt_remaining,federal_quarterly=fed_remaining/i.installments_remaining,
             nys_quarterly=max(0.,ny_tax-i.nys_withholding)/i.installments_remaining,
             nyc_quarterly=max(0.,nyc_net_with_school-i.nyc_withholding)/i.installments_remaining,
             mct_quarterly=mct/i.installments_remaining, ny_combined_quarterly=ny_remaining/i.installments_remaining,
             ubt_quarterly=ubt_remaining/i.installments_remaining,
             total_quarterly=(fed_remaining+ny_remaining+ubt_remaining)/i.installments_remaining)
    return r


FIELDS = [
    ('Business & W-2', [
        ('Gross Schedule C business receipts', 'gross_business'),
        ('Schedule C deductible expenses (before Solo/HSA/health insurance)', 'business_expenses'),
        ('W-2 box 1 wages (already net of W-2 pretax deferrals)', 'w2_box1'),
        ('W-2 box 3 Social Security wages', 'w2_social_security_wages'),
        ('W-2 box 5 Medicare wages (for Additional Medicare)', 'w2_medicare_wages'),
        ('W-2 Additional Medicare already withheld (Form 8959 credit)', 'w2_additional_medicare_withheld'),
        ('401(k)/403(b) employee deferrals at OTHER employers', 'w2_prior_401k_deferrals'),
        ('Short-term capital gains (net; ordinary rate)', 'short_gains'),
        ('Long-term capital gains (net; preferential rate)', 'long_gains'),
        ('Qualified dividends (separate from other income)', 'qualified_dividends'),
        ('Interest / nonqualified dividends / other ordinary income', 'interest_other_ordinary'),
        ('Other deductible federal AGI adjustments (non-retirement)', 'other_agi_adjustments'),
        ('ACA MAGI addbacks (tax-exempt interest, etc.)', 'aca_magi_addbacks'),
    ]),
    ('Retirement & HSA', [
        ('Age on December 31, 2026', 'age'),
        ('ACTUAL planned traditional Solo 401(k) total contribution', 'solo_actual'),
        ('ACTUAL deductible direct/personal HSA contributions', 'hsa_actual'),
        ('HSA employer + pretax payroll contributions (already excluded from W-2)', 'employer_hsa'),
        ('HSA eligibility START month (1-12)', 'hsa_first_month'),
        ('HSA eligibility END month (1-12)', 'hsa_last_month'),
        ('HSA FAMILY coverage rather than self-only?', 'hsa_family'),
        ('Claim HSA December last-month rule (testing period required)?', 'use_hsa_last_month_rule'),
    ]),
    ('Marketplace / insurance', [
        ('Gross MONTHLY actual Marketplace plan premium (before APTC)', 'enrollment_monthly'),
        ('Gross MONTHLY benchmark SLCSP premium', 'benchmark_monthly'),
        ('MONTHLY advance PTC paid to insurer (APTC)', 'aptc_monthly'),
        ('ACA Marketplace coverage START month (1-12)', 'aca_first_month'),
        ('ACA Marketplace coverage END month (1-12)', 'aca_last_month'),
        ('Qualified for self-employed health insurance deduction?', 'self_emp_coverage_eligible'),
        ('Qualified for Marketplace PTC all covered months?', 'qhp_ptc_eligible'),
    ]),
    ('New York & other adjustments', [
        ('MCT business earnings allocated to Zone 1, %', 'mct_zone1_percent'),
        ('MCT business earnings allocated to Zone 2, %', 'mct_zone2_percent'),
        ('Tax year for thresholds (2026: $150k; 2025: $50k)', 'tax_year'),
        ('MCTMT threshold override (0 = auto: $150k for 2026, $50k for 2025)', 'mct_threshold_override'),
        ('Automatically estimate NYC Unincorporated Business Tax?', 'ubt_auto_estimate'),
        ('NYC-202 business net-income adjustment (+/-), optional', 'ubt_city_net_income_adjustment'),
        ('Annual NYC UBT NET after its business credit (manual if auto OFF)', 'nyc_ubt_annual_manual'),
        ('Other NYC nonrefundable resident income-tax credits', 'nyc_other_tax_credits'),
        ('Enable estimated federal QBI deduction?', 'qbi_enabled'),
        ('Federal standard/itemized deduction (annual)', 'federal_standard_override'),
        ('NYS standard/itemized deduction (annual)', 'ny_standard_override'),
    ]),
    ('Withholdings & already paid', [
        ('Total estimated annual FEDERAL income tax withholding', 'fed_withholding'),
        ('Total estimated annual NYS tax withholding', 'nys_withholding'),
        ('Total estimated annual NYC income tax withholding', 'nyc_withholding'),
        ('Already paid federal 2026 quarterly estimated taxes', 'federal_estimates_already_paid'),
        ('Remaining estimated-tax installment dates (1-4)', 'installments_remaining'),
        ('Already paid NYS/NYC/MCT combined estimates', 'ny_estimates_already_paid'),
        ('Already paid separate NYC UBT estimates', 'ubt_estimates_already_paid'),
    ]),
]

MONTHS = ['January','February','March','April','May','June',
          'July','August','September','October','November','December']

DEFAULT_EXCEL_FILENAME = 'tax_inputs_2026.xlsx'
DEFAULT_TEMPLATE_FILENAME = 'tax_inputs_template.xlsx'

FIELD_META: dict[str, tuple[str, str]] = {
    'gross_business': ('Currency ($)', 'Annual gross receipts from Schedule C business'),
    'business_expenses': ('Currency ($)', 'Schedule C deductible business expenses (before Solo/HSA/health ins)'),
    'w2_box1': ('Currency ($)', 'W-2 Box 1 wages (already net of W-2 pretax retirement deferrals)'),
    'w2_social_security_wages': ('Currency ($)', 'W-2 Box 3 Social Security wages (used to cap SE Social Security tax)'),
    'w2_medicare_wages': ('Currency ($)', 'W-2 Box 5 Medicare wages (for 0.9% Additional Medicare tax base)'),
    'w2_additional_medicare_withheld': ('Currency ($)', 'W-2 Additional Medicare already withheld (Form 8959 credit)'),
    'w2_prior_401k_deferrals': ('Currency ($)', '401(k)/403(b) employee deferrals made at OTHER employers'),
    'short_gains': ('Currency ($)', 'Net short-term capital gains (taxed at ordinary income rates)'),
    'long_gains': ('Currency ($)', 'Net long-term capital gains (taxed at preferential rates)'),
    'qualified_dividends': ('Currency ($)', 'Qualified dividends (taxed at preferential capital gains rates)'),
    'interest_other_ordinary': ('Currency ($)', 'Taxable interest, nonqualified dividends, other ordinary income'),
    'other_agi_adjustments': ('Currency ($)', 'Other deductible federal AGI adjustments (non-retirement above-the-line)'),
    'aca_magi_addbacks': ('Currency ($)', 'ACA MAGI addbacks (tax-exempt interest, untaxed Social Security, foreign income)'),
    'age': ('Integer (years)', 'Age on December 31, 2026 (determines 401(k) and HSA catch-up limits)'),
    'solo_actual': ('Currency ($)', 'ACTUAL planned traditional Solo 401(k) total contribution (capped at allowable max)'),
    'hsa_actual': ('Currency ($)', 'ACTUAL deductible direct personal HSA contributions'),
    'employer_hsa': ('Currency ($)', 'HSA employer + pretax payroll contributions (already excluded from W-2 Box 1)'),
    'hsa_first_month': ('Integer (1-12)', 'HSA eligibility START month (1=Jan, 2=Feb, ..., 12=Dec)'),
    'hsa_last_month': ('Integer (1-12)', 'HSA eligibility END month (1=Jan, 2=Feb, ..., 12=Dec)'),
    'hsa_family': ('Boolean (TRUE/FALSE)', 'HSA FAMILY coverage ($8,750 cap) vs self-only ($4,400 cap)'),
    'use_hsa_last_month_rule': ('Boolean (TRUE/FALSE)', 'Claim HSA December last-month rule? (Requires remaining eligible through Dec 31, 2027)'),
    'enrollment_monthly': ('Currency ($/mo)', 'Gross MONTHLY actual Marketplace plan premium (before APTC subsidy)'),
    'benchmark_monthly': ('Currency ($/mo)', 'Gross MONTHLY benchmark SLCSP (Second Lowest Cost Silver Plan) premium'),
    'aptc_monthly': ('Currency ($/mo)', 'MONTHLY advance PTC paid directly to insurer (APTC)'),
    'aca_first_month': ('Integer (1-12)', 'ACA Marketplace coverage START month (1-12)'),
    'aca_last_month': ('Integer (1-12)', 'ACA Marketplace coverage END month (1-12)'),
    'self_emp_coverage_eligible': ('Boolean (TRUE/FALSE)', 'Qualified for self-employed health insurance deduction (IRC 162(l))?'),
    'qhp_ptc_eligible': ('Boolean (TRUE/FALSE)', 'Qualified for Marketplace Premium Tax Credit during covered months?'),
    'mct_zone1_percent': ('Percent (0-100)', 'MCT business earnings allocated to Zone 1, % (NYC, Nassau, Suffolk, Westchester)'),
    'mct_zone2_percent': ('Percent (0-100)', 'MCT business earnings allocated to Zone 2, % (Dutchess, Orange, Putnam, Rockland)'),
    'tax_year': ('Integer (2025 or 2026)', 'Tax year (controls MCTMT threshold: $150,000 for 2026 vs $50,000 for 2025)'),
    'mct_threshold_override': ('Currency ($)', 'Optional manual override for MCTMT threshold (0 = auto by tax year)'),
    'ubt_auto_estimate': ('Boolean (TRUE/FALSE)', 'Automatically estimate NYC Unincorporated Business Tax?'),
    'ubt_city_net_income_adjustment': ('Currency ($)', 'NYC-202 business net-income adjustment (+/-), optional'),
    'nyc_ubt_annual_manual': ('Currency ($)', 'Annual NYC UBT NET after business credit (manual entry if auto-estimate is FALSE)'),
    'nyc_other_tax_credits': ('Currency ($)', 'Other NYC nonrefundable resident personal income tax credits'),
    'qbi_enabled': ('Boolean (TRUE/FALSE)', 'Enable estimated federal Section 199A Qualified Business Income deduction?'),
    'federal_standard_override': ('Currency ($)', 'Federal standard/itemized deduction override (single default: $16,100)'),
    'ny_standard_override': ('Currency ($)', 'NYS standard/itemized deduction override (single default: $8,000)'),
    'fed_withholding': ('Currency ($)', 'Total estimated annual FEDERAL income tax withholding'),
    'nys_withholding': ('Currency ($)', 'Total estimated annual NYS tax withholding'),
    'nyc_withholding': ('Currency ($)', 'Total estimated annual NYC income tax withholding'),
    'federal_estimates_already_paid': ('Currency ($)', 'Already paid federal 2026 quarterly estimated taxes'),
    'installments_remaining': ('Integer (1-4)', 'Remaining estimated-tax installment dates (1-4; defaults to 1 for Jan 15 2027)'),
    'ny_estimates_already_paid': ('Currency ($)', 'Already paid NYS/NYC/MCT combined estimates (IT-2105)'),
    'ubt_estimates_already_paid': ('Currency ($)', 'Already paid separate NYC UBT estimates (NYC-5UBTI)'),
}


def create_excel_template(path: str, initial_inputs: Inputs | None = None) -> None:
    """Generate a cleanly formatted, annotated Excel input workbook."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to generate Excel files. Install with 'pip install openpyxl'.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tax_Inputs_2026"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Segoe UI", size=14, bold=True, color="163E57")
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="444444")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_sec = Font(name="Segoe UI", size=10, bold=True, color="163E57")
    font_data = Font(name="Segoe UI", size=10)
    font_var = Font(name="Consolas", size=9, color="333333")
    font_note = Font(name="Segoe UI", size=9, italic=True, color="555555")

    fill_header = PatternFill(start_color="163E57", end_color="163E57", fill_type="solid")
    fill_sec = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid")
    fill_val = PatternFill(start_color="FFFEE6", end_color="FFFEE6", fill_type="solid")

    border_thin = Side(border_style="thin", color="D0D5DD")
    border_data = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    ws["A1"] = "2026 NYC Tax / Solo 401(k) / HSA / ACA Planner — Input Workbook"
    ws["A1"].font = font_title
    ws["A2"] = "Edit the yellow 'Value' column (Column D). You can also override any values in the GUI during your session to model results."
    ws["A2"].font = font_subtitle

    headers = ["Category", "Field Description", "Variable Name", "Value (Enter Here)", "Format / Units", "Guidance / Notes"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_data
        cell.alignment = Alignment(horizontal="center" if col_idx in (3, 4, 5) else "left", vertical="center")

    current_row = 5
    inp = initial_inputs or Inputs()

    for sec_name, fields in FIELDS:
        for c in range(1, 7):
            cell = ws.cell(row=current_row, column=c)
            cell.fill = fill_sec
            cell.border = border_data
        sec_cell = ws.cell(row=current_row, column=1, value=sec_name.upper())
        sec_cell.font = font_sec
        current_row += 1

        for label, var_name in fields:
            val = getattr(inp, var_name)
            meta_unit, meta_desc = FIELD_META.get(var_name, ("Value", ""))

            c1 = ws.cell(row=current_row, column=1, value=sec_name)
            c2 = ws.cell(row=current_row, column=2, value=label)
            c3 = ws.cell(row=current_row, column=3, value=var_name)
            c4 = ws.cell(row=current_row, column=4, value=val)
            c5 = ws.cell(row=current_row, column=5, value=meta_unit)
            c6 = ws.cell(row=current_row, column=6, value=meta_desc)

            c1.font = font_data
            c2.font = font_data
            c3.font = font_var
            c4.font = Font(name="Segoe UI", size=10, bold=True)
            c4.fill = fill_val
            c5.font = font_note
            c6.font = font_note

            for c in (c1, c2, c3, c4, c5, c6):
                c.border = border_data
                c.alignment = Alignment(vertical="center")

            c3.alignment = Alignment(horizontal="center", vertical="center")
            c4.alignment = Alignment(horizontal="right", vertical="center")
            c5.alignment = Alignment(horizontal="center", vertical="center")

            if isinstance(val, bool):
                c4.number_format = 'BOOLEAN'
                c4.alignment = Alignment(horizontal="center", vertical="center")
            elif isinstance(val, int):
                c4.number_format = '#,##0'
            elif 'percent' in var_name:
                c4.number_format = '0.00'
            else:
                c4.number_format = '$#,##0.00'

            current_row += 1

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 54
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 65

    ws.freeze_panes = "A5"
    wb.save(path)


def load_inputs_from_excel(path: str) -> tuple[Inputs, list[str]]:
    """Load inputs from an Excel file, supporting both formatted templates and raw tables."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to load Excel files. Install with 'pip install openpyxl'.")
    warnings: list[str] = []
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    var_col = None
    val_col = None
    desc_col = None
    header_row = None

    for row_idx in range(1, 10):
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, 15)]
        row_strs = [str(v).strip().lower() if v is not None else "" for v in row_vals]
        for c_idx, s in enumerate(row_strs, 1):
            if "variable" in s or s == "var_name" or s == "name":
                var_col = c_idx
            if "value" in s or s == "amount" or s == "val":
                val_col = c_idx
            if "description" in s or "label" in s:
                desc_col = c_idx
        if var_col is not None and val_col is not None:
            header_row = row_idx
            break

    if var_col is None or val_col is None:
        var_col = 3
        val_col = 4
        header_row = 4

    start_row = header_row + 1 if header_row else 1
    read_data: dict[str, Any] = {}
    default_inputs = Inputs()
    field_defaults = asdict(default_inputs)

    desc_to_var: dict[str, str] = {}
    for _, fields in FIELDS:
        for lbl, vname in fields:
            desc_to_var[lbl.strip().lower()] = vname

    for r in range(start_row, ws.max_row + 1):
        raw_var = ws.cell(row=r, column=var_col).value
        raw_val = ws.cell(row=r, column=val_col).value
        raw_desc = ws.cell(row=r, column=desc_col).value if desc_col else None

        v_name = None
        if raw_var is not None and str(raw_var).strip() in field_defaults:
            v_name = str(raw_var).strip()
        elif raw_desc is not None and str(raw_desc).strip().lower() in desc_to_var:
            v_name = desc_to_var[str(raw_desc).strip().lower()]

        if not v_name or raw_val is None:
            continue

        target_type = type(field_defaults[v_name])

        try:
            if target_type is bool:
                if isinstance(raw_val, bool):
                    parsed_val = raw_val
                elif isinstance(raw_val, (int, float)):
                    parsed_val = bool(raw_val)
                else:
                    s = str(raw_val).strip().lower()
                    parsed_val = s in ("true", "1", "yes", "y", "t")
            elif target_type is int:
                if isinstance(raw_val, (int, float)):
                    parsed_val = int(raw_val)
                else:
                    s = str(raw_val).replace("$", "").replace(",", "").strip()
                    parsed_val = int(float(s))
            else:  # float
                if isinstance(raw_val, (int, float)):
                    parsed_val = float(raw_val)
                else:
                    s = str(raw_val).replace("$", "").replace(",", "").replace("%", "").strip()
                    parsed_val = float(s)
            read_data[v_name] = parsed_val
        except Exception as e:
            warnings.append(f"Row {r} ({v_name}): could not parse value '{raw_val}': {e}. Using default.")

    final_kwargs = dict(field_defaults)
    final_kwargs.update(read_data)
    return Inputs(**final_kwargs), warnings


def create_excel_backup(path: str) -> str | None:
    """Create a timestamped safety backup copy of the target Excel file before saving."""
    if not os.path.exists(path):
        return None
    abs_path = os.path.abspath(path)
    dirname, filename = os.path.split(abs_path)
    name_root, ext = os.path.splitext(filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"{name_root}.backup_{timestamp}{ext}"
    backup_path = os.path.join(dirname, backup_filename)
    counter = 1
    while os.path.exists(backup_path):
        backup_filename = f"{name_root}.backup_{timestamp}_{counter}{ext}"
        backup_path = os.path.join(dirname, backup_filename)
        counter += 1
    shutil.copy2(abs_path, backup_path)
    return backup_path


def save_inputs_to_excel(inputs: Inputs, path: str, make_backup: bool = True) -> str | None:
    """Save an Inputs dataclass into an Excel file, creating a backup first if the file exists."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to save Excel files. Install with 'pip install openpyxl'.")
    backup_path = None
    if os.path.exists(path):
        if make_backup:
            backup_path = create_excel_backup(path)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        var_col = 3
        val_col = 4
        start_row = 5
        for r in range(1, 10):
            row_vals = [str(ws.cell(row=r, column=c).value).strip().lower() for c in range(1, 10) if ws.cell(row=r, column=c).value is not None]
            for c_idx, s in enumerate(row_vals, 1):
                if "variable" in s or s == "var_name":
                    var_col = c_idx
                if "value" in s or s == "amount":
                    val_col = c_idx
            if "variable" in " ".join(row_vals) and "value" in " ".join(row_vals):
                start_row = r + 1
                break

        inp_dict = asdict(inputs)
        for r in range(start_row, ws.max_row + 1):
            raw_var = ws.cell(row=r, column=var_col).value
            if raw_var is not None and str(raw_var).strip() in inp_dict:
                ws.cell(row=r, column=val_col, value=inp_dict[str(raw_var).strip()])
        wb.save(path)
    else:
        create_excel_template(path, inputs)
    return backup_path


class Planner(tk.Tk):
    def __init__(self, excel_path: str | None = None):
        super().__init__()
        self.title('2026 NYC Tax / Solo 401(k) / HSA / ACA Planner')
        self.geometry('1180x860')
        self.minsize(900, 680)

        # Resolve initial Excel file location
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        candidate = excel_path or os.path.join(os.getcwd(), DEFAULT_EXCEL_FILENAME)
        if not os.path.exists(candidate) and os.path.exists(os.path.join(script_dir, DEFAULT_EXCEL_FILENAME)):
            candidate = os.path.join(script_dir, DEFAULT_EXCEL_FILENAME)
        elif not os.path.exists(candidate) and not excel_path:
            # Check if template file exists to copy from
            tmpl = os.path.join(os.getcwd(), DEFAULT_TEMPLATE_FILENAME)
            if not os.path.exists(tmpl):
                tmpl = os.path.join(script_dir, DEFAULT_TEMPLATE_FILENAME)
            if os.path.exists(tmpl):
                import shutil
                try:
                    shutil.copyfile(tmpl, candidate)
                except Exception:
                    pass
        self.excel_path = os.path.abspath(candidate)

        self.excel_load_warning = ""
        # Create default template if it does not exist yet
        if not os.path.exists(self.excel_path) and OPENPYXL_AVAILABLE:
            try:
                create_excel_template(self.excel_path, Inputs())
            except Exception as e:
                self.excel_load_warning = f"Could not create default Excel file: {e}"

        # Load inputs from Excel file
        self.excel_inputs = Inputs()
        if os.path.exists(self.excel_path) and OPENPYXL_AVAILABLE:
            try:
                loaded, warns = load_inputs_from_excel(self.excel_path)
                self.excel_inputs = loaded
                if warns:
                    self.excel_load_warning = "; ".join(warns)
            except Exception as e:
                self.excel_load_warning = f"Error reading Excel file: {e}"

        # Form fields initialize to values loaded from Excel
        self.default = self.excel_inputs
        self.vars: dict[str, tk.Variable] = {}
        self._style()

        # Top control bar
        bar = ttk.Frame(self, padding=(12, 10, 12, 6))
        bar.pack(fill='x')
        ttk.Label(bar, text='2026 NYC Tax & Benefits Planner', style='Heading.TLabel').pack(side='left')

        # Action buttons (ordered for clear, intuitive workflow)
        ttk.Button(bar, text='Calculate', command=self.calculate, style='Accent.TButton').pack(side='right', padx=3)
        ttk.Button(bar, text='Save to Excel (Overwrite)', command=self.save_excel_overwrite, style='Save.TButton').pack(side='right', padx=3)
        ttk.Button(bar, text='Reload from Excel', command=self.reload_excel).pack(side='right', padx=3)
        ttk.Button(bar, text='Reset GUI (0s)', command=self.reset_gui_zeros).pack(side='right', padx=3)
        ttk.Button(bar, text='Load Excel…', command=self.load_excel).pack(side='right', padx=3)
        ttk.Button(bar, text='Save As…', command=self.save_excel_as).pack(side='right', padx=3)
        ttk.Button(bar, text='Load JSON…', command=self.load).pack(side='right', padx=2)
        ttk.Button(bar, text='Save JSON…', command=self.save).pack(side='right', padx=2)

        # Status & override notification banner
        status_frame = ttk.Frame(self, padding=(14, 2, 12, 4))
        status_frame.pack(fill='x')
        self.status_var = tk.StringVar(
            value=f"Excel Source: {os.path.basename(self.excel_path)}  •  Feel free to manually override any field below to model what-if scenarios for this session."
        )
        self.status_lbl = ttk.Label(status_frame, textvariable=self.status_var, style='Status.TLabel')
        self.status_lbl.pack(side='left')

        ttk.Label(self, text='SINGLE filer · Full-year NYC resident · 1 Schedule C business · planning estimate, not a tax return',
                  foreground='#8e3d15', padding=(14, 0, 12, 6)).pack(fill='x')

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        self.inputs_tab = ttk.Frame(self.tabs)
        self.result_tab = ttk.Frame(self.tabs)
        self.hsa_tab = ttk.Frame(self.tabs)
        self.method_tab = ttk.Frame(self.tabs)
        for widget, label in [(self.inputs_tab, 'Enter income & contributions'),
                             (self.result_tab, 'Calculation / quarterly estimates'),
                             (self.hsa_tab, 'HSA monthly schedule'),
                             (self.method_tab, 'Method, limitations & sources')]:
            self.tabs.add(widget, text=label)
        self._create_inputs()
        self._create_text_tab(self.result_tab, 'results')
        self._create_text_tab(self.hsa_tab, 'hsa')
        self._create_text_tab(self.method_tab, 'method')
        self._render_method()
        self.calculate()

    def _style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Heading.TLabel', font=('Segoe UI', 15, 'bold'), foreground='#163e57')
        style.configure('Section.TLabel', font=('Segoe UI', 11, 'bold'), foreground='#163e57')
        style.configure('Status.TLabel', font=('Segoe UI', 9, 'bold'), foreground='#0d5c75')
        style.configure('StatusWarning.TLabel', font=('Segoe UI', 9, 'bold'), foreground='#b33800')
        style.configure('Accent.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Save.TButton', font=('Segoe UI', 9, 'bold'), foreground='#1b5e20')
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('TEntry', padding=4, font=('Segoe UI', 10))
        style.configure('TButton', padding=5)
        style.configure('TCheckbutton', font=('Segoe UI', 10))

    def _create_inputs(self):
        canvas=tk.Canvas(self.inputs_tab,highlightthickness=0)
        scrollbar=ttk.Scrollbar(self.inputs_tab,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right',fill='y')
        canvas.pack(side='left',fill='both',expand=True)
        body=ttk.Frame(canvas,padding=12)
        wid=canvas.create_window((0,0),window=body,anchor='nw')
        body.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',lambda e:canvas.itemconfigure(wid,width=e.width))
        for sec, fields in FIELDS:
            lf=ttk.LabelFrame(body,text=sec,padding=(14,10))
            lf.pack(fill='x',pady=7)
            lf.columnconfigure(0,weight=1)
            for row,(label,name) in enumerate(fields):
                val=getattr(self.default,name)
                var=tk.BooleanVar(value=val) if isinstance(val,bool) else tk.StringVar(value=str(val))
                self.vars[name]=var
                ttk.Label(lf,text=label).grid(row=row,column=0,sticky='w',pady=4)
                if isinstance(val,bool):
                    widget=ttk.Checkbutton(lf,variable=var)
                else:
                    widget=ttk.Entry(lf,textvariable=var,width=19,justify='right')
                widget.grid(row=row,column=1,sticky='e',padx=10,pady=4)
        def on_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), 'units')
        canvas.bind('<Enter>',lambda e:canvas.bind_all('<MouseWheel>',on_wheel))
        canvas.bind('<Leave>',lambda e:canvas.unbind_all('<MouseWheel>'))

    def _create_text_tab(self,tab,name):
        outer=ttk.Frame(tab,padding=10)
        outer.pack(fill='both',expand=True)
        output=tk.Text(outer,wrap='word',font=('Consolas',10),padx=12,pady=12,
                       bg='#fbfcff',fg='#1c2430',undo=False)
        output.pack(side='left',fill='both',expand=True)
        scrollbar=ttk.Scrollbar(outer,orient='vertical',command=output.yview)
        scrollbar.pack(side='right',fill='y')
        output.config(yscrollcommand=scrollbar.set,state='disabled')
        setattr(self,name+'_text',output)
        if name=='results':
            btn=ttk.Button(tab,text='Copy result text',command=lambda:self.clipboard_append(self._result_content))
            btn.pack(pady=(0,6))

    def _put(self,tag,text):
        t=getattr(self,tag+'_text')
        t.config(state='normal')
        t.delete('1.0','end')
        t.insert('end',text)
        t.config(state='disabled')

    def _read(self) -> Inputs:
        data: dict[str, Any] = {}
        for key, initial in asdict(Inputs()).items():
            if key not in self.vars:
                continue
            raw = self.vars[key].get()
            if isinstance(initial, bool):
                data[key] = bool(raw)
            elif isinstance(initial, int):
                s = str(raw).replace(',', '').replace('$', '').strip()
                if not s:
                    val = 0
                else:
                    try:
                        val = int(float(s))
                    except ValueError:
                        val = 0
                if key == 'installments_remaining' and val not in (1, 2, 3, 4):
                    val = 1
                elif key in ('hsa_first_month', 'aca_first_month') and not (1 <= val <= 12):
                    val = 1
                elif key in ('hsa_last_month', 'aca_last_month') and not (1 <= val <= 12):
                    val = 12
                elif key == 'age' and not (18 <= val <= 110):
                    val = 37
                data[key] = val
            else:
                s = str(raw).replace(',', '').replace('$', '').replace('%', '').strip()
                if not s:
                    val = 0.0
                else:
                    try:
                        val = float(s)
                    except ValueError:
                        val = 0.0
                data[key] = val
        return Inputs(**data)

    def reload_excel(self):
        """Reload inputs from the active Excel workbook, discarding temporary GUI changes."""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror('Excel Support', 'openpyxl is not installed. Run: pip install openpyxl')
            return
        if not os.path.exists(self.excel_path):
            messagebox.showerror('File Not Found', f'Could not find Excel file:\n{self.excel_path}')
            return
        try:
            loaded, warns = load_inputs_from_excel(self.excel_path)
            self.excel_inputs = loaded
            for name, val in asdict(loaded).items():
                if name in self.vars:
                    self.vars[name].set(val)
            self.status_var.set(
                f"✅ Reloaded from: {os.path.basename(self.excel_path)}  •  All temporary GUI overrides discarded."
            )
            self.status_lbl.configure(style='Status.TLabel')
            self.calculate()
            if warns:
                messagebox.showwarning('Excel Warnings', '\n'.join(warns))
            else:
                messagebox.showinfo('Reloaded', f'Successfully reloaded inputs from:\n{self.excel_path}\n\nAll temporary GUI changes discarded.')
        except Exception as e:
            messagebox.showerror('Excel Load Error', str(e))

    def reset_gui_zeros(self):
        """Reset all GUI fields to 0 / empty for hypothetical scenario testing. Excel file is NOT modified."""
        confirm = messagebox.askyesno(
            'Reset GUI to Zeros / Empty',
            f'Reset all input values in the GUI to 0 / empty for hypothetical what-if testing?\n\n'
            f'Note: Your active Excel file ({os.path.basename(self.excel_path)}) will NOT be touched on disk.\n'
            f'You can click "Reload from Excel" at any time to restore your saved values.'
        )
        if not confirm:
            return

        for key, initial in asdict(Inputs()).items():
            if key not in self.vars:
                continue
            if isinstance(initial, bool):
                self.vars[key].set(False)
            elif key == 'age':
                self.vars[key].set('37')
            elif key == 'installments_remaining':
                self.vars[key].set('1')
            elif key in ('hsa_first_month', 'aca_first_month'):
                self.vars[key].set('1')
            elif key in ('hsa_last_month', 'aca_last_month'):
                self.vars[key].set('12')
            elif isinstance(initial, int):
                self.vars[key].set('0')
            else:
                self.vars[key].set('0.00')

        self.status_var.set(
            f"⚠️ ZEROED BASELINE: GUI inputs set to 0/empty (hypothetical)  •  Excel '{os.path.basename(self.excel_path)}' remains untouched on disk."
        )
        self.status_lbl.configure(style='StatusWarning.TLabel')
        self.calculate()
        messagebox.showinfo(
            'GUI Reset to Zeros',
            'All input fields on the GUI have been reset to 0/empty for hypothetical scenario modeling.\n\n'
            f'Your active Excel file ({os.path.basename(self.excel_path)}) was NOT modified.\n\n'
            '• To discard these hypothetical values and reload from Excel: click "Reload from Excel".\n'
            '• To overwrite the Excel file with these hypothetical values: click "Save to Excel (Overwrite)".'
        )

    def load_excel(self):
        """Browse and load a different Excel workbook."""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror('Excel Support', 'openpyxl is not installed. Run: pip install openpyxl')
            return
        path = filedialog.askopenfilename(
            title='Select Excel Input File',
            filetypes=[('Excel Workbooks', '*.xlsx;*.xlsm'), ('All Files', '*.*')]
        )
        if not path:
            return
        try:
            loaded, warns = load_inputs_from_excel(path)
            self.excel_path = os.path.abspath(path)
            self.excel_inputs = loaded
            for name, val in asdict(loaded).items():
                if name in self.vars:
                    self.vars[name].set(val)
            self.status_var.set(
                f"✅ Excel Source: {os.path.basename(self.excel_path)}  •  Override any field below in GUI for this session."
            )
            self.status_lbl.configure(style='Status.TLabel')
            self.calculate()
            if warns:
                messagebox.showwarning('Excel Warnings', '\n'.join(warns))
        except Exception as e:
            messagebox.showerror('Excel Load Error', str(e))

    def save_excel_overwrite(self):
        """Overwrite the active Excel workbook with current GUI values, creating a safety backup first."""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror('Excel Support', 'openpyxl is not installed. Run: pip install openpyxl')
            return
        try:
            current_inputs = self._read()
        except (ValueError, TypeError) as e:
            messagebox.showerror('Invalid Values', str(e))
            return

        target_file = os.path.basename(self.excel_path)
        confirm = messagebox.askyesno(
            'Confirm Overwrite & Backup',
            f'Overwrite active Excel file with current GUI values?\n\n'
            f'Target file: {target_file}\n'
            f'Location: {self.excel_path}\n\n'
            f'A timestamped safety backup will be created automatically before saving.'
        )
        if not confirm:
            return

        try:
            backup_file = save_inputs_to_excel(current_inputs, self.excel_path, make_backup=True)
            self.excel_inputs = current_inputs
            b_info = f"Backup: {os.path.basename(backup_file)}" if backup_file else "No prior file to back up."
            self.status_var.set(
                f"💾 Saved to: {target_file}  •  {b_info}  •  Excel and GUI in sync."
            )
            self.status_lbl.configure(style='Status.TLabel')
            self.calculate()
            msg = f"Successfully overwritten:\n{self.excel_path}"
            if backup_file:
                msg += f"\n\nA safety backup was created at:\n{backup_file}"
            messagebox.showinfo('Excel Saved & Backed Up', msg)
        except Exception as e:
            messagebox.showerror('Excel Save Error', f"Could not save to Excel:\n{e}")

    def save_excel_as(self):
        """Save current GUI inputs into a new Excel workbook via file picker."""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror('Excel Support', 'openpyxl is not installed. Run: pip install openpyxl')
            return
        try:
            current_inputs = self._read()
        except (ValueError, TypeError) as e:
            messagebox.showerror('Invalid Values', str(e))
            return
        path = filedialog.asksaveasfilename(
            title='Save Inputs to New Excel Workbook',
            defaultextension='.xlsx',
            filetypes=[('Excel Workbooks', '*.xlsx')],
            initialfile=os.path.basename(self.excel_path)
        )
        if not path:
            return
        try:
            backup_file = save_inputs_to_excel(current_inputs, path, make_backup=True)
            self.excel_path = os.path.abspath(path)
            self.excel_inputs = current_inputs
            b_info = f"Backup: {os.path.basename(backup_file)}" if backup_file else "New file created."
            self.status_var.set(
                f"💾 Saved to: {os.path.basename(self.excel_path)}  •  {b_info}  •  Active workbook switched."
            )
            self.status_lbl.configure(style='Status.TLabel')
            self.calculate()
            msg = f"Inputs successfully saved to:\n{path}"
            if backup_file:
                msg += f"\n\nA safety backup was created at:\n{backup_file}"
            messagebox.showinfo('Saved', msg)
        except Exception as e:
            messagebox.showerror('Excel Save Error', str(e))

    def calculate(self):
        try:
            i=self._read()
            r=evaluate(i)
        except (ValueError,TypeError,OverflowError) as e:
            messagebox.showerror('Check your inputs',str(e))
            return
        V=r.values

        # Detect any GUI overrides relative to the loaded Excel inputs
        gui_dict = asdict(i)
        excel_dict = asdict(self.excel_inputs)
        overrides = []
        for k, v in gui_dict.items():
            ev = excel_dict.get(k)
            if ev is not None:
                if isinstance(v, float) and isinstance(ev, float):
                    if abs(v - ev) > 1e-4:
                        overrides.append((k, ev, v))
                elif v != ev:
                    overrides.append((k, ev, v))

        if overrides:
            self.status_var.set(
                f"⚠️ HYPOTHETICAL MODE: {len(overrides)} field(s) modified in GUI  •  Excel '{os.path.basename(self.excel_path)}' is unchanged on disk  •  Click 'Save to Excel (Overwrite)' or 'Reload from Excel'."
            )
            self.status_lbl.configure(style='StatusWarning.TLabel')
        else:
            self.status_var.set(
                f"✅ Excel Source: {os.path.basename(self.excel_path)}  •  All GUI fields match Excel file  •  Edit fields in GUI for hypothetical modeling."
            )
            self.status_lbl.configure(style='Status.TLabel')

        rows=[]
        def title(s):rows.extend(['','='*82,s.upper(),'='*82])
        def line(lbl,key):rows.append(f'  {lbl:<58} {money(V[key]):>17}')
        def metric(lbl,n,fmt=lambda v:money(v)):rows.append(f'  {lbl:<58} {fmt(n):>17}')
        rows.append('2026 • NYC FULL-YEAR RESIDENT • SINGLE FILER • ONE SOLE PROPRIETORSHIP')
        rows.append('ESTIMATE, not a filed-return calculation. Read important warnings below.')

        # Input source & session override summary
        rows.append('')
        rows.append(f'INPUT SOURCE FILE: {self.excel_path}')
        if overrides:
            rows.append(f'⚠️ HYPOTHETICAL SESSION OVERRIDES ACTIVE ({len(overrides)} field(s) modified in GUI for this calculation):')
            rows.append('   (The active Excel file is NOT modified yet. Click "Save to Excel (Overwrite)" to persist these values.)')
            for var_name, old_v, new_v in overrides:
                if isinstance(old_v, bool):
                    old_str = "TRUE" if old_v else "FALSE"
                    new_str = "TRUE" if new_v else "FALSE"
                elif isinstance(old_v, int):
                    old_str = str(old_v)
                    new_str = str(new_v)
                elif 'percent' in var_name:
                    old_str = f"{old_v:.2f}%"
                    new_str = f"{new_v:.2f}%"
                else:
                    old_str = money(old_v)
                    new_str = money(new_v)
                rows.append(f'   • {var_name:<30} Excel: {old_str:>14}  -->  Hypothetical GUI: {new_str:>14}')
        else:
            rows.append('STATUS: All modeled inputs match the active Excel file (no session overrides).')

        title('Solo 401(k): 2026 maximum and entered contribution')
        line('Net Schedule C profit', 'net_business')
        line('SE earnings for SE tax & MCTMT (92.35% of profit)', 'net_se_earnings')
        line('Half of Social Security/Medicare SE tax deductible', 'half_se_deduction')
        line('Solo earned-income compensation', 'solo_earned_income')
        line('MAX total traditional Solo 401(k)', 'solo_max')
        line('  Employee elective deferral allowed', 'solo_employee_limit')
        line('  Employer profit-sharing additional allowed', 'solo_employer_limit')
        line('ACTUAL Solo 401(k) deduction chosen', 'solo_actual') if 'solo_actual' in V else metric('ACTUAL Solo 401(k) deduction chosen',i.solo_actual)
        line('  Employee component of entered contribution','solo_employee_actual')
        line('  Employer component of entered contribution','solo_employer_actual')
        title('HSA and Marketplace premium tax credit')
        line('Max deductible HSA (before employer contributions)', 'hsa_max')
        metric('Monthly HSA limit in a FULLY eligible month',V['hsa_monthly_full'])
        metric('Remaining potential HSA capacity',max(0,V['hsa_max']-i.hsa_actual-i.employer_hsa))
        line('Direct HSA deduction used in calculation','deductible_hsa')
        line('Full gross premium for Marketplace coverage months','gross_premium')
        metric('Gross SLCSP benchmark premium PER MONTH',i.benchmark_monthly)
        line('Self-employed insurance deduction (jointly solved)','insurance_deduction')
        line('AGI after retirement, HSA and insurance deductions','agi')
        line('ACA MAGI (with required addbacks)','aca_magi')
        metric('Percent of federal poverty line',V['fpl_pct'],lambda x:f'{x:,.2f}%')
        line('Annual PREMIUM TAX CREDIT (PTC)','ptc')
        line('Advance PTC already paid to insurer (APTC)','aptc')
        line('Additional credit / (repayment) at federal filing','ptc_reconciliation')
        line('Annual effective health premiums after final PTC','insurance_net_after_ptc')
        line('Average PTC per covered month','expected_monthly_ptc')
        line('Average net premium per covered month','expected_monthly_net_premium')
        title('Annual estimated federal taxes')
        line('QBI / Section 199A deduction (NOT an ACA MAGI deduction)','qbi_deduction')
        line('Federal taxable income after std and QBI','fed_taxable')
        line('Federal income tax (ordinary + preferential gains)','federal_regular_income_tax')
        line('Social Security + Medicare self-employment tax','se_tax')
        line('Additional Medicare tax (uses W-2 box 5)','additional_medicare')
        metric('Less W-2 Additional Medicare already withheld',i.w2_additional_medicare_withheld)
        line('Net Investment Income Tax (approximation)','niit')
        line('Total federal taxes BEFORE refundable PTC','federal_total_before_ptc')
        metric('Less annual federal withholding',i.fed_withholding)
        line('Federal balance after PTC/APTC and withholding','federal_balance')
        metric('Federal estimated payments already made',i.federal_estimates_already_paid)
        line('Further projected federal balance, floor at 0','federal_remaining')
        title('Annual New York State, NYC, and MCTMT')
        line('NYS/NYC taxable income','ny_taxable')
        line('NYS resident income tax (bracket estimate)','nys_tax')
        line('NYC resident income tax before credits','nyc_gross')
        line('NYC fixed refundable school tax credit','nyc_school_fixed')
        line('NYC school rate-reduction refundable credit','nyc_school_rate')
        line('NYC resident UBT income-tax credit','ubt_resident_credit')
        line('NYC resident tax after those credits (may be negative)','nyc_tax')
        title('NYC Unincorporated Business Tax (UBT — NYC Department of Finance)')
        line('Schedule C gross business receipts (UBT filing threshold test)', 'ubt_gross_receipts')
        line('Schedule C net business income (+/- NYC adjustments)', 'ubt_business_net')
        line('Personal services deduction (up to $10k proprietor cap)', 'ubt_personal_services_deduction')
        metric('Business specific exemption (annual full year)', 5000.0)
        line('Taxable NYC UBT base', 'ubt_taxable')
        line('Gross NYC UBT (4.00% rate)', 'ubt_precredit')
        line('NYC small-business tax credit (eliminates tax <= $3,400)', 'ubt_business_credit')
        line('Net NYC UBT owed (before personal income tax credit)', 'nyc_ubt')
        line('NYC resident personal income tax credit (Form IT-219)', 'ubt_resident_credit')
        if V.get('ubt_filing_required', 0) > 0:
            rows.append('')
            rows.append(f'  >>> NYC UBT FILING STATUS: RETURN REQUIRED ({r.ubt_form})')
            if V['ubt_gross_receipts'] > 95000:
                rows.append(f'      Gross receipts ({money(V["ubt_gross_receipts"])}) exceed the $95,000 threshold.')
                rows.append('      You MUST file a UBT return with NYC Dept of Finance even if tax owed is $0.00.')
            if i.solo_actual > 0:
                rows.append('      Note: Solo 401(k) deductions require the full Form NYC-202 (not simplified NYC-202S).')
        else:
            rows.append('')
            rows.append('  >>> NYC UBT FILING STATUS: NO RETURN REQUIRED')
            rows.append('      Gross receipts <= $95,000 and gross tax <= $3,400.')
        metric('NY combined withholding (NYS + NYC)',i.nys_withholding+i.nyc_withholding)
        metric('NY combined estimates already paid',i.ny_estimates_already_paid)
        title(f'Suggested payments across {i.installments_remaining} REMAINING installments (NOT safe-harbor)')
        line('FEDERAL / quarter','federal_quarterly')
        line('NYS tax component / quarter (before prior NY estimates)','nys_quarterly')
        line('NYC income tax component / quarter','nyc_quarterly')
        line('MCTMT component / quarter','mct_quarterly')
        line('NYS+NYC+MCTMT combined payment / quarter','ny_combined_quarterly')
        line('NYC separate UBT payment / quarter','ubt_quarterly')
        line('TOTAL federal + NY + UBT per quarter','total_quarterly')
        rows.extend(['','  2026 due dates: April 15, June 15, September 15, 2026; January 15, 2027.',
                     f'  Shown amount divides remaining estimated cash liability by {i.installments_remaining};',
                     '  due dates already missed may generate penalties; these are not safe-harbor amounts.',
                     '  NYS, NYC resident tax and MCTMT normally share Form IT-2105 estimates.',
                     '  NYC UBT has its own filing and payment requirements.'])
        title('Important limitations and input-dependent warnings')
        rows.extend('  • '+w for w in r.warnings)
        title('Method notes')
        rows.extend('  • '+n for n in r.notes)
        self._result_content='\n'.join(rows)+'\n'
        self._put('results',self._result_content)
        hsa=['2026 HSA MAXIMUM BY MONTH (BASED ON FIRST-DAY ELIGIBILITY)','-'*60,
             f'Family coverage: {i.hsa_family}   Age: {i.age}',
             f'Prorated eligible-year max: {money(V["hsa_max"])}',
             f'Actual personal contribution: {money(i.hsa_actual)}',
             f'Employer/payroll HSA: {money(i.employer_hsa)}',
             f'Additional deductible personal capacity: {money(V["hsa_max"]-i.hsa_actual-i.employer_hsa)}','']
        running=0.
        for month,label,allowed in r.hsa_schedule:
            running+=allowed
            hsa.append(f'  {MONTHS[month-1]:<12}  eligibility: {"YES" if allowed else "NO ":<3}  '
                       f'normal limit: {money(allowed):>12}  cumulative: {money(running):>12}')
        if i.use_hsa_last_month_rule:
            hsa.extend(['','  Last-month rule is a FULL-YEAR special election; displayed monthly',
                        '  normal prorated schedule remains informative, but max may be greater.',
                        '  You must remain HSA-eligible through Dec 31, 2027.'])
        hsa.extend(['','  IMPORTANT: the limit is an annual combined contribution cap.',
                    '  The monthly amounts are a prorating formula, NOT monthly IRS deposit caps.',
                    '  Contributions can be made in a lump sum until the tax filing deadline.'])
        self._put('hsa','\n'.join(hsa))
        self.tabs.select(self.result_tab)

    def _render_method(self):
        s='''2026 NYC TAX & BENEFITS PLANNER — METHOD AND IMPORTANT LIMITATIONS
======================================================================

This is a single-filer planning model, not full Form 1040 / IT-201 preparation.
Model assumptions: 2026 calendar year, one self-employed Schedule C business,
full-year NYS and NYC residency, all 12 months employer-plan/PTC eligibility
specified by the user, no dependents, no other trade/business, standard deductions
unless overridden, no pension plan other than the entered W-2 and Solo accounts.

MATHEMATICAL ORDER
  P = Schedule C gross receipts - business expenses.
  SEbase = max(0, 0.9235 P).
  SETax = 12.4% * min(SEbase, max(0, 184500 - W2 SS wages))
          + 2.9% * SEbase.
  HalfSETax = 1/2 SETax.
  Earned = max(0, P - HalfSETax).
  Solo employee: limited by earnings and 2026 individual deferral limit MINUS
       employee elective deferrals to unrelated employers; age catch-up as applicable.
  Solo employer: at most 20% of adjusted net earnings (P - HalfSETax).
  Combined Solo limit: earnings and 2026 415(c) $72,000 (catch-up extra).
  HSA: 2026 per-month allowed amount times first-of-month eligible months, optionally
       last-month rule with 2027 eligibility testing period.
  AGI-before-health = wages + P + gains + interest - HalfSETax
       - deductible Solo - personal HSA - other explicitly entered adjustments.
  Health deduction <= max(0, Earned - Solo) AND <= premiums not covered by PTC.
  ACA MAGI = AGI - health insurance deduction + ACA MAGI addbacks.
  PTC(D) = sum for eligible months min(actual premium, max(0,
       SLCSP - applicable_pct(MAGI(D))*MAGI(D)/12)).
  Solve maximum D with D + PTC(D) <= actual annual Marketplace premiums.
  Model partitions at 100%/133%/400% FPL and uses bisection on each smooth branch;
       it does NOT blindly use fixed-point iteration across discontinuities.
  2026 PTC income cap is 400% FPL, based on 2025 one-person FPL $15,650.
  Federal QBI: lower-income simplified Section 199A deduction; does not reduce MAGI.
  Federal regular tax: ordinary rate brackets + preferential net LT gains/qualified
       dividends stacking; plus SE taxes, NIIT (simplified), Add'l Medicare (approx).
  NYS / NYC taxable income based on AGI, not QBI; NY 2026 reduced brackets.
  MCTMT (Metropolitan Commuter Transportation Mobility Tax - Article 23):
       Threshold is evaluated based strictly on net earnings from self-employment
       attributable to each relevant MCTMT zone, NOT federal AGI.
       Sole proprietor regular method: Net SE earnings ≈ Schedule C profit × 0.9235.
       W-2 wages and capital gains do NOT count toward this threshold.
       Solo 401(k) contributions do NOT reduce this earnings figure either.
       Threshold evaluated SEPARATELY for Zone 1 (NYC: Manhattan, Brooklyn, Queens,
       Bronx, Staten Island) and Zone 2 (Nassau, Suffolk, Westchester, Rockland,
       Orange, Putnam, Dutchess).
       Cliff threshold: if earnings in a zone cross the threshold, the tax applies to
       the FULL allocated net earnings in that zone (0.60% Zone 1, 0.34% Zone 2),
       not merely the amount above the threshold.
       Threshold rule: $150,000 threshold applies beginning in tax year 2026.
       The prior $50,000 threshold still applies to 2025 returns (even if filed in 2026).
  Installments: remaining estimated annual liability less withholding and prior
       payment amounts, divided by entered REMAINING payment count; this is not
       proof of safe-harbor compliance or a late-payment penalty calculation.

IMPORTANT UNSUPPORTED OR SPECIAL CASES
 • Not supported: MFJ/HOH/MFS/dependents, multiple businesses, partnerships/S corps,
   Roth Solo contributions, 403(b) special deferral rules, spouse coverage,
   itemized AMT/SALT interactions, tax credits other than PTC, taxes on investment
   income requiring special schedules, foreign income or multi-state allocations,
   business net operating losses, substantial capital losses/carryforwards,
   ACA multi-policy monthly allocations or alternative employer coverage periods.
 • NY/NYC tax-table rounding and New York benefit recapture at higher income can
   differ from the continuous progressive bracket estimate. Other NY/NYC credits
   and filing adjustments are not comprehensively included.
 • Additional Medicare uses W-2 Medicare wages BOX 5 and allows for employer
   withholding on Form 8959; box 1 is only a proxy if box 5 is missing.
 • The extra 2026 HSA contribution allowed via the last-month rule is conditional
   on maintaining eligibility through the following calendar year.
 • Self-employed health deduction depends on exclusion for ANY month eligible
   to participate in subsidized employer or spouse employer health coverage.
 • ACA eligibility depends on OTHER health coverage and Marketplace tax household;
   NY Medicaid/Essential Plan eligibility can block QHP PTC. NY Essential Plan
   threshold changed mid-2026, so a single-year MAGI cannot resolve eligibility.
 • Pub 974 allows specific iterative/simplified calculations; while this program
   enforces the deduction + premium credit constraint with robust numeric search,
   it is NOT a form-complete implementation of ALL Pub 974 Worksheets W/X/Y/Z.
   Certain discontinuity/corner cases require filing software/manual review.
 • 2026 PTC excess APTC repayment is FULL, without prior-year repayment caps.
 • NYC UBT (Unincorporated Business Tax - NYC Admin Code Title 11 Chapter 5):
   - Administered separately by the NYC Department of Finance (Form NYC-202/NYC-202S).
   - Rate: 4.00% on unincorporated business taxable income.
   - Filing threshold: >$95,000 in annual gross business income (based on gross receipts,
     NOT net profit). A return is generally required even if tax liability is $0!
   - Personal services deduction: 20% of net income, capped at $10,000 per proprietor.
   - Specific business exemption: $5,000 full-year exemption.
   - Small business credit: completely eliminates UBT when calculated gross tax is
     $3,400 or less (e.g. $100k profit -> $85k taxable base -> $3,400 tax -> $0 owed).
     Phases out between $3,400 and $5,400, and drops to $0 at $5,400+.
   - Form selection: Solo 401(k) contributions or multiple unincorporated businesses
     generally require filing the full Form NYC-202 rather than simplified NYC-202S.
   - NYC resident personal credit: Form IT-219 provides a partial credit on NYC
     personal income tax for UBT paid (phased from 100% down to 15% by NY taxable income).
 • NYC school credit uses the latest published full-year 2025 IT-201 rules.
 • Earnings may be seasonal; annualized-income estimated tax (Form 2210, IT-2105.9)
   and prior-year safe-harbor rules require extra prior-year and timing inputs.
   Equal quarterly allocations are NOT safe-harbor certification, especially late
   in October after three 2026 installment dates have passed.
 • The self-employment cap may interact with individual health plan enrollment
   dates and compensation limits in specific businesses. In such cases, review
   Form 7206 and Publication 974 carefully.

SOURCES (official where possible):
'''
        for desc,url in SOURCES:
            s+=f'  {desc}\n  {url}\n\n'
        self._put('method',s)

    def save(self):
        try: data=asdict(self._read())
        except (ValueError,TypeError) as e:
            messagebox.showerror('Invalid values',str(e));return
        path=filedialog.asksaveasfilename(defaultextension='.json',filetypes=[('JSON files','*.json')],
                                          initialfile='nyc_tax_inputs_2026.json')
        if path:
            with open(path,'w',encoding='utf-8') as f:json.dump(data,f,indent=2)

    def load(self):
        path=filedialog.askopenfilename(filetypes=[('JSON files','*.json')])
        if not path:return
        try:
            with open(path,encoding='utf-8') as f:data=json.load(f)
            for name,value in data.items():
                if name in self.vars:self.vars[name].set(value)
            self.status_var.set("Inputs loaded from JSON file (custom scenario).")
            self.calculate()
        except (ValueError,TypeError, OSError) as e:
            messagebox.showerror('Could not load',str(e))


def run_tests():
    from unittest import TestCase, main
    class CoreTests(TestCase):
        def test_zero(self):
            r=evaluate(Inputs(gross_business=0))
            self.assertAlmostEqual(r.values['solo_max'],0)
            self.assertAlmostEqual(r.values['mct_tax'],0)
        def test_solo_limits_100k(self):
            v=evaluate(Inputs()).values
            self.assertAlmostEqual(v['se_tax'],100000*.9235*.153, places=4)
            self.assertAlmostEqual(v['solo_max'],24500+.2*(100000-v['half_se_deduction']),places=4)
        def test_se_tax_small_profit(self):
            v=evaluate(Inputs(gross_business=200)).values
            self.assertEqual(v['se_tax'],0)
            self.assertGreater(v['net_se_earnings'],0)
        def test_mct_below_150k(self):
            # 2026 rule: $100k Schedule C profit = $92,350 net SE earnings, below $150k threshold -> $0 tax
            v = evaluate(Inputs(gross_business=100000, tax_year=2026)).values
            self.assertEqual(v['mct_tax'], 0)
            self.assertEqual(v['mct_threshold'], 150000.0)
        def test_mct_2025_threshold_50k(self):
            # 2025 distinction: $100k Schedule C profit = $92,350 net SE earnings.
            # In 2025 (threshold $50,000), it exceeds the threshold! Tax applies to full $92,350 at 0.60% = $554.10.
            v_2025 = evaluate(Inputs(gross_business=100000, tax_year=2025)).values
            self.assertEqual(v_2025['mct_threshold'], 50000.0)
            self.assertAlmostEqual(v_2025['mct_tax'], 92350 * 0.006, places=2)
            # In 2026, it is exempt
            v_2026 = evaluate(Inputs(gross_business=100000, tax_year=2026)).values
            self.assertEqual(v_2026['mct_tax'], 0)
        def test_mct_cliff_rule_full_earnings(self):
            # Cliff rule: Tax applies to FULL allocated net earnings in that zone, not just excess over $150k.
            v = evaluate(Inputs(gross_business=200000, tax_year=2026)).values
            net_se = 200000 * 0.9235  # 184,700
            self.assertAlmostEqual(v['mct_tax'], net_se * 0.006, places=3)
        def test_mct_w2_solo_independence(self):
            # W-2 wages and Solo 401(k) deductions do not alter MCTMT net earnings
            v1 = evaluate(Inputs(gross_business=100000)).values
            v2 = evaluate(Inputs(gross_business=100000, w2_box1=80000, solo_actual=20000)).values
            self.assertEqual(v1['mct_zone1'], v2['mct_zone1'])
            self.assertEqual(v1['mct_tax'], v2['mct_tax'])
        def test_nyc_ubt_credit_and_filing_threshold(self):
            # Test example from NYC UBT guidance: $100,000 net profit, $100,000 gross receipts
            # Profit: $100,000
            # Personal services deduction: -$10,000 (20% of $100k, capped at $10k)
            # Business exemption: -$5,000
            # Taxable UBT income: $85,000
            # Gross UBT (4%): $3,400
            # Small business tax credit: -$3,400
            # Net UBT owed: $0.00
            # Filing required: True (gross receipts $100,000 > $95,000 threshold!)
            r = evaluate(Inputs(gross_business=100000))
            v = r.values
            self.assertAlmostEqual(v['ubt_personal_services_deduction'], 10000.0)
            self.assertAlmostEqual(v['ubt_taxable'], 85000.0)
            self.assertAlmostEqual(v['ubt_precredit'], 3400.0)
            self.assertAlmostEqual(v['ubt_business_credit'], 3400.0)
            self.assertAlmostEqual(v['nyc_ubt'], 0.0)
            self.assertEqual(v['ubt_filing_required'], 1.0)
            self.assertEqual(r.ubt_form, "NYC-202S (Simplified)")

            # Test Solo 401(k) triggers full Form NYC-202
            r_solo = evaluate(Inputs(gross_business=100000, solo_actual=15000))
            self.assertEqual(r_solo.ubt_form, "NYC-202 (Full)")

            # Test below $95k filing threshold: $80,000 gross receipts, $80,000 profit
            # 20% of 80k is 16k, capped at $10k personal services deduction
            # Taxable: 80k - 10k - 5k = 65k
            # Gross UBT: 65k * 0.04 = $2,600 <= $3,400 -> UBT owed = $0
            # Filing required: False ($80k <= $95k and gross tax <= $3,400)
            r2 = evaluate(Inputs(gross_business=80000))
            self.assertAlmostEqual(r2.values['nyc_ubt'], 0.0)
            self.assertEqual(r2.values['ubt_filing_required'], 0.0)

            # Test phaseout above $3,400: $140,000 profit
            # Personal services: $10,000; Exemption: $5,000; Taxable: $125,000
            # Gross UBT: 125,000 * 0.04 = $5,000
            # Credit: 5000 * (5400 - 5000) / 2000 = $1,000
            # Net UBT: $5,000 - $1,000 = $4,000
            r3 = evaluate(Inputs(gross_business=140000))
            self.assertAlmostEqual(r3.values['ubt_precredit'], 5000.0)
            self.assertAlmostEqual(r3.values['ubt_business_credit'], 1000.0)
            self.assertAlmostEqual(r3.values['nyc_ubt'], 4000.0)
            self.assertEqual(r3.values['ubt_filing_required'], 1.0)
        def test_hsa_dec_start(self):
            a=Inputs(hsa_first_month=12)
            self.assertAlmostEqual(evaluate(a).values['hsa_max'],4400/12)
            a.use_hsa_last_month_rule=True
            self.assertAlmostEqual(evaluate(a).values['hsa_max'],4400)
        def test_ptc_cliff(self):
            a=Inputs(gross_business=0)
            self.assertEqual(ptc_for_magi(a,62600.01),0)
            self.assertGreater(ptc_for_magi(a,62600),0)
        def test_solo_hsa_agi_impact(self):
            a=Inputs(gross_business=90000)
            baseline=evaluate(a).values
            a.hsa_actual=3000
            lower=evaluate(a).values
            self.assertAlmostEqual(baseline['aca_magi']-lower['aca_magi'],3000,delta=3000)
        def test_no_double_hsa_deduction(self):
            baseline=evaluate(Inputs()).values['agi']
            same=evaluate(Inputs(employer_hsa=2000)).values['agi']
            self.assertAlmostEqual(baseline,same)
        def test_insurance_sum_feasible(self):
            for income in (0,20000,40000,60000,62000,65000,80000,100000):
                v=evaluate(Inputs(gross_business=income)).values
                self.assertLessEqual(v['ptc']+v['insurance_deduction'],v['gross_premium']+.02)
        def test_other_plan_deferrals(self):
            v=evaluate(Inputs(w2_prior_401k_deferrals=24500)).values
            self.assertEqual(v['solo_employee_limit'],0)
        def test_ptc_133_discontinuity(self):
            self.assertAlmostEqual(ptc_percentage(FPL*1.32999),.021,places=5)
            self.assertAlmostEqual(ptc_percentage(FPL*1.33),.0314,places=6)
        def test_solver_133_discontinuity(self):
            # Confirm search remains feasible when the price schedule jumps at 133% FPL.
            for income in [FPL*1.33+x for x in (-500,-.001,0,.001,500)]:
                q=Inputs(gross_business=0,w2_box1=income,enrollment_monthly=350,benchmark_monthly=370)
                x=evaluate(q).values
                self.assertLessEqual(x['ptc']+x['insurance_deduction'],x['gross_premium']+.02)
        def test_remaining_installments(self):
            a=Inputs(gross_business=100000,installments_remaining=1)
            b=Inputs(gross_business=100000,installments_remaining=4)
            self.assertAlmostEqual(evaluate(a).values['total_quarterly'],4*evaluate(b).values['total_quarterly'])
        def test_capgain_stack(self):
            self.assertAlmostEqual(ltcg_tax(49000,1000),(.15*550),places=3)
        def test_bracket(self):
            self.assertAlmostEqual(progressive(12401,FED_BRACKETS),1240+.12)
        def test_apct_reconciliation(self):
            v=evaluate(Inputs(aptc_monthly=200)).values
            self.assertAlmostEqual(v['ptc_reconciliation'],v['ptc']-2400)
        def test_excel_template_roundtrip(self):
            if not OPENPYXL_AVAILABLE:
                return
            import tempfile
            with tempfile.TemporaryDirectory() as tmpdir:
                p = os.path.join(tmpdir, "test.xlsx")
                create_excel_template(p, Inputs(gross_business=125000, solo_actual=20000, benchmark_monthly=910.0))
                loaded, warns = load_inputs_from_excel(p)
                self.assertEqual(len(warns), 0)
                self.assertAlmostEqual(loaded.gross_business, 125000)
                self.assertAlmostEqual(loaded.solo_actual, 20000)
                self.assertAlmostEqual(loaded.benchmark_monthly, 910.0)
                # Test update and save
                loaded.hsa_first_month = 6
                loaded.hsa_family = True
                save_inputs_to_excel(loaded, p)
                reloaded, _ = load_inputs_from_excel(p)
                self.assertEqual(reloaded.hsa_first_month, 6)
                self.assertTrue(reloaded.hsa_family)
        def test_excel_parsing_formatting(self):
            if not OPENPYXL_AVAILABLE:
                return
            import tempfile
            with tempfile.TemporaryDirectory() as tmpdir:
                p = os.path.join(tmpdir, "format_test.xlsx")
                create_excel_template(p)
                wb = openpyxl.load_workbook(p)
                ws = wb.active
                # Write string-formatted numbers with currency/commas to verify robust parsing
                for r in range(5, ws.max_row + 1):
                    vname = str(ws.cell(row=r, column=3).value).strip()
                    if vname == 'gross_business':
                        ws.cell(row=r, column=4, value="$180,000.50")
                    elif vname == 'hsa_family':
                        ws.cell(row=r, column=4, value="YES")
                    elif vname == 'age':
                        ws.cell(row=r, column=4, value="42")
                wb.save(p)
                loaded, _ = load_inputs_from_excel(p)
                self.assertAlmostEqual(loaded.gross_business, 180000.50)
                self.assertTrue(loaded.hsa_family)
                self.assertEqual(loaded.age, 42)
        def test_excel_backup_and_reset_zeros(self):
            if not OPENPYXL_AVAILABLE:
                return
            import tempfile
            with tempfile.TemporaryDirectory() as tmpdir:
                p = os.path.join(tmpdir, "my_taxes.xlsx")
                create_excel_template(p, Inputs(gross_business=100000, solo_actual=15000))
                
                # First overwrite save creates a backup of the original file
                inp_modified = Inputs(gross_business=160000, solo_actual=25000)
                backup_path = save_inputs_to_excel(inp_modified, p, make_backup=True)
                self.assertIsNotNone(backup_path)
                self.assertTrue(os.path.exists(backup_path))
                self.assertTrue(".backup_" in backup_path)
                
                # Check backup has original values (100,000)
                bk_loaded, _ = load_inputs_from_excel(backup_path)
                self.assertAlmostEqual(bk_loaded.gross_business, 100000)
                self.assertAlmostEqual(bk_loaded.solo_actual, 15000)
                
                # Check current target file has new values (160,000)
                cur_loaded, _ = load_inputs_from_excel(p)
                self.assertAlmostEqual(cur_loaded.gross_business, 160000)
                self.assertAlmostEqual(cur_loaded.solo_actual, 25000)

                # Test evaluation with zeroed baseline (Reset GUI to 0s)
                zero_inputs = Inputs(
                    gross_business=0.0, business_expenses=0.0, w2_box1=0.0,
                    short_gains=0.0, long_gains=0.0, solo_actual=0.0, hsa_actual=0.0,
                    benchmark_monthly=0.0, enrollment_monthly=0.0
                )
                r_zero = evaluate(zero_inputs)
                self.assertAlmostEqual(r_zero.values['federal_balance'], 0.0)
                self.assertAlmostEqual(r_zero.values['total_quarterly'], 0.0)
    suite=__import__('unittest').defaultTestLoader.loadTestsFromTestCase(CoreTests)
    run=__import__('unittest').TextTestRunner(verbosity=2).run(suite)
    sys.exit(0 if run.wasSuccessful() else 1)

if __name__=='__main__':
    if len(sys.argv)>1 and sys.argv[1]=='--test':
        run_tests()
    else:
        cli_excel = None
        if len(sys.argv) > 1:
            if sys.argv[1] == '--excel' and len(sys.argv) > 2:
                cli_excel = sys.argv[2]
            elif sys.argv[1].endswith('.xlsx') or sys.argv[1].endswith('.xlsm'):
                cli_excel = sys.argv[1]
        Planner(excel_path=cli_excel).mainloop()
