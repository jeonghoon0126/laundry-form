#!/usr/bin/env python3
"""
월별 세탁물 정산 자동화 스크립트
- Supabase에서 데이터 조회
- 사업자별 거래명세서 PDF 생성
- 홈택스 세금계산서 Excel 생성
- 이메일 발송
"""

import os
import sys
import smtplib
import calendar
import re
from datetime import datetime, date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

import psycopg2
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook

# ============================================================
# 설정
# ============================================================

# Supabase 연결 정보 (strip으로 공백/줄바꿈 제거)
SUPABASE_URI = os.environ.get('SUPABASE_URI',
    "postgresql://postgres.xlbckmqdzzkwtboscjgr:pmjeonghoon4189@aws-1-ap-northeast-2.pooler.supabase.com:5432/postgres?sslmode=require").strip()

# 이메일 설정
EMAIL_FROM = "kham0126@gmail.com"
EMAIL_TO = os.environ.get('EMAIL_TO', 'kham0126@gmail.com').strip()
EMAIL_CC = os.environ.get('EMAIL_CC', '').strip()  # 추가 수신자 (선택)
EMAIL_PASSWORD = os.environ.get('EMAIL_PASSWORD', '').strip()

# 공급자 정보
SUPPLIER = {
    'name': '캐리',
    'reg_no': '521-23-01693',
    'owner': '함정훈',
    'address': '',
    'bank': '카카오뱅크 3333349200339 (예금주: 함정훈(캐리))'
}

# 숙소 → 사업자 매핑
BUSINESS_MAP = {
    '중구 장충단로 225': ('554-88-03481', '주식회사 모어브릿지', '홍석화'),
    '서대문구 연희로4길 25-7': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 고산자로 508-3': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 왕산로 200, 1004호': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 장한로26나길 21': ('767-87-02214', '주식회사 콥스', '남택호'),
    '광진구 능동로 165-1': ('767-87-02214', '주식회사 콥스', '남택호'),
    '강남구 봉은사로37길 8': ('767-87-02214', '주식회사 콥스', '남택호'),
    '송파구 가락로28길 3-10': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 회기로 189': ('419-11-02853', '오를리(Orly)', '김지혜'),
    '관악구 신림동1길 19-5': ('461-86-03598', '주식회사스테이모먼트', '유경민'),
}

# 품목 단가 (기본)
PRICES = {
    'blanket': 3000,       # 이불
    'mat': 2000,           # 매트
    'pillow_cover': 250,   # 베개커버
    'towel': 500,          # 타올
    'body_towel': 800,     # 바디타올
    'pillow_fill': 5000,   # 배게 솜
    'cotton_blanket': 15000, # 솜 이불
}

# 사업자번호별 단가 override (기본값과 다른 항목만 기재)
BUSINESS_PRICES = {
    '419-11-02853': {      # 오를리(Orly) 김지혜
        'blanket': 3000,   # 이불
        'pillow_cover': 250,  # 베개커버
        'mat': 1500,       # 매트 (기본 2,000 → 오를리 1,500)
    },
}

LOCATION_PRICES = {
    '중구 장충단로 225': {
        'body_towel': 1100,  # 메종드브릭 바디타올
    },
}

STAYMOMENT_LOCATION = '관악구 신림동1길 19-5'
STAYMOMENT_SETTLEMENT_END_DATE = date(2026, 5, 1)
JANGHANPYEONG_LOCATION = '동대문구 장한로26나길 21'
JANGHANPYEONG_SETTLEMENT_END_DATE = date(2026, 6, 1)
LAST_DAY_CARRYOVER_START_DATE = date(2026, 4, 30)

ITEM_NAMES = {
    'blanket': '이불',
    'mat': '매트',
    'pillow_cover': '베개커버',
    'towel': '타올',
    'body_towel': '바디타올',
    'pillow_fill': '배게 솜',
    'cotton_blanket': '솜 이불',
}

# ============================================================
# Google Sheets 연동 설정
# ============================================================

CARRY_KOEUN_SPREADSHEET_ID = '1awoBzifmkDdnbQ-t3JAO5t9Ktnd9Qpn-PZgbQj_6apA'  # 캐리_고객
CARRY_JUNGSAN_SPREADSHEET_ID = '1yBDSsG8vgvM-e2oswAqDh2g1K0hzGzANstnesPWPBg4'  # 캐리_정산
CARRY_PROFIT_SHEET_ID = 1239394038  # 캐리_고객 > 영업이익계산
PROFIT_FORMAT_TEMPLATE_ROW = 4
KOPS_REG_NO = '767-87-02214'

SHEETS_FIXED_COSTS = {
    'hourly_wage': 12000,          # D열: 시급
    'labor_hours_baseline': 190,   # 인건비 산정 기준 시간
    'labor_revenue_baseline': 7457500, # 인건비 비율 기준 매출
    'logistics_count': 9,          # F열: 물류 횟수
    'logistics_cost_per': 120000,  # G열: 물류 1회 비용
    'rent_utility': 770000,        # I열: 월세+관리비
    'electricity': 250000,         # J열: 전기세
    'water': 100000,               # K열: 수도세
    'insurance': 60000,            # L열: 보험
    'supplies_reference_cost': 480000, # 2026-05 실제 소모품 지출
    'supplies_rate': 480000 / 7457500, # M열: 2026-05 매출 대비 소모품률
    'withholding_national_rate': 0.03,
    'withholding_local_rate': 0.003,
    'vat_inclusive_divisor': 11,
}

INCOME_TAX_BRACKETS = [
    (14_000_000, 0.06, 0),
    (50_000_000, 0.15, 1_260_000),
    (88_000_000, 0.24, 5_760_000),
    (150_000_000, 0.35, 15_440_000),
    (300_000_000, 0.38, 19_940_000),
    (500_000_000, 0.40, 25_940_000),
    (1_000_000_000, 0.42, 35_940_000),
    (float('inf'), 0.45, 65_940_000),
]
LOCAL_INCOME_TAX_RATE = 0.10

INVOICE_SHEET_MAP = {
    '중구 장충단로 225':            'invoice(거래명세서)_장충동 메종드브릭',
    '동대문구 회기로 189':          'invoice(거래명세서)_Orly',
    '관악구 신림동1길 19-5':        'invoice(거래명세서)_스테이모먼트',
    '서대문구 연희로4길 25-7':      'invoice(거래명세서)_서대문구 연희로4길 25-7, 연남에코리빙',
    '동대문구 고산자로 508-3':      'invoice(거래명세서)_동대문구 고산자로 508-3, 스테이브리즈',
    '동대문구 왕산로 200, 1004호':  'invoice(거래명세서)_동대문구 왕산로 200, 청량리역 롯데캐슬 SKY-L65',
    '강남구 봉은사로37길 8':        'invoice(거래명세서)_강남구 봉은사로37길 8',
    '송파구 가락로28길 3-10':       'invoice(거래명세서)_송파구 가락로28길 3-10 스테이브리즈 송파',
    '광진구 능동로 165-1':          'invoice(거래명세서)_능동로 165-1 화양프라하임',
    '동대문구 장한로26나길 21':     'invoice(거래명세서)_가회',
}

INVOICE_JOB_MODES = {'full', 'invoice_sheets_only', 'list_invoice_sheets'}

INVOICE_SHEET_ALIASES = {
    '중구 장충단로 225': ['장충동 메종드브릭', '메종드브릭', '장충단로 225'],
    '강남구 봉은사로37길 8': ['강남구 봉은사로37길 8', '봉은사로37길 8', '봉은사로 37길 8'],
}


def get_prices(reg_no: str) -> dict:
    """사업자번호에 맞는 단가 반환 (override 적용)"""
    prices = dict(PRICES)
    if reg_no in BUSINESS_PRICES:
        prices.update(BUSINESS_PRICES[reg_no])
    return prices


def get_location_prices(location: str) -> dict:
    """숙소별 예외 단가까지 반영한 단가 반환"""
    reg_no = BUSINESS_MAP[location][0]
    prices = get_prices(reg_no)
    if location in LOCATION_PRICES:
        prices.update(LOCATION_PRICES[location])
    return prices


def is_settlement_location_active(location: str, record_date: date) -> bool:
    """정산 대상 숙소 여부"""
    if location == STAYMOMENT_LOCATION and record_date >= STAYMOMENT_SETTLEMENT_END_DATE:
        return False
    if location == JANGHANPYEONG_LOCATION and record_date >= JANGHANPYEONG_SETTLEMENT_END_DATE:
        return False
    return True


def get_settlement_period(year: int, month: int) -> tuple[date, date]:
    """정산 대상 기간 반환: 시작일 포함, 종료일 제외"""
    month_start = date(year, month, 1)
    month_last_day = calendar.monthrange(year, month)[1]
    month_last_date = date(year, month, month_last_day)
    prev_month_last_date = month_start - timedelta(days=1)

    start_date = (
        prev_month_last_date
        if prev_month_last_date >= LAST_DAY_CARRYOVER_START_DATE
        else month_start
    )
    end_exclusive = (
        month_last_date
        if month_last_date >= LAST_DAY_CARRYOVER_START_DATE
        else month_last_date + timedelta(days=1)
    )
    return start_date, end_exclusive


def get_record_settlement_month(record_date: date) -> tuple[int, int]:
    """세탁 기록 1건이 속하는 정산 월 반환"""
    last_day = calendar.monthrange(record_date.year, record_date.month)[1]
    if record_date >= LAST_DAY_CARRYOVER_START_DATE and record_date.day == last_day:
        if record_date.month == 12:
            return record_date.year + 1, 1
        return record_date.year, record_date.month + 1
    return record_date.year, record_date.month


def format_settlement_period(year: int, month: int) -> str:
    """정산 기간 표시 문자열"""
    start_date, end_exclusive = get_settlement_period(year, month)
    end_date = end_exclusive - timedelta(days=1)
    return f"{start_date.month}/{start_date.day:02d}~{end_date.month}/{end_date.day:02d}"


def calculate_supplies_cost(total_amount: int) -> int:
    """매출 연동 소모품 비용"""
    return round(total_amount * SHEETS_FIXED_COSTS['supplies_rate'])


def calculate_included_vat(amount: int) -> int:
    """부가세 포함 금액에서 부가세 부분 계산"""
    return split_vat_inclusive_amount(amount)[1]


def split_vat_inclusive_amount(amount: int) -> tuple[int, int]:
    """부가세 포함 금액을 공급가액과 부가세로 분리"""
    supply_amount = amount * 10 // SHEETS_FIXED_COSTS['vat_inclusive_divisor']
    return supply_amount, amount - supply_amount


def calculate_input_vat_credit(*amounts: int) -> int:
    """매입처리 가능한 부가세 합계 계산"""
    return sum(calculate_included_vat(amount) for amount in amounts)


def calculate_income_tax_estimate(annual_tax_base: int) -> dict:
    """종소세+지방소득세 예상액 계산"""
    tax_base = max(0, annual_tax_base)
    for upper_limit, tax_rate, progressive_deduction in INCOME_TAX_BRACKETS:
        if tax_base <= upper_limit:
            national_tax = max(0, round(tax_base * tax_rate - progressive_deduction))
            local_tax = round(national_tax * LOCAL_INCOME_TAX_RATE)
            return {
                'annual_tax_base': tax_base,
                'income_tax_rate': tax_rate,
                'progressive_deduction': progressive_deduction,
                'annual_income_tax': national_tax,
                'annual_local_income_tax': local_tax,
                'annual_income_tax_total': national_tax + local_tax,
            }
    raise ValueError('income tax bracket not found')


def calculate_monthly_income_tax_reserve(monthly_profit: int) -> dict:
    """해당 월 이익이 12개월 유지된다고 보고 월별 종소세 적립액 계산"""
    annual_tax_base = max(0, monthly_profit) * 12
    estimate = calculate_income_tax_estimate(annual_tax_base)
    monthly_reserve = round(estimate['annual_income_tax_total'] / 12)
    effective_rate = monthly_reserve / monthly_profit if monthly_profit > 0 else 0
    return {
        **estimate,
        'monthly_income_tax_reserve': monthly_reserve,
        'income_tax_effective_rate': effective_rate,
    }


def calculate_labor_cost(total_amount: int) -> int:
    """기준 월 인건비를 매출 비율로 환산"""
    FC = SHEETS_FIXED_COSTS
    baseline_labor_cost = FC['hourly_wage'] * FC['labor_hours_baseline']
    return round(total_amount * baseline_labor_cost / FC['labor_revenue_baseline'])


def calculate_profit_summary(total_amount: int) -> dict:
    """정산 매출 기준 영업이익 요약 계산"""
    FC = SHEETS_FIXED_COSTS
    labor_cost = calculate_labor_cost(total_amount)
    logistics_cost = FC['logistics_count'] * FC['logistics_cost_per']
    supplies_cost = calculate_supplies_cost(total_amount)
    withholding_national_tax = round(
        (logistics_cost + labor_cost) * FC['withholding_national_rate']
    )
    withholding_local_tax = round(
        (logistics_cost + labor_cost) * FC['withholding_local_rate']
    )
    withholding_tax = withholding_national_tax + withholding_local_tax
    output_vat = calculate_included_vat(total_amount)
    input_vat_credit = calculate_input_vat_credit(
        FC['rent_utility'],
        FC['electricity'],
        FC['water'],
        FC['insurance'],
        supplies_cost,
    )
    vat = max(0, output_vat - input_vat_credit)
    pre_vat_cost = (
        labor_cost
        + logistics_cost
        + FC['rent_utility']
        + FC['electricity']
        + FC['water']
        + FC['insurance']
        + supplies_cost
        + withholding_tax
    )
    total_cost = pre_vat_cost + vat
    operating_profit = total_amount - total_cost
    operating_margin = operating_profit / total_amount if total_amount else 0
    return {
        'revenue': total_amount,
        'labor_cost': labor_cost,
        'logistics_cost': logistics_cost,
        'rent_utility': FC['rent_utility'],
        'electricity': FC['electricity'],
        'water': FC['water'],
        'insurance': FC['insurance'],
        'supplies_cost': supplies_cost,
        'withholding_tax': withholding_tax,
        'withholding_national_tax': withholding_national_tax,
        'withholding_local_tax': withholding_local_tax,
        'output_vat': output_vat,
        'input_vat_credit': input_vat_credit,
        'vat': vat,
        'total_cost': total_cost,
        'operating_profit': operating_profit,
        'operating_margin': operating_margin,
    }


def calculate_logistics_payments(logistics_cost: int) -> list[dict]:
    """물류비 15일/30일 지급 일정 계산"""
    first_payment = logistics_cost // 2
    second_payment = logistics_cost - first_payment
    return [
        {'day': 15, 'amount': first_payment},
        {'day': 30, 'amount': second_payment},
    ]


def calculate_monthly_close_summary(
    total_amount: int,
    *,
    kops_receivable_amount: int = 0,
    current_bank_balance: int = 0,
    late_receipt_amount: int = 0,
    payroll_to_pay: int = 0,
    owner_draw: int = 0,
    income_tax_reserve_rate=None,
) -> dict:
    """발생 손익과 입금/지급 기준 현금 마감을 분리해 계산"""
    profit_summary = calculate_profit_summary(total_amount)
    operating_profit = profit_summary['operating_profit']
    if income_tax_reserve_rate is not None:
        income_tax_reserve_detail = {
            'annual_tax_base': max(0, operating_profit) * 12,
            'income_tax_rate': income_tax_reserve_rate,
            'progressive_deduction': 0,
            'annual_income_tax': 0,
            'annual_local_income_tax': 0,
            'annual_income_tax_total': round(operating_profit * income_tax_reserve_rate * 12),
            'monthly_income_tax_reserve': round(operating_profit * income_tax_reserve_rate),
            'income_tax_effective_rate': income_tax_reserve_rate if operating_profit > 0 else 0,
        }
    else:
        income_tax_reserve_detail = calculate_monthly_income_tax_reserve(operating_profit)
    income_tax_reserve = income_tax_reserve_detail['monthly_income_tax_reserve']
    net_profit_after_income_tax_reserve = operating_profit - income_tax_reserve
    cash_after_late_receipt = current_bank_balance + late_receipt_amount
    cash_after_payroll = cash_after_late_receipt - payroll_to_pay

    return {
        'profit_summary': profit_summary,
        'income_tax_reserve_rate': income_tax_reserve_rate,
        'income_tax_reserve_detail': income_tax_reserve_detail,
        'income_tax_reserve': income_tax_reserve,
        'net_profit_after_income_tax_reserve': net_profit_after_income_tax_reserve,
        'owner_draw': owner_draw,
        'remaining_profit_after_owner_draw': net_profit_after_income_tax_reserve - owner_draw,
        'cash_timing': {
            'month_end_collection': total_amount - kops_receivable_amount,
            'delayed_kops_collection': kops_receivable_amount,
            'current_bank_balance': current_bank_balance,
            'late_receipt_amount': late_receipt_amount,
            'cash_after_late_receipt': cash_after_late_receipt,
            'payroll_to_pay': payroll_to_pay,
            'cash_after_payroll': cash_after_payroll,
            'logistics_payments': calculate_logistics_payments(profit_summary['logistics_cost']),
        },
    }


def calculate_business_total(data: dict) -> int:
    """사업자 1곳의 정산 총액 계산"""
    total = 0
    for location, loc_data in data.get('locations', {}).items():
        prices = get_location_prices(location)
        for item_key in prices:
            total += (loc_data.get(item_key, 0) or 0) * prices[item_key]
    for item in data.get('extra_items', []):
        total += item['amount']
    return total


def calculate_total_amount_from_business_data(business_data: dict) -> int:
    """사업자별 집계에서 전체 정산 총액 계산"""
    return sum(calculate_business_total(data) for data in business_data.values())


def calculate_kops_receivable_amount(business_data: dict) -> int:
    """콥스 익월 입금 대상 정산액 계산"""
    return calculate_business_total(business_data.get(KOPS_REG_NO, {}))


def format_won(amount: int) -> str:
    """원화 숫자 포맷"""
    return f"{amount:,}원"


def get_vat_credit_breakdown(summary: dict) -> list[tuple[str, int]]:
    """매입세액공제 항목별 부가세 계산"""
    return [
        ('월세+관리비 매입세액', calculate_included_vat(summary['rent_utility'])),
        ('전기세 매입세액', calculate_included_vat(summary['electricity'])),
        ('수도세 매입세액', calculate_included_vat(summary['water'])),
        ('보험 매입세액', calculate_included_vat(summary['insurance'])),
        ('소모품 매입세액', calculate_included_vat(summary['supplies_cost'])),
    ]


def get_operating_expense_rows(summary: dict) -> list[tuple[str, int]]:
    """운영비 항목"""
    return [
        ('인건비', summary['labor_cost']),
        ('기사비', summary['logistics_cost']),
        ('월세+관리비', summary['rent_utility']),
        ('전기세', summary['electricity']),
        ('수도세', summary['water']),
        ('보험', summary['insurance']),
        ('소모품', summary['supplies_cost']),
    ]


def get_tax_reserve_rows(summary: dict, close_summary: dict = None) -> list[tuple[str, int]]:
    """세금과 세금 적립 항목"""
    rows = [
        ('원천세', summary['withholding_tax']),
        ('부가세 납부 예상액', summary['vat']),
    ]
    if close_summary:
        rows.append(('종소세+지방소득세 적립액(예상)', close_summary['income_tax_reserve']))
    return rows


def calculate_operating_expense_total(summary: dict) -> int:
    """운영비 합계"""
    return sum(amount for _, amount in get_operating_expense_rows(summary))


def calculate_tax_reserve_total(summary: dict, close_summary: dict = None) -> int:
    """세금과 세금 적립 합계"""
    return sum(amount for _, amount in get_tax_reserve_rows(summary, close_summary))


def format_profit_summary_text(summary: dict, close_summary: dict = None) -> str:
    """메일 본문용 영업이익 요약 텍스트"""
    operating_expense_total = calculate_operating_expense_total(summary)
    tax_reserve_total = calculate_tax_reserve_total(summary, close_summary)
    final_profit = (
        close_summary['net_profit_after_income_tax_reserve']
        if close_summary
        else summary['operating_profit']
    )
    operating_expense_lines = '\n'.join(
        f"- {label}: {format_won(amount)}"
        for label, amount in get_operating_expense_rows(summary)
    )
    tax_reserve_lines = '\n'.join(
        f"- {label}: {format_won(amount)}"
        for label, amount in get_tax_reserve_rows(summary, close_summary)
    )
    vat_credit_lines = '\n'.join(
        f"- {label}: {format_won(amount)}"
        for label, amount in get_vat_credit_breakdown(summary)
    )
    income_tax_note = ""
    if close_summary:
        detail = close_summary['income_tax_reserve_detail']
        income_tax_note = f"""
종소세 적립 기준
- 월 영업이익을 12개월로 환산한 예상 과세표준: {format_won(detail['annual_tax_base'])}
- 적용 세율: {detail['income_tax_rate']:.0%}, 누진공제: {format_won(detail['progressive_deduction'])}
- 연간 예상 종소세+지방소득세: {format_won(detail['annual_income_tax_total'])}
- 월 적립액: {format_won(close_summary['income_tax_reserve'])}"""
    return f"""수입
- 정산 매출: {format_won(summary['revenue'])}

지출
- 운영비 합계: {format_won(operating_expense_total)}
{operating_expense_lines}
- 세금/적립 합계: {format_won(tax_reserve_total)}
{tax_reserve_lines}

매입세액공제 항목
{vat_credit_lines}
- 매입세액공제 합계: {format_won(summary['input_vat_credit'])}

부가세 계산
- 매출 부가세: {format_won(summary['output_vat'])}
- 매입세액공제: {format_won(summary['input_vat_credit'])}
- 납부 예상 부가세: {format_won(summary['vat'])}
{income_tax_note}

최종 순수익
- 부가세·원천세 반영 후 영업이익: {format_won(summary['operating_profit'])}
- 종소세 적립 후 순수익: {format_won(final_profit)}
- 영업이익률: {summary['operating_margin']:.1%}"""


def format_monthly_close_text(summary: dict) -> str:
    """메일 본문용 월마감 현금 요약 텍스트"""
    cash = summary['cash_timing']
    logistics_payments = ' / '.join(
        f"{item['day']}일 {format_won(item['amount'])}"
        for item in cash['logistics_payments']
    )
    lines = [
        "월마감 현금 기준",
        f"- 말일 입금(콥스 제외): {format_won(cash['month_end_collection'])}",
        f"- 콥스 익월 20일 입금: {format_won(cash['delayed_kops_collection'])}",
        f"- 물류비 지급 일정: {logistics_payments}",
    ]
    if cash['current_bank_balance']:
        lines.append(f"- 현재 통장 잔액: {format_won(cash['current_bank_balance'])}")
    if cash['late_receipt_amount']:
        lines.append(f"- 콥스 입금 후 현금: {format_won(cash['cash_after_late_receipt'])}")
    if cash['payroll_to_pay']:
        lines.append(f"- 인건비 지급 후 현금: {format_won(cash['cash_after_payroll'])}")
    lines.extend([
        f"- 종소세+지방소득세 적립액(예상): {format_won(summary['income_tax_reserve'])}",
        f"- 종소세 적립 후 순수익: {format_won(summary['net_profit_after_income_tax_reserve'])}",
    ])
    if summary['owner_draw']:
        lines.extend([
            f"- 이미 빼간 영업이익: {format_won(summary['owner_draw'])}",
            f"- 추가로 남길 수 있는 이익: {format_won(summary['remaining_profit_after_owner_draw'])}",
        ])
    return '\n'.join(lines)


def format_profit_summary_html(summary: dict, close_summary: dict = None) -> str:
    """내부 레포트 이메일용 영업이익 요약 HTML"""
    operating_expense_total = calculate_operating_expense_total(summary)
    tax_reserve_total = calculate_tax_reserve_total(summary, close_summary)
    final_profit = (
        close_summary['net_profit_after_income_tax_reserve']
        if close_summary
        else summary['operating_profit']
    )
    vat_credit_rows = get_vat_credit_breakdown(summary) + [
        ('매입세액공제 합계', summary['input_vat_credit']),
    ]
    vat_rows = [
        ('매출 부가세', summary['output_vat']),
        ('매입세액공제', summary['input_vat_credit']),
        ('납부 예상 부가세', summary['vat']),
    ]
    income_tax_rows = []
    if close_summary:
        detail = close_summary['income_tax_reserve_detail']
        income_tax_rows = [
            ('월 영업이익 연간 환산액', detail['annual_tax_base']),
            ('누진공제', detail['progressive_deduction']),
            ('연간 예상 종소세+지방소득세', detail['annual_income_tax_total']),
            ('월 적립액', close_summary['income_tax_reserve']),
        ]
    final_rows = [
        ('부가세·원천세 반영 후 영업이익', summary['operating_profit']),
        ('종소세 적립 후 순수익', final_profit),
    ]
    final_margin = final_profit / summary['revenue'] if summary['revenue'] else 0

    summary_rows = [
        ('수입', '정산 매출', format_won(summary['revenue']), '이번 달 청구 기준 총매출', True),
    ]
    summary_rows.extend(
        ('지출 - 운영비', label, format_won(amount), '매출을 만들기 위해 실제로 들어간 비용', False)
        for label, amount in get_operating_expense_rows(summary)
    )
    summary_rows.append(
        ('지출 - 운영비', '운영비 합계', format_won(operating_expense_total), '인건비·기사비·공간비·소모품 합계', True)
    )
    summary_rows.extend(
        ('지출 - 세금/적립', label, format_won(amount), '납부 또는 별도 적립 대상', False)
        for label, amount in get_tax_reserve_rows(summary, close_summary)
    )
    summary_rows.append(
        ('지출 - 세금/적립', '세금/적립 합계', format_won(tax_reserve_total), '원천세·부가세·종소세 적립 합계', True)
    )
    summary_rows.extend([
        ('최종', '부가세·원천세 반영 후 영업이익', format_won(summary['operating_profit']), '종소세 적립 전 남은 돈', False),
        ('최종', '종소세 적립 후 순수익', format_won(final_profit), '이번 달 실제로 가져갈 수 있는 돈', True),
        ('최종', '최종 순수익률', f"{final_margin:.1%}", '정산 매출 대비 실제 순수익', True),
    ])

    def render_rows(rows: list[tuple[str, int]], strong_last: bool = False) -> str:
        html_rows = []
        for i, (label, amount) in enumerate(rows):
            strong = strong_last and i == len(rows) - 1
            label_style = 'font-weight:bold;color:#111827;' if strong else 'color:#475569;'
            amount_style = 'font-weight:bold;color:#111827;' if strong else 'color:#111827;'
            html_rows.append(
                f'<tr><td style="padding:4px 0;{label_style}">{label}</td>'
                f'<td style="padding:4px 0;text-align:right;{amount_style}">{format_won(amount)}</td></tr>'
            )
        return ''.join(html_rows)

    def render_section(title: str, rows: list[tuple[str, int]], strong_last: bool = False) -> str:
        return f"""
  <h4 style="margin:14px 0 6px;font-size:13px;color:#334155;">{title}</h4>
  <table style="border-collapse:collapse;width:100%;">
    {render_rows(rows, strong_last)}
  </table>"""

    def render_summary_table() -> str:
        html_rows = []
        for section, label, value, note, strong in summary_rows:
            row_style = 'background:#f8fafc;' if strong else ''
            label_style = 'font-weight:bold;color:#111827;' if strong else 'color:#111827;'
            value_style = 'font-weight:bold;color:#047857;' if section == '최종' and strong else 'font-weight:bold;color:#111827;' if strong else 'color:#111827;'
            html_rows.append(
                f'<tr style="{row_style}">'
                f'<td style="padding:7px 8px;border-bottom:1px solid #e5e7eb;color:#475569;font-size:13px;white-space:nowrap;">{section}</td>'
                f'<td style="padding:7px 8px;border-bottom:1px solid #e5e7eb;font-size:13px;{label_style}">{label}</td>'
                f'<td style="padding:7px 8px;border-bottom:1px solid #e5e7eb;text-align:right;font-size:13px;white-space:nowrap;{value_style}">{value}</td>'
                f'<td style="padding:7px 8px;border-bottom:1px solid #e5e7eb;color:#64748b;font-size:12px;">{note}</td>'
                f'</tr>'
            )
        return f"""
  <table style="border-collapse:collapse;width:100%;margin-top:8px;">
    <thead>
      <tr>
        <th style="background:#047857;color:#fff;padding:8px;text-align:left;font-size:12px;">구분</th>
        <th style="background:#047857;color:#fff;padding:8px;text-align:left;font-size:12px;">항목</th>
        <th style="background:#047857;color:#fff;padding:8px;text-align:right;font-size:12px;">금액</th>
        <th style="background:#047857;color:#fff;padding:8px;text-align:left;font-size:12px;">설명</th>
      </tr>
    </thead>
    <tbody>{''.join(html_rows)}</tbody>
  </table>"""

    return f"""
<div style="background:#ecfdf5;border:1px solid #bbf7d0;border-radius:8px;padding:16px;margin-bottom:20px;">
  <h3 style="margin:0 0 12px;font-size:15px;color:#047857;">수입·지출·순수익 요약</h3>
  {render_summary_table()}
  {render_section('매입세액공제 항목', vat_credit_rows, True)}
  {render_section('부가세 계산', vat_rows, True)}
  {render_section('종소세 적립 기준', income_tax_rows, True) if income_tax_rows else ''}
  {render_section('최종 순수익', final_rows, True)}
  <p style="margin:6px 0 0;text-align:right;font-size:13px;color:#047857;font-weight:bold;">영업이익률 {summary['operating_margin']:.1%}</p>
</div>"""


def format_monthly_close_html(summary: dict) -> str:
    """내부 레포트 이메일용 월마감 현금 요약 HTML"""
    cash = summary['cash_timing']
    logistics_payments = ' / '.join(
        f"{item['day']}일 {format_won(item['amount'])}"
        for item in cash['logistics_payments']
    )
    rows = [
        ('말일 입금(콥스 제외)', format_won(cash['month_end_collection'])),
        ('콥스 익월 20일 입금', format_won(cash['delayed_kops_collection'])),
        ('물류비 지급 일정', logistics_payments),
    ]
    if cash['current_bank_balance']:
        rows.append(('현재 통장 잔액', format_won(cash['current_bank_balance'])))
    if cash['late_receipt_amount']:
        rows.append(('콥스 입금 후 현금', format_won(cash['cash_after_late_receipt'])))
    if cash['payroll_to_pay']:
        rows.append(('인건비 지급 후 현금', format_won(cash['cash_after_payroll'])))
    rows.extend([
        ('종소세+지방소득세 적립액(예상)', format_won(summary['income_tax_reserve'])),
        ('종소세 적립 후 순수익', format_won(summary['net_profit_after_income_tax_reserve'])),
    ])
    if summary['owner_draw']:
        rows.extend([
            ('이미 빼간 영업이익', format_won(summary['owner_draw'])),
            ('추가로 남길 수 있는 이익', format_won(summary['remaining_profit_after_owner_draw'])),
        ])
    row_html = ''.join(
        f'<tr><td style="padding:5px 0;color:#475569;">{label}</td>'
        f'<td style="padding:5px 0;text-align:right;color:#111827;">{value}</td></tr>'
        for label, value in rows
    )
    return f"""
<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin-bottom:20px;">
  <h3 style="margin:0 0 12px;font-size:15px;color:#334155;">월마감 현금 기준</h3>
  <table style="border-collapse:collapse;width:100%;">
    {row_html}
  </table>
</div>"""


def calculate_record_amount(row: tuple) -> int:
    """세탁 기록 1건의 정산 금액 계산"""
    record_date, location, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket = row
    if location not in BUSINESS_MAP:
        return 0
    if not is_settlement_location_active(location, record_date):
        return 0

    prices = get_location_prices(location)
    return (
        (blanket or 0) * prices.get('blanket', 0) +
        (mat or 0) * prices.get('mat', 0) +
        (pillow_cover or 0) * prices.get('pillow_cover', 0) +
        (towel or 0) * prices.get('towel', 0) +
        (body_towel or 0) * prices.get('body_towel', 0) +
        (pillow_fill or 0) * prices.get('pillow_fill', 0) +
        (cotton_blanket or 0) * prices.get('cotton_blanket', 0)
    )


# ============================================================
# 데이터 조회
# ============================================================

def get_monthly_data(year: int, month: int) -> list:
    """해당 월의 세탁물 데이터 조회"""
    start_date, end_exclusive = get_settlement_period(year, month)
    conn = psycopg2.connect(SUPABASE_URI)
    cur = conn.cursor()

    cur.execute("""
        SELECT record_date, location, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket
        FROM laundry_records
        WHERE record_date >= %s
          AND record_date < %s
        ORDER BY location, record_date
    """, (start_date, end_exclusive))

    rows = cur.fetchall()
    cur.close()
    conn.close()

    return rows


def aggregate_by_business(rows: list) -> dict:
    """사업자별로 데이터 집계"""
    business_data = {}

    for row in rows:
        record_date, location, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket = row

        if location not in BUSINESS_MAP:
            print(f"알 수 없는 숙소: {location}")
            continue
        if not is_settlement_location_active(location, record_date):
            continue

        reg_no, biz_name, owner = BUSINESS_MAP[location]

        if reg_no not in business_data:
            business_data[reg_no] = {
                'name': biz_name,
                'owner': owner,
                'locations': {},
                'extra_items': []
            }

        if location not in business_data[reg_no]['locations']:
            business_data[reg_no]['locations'][location] = {
                'blanket': 0, 'mat': 0, 'pillow_cover': 0,
                'towel': 0, 'body_towel': 0, 'pillow_fill': 0, 'cotton_blanket': 0
            }

        loc_data = business_data[reg_no]['locations'][location]
        loc_data['blanket'] += blanket or 0
        loc_data['mat'] += mat or 0
        loc_data['pillow_cover'] += pillow_cover or 0
        loc_data['towel'] += towel or 0
        loc_data['body_towel'] += body_towel or 0
        loc_data['pillow_fill'] += pillow_fill or 0
        loc_data['cotton_blanket'] += cotton_blanket or 0

    return business_data


# ============================================================
# PDF 생성
# ============================================================

def register_font():
    """한글 폰트 등록"""
    # 스크립트와 같은 디렉토리의 번들 폰트를 최우선 사용
    script_dir = os.path.dirname(os.path.abspath(__file__))
    font_paths = [
        os.path.join(script_dir, 'NotoSansKR.ttf'),           # 번들 폰트 (최우선)
        '/usr/share/fonts/truetype/noto/NotoSansKR-Regular.ttf',  # Ubuntu
        'NotoSansKR.ttf',                                      # 로컬 현재 디렉토리
    ]

    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('Korean', path))
                return True
            except:
                continue

    return False


def generate_pdf(reg_no: str, data: dict, year: int, month: int) -> BytesIO:
    """거래명세서 PDF 생성"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                           leftMargin=15*mm, rightMargin=15*mm,
                           topMargin=15*mm, bottomMargin=15*mm)

    # 폰트 등록
    font_registered = register_font()
    font_name = 'Korean' if font_registered else 'Helvetica'

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Title'],
                                  fontName=font_name, fontSize=18, alignment=1)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'],
                                   fontName=font_name, fontSize=10)
    heading_style = ParagraphStyle('Heading', parent=styles['Normal'],
                                    fontName=font_name, fontSize=12,
                                    textColor=colors.HexColor('#1e40af'))

    elements = []

    # 제목
    elements.append(Paragraph(f"{year}년 {month}월 거래명세서", title_style))
    elements.append(Spacer(1, 10*mm))

    # 공급자/공급받는자 정보
    info_data = [
        ['공급자', '', '공급받는자', ''],
        ['상호', SUPPLIER['name'], '상호', data['name']],
        ['사업자번호', SUPPLIER['reg_no'], '사업자번호', reg_no],
        ['대표자', SUPPLIER['owner'], '대표자', data['owner']],
    ]

    info_table = Table(info_data, colWidths=[25*mm, 55*mm, 25*mm, 55*mm])
    info_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#f0f0f0')),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#f0f0f0')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 8*mm))

    # 숙소별 내역
    grand_total = 0

    for idx, (location, loc_data) in enumerate(data['locations'].items(), 1):
        elements.append(Paragraph(f"[{idx}] {location}", heading_style))
        elements.append(Spacer(1, 3*mm))

        # 품목 테이블
        item_rows = [['품목', '수량', '단가', '금액']]
        loc_total = 0

        prices = get_location_prices(location)
        for item_key, item_name in ITEM_NAMES.items():
            qty = loc_data.get(item_key, 0)
            if qty > 0:
                price = prices[item_key]
                amount = qty * price
                loc_total += amount
                item_rows.append([item_name, f"{qty:,}", f"{price:,}원", f"{amount:,}원"])

        item_rows.append(['소계', '', '', f"{loc_total:,}원"])
        grand_total += loc_total

        item_table = Table(item_rows, colWidths=[50*mm, 30*mm, 35*mm, 45*mm])
        item_table.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), font_name),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3b82f6')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eff6ff')),
            ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        elements.append(item_table)
        elements.append(Spacer(1, 5*mm))

    # 기타 항목
    if data.get('extra_items'):
        elements.append(Paragraph("[기타]", heading_style))
        elements.append(Spacer(1, 3*mm))

        extra_rows = [['항목', '수량', '단가', '금액']]
        for item in data['extra_items']:
            grand_total += item['amount']
            extra_rows.append([
                item['name'],
                f"{item['qty']:,}",
                f"{item['price']:,}원",
                f"{item['amount']:,}원"
            ])

        extra_table = Table(extra_rows, colWidths=[50*mm, 30*mm, 35*mm, 45*mm])
        extra_table.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), font_name),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f97316')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ]))
        elements.append(extra_table)
        elements.append(Spacer(1, 5*mm))

    # 합계
    supply_amount, tax_amount = split_vat_inclusive_amount(grand_total)

    total_data = [
        ['공급가액', f"{supply_amount:,}원"],
        ['부가세', f"{tax_amount:,}원"],
        ['합계', f"{grand_total:,}원"],
    ]

    total_table = Table(total_data, colWidths=[50*mm, 110*mm])
    total_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('FONTSIZE', (0,0), (-1,-1), 11),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#f0f0f0')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#1e40af')),
        ('TEXTCOLOR', (0,-1), (-1,-1), colors.white),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(total_table)
    elements.append(Spacer(1, 8*mm))

    # 입금 계좌
    elements.append(Paragraph(f"입금계좌: {SUPPLIER['bank']}", normal_style))

    # 통장 사본 이미지
    script_dir = os.path.dirname(os.path.abspath(__file__))
    bank_img_path = os.path.join(script_dir, '캐리_사업자통장사본.png')
    if os.path.exists(bank_img_path):
        from reportlab.lib.utils import ImageReader
        ir = ImageReader(bank_img_path)
        orig_w, orig_h = ir.getSize()
        target_w = 160 * mm
        target_h = target_w * orig_h / orig_w
        elements.append(Spacer(1, 8*mm))
        img = Image(bank_img_path, width=target_w, height=target_h)
        img.hAlign = 'LEFT'
        elements.append(img)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ============================================================
# Excel 생성
# ============================================================

def generate_excel(business_data: dict, year: int, month: int) -> BytesIO:
    """홈택스 세금계산서 Excel 생성"""
    wb = Workbook()
    ws = wb.active
    ws.title = "세금계산서"

    # 헤더
    headers = ['작성일자', '공급받는자 사업자번호', '공급받는자 상호',
               '공급받는자 대표자', '공급가액', '세액', '합계금액', '품목']
    ws.append(headers)

    # 작성일자 (해당 월 말일)
    last_day = calendar.monthrange(year, month)[1]
    write_date = f"{year}{month:02d}{last_day:02d}"

    for reg_no, data in business_data.items():
        # 총액 계산
        total = calculate_business_total(data)

        supply_amount, tax_amount = split_vat_inclusive_amount(total)

        ws.append([
            write_date,
            reg_no.replace('-', ''),
            data['name'],
            data['owner'],
            supply_amount,
            tax_amount,
            total,
            '세탁 서비스'
        ])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# 이메일 발송
# ============================================================

def send_email(subject: str, body: str, attachments: list):
    """이메일 발송"""
    if not EMAIL_PASSWORD:
        print("EMAIL_PASSWORD가 설정되지 않음. 이메일 발송 건너뜀.")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
    if EMAIL_CC:
        msg['Cc'] = EMAIL_CC
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    for filename, data in attachments:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(data.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

    recipients = [EMAIL_TO] + ([EMAIL_CC] if EMAIL_CC else [])
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        refused = server.sendmail(EMAIL_FROM, recipients, msg.as_string())
        server.quit()
        if refused:
            print(f"이메일 일부 수신 거부: {refused}")
        else:
            print(f"이메일 발송 완료 → {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"이메일 발송 실패: {e}")
        return False


def _get_6month_trend(year: int, month: int) -> list:
    """최근 6개월 월별 매출 합계 반환 [(year, month, total), ...]"""
    # 시작월 계산 (5개월 전)
    months = []
    y, m = year, month
    for _ in range(6):
        months.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    months.reverse()  # 오래된 순서로

    try:
        conn = psycopg2.connect(SUPABASE_URI)
        cur = conn.cursor()
        start_y, start_m = months[0]
        start_date, _ = get_settlement_period(start_y, start_m)
        _, end_exclusive = get_settlement_period(year, month)
        cur.execute("""
            SELECT record_date,
                   location,
                   COALESCE(SUM(blanket),0)      AS blanket,
                   COALESCE(SUM(mat),0)          AS mat,
                   COALESCE(SUM(pillow_cover),0) AS pillow_cover,
                   COALESCE(SUM(towel),0)        AS towel,
                   COALESCE(SUM(body_towel),0)   AS body_towel,
                   COALESCE(SUM(pillow_fill),0)  AS pillow_fill,
                   COALESCE(SUM(cotton_blanket),0) AS cotton_blanket
            FROM laundry_records
            WHERE record_date >= %s
              AND record_date < %s
            GROUP BY record_date, location
            ORDER BY record_date
        """, (start_date, end_exclusive))
        loc_rows = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"6개월 추이 조회 실패: {e}")
        return [(y, m, 0) for y, m in months]

    # 월별 합산
    monthly = {(y, m): 0 for y, m in months}
    for record_date, loc, bl, mt, pc, tw, bt, pf, cb in loc_rows:
        if loc not in BUSINESS_MAP:
            continue
        if not is_settlement_location_active(loc, record_date):
            continue
        p = get_location_prices(loc)
        amt = bl*p['blanket'] + mt*p['mat'] + pc*p['pillow_cover'] + tw*p['towel'] + bt*p['body_towel'] + pf*p['pillow_fill'] + cb*p.get('cotton_blanket', 0)
        key = get_record_settlement_month(record_date)
        if key in monthly:
            monthly[key] += amt

    return [(y, m, monthly[(y, m)]) for y, m in months]


def send_report_email(year: int, month: int, rows: list, business_data: dict) -> bool:
    """내부 검토용 정산 레포트 이메일 발송 (HTML 테이블)"""
    if not EMAIL_PASSWORD:
        print("EMAIL_PASSWORD 미설정. 레포트 이메일 건너뜀.")
        return False

    from collections import defaultdict

    weekday_ko = ['월', '화', '수', '목', '금', '토', '일']

    daily = defaultdict(list)
    for row in rows:
        record_date, location, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket = row
        if not is_settlement_location_active(location, record_date):
            continue
        daily[location].append((record_date, blanket or 0, mat or 0,
                                 pillow_cover or 0, towel or 0, body_towel or 0, pillow_fill or 0, cotton_blanket or 0))
    for loc in daily:
        daily[loc].sort(key=lambda x: x[0])

    # 공통 테이블 스타일 (인라인 CSS, Gmail 호환)
    TH = 'style="background:#1e40af;color:#fff;padding:6px 10px;text-align:right;font-size:13px;white-space:nowrap;"'
    TH_L = 'style="background:#1e40af;color:#fff;padding:6px 10px;text-align:left;font-size:13px;"'
    TD = 'style="padding:5px 10px;text-align:right;font-size:13px;border-bottom:1px solid #e5e7eb;"'
    TD_L = 'style="padding:5px 10px;text-align:left;font-size:13px;border-bottom:1px solid #e5e7eb;"'
    TD_SUB = 'style="padding:5px 10px;text-align:right;font-size:13px;background:#eff6ff;font-weight:bold;"'
    TD_SUB_L = 'style="padding:5px 10px;text-align:left;font-size:13px;background:#eff6ff;font-weight:bold;"'

    # 주별/숙소별 매출 브리핑
    weekly = defaultdict(lambda: defaultdict(int))
    weekly_totals = defaultdict(int)
    period_start, period_end_exclusive = get_settlement_period(year, month)
    period_end = period_end_exclusive - timedelta(days=1)
    for row in rows:
        record_date, location, *_ = row
        amount = calculate_record_amount(row)
        if amount == 0:
            continue
        week_start = record_date - timedelta(days=record_date.weekday())
        week_end = week_start + timedelta(days=6)
        label_start = max(week_start, period_start)
        label_end = min(week_end, period_end)
        week_label = f"{label_start.month}/{label_start.day:02d}~{label_end.month}/{label_end.day:02d}"
        weekly[week_label][location] += amount
        weekly_totals[week_label] += amount

    location_names = sorted({loc for values in weekly.values() for loc in values.keys()})
    week_header = ''.join(f'<th {TH}>{loc}</th>' for loc in location_names)
    week_rows = []
    for week_label in sorted(weekly.keys()):
        cells = ''.join(f'<td {TD}>{weekly[week_label].get(loc, 0):,}원</td>' for loc in location_names)
        week_rows.append(
            f'<tr><td {TD_L}>{week_label}</td>{cells}<td {TD_SUB}>{weekly_totals[week_label]:,}원</td></tr>'
        )
    weekly_html = f"""
<div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:16px;margin-bottom:20px;">
  <h3 style="margin:0 0 12px;font-size:15px;color:#c2410c;">주별 매출 브리핑</h3>
  <table style="border-collapse:collapse;width:100%;">
    <thead><tr><th {TH_L}>주차</th>{week_header}<th {TH}>합계</th></tr></thead>
    <tbody>{''.join(week_rows)}</tbody>
  </table>
</div>""" if week_rows else ""

    grand_total = 0
    sections = []

    for reg_no, data in business_data.items():
        biz_total = 0
        loc_tables = []

        for location in sorted(data['locations'].keys()):
            prices = get_location_prices(location)
            loc_rows = daily.get(location, [])
            loc_total = 0
            data_rows_html = []

            for record_date, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket in loc_rows:
                if blanket + mat + pillow_cover + towel + body_towel + pillow_fill + cotton_blanket == 0:
                    continue
                row_amount = (blanket * prices.get('blanket', 0) +
                              mat * prices.get('mat', 0) +
                              pillow_cover * prices.get('pillow_cover', 0) +
                              towel * prices.get('towel', 0) +
                              body_towel * prices.get('body_towel', 0) +
                              pillow_fill * prices.get('pillow_fill', 0) +
                              cotton_blanket * prices.get('cotton_blanket', 0))
                loc_total += row_amount
                wd = weekday_ko[record_date.weekday()]
                date_str = f"{record_date.month}/{record_date.day:02d}({wd})"
                data_rows_html.append(
                    f"<tr><td {TD_L}>{date_str}</td>"
                    f"<td {TD}>{blanket}</td><td {TD}>{mat}</td>"
                    f"<td {TD}>{pillow_cover}</td><td {TD}>{towel}</td>"
                    f"<td {TD}>{body_towel}</td><td {TD}>{pillow_fill}</td>"
                    f"<td {TD}>{cotton_blanket}</td>"
                    f"<td {TD}>{row_amount:,}원</td></tr>"
                )

            loc_qty = data['locations'][location]
            biz_total += loc_total
            sub_bt = loc_qty.get('body_towel', 0)
            sub_pf = loc_qty.get('pillow_fill', 0)
            sub_cb = loc_qty.get('cotton_blanket', 0)
            data_rows_html.append(
                f"<tr><td {TD_SUB_L}>소계</td>"
                f"<td {TD_SUB}>{loc_qty['blanket']}</td><td {TD_SUB}>{loc_qty['mat']}</td>"
                f"<td {TD_SUB}>{loc_qty['pillow_cover']}</td><td {TD_SUB}>{loc_qty['towel']}</td>"
                f"<td {TD_SUB}>{sub_bt}</td><td {TD_SUB}>{sub_pf}</td>"
                f"<td {TD_SUB}>{sub_cb}</td>"
                f"<td {TD_SUB}>{loc_total:,}원</td></tr>"
            )

            loc_tables.append(f"""
<p style="margin:16px 0 6px;font-size:14px;font-weight:bold;color:#374151;">▶ {location}</p>
<table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
  <thead><tr>
    <th {TH_L}>날짜</th><th {TH}>이불</th><th {TH}>매트</th>
    <th {TH}>베개커버</th><th {TH}>타올</th><th {TH}>바디타올</th><th {TH}>배게 솜</th><th {TH}>솜 이불</th><th {TH}>금액</th>
  </tr></thead>
  <tbody>{''.join(data_rows_html)}</tbody>
</table>""")

        if data.get('extra_items'):
            extra_rows = []
            for item in data['extra_items']:
                biz_total += item['amount']
                extra_rows.append(
                    f"<tr><td {TD_L}>{item['name']}</td>"
                    f"<td {TD}>{item['qty']:,}</td><td {TD}>{item['price']:,}원</td>"
                    f"<td {TD}>{item['amount']:,}원</td></tr>"
                )
            loc_tables.append(f"""
<p style="margin:16px 0 6px;font-size:14px;font-weight:bold;color:#374151;">▶ 기타 차감/추가</p>
<table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
  <thead><tr>
    <th {TH_L}>항목</th><th {TH}>수량</th><th {TH}>단가</th><th {TH}>금액</th>
  </tr></thead>
  <tbody>{''.join(extra_rows)}</tbody>
</table>""")

        grand_total += biz_total
        sections.append(f"""
<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin-bottom:20px;">
  <h3 style="margin:0 0 12px;font-size:16px;color:#1e40af;">
    {data['name']} <span style="font-size:13px;color:#6b7280;font-weight:normal;">({reg_no})</span>
  </h3>
  {''.join(loc_tables)}
  <p style="text-align:right;font-size:15px;font-weight:bold;color:#1e40af;margin:8px 0 0;">
    {data['name']} 합계: {biz_total:,}원
  </p>
</div>""")

    # 최근 6개월 추이
    trend = _get_6month_trend(year, month)
    max_amt = max(t[2] for t in trend) or 1
    trend_rows = []
    for t_y, t_m, t_amt in trend:
        bar_pct = int(t_amt / max_amt * 100)
        is_cur = (t_y == year and t_m == month)
        bar_color = '#1e40af' if is_cur else '#93c5fd'
        label_style = 'font-weight:bold;color:#1e40af;' if is_cur else 'color:#374151;'
        trend_rows.append(
            f'<tr>'
            f'<td style="padding:4px 10px;font-size:13px;{label_style}white-space:nowrap;">{t_y}년 {t_m}월</td>'
            f'<td style="padding:4px 10px;width:100%;">'
            f'  <div style="background:{bar_color};height:16px;width:{bar_pct}%;border-radius:3px;min-width:2px;"></div>'
            f'</td>'
            f'<td style="padding:4px 10px;text-align:right;font-size:13px;{label_style}white-space:nowrap;">'
            f'  {t_amt:,}원{"  ◀ 이번달" if is_cur else ""}'
            f'</td>'
            f'</tr>'
        )
    trend_html = f"""
<div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:16px;margin-bottom:20px;">
  <h3 style="margin:0 0 12px;font-size:15px;color:#0369a1;">최근 6개월 매출 추이</h3>
  <table style="border-collapse:collapse;width:100%;">
    {''.join(trend_rows)}
  </table>
</div>"""
    profit_summary = calculate_profit_summary(grand_total)
    monthly_close_summary = calculate_monthly_close_summary(
        grand_total,
        kops_receivable_amount=calculate_kops_receivable_amount(business_data),
    )
    profit_summary_html = format_profit_summary_html(profit_summary, monthly_close_summary)
    monthly_close_html = format_monthly_close_html(monthly_close_summary)

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:'Apple SD Gothic Neo',Arial,sans-serif;max-width:720px;margin:0 auto;padding:20px;color:#111827;">
  <h2 style="margin:0 0 4px;font-size:20px;color:#111827;">[캐리] {year}년 {month}월 정산 내부 레포트</h2>
  <p style="margin:0 0 20px;font-size:13px;color:#6b7280;">정산 기간 {format_settlement_period(year, month)} · 레코드 {len(rows)}건 · 사업자 {len(business_data)}개</p>
  {trend_html}
  {profit_summary_html}
  {monthly_close_html}
  {weekly_html}
  {''.join(sections)}
  <div style="background:#1e40af;color:#fff;border-radius:8px;padding:14px 20px;text-align:right;font-size:18px;font-weight:bold;">
    {month}월 총 매출: {grand_total:,}원
  </div>
</body></html>"""

    subject = f"[캐리 레포트] {year}년 {month}월 정산 검토"
    try:
        msg = MIMEMultipart('alternative')
        recipients = [EMAIL_TO] + ([EMAIL_CC] if EMAIL_CC else [])
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        if EMAIL_CC:
            msg['Cc'] = EMAIL_CC
        msg['Subject'] = subject
        msg.attach(MIMEText(html, 'html', 'utf-8'))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            refused = server.sendmail(EMAIL_FROM, recipients, msg.as_string())
        if refused:
            print(f"레포트 이메일 일부 수신 거부: {refused}")
        else:
            print(f"레포트 이메일 발송 완료 → {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"레포트 이메일 발송 실패: {e}")
        return False


# ============================================================
# Google Sheets 연동 함수
# ============================================================

def get_sheets_token() -> str:
    """서비스 계정 JSON 환경변수 → Google Sheets API access token"""
    import json as _json
    key_json = os.environ.get('SHEETS_SERVICE_ACCOUNT', '').strip()
    if not key_json:
        return ''
    try:
        from google.oauth2 import service_account
        import google.auth.transport.requests
        key_dict = _json.loads(key_json)
        creds = service_account.Credentials.from_service_account_info(
            key_dict,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        creds.refresh(google.auth.transport.requests.Request())
        return creds.token
    except ImportError as e:
        import traceback
        print(f"google-auth import 실패: {e}")
        traceback.print_exc()
        return ''
    except Exception as e:
        import traceback
        print(f"Sheets 인증 실패: {e}")
        traceback.print_exc()
        return ''


def _sheets_get(spreadsheet_id: str, token: str, range_str: str) -> list:
    """Sheets API values.get"""
    import urllib.request as _req
    import urllib.parse as _up
    import json as _json
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{_up.quote(range_str)}'
    req = _req.Request(url, headers={'Authorization': f'Bearer {token}'})
    with _req.urlopen(req) as resp:
        return _json.loads(resp.read()).get('values', [])


def _sheets_batch_update(spreadsheet_id: str, token: str, data: list):
    """Sheets API values:batchUpdate"""
    import urllib.request as _req
    import json as _json
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values:batchUpdate'
    body = _json.dumps({'valueInputOption': 'RAW', 'data': data}).encode()
    req = _req.Request(url, data=body, method='POST',
                       headers={'Authorization': f'Bearer {token}',
                                'Content-Type': 'application/json'})
    with _req.urlopen(req) as resp:
        return _json.loads(resp.read())


def _sheets_batch_request(spreadsheet_id: str, token: str, requests: list):
    """Sheets API spreadsheets.batchUpdate"""
    import urllib.request as _req
    import json as _json
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate'
    body = _json.dumps({'requests': requests}).encode()
    req = _req.Request(url, data=body, method='POST',
                       headers={'Authorization': f'Bearer {token}',
                                'Content-Type': 'application/json'})
    with _req.urlopen(req) as resp:
        return _json.loads(resp.read())


def _sheets_get_metadata(spreadsheet_id: str, token: str) -> dict:
    """Sheets API spreadsheets.get"""
    import urllib.request as _req
    import urllib.parse as _up
    import json as _json
    fields = 'sheets(properties(sheetId,title))'
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}?fields={_up.quote(fields)}'
    req = _req.Request(url, headers={'Authorization': f'Bearer {token}'})
    with _req.urlopen(req) as resp:
        return _json.loads(resp.read())


def get_sheet_titles(spreadsheet_id: str, token: str) -> list:
    """스프레드시트의 탭 제목 목록 반환"""
    metadata = _sheets_get_metadata(spreadsheet_id, token)
    return [
        sheet.get('properties', {}).get('title', '')
        for sheet in metadata.get('sheets', [])
        if sheet.get('properties', {}).get('title')
    ]


def normalize_sheet_key(value: str) -> str:
    """탭 제목 비교용 정규화 키"""
    return re.sub(r'[\s\W_]+', '', value or '').casefold()


def resolve_invoice_sheet_name(location: str, sheet_titles: list) -> str:
    """숙소 위치에 대응하는 실제 invoice 탭 제목 반환"""
    expected = INVOICE_SHEET_MAP.get(location, '')
    if not expected or not sheet_titles:
        return expected

    if expected in sheet_titles:
        return expected

    normalized_titles = {
        normalize_sheet_key(title): title
        for title in sheet_titles
    }
    normalized_expected = normalize_sheet_key(expected)
    if normalized_expected in normalized_titles:
        return normalized_titles[normalized_expected]

    aliases = [expected, location] + INVOICE_SHEET_ALIASES.get(location, [])
    for alias in aliases:
        alias_key = normalize_sheet_key(alias)
        if not alias_key:
            continue
        matches = [
            title
            for title in sheet_titles
            if normalize_sheet_key(title).find(alias_key) != -1
        ]
        if len(matches) == 1:
            return matches[0]

    return expected


def list_invoice_sheets() -> bool:
    """캐리_정산 스프레드시트의 invoice 탭 목록 출력"""
    token = get_sheets_token()
    if not token:
        print("Sheets 토큰 없음. invoice 시트 목록 조회 건너뜀.")
        return False

    try:
        sheet_titles = get_sheet_titles(CARRY_JUNGSAN_SPREADSHEET_ID, token)
        invoice_titles = [
            title for title in sheet_titles
            if title.startswith('invoice(거래명세서)_')
        ]
        print(f"invoice 시트 수: {len(invoice_titles)}개")
        for title in invoice_titles:
            print(f"  - {title}")
        return True
    except Exception as e:
        print(f"invoice 시트 목록 조회 실패: {e}")
        return False


def format_profit_sheet_row(token: str, target_row: int):
    """영업이익계산 행 서식을 기존 완성 행 기준으로 복사"""
    if target_row == PROFIT_FORMAT_TEMPLATE_ROW:
        return

    source_row_index = PROFIT_FORMAT_TEMPLATE_ROW - 1
    target_row_index = target_row - 1
    _sheets_batch_request(CARRY_KOEUN_SPREADSHEET_ID, token, [{
        'copyPaste': {
            'source': {
                'sheetId': CARRY_PROFIT_SHEET_ID,
                'startRowIndex': source_row_index,
                'endRowIndex': source_row_index + 1,
                'startColumnIndex': 0,
                'endColumnIndex': 18,
            },
            'destination': {
                'sheetId': CARRY_PROFIT_SHEET_ID,
                'startRowIndex': target_row_index,
                'endRowIndex': target_row_index + 1,
                'startColumnIndex': 0,
                'endColumnIndex': 18,
            },
            'pasteType': 'PASTE_FORMAT',
        }
    }])


def update_profit_sheet(year: int, month: int, total_amount: int) -> bool:
    """캐리_고객 '영업이익계산' 시트 업데이트"""
    token = get_sheets_token()
    if not token:
        print("Sheets 토큰 없음. 영업이익계산 시트 업데이트 건너뜀.")
        return False

    try:
        rows = _sheets_get(CARRY_KOEUN_SPREADSHEET_ID, token, '영업이익계산!A:A')
        month_str = f'{year}-{month}'

        target_row = None
        for i, row in enumerate(rows, 1):
            if row and str(row[0]).strip() == month_str:
                target_row = i
                break

        if target_row is None:
            target_row = len(rows) + 1

        FC = SHEETS_FIXED_COSTS
        profit_summary = calculate_profit_summary(total_amount)

        data = [
            {'range': f'영업이익계산!A{target_row}', 'values': [[month_str]]},
            {'range': f'영업이익계산!B{target_row}', 'values': [[total_amount]]},
            {'range': f'영업이익계산!D{target_row}', 'values': [[FC['hourly_wage']]]},
            {'range': f'영업이익계산!E{target_row}', 'values': [[profit_summary['labor_cost']]]},
            {'range': f'영업이익계산!F{target_row}', 'values': [[FC['logistics_count']]]},
            {'range': f'영업이익계산!G{target_row}', 'values': [[FC['logistics_cost_per']]]},
            {'range': f'영업이익계산!H{target_row}', 'values': [[profit_summary['logistics_cost']]]},
            {'range': f'영업이익계산!I{target_row}', 'values': [[FC['rent_utility']]]},
            {'range': f'영업이익계산!J{target_row}', 'values': [[FC['electricity']]]},
            {'range': f'영업이익계산!K{target_row}', 'values': [[FC['water']]]},
            {'range': f'영업이익계산!L{target_row}', 'values': [[FC['insurance']]]},
            {'range': f'영업이익계산!M{target_row}', 'values': [[profit_summary['supplies_cost']]]},
            {'range': f'영업이익계산!N{target_row}', 'values': [[profit_summary['withholding_tax']]]},
            {'range': f'영업이익계산!O{target_row}', 'values': [[profit_summary['vat']]]},
            {'range': f'영업이익계산!P{target_row}', 'values': [[profit_summary['total_cost']]]},
            {'range': f'영업이익계산!Q{target_row}', 'values': [[profit_summary['operating_profit']]]},
            {'range': f'영업이익계산!R{target_row}', 'values': [[profit_summary['operating_margin']]]},
        ]

        _sheets_batch_update(CARRY_KOEUN_SPREADSHEET_ID, token, data)
        format_profit_sheet_row(token, target_row)
        print(f"영업이익계산 시트 업데이트 완료: {month_str} ({target_row}행), 매출 {total_amount:,}원")
        return True
    except Exception as e:
        print(f"영업이익계산 시트 업데이트 실패: {e}")
        return False


def update_invoice_sheets(year: int, month: int, rows: list) -> bool:
    """캐리_정산 각 invoice 시트 수량/월 업데이트"""
    token = get_sheets_token()
    if not token:
        print("Sheets 토큰 없음. 거래명세서 시트 업데이트 건너뜀.")
        return False

    location_totals = {}
    for row in rows:
        record_date, location, blanket, mat, pillow_cover, towel, _, _, _ = row
        if not is_settlement_location_active(location, record_date):
            continue
        if location not in location_totals:
            location_totals[location] = {'blanket': 0, 'mat': 0, 'pillow_cover': 0, 'towel': 0}
        location_totals[location]['blanket'] += blanket or 0
        location_totals[location]['mat'] += mat or 0
        location_totals[location]['pillow_cover'] += pillow_cover or 0
        location_totals[location]['towel'] += towel or 0

    month_str = f'{year}년 {month}월'
    try:
        sheet_titles = get_sheet_titles(CARRY_JUNGSAN_SPREADSHEET_ID, token)
    except Exception as e:
        print(f"invoice 시트 목록 조회 실패. 기존 매핑으로 진행: {e}")
        sheet_titles = []

    failures = []
    for location, qty in location_totals.items():
        sheet_name = resolve_invoice_sheet_name(location, sheet_titles)
        if not sheet_name:
            print(f"  invoice 매핑 없음: {location}")
            failures.append(location)
            continue
        try:
            data = [
                {'range': f"'{sheet_name}'!B8",  'values': [[month_str]]},
                {'range': f"'{sheet_name}'!D12", 'values': [[qty['blanket']]]},
                {'range': f"'{sheet_name}'!D13", 'values': [[qty['mat']]]},
                {'range': f"'{sheet_name}'!D14", 'values': [[qty['pillow_cover']]]},
                {'range': f"'{sheet_name}'!D15", 'values': [[qty['towel']]]},
            ]
            _sheets_batch_update(CARRY_JUNGSAN_SPREADSHEET_ID, token, data)
            print(f"  invoice 시트 업데이트 완료: {sheet_name}")
        except Exception as e:
            print(f"  invoice 시트 업데이트 실패 ({sheet_name}): {e}")
            failures.append(sheet_name)

    if failures:
        print(f"invoice 시트 업데이트 실패 {len(failures)}건: {', '.join(failures)}")
        return False
    return True


# ============================================================
# 메인
# ============================================================

def get_invoice_job_mode() -> str:
    """정산 배치 실행 모드 반환"""
    job_mode = os.environ.get('INVOICE_JOB_MODE', 'full').strip() or 'full'
    if job_mode not in INVOICE_JOB_MODES:
        allowed = ', '.join(sorted(INVOICE_JOB_MODES))
        raise SystemExit(f"지원하지 않는 INVOICE_JOB_MODE: {job_mode} (허용: {allowed})")
    return job_mode


def main():
    # 대상 월 결정 (실행일이 말일이면 해당 월, 아니면 전월)
    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]

    if today.day == last_day:
        year, month = today.year, today.month
    else:
        # 테스트용: 전월 데이터
        if today.month == 1:
            year, month = today.year - 1, 12
        else:
            year, month = today.year, today.month - 1

    # 명령행 인수로 년월 지정 가능
    if len(sys.argv) >= 3:
        year = int(sys.argv[1])
        month = int(sys.argv[2])

    print(f"=== {year}년 {month}월 세탁물 정산 ===")
    print(f"정산 기간: {format_settlement_period(year, month)}")
    job_mode = get_invoice_job_mode()
    print(f"실행 모드: {job_mode}")

    if job_mode == 'list_invoice_sheets':
        if not list_invoice_sheets():
            sys.exit(1)
        return

    # 데이터 조회
    rows = get_monthly_data(year, month)
    if not rows:
        print("데이터 없음")
        return

    print(f"조회된 레코드: {len(rows)}건")

    # 사업자별 집계
    business_data = aggregate_by_business(rows)
    print(f"사업자 수: {len(business_data)}개")

    if job_mode == 'invoice_sheets_only':
        if not update_invoice_sheets(year, month, rows):
            sys.exit(1)
        return

    # PDF 생성
    attachments = []

    for reg_no, data in business_data.items():
        pdf_buffer = generate_pdf(reg_no, data, year, month)
        filename = f"{year}년 {month}월 거래명세서 ({data['name']}).pdf"
        attachments.append((filename, pdf_buffer))
        print(f"PDF 생성: {filename}")

    # Excel 생성
    excel_buffer = generate_excel(business_data, year, month)
    excel_filename = f"{year}년 {month}월 세금계산서 (홈택스).xlsx"
    attachments.append((excel_filename, excel_buffer))
    print(f"Excel 생성: {excel_filename}")

    # 합계 계산
    total_amount = calculate_total_amount_from_business_data(business_data)

    # 이메일 발송
    profit_summary = calculate_profit_summary(total_amount)
    monthly_close_summary = calculate_monthly_close_summary(
        total_amount,
        kops_receivable_amount=calculate_kops_receivable_amount(business_data),
    )
    subject = f"[캐리] {year}년 {month}월 거래명세서"
    body = f"""{year}년 {month}월 세탁물 정산 내역입니다.

정산 기간: {format_settlement_period(year, month)}
총 금액: {total_amount:,}원 (사업자 {len(business_data)}개)

{format_profit_summary_text(profit_summary, monthly_close_summary)}

{format_monthly_close_text(monthly_close_summary)}

첨부:
- 거래명세서 PDF {len(business_data)}개
- 세금계산서 Excel (홈택스 일괄발행용) 1개

감사합니다.
"""

    # 버퍼 위치 리셋
    for filename, buffer in attachments:
        buffer.seek(0)

    send_report_email(year, month, rows, business_data)
    send_email(subject, body, attachments)

    # Google Sheets 업데이트
    update_profit_sheet(year, month, total_amount)
    if not update_invoice_sheets(year, month, rows):
        sys.exit(1)

    # 로컬 저장 (iCloud Drive 월별 폴더)
    if os.environ.get('SAVE_LOCAL'):
        icloud_base = os.path.expanduser(
            "~/Library/Mobile Documents/com~apple~CloudDocs/정훈_정산"
        )
        output_dir = os.path.join(icloud_base, f"{year}년_{month}월")
        os.makedirs(output_dir, exist_ok=True)
        for filename, buffer in attachments:
            buffer.seek(0)
            with open(os.path.join(output_dir, filename), 'wb') as f:
                f.write(buffer.read())
        print(f"로컬 저장: {output_dir}")


if __name__ == '__main__':
    main()
