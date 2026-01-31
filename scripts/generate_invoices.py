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
from datetime import datetime, date
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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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
EMAIL_TO = "kham0126@gmail.com"
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
    '서대문구 연희로4길 25-7': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 고산자로 508-3': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 왕산로 200, 1004호': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 장한로26나길 21': ('767-87-02214', '주식회사 콥스', '남택호'),
    '광진구 능동로 165-1': ('767-87-02214', '주식회사 콥스', '남택호'),
    '송파구 가락로28길 3-10': ('767-87-02214', '주식회사 콥스', '남택호'),
    '동대문구 회기로 189': ('419-11-02853', '오를리(Orly)', '김지혜'),
    '관악구 신림동1길 19-5': ('461-86-03598', '주식회사스테이모먼트', '유경민'),
}

# 품목 단가
PRICES = {
    'blanket': 6500,       # 이불
    'mat': 4500,           # 요
    'pillow_cover': 1500,  # 베개커버
    'towel': 1000,         # 수건
    'body_towel': 700,     # 바디타월
}

ITEM_NAMES = {
    'blanket': '이불',
    'mat': '요',
    'pillow_cover': '베개커버',
    'towel': '수건',
    'body_towel': '바디타월',
}


# ============================================================
# 데이터 조회
# ============================================================

def get_monthly_data(year: int, month: int) -> list:
    """해당 월의 세탁물 데이터 조회"""
    conn = psycopg2.connect(SUPABASE_URI)
    cur = conn.cursor()

    cur.execute("""
        SELECT record_date, location, blanket, mat, pillow_cover, towel, body_towel
        FROM laundry_records
        WHERE EXTRACT(YEAR FROM record_date) = %s
          AND EXTRACT(MONTH FROM record_date) = %s
        ORDER BY location, record_date
    """, (year, month))

    rows = cur.fetchall()
    cur.close()
    conn.close()

    return rows


def aggregate_by_business(rows: list) -> dict:
    """사업자별로 데이터 집계"""
    business_data = {}

    for row in rows:
        record_date, location, blanket, mat, pillow_cover, towel, body_towel = row

        if location not in BUSINESS_MAP:
            print(f"알 수 없는 숙소: {location}")
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
                'towel': 0, 'body_towel': 0
            }

        loc_data = business_data[reg_no]['locations'][location]
        loc_data['blanket'] += blanket or 0
        loc_data['mat'] += mat or 0
        loc_data['pillow_cover'] += pillow_cover or 0
        loc_data['towel'] += towel or 0
        loc_data['body_towel'] += body_towel or 0

    return business_data


# ============================================================
# PDF 생성
# ============================================================

def register_font():
    """한글 폰트 등록"""
    font_paths = [
        '/usr/share/fonts/truetype/noto/NotoSansKR-Regular.ttf',  # Ubuntu
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',  # Ubuntu alt
        '/System/Library/Fonts/AppleSDGothicNeo.ttc',  # macOS
        'NotoSansKR-Regular.ttf',  # 로컬
    ]

    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('Korean', path))
                return True
            except:
                continue

    # GitHub Actions에서 설치된 폰트 경로
    noto_path = '/usr/share/fonts/truetype/noto-cjk/NotoSansCJK-Regular.ttc'
    if os.path.exists(noto_path):
        pdfmetrics.registerFont(TTFont('Korean', noto_path))
        return True

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

        for item_key, item_name in ITEM_NAMES.items():
            qty = loc_data.get(item_key, 0)
            if qty > 0:
                price = PRICES[item_key]
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
    supply_amount = int(grand_total / 1.1)
    tax_amount = grand_total - supply_amount

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
        total = 0
        for loc_data in data['locations'].values():
            for item_key in PRICES:
                total += (loc_data.get(item_key, 0) or 0) * PRICES[item_key]

        for item in data.get('extra_items', []):
            total += item['amount']

        supply_amount = int(total / 1.1)
        tax_amount = total - supply_amount

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
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    for filename, data in attachments:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(data.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("이메일 발송 완료")
        return True
    except Exception as e:
        print(f"이메일 발송 실패: {e}")
        return False


# ============================================================
# 메인
# ============================================================

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

    # 데이터 조회
    rows = get_monthly_data(year, month)
    if not rows:
        print("데이터 없음")
        return

    print(f"조회된 레코드: {len(rows)}건")

    # 사업자별 집계
    business_data = aggregate_by_business(rows)
    print(f"사업자 수: {len(business_data)}개")

    # PDF 생성
    attachments = []

    for reg_no, data in business_data.items():
        pdf_buffer = generate_pdf(reg_no, data, year, month)
        filename = f"{year}년_{month}월_거래명세서_{data['name'].replace(' ', '_')}.pdf"
        attachments.append((filename, pdf_buffer))
        print(f"PDF 생성: {filename}")

    # Excel 생성
    excel_buffer = generate_excel(business_data, year, month)
    excel_filename = f"홈택스_세금계산서_{year}년{month}월.xlsx"
    attachments.append((excel_filename, excel_buffer))
    print(f"Excel 생성: {excel_filename}")

    # 합계 계산
    total_amount = 0
    for data in business_data.values():
        for loc_data in data['locations'].values():
            for item_key in PRICES:
                total_amount += (loc_data.get(item_key, 0) or 0) * PRICES[item_key]
        for item in data.get('extra_items', []):
            total_amount += item['amount']

    # 이메일 발송
    subject = f"[캐리] {year}년 {month}월 세탁물 정산 - 세금계산서 발행 요청"
    body = f"""안녕하세요, {year}년 {month}월 세탁물 정산 내역입니다.

총 금액: {total_amount:,}원
사업자 수: {len(business_data)}개

첨부파일:
- 사업자별 거래명세서 PDF ({len(business_data)}개)
- 홈택스 세금계산서 Excel (1개)

홈택스에서 세금계산서 발행 후 회신 부탁드립니다.

감사합니다.
"""

    # 버퍼 위치 리셋
    for filename, buffer in attachments:
        buffer.seek(0)

    send_email(subject, body, attachments)

    # 로컬 저장 (테스트용)
    if os.environ.get('SAVE_LOCAL'):
        output_dir = f"/tmp/invoice_{year}_{month}"
        os.makedirs(output_dir, exist_ok=True)
        for filename, buffer in attachments:
            buffer.seek(0)
            with open(f"{output_dir}/{filename}", 'wb') as f:
                f.write(buffer.read())
        print(f"로컬 저장: {output_dir}")


if __name__ == '__main__':
    main()
