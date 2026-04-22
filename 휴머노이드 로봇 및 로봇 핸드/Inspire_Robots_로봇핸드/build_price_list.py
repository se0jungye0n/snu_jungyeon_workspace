# -*- coding: utf-8 -*-
"""
Inspire Robots 이메일 회신 기반 참고 가격표 PDF 생성
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER


# 한글 폰트 등록
FONT_REG = 'C:/Windows/Fonts/malgun.ttf'
FONT_BOLD = 'C:/Windows/Fonts/malgunbd.ttf'
pdfmetrics.registerFont(TTFont('Malgun', FONT_REG))
pdfmetrics.registerFont(TTFont('MalgunBold', FONT_BOLD))


def build():
    out = os.path.join(os.path.dirname(__file__),
                       'Inspire_Robots_참고가격_260421.pdf')
    doc = SimpleDocTemplate(
        out, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle('title', parent=styles['Title'],
                           fontName='MalgunBold', fontSize=16,
                           alignment=TA_CENTER, spaceAfter=6)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'],
                        fontName='MalgunBold', fontSize=12,
                        spaceBefore=12, spaceAfter=4)
    body = ParagraphStyle('body', parent=styles['BodyText'],
                          fontName='Malgun', fontSize=10,
                          leading=15, alignment=TA_LEFT)
    note = ParagraphStyle('note', parent=body, fontSize=9,
                          textColor=colors.HexColor('#555555'))

    elements = []

    # 제목
    elements.append(Paragraph('Inspire Robots 로봇핸드 참고 가격', title))
    elements.append(Paragraph(
        '(Inspire Robots 이메일 회신 기반 · 정식 견적서 아님)', note))
    elements.append(Spacer(1, 6))

    # 수신 정보
    info_data = [
        ['수신 일자', '2026-04-21'],
        ['회신처', 'Inspire Robots'],
        ['수신자', '서정연 (서울대학교 AI 공학연구원 로봇 기술혁신센터)'],
        ['문서 구분', '참고 가격 안내 (Quick Reference Price List)'],
    ]
    info_tbl = Table(info_data, colWidths=[3.5*cm, 13*cm])
    info_tbl.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Malgun', 10),
        ('FONT', (0, 0), (0, -1), 'MalgunBold', 10),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F2F2F2')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#B0B0B0')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(info_tbl)

    # 가격표
    elements.append(Paragraph('1. 모델별 단가', h2))

    price_header = ['모델', '촉각 센서', '단가 (USD, 개당)', '비고']
    price_rows = [
        ['RH56BFX / DFX', '없음 (Standard)', '$ 2,608', '기본형'],
        ['RH56F1',       '포함 (with tactile)', '$ 4,056', '촉각 센서 탑재'],
        ['RH56E2',       '포함 (with tactile)', '$ 5,795', '촉각 센서 탑재 / 상위 모델'],
    ]
    data = [price_header] + price_rows
    tbl = Table(data, colWidths=[4.5*cm, 4.2*cm, 3.8*cm, 4.0*cm])
    tbl.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, 0), 'MalgunBold', 10),
        ('FONT', (0, 1), (-1, -1), 'Malgun', 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#305496')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#B0B0B0')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
         [colors.white, colors.HexColor('#F7F9FC')]),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(tbl)

    elements.append(Spacer(1, 6))
    elements.append(Paragraph('※ 상기 단가는 <b>배송비 제외</b> 기준입니다.', note))

    # 리드타임
    elements.append(Paragraph('2. 리드타임 (Lead Time)', h2))
    elements.append(Paragraph(
        '5 ~ 21일 (모델에 따라 상이)', body))

    # 후속 안내
    elements.append(Paragraph('3. 후속 안내', h2))
    elements.append(Paragraph(
        '세부 사양 및 활용 용도가 확정되는 대로 Inspire Robots 측에 재문의하여 '
        '정식 견적서(shipping 포함, 수량·옵션별 단가)를 수령할 예정.', body))

    # 원문 발췌
    elements.append(Paragraph('4. 원문 메일 발췌', h2))
    raw = (
        '<i>Thank you for your message.<br/><br/>'
        'Please find below a quick reference price list excluding shipping, '
        'for your initial evaluation:<br/><br/>'
        '&nbsp;&nbsp;• RH56BFX/DFX (standard, no tactile): USD 2,608/PC<br/>'
        '&nbsp;&nbsp;• RH56F1 (with tactile): USD 4,056/PC<br/>'
        '&nbsp;&nbsp;• RH56E2 (with tactile): USD 5,795/PC<br/><br/>'
        'Lead time: 5-21 days depending on model.<br/><br/>'
        'When you have more clarity on your application or specific '
        'requirements, feel free to reach out.</i>'
    )
    elements.append(Paragraph(raw, note))

    doc.build(elements)
    print(f'생성 완료: {out}')


if __name__ == '__main__':
    build()
