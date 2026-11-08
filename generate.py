#!/usr/bin/env python3
"""Generate a Chinese career-development PDF report and 36-slide PPTX."""
from pathlib import Path
import re
import html
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, KeepTogether

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'output'
REPORT = ROOT / 'career_report.md'
SLIDES = ROOT / 'presentation.md'

FONT_CANDIDATES = [
    os.environ.get('CJK_FONT', ''),
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    '/System/Library/Fonts/PingFang.ttc',
    'C:/Windows/Fonts/msyh.ttc',
    'C:/Windows/Fonts/simhei.ttf',
]


def register_pdf_font():
    for candidate in FONT_CANDIDATES:
        if candidate and Path(candidate).is_file():
            try:
                pdfmetrics.registerFont(TTFont('Chinese', candidate, subfontIndex=0))
                return 'Chinese'
            except Exception:
                continue
    raise SystemExit('未找到可用中文字体。请安装 Noto Sans CJK，或设置环境变量 CJK_FONT 指向中文字体文件。')


def md_inline(text):
    text = html.escape(text.strip())
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'`(.+?)`', r'<font name="Courier">\1</font>', text)
    return text


def make_pdf(font_name):
    styles = getSampleStyleSheet()
    title = ParagraphStyle('TitleCN', parent=styles['Title'], fontName=font_name, fontSize=20,
                           leading=28, textColor=colors.HexColor('#16324F'), alignment=TA_CENTER,
                           spaceAfter=18)
    h1 = ParagraphStyle('H1CN', parent=styles['Heading1'], fontName=font_name, fontSize=15,
                        leading=22, textColor=colors.HexColor('#16324F'), spaceBefore=13, spaceAfter=7,
                        keepWithNext=True)
    h2 = ParagraphStyle('H2CN', parent=styles['Heading2'], fontName=font_name, fontSize=12,
                        leading=18, textColor=colors.HexColor('#177E89'), spaceBefore=9, spaceAfter=5,
                        keepWithNext=True)
    body = ParagraphStyle('BodyCN', parent=styles['BodyText'], fontName=font_name, fontSize=9.5,
                          leading=16, textColor=colors.HexColor('#263238'), spaceAfter=6,
                          wordWrap='CJK', alignment=TA_LEFT)
    bullet = ParagraphStyle('BulletCN', parent=body, leftIndent=14, firstLineIndent=-9, spaceAfter=4)
    story = []
    first_h1 = True
    for raw in REPORT.read_text(encoding='utf-8').splitlines():
        line = raw.strip()
        if not line or line == '---':
            continue
        if line.startswith('# '):
            story.append(Paragraph(md_inline(line[2:]), title))
        elif line.startswith('## '):
            if not first_h1:
                story.append(Spacer(1, 4))
            first_h1 = False
            story.append(Paragraph(md_inline(line[3:]), h1))
        elif line.startswith('### '):
            story.append(Paragraph(md_inline(line[4:]), h2))
        elif line.startswith('- '):
            story.append(Paragraph('• ' + md_inline(line[2:]), bullet))
        elif re.match(r'^\d+\.\s', line):
            story.append(Paragraph(md_inline(line), bullet))
        else:
            story.append(Paragraph(md_inline(line), body))
    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font_name, 8)
        canvas.setFillColor(colors.HexColor('#687784'))
        canvas.drawString(20*mm, 12*mm, '基层田径教练员生涯发展报告')
        canvas.drawRightString(A4[0]-20*mm, 12*mm, f'{doc.page}')
        canvas.restoreState()
    path = OUT / '基层田径教练员生涯发展报告.pdf'
    doc = SimpleDocTemplate(str(path), pagesize=A4, rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=18*mm, bottomMargin=20*mm, title='基层田径教练员生涯发展报告',
                            author='')
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return path


def parse_slides():
    slides = []
    current = None
    for raw in SLIDES.read_text(encoding='utf-8').splitlines():
        line = raw.strip()
        if line.startswith('## '):
            if current:
                slides.append(current)
            title = line[3:]
            current = {'title': re.sub(r'^\d+\s*\|\s*', '', title), 'bullets': []}
        elif line.startswith('- ') and current:
            current['bullets'].append(line[2:])
    if current:
        slides.append(current)
    if len(slides) != 36:
        raise ValueError(f'展示文稿应为36页，当前解析到 {len(slides)} 页')
    return slides


def make_pptx():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    navy = RGBColor(22, 50, 79)
    teal = RGBColor(23, 126, 137)
    gold = RGBColor(232, 173, 64)
    ink = RGBColor(38, 50, 56)
    muted = RGBColor(95, 111, 121)
    for index, item in enumerate(parse_slides(), start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background.fill
        bg.solid(); bg.fore_color.rgb = RGBColor(248, 250, 252)
        # top accent and section marker
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(.16))
        bar.fill.solid(); bar.fill.fore_color.rgb = teal; bar.line.fill.background()
        if index == 1:
            panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(.16), prs.slide_width, prs.slide_height-Inches(.16))
            panel.fill.solid(); panel.fill.fore_color.rgb = navy; panel.line.fill.background()
            panel.name = 'Cover background'
            # Decorative track-like lines
            for j in range(3):
                shape = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(8.7+j*.45), Inches(1.0+j*.35), Inches(4.0), Inches(5.2))
                shape.line.color.rgb = RGBColor(48, 82, 110); shape.line.width = Pt(2+j*.4)
                shape.fill.background()
            box = slide.shapes.add_textbox(Inches(.9), Inches(1.55), Inches(10.8), Inches(2.5))
            tf = box.text_frame; tf.clear(); tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = item['title']; p.font.name='Microsoft YaHei'; p.font.size=Pt(32); p.font.bold=True; p.font.color.rgb=RGBColor(255,255,255)
            p2 = tf.add_paragraph(); p2.text = '从入门助教到专业引领'; p2.font.name='Microsoft YaHei'; p2.font.size=Pt(23); p2.font.color.rgb=gold; p2.space_before=Pt(16)
            p3 = tf.add_paragraph(); p3.text = '以能力成长为主线｜以运动员长期发展为中心'; p3.font.name='Microsoft YaHei'; p3.font.size=Pt(16); p3.font.color.rgb=RGBColor(220,230,238); p3.space_before=Pt(12)
            if item['bullets']:
                small=slide.shapes.add_textbox(Inches(.95), Inches(5.75), Inches(8), Inches(.8))
                pp=small.text_frame.paragraphs[0]; pp.text='汇报人：____    单位：____    日期：____'; pp.font.name='Microsoft YaHei'; pp.font.size=Pt(14); pp.font.color.rgb=RGBColor(230,235,240)
        else:
            titlebox = slide.shapes.add_textbox(Inches(.72), Inches(.42), Inches(11.9), Inches(.8))
            p = titlebox.text_frame.paragraphs[0]; p.text = item['title']; p.font.name='Microsoft YaHei'; p.font.size=Pt(25); p.font.bold=True; p.font.color.rgb=navy
            # Gold accent beneath heading
            accent=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.74), Inches(1.28), Inches(.9), Inches(.055))
            accent.fill.solid(); accent.fill.fore_color.rgb=gold; accent.line.fill.background()
            # Content card
            card=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(.72), Inches(1.65), Inches(11.9), Inches(4.95))
            card.fill.solid(); card.fill.fore_color.rgb=RGBColor(255,255,255); card.line.color.rgb=RGBColor(226,233,238)
            tf=slide.shapes.add_textbox(Inches(1.15), Inches(1.95), Inches(11.05), Inches(4.35)).text_frame
            tf.clear(); tf.word_wrap=True; tf.margin_left=Inches(.08); tf.margin_right=Inches(.08); tf.vertical_anchor=MSO_ANCHOR.MIDDLE
            for bi, text in enumerate(item['bullets']):
                p = tf.paragraphs[0] if bi == 0 else tf.add_paragraph()
                p.text = '•  ' + text
                p.font.name='Microsoft YaHei'; p.font.size=Pt(21 if len(item['bullets']) <= 3 else 18); p.font.color.rgb=ink
                p.space_after=Pt(20 if len(item['bullets']) <= 3 else 13); p.level=0
            # Footer
            foot=slide.shapes.add_textbox(Inches(.75), Inches(6.88), Inches(11.8), Inches(.3))
            fp=foot.text_frame.paragraphs[0]; fp.text='基层田径教练员生涯发展  ·  能力成长与实践证据'; fp.font.name='Microsoft YaHei'; fp.font.size=Pt(9); fp.font.color.rgb=muted
            num=slide.shapes.add_textbox(Inches(12.1), Inches(6.83), Inches(.45), Inches(.35))
            np=num.text_frame.paragraphs[0]; np.text=f'{index:02d}'; np.alignment=PP_ALIGN.RIGHT; np.font.name='Microsoft YaHei'; np.font.size=Pt(11); np.font.color.rgb=teal
    path = OUT / '基层田径教练员生涯发展展示.pptx'
    prs.save(path)
    return path


def main():
    OUT.mkdir(exist_ok=True)
    font = register_pdf_font()
    pdf = make_pdf(font)
    pptx = make_pptx()
    print(f'已生成：\n- {pdf}\n- {pptx}\nPPT页数：{len(parse_slides())}')

if __name__ == '__main__':
    main()
