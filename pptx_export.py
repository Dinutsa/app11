"""
Модуль експорту звіту у формат PowerPoint (.pptx).
Версія: CLEAN & STABLE.
- Без фонових зображень.
- Без шаблонів.
- З примусовими чорними рамками для таблиць (через XML).
"""

import io
import textwrap
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# --- ІМПОРТИ ДЛЯ XML (Щоб намалювати рамки) ---
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List

# --- НАЛАШТУВАННЯ ---
CHART_DPI = 150
FONT_SIZE_CHART = 12
FONT_SIZE_TABLE_HEADER = 12
FONT_SIZE_TABLE_DATA = 11
BAR_WIDTH = 0.6

# --- ФУНКЦІЇ ДЛЯ РАМОК ---
def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element

def set_cell_border(cell, border_color="000000", border_width='12700'):
    """
    Малює рамки навколо клітинки. 
    12700 EMU = 1pt.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Лінії: зліва, справа, зверху, знизу
    lines = [('a:lnL', border_width), ('a:lnR', border_width), 
             ('a:lnT', border_width), ('a:lnB', border_width)]
    
    for line_tag, w in lines:
        tcPr.remove(tcPr.find(qn(line_tag))) if tcPr.find(qn(line_tag)) is not None else None
        
        ln = SubElement(tcPr, line_tag, w=w, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
        prstDash = SubElement(ln, 'a:prstDash', val='solid')
        SubElement(ln, 'a:round')
        SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')

# --- ДІАГРАМИ ---
def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.clf()
    plt.rcParams.update({'font.size': FONT_SIZE_CHART})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    if qs.question.qtype == QuestionType.SCALE:
        fig = plt.figure(figsize=(6.0, 4.5))
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        plt.xticks(rotation=0)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
        fig = plt.figure(figsize=(6.0, 5.0))
        colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
        c_arg = colors[:len(values)] if len(values) <= len(colors) else None
        
        wedges, texts, autotexts = plt.pie(
            values, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.8, colors=c_arg, radius=1.1,
            textprops={'fontsize': FONT_SIZE_CHART}
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_weight('bold')
            import matplotlib.patheffects as path_effects
            autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])

        plt.axis('equal')
        cols = 2 if len(labels) > 2 else 1
        plt.legend(wrapped_labels, loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=cols, frameon=False, fontsize=10)

    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=CHART_DPI, bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

def build_pptx_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    prs = Presentation() # Чиста презентація (біла)

    # 1. Титульний
    slide_layout = prs.slide_layouts[0] 
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = "Звіт про результати опитування"
        slide.placeholders[1].text = f"Всього анкет: {len(original_df)}\nОброблено: {len(sliced_df)}\n{range_info}"
    except: pass

    # 2. Технічний
    slide_layout = prs.slide_layouts[1] 
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = "Технічна інформація"
        tf = slide.placeholders[1].text_frame
        tf.text = "Параметри вибірки:"
        for txt in [f"Загальна кількість: {len(original_df)}", f"У звіті: {len(sliced_df)}", f"Діапазон: {range_info}"]:
            p = tf.add_paragraph()
            p.text = txt
            p.font.size = Pt(20)
    except: pass

    # 3. Питання
    layout_index = 5 
    if len(prs.slide_layouts) <= 5: layout_index = len(prs.slide_layouts) - 1
    
    for qs in summaries:
        if qs.table.empty: continue
        
        slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
        
        # Заголовок
        try:
            title = slide.shapes.title
            title.text = f"{qs.question.code}. {qs.question.text}"
            if len(title.text) > 60:
                title.text_frame.paragraphs[0].font.size = Pt(24)
        except: pass

        # Таблиця
        rows = len(qs.table) + 1
        cols = 3
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(2.0), Inches(4.5), Inches(0.8)).table

        # Ширина
        table.columns[0].width = Inches(2.5)
        table.columns[1].width = Inches(1.0)
        table.columns[2].width = Inches(1.0)

        # Хедер
        headers = ["Варіант", "Кільк.", "%"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_TABLE_HEADER)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(230, 230, 230)
            set_cell_border(cell, border_width='12700') # 1pt

        # Дані
        for i, row in enumerate(qs.table.itertuples(index=False)):
            for j, val in enumerate(row):
                cell = table.cell(i+1, j)
                cell.text = str(val)
                cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_TABLE_DATA)
                if j > 0: cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                set_cell_border(cell, border_width='12700')

        # Діаграма
        try:
            img_stream = create_chart_image(qs)
            slide.shapes.add_picture(img_stream, Inches(5.2), Inches(2.0), width=Inches(4.6))
        except: pass

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()