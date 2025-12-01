"""
Модуль експорту звіту у формат PDF.
Забезпечує підтримку кирилиці (DejaVuSans) та генерацію статичних графіків.
"""

import io
import os
import tempfile
import requests
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List

# Надійне посилання на шрифт DejaVuSans (стандарт для кирилиці)
FONT_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"
FONT_NAME = "DejaVuSans"

class PDFReport(FPDF):
    def __init__(self, font_path):
        super().__init__()
        self.font_path = font_path
        
        # Реєстрація шрифту
        if font_path:
            self.add_font(FONT_NAME, "", font_path, uni=True)
            self.set_font(FONT_NAME, size=12)
        else:
            self.set_font("Arial", size=12)

    def header(self):
        # Використовуємо registered font name або стандартний
        font = FONT_NAME if self.font_path else "Arial"
        self.set_font(font, "", 10)
        self.cell(0, 10, "Звіт за результатами опитування", border=False, align="R")
        self.ln(10)

    def footer(self):
        font = FONT_NAME if self.font_path else "Arial"
        self.set_y(-15)
        self.set_font(font, "", 8)
        self.cell(0, 10, f"Сторінка {self.page_no()}", align="C")

    def chapter_title(self, text):
        font = FONT_NAME if self.font_path else "Arial"
        self.set_font(font, "", 12)
        self.set_fill_color(220, 230, 241) 
        # text=str(text) на випадок якщо там не рядок
        self.multi_cell(0, 10, str(text), fill=True, align='L')
        self.ln(2)

    def add_table(self, df: pd.DataFrame):
        font = FONT_NAME if self.font_path else "Arial"
        self.set_font(font, "", 10)
        
        line_height = self.font_size * 2
        col_width = [110, 30, 20] 

        headers = df.columns.tolist() 
        self.set_fill_color(240, 240, 240)
        
        # Заголовок таблиці
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

        # Рядки таблиці
        for row in df.itertuples(index=False):
            x_start = self.get_x()
            y_start = self.get_y()
            
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])

            # 1. Текстова комірка (може бути багаторядковою)
            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            # Визначаємо, де ми опинилися після multi_cell
            x_next = self.get_x()
            y_next = self.get_y()
            h_curr = y_next - y_start
            
            # Повертаємо курсор наверх для малювання сусідніх комірок
            self.set_xy(x_start + col_width[0], y_start)
            
            # 2. Числові комірки (висота така сама, як у текстової)
            self.cell(col_width[1], h_curr, count_val, border=1, align='C')
            self.cell(col_width[2], h_curr, perc_val, border=1, align='C')
            
            # Перехід на новий рядок (на позицію під найвищою коміркою)
            self.set_xy(x_start, y_next)
            # self.ln() не потрібен, бо ми вже на правильному Y

    def add_chart(self, qs: QuestionSummary):
        if qs.table.empty:
            return

        # Створення графіку
        plt.figure(figsize=(6, 3))
        labels = qs.table["Варіант відповіді"]
        values = qs.table["Кількість"]

        if qs.question.qtype == QuestionType.SCALE:
            bars = plt.bar(labels, values, color='#4F81BD')
            plt.ylabel('Кількість')
            # Підписи значень
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height,
                         f'{int(height)}', ha='center', va='bottom')
            plt.xticks(rotation=0)
        else:
            plt.pie(values, labels=None, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
            plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5), fontsize='small')
            plt.axis('equal') 

        plt.tight_layout()

        # Збереження у тимчасовий файл
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            plt.savefig(tmp_img.name, format='png', dpi=100)
            tmp_img_path = tmp_img.name
        
        plt.close()

        # Перевірка місця на сторінці
        if self.get_y() > 200:
            self.add_page()
            
        self.image(tmp_img_path, w=150)
        self.ln(5)
        
        # Видалення файлу
        try:
            os.remove(tmp_img_path)
        except:
            pass

def ensure_font_exists():
    """Завантажує шрифт, якщо його немає або файл пошкоджено."""
    font_path = "DejaVuSans.ttf"
    
    # Видаляємо, якщо файл надто малий (помилка завантаження)
    if os.path.exists(font_path):
        if os.path.getsize(font_path) < 10000:
            try:
                os.remove(font_path)
            except:
                pass

    if not os.path.exists(font_path):
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(FONT_URL, headers=headers)
            response.raise_for_status()
            with open(font_path, "wb") as f:
                f.write(response.content)
        except Exception as e:
            print(f"Font download error: {e}")
            return None
            
    return font_path

def build_pdf_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    font_path = ensure_font_exists()
    
    # Ініціалізація PDF
    pdf = PDFReport(font_path)
    pdf.add_page()

    # --- Титульна сторінка ---
    font_name = FONT_NAME if font_path else "Arial"
    pdf.set_font(font_name, "", 16)
    
    # Використовуємо .encode('latin-1', 'replace') для заголовків тільки якщо немає шрифту,
    # але оскільки ми використовуємо unicode=True в fpdf2, encode не потрібен для тексту методу cell/multi_cell.
    
    pdf.cell(0, 10, "Звіт про результати опитування", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font(font_name, "", 12)
    pdf.cell(0, 10, f"Всього анкет: {len(original_df)}", ln=True)
    pdf.cell(0, 10, f"Оброблено анкет: {len(sliced_df)}", ln=True)
    pdf.cell(0, 10, f"Діапазон: {range_info}", ln=True)
    pdf.ln(10)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # --- Основний контент ---
    for qs in summaries:
        if pdf.get_y() > 250:
            pdf.add_page()

        title = f"{qs.question.code}. {qs.question.text}"
        pdf.chapter_title(title)

        if qs.table.empty:
            pdf.set_font(font_name, "", 10)
            pdf.cell(0, 10, "Немає даних або відкриті відповіді.", ln=True)
            pdf.ln(5)
            continue

        try:
            pdf.add_table(qs.table)
        except Exception as e:
            pdf.cell(0, 10, f"Table error: {str(e)}", ln=True)

        pdf.ln(5)
        
        # Якщо мало місця для графіка - нова сторінка
        if pdf.get_y() > 180:
            pdf.add_page()
            
        try:
            pdf.add_chart(qs)
        except Exception:
            pass

        pdf.ln(10)

    # ВАЖЛИВО: fpdf2.output() повертає bytearray. Конвертуємо в bytes.
    # .encode() не потрібен.
    return bytes(pdf.output())