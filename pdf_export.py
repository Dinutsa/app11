"""
Модуль експорту звіту у формат PDF.
Забезпечує:
- Підтримку кирилиці (DejaVuSans).
- Нерозривність таблиць (перенос на нову сторінку, якщо не влазять).
- УНІФІКОВАНИЙ дизайн діаграм (однаковий розмір, великі шрифти).
"""

import io
import os
import math
import tempfile
import requests
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# --- КОНСТАНТИ ---
FONT_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"
FONT_NAME = "DejaVuSans"

# Уніфіковані налаштування для графіків
CHART_FIGSIZE = (11, 6)   # Однаковий розмір полотна для всіх типів
CHART_DPI = 150           # Якість зображення
FONT_SIZE_BASE = 12       # Базовий шрифт для осей і легенди
FONT_SIZE_TITLE = 14      # Шрифт (хоча заголовок ми пишемо в PDF, але для осей корисно)
BAR_WIDTH = 0.5           # Ширина стовпчиків (щоб не були товстими)

def get_font_path() -> Optional[str]:
    """Шукає або завантажує шрифт DejaVuSans.ttf."""
    local_path = "DejaVuSans.ttf"
    
    # 1. Локальна перевірка
    if os.path.exists(local_path) and os.path.getsize(local_path) > 10000:
        return local_path

    # 2. Системна перевірка
    system_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf"
    ]
    for path in system_paths:
        if os.path.exists(path):
            return path

    # 3. Завантаження
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(FONT_URL, headers=headers, timeout=10)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            f.write(response.content)
        return local_path
    except Exception as e:
        print(f"Font download failed: {e}")
        return None

class PDFReport(FPDF):
    def __init__(self, font_path):
        super().__init__()
        self.font_path = font_path
        if not font_path:
            raise RuntimeError("Font not found")
        self.add_font(FONT_NAME, "", font_path, uni=True)
        self.set_font(FONT_NAME, size=12)

    def header(self):
        self.set_font(FONT_NAME, "", 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, "Звіт за результатами опитування", border=False, align="R")
        self.ln(10)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-15)
        self.set_font(FONT_NAME, "", 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"Сторінка {self.page_no()}", align="C")
        self.set_text_color(0, 0, 0)

    def chapter_title(self, text):
        # Якщо заголовок внизу сторінки, переносимо
        if self.get_y() > 250:
            self.add_page()
            
        self.set_font(FONT_NAME, "", 12)
        self.set_fill_color(220, 230, 241) 
        self.multi_cell(0, 8, str(text), fill=True, align='L')
        self.ln(2)

    def calculate_row_height(self, text_val, col_width, line_height):
        """Розраховує висоту рядка для тексту."""
        text_len = len(str(text_val))
        # Приблизно 50 символів в рядок шириною 110мм
        lines_count = math.ceil(text_len / 50)
        if lines_count < 1: lines_count = 1
        newlines = str(text_val).count('\n')
        lines_count += newlines
        return lines_count * line_height

    def add_table(self, df: pd.DataFrame):
        self.set_font(FONT_NAME, "", 10)
        line_height = 6
        col_width = [110, 30, 20] 
        headers = df.columns.tolist()

        # 1. Розрахунок повної висоти таблиці
        total_table_height = line_height 
        row_heights = []
        for row in df.itertuples(index=False):
            text_val = str(row[0])
            h = self.calculate_row_height(text_val, col_width[0], line_height)
            row_heights.append(h)
            total_table_height += h

        # 2. Логіка "Keep Together"
        page_break_trigger = 270
        space_left = page_break_trigger - self.get_y()

        # Якщо таблиця середня (до 240мм) і не влазить -> нова сторінка
        if total_table_height < 240 and total_table_height > space_left:
            self.add_page()

        # 3. Друк шапки
        self.set_fill_color(240, 240, 240)
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

        # 4. Друк рядків
        for idx, row in enumerate(df.itertuples(index=False)):
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])
            
            curr_h = row_heights[idx]

            # Аварійний розрив для дуже довгих таблиць
            if self.get_y() + curr_h > page_break_trigger:
                self.add_page()
                # Повтор шапки
                for i, h in enumerate(headers):
                    w = col_width[i] if i < len(col_width) else 20
                    self.cell(w, line_height, str(h), border=1, fill=True, align='C')
                self.ln(line_height)

            x_start = self.get_x()
            y_start = self.get_y()

            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            x_next = self.get_x()
            y_next = self.get_y()
            h_real = y_next - y_start 
            final_h = max(h_real, curr_h)

            self.set_xy(x_start + col_width[0], y_start)
            self.cell(col_width[1], final_h, count_val, border=1, align='C')
            self.cell(col_width[2], final_h, perc_val, border=1, align='C')
            
            self.set_xy(x_start, y_start + final_h)

    def add_chart(self, qs: QuestionSummary):
        if qs.table.empty:
            return

        # Завжди перевіряємо місце. Картинка займає ~90-100 мм висоти.
        if self.get_y() > 180:
            self.add_page()

        # --- НАЛАШТУВАННЯ MATPLOTLIB ---
        # Скидаємо налаштування і задаємо глобальний шрифт для цього малюнка
        plt.rcParams.update({'font.size': FONT_SIZE_BASE}) 
        
        # Фіксований розмір для ВСІХ типів
        plt.figure(figsize=CHART_FIGSIZE) 
        
        labels = qs.table["Варіант відповіді"]
        values = qs.table["Кількість"]

        if qs.question.qtype == QuestionType.SCALE:
            # --- СТОВПЧИКОВА ---
            # width=BAR_WIDTH робить їх тоншими
            bars = plt.bar(labels, values, color='#4F81BD', width=BAR_WIDTH)
            
            plt.ylabel('Кількість', fontsize=FONT_SIZE_BASE)
            plt.grid(axis='y', linestyle='--', alpha=0.5)
            
            # Підписи осей
            plt.xticks(fontsize=FONT_SIZE_BASE)
            plt.yticks(fontsize=FONT_SIZE_BASE)
            
            # Числа над стовпчиками
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                         f'{int(height)}', ha='center', va='bottom', 
                         fontsize=FONT_SIZE_BASE, fontweight='bold')
        
        else:
            # --- КРУГОВА ---
            colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
            c_arg = colors[:len(values)] if len(values) <= len(colors) else None
            
            # radius=1.1 – трохи більше стандарту (1.0)
            wedges, texts, autotexts = plt.pie(
                values, labels=None, autopct='%1.1f%%', startangle=90, 
                pctdistance=0.8, colors=c_arg, radius=1.1,
                textprops={'fontsize': FONT_SIZE_BASE}
            )
            
            # Жирні відсотки
            for autotext in autotexts:
                autotext.set_color('white') # Білий текст на кольоровому фоні краще
                autotext.set_weight('bold')
                autotext.set_fontsize(FONT_SIZE_BASE + 1)
                # Додаємо обводку (stroke), щоб читалось на світлому фоні
                import matplotlib.patheffects as path_effects
                autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='black')])

            plt.axis('equal')
            
            # Легенда ЗАВЖДИ праворуч, однаковим шрифтом
            plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5), 
                       fontsize=FONT_SIZE_BASE)

        # Важливо: tight_layout з pad=2 дає достатньо повітря навколо
        plt.tight_layout(pad=2.0)

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            # Зберігаємо з білими полями (bbox_inches='tight' іноді ріже легенду, якщо вона велика)
            # Тому ми покладаємось на tight_layout і фіксований розмір
            plt.savefig(tmp_img.name, format='png', dpi=CHART_DPI, bbox_inches='tight')
            tmp_img_path = tmp_img.name
        
        plt.close()

        # Вставка в PDF з фіксованою шириною 170 мм
        # Це гарантує, що візуально всі блоки будуть однакової ширини на сторінці
        self.image(tmp_img_path, x=20, w=170)
        self.ln(5)
        
        try:
            os.remove(tmp_img_path)
        except:
            pass

def build_pdf_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    font_path = get_font_path()
    
    if not font_path:
        err_pdf = FPDF()
        err_pdf.add_page()
        err_pdf.set_font("Helvetica", size=12)
        err_pdf.multi_cell(0, 10, "CRITICAL ERROR: Cyrillic font not found.")
        return bytes(err_pdf.output())

    pdf = PDFReport(font_path)
    pdf.add_page()

    # --- Титульна ---
    pdf.set_font(FONT_NAME, "", 16)
    pdf.cell(0, 10, "Звіт про результати опитування", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font(FONT_NAME, "", 12)
    pdf.cell(0, 8, f"Всього анкет: {len(original_df)}", ln=True)
    pdf.cell(0, 8, f"Оброблено анкет: {len(sliced_df)}", ln=True)
    pdf.cell(0, 8, f"Діапазон: {range_info}", ln=True)
    pdf.ln(10)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # --- Основний цикл ---
    for qs in summaries:
        title = f"{qs.question.code}. {qs.question.text}"
        pdf.chapter_title(title)

        if qs.table.empty:
            pdf.set_font(FONT_NAME, "", 10)
            pdf.cell(0, 10, "Немає даних або відкриті відповіді.", ln=True)
            pdf.ln(5)
            continue

        try:
            pdf.add_table(qs.table)
        except Exception as e:
            pdf.cell(0, 10, f"Table Error: {e}", ln=True)

        pdf.ln(5)
        
        try:
            pdf.add_chart(qs)
        except Exception:
            pass

        pdf.ln(5)

    return bytes(pdf.output())