"""
Модуль експорту звіту у формат PDF.
ВЕРСІЯ: DATE IN HEADER (Дата в хедері).
- Хедер: Поточна дата справа зверху.
- Тіло: Назва звіту по центру.
- Виправлено всі імпорти.
"""

import io
import os
import urllib.request
import textwrap
import tempfile
from datetime import datetime  # <-- ВАЖЛИВО: Імпорт дати
import pandas as pd            # <-- ВАЖЛИВО: Імпорт pandas
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from fpdf import FPDF

from classification import QuestionType
from summary import QuestionSummary

# --- НАЛАШТУВАННЯ ---
CHART_DPI = 150
BAR_WIDTH = 0.6

# Налаштування шрифту (Tinos - аналог Times New Roman)
FONT_FILENAME = "Tinos-Regular.ttf"
FONT_PATH = os.path.join(os.getcwd(), FONT_FILENAME)
FONT_URL = "https://github.com/google/fonts/raw/main/apache/tinos/Tinos-Regular.ttf"

def ensure_font_exists():
    """Гарантує, що файл шрифту є на диску."""
    if not os.path.exists(FONT_PATH) or os.path.getsize(FONT_PATH) == 0:
        try:
            print(f"🔄 Завантажую шрифт: {FONT_PATH}")
            opener = urllib.request.build_opener()
            opener.addheaders = [('User-agent', 'Mozilla/5.0')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(FONT_URL, FONT_PATH)
            print("✅ Шрифт успішно завантажено!")
        except Exception as e:
            print(f"❌ Помилка завантаження шрифту: {e}")

class PDFReport(FPDF):
    def header(self):
        # Отримуємо поточну дату
        date_str = datetime.now().strftime("%d.%m.%Y")
        
        # --- ДАТА СПРАВА ЗВЕРХУ ---
        try:
            self.set_font("TimesUA", size=10)
        except:
            self.set_font("Times", "I", 10) # Fallback
            
        # ln=1 переводить на новий рядок після дати
        self.cell(0, 10, date_str, ln=1, align='R')

    def footer(self):
        self.set_y(-15)
        try:
            self.set_font("TimesUA", size=8)
        except:
            self.set_font("Times", "I", 8)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.close('all')
    plt.clf()
    plt.rcParams.update({
        'font.size': 10,
        'font.family': 'serif'
    })
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    is_scale = (qs.question.qtype == QuestionType.SCALE)
    if not is_scale:
        try:
            vals = pd.to_numeric(qs.table["Варіант відповіді"], errors='coerce')
            if vals.notna().all() and vals.min() >= 0 and vals.max() <= 10:
                is_scale = True
        except: pass

    if is_scale:
        fig = plt.figure(figsize=(6.0, 4.0))
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
        fig = plt.figure(figsize=(6.0, 4.0))
        colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
        c_arg = colors[:len(values)] if len(values) <= len(colors) else None
        wedges, texts, autotexts = plt.pie(
            values, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.8, colors=c_arg, radius=1.0
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_weight('bold')
            import matplotlib.patheffects as path_effects
            autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])
        
        plt.axis('equal')
        cols = 2 if len(labels) > 3 else 1
        plt.legend(wrapped_labels, loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=cols, frameon=False, fontsize=8)

    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=CHART_DPI, bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

def build_pdf_report(original_df, sliced_df, summaries, range_info) -> bytes:
    ensure_font_exists()
    
    pdf = PDFReport()
    
    font_ok = False
    if os.path.exists(FONT_PATH) and os.path.getsize(FONT_PATH) > 0:
        try:
            pdf.add_font("TimesUA", fname=FONT_PATH)
            font_ok = True
        except Exception as e:
            print(f"⚠️ Font error: {e}")

    pdf.add_page()
    
    # --- ТИТУЛЬНА ЧАСТИНА (Тіло документа) ---
    if font_ok:
        pdf.set_font("TimesUA", size=16)
        pdf.cell(0, 10, "Звіт про результати опитування", ln=1, align='C')
        
        pdf.set_font("TimesUA", size=12)
        safe_range = range_info.replace('–', '-').replace('—', '-')
        
        pdf.cell(0, 10, f"Всього анкет: {len(original_df)}", ln=1, align='C')
        pdf.cell(0, 10, f"Оброблено: {len(sliced_df)}", ln=1, align='C')
        pdf.cell(0, 10, safe_range, ln=1, align='C')
    else:
        # Fallback
        pdf.set_font("Times", "B", 16)
        pdf.cell(0, 10, "Survey Report", ln=1, align='C')
        pdf.set_font("Times", size=12)
        pdf.cell(0, 10, f"Count: {len(sliced_df)}", ln=1, align='C')
    
    pdf.ln(5)

    for qs in summaries:
        if qs.table.empty: continue
        
        title = f"{qs.question.code}. {qs.question.text}"
        title = title.replace('–', '-').replace('—', '-').replace('’', "'")
        
        if font_ok:
            pdf.set_font("TimesUA", size=12)
            pdf.multi_cell(0, 6, title)
        else:
            pdf.set_font("Times", "B", 12)
            pdf.multi_cell(0, 6, f"Question {qs.question.code}")

        pdf.ln(2)

        # Таблиця
        if font_ok: pdf.set_font("TimesUA", size=11)
        else: pdf.set_font("Times", size=10)

        col_w1 = 110
        col_w2 = 30
        
        h1 = "Варіант" if font_ok else "Option"
        h2 = "Кільк." if font_ok else "Count"
        h3 = "%"
        
        pdf.cell(col_w1, 8, h1, border=1, ln=0)
        pdf.cell(col_w2, 8, h2, border=1, ln=0)
        pdf.cell(col_w2, 8, h3, border=1, ln=1)
        
        for row in qs.table.itertuples(index=False):
            val_text = str(row[0])[:60].replace('–', '-').replace('—', '-').replace('’', "'")
            
            if not font_ok and not val_text.isascii():
                val_text = "..."

            pdf.cell(col_w1, 8, val_text, border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[1]), border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[2]), border=1, ln=1)
            
        pdf.ln(5)

        try:
            img = create_chart_image(qs)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(img.getvalue())
                name = tmp.name
            
            pdf.image(name, w=140, x=35)
            os.unlink(name)
            pdf.ln(10)
        except:
            pdf.cell(0, 10, "[Chart Error]", ln=1)

        if pdf.get_y() > 240:
            pdf.add_page()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        pdf.output(tmp_pdf.name)
        tmp_name = tmp_pdf.name
        
    with open(tmp_name, 'rb') as f:
        pdf_bytes = f.read()
    os.unlink(tmp_name)
    
    return pdf_bytes