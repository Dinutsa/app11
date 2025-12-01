"""
Модуль формування Excel-звіту з результатами опитування.

Застосування звітів у форматі XLSX відповідає практиці документування
результатів соціологічних досліджень у ЗВО та вимогам внутрішніх
систем забезпечення якості освіти.
"""
"""
Модуль формування Excel-звіту з результатами опитування.

Застосування звітів у форматі XLSX відповідає практиці документування
результатів соціологічних досліджень у ЗВО та вимогам внутрішніх
систем забезпечення якості освіти.
"""

"""
Модуль формування Excel-звіту з результатами опитування.
"""

from __future__ import annotations

import io
from typing import Dict, List

import pandas as pd
import xlsxwriter

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary


def build_excel_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    qinfo: Dict[str, QuestionInfo],
    summaries: List[QuestionSummary],
    range_info: str,
) -> bytes:
    """
    Створює Excel-звіт та повертає його у вигляді байтів.
    Використовує прямий запис xlsxwriter для уникнення конфліктів з pandas.
    """
    output = io.BytesIO()

    # Створюємо об'єкт workbook
    # ВАЖЛИВО: engine_kwargs={'options': {'nan_inf_to_errors': True}} допомагає уникнути помилок з NaN
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # --- Стилі ---
        header_fmt = workbook.add_format({
            "bold": True, 
            "font_size": 11,
            "bottom": 1,
            "bg_color": "#F2F2F2"
        })
        title_fmt = workbook.add_format({
            "bold": True, 
            "font_size": 12, 
            "fg_color": "#DCE6F1",
            "border": 1
        })
        # Стиль для відсотків (опціонально)
        percent_fmt = workbook.add_format({'num_format': '0.0'})
        
        # ---------------------------------------------------------
        # 1. Технічна інформація (використовуємо pandas для простоти)
        # ---------------------------------------------------------
        meta_df = pd.DataFrame(
            {
                "Параметр": [
                    "Загальна кількість анкет у файлі",
                    "Кількість анкет у вибраному діапазоні",
                    "Діапазон обробки",
                ],
                "Значення": [
                    len(original_df),
                    len(sliced_df),
                    range_info,
                ],
            }
        )
        meta_df.to_excel(writer, sheet_name="Технічна_інформація", index=False)
        ws_meta = writer.sheets["Технічна_інформація"]
        ws_meta.set_column(0, 0, 40)
        ws_meta.set_column(1, 1, 50)

        # ---------------------------------------------------------
        # 2. Вихідні дані
        # ---------------------------------------------------------
        sliced_df.to_excel(writer, sheet_name="Вихідні_дані", index=False)

        # ---------------------------------------------------------
        # 3. Таблиці підсумків та діаграми
        # ---------------------------------------------------------
        sheet_name = "Підсумки"
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Налаштування колонок
        worksheet.set_column(0, 0, 50)  # Варіант відповіді
        worksheet.set_column(1, 2, 12)  # Числа
        
        current_row = 0

        for qs in summaries:
            # -- 3.1. Заголовок питання --
            q_title = f"{qs.question.code}. {qs.question.text}"
            worksheet.merge_range(current_row, 0, current_row, 2, q_title, title_fmt)
            current_row += 1

            if qs.table.empty:
                worksheet.write(current_row, 0, "Немає даних або текстові відповіді.")
                current_row += 2
                continue

            # -- 3.2. Ручний запис таблиці (щоб не ламати worksheet через pandas) --
            # Заголовки таблиці
            columns = qs.table.columns.tolist() # ["Варіант", "Кількість", "%"]
            for col_idx, col_name in enumerate(columns):
                worksheet.write(current_row, col_idx, col_name, header_fmt)
            
            # Дані таблиці
            # Конвертуємо в numpy array або list of lists
            data_rows = qs.table.values.tolist()
            start_data_row = current_row + 1
            
            for i, row_data in enumerate(data_rows):
                # row_data[0] -> Варіант, row_data[1] -> Кількість, row_data[2] -> %
                worksheet.write(start_data_row + i, 0, row_data[0])
                worksheet.write(start_data_row + i, 1, row_data[1])
                worksheet.write(start_data_row + i, 2, row_data[2], percent_fmt)

            n_items = len(data_rows)
            end_data_row = start_data_row + n_items - 1
            
            # Якщо даних немає (на всяк випадок), пропускаємо діаграму
            if n_items == 0:
                current_row = start_data_row + 2
                continue

            # -- 3.3. Побудова діаграми --
            # Вибір типу
            if qs.question.qtype == QuestionType.SCALE:
                chart_type = 'column'
            else:
                chart_type = 'pie'

            chart = workbook.add_chart({'type': chart_type})

            # Посилання на дані: [sheet, first_row, first_col, last_row, last_col]
            # Колонка 0 - категорії, Колонка 1 - значення (Кількість)
            
            chart.add_series({
                'name':       'Кількість',
                'categories': [sheet_name, start_data_row, 0, end_data_row, 0],
                'values':     [sheet_name, start_data_row, 1, end_data_row, 1],
                'data_labels': {'value': True, 'percentage': (chart_type == 'pie')},
            })

            # Назва діаграми (коротка, тільки код питання, бо текст довгий)
            chart.set_title({'name': str(qs.question.code)})
            chart.set_style(10)

            if chart_type == 'column':
                 chart.set_legend({'position': 'none'})
                 chart.set_x_axis({'name': 'Варіант'})
                 chart.set_y_axis({'name': 'Кількість'})

            # Вставка діаграми
            worksheet.insert_chart(current_row, 4, chart)

            # Відступ для наступного блоку
            # Висота блоку = заголовок + хедер + рядки даних
            # Або висота діаграми (~15 рядків)
            block_height = max(n_items + 2, 18)
            current_row += block_height

    return output.getvalue()