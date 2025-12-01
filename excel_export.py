"""
Модуль формування Excel-звіту з результатами опитування.

Застосування звітів у форматі XLSX відповідає практиці документування
результатів соціологічних досліджень у ЗВО та вимогам внутрішніх
систем забезпечення якості освіти.
"""

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
    """
    output = io.BytesIO()

    # Створюємо об'єкт writer з рушієм xlsxwriter
    # nan_inf_to_errors=True допомагає уникнути помилок при записі NaN
    with pd.ExcelWriter(output, engine="xlsxwriter", engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
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
            "border": 1,
            "text_wrap": True
        })
        percent_fmt = workbook.add_format({'num_format': '0.0'})

        # ---------------------------------------------------------
        # 1. Технічна інформація
        # ---------------------------------------------------------
        meta_df = pd.DataFrame({
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
        })
        meta_df.to_excel(writer, sheet_name="Технічна_інформація", index=False)
        
        # Налаштування ширини колонок
        ws_meta = writer.sheets["Технічна_інформація"]
        ws_meta.set_column(0, 0, 40)
        ws_meta.set_column(1, 1, 50)

        # ---------------------------------------------------------
        # 2. Вихідні дані
        # ---------------------------------------------------------
        sliced_df.to_excel(writer, sheet_name="Вихідні_дані", index=False)

        # ---------------------------------------------------------
        # 3. Підсумки та діаграми
        # ---------------------------------------------------------
        sheet_name = "Підсумки"
        worksheet = workbook.add_worksheet(sheet_name)

        # Налаштування колонок: A (Варіант), B (Кількість), C (%)
        worksheet.set_column(0, 0, 50)
        worksheet.set_column(1, 2, 12)

        current_row = 0

        for qs in summaries:
            # -- 3.1. Заголовок питання --
            # Об'єднуємо комірки A, B, C для заголовка
            q_title = f"{qs.question.code}. {qs.question.text}"
            worksheet.merge_range(current_row, 0, current_row, 2, q_title, title_fmt)
            current_row += 1

            # Якщо таблиця порожня
            if qs.table.empty:
                worksheet.write(current_row, 0, "Немає даних або текстові відповіді.")
                current_row += 2
                continue

            # -- 3.2. Ручний запис таблиці --
            # Заголовки: Варіант, Кількість, %
            columns = qs.table.columns.tolist()
            for col_idx, col_name in enumerate(columns):
                worksheet.write(current_row, col_idx, col_name, header_fmt)
            
            # Дані
            data_rows = qs.table.values.tolist()
            start_data_row = current_row + 1
            
            for i, row_data in enumerate(data_rows):
                # row_data[0] -> Варіант (текст)
                # row_data[1] -> Кількість (int)
                # row_data[2] -> Відсоток (float)
                worksheet.write(start_data_row + i, 0, row_data[0])
                worksheet.write(start_data_row + i, 1, row_data[1])
                worksheet.write(start_data_row + i, 2, row_data[2], percent_fmt)

            n_items = len(data_rows)
            end_data_row = start_data_row + n_items - 1

            # Якщо раптом даних 0, йдемо далі
            if n_items == 0:
                current_row = start_data_row + 2
                continue

            # -- 3.3. Побудова діаграми --
            if qs.question.qtype == QuestionType.SCALE:
                chart_type = 'column'
            else:
                chart_type = 'pie'

            chart = workbook.add_chart({'type': chart_type})

            # Синтаксис: [sheetname, first_row, first_col, last_row, last_col]
            # Колонка 0 (A) - категорії, Колонка 1 (B) - значення
            categories_ref = [sheet_name, start_data_row, 0, end_data_row, 0]
            values_ref = [sheet_name, start_data_row, 1, end_data_row, 1]

            chart.add_series({
                'name':       'Кількість',
                'categories': categories_ref,
                'values':     values_ref,
                'data_labels': {'value': True, 'percentage': (chart_type == 'pie')},
            })

            # Назва діаграми (код питання)
            chart.set_title({'name': str(qs.question.code)})
            chart.set_style(10)

            if chart_type == 'column':
                chart.set_legend({'position': 'none'})
                chart.set_x_axis({'name': 'Варіант'})
                chart.set_y_axis({'name': 'Кількість'})

            # Вставка діаграми у клітинку E (індекс 4) навпроти початку питання
            # Зміщення: current_row вказував на заголовок таблиці
            chart_insert_row = current_row 
            worksheet.insert_chart(chart_insert_row, 4, chart)

            # Розрахунок відступу для наступного питання
            # Висота блоку = заголовок(1) + хедер(1) + рядки даних + відступ
            # Або висота діаграми (~15 рядків)
            needed_rows_for_data = n_items + 3
            needed_rows_for_chart = 18
            block_height = max(needed_rows_for_data, needed_rows_for_chart)
            
            # Оновлюємо current_row для наступного циклу (відраховуємо від заголовка питання)
            # Поточний current_row був на "хедері" таблиці, тому повертаємось на рядок заголовка (-1) і додаємо висоту
            current_row = (current_row - 1) + block_height

    return output.getvalue()