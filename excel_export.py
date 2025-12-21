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

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        workbook = writer.book

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
        # Підсумки та діаграми
        # ---------------------------------------------------------
        sheet_name = "Підсумки"
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.set_column(0, 0, 50)
        worksheet.set_column(1, 2, 12)

        current_row = 0

        for qs in summaries:
            # --Заголовок питання --
            # Об'єднуємо комірки A, B, C для заголовка
            q_title = f"{qs.question.code}. {qs.question.text}"
            worksheet.merge_range(current_row, 0, current_row, 2, q_title, title_fmt)
            current_row += 1

            if qs.table.empty:
                worksheet.write(current_row, 0, "Немає даних або текстові відповіді.")
                current_row += 2
                continue

            # -- Запис таблиці --
            # Заголовки: Варіант, Кількість, %
            columns = qs.table.columns.tolist()
            for col_idx, col_name in enumerate(columns):
                worksheet.write(current_row, col_idx, col_name, header_fmt)
            
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

            if n_items == 0:
                current_row = start_data_row + 2
                continue

            # -- Побудова діаграми --
            if qs.question.qtype == QuestionType.SCALE:
                chart_type = 'column'
            else:
                chart_type = 'pie'

            chart = workbook.add_chart({'type': chart_type})

            categories_ref = [sheet_name, start_data_row, 0, end_data_row, 0]
            values_ref = [sheet_name, start_data_row, 1, end_data_row, 1]

            chart.add_series({
                'name':       'Кількість',
                'categories': categories_ref,
                'values':     values_ref,
                'data_labels': {'value': True, 'percentage': (chart_type == 'pie')},
            })

            chart.set_title({'name': str(qs.question.code)})
            chart.set_style(10)

            if chart_type == 'column':
                chart.set_legend({'position': 'none'})
                chart.set_x_axis({'name': 'Варіант'})
                chart.set_y_axis({'name': 'Кількість'})

            chart_insert_row = current_row 
            worksheet.insert_chart(chart_insert_row, 4, chart)

            needed_rows_for_data = n_items + 3
            needed_rows_for_chart = 18
            block_height = max(needed_rows_for_data, needed_rows_for_chart)
        
            current_row = (current_row - 1) + block_height

        # Технічна інформація
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
        
        ws_meta = writer.sheets["Технічна_інформація"]
        ws_meta.set_column(0, 0, 40)
        ws_meta.set_column(1, 1, 50)

        # Вихідні дані
        sliced_df.to_excel(writer, sheet_name="Вихідні_дані", index=False)

    return output.getvalue()