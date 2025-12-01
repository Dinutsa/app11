"""
Головний модуль веб-застосунку для обробки результатів студентських опитувань.

Забезпечує:
- завантаження файлів Excel із Google Forms;
- вибір діапазону рядків;
- класифікацію типів питань;
- обчислення частот і відсотків;
- побудову кругових діаграм;
- формування Excel-звіту.

Такий підхід відповідає рекомендаціям з автоматизації обробки
соціологічних досліджень та потребам систем внутрішнього забезпечення
якості освіти ЗВО.
"""
import streamlit as st
import plotly.express as px
import pandas as pd
from typing import List

from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries
from excel_export import build_excel_report

st.set_page_config(
    page_title="Обробка результатів студентських опитувань",
    layout="wide",
)


def ensure_session_state_keys():
    keys = {
        "uploaded_files_info": None,  # list of filenames (to detect change)
        "ld": None,                   # LoadedData
        "sliced": None,               # sliced DataFrame
        "qinfo": None,                # question metadata dict
        "summaries": None,            # list of summaries
        "processed": False,           # flag processed
        "selected_code": None,        # last selected question code
        "from_row": None,
        "to_row": None,
    }
    for k, v in keys.items():
        if k not in st.session_state:
            st.session_state[k] = v


def files_changed(uploaded_files: List[st.runtime.uploaded_file_manager.UploadedFile]) -> bool:
    """Проста перевірка чи змінились файли (по іменах та кількості)."""
    if not uploaded_files:
        return False
    names = [getattr(f, "name", str(i)) for i, f in enumerate(uploaded_files)]
    prev = st.session_state.uploaded_files_info
    return prev != names


def store_uploaded_files_info(uploaded_files):
    if not uploaded_files:
        st.session_state.uploaded_files_info = None
    else:
        st.session_state.uploaded_files_info = [getattr(f, "name", str(i)) for i, f in enumerate(uploaded_files)]


def process_range(uploaded_files, from_row: int, to_row: int):
    """
    Завантажує файли (якщо потрібно), робить slice, класифікацію, summaries
    і зберігає результати у session_state.
    """
    # Завантаження й підготовка
    ld = load_excels(uploaded_files)
    sliced = slice_range(ld, int(from_row), int(to_row))

    # класифікація (припускаємо 1 технічний стовпець - позначка часу)
    qinfo = classify_questions(sliced, technical_columns=1)

    # summary
    summaries = build_all_summaries(sliced, qinfo)

    # запишемо у session_state
    st.session_state.ld = ld
    st.session_state.sliced = sliced
    st.session_state.qinfo = qinfo
    st.session_state.summaries = summaries
    st.session_state.processed = True
    st.session_state.from_row = int(from_row)
    st.session_state.to_row = int(to_row)

    # обрати дефолтний код (якщо ще не було вибору)
    codes = [qs.question.code for qs in summaries]
    if codes:
        st.session_state.selected_code = st.session_state.selected_code or codes[0]
    else:
        st.session_state.selected_code = None


def clear_processing_state():
    st.session_state.ld = None
    st.session_state.sliced = None
    st.session_state.qinfo = None
    st.session_state.summaries = None
    st.session_state.processed = False
    st.session_state.selected_code = None
    st.session_state.uploaded_files_info = None
    st.session_state.from_row = None
    st.session_state.to_row = None


def main():
    ensure_session_state_keys()

    st.title("Система обробки результатів студентських опитувань")
    st.markdown(
        """Завантажте Excel-файли (експорт з Google Forms), оберіть діапазон рядків і натисніть
        «Обробити діапазон». Після обробки ви можете вибирати любое питання для перегляду таблиці/діаграми
        без повторної обробки."""
    )

    st.sidebar.header("Налаштування аналізу")

    uploaded_files = st.sidebar.file_uploader(
        "Завантажте файл(и) Excel з відповідями",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Підтримується кілька файлів Google Forms; вони будуть об'єднані."
    )

    # Кнопки: обробити та скидання
    process_button = st.sidebar.button("Обробити діапазон")
    reset_button = st.sidebar.button("Скинути обробку")

    # Якщо файли змінилися з моменту останнього збереження — скинути попередні результати
    if uploaded_files:
        if files_changed(uploaded_files):
            # помічаємо нові файли та очищуємо обробку
            store_uploaded_files_info(uploaded_files)
            clear_processing_state()
    else:
        # якщо немає файлів в UI — нічого не робимо (але не обнуляємо session_state, користувач може вже мати оброблені дані)
        pass

    # Якщо користувач натиснув reset — очищуємо
    if reset_button:
        clear_processing_state()
        st.experimental_rerun()

    # Якщо раніше вже були оброблені дані, покажемо короткий огляд
    if st.session_state.processed and st.session_state.ld is not None:
        ld = st.session_state.ld
        st.subheader("Короткий огляд останньо оброблених відповідей")
        st.write(f"Кількість рядків у завантажених даних: **{ld.n_rows}**, стовпців: **{ld.n_cols}**")
        st.dataframe(ld.df.head())
        # дозволяємо змінювати діапазон, але не переобробляти автоматично
        min_row, max_row = get_row_bounds(ld)
        from_row = st.sidebar.number_input(
            "Від (рядок)", min_value=min_row, max_value=max_row,
            value=st.session_state.from_row or min_row, step=1
        )
        to_row = st.sidebar.number_input(
            "До (рядок)", min_value=min_row, max_value=max_row,
            value=st.session_state.to_row or max_row, step=1
        )
        # Якщо користувач змінив діапазон та натискає процес — процесимо заново
        if process_button:
            try:
                process_range(uploaded_files or [ ], from_row, to_row)
            except Exception as e:
                st.error(f"Помилка обробки: {e}")
                return
    else:
        # Ніяких оброблених даних ще немає — показуємо інтерфейс для вибору діапазону на основі файлів (якщо вони є)
        if not uploaded_files:
            st.info("Завантажте хоча б один файл Excel, щоб розпочати аналіз.")
            return
        # визначаємо row bounds на основі тимчасового завантаження (без збереження у session_state)
        try:
            tmp_ld = load_excels(uploaded_files)
            min_row, max_row = get_row_bounds(tmp_ld)
        except Exception as e:
            st.error(str(e))
            return

        from_row = st.sidebar.number_input("Від (рядок)", min_value=min_row, max_value=max_row, value=min_row, step=1)
        to_row = st.sidebar.number_input("До (рядок)", min_value=min_row, max_value=max_row, value=max_row, step=1)

        if process_button:
            try:
                process_range(uploaded_files, from_row, to_row)
            except Exception as e:
                st.error(f"Помилка обробки: {e}")
                return
        else:
            st.info("Оберіть діапазон та натисніть «Обробити діапазон».")

    # --- Після успішної обробки показуємо результати (без повторної обробки) ---
    if st.session_state.processed:
        summaries = st.session_state.summaries or []
        qinfo = st.session_state.qinfo or {}
        sliced = st.session_state.sliced

        st.subheader("Класифікація запитань")
        q_table = pd.DataFrame(
            [
                {"Код": info.code, "Назва стовпця": col, "Тип": info.qtype.value}
                for col, info in qinfo.items()
                if info.qtype != QuestionType.TECHNICAL
            ]
        )
        st.dataframe(q_table)

        st.subheader("Таблиці відповідей за всіма запитаннями")
        for qs in summaries:
            with st.expander(f"{qs.question.code}. {qs.question.text}"):
                st.dataframe(qs.table)

        # вибір питання для діаграми — використовуємо session_state.selected_code
        st.subheader("Окреме питання (діаграма)")

        available_codes = [qs.question.code for qs in summaries]
        if not available_codes:
            st.warning("Немає запитань з варіантами для побудови діаграм.")
        else:
            # Якщо selected_code вже є в session_state і все ще доступний — залишити його
            if st.session_state.selected_code not in available_codes:
                st.session_state.selected_code = available_codes[0]

            selected_code = st.selectbox(
                "Оберіть код питання:",
                options=available_codes,
                index=available_codes.index(st.session_state.selected_code) if st.session_state.selected_code in available_codes else 0,
                key="select_question_box"
            )

            # оновлюємо збережений вибір без повторної обробки
            if selected_code != st.session_state.selected_code:
                st.session_state.selected_code = selected_code

            selected = next((qs for qs in summaries if qs.question.code == st.session_state.selected_code), None)

            if selected is None or selected.table.empty:
                st.warning("Для цього питання немає даних для побудови діаграми.")
            else:
                fig = px.pie(
                    selected.table,
                    names="Варіант відповіді",
                    values="Кількість",
                    title=f"{selected.question.code}. {selected.question.text}",
                    hole=0.0,
                )
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(selected.table)

        # --- Експорт звіту (take from session_state to avoid recompute) ---
        st.subheader("Експорт результатів")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row} (усього {len(sliced)} анкет)"
        report_bytes = build_excel_report(
            original_df=st.session_state.ld.df,
            sliced_df=st.session_state.sliced,
            qinfo=st.session_state.qinfo,
            summaries=st.session_state.summaries,
            range_info=range_info,
        )

        st.download_button(
            label="Завантажити звіт (Excel)",
            data=report_bytes,
            file_name="opituvalny_analiz_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
