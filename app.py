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
# вставити зверху файлу
import io
import streamlit as st
import plotly.express as px
import pandas as pd

from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries
from excel_export import build_excel_report

st.set_page_config(page_title="Обробка результатів студентських опитувань", layout="wide")

# ----------------- допоміжні функції -----------------

def init_state():
    keys = [
        "uploaded_files_store",  # [{"name":..., "bytes": b'...'}, ...]
        "ld", "sliced", "qinfo", "summaries",
        "processed", "selected_code",
        "from_row", "to_row"
    ]
    for k in keys:
        if k not in st.session_state:
            st.session_state[k] = None
    if st.session_state.processed is None:
        st.session_state.processed = False

def store_uploaded_files(uploaded_files):
    """
    Зчитати байти з UploadedFile і зберегти в session_state.
    Це робить завантажені файли незалежними від тимчасових UploadedFile-об'єктів.
    """
    if not uploaded_files:
        return
    store = []
    for f in uploaded_files:
        # f.read() дає bytes; збережемо і назву
        try:
            b = f.read()
            name = getattr(f, "name", None) or "file"
            store.append({"name": name, "bytes": b})
        except Exception as e:
            st.error(f"Не вдалося зчитати файл {getattr(f,'name', '')}: {e}")
            return
    st.session_state.uploaded_files_store = store

def get_uploaded_files_from_store():
    """
    Повернути список file-like об'єктів (BytesIO) для подачі в load_excels,
    або None якщо нічого не збережено.
    """
    if not st.session_state.uploaded_files_store:
        return None
    objs = []
    for item in st.session_state.uploaded_files_store:
        bio = io.BytesIO(item["bytes"])
        # Pandas може приймати BytesIO
        bio.name = item["name"]
        objs.append(bio)
    return objs

def clear_all_state():
    st.session_state.uploaded_files_store = None
    st.session_state.ld = None
    st.session_state.sliced = None
    st.session_state.qinfo = None
    st.session_state.summaries = None
    st.session_state.processed = False
    st.session_state.selected_code = None
    st.session_state.from_row = None
    st.session_state.to_row = None

def process_range_from_store(from_row: int, to_row: int):
    """
    Беремо байти з session_state.uploaded_files_store, створюємо BytesIO
    і запускаємо обробку (load_excels, slice, classify, summary).
    Зберігаємо результат у session_state.
    """
    files_for_load = get_uploaded_files_from_store()
    if not files_for_load:
        raise ValueError("Немає збережених файлів для обробки.")
    # load_excels у data_loader приймає список file-like
    ld = load_excels(files_for_load)
    sliced = slice_range(ld, int(from_row), int(to_row))
    qinfo = classify_questions(sliced, technical_columns=1)
    summaries = build_all_summaries(sliced, qinfo)

    st.session_state.ld = ld
    st.session_state.sliced = sliced
    st.session_state.qinfo = qinfo
    st.session_state.summaries = summaries
    st.session_state.processed = True
    st.session_state.from_row = int(from_row)
    st.session_state.to_row = int(to_row)

    # default selected code
    codes = [qs.question.code for qs in summaries]
    st.session_state.selected_code = codes[0] if codes else None

# ----------------- UI -----------------

def main():
    init_state()
    st.title("Система обробки результатів студентських опитувань")

    st.sidebar.header("Керування")
    uploaded_files = st.sidebar.file_uploader(
        "Завантажте Excel (.xlsx) файли",
        type=["xlsx"],
        accept_multiple_files=True
    )

    # Якщо користувач завантажив нові файли у UI — зберігаємо їхні байти в session_state
    if uploaded_files:
        # Якщо в сесії ще немає файлів або змінилися імена файлів — зберігаємо нові
        names_now = [getattr(f, "name", "") for f in uploaded_files]
        names_prev = [it["name"] for it in (st.session_state.uploaded_files_store or [])]
        if names_now != names_prev:
            store_uploaded_files(uploaded_files)
            # скидаємо попередню обробку при завантаженні інших файлів
            clear_all_state()
            # але після збереження, змінюємо uploaded_files_store — треба поставити назад у session_state.processed False
            st.session_state.processed = False

    # Кнопки
    process_btn = st.sidebar.button("Обробити діапазон")
    reset_btn = st.sidebar.button("Скинути")

    # Якщо немає збережених файлів — показати підказку
    if not st.session_state.uploaded_files_store:
        st.info("Завантажте файл(и) Excel у боковій панелі для початку.")
        # поки немає файлів — нічого більше не показуємо
        return

    # Показати короткий перегляд (без повторного load_excels)
    if st.session_state.ld:
        st.subheader("Огляд останньо оброблених даних")
        st.write(f"Рядків у джерелі: **{st.session_state.ld.n_rows}**, стовпців: **{st.session_state.ld.n_cols}**")
        st.dataframe(st.session_state.ld.df.head())
        min_row, max_row = get_row_bounds(st.session_state.ld)
    else:
        # тимчасово дізнаємось межі з файлів у store (не зберігаємо ld!)
        try:
            tmp_objs = get_uploaded_files_from_store()
            tmp_ld = load_excels(tmp_objs)
            min_row, max_row = get_row_bounds(tmp_ld)
            st.write(f"Файли завантажені. Дані готові для обробки. (Доступні рядки: {min_row}–{max_row})")
        except Exception as e:
            st.error(f"Помилка читання файлів: {e}")
            return

    # Поля діапазону (заповнюються з session_state якщо є)
    from_row = st.sidebar.number_input("Від (рядок)", min_value=min_row, max_value=max_row, value=st.session_state.from_row or min_row, step=1)
    to_row = st.sidebar.number_input("До (рядок)", min_value=min_row, max_value=max_row, value=st.session_state.to_row or max_row, step=1)

    # Обробка — викликаємо нашу функцію, яка працює з байтами у state
    if process_btn:
        try:
            process_range_from_store(from_row, to_row)
            st.success("Обробку завершено.")
        except Exception as e:
            st.error(f"Помилка обробки: {e}")
            return

    # Reset
    if reset_btn:
        clear_all_state()
        st.experimental_rerun()

    # Якщо оброблено — показуємо результати (без повторної обробки)
    if st.session_state.processed:
        summaries = st.session_state.summaries or []
        qinfo = st.session_state.qinfo or {}
        sliced = st.session_state.sliced

        st.subheader("Класифікація запитань")
        q_table = pd.DataFrame([{"Код":info.code, "Назва стовпця": col, "Тип": info.qtype.value} for col, info in qinfo.items() if info.qtype != QuestionType.TECHNICAL])
        st.dataframe(q_table)

        st.subheader("Таблиці відповідей")
        for qs in summaries:
            with st.expander(f"{qs.question.code}. {qs.question.text}"):
                st.dataframe(qs.table)

        st.subheader("Окреме питання (діаграма)")
        available_codes = [qs.question.code for qs in summaries]
        if not available_codes:
            st.warning("Немає питань для відображення.")
        else:
            # використай key щоб запам'ятати вибір; selected_code зберігаємо в session_state
            default_idx = 0
            if st.session_state.selected_code and st.session_state.selected_code in available_codes:
                default_idx = available_codes.index(st.session_state.selected_code)
            selected_code = st.selectbox("Оберіть код питання:", options=available_codes, index=default_idx, key="select_question_box")

            # оновити тільки selected_code — дані в summaries не перераховуються
            st.session_state.selected_code = selected_code

            selected = next((qs for qs in summaries if qs.question.code == st.session_state.selected_code), None)
            if selected and not selected.table.empty:
                fig = px.pie(selected.table, names="Варіант відповіді", values="Кількість", title=f"{selected.question.code}. {selected.question.text}")
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(selected.table)
            else:
                st.info("Немає даних для цього питання.")

        # Експорт звіту
        st.subheader("Експорт звіту")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row} (усього {len(st.session_state.sliced)})"
        report_bytes = build_excel_report(original_df=st.session_state.ld.df, sliced_df=st.session_state.sliced, qinfo=st.session_state.qinfo, summaries=st.session_state.summaries, range_info=range_info)
        st.download_button("Завантажити звіт (Excel)", data=report_bytes, file_name="opituvalny_analiz_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
