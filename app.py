import io
import os
import streamlit as st
import plotly.express as px
import pandas as pd

# Імпорти модулів логіки
from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries

# Імпорти модулів експорту
from excel_export import build_excel_report
from pdf_export import build_pdf_report
from docx_export import build_docx_report
from pptx_export import build_pptx_report

st.set_page_config(
    page_title="Обробка результатів студентських опитувань",
    layout="wide",
)

def init_state():
    defaults = {
        "uploaded_files_store": None,
        "ld": None,
        "sliced": None,
        "qinfo": None,
        "summaries": None,
        "processed": False,
        "selected_code": None,
        "from_row": 0,
        "to_row": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

st.title("Аналіз результатів опитувань (Google Forms)")

# --- БІЧНА ПАНЕЛЬ ---
with st.sidebar:
    st.header("1. Завантаження та Налаштування")
    
    # Завантаження файлів
    uploaded_files = st.file_uploader(
        "Оберіть Excel-файли (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    # ЛОГІКА: Якщо файли обрані, одразу вантажимо їх у пам'ять, щоб показати слайдер
    if uploaded_files:
        # Перевіряємо, чи це нові файли, щоб не перечитувати зайвий раз
        if st.session_state.ld is None or uploaded_files != st.session_state.uploaded_files_store:
            try:
                ld = load_excels(uploaded_files)
                st.session_state.ld = ld
                st.session_state.uploaded_files_store = uploaded_files
                
                # Скидаємо межі на дефолтні при новому файлі
                min_r, max_r = get_row_bounds(ld)
                st.session_state.from_row = min_r
                st.session_state.to_row = max_r
                st.session_state.processed = False # Скидаємо прапорець обробки
                
            except Exception as e:
                st.error(f"Помилка читання файлу: {e}")

    # Якщо дані завантажені, показуємо слайдер і кнопки керування
    if st.session_state.ld:
        st.success(f"Файл успішно зчитано: {st.session_state.ld.n_rows} анкет.")
        st.divider()
        st.header("2. Вибір діапазону")
        
        min_r, max_r = get_row_bounds(st.session_state.ld)
        
        if max_r > min_r:
            r_range = st.slider(
                "Оберіть рядки для обробки:",
                min_value=min_r,
                max_value=max_r,
                value=(st.session_state.from_row, st.session_state.to_row)
            )
            st.session_state.from_row = r_range[0]
            st.session_state.to_row = r_range[1]
        
        st.divider()
        
        # Кнопки дій
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("🚀 Обробити / Оновити", type="primary"):
                # Основна логіка обробки
                sliced = slice_range(st.session_state.ld, st.session_state.from_row, st.session_state.to_row)
                st.session_state.sliced = sliced
                
                qinfo = classify_questions(sliced)
                st.session_state.qinfo = qinfo
                
                summaries = build_all_summaries(sliced, qinfo)
                st.session_state.summaries = summaries
                
                st.session_state.processed = True
        
        with col_btn2:
            if st.button("❌ Скинути"):
                # Повне очищення стану
                st.session_state.ld = None
                st.session_state.uploaded_files_store = None
                st.session_state.processed = False
                st.session_state.sliced = None
                st.session_state.summaries = None
                st.rerun()

# --- ОСНОВНА ЧАСТИНА ---
if st.session_state.processed and st.session_state.sliced is not None:
    sliced = st.session_state.sliced
    summaries = st.session_state.summaries
    
    tab1, tab2 = st.tabs(["📊 Аналіз даних", "📥 Експорт звітів"])
    
    # ---------------- ВКЛАДКА АНАЛІЗУ ----------------
    with tab1:
        st.info(f"**В роботі {len(sliced)} анкет** (рядки {st.session_state.from_row}–{st.session_state.to_row})")
        
        # 1. ПЕРЕГЛЯД ВИХІДНИХ ДАНИХ
        with st.expander("🔍 Перегляд таблиці вихідних даних", expanded=False):
            st.dataframe(sliced, use_container_width=True)
        
        st.divider()
        
        # 2. ВИБІРКОВИЙ ПЕРЕГЛЯД
        st.subheader("Детальний перегляд окремого питання")
        options = [qs.question.code for qs in summaries]
        selected_code = st.selectbox("Оберіть питання зі списку:", options)
        
        if selected_code:
            st.session_state.selected_code = selected_code
            selected = next((qs for qs in summaries if qs.question.code == st.session_state.selected_code), None)

            if selected is None or selected.table.empty:
                st.warning("Немає даних для відображення.")
            else:
                st.markdown(f"**{selected.question.code}. {selected.question.text}**")
                
                col_chart, col_table = st.columns([1.5, 1])
                
                with col_chart:
                    # Повна кругова діаграма (hole=0)
                    fig = px.pie(
                        selected.table,
                        names="Варіант відповіді",
                        values="Кількість",
                        hole=0, 
                        title="Розподіл відповідей"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                with col_table:
                    st.write("Статистика:")
                    st.dataframe(selected.table, use_container_width=True)

        # 3. ПОВНИЙ СПИСОК УСІХ ПИТАНЬ (Знизу)
        st.divider()
        st.subheader("📋 Повний огляд всіх питань")
        
        for qs in summaries:
            if qs.table.empty:
                continue
                
            # Використовуємо expander, щоб не захаращувати екран, але він відкритий
            with st.expander(f"{qs.question.code}. {qs.question.text}", expanded=True):
                c_chart, c_tbl = st.columns([1, 1])
                
                with c_chart:
                     fig_all = px.pie(
                        qs.table,
                        names="Варіант відповіді",
                        values="Кількість",
                        hole=0
                    )
                     st.plotly_chart(fig_all, use_container_width=True, key=f"chart_{qs.question.code}")
                
                with c_tbl:
                    st.dataframe(qs.table, use_container_width=True)


    # ---------------- ВКЛАДКА ЕКСПОРТУ ----------------
    with tab2:
        st.subheader("Завантаження звітів")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row} (усього {len(sliced)} анкет)"
        
        # --- КЕШОВАНІ ФУНКЦІЇ ГЕНЕРАЦІЇ ---
        
        @st.cache_data(show_spinner="Генеруємо Excel...")
        def get_excel_data(_original_df, _sliced_df, _qinfo, _summaries, _range_info):
            return build_excel_report(_original_df, _sliced_df, _qinfo, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо PDF...")
        def get_pdf_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_pdf_report(_original_df, _sliced_df, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо DOCX...")
        def get_docx_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_docx_report(_original_df, _sliced_df, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо PowerPoint...")
        def get_pptx_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_pptx_report(_original_df, _sliced_df, _summaries, _range_info)

        # --- КНОПКИ ---
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            if st.button("📊 Excel звіт"):
                with st.spinner("Створення файлу..."):
                    try:
                        excel_bytes = get_excel_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.qinfo, st.session_state.summaries, range_info)
                        st.download_button("📥 Завантажити Excel", excel_bytes, "survey_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e: st.error(f"Error: {e}")

        with c2:
            if st.button("📄 PDF звіт"):
                with st.spinner("Створення файлу..."):
                    try:
                        pdf_bytes = get_pdf_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                        st.download_button("📥 Завантажити PDF", pdf_bytes, "survey_results.pdf", "application/pdf")
                    except Exception as e: st.error(f"Error: {e}")

        with c3:
            if st.button("📝 Word звіт"):
                with st.spinner("Створення файлу..."):
                    try:
                        docx_bytes = get_docx_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                        st.download_button("📥 Завантажити Word", docx_bytes, "survey_results.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e: st.error(f"Error: {e}")

        with c4:
            if st.button("🖥️ PPTX звіт"):
                with st.spinner("Створення файлу..."):
                    try:
                        pptx_bytes = get_pptx_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                        st.download_button("📥 Завантажити PPTX", pptx_bytes, "survey_results.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
                    except Exception as e: st.error(f"Error: {e}")

elif not st.session_state.ld:
    st.info("👈 Будь ласка, завантажте Excel-файл у меню ліворуч.")