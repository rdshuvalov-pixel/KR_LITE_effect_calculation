"""Объединение отчётов Энтерсайт — склейка периодов с подённой разбивкой."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import streamlit as st
from scripts.generate import merge_entersite_from_uploads

st.set_page_config(page_title="Объединение Энтерсайт", page_icon="📎", layout="centered")
st.title("Объединение отчётов Энтерсайт")
st.caption("Склейка периодов: агрегат за период → подённая разбивка. При пересечении приоритет у более позднего файла.")

uploaded = st.file_uploader("Загрузите xlsx с периодом в имени (*_с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ_*)", type=["xlsx"], accept_multiple_files=True)

if uploaded:
    if st.button("Объединить и нарезать"):
        files = [(f.name, f.read()) for f in uploaded]
        result, msg = merge_entersite_from_uploads(files)
        if result is not None:
            st.success(msg)
            st.download_button(
                "Скачать Итоговые_данные_Энтерсайт.xlsx",
                data=result,
                file_name="Итоговые_данные_Энтерсайт.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(msg)

with st.expander("Формат файлов (см. SKILL.md)"):
    st.markdown("""
- **Имя:** `*_с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ_*.xlsx`
- **Лист:** B=Артикул, C=Товар, L=Количество, M=Закуп, N=Розница
- **Выход:** листы Себестоимость и Продажи (готовы для калькулятора)
    """)
