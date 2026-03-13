"""Инструменты: объединение Энтерсайт, эталон, восстановление себестоимости."""
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import streamlit as st
from scripts.generate import merge_entersite_from_uploads
from create_etalon_file import get_etalon_bytes
from restore_cost_from_sales import restore_cost_from_sales
from restore_cost_history import restore_cost_history

st.set_page_config(page_title="Инструменты", page_icon="🔧", layout="centered")
if st.sidebar.button("← К калькулятору", use_container_width=True):
    st.switch_page("app.py")

st.title("Инструменты")
st.caption("Предобработка данных перед расчётом эффекта.")

# --- 1. Объединение Энтерсайт ---
st.header("Объединение отчётов Энтерсайт")
st.caption("Склейка периодов: агрегат за период → подённая разбивка. При пересечении приоритет у более позднего файла.")

uploaded_entersite = st.file_uploader(
    "Загрузите xlsx с периодом в имени (*_с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ_*)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="entersite_upload",
)

if uploaded_entersite:
    if st.button("Объединить и нарезать", key="btn_merge"):
        files = [(f.name, f.read()) for f in uploaded_entersite]
        result, msg = merge_entersite_from_uploads(files)
        if result is not None:
            st.success(msg)
            st.download_button(
                "Скачать Итоговые_данные_Энтерсайт.xlsx",
                data=result,
                file_name="Итоговые_данные_Энтерсайт.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_merge",
            )
        else:
            st.error(msg)

with st.expander("Формат файлов (см. SKILL.md)", expanded=False):
    st.markdown("""
- **Имя:** `*_с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ_*.xlsx`
- **Лист:** B=Артикул, C=Товар, L=Количество, M=Закуп, N=Розница
- **Выход:** листы Себестоимость и Продажи (готовы для калькулятора)
    """)

st.divider()

# --- 2. Эталон для проверки (свёрнутый) ---
with st.expander("📋 Эталон для проверки", expanded=False):
    etalon_bytes = get_etalon_bytes()
    st.download_button(
        "Скачать etalon_check.xlsx",
        data=etalon_bytes,
        file_name="etalon_check.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_etalon",
    )
    st.markdown("**Как понять результат эталона:**")
    st.markdown("- ✅ **Правильно:** «Обработка завершена успешно», валидных недель > 0, эффект не нулевой, есть категории рост / без изменений / падение.")
    st.markdown("- ❌ **Неправильно:** «сводка пуста», валидных недель 0, все метрики нулевые.")

st.divider()

# --- 3. Восстановление себестоимости ---
st.header("Восстановление себестоимости")
st.caption("Предобработка файла перед расчётом. Результат — скачать изменённый Excel.")

uploaded_restore = st.file_uploader("Загрузите файл XLSX с листами Продажи и Себестоимость", type=["xlsx"], key="restore_upload")

if uploaded_restore:
    if st.session_state.get("restore_file_key") != uploaded_restore.name:
        st.session_state.pop("restore_download", None)
        st.session_state["restore_file_key"] = uploaded_restore.name

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Восстановить себестоимость из продаж", key="btn_restore_sales"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_restore.getvalue())
                tmp_path = tmp.name
            try:
                result = restore_cost_from_sales(tmp_path)
                with open(tmp_path, "rb") as f:
                    st.session_state["restore_download"] = {
                        "data": f.read(),
                        "name": f"restored_from_sales_{uploaded_restore.name}",
                        "msg": f"{result['rows']} строк, {result['products']} товаров",
                    }
                Path(tmp_path).unlink(missing_ok=True)
            except Exception as e:
                Path(tmp_path).unlink(missing_ok=True)
                st.session_state["restore_error"] = str(e)
    with col2:
        if st.button("Восстановить историю себестоимости", key="btn_restore_history"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_restore.getvalue())
                tmp_path = tmp.name
            try:
                result = restore_cost_history(tmp_path)
                if result and result.get("rows", 0) > 0:
                    with open(tmp_path, "rb") as f:
                        st.session_state["restore_download"] = {
                            "data": f.read(),
                            "name": f"restored_history_{uploaded_restore.name}",
                            "msg": f"Добавлено {result['rows']} строк",
                        }
                    Path(tmp_path).unlink(missing_ok=True)
                else:
                    Path(tmp_path).unlink(missing_ok=True)
                    st.session_state["restore_info"] = result.get("message", "Нет данных для восстановления.")
            except Exception as e:
                Path(tmp_path).unlink(missing_ok=True)
                st.session_state["restore_error"] = str(e)

    if st.session_state.get("restore_error"):
        st.error(st.session_state.pop("restore_error"))
    if st.session_state.get("restore_info"):
        st.info(st.session_state.pop("restore_info"))
    if st.session_state.get("restore_download"):
        dl = st.session_state["restore_download"]
        st.success(f"Готово: {dl['msg']}")
        st.download_button(
            "📥 Скачать результат",
            data=dl["data"],
            file_name=dl["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_restore",
        )
