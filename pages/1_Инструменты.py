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

# --- 1. Объединение данных Энтерсайт ---
with st.expander("Объединение данных Энтерсайт", expanded=False):
    st.caption("Склейка периодов: агрегат за период → подённая разбивка. При пересечении приоритет у более позднего файла.")
    uploaded_entersite = st.file_uploader(
        "Загрузите xlsx с периодом в имени (с_ДД.ММ.ГГГГ_по_ДД.ММ.ГГГГ или с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ)",
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
- **Имя:** `с_ДД.ММ.ГГГГ_по_ДД.ММ.ГГГГ` или `с_ДД_ММ_ГГГГ_по_ДД_ММ_ГГГГ`
- **Лист:** B=Артикул, C=Товар, L=Количество, M=Закуп, N=Розница
- **Выход:** листы Себестоимость и Продажи
        """)

# --- 2. Эталонный файл ---
with st.expander("Эталонный файл", expanded=False):
    etalon_bytes = get_etalon_bytes()
    st.download_button(
        "Скачать etalon_check.xlsx",
        data=etalon_bytes,
        file_name="etalon_check.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_etalon",
    )
    st.markdown("**Как понять результат эталона:**")
    st.markdown("- ✅ **Правильно:** «Обработка завершена успешно», валидных недель > 0, эффект не нулевой.")
    st.markdown("- ❌ **Неправильно:** «сводка пуста», валидных недель 0, все метрики нулевые.")

# --- 3. Восстановление истории себестоимости ---
with st.expander("Восстановление истории себестоимости", expanded=False):
    st.caption("Для каждого товара генерируются строки с датами от 22.09.2025 до первой известной даты. Добавляет лист «Восстановленные данные».")
    uploaded_history = st.file_uploader("Загрузите XLSX с листом Себестоимость", type=["xlsx"], key="restore_history_upload")
    if uploaded_history:
        if st.button("Восстановить историю", key="btn_restore_history"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_history.getvalue())
                tmp_path = tmp.name
            try:
                result = restore_cost_history(tmp_path)
                if result and result.get("rows", 0) > 0:
                    with open(tmp_path, "rb") as f:
                        st.session_state["restore_history_dl"] = {
                            "data": f.read(),
                            "name": f"restored_history_{uploaded_history.name}",
                            "msg": f"Добавлено {result['rows']} строк, {result['products']} товаров",
                        }
                    Path(tmp_path).unlink(missing_ok=True)
                else:
                    Path(tmp_path).unlink(missing_ok=True)
                    st.session_state["restore_history_info"] = result.get("message", "Нет данных для восстановления.")
            except Exception as e:
                Path(tmp_path).unlink(missing_ok=True)
                st.session_state["restore_history_error"] = str(e)
    if st.session_state.get("restore_history_error"):
        st.error(st.session_state.pop("restore_history_error"))
    if st.session_state.get("restore_history_info"):
        st.info(st.session_state.pop("restore_history_info"))
    if st.session_state.get("restore_history_dl"):
        dl = st.session_state["restore_history_dl"]
        st.success(f"Готово: {dl['msg']}")
        st.download_button(
            "📥 Скачать результат",
            data=dl["data"],
            file_name=dl["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_restore_history",
        )

# --- 4. Восстановление себестоимости из продаж ---
with st.expander("Восстановление себестоимости из продаж", expanded=False):
    st.caption("Для каждого product_id из Продаж создаются записи себестоимости на каждую дату (stock=1, cost=1). Объединяет с существующим листом Себестоимость.")
    uploaded_sales = st.file_uploader("Загрузите XLSX с листами Продажи и Себестоимость", type=["xlsx"], key="restore_sales_upload")
    if uploaded_sales:
        if st.button("Восстановить из продаж", key="btn_restore_sales"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_sales.getvalue())
                tmp_path = tmp.name
            try:
                result = restore_cost_from_sales(tmp_path)
                with open(tmp_path, "rb") as f:
                    st.session_state["restore_sales_dl"] = {
                        "data": f.read(),
                        "name": f"restored_from_sales_{uploaded_sales.name}",
                        "msg": f"{result['rows']} строк, {result['products']} товаров",
                    }
                Path(tmp_path).unlink(missing_ok=True)
            except Exception as e:
                Path(tmp_path).unlink(missing_ok=True)
                st.session_state["restore_sales_error"] = str(e)
    if st.session_state.get("restore_sales_error"):
        st.error(st.session_state.pop("restore_sales_error"))
    if st.session_state.get("restore_sales_dl"):
        dl = st.session_state["restore_sales_dl"]
        st.success(f"Готово: {dl['msg']}")
        st.download_button(
            "📥 Скачать результат",
            data=dl["data"],
            file_name=dl["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_restore_sales",
        )
