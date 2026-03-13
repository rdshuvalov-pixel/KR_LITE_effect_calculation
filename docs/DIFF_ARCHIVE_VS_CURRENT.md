# Сравнение: архив vs текущая версия (после отката)

**Архив:** `Замер_результата_пилота_KeepRise_Lite_back_main (1).zip` (25.02.2026)  
**Текущая:** GitHub b31ae65 (19.02.2026) + локальные файлы

---

## Что есть в АРХИВЕ, но НЕТ в текущей версии

### 1. calculator.py
- **Категория «Без изменений»** — отдельная маска `unchanged_mask` и `unchanged_df` (раньше decline = всё, что не growth)
- **product_ids** в `growth_stats`, `decline_stats`, `unchanged_stats` — для фильтрации и синхронизации
- **unchanged_stats** в возвращаемом summary

### 2. app.py
- **Блок «Без изменений»** — 3 колонки (Рост / Без изменений / Падение) вместо 2
- **Синхронизация товара** между вкладками «Анализ товара» и «Активация цен» (`synced_product_id`)
- **Фильтр «Без изменений»** в списке товаров
- **Фильтр по товару** во вкладке «Активация цен»
- **included_pids** считается из `product_ids` в summary, а не из `results['product_id']`
- Иконка ➡️ для позиций без изменений

---

## Что есть в ТЕКУЩЕЙ версии, но НЕТ в архиве

### 1. Презентация и PDF
- **presentation_builder.py** — генерация HTML-презентации
- **pdf_generator.py** — экспорт в PDF через Playwright
- **presentation/presentation.html** — шаблон
- Автогенерация презентации при расчёте
- Кнопка «Скачать презентацию в PDF»

### 2. report_generator.py
- **Обработка NaT** — `val.strftime(...) if not pd.isna(val) else "-"` (в архиве будет ошибка на пустых датах)

### 3. Инфраструктура
- **packages.txt** — зависимости для Streamlit Cloud (Playwright)
- **DEPLOY.md** — инструкции деплоя
- **.streamlit/config.toml**
- **PRESENTATION_FLOW.md**, **PRESENTATION_RULES.md**

---

## Итоговая матрица

| Функция                         | Архив | Текущая |
|---------------------------------|-------|---------|
| Категория «Без изменений»       | ✅    | ❌      |
| product_ids в summary          | ✅    | ❌      |
| Синхронизация товара между вкладками | ✅ | ❌      |
| Фильтр по товару в Активации   | ✅    | ❌      |
| Презентация HTML/PDF           | ❌    | ✅      |
| Обработка NaT в Word-отчёте    | ❌    | ✅      |
| Streamlit Cloud / packages.txt | ❌    | ✅      |

---

## Рекомендации по объединению

Чтобы совместить обе версии:
1. Взять из архива: `unchanged_stats`, `product_ids`, синхронизацию товара, фильтры в app.
2. Сохранить из текущей: `presentation_builder`, `pdf_generator`, NaT в report_generator, `packages.txt`.
