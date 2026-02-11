# KeepRise Lite: A/B Test Calculator

Калькулятор эффективности ценообразования для расчёта эффекта переоценки. Анализ по методу Test vs Control с учётом дотестового периода, активации цен и фильтрации по остаткам.

## Возможности

- Загрузка Excel с продажами, тестовыми ценами и себестоимостью
- Расчёт эффекта переоценки по выручке и прибыли
- Экспорт в Excel и Word
- Генерация HTML-презентации по расчёту (последние 3 сохраняются)

## Локальный запуск

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Деплой

### Streamlit Community Cloud (рекомендуется)

1. Форкните репозиторий или создайте свой на GitHub
2. Перейдите на [share.streamlit.io](https://share.streamlit.io)
3. Подключите GitHub и выберите репозиторий
4. **Main file path:** `app.py`
5. Deploy

### Vercel

Vercel не поддерживает запуск Streamlit (Python-сервер). Варианты:

- **Вариант A:** Деплоить основной сервис на [Streamlit Cloud](https://share.streamlit.io), а на Vercel — лендинг/редирект на него
- **Вариант B:** Переписать backend на Next.js API Routes + фронт на React — трудоёмко

## Структура проекта

```
app.py              # Основное Streamlit-приложение
calculator.py       # Логика расчёта эффекта
report_generator.py # Генерация Word-отчёта
presentation_builder.py  # Генерация HTML-презентации
presentation.html   # Шаблон презентации
```

## Формат данных

Excel с листами: **Тестовые цены**, **Продажи**, **Себестоимость**. Подробнее в `METHODOLOGY.md`.
