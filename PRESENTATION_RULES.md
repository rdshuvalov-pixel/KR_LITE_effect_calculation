# Правила оформления презентации

Оформление презентации по расчёту эффекта переоценки (перенесено из эксперимента).

## Структура слайдов (9 шт.)

| Слайд | Заголовок | Содержание |
|-------|-----------|------------|
| 1 | Титульный | KeepRise, приветствие |
| 2 | Пример позиции | Объект анализа, периоды, карточка товара |
| 3 | Динамика выручки | График ECharts (позиция + магазин) |
| 4 | Расчёт эффекта переоценки | 5 шагов формул |
| 5 | Итоги по позиции | 3 карточки результата |
| 6 | Волны переоценок | Таблица дат |
| 7 | Суммарный эффект, результат по позициям, выручка | Эффект + pie + бар выручки |
| 8 | Охват и сценарии | Доли + 3 сценария |
| 9 | Ограничения и фильтрация | 4 карточки методологии + params внизу |

## CSS-правила

- **Формат:** 16:9 (1265×711px), `display: flex; flex-direction: column`
- **Заголовки:** `.slide-header`, `.slide > h2` — `align-self: flex-start` (всегда верхний левый угол)
- **Слайды:** `data-slide="N"`, у каждого `.slide-header` с `.slide-num` и `.slide-title-text`
- **Карточки:** `.card-3d` — 3D-эффект на экране, при печати — `transform: none`, `box-shadow: none`, только `border`

## Печать (@media print)

- `@page { size: A4 landscape; margin: 0 }`
- `.slide`: `height: 210mm`, `page-break-before: always` (кроме первого)
- Блоки: без теней, `border: 1px solid rgba(255,255,255,0.12)`
- График (слайд 3): `height: 280px`, `max-width: 100%`
- Слайд 8: `.scenario-box` — нейтральный фон `rgba(0,0,0,0.15)` (без цветного)

## Инъекция данных

`presentation_builder.generate_html()` вставляет `window.PRESENTATION_DATA` и `window.STATS_DATA` в начало скрипта с `function fmtWeek`. Шаблон должен содержать строку `  <script>\n    function fmtWeek` для замены.

## PDF (Playwright)

- Движок: Chromium (WebKit не поддерживает PDF)
- Ожидание: `document.fonts.ready`, 2.5 с для ECharts, 500 мс после `emulate_media("print")`
- Перед PDF: `echarts.getInstanceByDom().resize()` для chart и pieChart
- `page.pdf(margin=0, print_background=True)`

## Файлы

- `presentation.html` — шаблон
- `pdf_generator.py` — `export_html_to_pdf(html_path, pdf_path)`
- `presentation_builder.py` — `generate_html()`, `save_presentation_and_manage_history()`
