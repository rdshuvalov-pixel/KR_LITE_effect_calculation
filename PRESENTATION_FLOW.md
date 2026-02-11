# Потоковая схема: от данных до презентации

```mermaid
flowchart TB
    subgraph input["1. Вход"]
        A[Новый Excel с данными<br/>S-Market.xlsx, эффект_ДДММГГГГ.xlsx]
    end

    subgraph app["2. Приложение Streamlit"]
        B[Загрузка файла в app.py]
        C[Расчёт через EffectCalculator]
        D[Результаты + summary]
        
        B --> C --> D
    end

    subgraph export["3. Экспорт данных расчёта"]
        E[stats_data.json<br/>переоценки, эффекты, сценарии,<br/>sourceFileName, generatedAt]
        F[presentation_data.json<br/>пример позиции: weeks,<br/>revenuePosition, revenueControl]
        
        D --> E
        D --> F
    end

    subgraph presentation["4. Генерация презентации"]
        G[Новый HTML-файл<br/>presentation_YYYYMMDD_HHMM.html<br/>или по имени файла]
        H[Самодостаточная презентация<br/>данные встроены / рядом]
        
        E --> G
        F --> G
        G --> H
    end

    A --> B

    style input fill:#e8f5e9
    style app fill:#e3f2fd
    style export fill:#fff3e0
    style presentation fill:#f3e5f5
```

## Шаги потока

| № | Действие | Результат |
|---|----------|-----------|
| 1 | Берёшь новый Excel с данными | Файл готов к загрузке |
| 2 | Загружаешь в Streamlit-приложение, настраиваешь параметры, нажимаешь «Применить» | `calc.calculate()` → `results`, `summary` |
| 3 | Данные экспортируются в JSON | `stats_data.json` (общая статистика) + `presentation_data.json` (пример позиции) |
| 4 | Формируется новый HTML-файл | `presentation_<имя_замера>_<дата>.html` — презентация только под этот расчёт |

## Текущее vs целевое

- **Сейчас**: `stats_data.json` и `presentation_data.json` обновляются скриптами `generate_stats.py` и `extract_data.py` вручную. `presentation.html` читает их через `fetch()` при открытии.
- **Целевое**: из приложения после расчёта — экспорт JSON и генерация отдельного HTML-файла под текущий расчёт (файл самодостаточен или всегда идёт со своим набором JSON).
