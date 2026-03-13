# Деплой на GitHub и в облако

## 1. GitHub (отдельный репозиторий)

Папка сейчас внутри большого репо Cursor. Для чистого деплоя — отдельный репо только для калькулятора.

### Вариант A: Новый репо из этой папки

```bash
cd "/Users/luqy/Documents/Cursor/Приложение для рассчета эффекта"

# Создать новый git внутри папки (будет вложенный репо)
git init
git add .
git commit -m "Initial: KeepRise Lite A/B Test Calculator"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/keeprise-lite-calculator.git
git push -u origin main
```

На GitHub: New repository → `keeprise-lite-calculator` (пустой, без README).

### Вариант B: Только эта папка из текущего репо

Если хочешь оставить всё в родительском репо Cursor:

```bash
cd /Users/luqy/Documents/Cursor
git add "Приложение для рассчета эффекта/"
git commit -m "Add KeepRise Lite Calculator"
git push
```

Тогда в GitHub будет весь Cursor, включая эту папку.

---

## 2. Деплой приложения

### Streamlit Community Cloud (предпочтительно)

1. [share.streamlit.io](https://share.streamlit.io) → Sign in с GitHub
2. New app → выбери репозиторий `keeprise-lite-calculator`
3. **Main file path:** `app.py`
4. **App URL (опционально):** оставь автоматический
5. Deploy

Через 1–2 минуты приложение будет доступно по ссылке вида  
`https://keeprise-lite-calculator-xxx.streamlit.app`.

### Vercel

Vercel не умеет запускать Streamlit (нужен Python-сервер).

Имеет смысл так:

- **Деплой приложения** — на Streamlit Community Cloud
- **Vercel** — только лендинг или статика, которая ведёт на Streamlit

Пример для лендинга на Vercel:

1. Создать в репо папку `vercel-landing/` с `index.html`:

```html
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>KeepRise Lite</title></head>
<body>
  <h1>KeepRise Lite: Калькулятор эффекта</h1>
  <p><a href="https://YOUR-APP.streamlit.app">Открыть приложение</a></p>
</body>
</html>
```

2. В корне — `vercel.json`:

```json
{
  "rewrites": [{ "source": "/(.*)", "destination": "/vercel-landing/index.html" }]
}
```

3. Подключить репо к Vercel и задеплоить.

---

## 3. Переменные окружения

Для Streamlit Cloud обычно ничего не нужно. Если появятся секреты — задать в Settings → Secrets.
