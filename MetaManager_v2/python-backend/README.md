# Python Backend For Netlify MVP

Этот сервис использует оригинальные генераторы проекта:

- `generator.py` (КП)
- `contract_generator.py` (договор)

и отправляет итоговые документы в Telegram.

## Запуск локально (Windows)

```bash
cd python-backend
pip install -r requirements.txt
set TELEGRAM_BOT_TOKEN=YOUR_TOKEN
set TELEGRAM_CHAT_ID=YOUR_CHAT_ID
set CHECKO_API_KEY=YOUR_CHECKO_KEY
uvicorn app:app --host 0.0.0.0 --port 8000
```

Проверка:

- `GET http://localhost:8000/health`

## Endpoints

- `POST /generate/kp`
- `POST /generate/contract`
- `GET /lookup/company?inn=...` (проксируется с Netlify как `/api/lookup/company`)
- `GET /lookup/bank?bic=...` (проксируется с Netlify как `/api/lookup/bank`)

Формат payload приходит из `netlify-mvp/public/app.js`.

## PDF на Linux (Render и т.п.)

Генератор пытается конвертировать `.docx` в `.pdf` через LibreOffice, если конвертер недоступен — поле `pdf` в ответе будет `null`.

Для Render рекомендуется Docker-деплой с LibreOffice.

- Если build context — **корень этого проекта** (рядом `generator.py`), используйте `Dockerfile.backend`.
- Если на GitHub проект лежит внутри папки `MetaManager_v2/`, а build context — **корень монорепозитория**, используйте `Dockerfile.render-backend`.

Пример переменных на backend:

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`
- `CHECKO_API_KEY`
