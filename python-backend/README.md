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
uvicorn app:app --host 0.0.0.0 --port 8000
```

Проверка:

- `GET http://localhost:8000/health`

## Endpoints

- `POST /generate/kp`
- `POST /generate/contract`

Формат payload приходит из `netlify-mvp/public/app.js`.
