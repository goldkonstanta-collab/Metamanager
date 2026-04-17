# Netlify MVP (mobile + Telegram)

Этот MVP разворачивается отдельно от desktop-приложения и дает:

- мобильную веб-форму для КП/Договора;
- Netlify Functions для обработки заявок;
- отправку результата в Telegram через Python backend.

## Что уже умеет MVP

- Открывается в браузере на телефоне.
- При отправке формы вызывает serverless функцию Netlify.
- Передает данные в Python backend.
- Python backend генерирует документы тем же кодом, что и desktop-версия
  (`generator.py`, `contract_generator.py`) и отправляет их в Telegram.

## Важно

Netlify Functions сами по себе не подходят для запуска вашего Python-генератора "как на ПК".
Поэтому используется связка:

`Netlify frontend -> Netlify function proxy -> Python backend (FastAPI) -> Telegram`

## Развертывание в Netlify

1. В Netlify создайте новый сайт из репозитория.
2. В настройках Build:
   - **Base directory**: `netlify-mvp`
   - **Build command**: оставить пустым
   - **Publish directory**: `public`
3. Добавьте переменные окружения:
   - `TELEGRAM_BOT_TOKEN`
   - `TELEGRAM_CHAT_ID`
   - `BACKEND_URL` (URL Python backend, пример: `https://your-backend.onrender.com`)
4. Deploy.

## Локальный запуск

```bash
cd netlify-mvp
npm install
npm run dev
```

После этого откройте локальный URL от Netlify CLI.
Локальные переменные можно задать в файле `.env` (в MVP уже есть шаблон `.env.example`).

## Telegram настройки

- `TELEGRAM_BOT_TOKEN` — токен от BotFather.
- `TELEGRAM_CHAT_ID` — ваш chat_id.
- `BACKEND_URL` — URL backend сервиса.

## Python backend (обязательно для Word/PDF как в ПК версии)

В проекте есть папка `python-backend`.

Локально:

```bash
cd python-backend
pip install -r requirements.txt
set TELEGRAM_BOT_TOKEN=...
set TELEGRAM_CHAT_ID=...
uvicorn app:app --host 0.0.0.0 --port 8000
```

Для продакшена задеплойте `python-backend` на Render/Railway/VPS
и укажите его адрес в `BACKEND_URL` в Netlify.
