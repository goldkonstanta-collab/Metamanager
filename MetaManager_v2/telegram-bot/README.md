# MetaManager Telegram Bot (@gidroservbot)

Бот нужен только для одной задачи: выдать пользователю его `chat_id`,
чтобы он вставил его на сайте MetaManager в поле **Telegram chat ID**
и нажал **Запомнить**. После этого все сгенерированные КП и договоры
будут приходить лично в его чат с `@gidroservbot`.

## Команды / кнопки

- `/start` — приветствие и кнопка **🔑 Получить ключ**.
- `/key`, `/id` — прислать `chat_id` текстом (удобно копировать на телефоне).
- Кнопка **🔑 Получить ключ** — то же самое, но в одно нажатие.

## Быстрый запуск локально

```bash
cd telegram-bot
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt

export TELEGRAM_BOT_TOKEN="<токен @gidroservbot>"   # Linux/macOS
# Windows PowerShell:
# $env:TELEGRAM_BOT_TOKEN = "<токен @gidroservbot>"

python bot.py
```

Бот работает через long polling — никаких публичных URL/вебхуков не требуется.
Достаточно, чтобы процесс `python bot.py` был запущен (локально, на сервере,
в Docker или в виде Render/Railway Background Worker).

## Как пользоваться (для менеджеров)

1. Открыть [t.me/gidroservbot](https://t.me/gidroservbot).
2. Нажать **Start** (команда `/start`).
3. Нажать кнопку **🔑 Получить ключ** — бот пришлёт число, например `779238503`.
4. Открыть сайт MetaManager → блок **Telegram** → вставить ключ → **Запомнить**.
5. Всё: следующие КП и договоры будут приходить вам в этот чат.

## Как добавить кнопку «Получить ключ» в бота (для владельца @gidroservbot)

Кнопка реализована внутри `bot.py` как inline-кнопка под приветственным
сообщением (`InlineKeyboardButton("🔑 Получить ключ", callback_data="get_key")`).
Она появляется автоматически после отправки пользователем `/start`.

Дополнительно удобно прописать команды в меню Telegram (синяя кнопка «Меню»
слева от поля ввода):

1. Откройте чат с [@BotFather](https://t.me/BotFather).
2. `/setcommands` → выбрать `@gidroservbot`.
3. Вставить:

```
start - Приветствие и кнопка "Получить ключ"
key - Прислать ваш chat ID
id - Прислать ваш chat ID
help - Помощь
```

4. После этого у пользователей появится меню с этими командами,
   а под `/start` всегда будет inline-кнопка **🔑 Получить ключ**.

## Деплой (варианты)

- **Локально / на рабочем ПК** — просто держать `python bot.py` запущенным.
- **Render (Background Worker)** — `pip install -r requirements.txt` + команда
  запуска `python bot.py`, переменная окружения `TELEGRAM_BOT_TOKEN`.
- **Docker** — любой минимальный образ `python:3.11-slim`, копируем файлы,
  ставим зависимости, `CMD ["python", "bot.py"]`.
