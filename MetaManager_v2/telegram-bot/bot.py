"""
MetaManager Telegram bot (@gidroservbot).

Задача бота: выдать пользователю его `chat_id`, который он потом
вставит на сайте MetaManager в поле "Telegram chat ID" и нажмёт
кнопку "Запомнить". После этого все сгенерированные КП и договоры
будут приходить лично ему.

Поддерживаемые действия:
- /start           — приветствие + inline-кнопка "Получить ключ".
- /key, /id        — прислать chat_id текстом (удобно копировать).
- Кнопка "Получить ключ" (callback) — бот отвечает тем же chat_id.

Запуск:
  1. pip install -r requirements.txt
  2. setx TELEGRAM_BOT_TOKEN "<токен @gidroservbot>"
     (или положить токен в переменные окружения процесса)
  3. python bot.py

Бот использует long polling, т.е. никаких публичных URL и вебхуков
не требуется — достаточно компьютера/сервера с выходом в интернет.
"""

from __future__ import annotations

import logging
import os
import sys

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("metamanager-bot")


WELCOME_TEXT = (
    "Привет! Я бот MetaManager.\n\n"
    "Нажмите кнопку <b>Получить ключ</b> ниже — я пришлю ваш "
    "<code>chat_id</code>. Его нужно вставить на сайте MetaManager в поле "
    "<b>Telegram chat ID</b> и нажать <b>Запомнить</b>.\n\n"
    "После этого все сгенерированные КП и договоры будут приходить в этот чат."
)


def _main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("🔑 Получить ключ", callback_data="get_key")]]
    )


def _format_key_message(chat_id: int) -> str:
    return (
        "Ваш ключ (chat ID):\n"
        f"<code>{chat_id}</code>\n\n"
        "Скопируйте его, откройте сайт MetaManager, вставьте в поле "
        "<b>Telegram chat ID</b> и нажмите <b>Запомнить</b>."
    )


async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.effective_chat:
        return
    await update.effective_chat.send_message(
        WELCOME_TEXT,
        parse_mode=ParseMode.HTML,
        reply_markup=_main_keyboard(),
    )


async def key_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat = update.effective_chat
    if not chat:
        return
    await chat.send_message(
        _format_key_message(chat.id),
        parse_mode=ParseMode.HTML,
        reply_markup=_main_keyboard(),
    )


async def get_key_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return
    await query.answer("Ключ готов")
    chat = update.effective_chat
    if not chat:
        return
    await chat.send_message(
        _format_key_message(chat.id),
        parse_mode=ParseMode.HTML,
        reply_markup=_main_keyboard(),
    )


def build_app(token: str) -> Application:
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("help", start_cmd))
    app.add_handler(CommandHandler("key", key_cmd))
    app.add_handler(CommandHandler("id", key_cmd))
    app.add_handler(CallbackQueryHandler(get_key_callback, pattern=r"^get_key$"))
    return app


def main() -> None:
    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip()
    if not token:
        print(
            "ERROR: переменная окружения TELEGRAM_BOT_TOKEN не задана.\n"
            "Задайте токен бота @gidroservbot и запустите снова.",
            file=sys.stderr,
        )
        sys.exit(1)

    app = build_app(token)
    logger.info("Bot is starting (long polling)…")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
