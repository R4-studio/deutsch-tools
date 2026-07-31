"""
post_quiz.py — разовый скрипт: генерирует "слово дня" и постит его в Telegram-канал
обычным сообщением с inline-кнопками-вариантами ответа. Рассчитан на запуск
по Планировщику/cron раз в сутки.

Почему не нативный Quiz: Telegram не разрешает не-анонимные опросы в каналах,
а у анонимных опросов бот не получает, кто именно ответил — только агрегированные
цифры. Поэтому вариант "квиза" тут — обычное сообщение с InlineKeyboardMarkup:
клик по кнопке всегда приходит боту вместе с user_id, независимо от анонимности.

Запуск вручную:
    python post_quiz.py

Только постит и выходит. Сами клики по кнопкам обрабатывает bot_listener.py —
он должен работать постоянно (см. README.md), иначе кнопки будут "висеть"
до следующего запуска бота.
"""
import asyncio
import json
import logging
import uuid

from telegram import Bot, InlineKeyboardButton, InlineKeyboardMarkup

import config
from quiz_data import build_quiz, load_vocab

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("post_quiz")


def _load_posted_polls() -> dict:
    if not config.POSTED_POLLS_PATH.exists():
        return {}
    try:
        return json.loads(config.POSTED_POLLS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _save_posted_polls(polls: dict) -> None:
    config.STATE_DIR.mkdir(parents=True, exist_ok=True)
    config.POSTED_POLLS_PATH.write_text(json.dumps(polls, ensure_ascii=False, indent=2), encoding="utf-8")


def _render_text(question: str, answered: int = 0, correct: int = 0) -> str:
    text = f"🇩🇪 <b>Слово дня</b>\n\n{question}"
    if answered:
        text += f"\n\n👥 Ответили: {answered} · ✅ верно: {correct}"
    return text


async def main():
    vocab = load_vocab()
    log.info("Загружено %d слов из data.js", len(vocab))

    quiz = build_quiz(vocab)
    word = quiz["word"]
    poll_id = uuid.uuid4().hex[:10]

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton(opt, callback_data=f"ans:{poll_id}:{i}")]
        for i, opt in enumerate(quiz["options"])
    ])
    text = _render_text(quiz["question"])

    bot = Bot(token=config.BOT_TOKEN)
    message = await bot.send_message(
        chat_id=config.CHANNEL_ID,
        text=text,
        parse_mode="HTML",
        reply_markup=keyboard,
    )

    polls = _load_posted_polls()
    polls[poll_id] = {
        "chat_id": message.chat_id,
        "message_id": message.message_id,
        "question": quiz["question"],
        "word": word,
        "options": quiz["options"],
        "correct_option_id": quiz["correct_option_id"],
        "answered_user_ids": [],
        "stats": {"answered": 0, "correct": 0},
    }
    _save_posted_polls(polls)
    log.info("Запостил «%s» (poll_id=%s)", word["de"], poll_id)


if __name__ == "__main__":
    asyncio.run(main())
