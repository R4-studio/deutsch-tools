"""
bot_listener.py — постоянно работающий процесс: слушает клики по кнопкам
под постами из post_quiz.py, отвечает всплывающей подсказкой (верно/неверно),
пишет каждый ответ в logs/answers.jsonl и обновляет счётчик под сообщением.

Это не разовый скрипт — его нужно держать запущенным всё время (см. README.md),
иначе клики по кнопкам будут просто крутиться у пользователя до перезапуска бота.

Запуск:
    python bot_listener.py
"""
import json
import logging
from datetime import datetime, timezone

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.error import BadRequest
from telegram.ext import Application, CallbackQueryHandler, ContextTypes

import config
from explanations import build_explanation

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("bot_listener")


def _load_polls() -> dict:
    if not config.POSTED_POLLS_PATH.exists():
        return {}
    try:
        return json.loads(config.POSTED_POLLS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _save_polls(polls: dict) -> None:
    config.STATE_DIR.mkdir(parents=True, exist_ok=True)
    config.POSTED_POLLS_PATH.write_text(json.dumps(polls, ensure_ascii=False, indent=2), encoding="utf-8")


def _log_answer(entry: dict) -> None:
    config.LOGS_DIR.mkdir(parents=True, exist_ok=True)
    with config.ANSWERS_LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def _render_text(question: str, answered: int, correct: int) -> str:
    text = f"🇩🇪 <b>Слово дня</b>\n\n{question}"
    if answered:
        text += f"\n\n👥 Ответили: {answered} · ✅ верно: {correct}"
    return text


async def _safe_answer(query, text: str) -> None:
    """query.answer() падает с BadRequest, если клик "протух" (бот был офлайн,
    двойной тап и т.п.) — сама попытка ответить пользователю в таком случае
    просто теряется, но остальную обработку (лог, счётчик) это не должно рушить."""
    try:
        await query.answer(text, show_alert=True)
    except BadRequest as e:
        log.warning("Не удалось ответить на callback (вероятно, устарел): %s", e)


async def on_answer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    try:
        _, poll_id, option_idx_raw = query.data.split(":", 2)
        option_idx = int(option_idx_raw)
    except (ValueError, AttributeError):
        await _safe_answer(query, "Что-то не так с кнопкой 🤔")
        return

    polls = _load_polls()
    poll = polls.get(poll_id)
    if not poll:
        await _safe_answer(query, "Этот вопрос уже устарел")
        return

    user = query.from_user
    if user.id in poll["answered_user_ids"]:
        await _safe_answer(query, "Ты уже отвечал на этот вопрос 🙂")
        return

    is_correct = option_idx == poll["correct_option_id"]
    word = poll["word"]

    verdict = "✅ Верно!" if is_correct else f"❌ Неверно. Правильный ответ: {poll['options'][poll['correct_option_id']]}"
    extra = build_explanation(word, config.SITE_URL, config.LINK_PROBABILITY)
    popup = verdict if not extra else f"{verdict}\n\n{extra}"
    await _safe_answer(query, popup[:200])

    poll["answered_user_ids"].append(user.id)
    poll["stats"]["answered"] += 1
    if is_correct:
        poll["stats"]["correct"] += 1
    polls[poll_id] = poll
    _save_polls(polls)

    _log_answer({
        "ts": datetime.now(timezone.utc).isoformat(),
        "poll_id": poll_id,
        "word_id": word["id"],
        "de": word["de"],
        "ru": word["ru"],
        "user_id": user.id,
        "username": user.username,
        "first_name": user.first_name,
        "chosen_option": poll["options"][option_idx],
        "correct": is_correct,
    })

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton(opt, callback_data=f"ans:{poll_id}:{i}")]
        for i, opt in enumerate(poll["options"])
    ])
    try:
        await context.bot.edit_message_text(
            chat_id=poll["chat_id"],
            message_id=poll["message_id"],
            text=_render_text(poll["question"], poll["stats"]["answered"], poll["stats"]["correct"]),
            parse_mode="HTML",
            reply_markup=keyboard,
        )
    except Exception as e:
        log.warning("Не удалось обновить счётчик под сообщением: %s", e)

    log.info("%s (%s) ответил на «%s»: %s", user.username or user.id, user.id, word["de"], "верно" if is_correct else "неверно")


def main() -> None:
    app = Application.builder().token(config.BOT_TOKEN).build()
    app.add_handler(CallbackQueryHandler(on_answer, pattern=r"^ans:"))
    log.info("bot_listener запущен, жду ответов...")
    app.run_polling(allowed_updates=["callback_query"])


if __name__ == "__main__":
    main()
