"""Конфиг бота — читает переменные окружения (см. .env.example)."""
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

BOT_TOKEN = os.environ["BOT_TOKEN"]
CHANNEL_ID = os.environ["CHANNEL_ID"]  # напр. @deutsch_tools или -1001234567890
SITE_URL = os.environ.get("SITE_URL", "https://r4-studio.github.io/deutsch-tools")

# Шанс добавить в explanation ссылку на сайт с призывом подтянуть знания (0..1)
LINK_PROBABILITY = float(os.environ.get("LINK_PROBABILITY", "0.5"))

STATE_DIR = Path(__file__).resolve().parent / "state"
LOGS_DIR = Path(__file__).resolve().parent / "logs"
POSTED_POLLS_PATH = STATE_DIR / "posted_polls.json"
ANSWERS_LOG_PATH = LOGS_DIR / "answers.jsonl"
