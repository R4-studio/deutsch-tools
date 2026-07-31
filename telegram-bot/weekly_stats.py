"""
weekly_stats.py — вспомогательный скрипт для еженедельного топа: читает
logs/answers.jsonl и печатает таблицу по пользователям за последние N дней
(по умолчанию 7). Дальше уже руками оформляется в картинку.

Запуск:
    python weekly_stats.py          # за последние 7 дней
    python weekly_stats.py 14       # за последние 14 дней
"""
import json
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone

import config


def main() -> None:
    days = int(sys.argv[1]) if len(sys.argv) > 1 else 7
    since = datetime.now(timezone.utc) - timedelta(days=days)

    if not config.ANSWERS_LOG_PATH.exists():
        print("Лога пока нет — answers.jsonl не найден.")
        return

    stats = defaultdict(lambda: {"total": 0, "correct": 0, "name": ""})
    with config.ANSWERS_LOG_PATH.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            entry = json.loads(line)
            ts = datetime.fromisoformat(entry["ts"])
            if ts < since:
                continue
            uid = entry["user_id"]
            stats[uid]["total"] += 1
            stats[uid]["correct"] += int(entry["correct"])
            stats[uid]["name"] = entry.get("username") or entry.get("first_name") or str(uid)

    if not stats:
        print(f"За последние {days} дн. ответов нет.")
        return

    rows = sorted(stats.items(), key=lambda kv: (-kv[1]["correct"], -kv[1]["total"]))
    print(f"Топ за последние {days} дн.:\n")
    print(f"{'#':<3} {'Игрок':<20} {'Верно':<7} {'Всего':<7} {'%':<5}")
    for i, (uid, s) in enumerate(rows, 1):
        pct = round(100 * s["correct"] / s["total"]) if s["total"] else 0
        print(f"{i:<3} {s['name']:<20} {s['correct']:<7} {s['total']:<7} {pct:<5}")


if __name__ == "__main__":
    main()
