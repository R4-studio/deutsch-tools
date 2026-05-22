"""
Миграция data.js: переразметка осиротевших topics в QUESTIONS.

Что делает:
1. Проходит по всем QUESTIONS
2. Перенаправляет topics-сирот (типа `trainer:basic`) в правильные
   (`trainer:verbs:basic`) на основе известного маппинга
3. Сохраняет data.js обратно
"""
import re
from pathlib import Path

DATA_JS = Path("/home/claude/repo/data.js")

# Карта переразметки: что было → что должно стать
# Базируется на BLOCKS структуре
TOPIC_REMAP = {
    # Глаголы (сабблоки)
    "trainer:basic":         "trainer:verbs:basic",
    "trainer:modal":         "trainer:verbs:modal",
    "trainer:movement":      "trainer:verbs:movement",
    "trainer:food":          "trainer:verbs:food",
    "trainer:home":          "trainer:verbs:home",
    "trainer:communication": "trainer:verbs:communication",
    "trainer:study":         "trainer:verbs:study",
    "trainer:leisure":       "trainer:verbs:leisure",
    "trainer:impersonal":    "trainer:verbs:impersonal",
    # Существительные (часть сабблоков)
    "trainer:family":        "trainer:nouns:family",
    "trainer:misc":          "trainer:nouns:misc",
    # Топ-уровневые блоки остаются как есть:
    #   trainer:adj, trainer:adv, trainer:mestoim, trainer:neu,
    #   trainer:nums, trainer:phrases, trainer:rules, trainer:sounds,
    #   trainer:termin, trainer:unregel
    # — это правильные topics (блоки без сабблоков или особые kind-блоки)
}

print(f"Читаю {DATA_JS}...")
content = DATA_JS.read_text(encoding="utf-8")

# Считаем сколько вопросов изменим
remapped = 0
for old_topic, new_topic in TOPIC_REMAP.items():
    # Находим все вхождения `topic: "trainer:basic"` и заменяем
    pattern = re.compile(r'topic:\s*"' + re.escape(old_topic) + r'"')
    matches = pattern.findall(content)
    if matches:
        count = len(matches)
        content = pattern.sub(f'topic: "{new_topic}"', content)
        print(f"  {old_topic} → {new_topic}: {count} вопросов")
        remapped += count

print(f"\n✓ Переразмечено: {remapped} вопросов")

DATA_JS.write_text(content, encoding="utf-8")
print(f"✓ Сохранено: {DATA_JS}")
