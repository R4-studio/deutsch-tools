"""
quiz_data.py — читает VOCAB прямо из ../data.js (без дублирования данных)
и умеет собрать один quiz-вопрос: слово + правильный перевод + 3 дистрактора
из той же темы (topics).

data.js — не JSON (ключи объектов не в кавычках: `{ id: "u001", de: "..." }`),
поэтому парсим регуляркой: находим массив VOCAB, дописываем кавычки к ключам,
убираем висячие запятые перед `]`/`}` — и дальше это уже валидный JSON.
"""
import json
import random
import re
from pathlib import Path

DATA_JS_PATH = Path(__file__).resolve().parent.parent / "data.js"
RECENT_WORDS_PATH = Path(__file__).resolve().parent / "state" / "recent_words.json"
RECENT_WORDS_KEEP = 60  # не повторять слово, пока не наберётся столько других

_KEY_RE = re.compile(r'([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:')
_TRAILING_COMMA_RE = re.compile(r',(\s*[\]}])')
_LINE_COMMENT_RE = re.compile(r'^\s*//.*$', re.MULTILINE)

# Похожесть по звучанию — тот же алгоритм, что и в trainer.html
# (normDeForSim/trigramsOf/trigramSimilarity, см. pickDistractors в trainer.html),
# чтобы дистракторы в боте подбирались так же, как в тестах на сайте.
_LEADING_ARTICLE_RE = re.compile(
    r'^(der|die|das|den|dem|des|ein|eine|einen|einem|einer|eines)\s+', re.IGNORECASE
)
DIST_SIM_THRESHOLD = 0.2


def _norm_de_for_sim(s: str) -> str:
    if not s:
        return ""
    return _LEADING_ARTICLE_RE.sub("", s.lower()).strip()


def _trigrams(s: str) -> set:
    if not s or len(s) < 3:
        return set()
    return {s[i:i + 3] for i in range(len(s) - 2)}


def _trigram_similarity(de1: str, de2: str) -> float:
    a, b = _trigrams(_norm_de_for_sim(de1)), _trigrams(_norm_de_for_sim(de2))
    if not a or not b:
        return 0.0
    return len(a & b) / max(len(a), len(b))


def _extract_const_array(js_text: str, const_name: str) -> list:
    """Достаёт `const NAME = [ ... ];` из текста data.js и парсит как JSON."""
    m = re.search(rf'const {const_name} = (\[.*?\n\]);', js_text, re.S)
    if not m:
        raise RuntimeError(f"Не нашёл `const {const_name} = [...]` в data.js")
    raw = m.group(1)
    raw = _LINE_COMMENT_RE.sub('', raw)
    raw = _KEY_RE.sub(r'\1"\2":', raw)
    raw = _TRAILING_COMMA_RE.sub(r'\1', raw)
    return json.loads(raw)


def load_vocab() -> list:
    """Возвращает список слов VOCAB (только те, у кого есть de+ru — вопрос строится на них)."""
    js_text = DATA_JS_PATH.read_text(encoding="utf-8")
    vocab = _extract_const_array(js_text, "VOCAB")
    return [v for v in vocab if v.get("de") and v.get("ru")]


def _load_recent_ids() -> list:
    if not RECENT_WORDS_PATH.exists():
        return []
    try:
        return json.loads(RECENT_WORDS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []


def _save_recent_ids(ids: list) -> None:
    RECENT_WORDS_PATH.parent.mkdir(parents=True, exist_ok=True)
    RECENT_WORDS_PATH.write_text(json.dumps(ids[-RECENT_WORDS_KEEP:], ensure_ascii=False), encoding="utf-8")


def pick_word(vocab: list) -> dict:
    """Случайное слово, по возможности не из недавно использованных."""
    recent = set(_load_recent_ids())
    pool = [v for v in vocab if v["id"] not in recent] or vocab
    word = random.choice(pool)
    ids = _load_recent_ids()
    ids.append(word["id"])
    _save_recent_ids(ids)
    return word


def pick_distractors(vocab: list, word: dict, n: int = 3) -> list:
    """Каскад — как pickDistractors в trainer.html (genL1_translate_DE_RU):
    1) та же часть речи (pos) + общая тема
    2) та же часть речи + похожесть по звучанию (триграммы de) >= порога
    3) любая часть речи + похожесть по звучанию >= порога
    4) та же часть речи, рандом (фолбэк при нехватке данных)
    5) любая часть речи, рандом (крайний фолбэк — в trainer.html его нет,
       там генератор вопроса просто возвращает null; боту всегда нужно 3
       варианта, поэтому подстраховываемся)
    Дедуп по русскому переводу — чтобы не было двух одинаковых вариантов ответа."""
    word_topics = set(word.get("topics") or [])
    correct_ru = word["ru"].strip().lower()

    taken, used_ru = [], {correct_ru}

    def try_add(v):
        if v["id"] == word["id"]:
            return
        ru = v["ru"].strip().lower()
        if not ru or ru in used_ru:
            return
        taken.append(v)
        used_ru.add(ru)

    def by_similarity(pool):
        scored = [(v, _trigram_similarity(v["de"], word["de"])) for v in pool]
        scored = [(v, s) for v, s in scored if s >= DIST_SIM_THRESHOLD]
        scored.sort(key=lambda x: -x[1])
        return [v for v, _ in scored]

    if word_topics:
        level1 = [
            v for v in vocab
            if v["id"] != word["id"] and v.get("pos") == word.get("pos")
            and word_topics & set(v.get("topics") or [])
        ]
        random.shuffle(level1)
        for v in level1:
            if len(taken) >= n:
                break
            try_add(v)

    if len(taken) < n:
        level2 = [v for v in vocab if v["id"] != word["id"] and v.get("pos") == word.get("pos")]
        for v in by_similarity(level2):
            if len(taken) >= n:
                break
            try_add(v)

    if len(taken) < n:
        level3 = [v for v in vocab if v["id"] != word["id"]]
        for v in by_similarity(level3):
            if len(taken) >= n:
                break
            try_add(v)

    if len(taken) < n:
        level4 = [v for v in vocab if v["id"] != word["id"] and v.get("pos") == word.get("pos")]
        random.shuffle(level4)
        for v in level4:
            if len(taken) >= n:
                break
            try_add(v)

    if len(taken) < n:
        level5 = [v for v in vocab if v["id"] != word["id"]]
        random.shuffle(level5)
        for v in level5:
            if len(taken) >= n:
                break
            try_add(v)

    return taken[:n]


def build_quiz(vocab: list) -> dict:
    """Возвращает {question, options, correct_option_id, word} — готово для send_poll."""
    word = pick_word(vocab)
    distractors = pick_distractors(vocab, word)
    options = [word["ru"]] + [d["ru"] for d in distractors]
    # de-dup на всякий случай (не должно срабатывать благодаря used_ru, но не доверяем слепо)
    seen, uniq = set(), []
    for o in options:
        key = o.strip().lower()
        if key not in seen:
            seen.add(key)
            uniq.append(o)
    options = uniq
    order = list(range(len(options)))
    random.shuffle(order)
    shuffled = [options[i] for i in order]
    correct_option_id = shuffled.index(word["ru"])
    return {
        "question": f'Как переводится «{word["de"]}»?',
        "options": shuffled,
        "correct_option_id": correct_option_id,
        "word": word,
    }


if __name__ == "__main__":
    # Быстрый прогон парсера + генератора на реальных данных — для проверки руками:
    #   python quiz_data.py
    v = load_vocab()
    print(f"VOCAB: {len(v)} слов с de+ru")
    for _ in range(3):
        q = build_quiz(v)
        print("\n" + q["question"])
        for i, opt in enumerate(q["options"]):
            mark = " <-- верно" if i == q["correct_option_id"] else ""
            print(f"  {i}. {opt}{mark}")
