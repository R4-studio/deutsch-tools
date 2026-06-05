"""
update_data.py — собирает data.js из database.xlsx

Запуск: python update_data.py

Что делает:
1. Читает все листы database.xlsx (nouns, verbs, adjectives, adverbs, pronouns,
   numbers, phrases, sounds, terms, rules, questions) + лист Settings-XLSX.
2. Для verbs: regel без форм → автогенератор; regel с формами → ручные.
   Для unregel → берёт формы из ячеек.
3. Сохраняет BLOCKS, TOPIC_TITLES, TAB_TITLES, PHRASE_UNITS, SENTENCE_TEMPLATES
   из существующего data.js (не пересоздаёт их — это код фронта, не контент).
4. Settings-XLSX → const TAXONOMY (двухуровневая иерархия domen → group).
5. Финальный файл data.js пишется заново.

СХЕМА (новая, domen/group):
- topic заменён на пару domen + group на всех контентных листах.
- Settings-XLSX — мастер-список валидных пар (page, domen, group, …).
- Глаголы: de / er_sie_es / sie_Sie + флаги separable, prefix, reflexive,
  impersonal, case, praeteritum.
- Поля example_de / example_ru + quiz_use (годится ли ПРИМЕР для теста-сборки).

ПРИНЦИПЫ:
- Читаются ВСЕ колонки схемы (даже сейчас пустые). Маппинг объявлен в SCHEMA —
  заполнишь колонку в xlsx, и она поедет в data.js без правок кода.
- ID берутся из xlsx как есть (уважаются). Пустой id → автоген (prefix + номер) + варн.
  Дубликаты id внутри листа → варн.
- Связь между сущностями — по тексту слова (`de`).
- Тип вопроса build трактуется как алиас tiles (движок тренажёра знает только tiles).
- Колонка studied игнорируется (пережиток, удаляется из xlsx).
"""
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("Нужен openpyxl. pip install openpyxl")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).parent
XLSX_PATH = SCRIPT_DIR / "database.xlsx"
OLD_DATA_JS = SCRIPT_DIR / "data.js"
OUT_DATA_JS = SCRIPT_DIR / "data.js"

# Колонки, которые сознательно игнорируются (пережитки)
IGNORED_COLS = {"studied"}

# ═══════════════════════════════════════════════════════════════
# АВТОГЕНЕРАТОР Präsens (regelmäßig)
# ═══════════════════════════════════════════════════════════════
def conj_regelmaessig(infinitive, pronoun):
    is_eln = infinitive.endswith("eln")
    is_ern = infinitive.endswith("ern")
    is_eln_or_rn = is_eln or is_ern

    if is_eln_or_rn:
        stem = infinitive[:-1]
    elif infinitive.endswith("en"):
        stem = infinitive[:-2]
    elif infinitive.endswith("n"):
        stem = infinitive[:-1]
    else:
        stem = infinitive

    ends_in_td = bool(re.search(r"[td]$", stem))
    needs_e_cluster = bool(re.search(r"(chn|ffn|tm|dm|gn)$", stem))
    needs_e = ends_in_td or needs_e_cluster
    sibilant_end = bool(re.search(r"[sßzx]$", stem))

    if pronoun == "ich":
        if is_eln: return stem[:-2] + "le"
        return stem + "e"
    if pronoun == "du":
        if sibilant_end: return stem + "t"
        if needs_e: return stem + "est"
        return stem + "st"
    if pronoun in ("er", "sie", "es"):
        if needs_e: return stem + "et"
        return stem + "t"
    if pronoun in ("wir", "Sie", "sie_pl"):
        if is_eln_or_rn: return stem + "n"
        return stem + "en"
    if pronoun == "ihr":
        if needs_e: return stem + "et"
        return stem + "t"
    return stem + "en"

UNREGELMAESSIG_VERBS = {
    "sein", "haben", "werden", "wissen", "mögen", "müssen", "können",
    "dürfen", "sollen", "wollen", "lassen", "fahren", "laufen", "schlafen",
    "lesen", "sehen", "geben", "nehmen", "kommen", "gehen", "stehen",
    "tragen", "waschen", "fallen", "halten", "schreiben",
}

def conj_all_forms(infinitive):
    """Возвращает 6 форм Präsens: [ich, du, er, wir, ihr, sie]"""
    if infinitive in UNREGELMAESSIG_VERBS:
        raise ValueError(
            f"conj_all_forms: «{infinitive}» — нерегулярный глагол, "
            f"автоген запрещён. Укажи тип unregel и формы вручную в xlsx."
        )
    return [conj_regelmaessig(infinitive, p)
            for p in ("ich", "du", "er", "wir", "ihr", "sie_pl")]

# ═══════════════════════════════════════════════════════════════
# АВТОГЕНЕРАТОР Partizip II (regelmäßig)
# ═══════════════════════════════════════════════════════════════
INSEPARABLE_PREFIXES = ("be", "ent", "er", "ge", "ver", "zer", "miss", "emp")
SEPARABLE_PREFIXES = (
    "ab", "an", "auf", "aus", "bei", "ein", "fest", "fort", "her", "hin",
    "los", "mit", "nach", "vor", "weg", "weiter", "zu", "zurück", "zusammen"
)

def partizip2_regelmaessig(infinitive):
    if infinitive.endswith("ieren"):
        return infinitive[:-2] + "t"

    if infinitive.endswith("en"):
        stem = infinitive[:-2]
    elif infinitive.endswith("n"):
        stem = infinitive[:-1]
    else:
        stem = infinitive

    needs_et = bool(re.search(r"[td]$", stem)) or \
               bool(re.search(r"(chn|ffn|tm|dm|gn)$", stem))
    ending = "et" if needs_et else "t"

    for prefix in INSEPARABLE_PREFIXES:
        if infinitive.startswith(prefix) and len(infinitive) > len(prefix) + 2:
            return stem + ending

    for prefix in SEPARABLE_PREFIXES:
        if infinitive.startswith(prefix) and len(infinitive) > len(prefix) + 2:
            sub_inf = infinitive[len(prefix):]
            sub_p2 = partizip2_regelmaessig(sub_inf)
            return prefix + sub_p2

    return "ge" + stem + ending

# ═══════════════════════════════════════════════════════════════
# СЕРИАЛИЗАЦИЯ В JS
# ═══════════════════════════════════════════════════════════════
def js_str(s):
    if s is None: return "null"
    s = (str(s).replace("\\", "\\\\").replace('"', '\\"')
         .replace("\n", "\\n").replace("\r", "\\r").replace("\t", "\\t"))
    return f'"{s}"'

def js_value(v):
    if v is None: return "null"
    if isinstance(v, bool): return "true" if v else "false"
    if isinstance(v, (int, float)): return str(v)
    if isinstance(v, list):
        return "[" + ", ".join(js_value(x) for x in v) + "]"
    if isinstance(v, dict):
        # Всегда квотим ключи — они могут содержать `:` и др. спецсимволы.
        return "{ " + ", ".join(f"{js_str(str(k))}: {js_value(val)}" for k, val in v.items()) + " }"
    return js_str(v)

def item_to_js(item, key_order=None):
    if key_order:
        keys = [k for k in key_order if k in item]
        # любые ключи не из key_order — в конец (чтобы новые поля не терялись)
        keys += [k for k in item if k not in key_order]
    else:
        keys = list(item.keys())
    parts = [f"{k}: {js_value(item[k])}" for k in keys]
    return "  { " + ", ".join(parts) + " },"

# ═══════════════════════════════════════════════════════════════
# Утилиты для чтения xlsx
# ═══════════════════════════════════════════════════════════════
def clean(v):
    """None / пустую строку → None, иначе str.strip()."""
    if v is None: return None
    s = str(v).strip()
    return s if s else None

def is_true(v):
    """TRUE / true / 1 / x / yes / да → True. Пусто → False."""
    if v is None: return False
    return str(v).strip().lower() in ("true", "1", "x", "yes", "да")

def is_false(v):
    """Явное FALSE / 0 / no / нет → True (для bool с дефолтом True)."""
    if v is None: return False
    return str(v).strip().lower() in ("false", "0", "no", "нет")

def to_int(v):
    if v is None or str(v).strip() == "": return None
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return None

def read_sheet(wb, name):
    """Читает лист в список dict-ов. Пропускает строку легенды (опц./обяз.)
    и полностью пустые строки. Ключи — заголовки строки 1."""
    if name not in wb.sheetnames:
        print(f"  ⚠ Лист '{name}' не найден")
        return []
    ws = wb[name]
    row1_vals = [c.value for c in ws[1]]
    non_empty_r1 = [v for v in row1_vals if v is not None]
    # Если строка 1 — легенда (обяз./опц.), заголовки в строке 2
    if non_empty_r1 and all(str(v).strip() in ("опц.", "обяз.") for v in non_empty_r1):
        headers = [c.value for c in ws[2]]
        data_start = 3
    else:
        headers = row1_vals
        data_start = 2
    rows = []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        non_empty = [v for v in row if v is not None]
        if non_empty and all(str(v).strip() in ("опц.", "обяз.") for v in non_empty):
            continue
        if not any(v is not None and str(v).strip() for v in row):
            continue
        rows.append(dict(zip(headers, row)))
    return rows

# ═══════════════════════════════════════════════════════════════
# ДЕКЛАРАТИВНАЯ СХЕМА КОЛОНОК
# Формат: (xlsx_col, out_key, kind)
#   kind: "str"  — строка (None если пусто)
#         "bool" — булев флаг (TRUE/1/x/да)
#         "boolT"— булев с дефолтом True (в JS пишем только явный false)
#         "int"  — целое
#         "list" — pipe-separated → массив строк
# Колонки id / domen / group / new обрабатываются отдельно (общая логика).
# Сюда добавляешь новую колонку — она автоматически поедет в data.js.
# ═══════════════════════════════════════════════════════════════
COMMON_META = [
    ("level", "level", "str"),
    ("priority", "priority", "int"),
    ("source", "source", "str"),
    ("example_de", "exampleDe", "str"),
    ("example_ru", "exampleRu", "str"),
    ("quiz_use", "quizUse", "boolT"),   # годится ли ПРИМЕР для теста-сборки
    ("strict_order", "strictOrder", "boolT"),
    ("warning", "warning", "str"),
    ("label", "label", "str"),
    ("note", "note", "str"),
]

SCHEMA = {
    "nouns": [
        ("de", "de", "str"), ("alt-de", "altDe", "str"),
        ("gender", "gender", "str"), ("ru", "ru", "str"), ("alt-ru", "altRu", "str"),
        ("plural", "plural", "str"), ("alt-plural", "altPlural", "str"),
        ("antonym", "antonym", "str"),
    ] + COMMON_META,
    "adjectives": [
        ("de", "de", "str"), ("ru", "ru", "str"), ("alt-ru", "altRu", "str"),
        ("comparative", "comparative", "str"), ("superlative", "superlative", "str"),
        ("antonym", "antonym", "str"), ("derived_from", "derivedFrom", "str"),
    ] + COMMON_META,
    "adverbs": [
        ("de", "de", "str"), ("ru", "ru", "str"), ("alt-ru", "altRu", "str"),
        ("antonym", "antonym", "str"),
    ] + COMMON_META,
    "phrases": [
        ("de", "de", "str"), ("ru", "ru", "str"), ("context", "context", "str"),
    ] + COMMON_META,
    "pronouns": [
        ("de", "de", "str"), ("ru", "ru", "str"), ("kind", "kind", "str"),
        ("case", "case", "str"), ("gender", "gender", "str"),
        ("level", "level", "str"), ("priority", "priority", "int"),
        ("source", "source", "str"), ("note", "note", "str"),
    ],
    "numbers": [
        ("digit", "digit", "str"), ("de", "de", "str"), ("ru", "ru", "str"),
        ("kind", "kind", "str"), ("transcription", "transcription", "str"),
        ("level", "level", "str"), ("priority", "priority", "int"),
        ("source", "source", "str"), ("note", "note", "str"),
    ],
    "terms": [
        ("de", "term", "str"),          # слово термина теперь в колонке de
        ("plural", "plural", "str"), ("ru", "ru", "str"),
        ("level", "level", "str"), ("priority", "priority", "int"),
        ("source", "source", "str"), ("note", "note", "str"),
    ],
    "sounds": [
        ("combo", "combo", "str"), ("pronunciation", "pronunciation", "str"),
        ("example", "example", "str"), ("translation", "translation", "str"),
        ("note", "note", "str"),
    ],
    "rules": [
        ("title", "title", "str"), ("level", "level", "str"),
        ("content_md", "content_md", "str"), ("examples", "examples", "str"),
        ("note", "note", "str"),
    ],
}

def apply_schema(r, sheet):
    """Применяет SCHEMA[sheet] к строке r, возвращает dict (только непустые поля)."""
    out = {}
    for col, key, kind in SCHEMA[sheet]:
        v = r.get(col)
        if kind == "bool":
            if is_true(v): out[key] = True
        elif kind == "boolT":
            # дефолт True: в data.js пишем поле только если ЯВНО false
            if is_false(v): out[key] = False
        elif kind == "int":
            iv = to_int(v)
            if iv is not None: out[key] = iv
        elif kind == "list":
            cv = clean(v)
            if cv: out[key] = [s.strip() for s in cv.split("|") if s.strip()]
        else:  # str
            cv = clean(v)
            if cv is not None: out[key] = cv
    return out

# ═══════════════════════════════════════════════════════════════
# НЕМЕЦКИЕ ПОДПИСИ ДЛЯ ФРОНТА (BLOCKS / TOPIC_TITLES / TAB_TITLES)
# ─────────────────────────────────────────────────────────────────
# Маппинг ключей domen:group → (emoji, deutsch). Хардкод-фоллбэк.
# Если в Settings-XLSX появятся колонки label_de/icon — они
# автоматически переопределят значения отсюда (см. label_for_pair).
# Дописывай сюда новые пары при расширении таксономии.
# ═══════════════════════════════════════════════════════════════
GERMAN_LABELS = {
    # ─── nouns ─────────────────────────────────────────────────
    "communication:docs":       ("📄", "Dokumente"),
    "communication:general":    ("💬", "Kommunikation"),
    "communication:greetings":  ("👋", "Begrüßung"),
    "communication:media":      ("📺", "Medien"),
    "home:cleaning":            ("🧹", "Putzen"),
    "home:furniture":           ("🛋", "Möbel"),
    "leisure:activities":       ("🎯", "Aktivitäten"),
    "leisure:hobby":            ("🎨", "Hobby"),
    "leisure:places":           ("🎪", "Freizeitorte"),
    "movement:traffic":         ("🚦", "Verkehr"),
    "movement:transport":       ("🚗", "Transport"),
    "nature:animals":           ("🐾", "Tiere"),
    "nature:landscape":         ("🏞", "Landschaft"),
    "nature:weather":           ("☀", "Wetter"),
    "people:family":            ("👨‍👩‍👧", "Familie"),
    "people:health":            ("🏥", "Gesundheit"),
    "people:identity":          ("🪪", "Identität"),
    "people:professions":       ("💼", "Berufe"),
    "people:relations":         ("💑", "Beziehungen"),
    "people:roles":             ("🎭", "Rollen"),
    "people:traits":            ("🧠", "Eigenschaften"),
    "place:buildings":          ("🏛", "Gebäude"),
    "place:city":               ("🏙", "Stadt"),
    "place:housing":            ("🏠", "Wohnung"),
    "place:rooms":              ("🚪", "Zimmer"),
    "products:clothing":        ("👕", "Kleidung"),
    "products:drinks":          ("🥤", "Getränke"),
    "products:food":            ("🍎", "Lebensmittel"),
    "products:kitchenware":     ("🍳", "Geschirr"),
    "products:stationery":      ("✏", "Schreibwaren"),
    "products:tech":            ("💻", "Technik"),
    "shopping:general":         ("🛒", "Einkaufen"),
    "state:emotions":           ("😊", "Gefühle"),
    "state:safety":             ("🚨", "Sicherheit"),
    "study:places":             ("🏫", "Lernorte"),
    "study:writing":            ("📝", "Schreiben"),
    "time:clock":               ("🕐", "Uhrzeit"),
    "time:dates":               ("📅", "Datum"),
    "time:dayparts":            ("🌅", "Tageszeiten"),
    "time:periods":             ("⏳", "Zeitabschnitte"),
    "work:general":             ("💼", "Arbeit"),
    "work:office":              ("🏢", "Büro"),
    "work:places":              ("🏗", "Arbeitsorte"),
    # ─── verbs (доп.) ──────────────────────────────────────────
    "grammar:modal":            ("🎯", "Modalverben"),
    "home:chores":              ("🧹", "Hausarbeit"),
    "movement:general":         ("🚶", "Bewegung"),
    "products:cooking":         ("🍳", "Kochen"),
    "study:general":            ("📚", "Lernen"),
    "time:duration":            ("⏱", "Dauer"),
    # ─── adjectives ────────────────────────────────────────────
    "communication:style":      ("💬", "Stil"),
    "movement:speed":           ("⚡", "Tempo"),
    "people:marital":           ("💍", "Familienstand"),
    "people:physical":          ("💪", "Aussehen"),
    "place:size":               ("📏", "Größe"),
    "products:materials":       ("🪵", "Materialien"),
    "shopping:price":           ("💰", "Preis"),
    "state:colors":             ("🎨", "Farben"),
    "state:qualities":          ("⭐", "Qualität"),
    "time:chronology":          ("📅", "Zeitfolge"),
    # ─── adverbs ───────────────────────────────────────────────
    "communication:agreement":  ("🤝", "Zustimmung"),
    "movement:direction":       ("➡", "Richtung"),
    "place:location":           ("📍", "Ort"),
    "state:certainty":          ("🎯", "Sicherheit"),
    "state:feelings":           ("👍", "Befinden"),
    "time:frequency":           ("🔁", "Häufigkeit"),
    # ─── phrases (доп.) ────────────────────────────────────────
    "communication:questions":  ("❓", "Fragen"),
    "communication:wellbeing":  ("💆", "Befinden"),
    "study:classroom":          ("🎓", "Im Unterricht"),
    # ─── rules (домен grammar) ─────────────────────────────────
    "grammar:pronouns":         ("👤", "Pronomen"),
    "grammar:wordorder":        ("📐", "Satzbau"),
    "grammar:verbs":            ("⚡", "Verben"),
    "grammar:cases":            ("🎯", "Kasus"),
    "grammar:articles":         ("📰", "Artikel"),
    "grammar:wfragen":          ("❓", "W-Fragen"),
    "grammar:adjectives":       ("🎨", "Adjektive"),
    # ─── pronouns (домен grammar / специальные) ────────────────
    "grammar:personal":         ("👤", "Personalpronomen"),
    "grammar:possessive":       ("🔑", "Possessivpronomen"),
    "grammar:indefinite":       ("🌫", "Indefinitpronomen"),
    "grammar:negation":         ("🚫", "Negation"),
    "grammar:particles":        ("🎯", "Partikeln"),
    "grammar:prepositions":     ("🔗", "Präpositionen"),
    # ─── numbers (доп.) ────────────────────────────────────────
    "quantity:ordinal":         ("🥇", "Ordinalzahlen"),
    # ─── phonetics (звуки) ─────────────────────────────────────
    "phonetics:rules":          ("📜", "Regeln"),
    "phonetics:vowels":         ("🅰", "Vokale"),
    "phonetics:consonants":     ("🅱", "Konsonanten"),
    # ─── специальные (не из Settings) ──────────────────────────
    "quantity:cardinal":        ("🔢", "Kardinalzahlen"),
    "pron:personal":            ("👤", "Personalpronomen"),
    "pron:negation":            ("🚫", "Negation"),
    "new:new-nouns":            ("📦", "Neue Substantive"),
    "new:new-verbs":             ("⚡", "Neue Verben"),
    "new:new-adj":               ("🎨", "Neue Adjektive"),
    "new:new-adv":               ("⏱", "Neue Adverbien"),
    "new:new-phrases":           ("💬", "Neue Phrasen"),
    "new:new-pron":              ("👤", "Neue Pronomen"),
    "new:new-nums":              ("🔢", "Neue Zahlen"),
    "new:new-terms":             ("📖", "Neue Begriffe"),
    "new:new-sounds":            ("🔤", "Neue Laute"),
    "new:new-rules":             ("📐", "Neue Regeln"),
    "perfekt:regel":            ("📘", "Regelmäßig"),
    "perfekt:unregel":          ("🔥", "Unregelmäßig"),
}

# РУССКИЕ ПЕРЕВОДЫ (для тултипов на фронте)
# ─────────────────────────────────────────────────────────────────
RUSSIAN_LABELS = {
    # ─── существительные ────────────────────────────────────────
    "communication:docs":       "Документы",
    "communication:general":    "Общение",
    "communication:greetings":  "Приветствия",
    "communication:media":      "СМИ / медиа",
    "home:cleaning":            "Уборка",
    "home:furniture":           "Мебель",
    "leisure:activities":       "Занятия",
    "leisure:hobby":            "Хобби",
    "leisure:places":           "Места отдыха",
    "movement:traffic":         "Дорожное движение",
    "movement:transport":       "Транспорт",
    "nature:animals":           "Животные",
    "nature:landscape":         "Природа / пейзаж",
    "nature:weather":           "Погода",
    "people:family":            "Семья",
    "people:health":            "Здоровье",
    "people:identity":          "Личность / удостоверение",
    "people:professions":       "Профессии",
    "people:relations":         "Отношения",
    "people:roles":             "Роли",
    "people:traits":            "Черты характера",
    "place:buildings":          "Здания",
    "place:city":               "Город",
    "place:housing":            "Жильё",
    "place:rooms":              "Комнаты",
    "products:clothing":        "Одежда",
    "products:drinks":          "Напитки",
    "products:food":            "Еда / продукты",
    "products:kitchenware":     "Посуда",
    "products:stationery":      "Канцелярия",
    "products:tech":            "Техника",
    "shopping:general":         "Покупки",
    "state:emotions":           "Чувства / эмоции",
    "state:safety":             "Безопасность",
    "study:places":             "Места учёбы",
    "study:writing":            "Письмо / текст",
    "time:clock":               "Время (часы)",
    "time:dates":               "Даты / дни недели",
    "time:dayparts":            "Части дня",
    "time:periods":             "Периоды времени",
    "work:general":             "Работа",
    "work:office":              "Офис",
    "work:places":              "Рабочие места",
    # ─── глаголы (доп.) ─────────────────────────────────────────
    "grammar:modal":            "Модальные глаголы",
    "home:chores":              "Домашние дела",
    "movement:general":         "Движение",
    "products:cooking":         "Готовка",
    "study:general":            "Учёба",
    "time:duration":            "Продолжительность",
    "communication:style":      "Стиль общения",
    # ─── прилагательные ─────────────────────────────────────────
    "movement:speed":           "Скорость",
    "people:marital":           "Семейное положение",
    "people:physical":          "Внешность",
    "place:size":               "Размер",
    "products:materials":       "Материалы",
    "shopping:price":           "Цена",
    "state:colors":             "Цвета",
    "state:qualities":          "Качества",
    "time:chronology":          "Хронология",
    # ─── наречия ────────────────────────────────────────────────
    "communication:agreement":  "Согласие / несогласие",
    "movement:direction":       "Направление",
    "place:location":           "Местоположение",
    "state:certainty":          "Уверенность",
    "state:feelings":           "Самочувствие",
    "time:frequency":           "Частота",
    # ─── фразы ──────────────────────────────────────────────────
    "communication:questions":  "Вопросы",
    "communication:wellbeing":  "Самочувствие",
    "study:classroom":          "В классе",
    # ─── грамматика / правила ────────────────────────────────────
    "grammar:pronouns":         "Местоимения",
    "grammar:wordorder":        "Порядок слов",
    "grammar:verbs":            "Глаголы",
    "grammar:cases":            "Падежи",
    "grammar:articles":         "Артикли",
    "grammar:wfragen":          "W-вопросы",
    "grammar:adjectives":       "Прилагательные",
    "grammar:personal":         "Личные местоимения",
    "grammar:possessive":       "Притяжательные местоимения",
    "grammar:indefinite":       "Неопределённые местоимения",
    "grammar:negation":         "Отрицание",
    "grammar:particles":        "Частицы",
    "grammar:prepositions":     "Предлоги",
    # ─── числа / фонетика ───────────────────────────────────────
    "quantity:cardinal":        "Количественные числительные",
    "quantity:ordinal":         "Порядковые числительные",
    "phonetics:rules":          "Правила произношения",
    "phonetics:vowels":         "Гласные",
    "phonetics:consonants":     "Согласные",
    # ─── специальные ────────────────────────────────────────────
    "pron:personal":            "Личные местоимения",
    "pron:negation":            "Отрицание",
    "new:new-nouns":            "Новые существительные",
    "new:new-verbs":            "Новые глаголы",
    "new:new-adj":              "Новые прилагательные",
    "new:new-adv":              "Новые наречия",
    "new:new-phrases":          "Новые фразы",
    "new:new-pron":             "Новые местоимения",
    "new:new-nums":             "Новые числа",
    "new:new-terms":            "Новые термины",
    "new:new-sounds":           "Новые звуки",
    "new:new-rules":            "Новые правила",
    "perfekt:regel":            "Регулярные глаголы",
    "perfekt:unregel":          "Нерегулярные глаголы",
}

BLOCK_META_RU = {
    "neu":     "Новые слова (буфер)",
    "nums":    "Числа",
    "sounds":  "Произношение",
    "verbs":   "Глаголы",
    "unregel": "Нерегулярные глаголы в Präsens",
    "perfekt": "Прошедшее время (Perfekt)",
    "nouns":   "Существительные",
    "adj":     "Прилагательные",
    "adv":     "Наречия",
    "mestoim": "Местоимения",
    "termin":  "Грамматические термины",
    "rules":   "Правила грамматики",
    "phrases": "Разговорные фразы",
}

# Подписи и цвета верхнеуровневых блоков (POS / спец.).
BLOCK_META = {
    "neu":     {"label": "🆕 Neu",            "color": "var(--c-block-neu)", "desc": "Frische Wörter (Puffer)"},
    "verbs":   {"label": "⚡ Verben",         "color": "var(--c-block-verben)"},
    "nouns":   {"label": "📦 Substantive",    "color": "var(--c-block-substantive)"},
    "adj":     {"label": "🎨 Adjektive",      "color": "var(--c-block-adj)"},
    "adv":     {"label": "🔄 Adverbien",      "color": "var(--c-block-adv)"},
    "mestoim": {"label": "👤 Pronomen",       "color": "var(--c-block-pronomen)"},
    "phrases": {"label": "💬 Redewendungen",  "color": "var(--c-block-redewendungen)"},
    "unregel": {"label": "🔥 Unregelmäßig", "color": "var(--c-block-unregel)", "desc": "Unregelmäßige Verben in Präsens", "kind": "conjugations"},
    "nums":    {"label": "🔢 Zahlen",         "color": "#16a085"},
    "sounds":  {"label": "🔤 Aussprache",     "color": "#9b59b6", "desc": "Ausspracheregeln", "kind": "sounds"},
    
}

def label_for_pair(domen, group, taxonomy_entry=None):
    """Возвращает 'эмодзи Подпись' для domen:group.
    Override через колонки label_de/label/icon в Settings-XLSX."""
    if taxonomy_entry:
        de_label = taxonomy_entry.get("label_de") or taxonomy_entry.get("label")
        icon = taxonomy_entry.get("icon")
        if de_label:
            return f"{icon} {de_label}".strip() if icon else de_label
    key = f"{domen}:{group}"
    emoji, name = GERMAN_LABELS.get(key, ("", key))
    return f"{emoji} {name}".strip() if emoji else key

def build_block_from_pairs(bid, taxonomy_pairs, used_topics=None):
    """Блок POS: каждая пара domen:group → подблок.
    Если передан used_topics — отфильтровываем подблоки без слов."""
    meta = BLOCK_META[bid]
    block = {"id": bid, "label": meta["label"], "color": meta["color"]}
    if "desc" in meta: block["desc"] = meta["desc"]
    seen = set()
    subs = []
    for p in (taxonomy_pairs or []):
        d, g = p.get("domen"), p.get("group")
        if not (d and g) or (d, g) in seen: continue
        seen.add((d, g))
        topic = f"{d}:{g}"
        if used_topics is not None and topic not in used_topics:
            continue  # пустой подблок — не показываем
        subs.append({
            "id": f"{d}-{g}".replace("/", "-"),
            "label": label_for_pair(d, g, p),
            "topics": [topic],
        })
    if subs: block["subblocks"] = subs
    return block

def build_special_block(bid, subblocks=None):
    """Спец. блок (kind=sounds/terms/rules/conjugations/perfekt или с фикс. подблоками)."""
    meta = BLOCK_META[bid]
    block = {"id": bid, "label": meta["label"], "color": meta["color"]}
    if "desc" in meta: block["desc"] = meta["desc"]
    if "kind" in meta: block["kind"] = meta["kind"]
    if subblocks: block["subblocks"] = subblocks
    return block

def collect_used_topics_by_pos(vocab):
    """Для каждого pos — множество topics, реально встречающихся у слов."""
    out = {}
    for v in vocab:
        pos = v.get("pos")
        for t in v.get("topics", []) or []:
            out.setdefault(pos, set()).add(t)
    return out

def build_blocks(taxonomy, vocab=None):
    """BLOCKS строится из Settings-XLSX + спец. блоков. Порядок важен (UI).
    Если передан vocab — отфильтровываем пустые подблоки.

    Архитектура:
    - В тренажёре 7 блоков: Neu, Verben (+Unregel внутри), Substantive,
      Adjektive, Adverbien, Pronomen, Redewendungen.
    - Sounds остаётся в массиве (kind=sounds), но тренажёр его фильтрует —
      используется только в шпоре.
    - Unregel — сабблок Verben (kind=conjugations).
      trainer.html распознаёт sub.kind и роутит его в свой поток.
    - Блок Perfekt удалён: его функционал (тренировка Partizip II) полностью
      дублируется комбо-кнопкой «Partizip II» на экране select-test-type.
      Контекст задаёт фильтр автоматически: в Verben — все глаголы (regel+unregel),
      в Unregel-сабблоке — только unregel.
    """
    used = collect_used_topics_by_pos(vocab or [])

    # Verben + спец. сабблок Unregel в начале
    verbs_block = build_block_from_pairs("verbs", taxonomy.get("verbs", []), used.get("verb"))
    unregel_sub = {
        "id": "unregel",
        "label": BLOCK_META["unregel"]["label"],
        "color": BLOCK_META["unregel"]["color"],
        "desc": BLOCK_META["unregel"].get("desc", ""),
        "kind": "conjugations",
    }
    verbs_block["subblocks"] = [unregel_sub] + verbs_block.get("subblocks", [])

    return [
        # 1. Neu (буфер новых)
        build_special_block("neu", subblocks=[
            {"id": "nouns", "label": "📦 Substantive",  "topics": ["new:new-nouns"]},
            {"id": "verbs", "label": "⚡ Verben",       "topics": ["new:new-verbs"]},
            {"id": "adj",   "label": "🎨 Adjektive",   "topics": ["new:new-adj"]},
            {"id": "adv",   "label": "⏱ Adverbien",   "topics": ["new:new-adv"]},
        ]),
        # 2. Aussprache (kind=sounds, только для шпоры; trainer фильтрует)
        build_special_block("sounds"),
        # 3. Verben (+ Unregel-сабблок в начале)
        verbs_block,
        # 4. Substantive
        build_block_from_pairs("nouns", taxonomy.get("nouns", []), used.get("noun")),
        # 5. Adjektive
        build_block_from_pairs("adj", taxonomy.get("adjectives", []), used.get("adj")),
        # 6. Adverbien
        build_block_from_pairs("adv", taxonomy.get("adverbs", []), used.get("adv")),
        # 7. Pronomen
        build_block_from_pairs("mestoim", taxonomy.get("pronouns", []), used.get("pron")),
        # 8. Redewendungen
        build_block_from_pairs("phrases", taxonomy.get("phrases", []), used.get("phrase")),
    ]

def build_topic_titles(taxonomy):
    """TOPIC_TITLES — заголовки секций в шпоре."""
    out = {}
    for page, pairs in taxonomy.items():
        for p in pairs:
            d, g = p.get("domen"), p.get("group")
            if not (d and g): continue
            out[f"{d}:{g}"] = label_for_pair(d, g, p)
    for k in ("quantity:cardinal", "pron:personal", "pron:negation",
              "new:new-nouns", "new:new-verbs", "new:new-adj", "new:new-adv",
              "new:new-phrases", "new:new-pron", "new:new-nums",
              "new:new-terms", "new:new-sounds", "new:new-rules",
              "perfekt:regel", "perfekt:unregel"):
        emoji, name = GERMAN_LABELS.get(k, ("", k))
        out[k] = f"{emoji} {name}".strip() if emoji else k
    return out

def build_tab_titles():
    return {bid: meta["label"] for bid, meta in BLOCK_META.items()}

def build_ru_titles(taxonomy):
    """TOPIC_TITLES_RU и TAB_TITLES_RU — русские переводы для тултипов."""
    topic_ru = {}
    for pairs in taxonomy.values():
        for p in pairs:
            d, g = p.get("domen"), p.get("group")
            if not (d and g): continue
            key = f"{d}:{g}"
            if key in RUSSIAN_LABELS:
                topic_ru[key] = RUSSIAN_LABELS[key]
    # специальные ключи
    for k in ("quantity:cardinal", "pron:personal", "pron:negation",
              "new:new-nouns", "new:new-verbs", "new:new-adj", "new:new-adv",
              "new:new-phrases", "new:new-pron", "new:new-nums",
              "new:new-terms", "new:new-sounds", "new:new-rules",
              "perfekt:regel", "perfekt:unregel"):
        if k in RUSSIAN_LABELS:
            topic_ru[k] = RUSSIAN_LABELS[k]
    return topic_ru, dict(BLOCK_META_RU)

# ───────────────────────────────────────────────────────────────
# Хелперы domen/group → topics
# ───────────────────────────────────────────────────────────────
def dg(r):
    """Возвращает (domen, group) очищенные или (None, None)."""
    return clean(r.get("domen")), clean(r.get("group"))

def topic_key(domen, group):
    """Канонический ключ темы. domen:group, либо domen, либо None."""
    if domen and group: return f"{domen}:{group}"
    if domen: return domen
    return None

# ═══════════════════════════════════════════════════════════════
# ОБРАБОТКА ЛИСТОВ
# ═══════════════════════════════════════════════════════════════
NEW_TOPIC = {  # буфер «Новые» по pos
    "noun":   "new:new-nouns",
    "verb":   "new:new-verbs",
    "adj":    "new:new-adj",
    "adv":    "new:new-adv",
    "phrase": "new:new-phrases",
    "pron":   "new:new-pron",
    "num":    "new:new-nums",
    "term":   "new:new-terms",
    "sound":  "new:new-sounds",
    "rule":   "new:new-rules",
}

def attach_common(item, r, pos):
    """Добавляет id/domen/group/topics/new к vocab-элементу."""
    rid = clean(r.get("id"))
    if rid: item["id"] = rid
    domen, group = dg(r)
    if domen: item["domen"] = domen
    if group: item["group"] = group
    topics = []
    tk = topic_key(domen, group)
    if tk: topics.append(tk)
    if is_true(r.get("new")):
        item["new"] = True
        nt = NEW_TOPIC.get(pos)
        if nt and nt not in topics:
            topics.append(nt)
    item["topics"] = topics

def process_simple_vocab(rows, sheet, pos, warn):
    """nouns / adjectives / adverbs / phrases — общая логика."""
    items = []
    seen = set()
    for r in rows:
        de = clean(r.get("de"))
        if not de: continue
        # legacy: инлайн "X / Y" в de → de + altDe (на случай старых записей)
        alt_inline = None
        if " / " in de and not clean(r.get("alt-de")):
            parts = [p.strip() for p in de.split(" / ")]
            de = parts[0]
            alt_inline = " / ".join(parts[1:]) if len(parts) > 1 else None
        domen, group = dg(r)
        key = (de, domen or "", group or "")
        if key in seen:
            warn.append(f"{sheet}: дубль «{de}» ({domen}/{group})")
            continue
        seen.add(key)
        item = {"de": de}
        item.update(apply_schema(r, sheet))
        if alt_inline and "altDe" not in item:
            item["altDe"] = alt_inline
        item["pos"] = pos
        attach_common(item, r, pos)
        items.append(item)
    return items

def process_verbs(rows, warn):
    """→ (vocab_items, regel_verbs, conjugations)."""
    vocab_items, regel_verbs, conjugations = [], [], []
    seen = set()
    missing_p2, missing_aux = [], []

    for r in rows:
        verb = clean(r.get("de"))
        if not verb: continue
        if verb in seen:
            warn.append(f"verbs: дубль «{verb}»")
            continue
        seen.add(verb)

        vtype = (clean(r.get("type")) or "regel").lower()
        if vtype not in ("regel", "unregel", "composite"):
            vtype = "regel"

        # ── VOCAB-запись глагола (все поля схемы verbs) ──
        vi = {"de": verb}
        vi.update(apply_schema_verb(r))
        vi["pos"] = "verb"
        attach_common(vi, r, "verb")
        vocab_items.append(vi)

        if vtype == "composite":
            continue  # только в VOCAB

        # формы Präsens
        forms = [clean(r.get(c)) for c in ("ich", "du", "er_sie_es", "wir", "ihr", "sie_Sie")]
        forms = [f or "" for f in forms]
        has_all_forms = all(forms)

        aux = (clean(r.get("aux")) or "").lower()
        if aux not in ("haben", "sein", ""):
            aux = ""
        p2 = clean(r.get("partizip2")) or ""
        praet = clean(r.get("praeteritum"))
        modal = is_true(r.get("modal"))

        if vtype == "unregel":
            if not has_all_forms:
                warn.append(f"verbs: unregel без полных форм Präsens — «{verb}» (автоген)")
                forms = conj_all_forms(verb)
            if not p2: missing_p2.append(verb)
            if not aux: missing_aux.append(verb)
            conj = {
                "verb": verb,
                "ru": clean(r.get("ru")) or "",
                "tense": "Präsens",
                "pronouns": ["ich", "du", "er/sie/es", "wir", "ihr", "sie/Sie"],
                "forms": forms,
            }
            if modal: conj["modal"] = True
            if p2: conj["partizip2"] = p2
            if aux: conj["aux"] = aux
            if praet: conj["praeteritum"] = praet
            for fl in ("separable", "reflexive", "impersonal"):
                if is_true(r.get(fl)): conj[fl] = True
            if clean(r.get("case")): conj["case"] = clean(r.get("case"))
            if clean(r.get("level")): conj["level"] = clean(r.get("level"))
            conjugations.append(conj)
        else:  # regel
            if not has_all_forms:
                forms = conj_all_forms(verb)
            if not p2:
                p2 = partizip2_regelmaessig(verb)
            if not aux:
                aux = "haben"
            reg = {"verb": verb, "ru": clean(r.get("ru")) or "",
                   "forms": forms, "partizip2": p2, "aux": aux}
            if praet: reg["praeteritum"] = praet
            for fl in ("separable", "reflexive", "impersonal"):
                if is_true(r.get(fl)): reg[fl] = True
            if clean(r.get("case")): reg["case"] = clean(r.get("case"))
            regel_verbs.append(reg)

    if missing_p2:
        warn.append(f"verbs: unregel без partizip2 ({len(missing_p2)}): {missing_p2[:8]}")
    if missing_aux:
        warn.append(f"verbs: unregel без aux ({len(missing_aux)}): {missing_aux[:8]}")
    return vocab_items, regel_verbs, conjugations

VERB_VOCAB_SCHEMA = [
    ("ru", "ru", "str"), ("alt-ru", "altRu", "str"), ("type", "type", "str"),
    ("modal", "modal", "bool"), ("separable", "separable", "bool"),
    ("prefix", "prefix", "str"), ("reflexive", "reflexive", "bool"),
    ("impersonal", "impersonal", "bool"), ("case", "case", "str"),
    ("aux", "aux", "str"), ("partizip2", "partizip2", "str"),
    ("praeteritum", "praeteritum", "str"),
    ("level", "level", "str"), ("priority", "priority", "int"),
    ("source", "source", "str"),
    ("example_de", "exampleDe", "str"), ("example_ru", "exampleRu", "str"),
    ("quiz_use", "quizUse", "boolT"), ("strict_order", "strictOrder", "boolT"),
    ("warning", "warning", "str"), ("label", "label", "str"), ("note", "note", "str"),
]
def apply_schema_verb(r):
    out = {}
    for col, key, kind in VERB_VOCAB_SCHEMA:
        v = r.get(col)
        if kind == "bool":
            if is_true(v): out[key] = True
        elif kind == "boolT":
            if is_false(v): out[key] = False
        elif kind == "int":
            iv = to_int(v)
            if iv is not None: out[key] = iv
        else:
            cv = clean(v)
            if cv is not None: out[key] = cv
    return out

def process_pronouns(rows):
    items = []
    for r in rows:
        de = clean(r.get("de"))
        if not de: continue
        item = {"de": de, "pos": "pron"}
        item.update(apply_schema(r, "pronouns"))
        attach_common(item, r, "pron")
        items.append(item)
    return items

def process_numbers(rows):
    items = []
    for r in rows:
        de = clean(r.get("de"))
        if not de: continue
        item = {"de": de, "pos": "num"}
        item.update(apply_schema(r, "numbers"))
        rid = clean(r.get("id"))
        if rid: item["id"] = rid
        kind = item.get("kind")
        domen, group = dg(r)
        topics = []
        tk = topic_key(domen, group)
        if tk:
            topics.append(tk)
        else:
            topics.append("nums:" + kind if kind else "nums:basic")
        # буфер «Новые»
        if is_true(r.get("new")):
            item["new"] = True
            nt = NEW_TOPIC.get("num")
            if nt and nt not in topics:
                topics.append(nt)
        item["topics"] = topics
        if domen: item["domen"] = domen
        if group: item["group"] = group
        items.append(item)
    return items

def process_sounds(rows):
    items = []
    for r in rows:
        combo = clean(r.get("combo"))
        if not combo: continue
        item = apply_schema(r, "sounds")
        item = {"combo": combo, **{k: v for k, v in item.items() if k != "combo"}}
        rid = clean(r.get("id"))
        if rid: item["id"] = rid
        domen, group = dg(r)
        if domen: item["domen"] = domen
        if group: item["group"] = group
        # буфер «Новые»
        topics = []
        tk = topic_key(domen, group)
        if tk: topics.append(tk)
        if is_true(r.get("new")):
            item["new"] = True
            nt = NEW_TOPIC.get("sound")
            if nt and nt not in topics:
                topics.append(nt)
        if topics: item["topics"] = topics
        items.append(item)
    return items

def process_terms(rows):
    items = []
    for r in rows:
        term = clean(r.get("de"))
        if not term: continue
        item = apply_schema(r, "terms")  # de→term внутри схемы
        item = {"term": term, **{k: v for k, v in item.items() if k != "term"}}
        rid = clean(r.get("id"))
        if rid: item["id"] = rid
        domen, group = dg(r)
        tk = topic_key(domen, group)
        if tk: item["topic"] = tk
        if domen: item["domen"] = domen
        if group: item["group"] = group
        # буфер «Новые»
        topics = []
        if tk: topics.append(tk)
        if is_true(r.get("new")):
            item["new"] = True
            nt = NEW_TOPIC.get("term")
            if nt and nt not in topics:
                topics.append(nt)
        if topics: item["topics"] = topics
        items.append(item)
    return items

def process_rules(rows, warn):
    items = []
    for r in rows:
        title = clean(r.get("title"))
        if not title: continue
        item = {"title": title}
        item.update(apply_schema(r, "rules"))
        rid = clean(r.get("id"))
        if rid: item["id"] = rid
        domen, group = dg(r)
        tk = topic_key(domen, group)
        if tk:
            item["topic"] = tk
        else:
            warn.append(f"rules: нет topic (пустые domen/group) — «{title[:40]}»")
        if domen: item["domen"] = domen
        if group: item["group"] = group
        # буфер «Новые» — кладём и флаг, и topic (для рендера на вкладке Neu)
        topics = []
        if tk: topics.append(tk)
        if is_true(r.get("new")):
            item["new"] = True
            nt = NEW_TOPIC.get("rule")
            if nt and nt not in topics:
                topics.append(nt)
        if topics: item["topics"] = topics
        items.append(item)
    return items

DIFFICULTY_DEFAULTS = {"mc": 1, "fill": 4, "tiles": 4, "conj": 3, "open": 5}
TYPE_ALIASES = {"build": "tiles"}  # движок тренажёра знает только tiles

def process_questions(rows, warn):
    items = []
    no_topic = 0
    for r in rows:
        qtype_raw = clean(r.get("type"))
        q = clean(r.get("q"))
        if not (qtype_raw and q): continue
        qtype = TYPE_ALIASES.get(qtype_raw.lower(), qtype_raw)

        domen, group = dg(r)
        tk = topic_key(domen, group)
        item = {"type": qtype, "q": q}
        if tk:
            item = {"topic": tk, **item}
        else:
            no_topic += 1

        level = clean(r.get("level"))
        if level: item["level"] = level
        diff = to_int(r.get("difficulty"))
        item["difficulty"] = diff if diff is not None else DIFFICULTY_DEFAULTS.get(qtype, 1)

        if qtype == "mc":
            opts = clean(r.get("opts"))
            if opts: item["opts"] = [s.strip() for s in opts.split("|")]
            ans = to_int(r.get("ans_mc"))
            if ans is not None: item["ans"] = ans
        elif qtype == "fill":
            ans = clean(r.get("answer"))
            if ans: item["ans"] = ans
            alt = clean(r.get("altAns"))
            if alt: item["altAns"] = [s.strip() for s in alt.split("|")]
        elif qtype == "tiles":
            words = clean(r.get("words"))
            if words: item["words"] = [s.strip() for s in words.split("|")]
        elif qtype == "conj":
            pr = clean(r.get("pronouns"))
            if pr: item["pronouns"] = [s.strip() for s in pr.split("|")]
            fm = clean(r.get("forms"))
            if fm: item["ans"] = [s.strip() for s in fm.split("|")]
        if clean(r.get("hint")): item["hint"] = clean(r.get("hint"))
        if clean(r.get("explain")): item["explain"] = clean(r.get("explain"))
        rid = clean(r.get("id"))
        if rid: item["id"] = rid
        items.append(item)
    if no_topic:
        warn.append(f"questions: без topic (пустые domen/group): {no_topic} из {len(items)}")
    return items

# ═══════════════════════════════════════════════════════════════
# Settings-XLSX → TAXONOMY (генерически, все колонки)
# ═══════════════════════════════════════════════════════════════
def process_settings(wb):
    """Читает Settings-XLSX в структуру { page: [ {domen, group, ...extra} ] }.
    Читает ВСЕ колонки листа — будущие колонки подхватятся автоматически."""
    name = "Settings-XLSX"
    if name not in wb.sheetnames:
        print(f"  ⚠ Лист '{name}' не найден — TAXONOMY будет пустой")
        return {}
    ws = wb[name]
    headers = [clean(c.value) for c in ws[1]]
    tax = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = {h: clean(v) for h, v in zip(headers, row) if h}
        page = rec.get("page")
        if not page: continue
        entry = {k: v for k, v in rec.items() if k != "page" and v is not None}
        tax.setdefault(page, []).append(entry)
    return tax

def validate_against_taxonomy(items, page, tax, warn):
    """Варнит, если (domen, group) элемента нет в Settings для данной page."""
    if page not in tax: return
    valid = {(e.get("domen"), e.get("group")) for e in tax[page]}
    seen_bad = set()
    for it in items:
        d, g = it.get("domen"), it.get("group")
        if d is None and g is None: continue
        if (d, g) not in valid and (d, g) not in seen_bad:
            seen_bad.add((d, g))
            warn.append(f"{page}: пара ({d}/{g}) отсутствует в Settings-XLSX")

# ═══════════════════════════════════════════════════════════════
# ID: уважаем xlsx, автоген для пустых, детект дублей
# ═══════════════════════════════════════════════════════════════
def assign_ids(items, prefix, warn, label, width=4, warn_blanks=True):
    used = set()
    dupes = set()
    for it in items:
        rid = it.get("id")
        rid = str(rid).strip() if rid not in (None, "") else None
        if rid:
            if rid in used: dupes.add(rid)
            used.add(rid)
    # автоген свободных номеров для пустых
    counter = 1
    def next_free():
        nonlocal counter
        while True:
            cand = f"{prefix}{counter:0{width}d}"
            counter += 1
            if cand not in used:
                used.add(cand)
                return cand
    blanks = 0
    for it in items:
        rid = it.get("id")
        rid = str(rid).strip() if rid not in (None, "") else None
        if not rid:
            rid = next_free()
            blanks += 1
        # id первым ключом
        rest = {k: v for k, v in it.items() if k != "id"}
        it.clear()
        it["id"] = rid
        it.update(rest)
    if dupes:
        warn.append(f"{label}: дубликаты id ({len(dupes)}): {sorted(dupes)[:8]}")
    if blanks and warn_blanks:
        warn.append(f"{label}: пустых id заполнено автогеном: {blanks}")

# ═══════════════════════════════════════════════════════════════
# Сохранение констант из старого data.js
# ═══════════════════════════════════════════════════════════════
def extract_block(content, marker):
    start_marker = f"const {marker} = "
    start = content.find(start_marker)
    if start == -1: return None
    start += len(start_marker)
    if content[start] not in "[{": return None
    open_ch = content[start]
    close_ch = "]" if open_ch == "[" else "}"
    depth = 0
    pos = start
    while pos < len(content):
        c = content[pos]
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return content[start:pos+1]
        elif c == '"':
            pos += 1
            while pos < len(content) and content[pos] != '"':
                pos += 2 if content[pos] == '\\' else 1
        elif c == "'":
            pos += 1
            while pos < len(content) and content[pos] != "'":
                pos += 2 if content[pos] == '\\' else 1
        pos += 1
    return None

# ═══════════════════════════════════════════════════════════════
# ОСНОВНОЙ ПРОЦЕСС
# ═══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    WARN = []

    print(f"Читаю {XLSX_PATH}...")
    wb = load_workbook(XLSX_PATH, data_only=True)

    print("\n=== Settings-XLSX → TAXONOMY ===")
    TAXONOMY = process_settings(wb)
    print(f"  ✓ {sum(len(v) for v in TAXONOMY.values())} пар в {len(TAXONOMY)} страницах")

    print("\n=== nouns ===")
    nouns = process_simple_vocab(read_sheet(wb, "nouns"), "nouns", "noun", WARN)
    print(f"  ✓ {len(nouns)} существительных")

    print("\n=== verbs ===")
    verbs_vocab, regel_verbs, conjugations = process_verbs(read_sheet(wb, "verbs"), WARN)
    print(f"  ✓ {len(verbs_vocab)} в VOCAB, {len(regel_verbs)} regel, {len(conjugations)} unregel")

    print("\n=== adjectives ===")
    adjectives = process_simple_vocab(read_sheet(wb, "adjectives"), "adjectives", "adj", WARN)
    print(f"  ✓ {len(adjectives)} прилагательных")

    print("\n=== adverbs ===")
    adverbs = process_simple_vocab(read_sheet(wb, "adverbs"), "adverbs", "adv", WARN)
    print(f"  ✓ {len(adverbs)} наречий")

    print("\n=== phrases ===")
    phrases = process_simple_vocab(read_sheet(wb, "phrases"), "phrases", "phrase", WARN)
    print(f"  ✓ {len(phrases)} фраз")

    print("\n=== pronouns ===")
    pronouns = process_pronouns(read_sheet(wb, "pronouns"))
    print(f"  ✓ {len(pronouns)} местоимений")

    print("\n=== numbers ===")
    numbers = process_numbers(read_sheet(wb, "numbers"))
    print(f"  ✓ {len(numbers)} чисел")

    print("\n=== terms ===")
    terms = process_terms(read_sheet(wb, "terms"))
    print(f"  ✓ {len(terms)} терминов")

    print("\n=== sounds ===")
    sounds = process_sounds(read_sheet(wb, "sounds"))
    print(f"  ✓ {len(sounds)} звуков")

    print("\n=== rules ===")
    rules = process_rules(read_sheet(wb, "rules"), WARN)
    print(f"  ✓ {len(rules)} правил")

    print("\n=== questions ===")
    questions = process_questions(read_sheet(wb, "questions"), WARN)
    print(f"  ✓ {len(questions)} вопросов")

    # Валидация пар против Settings
    for items, page in [(nouns, "nouns"), (verbs_vocab, "verbs"), (adjectives, "adjectives"),
                        (adverbs, "adverbs"), (phrases, "phrases")]:
        validate_against_taxonomy(items, page, TAXONOMY, WARN)

    # ═══════════════════════════════════════════════════════════════
    # Считываем константы фронта из старого data.js
    # ═══════════════════════════════════════════════════════════════
    print("\n=== Читаю старый data.js (PHRASE_UNITS / SENTENCE_TEMPLATES) ===")
    old_content = OLD_DATA_JS.read_text(encoding="utf-8")
    phrase_units_block = extract_block(old_content, "PHRASE_UNITS")
    sentence_templates_block = extract_block(old_content, "SENTENCE_TEMPLATES")
    for nm, b in [("PHRASE_UNITS", phrase_units_block),
                  ("SENTENCE_TEMPLATES", sentence_templates_block)]:
        if b is None:
            WARN.append(f"не нашёл {nm} в старом data.js — заглушка")

    # ═══════════════════════════════════════════════════════════════
    # Назначаем id
    # ═══════════════════════════════════════════════════════════════
    all_vocab = []
    for items, prefix, lbl in [
        (nouns, "n", "nouns"), (verbs_vocab, "v", "verbs"), (adjectives, "a", "adjectives"),
        (adverbs, "d", "adverbs"), (pronouns, "p", "pronouns"),
        (numbers, "u", "numbers"), (phrases, "f", "phrases"),
    ]:
        assign_ids(items, prefix, WARN, lbl)
        all_vocab.extend(items)

    assign_ids(regel_verbs, "r", WARN, "regel_verbs", width=3, warn_blanks=False)
    assign_ids(conjugations, "c", WARN, "conjugations", width=3, warn_blanks=False)
    assign_ids(questions, "q", WARN, "questions")
    assign_ids(rules, "rl", WARN, "rules", width=3)
    assign_ids(terms, "t", WARN, "terms", width=3)
    assign_ids(sounds, "s", WARN, "sounds", width=3)

    print("\n=== Генерирую BLOCKS / TOPIC_TITLES / TAB_TITLES (немецкий) ===")
    BLOCKS_DATA = build_blocks(TAXONOMY, all_vocab)
    TOPIC_TITLES_DATA = build_topic_titles(TAXONOMY)
    TAB_TITLES_DATA = build_tab_titles()
    TOPIC_TITLES_RU_DATA, TAB_TITLES_RU_DATA = build_ru_titles(TAXONOMY)
    print(f"  ✓ BLOCKS: {len(BLOCKS_DATA)} блоков, "
          f"{sum(len(b.get('subblocks',[])) for b in BLOCKS_DATA)} подблоков")
    print(f"  ✓ TOPIC_TITLES: {len(TOPIC_TITLES_DATA)} ключей, RU: {len(TOPIC_TITLES_RU_DATA)}")

    # ═══════════════════════════════════════════════════════════════
    # Сериализация
    # ═══════════════════════════════════════════════════════════════
    print("\n=== Собираю data.js ===")

    VOCAB_KEYS = ["id", "de", "altDe", "ru", "altRu", "pos", "gender", "plural", "altPlural",
                  "level", "topics", "domen", "group", "note", "new",
                  "comparative", "superlative", "antonym", "derivedFrom",
                  "kind", "case", "digit", "transcription", "context",
                  "type", "modal", "separable", "prefix", "reflexive", "impersonal",
                  "aux", "partizip2", "praeteritum",
                  "priority", "source", "exampleDe", "exampleRu", "quizUse",
                  "strictOrder", "warning", "label"]
    REGEL_KEYS = ["id", "verb", "ru", "forms", "partizip2", "aux", "praeteritum",
                  "separable", "reflexive", "impersonal", "case"]
    CONJ_KEYS = ["id", "verb", "ru", "tense", "modal", "level", "pronouns", "forms",
                 "partizip2", "aux", "praeteritum", "separable", "reflexive",
                 "impersonal", "case"]
    Q_KEYS = ["id", "topic", "level", "difficulty", "type", "q",
              "opts", "ans", "words", "pronouns", "altAns", "hint", "explain"]
    R_KEYS = ["id", "title", "topic", "domen", "group", "level", "content_md",
              "examples", "note", "new"]
    T_KEYS = ["id", "term", "plural", "ru", "topic", "domen", "group", "level",
              "priority", "source", "note"]
    S_KEYS = ["id", "combo", "pronunciation", "example", "translation", "domen", "group", "note"]

    out = []
    out.append("// ═══════════════════════════════════════════════════════════════")
    out.append("// data.js — единый источник правды для шпоры и тренажёра")
    out.append("// Собрано автоматически из database.xlsx (см. update_data.py)")
    out.append("// ═══════════════════════════════════════════════════════════════")
    out.append(f"// VOCAB: {len(all_vocab)} | REGEL: {len(regel_verbs)} | CONJ: {len(conjugations)}")
    out.append(f"// QUESTIONS: {len(questions)} | RULES: {len(rules)} | TERMS: {len(terms)} | SOUNDS: {len(sounds)}")
    out.append("")

    vocab_by_pos = {}
    for v in all_vocab:
        vocab_by_pos.setdefault(v["pos"], []).append(v)
    out.append("const VOCAB = [")
    POS_ORDER = ["num", "noun", "verb", "adj", "adv", "pron", "phrase"]
    POS_LABEL = {"num": "Числа", "noun": "Существительные", "verb": "Глаголы",
                 "adj": "Прилагательные", "adv": "Наречия", "pron": "Местоимения",
                 "phrase": "Фразы"}
    for pos in POS_ORDER:
        if pos in vocab_by_pos:
            out.append(f"  // ─── {POS_LABEL[pos]} ({len(vocab_by_pos[pos])}) ─────────────")
            for item in vocab_by_pos[pos]:
                out.append(item_to_js(item, VOCAB_KEYS))
    out.append("];")
    out.append("")

    for const_name, data, keys in [
        ("CONJUGATIONS", conjugations, CONJ_KEYS),
        ("REGEL_VERBS", regel_verbs, REGEL_KEYS),
        ("QUESTIONS", questions, Q_KEYS),
        ("SOUNDS", sounds, S_KEYS),
        ("TERMS", terms, T_KEYS),
        ("RULES", rules, R_KEYS),
    ]:
        out.append(f"const {const_name} = [")
        for it in data:
            out.append(item_to_js(it, keys))
        out.append("];")
        out.append("")

    # TAXONOMY (из Settings-XLSX)
    out.append("const TAXONOMY = " + js_value(TAXONOMY) + ";")
    out.append("")

    # BLOCKS / TOPIC_TITLES / TAB_TITLES — сгенерированы (немецкий, новая таксономия)
    out.append("const BLOCKS = " + js_value(BLOCKS_DATA) + ";")
    out.append("")
    out.append("const TOPIC_TITLES = " + js_value(TOPIC_TITLES_DATA) + ";")
    out.append("")
    out.append("const TAB_TITLES = " + js_value(TAB_TITLES_DATA) + ";")
    out.append("")
    out.append("const TOPIC_TITLES_RU = " + js_value(TOPIC_TITLES_RU_DATA) + ";")
    out.append("")
    out.append("const TAB_TITLES_RU = " + js_value(TAB_TITLES_RU_DATA) + ";")
    out.append("")

    # PHRASE_UNITS — сохранили из старого data.js (если был)
    if phrase_units_block:
        out.append("const PHRASE_UNITS = " + phrase_units_block + ";")
        out.append("")
    if sentence_templates_block:
        out.append("const SENTENCE_TEMPLATES = " + sentence_templates_block + ";")
    else:
        out.append("const SENTENCE_TEMPLATES = [];")
    out.append("")

    out.append("if (typeof module !== 'undefined' && module.exports) {")
    out.append("  module.exports = { VOCAB, CONJUGATIONS, REGEL_VERBS, QUESTIONS,")
    out.append("    SOUNDS, TERMS, RULES, TAXONOMY, BLOCKS, TOPIC_TITLES, TAB_TITLES,")
    out.append("    TOPIC_TITLES_RU, TAB_TITLES_RU,")
    out.append("    PHRASE_UNITS, SENTENCE_TEMPLATES };")
    out.append("}")
    out.append("")

    OUT_DATA_JS.write_text("\n".join(out), encoding="utf-8")

    # ═══════════════════════════════════════════════════════════════
    # Отчёт
    # ═══════════════════════════════════════════════════════════════
    print(f"\n✓ Записано: {OUT_DATA_JS}")
    print(f"  размер: {OUT_DATA_JS.stat().st_size:,} байт")
    if WARN:
        print(f"\n⚠ Предупреждения ({len(WARN)}):")
        for w in WARN:
            print("   •", w)
    else:
        print("\n✓ Без предупреждений")
