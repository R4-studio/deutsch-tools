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
# БИБЛИОТЕКА SVG-ИКОНОК (единственный источник — эмитится в data.js
# как ICON_SVGS, потребители: cheatsheet.html (ic()) и trainer.html
# (компонент <Icon name=.../>). Stroke-style, viewBox 24x24.
# ═══════════════════════════════════════════════════════════════
ICON_LIBRARY = {
    "document": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 2h8l4 4v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z'/><path d='M14 2v4h4'/></svg>",
    "chat": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 4h16v12H8l-4 4V4z'/></svg>",
    "wave": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M7 11V6a1.5 1.5 0 0 1 3 0v4'/><path d='M10 10V4.5a1.5 1.5 0 0 1 3 0V10'/><path d='M13 10V5.5a1.5 1.5 0 0 1 3 0V11'/><path d='M16 11.5a1.5 1.5 0 0 1 3 0V14a6 6 0 0 1-6 6h-1a6 6 0 0 1-6-6v-1l-2-3a1.3 1.3 0 0 1 2-1.6L7 11'/></svg>",
    "tv": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='5' width='20' height='14' rx='2'/><path d='M8 21h8M12 17v4'/></svg>",
    "broom": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M18 3 8 13'/><path d='M8 13c-2 0-4 2-4 4l6-2z'/><path d='M4 21l4-4'/></svg>",
    "sofa": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 12v6h16v-6'/><path d='M4 12a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2'/><path d='M4 12V9a1 1 0 0 1 1-1h1a1 1 0 0 1 1 1v2M17 12V9a1 1 0 0 1 1-1h1a1 1 0 0 1 1 1v2'/><path d='M4 18v2M20 18v2'/></svg>",
    "target": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><circle cx='12' cy='12' r='5'/><circle cx='12' cy='12' r='1'/></svg>",
    "palette": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 2a10 10 0 1 0 0 20c1.1 0 2-.9 2-2 0-.5-.2-1-.5-1.4-.3-.4-.5-.8-.5-1.3 0-1.1.9-2 2-2h2.3A5.2 5.2 0 0 0 22 10c0-4.4-4.5-8-10-8z'/><circle cx='7' cy='11' r='1'/><circle cx='7.5' cy='7' r='1'/><circle cx='12' cy='5.5' r='1'/><circle cx='16.5' cy='7.5' r='1'/></svg>",
    "carousel": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 2v4M12 8l7 4-7 4-7-4z'/><path d='M5 12v6l7 4 7-4v-6'/></svg>",
    "traffic-light": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='8' y='2' width='8' height='16' rx='3'/><circle cx='12' cy='6' r='1.3'/><circle cx='12' cy='10' r='1.3'/><circle cx='12' cy='14' r='1.3'/><path d='M9 20h6'/></svg>",
    "car": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 17V12l2-6h14l2 6v5'/><path d='M3 17h18'/><circle cx='7' cy='17' r='2'/><circle cx='17' cy='17' r='2'/></svg>",
    "paw": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='7' cy='8' r='2'/><circle cx='12' cy='6' r='2'/><circle cx='17' cy='8' r='2'/><path d='M12 12c-4 0-6 3-6 5.5S8 21 12 21s6-1 6-3.5S16 12 12 12z'/></svg>",
    "mountain": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 20 9 8l4 6 2-3 6 9z'/></svg>",
    "weather-sun": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='4'/><path d='M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4'/></svg>",
    "users": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='9' cy='7' r='3'/><path d='M2 21c0-3.9 3.1-7 7-7s7 3.1 7 7'/><circle cx='17' cy='8' r='2.5'/><path d='M22 21c0-3-2-5.5-5-6.3'/></svg>",
    "user": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='8' r='4'/><path d='M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8'/></svg>",
    "health": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='3' width='18' height='18' rx='2'/><path d='M12 8v8M8 12h8'/></svg>",
    "id-card": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='5' width='20' height='14' rx='2'/><circle cx='8' cy='11' r='2'/><path d='M5 17c0-1.7 1.3-3 3-3s3 1.3 3 3M14 9h5M14 13h5M14 17h3'/></svg>",
    "briefcase": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='7' width='20' height='13' rx='2'/><path d='M8 7V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M2 12h20'/></svg>",
    "couple": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 21s-7-4.5-9.5-9C.7 8.6 2 5 5.5 5c2 0 3.3 1.2 4.5 2.7C11.2 6.2 12.5 5 14.5 5c3.5 0 4.8 3.6 3 7-2.5 4.5-9.5 9-9.5 9z'/></svg>",
    "mask": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='7' cy='12' r='5'/><circle cx='17' cy='12' r='5'/><path d='M5 10.5c1-1 3-1 4 0M15 10.5c1-1 3-1 4 0M5.5 14.5c1 1.5 3 1.5 4 0M14.5 14.5c1 1.5 3 1.5 4 0'/></svg>",
    "brain": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M9 4a3 3 0 0 0-3 3v1a3 3 0 0 0-2 5 3 3 0 0 0 2 5 3 3 0 0 0 5 1V6a3 3 0 0 0-2-2z'/><path d='M15 4a3 3 0 0 1 3 3v1a3 3 0 0 1 2 5 3 3 0 0 1-2 5 3 3 0 0 1-5 1V6a3 3 0 0 1 2-2z'/></svg>",
    "landmark": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 21h18M4 21V10M20 21V10M2 10l10-6 10 6M6 10v6M10 10v6M14 10v6M18 10v6'/></svg>",
    "city": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 21V9l5-4v16M4 21h16M13 21V4l6 3v14M9 8h1M9 12h1M9 16h1M17 10h1M17 14h1M17 18h1'/></svg>",
    "home": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 9.5 12 3l9 6.5'/><path d='M5 8v11a1 1 0 0 0 1 1h4v-6h4v6h4a1 1 0 0 0 1-1V8'/></svg>",
    "door": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='6' y='2' width='12' height='20' rx='1'/><circle cx='14' cy='12' r='1'/></svg>",
    "shirt": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M8 3 3 6l2 4 2-1v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V9l2 1 2-4-5-3a3 3 0 0 1-6 0z'/></svg>",
    "cup": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 8h12l-1.2 11.5a2 2 0 0 1-2 1.5H9.2a2 2 0 0 1-2-1.5L6 8Z'/><path d='M9 3h6l1 5H8z'/></svg>",
    "apple": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 7c-3 0-5 2.5-5 6 0 4 2.5 8 5 8s5-4 5-8c0-3.5-2-6-5-6Z'/><path d='M12 7c0-2 1-4 3-4'/></svg>",
    "pan": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='11' cy='12' r='7'/><path d='M18 10h4'/><path d='M8 10a3 3 0 0 1 6 0'/></svg>",
    "pencil": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 20h9'/><path d='M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z'/></svg>",
    "laptop": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='4' y='4' width='16' height='11' rx='1'/><path d='M2 19h20'/></svg>",
    "cart": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='9' cy='20' r='1.3'/><circle cx='18' cy='20' r='1.3'/><path d='M2 3h2l2.6 12.6a2 2 0 0 0 2 1.6h8.8a2 2 0 0 0 2-1.6L21 8H6'/></svg>",
    "smile": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><path d='M8 14s1.5 2 4 2 4-2 4-2'/><path d='M9 9h.01M15 9h.01'/></svg>",
    "warning": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3 2 20h20L12 3Z'/><path d='M12 10v4M12 17h.01'/></svg>",
    "school": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3 2 8l10 5 10-5-10-5Z'/><path d='M6 10.5V16c0 1.5 3 3 6 3s6-1.5 6-3v-5.5M22 8v6'/></svg>",
    "notes": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 3h13l3 3v15H4z'/><path d='M17 3v4h3M8 12h8M8 16h8M8 8h4'/></svg>",
    "clock": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><path d='M12 7v5l4 2'/></svg>",
    "calendar": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4' width='18' height='17' rx='2'/><path d='M16 2v4M8 2v4M3 10h18'/></svg>",
    "sunrise": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3v5'/><path d='M5.6 10.6 7 12M18.4 10.6 17 12M2 18h20M4 18a8 8 0 0 1 16 0'/></svg>",
    "hourglass": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 2h12M6 22h12M6 2c0 5 4 7 6 8-2 1-6 3-6 8M18 2c0 5-4 7-6 8 2 1 6 3 6 8'/></svg>",
    "stopwatch": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='13' r='8'/><path d='M12 9v4l3 2'/><path d='M9 2h6M12 2v3'/></svg>",
    "office": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='4' y='2' width='16' height='20' rx='1'/><path d='M9 7h1M14 7h1M9 11h1M14 11h1M9 15h1M14 15h1M10 22v-4h4v4'/></svg>",
    "crane": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 21V9l8-5v17'/><path d='M12 6h8l-3 5M17 11v10'/></svg>",
    "walk": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='13' cy='4' r='2'/><path d='M10 21l1-6-3-2 1-5 4-2 3 3h3'/><path d='M9 13l-3 2v6'/></svg>",
    "book": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2Z'/><path d='M19 17H6a2 2 0 0 0-2 2'/></svg>",
    "ring": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='15' r='6'/><path d='M9 9l3-6 3 6'/></svg>",
    "flex": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 14c2-3 4-4 6-4 3 0 3 2 5 2s3-2 3-2'/><path d='M14 8c1-2 3-3 5-2 2.5 1.3 2 5-1 7-3 2.5-8 3-11 1'/></svg>",
    "ruler": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 17 17 3l4 4L7 21z'/><path d='M9 11l2 2M12 8l2 2M6 14l2 2'/></svg>",
    "ruleset": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 3v14a4 4 0 0 0 4 4h14'/><path d='M8 3v10M3 8h10'/></svg>",
    "wood": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='9' width='20' height='6' rx='2'/><path d='M6 9v6M11 9v6M16 9v6'/></svg>",
    "coin": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><path d='M12 7v10M9.5 9.5c0-1.4 1.1-2.2 2.5-2.2s2.5.8 2.5 2c0 2.5-5 1.7-5 4.2 0 1.2 1.1 2 2.5 2s2.5-.8 2.5-2'/></svg>",
    "star": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 2l3 7h7l-5.5 4.5L18.5 21 12 16.5 5.5 21 7.5 13.5 2 9h7Z'/></svg>",
    "handshake": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M2 12h4l3-3 3 3h3'/><path d='M9 12l3 3 6-6 3 3-8 8-6-6'/></svg>",
    "arrow-right": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M5 12h14M13 6l6 6-6 6'/></svg>",
    "pin": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 22s7-7.5 7-13a7 7 0 0 0-14 0c0 5.5 7 13 7 13Z'/><circle cx='12' cy='9' r='2.5'/></svg>",
    "thumbs-up": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M7 22V11l4-9c1.5 0 2 1 2 2v6h6a2 2 0 0 1 2 2.3l-1.4 7A2 2 0 0 1 17.6 22H7Z'/><path d='M7 11H3v11h4'/></svg>",
    "repeat": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M17 2l4 4-4 4'/><path d='M3 11V9a4 4 0 0 1 4-4h14'/><path d='M7 22l-4-4 4-4'/><path d='M21 13v2a4 4 0 0 1-4 4H3'/></svg>",
    "graduation": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M2 9 12 4l10 5-10 5z'/><path d='M6 11v5c0 1.5 3 3 6 3s6-1.5 6-3v-5M22 9v6'/></svg>",
    "newspaper": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='5' width='13' height='16' rx='1'/><path d='M16 8h5v11a2 2 0 0 1-2 2H5M6 9h6M6 13h6M6 17h6'/></svg>",
    "key": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='7' cy='15' r='4'/><path d='M10 12l9-9M16 6l2 2M19 3l2 2'/></svg>",
    "fog": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 8h16M2 12h20M4 16h16M6 20h12'/></svg>",
    "ban": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><path d='M6 6l12 12'/></svg>",
    "link": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M9 17H7a5 5 0 0 1 0-10h2M15 7h2a5 5 0 0 1 0 10h-2M8 12h8'/></svg>",
    "medal": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='15' r='6'/><path d='M9 10 6 3M15 10l3-7M9 3h6'/><path d='M12 12v6'/></svg>",
    "scroll": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 4a2 2 0 0 0-2 2v1a2 2 0 0 0 2 2M6 4h13a2 2 0 0 1 2 2v1a2 2 0 0 1-2 2H6M6 4v14M6 18a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-1a2 2 0 0 0-2-2H6'/></svg>",
    "letter-a": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 20 10 4h4l6 16'/><path d='M7.5 14h9'/></svg>",
    "letter-b": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 4h8a4 4 0 0 1 0 8H6zM6 12h9a4 4 0 0 1 0 8H6Z'/></svg>",
    "hash": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M5 9h14M5 15h14M10 3 8 21M16 3l-2 18'/></svg>",
    "box": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M21 8 12 3 3 8l9 5 9-5Z'/><path d='M3 8v9l9 5 9-5V8M12 13v9'/></svg>",
    "bolt": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M13 2 4 14h6l-1 8 9-12h-6z'/></svg>",
    "book-open": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z'/><path d='M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z'/></svg>",
    "waveform": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M2 12h2l2-7 3 14 3-10 2 6 2-3h6'/></svg>",
    "book-check": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2Z'/><path d='M8 9l2 2 4-4'/></svg>",
    "flame": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 2c1 3-3 4-3 8a3 3 0 0 0 6 0c1 1 2 2 2 4a5 5 0 0 1-10 0c0-5 5-6 5-12Z'/></svg>",
    "sparkle": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 2l1.5 5.5L19 9l-5.5 1.5L12 16l-1.5-5.5L5 9l5.5-1.5Z'/></svg>",
    "shuffle": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M2 6h4l12 12'/><path d='M14 6h6M18 3l4 3-4 3'/><path d='M2 18h4l3-5'/></svg>",
    "question": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='9'/><path d='M9.5 9a2.5 2.5 0 0 1 5 0c0 2-2.5 2-2.5 4.5'/><path d='M12 17h.01'/></svg>",
    "spa": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 21c4-2 7-5 7-10a7 7 0 0 0-7-7 7 7 0 0 0-7 7c0 5 3 8 7 10Z'/><path d='M12 21V9'/></svg>",
    "bulb": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M9 18h6M10 21h4'/><path d='M12 3a6 6 0 0 0-4 10.5c.7.6 1 1.5 1 2.5h6c0-1 .3-1.9 1-2.5A6 6 0 0 0 12 3Z'/></svg>",
    "bus": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='5' width='18' height='12' rx='2'/><path d='M3 11h18'/><circle cx='7.5' cy='17.5' r='1.5'/><circle cx='16.5' cy='17.5' r='1.5'/></svg>",
    "gamepad": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='7' width='20' height='11' rx='5'/><path d='M7 10v4M5 12h4'/><circle cx='16' cy='11' r='1'/><circle cx='18' cy='14' r='1'/></svg>",
    "cloud-rain": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M17 13a4 4 0 0 0-1-7.9A6 6 0 0 0 4.5 9 4 4 0 0 0 5 17h12Z'/><path d='M8 19l-1 2M12 19l-1 2M16 19l-1 2'/></svg>",
    "music": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M9 18V5l11-2v13'/><circle cx='6' cy='18' r='3'/><circle cx='17' cy='16' r='3'/></svg>",
    "wine": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M8 3h8l-1 8a3 3 0 0 1-6 0Z'/><path d='M12 14v7M9 21h6'/></svg>",
    "bed": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M3 18v-6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v6'/><path d='M3 18v3M21 18v3M3 12V6h6v6'/></svg>",
    "monitor": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4' width='18' height='12' rx='1'/><path d='M8 20h8M12 16v4'/></svg>",
    "archive": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4' width='18' height='4' rx='1'/><path d='M4 8v10a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8M10 13h4'/></svg>",
    "shower": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 12a8 8 0 0 1 14-5'/><path d='M18 3l3 3'/><path d='M4 12h16'/><path d='M8 16v2M12 16v2M16 16v2M6 20v1M10 20v1M14 20v1M18 20v1'/></svg>",
    "toy": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='8' r='4'/><path d='M9 6l-1-2M15 6l1-2'/><path d='M6 21c0-4 3-7 6-7s6 3 6 7'/></svg>",
    "leaf": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M11 20A7 7 0 0 1 4 13c0-6 5-10 15-11 0 10-4 15-11 15Z'/><path d='M4 20c3-3 5-6 15-15'/></svg>",
    "dining": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M7 2v9M4 2v5a2 2 0 0 0 2 2M10 2v5a2 2 0 0 1-2 2M17 2c-2 0-3 2-3 5s1 4 3 4M17 2v20'/></svg>",
    "bag": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 7h12l1 14H5Z'/><path d='M9 7a3 3 0 0 1 6 0'/></svg>",
    "thermometer": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3a2 2 0 0 0-2 2v9a4 4 0 1 0 4 0V5a2 2 0 0 0-2-2Z'/><path d='M12 14v-6'/></svg>",
    "chart": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M4 20V10M10 20V4M16 20v-7M22 20H2'/></svg>",
    "history": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='13' r='8'/><path d='M12 9v4l3 2'/><path d='M3 8a9 9 0 0 1 2-3M3 8V4M3 8h4'/></svg>",
    "scale": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M12 3v18M5 8h14'/><path d='M5 8l-3 6a3 3 0 0 0 6 0ZM19 8l-3 6a3 3 0 0 0 6 0Z'/><path d='M8 21h8'/></svg>",
    "clipboard": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='5' y='4' width='14' height='17' rx='2'/><rect x='9' y='2' width='6' height='4' rx='1'/><path d='M9 12h6M9 16h6'/></svg>",
    "check": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M20 6 9 17l-5-5'/></svg>",
    "cross": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M18 6 6 18M6 6l12 12'/></svg>",
    "wrench": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M14.7 6.3a4 4 0 1 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.8 2.8-2-2Z'/></svg>",
    "gear": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='3'/><path d='M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1'/></svg>",
    "rewind": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M11 19 2 12l9-7v14Z'/><path d='M22 19 13 12l9-7v14Z'/></svg>",
    "dice": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='3' width='18' height='18' rx='3'/><circle cx='8' cy='8' r='1.3'/><circle cx='16' cy='8' r='1.3'/><circle cx='8' cy='16' r='1.3'/><circle cx='16' cy='16' r='1.3'/><circle cx='12' cy='12' r='1.3'/></svg>",
    "party": "<svg width='1em' height='1em' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M5.8 11.3 2 22l10.7-3.8'/><path d='M11.5 12.5c2-2 5-6 8.5-8.5'/><path d='M16 8c-2 2-6 5-8.5 8.5'/><path d='M4 3h.01M22 8h.01M15 2h.01M22 20h.01M22 14h.01'/></svg>",
}

# ═══════════════════════════════════════════════════════════════
# НЕМЕЦКИЕ ПОДПИСИ ДЛЯ ФРОНТА (BLOCKS / TOPIC_TITLES / TAB_TITLES)
# ─────────────────────────────────────────────────────────────────
# Маппинг ключей domen:group → (icon_key, deutsch). icon_key — ключ
# в ICON_LIBRARY (SVG вместо эмодзи). Хардкод-фоллбэк.
# Если в Settings-XLSX появятся колонки label_de/icon — они
# автоматически переопределят значения отсюда (см. label_for_pair).
# Дописывай сюда новые пары при расширении таксономии.
# ═══════════════════════════════════════════════════════════════
GERMAN_LABELS = {
    # ─── nouns ─────────────────────────────────────────────────
    "communication:docs":       ("document", "Dokumente"),
    "communication:general":    ("chat", "Kommunikation"),
    "communication:greetings":  ("wave", "Begrüßung"),
    "communication:media":      ("tv", "Medien"),
    "home:cleaning":            ("broom", "Putzen"),
    "home:furniture":           ("sofa", "Möbel"),
    "leisure:activities":       ("target", "Aktivitäten"),
    "leisure:hobby":            ("palette", "Hobby"),
    "leisure:places":           ("carousel", "Freizeitorte"),
    "movement:traffic":         ("traffic-light", "Verkehr"),
    "movement:transport":       ("car", "Transport"),
    "nature:animals":           ("paw", "Tiere"),
    "nature:landscape":         ("mountain", "Landschaft"),
    "nature:weather":           ("weather-sun", "Wetter"),
    "people:family":            ("users", "Familie"),
    "people:health":            ("health", "Gesundheit"),
    "people:identity":          ("id-card", "Identität"),
    "people:professions":       ("briefcase", "Berufe"),
    "people:relations":         ("couple", "Beziehungen"),
    "people:roles":             ("mask", "Rollen"),
    "people:traits":            ("brain", "Eigenschaften"),
    "place:buildings":          ("landmark", "Gebäude"),
    "place:city":               ("city", "Stadt"),
    "place:housing":            ("home", "Wohnung"),
    "place:rooms":              ("door", "Zimmer"),
    "products:clothing":        ("shirt", "Kleidung"),
    "products:drinks":          ("cup", "Getränke"),
    "products:food":            ("apple", "Lebensmittel"),
    "products:kitchenware":     ("pan", "Geschirr"),
    "products:stationery":      ("pencil", "Schreibwaren"),
    "products:tech":            ("laptop", "Technik"),
    "shopping:general":         ("cart", "Einkaufen"),
    "state:emotions":           ("smile", "Gefühle"),
    "state:safety":             ("warning", "Sicherheit"),
    "study:places":             ("school", "Lernorte"),
    "study:writing":            ("notes", "Schreiben"),
    "time:clock":               ("clock", "Uhrzeit"),
    "time:dates":               ("calendar", "Datum"),
    "time:dayparts":            ("sunrise", "Tageszeiten"),
    "time:periods":             ("hourglass", "Zeitabschnitte"),
    "work:general":             ("briefcase", "Arbeit"),
    "work:office":              ("office", "Büro"),
    "work:places":              ("crane", "Arbeitsorte"),
    # ─── verbs (доп.) ──────────────────────────────────────────
    "grammar:modal":            ("target", "Modalverben"),
    "home:chores":              ("broom", "Hausarbeit"),
    "movement:general":         ("walk", "Bewegung"),
    "products:cooking":         ("pan", "Kochen"),
    "study:general":            ("book", "Lernen"),
    "time:duration":            ("stopwatch", "Dauer"),
    # ─── adjectives ────────────────────────────────────────────
    "communication:style":      ("chat", "Stil"),
    "movement:speed":           ("bolt", "Tempo"),
    "people:marital":           ("ring", "Familienstand"),
    "people:physical":          ("flex", "Aussehen"),
    "place:size":               ("ruler", "Größe"),
    "products:materials":       ("wood", "Materialien"),
    "shopping:price":           ("coin", "Preis"),
    "state:colors":             ("palette", "Farben"),
    "state:qualities":          ("star", "Qualität"),
    "time:chronology":          ("calendar", "Zeitfolge"),
    # ─── adverbs ───────────────────────────────────────────────
    "communication:agreement":  ("handshake", "Zustimmung"),
    "movement:direction":       ("arrow-right", "Richtung"),
    "place:location":           ("pin", "Ort"),
    "state:certainty":          ("target", "Sicherheit"),
    "state:feelings":           ("thumbs-up", "Befinden"),
    "time:frequency":           ("repeat", "Häufigkeit"),
    # ─── phrases (доп.) ────────────────────────────────────────
    "communication:questions":  ("question", "Fragen"),
    "communication:wellbeing":  ("spa", "Befinden"),
    "study:classroom":          ("graduation", "Im Unterricht"),
    # ─── rules (домен grammar) ─────────────────────────────────
    "grammar:pronouns":         ("user", "Pronomen"),
    "grammar:wordorder":        ("ruleset", "Satzbau"),
    "grammar:verbs":            ("bolt", "Verben"),
    "grammar:cases":            ("target", "Kasus"),
    "grammar:articles":         ("newspaper", "Artikel"),
    "grammar:wfragen":          ("question", "W-Fragen"),
    "grammar:adjectives":       ("palette", "Adjektive"),
    # ─── pronouns (домен grammar / специальные) ────────────────
    "grammar:personal":         ("user", "Personalpronomen"),
    "grammar:possessive":       ("key", "Possessivpronomen"),
    "grammar:indefinite":       ("fog", "Indefinitpronomen"),
    "grammar:negation":         ("ban", "Negation"),
    "grammar:particles":        ("target", "Partikeln"),
    "grammar:prepositions":     ("link", "Präpositionen"),
    # ─── numbers (доп.) ────────────────────────────────────────
    "quantity:ordinal":         ("medal", "Ordinalzahlen"),
    # ─── phonetics (звуки) ─────────────────────────────────────
    "phonetics:rules":          ("scroll", "Regeln"),
    "phonetics:vowels":         ("letter-a", "Vokale"),
    "phonetics:consonants":     ("letter-b", "Konsonanten"),
    # ─── специальные (не из Settings) ──────────────────────────
    "quantity:cardinal":        ("hash", "Kardinalzahlen"),
    "pron:personal":            ("user", "Personalpronomen"),
    "pron:negation":            ("ban", "Negation"),
    "new:new-nouns":            ("box", "Neue Substantive"),
    "new:new-verbs":             ("bolt", "Neue Verben"),
    "new:new-adj":               ("palette", "Neue Adjektive"),
    "new:new-adv":               ("stopwatch", "Neue Adverbien"),
    "new:new-phrases":           ("chat", "Neue Phrasen"),
    "new:new-pron":              ("user", "Neue Pronomen"),
    "new:new-nums":              ("hash", "Neue Zahlen"),
    "new:new-terms":             ("book-open", "Neue Begriffe"),
    "new:new-sounds":            ("waveform", "Neue Laute"),
    "new:new-rules":             ("ruleset", "Neue Regeln"),
    "perfekt:regel":            ("book-check", "Regelmäßig"),
    "perfekt:unregel":          ("flame", "Unregelmäßig"),
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
    "neu":     {"label": "Neu",            "icon": "sparkle", "color": "var(--c-block-neu)", "desc": "Frische Wörter (Puffer)"},
    "verbs":   {"label": "Verben",         "icon": "bolt",    "color": "var(--c-block-verben)"},
    "nouns":   {"label": "Substantive",    "icon": "box",     "color": "var(--c-block-substantive)"},
    "adj":     {"label": "Adjektive",      "icon": "palette", "color": "var(--c-block-adj)"},
    "adv":     {"label": "Adverbien",      "icon": "shuffle", "color": "var(--c-block-adv)"},
    "mestoim": {"label": "Pronomen",       "icon": "user",    "color": "var(--c-block-pronomen)"},
    "phrases": {"label": "Redewendungen",  "icon": "chat",    "color": "var(--c-block-redewendungen)"},
    "unregel": {"label": "Unregelmäßig",   "icon": "flame",   "color": "var(--c-block-unregel)", "desc": "Unregelmäßige Verben in Präsens", "kind": "conjugations"},
    "nums":    {"label": "Zahlen",         "icon": "hash",    "color": "#16a085"},
    "sounds":  {"label": "Aussprache",     "icon": "waveform", "color": "#9b59b6", "desc": "Ausspracheregeln", "kind": "sounds"},
}

def label_for_pair(domen, group, taxonomy_entry=None):
    """Возвращает {"icon": icon_key, "label": Подпись} для domen:group.
    Override через колонки label_de/label/icon в Settings-XLSX (icon —
    ключ в ICON_LIBRARY, не эмодзи)."""
    if taxonomy_entry:
        de_label = taxonomy_entry.get("label_de") or taxonomy_entry.get("label")
        icon = taxonomy_entry.get("icon")
        if de_label:
            return {"icon": icon or "", "label": de_label}
    key = f"{domen}:{group}"
    icon, name = GERMAN_LABELS.get(key, ("", key))
    return {"icon": icon, "label": name}

def build_block_from_pairs(bid, taxonomy_pairs, used_topics=None):
    """Блок POS: каждая пара domen:group → подблок.
    Если передан used_topics — отфильтровываем подблоки без слов."""
    meta = BLOCK_META[bid]
    block = {"id": bid, "label": meta["label"], "icon": meta.get("icon", ""), "color": meta["color"]}
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
        lp = label_for_pair(d, g, p)
        subs.append({
            "id": f"{d}-{g}".replace("/", "-"),
            "label": lp["label"],
            "icon": lp["icon"],
            "topics": [topic],
        })
    if subs: block["subblocks"] = subs
    return block

def build_special_block(bid, subblocks=None):
    """Спец. блок (kind=sounds/terms/rules/conjugations/perfekt или с фикс. подблоками)."""
    meta = BLOCK_META[bid]
    block = {"id": bid, "label": meta["label"], "icon": meta.get("icon", ""), "color": meta["color"]}
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
      используется только в справочнике.
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
        "icon": BLOCK_META["unregel"].get("icon", ""),
        "color": BLOCK_META["unregel"]["color"],
        "desc": BLOCK_META["unregel"].get("desc", ""),
        "kind": "conjugations",
    }
    verbs_block["subblocks"] = [unregel_sub] + verbs_block.get("subblocks", [])

    return [
        # 1. Neu (буфер новых)
        build_special_block("neu", subblocks=[
            {"id": "nouns", "label": "Substantive", "icon": "box",       "topics": ["new:new-nouns"]},
            {"id": "verbs", "label": "Verben",      "icon": "bolt",      "topics": ["new:new-verbs"]},
            {"id": "adj",   "label": "Adjektive",   "icon": "palette",   "topics": ["new:new-adj"]},
            {"id": "adv",   "label": "Adverbien",   "icon": "stopwatch", "topics": ["new:new-adv"]},
        ]),
        # 2. Aussprache (kind=sounds, только для справочника; trainer фильтрует)
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
    """TOPIC_TITLES — заголовки секций в справочнике. Каждое значение — {icon, label}."""
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
        icon, name = GERMAN_LABELS.get(k, ("", k))
        out[k] = {"icon": icon, "label": name}
    return out

def build_tab_titles():
    return {bid: {"icon": meta.get("icon", ""), "label": meta["label"]} for bid, meta in BLOCK_META.items()}

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
    out.append("// data.js — единый источник правды для справочника и тренажёра")
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

    # ICON_SVGS — единственный источник SVG-иконок (потребители: cheatsheet.html, trainer.html)
    out.append("const ICON_SVGS = " + js_value(ICON_LIBRARY) + ";")
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
