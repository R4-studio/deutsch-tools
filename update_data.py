"""
update_data.py — собирает data.js из database.xlsx

Запуск: python update_data.py

Что делает:
1. Читает все листы database.xlsx (nouns, verbs, adjectives, adverbs, pronouns,
   numbers, phrases, sounds, terms, rules, questions).
2. Для verbs: regel без форм → автогенератор; regel с формами → ручные.
   Для unregel → берёт формы из ячеек.
3. Сохраняет BLOCKS, TOPIC_TITLES, TAB_TITLES, PHRASE_UNITS, SENTENCE_TEMPLATES
   из существующего data.js (не пересоздаёт их).
4. Финальный файл data.js пишется заново.

ВАЖНО:
- ID из xlsx игнорируются. Конвертер генерирует новые (n0001, v0001, ...).
- Связь между сущностями — по тексту слова (`de` или `verb`).
- BLOCKS, TOPIC_TITLES, TAB_TITLES, PHRASE_UNITS НЕ меняются конвертером —
  они часть кода тренажёра, не контент.
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

# ═══════════════════════════════════════════════════════════════
# АВТОГЕНЕРАТОР Präsens (regelmäßig)
# Дублирует логику из trainer.html для использования при сборке data.js
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

def conj_all_forms(infinitive):
    """Возвращает 6 форм Präsens: [ich, du, er, wir, ihr, sie]"""
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
    """
    Правила:
    - На -ieren: без ge-, stem + t (studieren → studiert)
    - Неотделяемые приставки: без ge- (besuchen → besucht)
    - Отделяемые приставки: prefix + ge + stem + t (aufmachen → aufgemacht)
    - Стандарт: ge + stem + t/et (machen → gemacht, arbeiten → gearbeitet)
    """
    # На -ieren
    if infinitive.endswith("ieren"):
        return infinitive[:-2] + "t"  # studieren → studiert

    # Stem
    if infinitive.endswith("en"):
        stem = infinitive[:-2]
    elif infinitive.endswith("n"):
        stem = infinitive[:-1]
    else:
        stem = infinitive

    # Окончание (с -et или -t)
    needs_et = bool(re.search(r"[td]$", stem)) or \
               bool(re.search(r"(chn|ffn|tm|dm|gn)$", stem))
    ending = "et" if needs_et else "t"

    # Неотделяемая приставка → без ge-
    for prefix in INSEPARABLE_PREFIXES:
        if infinitive.startswith(prefix) and len(infinitive) > len(prefix) + 2:
            # дальше должна быть хотя бы основа из 2+ букв
            return stem + ending

    # Отделяемая приставка → prefix + ge + stem-без-prefix + ending
    for prefix in SEPARABLE_PREFIXES:
        if infinitive.startswith(prefix) and len(infinitive) > len(prefix) + 2:
            sub_inf = infinitive[len(prefix):]
            sub_p2 = partizip2_regelmaessig(sub_inf)
            return prefix + sub_p2

    # Стандарт
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
        return "{ " + ", ".join(f"{k}: {js_value(val)}" for k, val in v.items()) + " }"
    return js_str(v)

def item_to_js(item, key_order=None):
    if key_order:
        keys = [k for k in key_order if k in item]
    else:
        keys = list(item.keys())
    parts = [f"{k}: {js_value(item[k])}" for k in keys]
    return "  { " + ", ".join(parts) + " },"

# ═══════════════════════════════════════════════════════════════
# Утилиты для чтения xlsx
# ═══════════════════════════════════════════════════════════════
def get_cell(row, idx):
    """Безопасно достать ячейку (None если короче)"""
    if idx >= len(row): return None
    v = row[idx]
    if v is None: return None
    s = str(v).strip()
    return s if s else None

def is_true(v):
    """Конверсия значения в bool (TRUE / true / 1 / X → True)"""
    if v is None: return False
    return str(v).strip().lower() in ("true", "1", "x", "yes", "да")

def read_sheet(wb, name, expected_headers=None):
    """Читает лист, пропуская строку легенды (строка 2 с 'обяз./опц.')."""
    if name not in wb.sheetnames:
        print(f"  ⚠ Лист '{name}' не найден")
        return []
    ws = wb[name]
    headers = [c.value for c in ws[1]]
    if expected_headers:
        missing = set(expected_headers) - set(headers)
        if missing:
            print(f"  ⚠ В '{name}' нет колонок: {missing}")
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Пропустить строку легенды (где все значения "опц." / "обяз.")
        non_empty = [v for v in row if v is not None]
        if non_empty and all(str(v).strip() in ("опц.", "обяз.") for v in non_empty):
            continue
        # Пропустить полностью пустые строки
        if not any(v is not None and str(v).strip() for v in row):
            continue
        rows.append(dict(zip(headers, row)))
    return rows

# ═══════════════════════════════════════════════════════════════
# ОБРАБОТКА КАЖДОГО ЛИСТА
# ═══════════════════════════════════════════════════════════════
def process_nouns(rows):
    items = []
    seen = set()  # ключ: (de, topic) — позволяет одно слово в разных темах
    for r in rows:
        de_raw = r.get("de")
        if not de_raw: continue
        de = str(de_raw).strip()

        # altDe из двух источников:
        # 1) явная колонка alt-de (приоритет — для будущих записей xlsx)
        # 2) инлайн "X / Y" в поле de (legacy: 14 текущих записей)
        alt_de = r.get("alt-de")
        if alt_de is not None and str(alt_de).strip():
            alt_de = str(alt_de).strip()
        elif " / " in de:
            parts = [p.strip() for p in de.split(" / ")]
            de = parts[0]
            alt_de = " / ".join(parts[1:]) if len(parts) > 1 else None
        else:
            alt_de = None

        topic = str(r.get("topic") or "").strip()
        key = (de, topic)
        if key in seen:
            print(f"  ⚠ Дубль noun: {de} ({topic or 'без topic'})")
            continue
        seen.add(key)
        item = {"de": de}
        if alt_de: item["altDe"] = alt_de
        if r.get("gender"): item["gender"] = str(r["gender"]).strip()
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        if r.get("plural"): item["plural"] = str(r["plural"]).strip()
        item["topics"] = [topic] if topic else []
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        if is_true(r.get("new")):
            item["new"] = True
            if "new:new-nouns" not in item["topics"]:
                item["topics"] = item["topics"] + ["new:new-nouns"]
        item["pos"] = "noun"
        items.append(item)
    return items

def process_verbs(rows):
    """
    Возвращает три списка:
    - vocab_items: записи для VOCAB (vocab-блок: id, de=verb, ru, topic, level, pos, new)
    - regel_verbs: для REGEL_VERBS (regel + composite где есть формы)
    - conjugations: для CONJUGATIONS (unregel со всеми формами)
    """
    vocab_items = []
    regel_verbs = []
    conjugations = []
    seen = set()
    missing_p2 = []   # unregel без partizip2
    missing_aux = []  # unregel без aux

    for r in rows:
        verb = r.get("verb")
        if not verb: continue
        verb = str(verb).strip()
        if verb in seen:
            print(f"  ⚠ Дубль verb: {verb}")
            continue
        seen.add(verb)

        vtype = (r.get("type") or "").strip().lower()
        if vtype not in ("regel", "unregel", "composite"):
            vtype = "regel"  # дефолт

        # VOCAB запись
        vi = {"de": verb, "pos": "verb"}
        if r.get("ru"): vi["ru"] = str(r["ru"]).strip()
        vi["topics"] = [str(r["topic"]).strip()] if r.get("topic") else []
        if r.get("level"): vi["level"] = str(r["level"]).strip()
        if r.get("note"): vi["note"] = str(r["note"]).strip()
        if is_true(r.get("new")):
            vi["new"] = True
            if "new:new-verbs" not in vi["topics"]:
                vi["topics"] = vi["topics"] + ["new:new-verbs"]
        vocab_items.append(vi)

        # Composite (Fahrrad fahren) — только в VOCAB, не в REGEL/CONJUGATIONS
        if vtype == "composite":
            continue

        # Достаём формы и метаданные
        forms = [get_cell(list(r.values()), i) for i in (7, 8, 9, 10, 11, 12)]
        # Через словарь надёжнее:
        forms = [r.get("ich"), r.get("du"), r.get("er/sie/es"),
                 r.get("wir"), r.get("ihr"), r.get("sie/Sie")]
        forms = [str(f).strip() if f else "" for f in forms]
        has_all_forms = all(forms)

        aux = (r.get("aux") or "").strip().lower() if r.get("aux") else ""
        if aux not in ("haben", "sein", ""):
            aux = ""
        p2 = (r.get("partizip2") or "").strip() if r.get("partizip2") else ""
        modal = is_true(r.get("modal"))

        if vtype == "unregel":
            # Для unregel — формы обязательны
            if not has_all_forms:
                print(f"  ⚠ unregel без полных форм Präsens: {verb}")
                # генерируем автогенератором как fallback (не идеально, но не пусто)
                forms = conj_all_forms(verb)
            if not p2:
                missing_p2.append(verb)
                # ничего не подставляем — будет null
            if not aux:
                missing_aux.append(verb)
            conj = {
                "verb": verb,
                "ru": str(r.get("ru") or "").strip(),
                "tense": "Präsens",
                "pronouns": ["ich", "du", "er/sie/es", "wir", "ihr", "sie/Sie"],
                "forms": forms,
            }
            if modal: conj["modal"] = True
            if p2: conj["partizip2"] = p2
            if aux: conj["aux"] = aux
            if r.get("level"): conj["level"] = str(r["level"]).strip()
            conjugations.append(conj)
        else:
            # regel: если форм нет — генерируем; если есть — берём из xlsx
            if not has_all_forms:
                forms = conj_all_forms(verb)
            # Partizip II: если нет — генерируем
            if not p2:
                p2 = partizip2_regelmaessig(verb)
            # aux: по умолчанию haben (большинство regel)
            if not aux:
                aux = "haben"
            reg = {
                "verb": verb,
                "ru": str(r.get("ru") or "").strip(),
                "forms": forms,
                "partizip2": p2,
                "aux": aux,
            }
            regel_verbs.append(reg)

    if missing_p2:
        print(f"  ⚠ unregel глаголов без partizip2 ({len(missing_p2)}): {missing_p2[:8]}")
        if len(missing_p2) > 8:
            print(f"    ... ещё {len(missing_p2)-8}")
    if missing_aux:
        print(f"  ⚠ unregel глаголов без aux ({len(missing_aux)}): {missing_aux[:8]}")

    return vocab_items, regel_verbs, conjugations

def process_adjectives(rows):
    items = []
    seen = set()
    for r in rows:
        de = r.get("de")
        if not de: continue
        de = str(de).strip()
        if de in seen:
            print(f"  ⚠ Дубль adj: {de}")
            continue
        seen.add(de)
        item = {"de": de, "pos": "adj"}
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        if r.get("comparative"): item["comparative"] = str(r["comparative"]).strip()
        if r.get("superlative"): item["superlative"] = str(r["superlative"]).strip()
        item["topics"] = [str(r["topic"]).strip()] if r.get("topic") else []
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("antonym"): item["antonym"] = str(r["antonym"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        if is_true(r.get("new")):
            item["new"] = True
            if "new:new-adj-adv" not in item["topics"]:
                item["topics"] = item["topics"] + ["new:new-adj-adv"]
        items.append(item)
    return items

def process_adverbs(rows):
    items = []
    for r in rows:
        de = r.get("de")
        if not de: continue
        item = {"de": str(de).strip(), "pos": "adv"}
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        item["topics"] = [str(r["topic"]).strip()] if r.get("topic") else []
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        if is_true(r.get("new")):
            item["new"] = True
            if "new:new-adj-adv" not in item["topics"]:
                item["topics"] = item["topics"] + ["new:new-adj-adv"]
        items.append(item)
    return items

def process_pronouns(rows):
    items = []
    for r in rows:
        de = r.get("de")
        if not de: continue
        item = {"de": str(de).strip(), "pos": "pron"}
        item["topics"] = []  # для совместимости с кодом который везде делает .topics.some()
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        if r.get("kind"): item["kind"] = str(r["kind"]).strip()
        if r.get("case"): item["case"] = str(r["case"]).strip()
        if r.get("gender"): item["gender"] = str(r["gender"]).strip()
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

def process_numbers(rows):
    items = []
    for r in rows:
        de = r.get("de")
        if not de: continue
        item = {"de": str(de).strip(), "pos": "num"}
        kind = str(r.get("kind") or "").strip()
        item["topics"] = ["nums:" + kind] if kind else ["nums:basic"]
        if r.get("digit") is not None: item["digit"] = str(r["digit"]).strip()
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        if kind: item["kind"] = kind
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

def process_phrases(rows):
    items = []
    for r in rows:
        de = r.get("de")
        if not de: continue
        item = {"de": str(de).strip(), "pos": "phrase"}
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        item["topics"] = [str(r["topic"]).strip()] if r.get("topic") else []
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("context"): item["context"] = str(r["context"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

def process_sounds(rows):
    items = []
    for r in rows:
        combo = r.get("combo")
        if not combo: continue
        item = {"combo": str(combo).strip()}
        if r.get("pronunciation"): item["pronunciation"] = str(r["pronunciation"]).strip()
        if r.get("example"): item["example"] = str(r["example"]).strip()
        if r.get("translation"): item["translation"] = str(r["translation"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

def process_terms(rows):
    items = []
    for r in rows:
        term = r.get("term")
        if not term: continue
        item = {"term": str(term).strip()}
        if r.get("ru"): item["ru"] = str(r["ru"]).strip()
        if r.get("topic"): item["topic"] = str(r["topic"]).strip()
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

def process_rules(rows):
    items = []
    for r in rows:
        title = r.get("title")
        if not title: continue
        item = {"title": str(title).strip()}
        if r.get("topic"): item["topic"] = str(r["topic"]).strip()
        if r.get("level"): item["level"] = str(r["level"]).strip()
        if r.get("content_md"): item["content_md"] = str(r["content_md"]).strip()
        if r.get("examples"): item["examples"] = str(r["examples"]).strip()
        if r.get("note"): item["note"] = str(r["note"]).strip()
        items.append(item)
    return items

DIFFICULTY_DEFAULTS = {"mc": 1, "fill": 4, "tiles": 4, "conj": 3, "open": 5}

def process_questions(rows):
    items = []
    for r in rows:
        topic = r.get("topic")
        qtype_raw = r.get("type")
        q = r.get("q")
        if not (topic and qtype_raw and q): continue
        qtype = str(qtype_raw).strip()
        item = {"topic": str(topic).strip(), "type": qtype, "q": str(q).strip()}
        # level = A1/A2 (языковой), difficulty = 1-5 (уровень в тренажёре)
        level = r.get("level")
        if level is not None:
            lvl_str = str(level).strip()
            if lvl_str:
                item["level"] = lvl_str
        difficulty = r.get("difficulty")
        if difficulty is not None and str(difficulty).strip() != "":
            try:
                item["difficulty"] = int(float(difficulty))
            except (ValueError, TypeError):
                item["difficulty"] = DIFFICULTY_DEFAULTS.get(qtype, 1)
        else:
            item["difficulty"] = DIFFICULTY_DEFAULTS.get(qtype, 1)
        # Разбор по типу
        if qtype == "mc":
            opts = r.get("opts")
            if opts:
                item["opts"] = [s.strip() for s in str(opts).split("|")]
            if r.get("ans_mc") is not None:
                try:
                    item["ans"] = int(r["ans_mc"])
                except (ValueError, TypeError):
                    pass
        elif qtype == "fill":
            if r.get("answer"): item["ans"] = str(r["answer"]).strip()
            if r.get("altAns"):
                item["altAns"] = [s.strip() for s in str(r["altAns"]).split("|")]
        elif qtype == "tiles":
            if r.get("words"):
                item["words"] = [s.strip() for s in str(r["words"]).split("|")]
        elif qtype == "conj":
            if r.get("pronouns"):
                item["pronouns"] = [s.strip() for s in str(r["pronouns"]).split("|")]
            if r.get("forms"):
                item["ans"] = [s.strip() for s in str(r["forms"]).split("|")]
        if r.get("hint"): item["hint"] = str(r["hint"]).strip()
        if r.get("explain"): item["explain"] = str(r["explain"]).strip()
        items.append(item)
    return items

# ═══════════════════════════════════════════════════════════════
# СОХРАНЕНИЕ КОНСТАНТ ИЗ СТАРОГО data.js
# (BLOCKS, TOPIC_TITLES, TAB_TITLES, PHRASE_UNITS, SENTENCE_TEMPLATES)
# ═══════════════════════════════════════════════════════════════
def extract_block(content, marker):
    """
    Извлекает блок const XXX = [...] или { ... } из data.js.
    marker — имя константы (например 'BLOCKS').
    """
    pattern = re.compile(
        r"const " + re.escape(marker) + r" = ([\[\{][\s\S]*?[\]\}]);",
        re.MULTILINE
    )
    # Это слишком жадно, нужен правильный балансировщик
    # Простая версия — ищем "const XXX = " и читаем до закрывающей скобки на уровне 0
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
                if content[pos] == '\\': pos += 2
                else: pos += 1
        elif c == "'":
            pos += 1
            while pos < len(content) and content[pos] != "'":
                if content[pos] == '\\': pos += 2
                else: pos += 1
        pos += 1
    return None

# ═══════════════════════════════════════════════════════════════
# ОСНОВНОЙ ПРОЦЕСС
# ═══════════════════════════════════════════════════════════════
print(f"Читаю {XLSX_PATH}...")
wb = load_workbook(XLSX_PATH, data_only=True)

print("\n=== nouns ===")
nouns_rows = read_sheet(wb, "nouns")
nouns = process_nouns(nouns_rows)
print(f"  ✓ {len(nouns)} существительных")

print("\n=== verbs ===")
verbs_rows = read_sheet(wb, "verbs")
verbs_vocab, regel_verbs, conjugations = process_verbs(verbs_rows)
print(f"  ✓ {len(verbs_vocab)} в VOCAB, {len(regel_verbs)} regel, {len(conjugations)} unregel")

print("\n=== adjectives ===")
adj_rows = read_sheet(wb, "adjectives")
adjectives = process_adjectives(adj_rows)
print(f"  ✓ {len(adjectives)} прилагательных")

print("\n=== adverbs ===")
adv_rows = read_sheet(wb, "adverbs")
adverbs = process_adverbs(adv_rows)
print(f"  ✓ {len(adverbs)} наречий")

print("\n=== pronouns ===")
pron_rows = read_sheet(wb, "pronouns")
pronouns = process_pronouns(pron_rows)
print(f"  ✓ {len(pronouns)} местоимений")

print("\n=== numbers ===")
num_rows = read_sheet(wb, "numbers")
numbers = process_numbers(num_rows)
print(f"  ✓ {len(numbers)} чисел")

print("\n=== phrases ===")
ph_rows = read_sheet(wb, "phrases")
phrases = process_phrases(ph_rows)
print(f"  ✓ {len(phrases)} фраз")

print("\n=== sounds ===")
sounds_rows = read_sheet(wb, "sounds")
sounds = process_sounds(sounds_rows)
print(f"  ✓ {len(sounds)} звуков")

print("\n=== terms ===")
terms_rows = read_sheet(wb, "terms")
terms = process_terms(terms_rows)
print(f"  ✓ {len(terms)} терминов")

print("\n=== rules ===")
rules_rows = read_sheet(wb, "rules")
rules = process_rules(rules_rows)
print(f"  ✓ {len(rules)} правил")

print("\n=== questions ===")
q_rows = read_sheet(wb, "questions")
questions = process_questions(q_rows)
print(f"  ✓ {len(questions)} вопросов")

# ═══════════════════════════════════════════════════════════════
# Считываем константы из старого data.js
# ═══════════════════════════════════════════════════════════════
print("\n=== Читаю старый data.js для констант ===")
old_content = OLD_DATA_JS.read_text(encoding="utf-8")

blocks_block = extract_block(old_content, "BLOCKS")
topic_titles_block = extract_block(old_content, "TOPIC_TITLES")
tab_titles_block = extract_block(old_content, "TAB_TITLES")
phrase_units_block = extract_block(old_content, "PHRASE_UNITS")
sentence_templates_block = extract_block(old_content, "SENTENCE_TEMPLATES")

for name, b in [("BLOCKS", blocks_block), ("TOPIC_TITLES", topic_titles_block),
                ("TAB_TITLES", tab_titles_block), ("PHRASE_UNITS", phrase_units_block),
                ("SENTENCE_TEMPLATES", sentence_templates_block)]:
    if b is None:
        print(f"  ⚠ Не нашёл {name} в старом data.js — поставлю заглушку")

# ═══════════════════════════════════════════════════════════════
# Назначаем id
# ═══════════════════════════════════════════════════════════════
def add_ids(items, prefix, width=4):
    for i, item in enumerate(items, 1):
        item_id = f"{prefix}{i:0{width}d}"
        # Вставить id первым ключом
        new_item = {"id": item_id}
        new_item.update(item)
        items[i-1] = new_item
    return items

# VOCAB = nouns + verbs + adj + adv + pron + num + phrase
all_vocab = []
for items, prefix in [
    (nouns, "n"), (verbs_vocab, "v"), (adjectives, "a"),
    (adverbs, "d"), (pronouns, "p"), (numbers, "u"), (phrases, "f"),
]:
    add_ids(items, prefix)
    all_vocab.extend(items)

add_ids(regel_verbs, "r", width=3)
add_ids(conjugations, "c", width=3)
add_ids(questions, "q")
add_ids(rules, "rl", width=3)
add_ids(terms, "t", width=3)
add_ids(sounds, "s", width=3)

# ═══════════════════════════════════════════════════════════════
# Финальная сериализация
# ═══════════════════════════════════════════════════════════════
print("\n=== Собираю data.js ===")

VOCAB_KEYS = ["id", "de", "altDe", "ru", "pos", "gender", "plural", "level", "topics",
              "note", "new", "comparative", "superlative", "antonym",
              "kind", "case", "digit", "context"]
REGEL_KEYS = ["id", "verb", "ru", "forms", "partizip2", "aux"]
CONJ_KEYS = ["id", "verb", "ru", "tense", "modal", "level", "pronouns", "forms",
             "partizip2", "aux"]
Q_KEYS = ["id", "topic", "level", "difficulty", "type", "q",
          "opts", "ans", "words", "pronouns", "altAns", "hint", "explain"]
R_KEYS = ["id", "title", "topic", "level", "content_md", "examples", "note"]
T_KEYS = ["id", "term", "ru", "topic", "level", "note"]
S_KEYS = ["id", "combo", "pronunciation", "example", "translation", "note"]

out = []
out.append("// ═══════════════════════════════════════════════════════════════")
out.append("// data.js — единый источник правды для шпоры и тренажёра")
out.append("// Собрано автоматически из database.xlsx (см. update_data.py)")
out.append("// ═══════════════════════════════════════════════════════════════")
out.append(f"// VOCAB: {len(all_vocab)} | REGEL: {len(regel_verbs)} | CONJ: {len(conjugations)}")
out.append(f"// QUESTIONS: {len(questions)} | RULES: {len(rules)} | TERMS: {len(terms)} | SOUNDS: {len(sounds)}")
out.append("")

# VOCAB по pos
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

out.append("const CONJUGATIONS = [")
for c in conjugations:
    out.append(item_to_js(c, CONJ_KEYS))
out.append("];")
out.append("")

out.append("const REGEL_VERBS = [")
for r in regel_verbs:
    out.append(item_to_js(r, REGEL_KEYS))
out.append("];")
out.append("")

out.append("const QUESTIONS = [")
for q in questions:
    out.append(item_to_js(q, Q_KEYS))
out.append("];")
out.append("")

out.append("const SOUNDS = [")
for s in sounds:
    out.append(item_to_js(s, S_KEYS))
out.append("];")
out.append("")

out.append("const TERMS = [")
for t in terms:
    out.append(item_to_js(t, T_KEYS))
out.append("];")
out.append("")

out.append("const RULES = [")
for r in rules:
    out.append(item_to_js(r, R_KEYS))
out.append("];")
out.append("")

# BLOCKS / TOPIC_TITLES / TAB_TITLES / PHRASE_UNITS / SENTENCE_TEMPLATES — из старого
if blocks_block:
    out.append("const BLOCKS = " + blocks_block + ";")
    out.append("")
if topic_titles_block:
    out.append("const TOPIC_TITLES = " + topic_titles_block + ";")
    out.append("")
if tab_titles_block:
    out.append("const TAB_TITLES = " + tab_titles_block + ";")
    out.append("")
if phrase_units_block:
    out.append("const PHRASE_UNITS = " + phrase_units_block + ";")
    out.append("")
if sentence_templates_block:
    out.append("const SENTENCE_TEMPLATES = " + sentence_templates_block + ";")
else:
    out.append("const SENTENCE_TEMPLATES = [];")
out.append("")

# Экспорт для node
out.append("if (typeof module !== 'undefined' && module.exports) {")
out.append("  module.exports = { VOCAB, CONJUGATIONS, REGEL_VERBS, QUESTIONS,")
out.append("    SOUNDS, TERMS, RULES, BLOCKS, TOPIC_TITLES, TAB_TITLES,")
out.append("    PHRASE_UNITS, SENTENCE_TEMPLATES };")
out.append("}")
out.append("")

OUT_DATA_JS.write_text("\n".join(out), encoding="utf-8")
print(f"\n✓ Записано: {OUT_DATA_JS}")
print(f"  размер: {OUT_DATA_JS.stat().st_size:,} байт")
