// ═══════════════════════════════════════════════════════════════
// РЕНДЕР КАРТОЧКИ ПРАВИЛА — общий для cheatsheet.html и trainer.html
// ═══════════════════════════════════════════════════════════════
// Вынесено из cheatsheet.html, чтобы не дублировать код карточки во
// втором месте показа (оверлей в тренажёре). Подключать после
// i18n/i18n.js и data.js — файл зависит от TOPIC_TITLES и RULES.

// ─── Глобальный хелпер экранирования HTML ────────────────────────────────
function escHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
// ic(key) — достаёт SVG-разметку иконки из ICON_SVGS (единый источник, генерируется в data.js)
function ic(key) { return (typeof ICON_SVGS !== 'undefined' && ICON_SVGS[key]) || ''; }

// Маркеры-иконки в начале строк content_md (правила): 💡 совет, ⚠ предупреждение и т.п.
var NOTE_ICON_MAP = { '💡': 'bulb', '⚠': 'warning', '📍': 'pin', '🔑': 'key', '✓': 'check', '✗': 'cross', '❗': 'warning', '👉': 'arrow-right' };

// ─── Утилиты для рендера правил ────────────────────────────────────────

// Превращает inline-разметку **жирное** в HTML
function renderInline(s) {
  return escHtml(s)
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/__([^_]+)__/g, '<b>$1</b>')
    // Разрешаем inline-окраску родов: <span class="g-m">текст</span>
    // (после escHtml теги превратились в &lt;span...&gt; — возвращаем обратно)
    .replace(/&lt;span class=&quot;(g-(?:m|f|n|pl))&quot;&gt;([\s\S]*?)&lt;\/span&gt;/g,
             '<span class="$1">$2</span>');
}

// ─── Карточный рендер правила — renderRuleCard() ───────────────────────
// Разметка внутри content_md (маркеры — ASCII, EN-пайплайн заворачивает их
// в <x>…</x>, DeepL не трогает; число строк и `|` сохраняется):
//   [sub] текст              — 1-я непустая строка → подзаголовок в шапке
//   1. … / 2. …              — подряд ≥2 строк → нумерованный чек-лист (кружки)
//   :::outcomes … :::        — парные карточки; строка «метка → акцент | подпись»
//   ### Заголовок [collapse] — секцию (до следующего ###) прячем в <details>
//   ### Заголовок            — обычная открытая подсекция
//   💡 ⚠ 📍 …                — callout мелким шрифтом
//   | a | b |                — таблица
//   прочее                   — <p> (ни одно предложение не теряется)
// Плашка-примечание снизу — из r.note/noteEn. Зелёная ссылка «см. также» —
// id вида r0053 в r.note, если резолвится в другое правило.
const RULES_BY_ID = (typeof RULES !== 'undefined')
  ? Object.fromEntries(RULES.map(r => [r.id, r])) : {};

// ─── Готовность теста по правилу ─────────────────────────────────────────
// Один источник для обеих страниц: справочник рисует по этому живую или мёртвую
// кнопку «к тесту», тренажёр — обычную или приглушённую карточку правила.

// Сколько вопросов нужно, чтобы тест считался готовым. Тест из одного вопроса —
// не тест: он открывается и сразу показывает «1/1, 100%».
const MIN_Q_FOR_TEST = 3;

// Вопросы одного правила. Тип open в тренажёре пока не рендерится — исключаем,
// иначе вопрос покажет пустой экран.
// TODO: снять фильтр, когда появится рендер типа open.
function questionsForRule(ruleId) {
  const qs = (typeof QUESTIONS !== 'undefined') ? QUESTIONS : [];
  return qs.filter(q => q.rule === ruleId && q.type !== 'open');
}

// Счётчик по всем правилам сразу — считается один раз при загрузке.
const RULE_Q_COUNT = (() => {
  const m = {};
  const qs = (typeof QUESTIONS !== 'undefined') ? QUESTIONS : [];
  for (const q of qs) {
    if (!q.rule || q.type === 'open') continue;
    m[q.rule] = (m[q.rule] || 0) + 1;
  }
  return m;
})();

function ruleHasTest(ruleId) { return (RULE_Q_COUNT[ruleId] || 0) >= MIN_Q_FOR_TEST; }

window.MIN_Q_FOR_TEST = MIN_Q_FOR_TEST;
window.RULE_Q_COUNT   = RULE_Q_COUNT;
window.questionsForRule = questionsForRule;
window.ruleHasTest    = ruleHasTest;

const RULE_CHEVRON = '<svg class="rule-collapse__chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>';

function ruleAccent(text) {
  const s = (text || '').toLowerCase();
  if (/\bdativ\b/.test(s)) return 'cyan';
  if (/\bakkusativ\b/.test(s)) return 'red';
  return null;
}

// Вертикальная окраска колонок таблицы. Падежи держат закреплённый цвет
// (Akkusativ — красный, Dativ — циан, Genitiv — янтарь, Nominativ — серый),
// остальные колонки — по порядку. Первая колонка обычно строковая «шапка
// строки» — оставляем нейтральной (серой).
const RULE_CASE_COL = [
  [/\bakkusativ\b|\bwen\?/i,  'rc-red'],
  [/\bdativ\b|\bwem\?/i,      'rc-cyan'],
  [/\bgenitiv\b|\bwessen\?/i, 'rc-amber'],
  [/\bnominativ\b|\bwer\?/i,  'rc-nom'],
];
const RULE_POS_COL = ['rc-cyan', 'rc-red', 'rc-amber', 'rc-green'];
function ruleTable(rows) {
  const colCls = (rows[0] || []).map((hCell, ci) => {
    const plain = String(hCell).replace(/\*\*/g, '');
    for (const [re, cls] of RULE_CASE_COL) if (re.test(plain)) return cls;
    if (ci === 0) return 'rc-nom';
    return RULE_POS_COL[(ci - 1) % RULE_POS_COL.length];
  });
  let h = '<table>';
  rows.forEach((cells, idx) => {
    const tag = idx === 0 ? 'th' : 'td';
    h += '<tr>';
    cells.forEach((c, ci) => {
      const cls = colCls[ci] ? ' class="' + colCls[ci] + '"' : '';
      h += '<' + tag + cls + '>' + renderInline(c) + '</' + tag + '>';
    });
    h += '</tr>';
  });
  return h + '</table>';
}

function ruleChecklist(steps) {
  let h = '<ol class="rule-steps">';
  steps.forEach((s, i) => {
    h += '<li><span class="rule-steps__n">' + (i + 1) + '</span><span>' + renderInline(s) + '</span></li>';
  });
  return h + '</ol>';
}

function ruleOutcomes(rows) {
  let h = '<div class="rule-outcomes">';
  rows.forEach((raw, i) => {
    let cap = '';
    const p = raw.indexOf('|');
    if (p !== -1) { cap = raw.slice(p + 1).trim(); raw = raw.slice(0, p).trim(); }
    const accent = ruleAccent(raw) || ['cyan', 'red', 'amber'][i];
    h += '<div class="rule-outcome' + (accent ? ' rule-outcome--' + accent : '') + '">'
       + renderInline(raw)
       + (cap ? '<span class="rule-outcome__cap">' + renderInline(cap) + '</span>' : '')
       + '</div>';
  });
  return h + '</div>';
}

// Буфер подряд идущих строк → чек-лист (1-я «N. …», строки без номера
// приклеиваются к предыдущему шагу). null — это не чек-лист.
function tryChecklist(buf) {
  if (!/^\d+\.\s+/.test(buf[0])) return null;
  const steps = [];
  for (const l of buf) {
    const m = l.match(/^(\d+)\.\s+(.*)$/);
    if (m) steps.push(m[2]);
    else if (steps.length) steps[steps.length - 1] += ' ' + l.trim();
    else return null;
  }
  return steps.length >= 2 ? steps : null;
}

function renderRuleBlocks(md) {
  const lines = (md || '').split(/\r?\n/);
  let out = '', buf = [], tableRows = [];

  const flushTable = () => { if (tableRows.length) { out += ruleTable(tableRows); tableRows = []; } };
  const flushBuf = () => {
    if (!buf.length) return;
    const steps = tryChecklist(buf);
    if (steps) { out += ruleChecklist(steps); buf = []; return; }
    for (const line of buf) {
      // флаг u — иначе emoji-класс матчит только суррогатную половину 💡/📍/🔑
      const m = line.match(/^([💡⚠📍🔑✓✗❗👉])\s*/u);
      if (m) {
        const ch = m[1];
        const icon = NOTE_ICON_MAP[ch];
        const mod = ch === '💡' ? ' rule-callout--bulb'
                  : (ch === '⚠' || ch === '❗' || ch === '✗') ? ' rule-callout--warn' : '';
        out += '<p class="rule-callout' + mod + '">' + (icon ? ic(icon) : '') + '<span>' + renderInline(line.slice(m[0].length)) + '</span></p>';
      } else out += '<p>' + renderInline(line) + '</p>';
    }
    buf = [];
  };
  const flushAll = () => { flushTable(); flushBuf(); };

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();

    // фенс :::outcomes … :::
    if (/^:::\s*outcomes\b/i.test(trimmed)) {
      flushAll();
      const rows = [];
      i++;
      for (; i < lines.length && !/^:::\s*$/.test(lines[i].trim()); i++) {
        if (lines[i].trim()) rows.push(lines[i].trim());
      }
      if (rows.length) out += ruleOutcomes(rows);
      continue;
    }

    // заголовок секции ### / ##
    const hm = trimmed.match(/^#{2,3}\s+(.*)$/);
    if (hm) {
      flushAll();
      let head = hm[1];
      const collapse = /\[collapse\]\s*$/i.test(head);
      if (collapse) head = head.replace(/\s*\[collapse\]\s*$/i, '');
      const body = [];
      let j = i + 1;
      for (; j < lines.length; j++) {
        const lt = lines[j].trim();
        if (/^#{2,3}\s+/.test(lt) || /^:::\s*outcomes\b/i.test(lt)) break;
        body.push(lines[j]);
      }
      const bodyHtml = renderRuleBlocks(body.join('\n'));
      if (collapse) {
        out += '<details class="rule-collapse"><summary><span>' + renderInline(head) + '</span>' + RULE_CHEVRON + '</summary>'
             + '<div class="rule-collapse__body">' + bodyHtml + '</div></details>';
      } else {
        out += '<div class="rule-sec"><div class="rule-sec__head">' + renderInline(head) + '</div>' + bodyHtml + '</div>';
      }
      i = j - 1;
      continue;
    }

    if (!trimmed) { flushAll(); continue; }

    // таблица: строка с 2+ '|'
    if ((trimmed.match(/\|/g) || []).length >= 2) {
      flushBuf();
      const cells = trimmed.split('|').map(c => c.trim()).filter(c => c !== '');
      if (cells.length) { tableRows.push(cells); continue; }
    }

    flushTable();
    buf.push(trimmed);
  }
  flushAll();
  return out;
}

function renderRulePlate(text) {
  if (!text) return '';
  let cls = 'rule-plate', icon = '';
  // note может начинаться с маркера 💡/⚠ — вырезаем его, ставим иконку и цвет
  const m = text.match(/^([💡⚠📍🔑✓✗❗👉])\s*/u);
  if (m) {
    text = text.slice(m[0].length);
    const key = NOTE_ICON_MAP[m[1]];
    if (key) icon = ic(key);
    if (m[1] === '⚠' || m[1] === '❗' || m[1] === '✗') cls += ' rule-plate--red';
  }
  return '<div class="' + cls + '">' + icon + renderInline(text) + '</div>';
}

// Примеры — отдельная колонка examples/examplesEn. Прогоняются через тот же
// renderRuleBlocks, что и тело: там встречаются таблицы (r0019) и подсказки.
function renderRuleExamples(rule) {
  const md = wordField(rule, 'examples', 'examplesEn');
  if (!md || !md.trim()) return '';
  return '<div class="rule-examples">'
       + '<div class="rule-examples__head">' + ic('notes') + '<span>' + escHtml(window.t('cheatsheet.rule_examples')) + '</span></div>'
       + renderRuleBlocks(md)
       + '</div>';
}

// Ссылки «см. также» кликабельны только когда страница задала обработчик:
// справочник подставляет свой gotoRule через window.onRuleLinkClick, тренажёр —
// свой (открытие правила в оверлее). Обработчика нет — ссылки не рисуются.
function renderRuleSeeAlso(rule) {
  if (typeof window.onRuleLinkClick !== 'function') return '';
  const ids = ((rule.note || '').match(/\br\d{4}\b/g) || []);   // id языконезависимы — из RU-note
  const seen = {};
  const links = [];
  for (const id of ids) {
    if (id === rule.id || seen[id] || !RULES_BY_ID[id]) continue;
    seen[id] = true;
    const title = wordField(RULES_BY_ID[id], 'title', 'titleEn');
    links.push('<button type="button" class="rule-seealso" onclick="window.onRuleLinkClick(\'' + id + '\')">'
           + ic('link') + '<span>' + escHtml(id) + ' — ' + escHtml(title) + '</span></button>');
  }
  if (!links.length) return '';
  return '<div class="rule-card__links">' + links.join('') + '</div>';
}

// Кнопка к тесту. Живой она становится, только если (а) у правила набран тест и
// (б) страница объявила window.onRuleTestClick. Справочник объявляет — и кнопка
// уводит в тренажёр; в оверлее тренажёра обработчик будет свой.
function renderRuleTestBtn(rule) {
  const id = rule && rule.id;
  const arrow = '<svg class="rule-testbtn__arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>';
  const ready = id && ruleHasTest(id) && typeof window.onRuleTestClick === 'function';

  if (!ready) {
    // как было: неактивная, «к тесту» → «скоро» на наведении и по тапу
    return '<button type="button" class="rule-testbtn" aria-disabled="true">'
         + '<span class="rule-testbtn__roll"><span class="rule-testbtn__roll-inner">'
         + '<span>' + escHtml(window.t('cheatsheet.rule_take_test')) + '</span>'
         + '<span>' + escHtml(window.t('cheatsheet.rule_soon')) + '</span></span></span>'
         + arrow + '</button>';
  }

  // Живая: подпись с числом вопросов, ролл крутится, но текст не меняется
  // (та же механика, что у .dt-cta на главной тренажёра).
  const label = escHtml(window.t('cheatsheet.rule_take_test')) + ' · ' + (RULE_Q_COUNT[id] || 0);
  return '<button type="button" class="rule-testbtn rule-testbtn--ready" '
       + 'onclick="window.onRuleTestClick(\'' + escHtml(id) + '\')">'
       + '<span class="rule-testbtn__roll"><span class="rule-testbtn__roll-inner">'
       + '<span>' + label + '</span><span>' + label + '</span></span></span>'
       + arrow + '</button>';
}

// Полная карточка правила. opts.compact — компактный отступ сверху (вкладка «Местоимения»).
function renderRuleCard(rule, opts) {
  opts = opts || {};
  const lines = (wordField(rule, 'content_md', 'content_md_en') || '').split(/\r?\n/);
  const eyebrow = [rule.domen, rule.group].filter(Boolean).join(' · ');
  const title = wordField(rule, 'title', 'titleEn');
  const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();

  let subtitle = '';
  for (let k = 0; k < lines.length; k++) {
    if (!lines[k].trim()) continue;
    const t = lines[k].trim();
    const sm = t.match(/^\[sub\]\s*(.*)$/i);
    if (sm) { subtitle = sm[1]; lines.splice(k, 1); break; }
    // первый ### дублирует заголовок карточки — убираем, тело секции остаётся
    const hm = t.match(/^#{2,3}\s+(.*?)(?:\s*\[collapse\])?$/i);
    if (hm && norm(hm[1]) === norm(title)) lines.splice(k, 1);
    break;
  }

  let h = '<div class="rule-card" data-rule-id="' + escHtml(rule.id || '') + '"'
        + (opts.compact ? ' style="margin-top:12px"' : '') + '>';
  // Иконка домена: пара domen:group уже описана в TOPIC_TITLES ({icon,label}),
  // отдельной привязки заводить не надо. Темы нет в справочнике — иконки нет,
  // шапка тогда собирается по-старому, без сетки.
  const topicKey = rule.topic || ((rule.domen && rule.group) ? rule.domen + ':' + rule.group : '');
  const topicMeta = (typeof TOPIC_TITLES !== 'undefined' && topicKey) ? TOPIC_TITLES[topicKey] : null;
  const iconSvg = (topicMeta && topicMeta.icon) ? ic(topicMeta.icon) : '';

  const headInner = (eyebrow ? '<div class="rule-card__eyebrow">' + escHtml(eyebrow) + '</div>' : '')
    + '<div class="rule-card__title' + ((subtitle || iconSvg) ? '' : ' is-last') + '">' + escHtml(title) + '</div>'
    + (subtitle ? '<div class="rule-card__subtitle">' + renderInline(subtitle) + '</div>' : '');
  h += iconSvg
    ? '<div class="rule-card__head"><span class="rule-icon">' + iconSvg + '</span>' + headInner + '</div>'
    : headInner;

  h += renderRuleBlocks(lines.join('\n'));
  h += renderRuleExamples(rule);
  h += renderRulePlate(wordField(rule, 'note', 'noteEn'));

  const seeAlso = renderRuleSeeAlso(rule);
  const testBtn = opts.noTestBtn ? '' : renderRuleTestBtn(rule);
  if (seeAlso || testBtn) {
    h += '<div class="rule-card__footer' + (seeAlso ? '' : ' is-endonly') + '">' + seeAlso + testBtn + '</div>';
  }

  return h + '</div>';
}
