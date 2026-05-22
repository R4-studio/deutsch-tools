# Deutsch Tools

Шпаргалка и тренажёр по немецкому языку (A1 → A2).

🌐 [r4-studio.github.io/deutsch-tools](https://r4-studio.github.io/deutsch-tools)

---

## Файлы в репозитории

| Файл | Назначение |
|---|---|
| `index.html` | Главная страница (выбор: шпаргалка / тренажёр) |
| `cheatsheet.html` | Шпаргалка: словарь, грамматика, правила, поиск |
| `trainer.html` | Тренажёр: карточки, тесты по уровням |
| `data.js` | База данных (генерируется автоматически) |
| `database.xlsx` | **Источник правды**: все слова, правила, тесты |
| `update_data.py` | Конвертер `database.xlsx` → `data.js` |

---

## Архитектура

```
database.xlsx  ──[update_data.py]──>  data.js  ──>  cheatsheet.html
   (источник)         (конвертер)      (готовый)    + trainer.html
```

**Один источник правды — `database.xlsx`.**
Все правки контента — в xlsx, потом запуск конвертера, потом коммит.

---

## Как добавить слово / правило / тест

1. Открыть `database.xlsx` (Excel или LibreOffice)
2. Перейти на нужный лист (nouns, verbs, adjectives, adverbs, pronouns, numbers, phrases, sounds, terms, rules, questions)
3. Добавить строку (id можно не заполнять — сгенерируется автоматически)
4. Сохранить xlsx
5. В терминале: `python update_data.py`
6. `git add . && git commit -m "..." && git push`

---

## Локальный запуск

Двойной клик по `index.html`. React и Babel подгружаются с CDN — никакого билда не нужно.

---

## Этапы

- ✅ Шпаргалка и тренажёр работают на общем `data.js`
- ✅ Тренажёр с 5 уровнями сложности, выбор уровня
- ✅ Конвертер `database.xlsx` → `data.js`
- ✅ Автогенератор Präsens и Partizip II для regelmäßig
- 🔄 Блок «Прошедшее время» (Perfekt)
- ⏳ Триграммный поиск
- ⏳ Прогресс пользователя (localStorage)
