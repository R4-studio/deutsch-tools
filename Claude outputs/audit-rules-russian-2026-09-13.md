# Аудит русских формулировок: 39 правил с ручным переводом

Дата: 13.09.2026 · Проверено: все правила с `source: manual`, то есть те, чей английский делался батчами, а русский не проверялся никогда.
Метод: шесть параллельных проверок, каждое утверждение о немецком сверено с публичными методичками (grammis IDS Mannheim, Duden, DWDS, DeutschAkademie, Deutsche Grammatik 2.0, verbformen, mein-deutschbuch, Dr. Bopp).

| | Количество |
|---|---|
| КРИТИЧНО | 8 |
| СРЕДНЕ | 34 |
| МЕЛКО | 43 |
| Чистых правил | 3 из 39 |

Чистые: `r0001` (личные местоимения), `r0002` (Reflexivpronomen), `r0036` (Zahl + -jährig).

Важное наблюдение: **ни в одном из 39 правил нет ошибочной немецкой формы.** Все напечатанные слова, артикли, склонения и примеры корректны. Ошибки лежат в формулировках — и они порождают неверный немецкий уже на стороне ученика, когда он применяет правило к новому материалу. Поэтому ошибку не видно при беглом просмотре карточки.

---

## 1. Критичные — ученик заучит неправильный немецкий

### r0056 — «dass/ob → возвратное местоимение в Dativ»
**Было:** «was / wer / ob / dass → придаточное само выступает тем, на что направлено действие → возвратное местоимение в Дательном»
**Механизм:** правило порождает `Ich frage **mir**, ob das stimmt`, `Ich freue **mir**, dass du kommst`, `Ich wundere **mir**, dass …` — грубо неверный немецкий в самых частотных B1-оборотах.
**Надо:** Dativ появляется только у глаголов, у которых есть отдельное Akkusativ-дополнение (sich **etwas** wünschen / merken / vorstellen): `Ich wünsche mir, dass du kommst`. Там, где возвратное местоимение само и есть Akkusativ-дополнение (sich freuen, sich fragen, sich erinnern, sich wundern), падеж не меняется никогда.
**Маскировка:** в примерах для sich freuen взят именно `weil`, а не `dass` — то есть единственный пример, который сломал бы правило, обойдён.
Источники: [DWDS fragen](https://www.dwds.de/wb/fragen), [DeutschAkademie Reflexive Verben](https://www.deutschakademie.de/online-deutschkurs/deutsche-grammatik/wortarten/verben/reflexive-verben/)

### r0056 — алгоритм подан как универсальный
**Было:** «### Порядок проверки. 1. Ищем глагол. 2. Ищем, на кого направлено действие → Да → Dativ / Нет → Akkusativ»
**Механизм:** падеж возвратного местоимения — лексическое свойство глагола, а не результат разбора предложения. У sich freuen, sich beeilen, sich erholen, sich verhalten он жёстко Akkusativ при любом окружении.
**Надо:** нулевым шагом — «сначала смотрим, какой глагол; тест применим только там, где у глагола бывает отдельное Akkusativ-дополнение».

### r0063 — «с auf всегда Akkusativ»
**Было:** «⚠ Падеж здесь задаёт предлог: с **auf всегда Akkusativ**, с mit всегда Dativ»
**Механизм:** порождает `Der Film beruht auf **eine wahre Geschichte**` (надо `auf einer wahren Geschichte`), `Er besteht auf **seinen Standpunkt**` (надо `auf seinem`), и буквально применённое — `Das Buch liegt auf **den** Tisch`. Утверждение противоречит базовому материалу A1 про Wechselpräpositionen.
**Надо:** mit всегда Dativ. auf — Wechselpräposition, падеж задаёт конкретный глагол: warten auf + Akk, но beruhen / bestehen / beharren auf + Dat.
Источник: [grammis — Dativ oder Akkusativ nach verbbestimmten Präpositionen](https://grammis.ids-mannheim.de/fragen/3140)

### r0058 — «при разных подлежащих zu-Infinitiv невозможен»
**Было:** «Если подлежащие разные – zu-Infinitiv невозможен, нужен полноценный dass-Satz»
**Механизм:** закрывает для ученика целый класс нормативных конструкций. `Ich bitte dich, die Tür zu schließen` он будет считать ошибкой и строить `Ich bitte dich, dass du die Tür schließt`.
**Надо:** подлежащее zu-инфинитива — это подлежащее главного предложения ИЛИ его дополнение, у глаголов bitten, empfehlen, raten, erlauben, verbieten, vorschlagen, helfen: `Mein Vater erlaubt mir, nach Italien zu fahren`.
Источник: [mein-deutschbuch — Infinitivsätze](https://mein-deutschbuch.de/infinitivsaetze.html)

### r0061 — «в Dativ Plural -n добавляется обязательно»
**Было:** «Dativ Plural, где -n добавляется обязательно: mit den Kindern, mit den Häusern»
**Механизм:** порождает `mit den **Autosn**`, `in den **Hotelsn**`, `mit den **Frauenn**`.
**Надо:** -n добавляется, если форма множественного ещё не оканчивается на -n или -s.
Источник: [verbformen — Auto](https://www.verbformen.com/declension/nouns/Auto.htm)

### r0024 — «Nominativ всегда только один в предложении»
**Было:** «💡 Nominativ – всегда только один в предложении, это «деятель»» + «Nominativ – всегда субъект»
**Механизм:** запрещает Gleichsetzungsnominativ. Ученик будет искать для второго существительного другой падеж: `Er ist **einen** Lehrer`, `Das ist **einen** Apfel`, `Ich werde **einen** Arzt`.
**Надо:** Nominativ — падеж подлежащего И именной части сказуемого после sein / werden / bleiben, поэтому в одном предложении может стоять дважды.
Источник: [Duden — Gleichsetzungsnominativ](https://www.duden.de/rechtschreibung/Gleichsetzungsnominativ)

### r0004 — заголовок про Perfekt, содержимое про копулу
**Было:** заголовок «Perfekt: sein или haben?», в содержимом таблица «sein — прилагательные, профессии / haben — существительные», ни одного Partizip II
**Механизм:** ученик, открывший правило по названию, применит «haben — с существительными» к перфекту и скажет `Ich **habe** nach Berlin gefahren` вместо `bin gefahren`. Правило по названию должно предотвращать ровно эту ошибку, а вместо этого её провоцирует.
**Надо:** либо переименовать в «sein и haben как самостоятельные глаголы», либо наполнить правилом выбора вспомогательного глагола.
Источник: [grammis — haben oder sein?](https://grammis.ids-mannheim.de/progr@mm/5316)

### r0067 — требование Konjunktiv, опровергаемое собственными примерами
**Было:** «### Следствие обязано быть таким же гипотетическим… Иначе форма ломается»
**Механизм:** три из пяти примеров этого же правила нарушают требование — `Sollten Sie Fragen haben, **melden Sie sich** gern` и `Sollte es regnen, **bleiben wir** zu Hause`. Ученик либо сочтёт примеры ошибочными, либо начнёт «чинить» стандартный оборот.
**Надо:** после Sollte… главная часть может стоять в Konjunktiv II, в индикативе или в императиве — всё нормативно. Konjunktiv выбирают для смягчения, а не по обязанности.
Источник: [grammis — Konditionalsätze](https://grammis.ids-mannheim.de/systematische-grammatik/2101)

---

## 2. Замечания к тому, что правили сегодня

### r0057 — блок про порядок слов неполон (моя правка)
**Было** (то, что я сформулировал сегодня): «sowohl...als auch и weder...noch ведут себя как сочинительные — порядок слов после них не меняется»
**Механизм:** верно, пока союз соединяет члены одного предложения. Но когда **weder…noch** соединяет два самостоятельных предложения, `noch` занимает первое место и глагол идёт сразу за ним: `Weder **hat er** angerufen, noch **hat er** geschrieben`. В отличие от und/aber, noch — не нулевая позиция. По моей формулировке ученик напишет `noch **er hat** geschrieben`.
**Надо:** добавить эту оговорку в блок «Порядок слов».
Источник: [Duden — Kommasetzung bei „weder – noch"](https://www.duden.de/sprachwissen/sprachratgeber/Kommasetzung-bei-%E2%80%9Eweder-%E2%80%93-noch%E2%80%9C)

Блок про entweder, который я добавил, проверку прошёл — оба варианта порядка слов подтверждены.

### r0032 — note провоцирует «mögte»
**Было:** «Умлаут исчезает: müssen→musste, können→konnte, dürfen→durfte. **Genauso**: sollen→sollte, mögen→mochte»
**Механизм:** «genauso» после тезиса про умлаут подталкивает к форме `mögte`. У mögen меняется не только умлаут, но и согласный: g → ch. Плюс в таблице стоит wollen, к которому объяснение про умлаут вообще неприменимо.
**Надо:** развести — у müssen/können/dürfen/mögen умлаут исчезает; у wollen и sollen его и не было; у mögen дополнительно g → ch.
Плюс пробел: у **möchte** своей формы Präteritum нет, в прошедшем используется wollte — без этой оговорки ученик произведёт «möchtete».

### r0059 — «sondern только после отрицания» сформулировано неточно
**Механизм:** по формулировке «sondern обязательно после nicht/kein» ученик после любого отрицания ставит sondern: `Er ist nicht reich, **sondern** glücklich` — бессмыслица, надо `aber`. Различие не в наличии отрицания, а в функции: sondern исправляет отрицаемое, aber добавляет контраст при сохранении отрицания.
Источник: [Duden — Besonderheiten von „sondern"](https://www.duden.de/sprachwissen/sprachratgeber/Besonderheiten-von-sondern)

---

## 3. Средние — правило подведёт в части случаев

Сгруппировано по типу.

**Ложные абсолюты «всегда / только / никогда»**
- `r0031` — «Wo? → всегда Dativ / Wohin? → всегда Akkusativ». Работает лишь для девяти Wechselpräpositionen. Порождает `Ich gehe **zu den** Arzt`.
- `r0026`, `r0027` — «если оба существительных, Dativ всегда первый» и 💡-подсказка «Dativ → Akkusativ» без оговорки про местоимения. Порождает `Ich gebe dem Mann **es**`.
- `r0026` — «оба местоимения → Akkusativ первый» верно только для личных: `Ich gebe **ihm eins**`, не «eins ihm».
- `r0023`, `r0029` — «глагол всегда последний в Nebensatz». Исключение — Ersatzinfinitiv: `…, weil er **hat** arbeiten müssen`.
- `r0033` — «в Pl Dativ существительное тоже +n», та же проблема, что в r0061.
- `r0050` — «erst = только/лишь». У erst есть второе значение «сначала», в паре с dann: `**Erst** die Arbeit, **dann** das Vergnügen`.
- `r0060` — «trotzdem, deshalb занимают позицию 1». Могут стоять и внутри предложения: `Ich gehe **trotzdem** spazieren`.

**Правило противоречит собственному примеру**
- `r0045` — проза говорит «инфинитив уходит в конец», пример даёт `damit sie besser Deutsch **sprechen konnten**`. В придаточном порядок обратный.
- `r0047` — заявлено окончание Genitiv `-s`, а собственный пример даёт `des **Mannes**`. Плюс не учтена n-Deklination: `des Studenten`, не «des Students».
- `r0065` — подсказка «Partizip II после werden = Passiv» опровергается разделом Futur II того же правила: `Ich werde das Buch **gelesen haben**`.
- `r0042` — «с am окончание всегда -sten», а строкой ниже `am **größten**`.
- `r0024` — Akkusativ и Dativ описаны как «субъект + объект», хотя подлежащее в немецком всегда Nominativ.

**Пробел, ведущий к ошибке**
- `r0044` — «seit … Perfekt/Präsens» разрешает `Ich **habe** seit zwei Jahren Deutsch **gelernt**`. При seit нужен Präsens, хотя по-русски прошедшее.
- `r0062` — «bis: оба действия идут одновременно и заканчиваются в одной точке» противоречит своему примеру `Warte, bis ich komme`. bis называет конечную точку.
- `r0066` — нет Zustandspassiv, из текста следует, что `Das Auto ist repariert` вообще не немецкий. Плюс `von dem Mechaniker` вместо нормативного `vom Mechaniker`.
- `r0066` — нет ограничения «только переходные с Akkusativ»: порождает `Der Mann wird geholfen` вместо `Dem Mann wird geholfen`.
- `r0041` — «прилагательное + -er» без склонения перед существительным: порождает `ein **schneller** Auto` вместо `ein schnelleres Auto`.
- `r0042` — вставка -e- перед -sten без исключений: порождает `am **spannendesten**`, `am **praktischesten**`.
- `r0039` — запрет «sein in» подан слишком широко: ученик не построит `der Mann **in** der blauen Jacke` и `**in** Uniform sein`.
- `r0003` — шаблон «kein + Nomen» без склонения: порождает `Man darf **kein** Hund haben`.
- `r0055` — «'s всегда = es» неверно: в ins/ans/aufs это das.
- `r0058` — безличные обороты (`Es ist wichtig, … zu`) стоят в списке «то же подлежащее», хотя проверка к ним неприменима.

---

## 4. Что делать

Порядок, в котором это имеет смысл чинить:

1. **Восемь критичных** — они порождают неверный немецкий прямо сейчас, в самых частотных конструкциях. Плюс три замечания к сегодняшним правкам.
2. **Ложные абсолюты** — их чинить дёшево: почти везде достаточно дописать оговорку, переписывать правило не надо.
3. **Противоречия примеру** — здесь придётся выбирать, что верно: текст или пример. В пяти случаях из пяти верен пример, а текст неточен.
4. **Пробелы** — по одному предложению в каждое правило.

После правки русского английский для затронутых правил переводится заново — старый батч на изменившийся русский не встанет.
