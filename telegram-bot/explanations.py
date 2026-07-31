"""Доп. текст, который подмешивается к popup-у с ответом (верно/неверно уже
показывает вызывающий код). Через раз (см. LINK_PROBABILITY) добавляем ссылку
на сайт — неформально, с самоиронией, а не "зайдите поучите матчасть"."""
import random

CTA_VARIANTS = [
    "Кстати, весь словарь и тренажёр лежат тут: {url} — если что, я не обижусь",
    "Есть подозрительно полезный сайт на такое: {url}",
    "На сайте можно потренироваться специально на такие слова: {url}",
    "{url} — там я всё это и собирал, вдруг пригодится",
    "Если промахнулся не в первый раз — сайт в помощь: {url}",
]


def build_explanation(word: dict, site_url: str, link_probability: float) -> str:
    base = word.get("exampleDe") or word.get("note") or ""
    parts = []
    if base:
        parts.append(base)
    if random.random() < link_probability:
        parts.append(random.choice(CTA_VARIANTS).format(url=site_url))
    text = " ".join(p for p in parts if p).strip()
    # popup через show_alert ограничен ~200 символами
    return text[:200] if text else None
