// i18n/i18n.js — общий рантайм переключения языка UI.
// Подключается ПОСЛЕ i18n/ru.js и i18n/en.js (нужны window.I18N_RU/I18N_EN).
// Слой 1 из брифа "Мультиязычный интерфейс": только статичный UI-текст,
// сами немецкие слова в data.js этот файл не трогает.
//
// Разметка:
//   data-i18n="key"             — innerHTML элемента = t(key)
//   data-i18n-title="key"       — атрибут title = t(key)
//   data-i18n-aria="key"        — атрибут aria-label = t(key)
//   data-i18n-placeholder="key" — атрибут placeholder = t(key)
//
// Страница вызывает applyI18n() сама (обычно в конце своего <script>),
// по желанию определив window.onI18nApplied(lang) до вызова — туда кладут
// то, что рантайм не знает как обновить сам (например title кнопки темы,
// который зависит и от языка, и от текущей темы одновременно).
var I18N_DICTS = { ru: window.I18N_RU || {}, en: window.I18N_EN || {} };

function getLang() {
  return localStorage.getItem('lang') || 'ru';
}

function t(key) {
  var dict = I18N_DICTS[getLang()] || I18N_DICTS.ru;
  if (dict && dict[key] != null) return dict[key];
  return (I18N_DICTS.ru[key] != null) ? I18N_DICTS.ru[key] : key;
}

function applyI18n() {
  var lang = getLang();
  document.documentElement.lang = lang;

  document.querySelectorAll('[data-i18n]').forEach(function(el) {
    el.innerHTML = t(el.getAttribute('data-i18n'));
  });
  document.querySelectorAll('[data-i18n-title]').forEach(function(el) {
    el.title = t(el.getAttribute('data-i18n-title'));
  });
  document.querySelectorAll('[data-i18n-aria]').forEach(function(el) {
    el.setAttribute('aria-label', t(el.getAttribute('data-i18n-aria')));
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach(function(el) {
    el.placeholder = t(el.getAttribute('data-i18n-placeholder'));
  });

  var langBtn = document.getElementById('lang-btn');
  if (langBtn) {
    langBtn.textContent = lang.toUpperCase();
    langBtn.title = t('nav.lang_title');
  }

  if (typeof window.onI18nApplied === 'function') window.onI18nApplied(lang);
}

function cycleLang() {
  localStorage.setItem('lang', getLang() === 'ru' ? 'en' : 'ru');
  applyI18n();
}
