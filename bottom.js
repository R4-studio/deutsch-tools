// bottom.js — нижний блок страницы: ряд кнопок и футер.
//
// Раньше всё это жило внутри index.html — разметка, стили и логика кнопки
// «Установка». Чтобы блок был на всех страницах, а правился в одном месте,
// он вынесен сюда. Странице достаточно поставить <div id="dt-bottom"></div>
// и подключить этот файл; стили лежат в theme.css.
//
// Файл приносит с собой всё, что блоку нужно: тост, showStub и установку.
// Строка Update рисуется только там, где страница уже загрузила changelog —
// тащить его 52 КБ ради одной даты на каждую страницу смысла нет.

(function () {
  var HTML = "  <div class=\"dt-info-row\">\n    <button class=\"dt-info-item\" onclick=\"showStub()\">\n      <span class=\"dt-info-item__icon\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><circle cx=\"12\" cy=\"8\" r=\"4\"/><path d=\"M4 21c0-4 3.5-7 8-7s8 3 8 7\"/></svg></span>\n      <span class=\"dt-info-item__text\">\n        <div class=\"dt-info-item__title\" data-i18n=\"info.profile.title\">Профиль</div>\n        <div class=\"dt-info-item__sub\" data-i18n=\"info.profile.sub\">Скоро</div>\n      </span>\n    </button>\n    <button class=\"dt-info-item\" onclick=\"installApp()\">\n      <span class=\"dt-info-item__icon\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M12 3v12\"/><path d=\"m7 10 5 5 5-5\"/><path d=\"M4 21h16\"/></svg></span>\n      <span class=\"dt-info-item__text\">\n        <div class=\"dt-info-item__title\" data-i18n=\"info.install.title\">Установка</div>\n        <div class=\"dt-info-item__sub\" id=\"install-sub\" data-i18n=\"info.install.sub\">Скоро</div>\n      </span>\n    </button>\n    <a class=\"dt-info-item\" href=\"faq.html\">\n      <span class=\"dt-info-item__icon\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><circle cx=\"12\" cy=\"12\" r=\"10\"/><path d=\"M9.5 9a2.5 2.5 0 0 1 4.9.8c0 1.7-2.4 2-2.4 3.7\"/><path d=\"M12 17h.01\"/></svg></span>\n      <span class=\"dt-info-item__text\">\n        <div class=\"dt-info-item__title\" data-i18n=\"info.faq.title\">FAQ</div>\n        <div class=\"dt-info-item__sub\" data-i18n=\"info.faq.sub\">Вопросы и ответы</div>\n      </span>\n    </a>\n    <a class=\"dt-info-item\" href=\"about.html\">\n      <span class=\"dt-info-item__icon\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><circle cx=\"12\" cy=\"12\" r=\"10\"/><path d=\"M12 16v-4\"/><path d=\"M12 8h.01\"/></svg></span>\n      <span class=\"dt-info-item__text\">\n        <div class=\"dt-info-item__title\" data-i18n=\"info.about.title\">О проекте</div>\n        <div class=\"dt-info-item__sub\" data-i18n=\"info.about.sub\">Подробнее</div>\n      </span>\n    </a>\n    <a class=\"dt-info-item\" href=\"faq.html#contact\">\n      <span class=\"dt-info-item__icon\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M21 11.5a8.38 8.38 0 0 1-8.5 8.4 8.5 8.5 0 0 1-4-1L3 20l1.1-5.5A8.4 8.4 0 0 1 12.5 3a8.5 8.5 0 0 1 8.5 8.5Z\"/></svg></span>\n      <span class=\"dt-info-item__text\">\n        <div class=\"dt-info-item__title\" data-i18n=\"info.feedback.title\">Обратная связь</div>\n        <div class=\"dt-info-item__sub\" data-i18n=\"info.feedback.sub\">Напишите нам</div>\n      </span>\n    </a>\n  </div>\n\n  <div class=\"dt-footer\">\n    <div class=\"dt-footer__social-group\">\n      <div class=\"dt-footer__social-label\" data-i18n=\"footer.social_label\">Связь и сообщества</div>\n      <div class=\"dt-footer__social\">\n        <a class=\"dt-footer__social-btn\" href=\"https://www.linkedin.com/in/r4-designer\" target=\"_blank\" rel=\"noopener\" title=\"LinkedIn\">in</a>\n        <a class=\"dt-footer__social-btn\" href=\"https://www.behance.net/r4-designer\" target=\"_blank\" rel=\"noopener\" title=\"Behance\">Bē</a>\n        <a class=\"dt-footer__social-btn\" href=\"https://www.instagram.com/r4.designer\" target=\"_blank\" rel=\"noopener\" title=\"Instagram\">\n          <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><rect x=\"3\" y=\"3\" width=\"18\" height=\"18\" rx=\"5\"/><circle cx=\"12\" cy=\"12\" r=\"4\"/><circle cx=\"17.5\" cy=\"6.5\" r=\"0.6\" fill=\"currentColor\" stroke=\"none\"/></svg>\n        </a>\n      </div>\n    </div>\n    <div class=\"dt-footer__meta\">\n      <span id=\"footer-update\">Update: – · открытый проект</span>\n      <a href=\"https://github.com/r4-studio/deutsch-tools\" target=\"_blank\">\n        <svg viewBox=\"0 0 24 24\" fill=\"currentColor\"><path d=\"M12 2a10 10 0 0 0-3.16 19.5c.5.1.68-.22.68-.48v-1.7c-2.78.6-3.37-1.34-3.37-1.34-.46-1.16-1.11-1.47-1.11-1.47-.9-.62.07-.6.07-.6 1 .07 1.53 1.03 1.53 1.03.9 1.52 2.34 1.08 2.91.83.09-.65.35-1.08.63-1.33-2.22-.25-4.56-1.11-4.56-4.94 0-1.09.39-1.98 1.03-2.68-.1-.25-.45-1.27.1-2.65 0 0 .84-.27 2.75 1.02a9.6 9.6 0 0 1 5 0c1.91-1.29 2.75-1.02 2.75-1.02.55 1.38.2 2.4.1 2.65.64.7 1.03 1.59 1.03 2.68 0 3.84-2.34 4.68-4.57 4.93.36.31.68.92.68 1.85v2.74c0 .27.18.58.69.48A10 10 0 0 0 12 2Z\"/></svg>\n        GitHub\n      </a>\n    </div>\n  </div>";

  function mount() {
    var slot = document.getElementById('dt-bottom');
    if (!slot) return;
    slot.innerHTML = HTML;

    if (!document.getElementById('dt-toast')) {
      var toast = document.createElement('div');
      toast.className = 'dt-toast';
      toast.id = 'dt-toast';
      document.body.appendChild(toast);
    }

    renderUpdate();
    if (dtIsStandalone()) dtInstallSub('info.install.done');
    // Юридические страницы объявлены lang="de" — их текст по-немецки и
    // не переводится. applyI18n() ставит lang по выбранному языку интерфейса,
    // поэтому на таких страницах возвращаем атрибут обратно: перевести надо
    // только подписи кнопок, а язык документа менять нельзя.
    if (typeof applyI18n === 'function') {
      var pageLang = document.documentElement.getAttribute('lang');
      var fixed = document.documentElement.hasAttribute('data-lang-fixed')
               || pageLang === 'de';
      applyI18n();
      if (fixed) document.documentElement.setAttribute('lang', pageLang);
    }
  }

  // ─── строка Update ────────────────────────────────────────────────
  function renderUpdate() {
    var el = document.getElementById('footer-update');
    if (!el) return;
    var open = (typeof t === 'function') ? t('footer.open_project') : 'открытый проект';
    var list = window.CHANGELOG;
    if (Array.isArray(list) && list.length && typeof window.formatDate === 'function') {
      var label = (typeof t === 'function') ? t('footer.update_label') : 'Update:';
      el.textContent = label + ' ' + window.formatDate(list[0].date)
                     + ' - ' + (window.SITE_VERSION || '') + ' · ' + open;
    } else {
      el.textContent = open;
    }
  }
  window.renderFooterUpdate = renderUpdate;   // index.html зовёт после загрузки changelog

  // ─── тост ─────────────────────────────────────────────────────────
  window.showStub = function (msg) {
    var el = document.getElementById('dt-toast');
    if (!el) return;
    el.textContent = msg || ((typeof t === 'function') ? t('toast.default') : '');
    el.classList.add('show');
    clearTimeout(showStub._t);
    var duration = 1800 + Math.max(0, (el.textContent.length - 20)) * 40;
    showStub._t = setTimeout(function () { el.classList.remove('show'); }, duration);
  };

  // ─── установка приложения ─────────────────────────────────────────
  // beforeinstallprompt приходит только в Chromium и только когда манифест
  // и иконки на месте. Событие надо перехватить и придержать: prompt()
  // можно вызвать один раз и только по жесту пользователя, поэтому вызов
  // живёт в installApp(), а не в обработчике.
  // На iOS события нет вовсе — там установка вручную, показываем инструкцию.
  // Подпись переключаем через data-i18n, иначе смена языка затрёт её обратно.
  var installEvent = null;

  function dtInstallSub(key) {
    var el = document.getElementById('install-sub');
    if (!el) return;
    el.setAttribute('data-i18n', key);
    el.innerHTML = (typeof t === 'function') ? t(key) : el.innerHTML;
  }

  function dtIsStandalone() {
    return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches)
        || window.navigator.standalone === true;
  }

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    installEvent = e;
    dtInstallSub('info.install.ready');
  });

  window.addEventListener('appinstalled', function () {
    installEvent = null;
    dtInstallSub('info.install.done');
  });

  window.installApp = function () {
    if (dtIsStandalone()) { dtInstallSub('info.install.done'); return; }
    if (installEvent) {
      installEvent.prompt();
      installEvent.userChoice.then(function (res) {
        if (res && res.outcome === 'accepted') dtInstallSub('info.install.done');
        installEvent = null;
      });
      return;
    }
    var ios = /iphone|ipad|ipod/i.test(navigator.userAgent)
           || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    showStub(t(ios ? 'toast.install_ios' : 'toast.install_na'));
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mount);
  } else {
    mount();
  }
})();
