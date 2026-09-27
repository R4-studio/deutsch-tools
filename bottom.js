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

  // ─── модалка «Установка на iPhone» ────────────────────────────────
  // На iOS beforeinstallprompt не существует — поставить приложение может
  // только сам пользователь через меню «Поделиться». Пункт меню называется
  // по-разному в зависимости от языка телефона («На экран „Домой“»,
  // «Add to Home Screen», «Zum Home-Bildschirm»), поэтому инструкция
  // опирается на значки: они одинаковые на любом языке.
  // Разметка лежит здесь, а не в HTML страниц: блок нужен на всех страницах,
  // а иконки инлайном — чтобы не тянуть три отдельных запроса ради модалки,
  // которую большинство никогда не откроет.
  var INSTALL_HTML = "<div class=\"dt-install-modal\" id=\"dt-install-modal\">\n  <div class=\"dt-install-modal__card\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"dt-install-title\">\n    <button class=\"dt-install-modal__close\" id=\"dt-install-close\" aria-label=\"Закрыть\" data-i18n-aria=\"install.modal.close\">×</button>\n    <h2 class=\"dt-install-modal__title\" id=\"dt-install-title\" data-i18n=\"install.modal.title\">Установка на iPhone</h2>\n    <div class=\"dt-install-modal__steps\">\n      <div class=\"dt-install-step\">\n        <span class=\"dt-install-step__icon\"><svg viewBox=\"0 0 49 63\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\" aria-hidden=\"true\"><path d=\"M23.0552 0.0498339C21.8897 0.268367 4.40708 16.4034 4.15212 17.5325C3.64221 19.5357 6.33745 21.2839 8.12213 20.082C8.5592 19.7906 11.7279 16.9861 15.1152 13.8538L21.3069 8.17197L21.4162 24.1613L21.4891 40.187L22.436 40.9155C23.6744 41.8624 24.4028 41.8624 25.6412 40.9155L26.5882 40.187L26.661 24.2341L26.7703 8.28123L33.5084 14.4002C39.7001 20.082 40.3557 20.5555 41.4848 20.4826C43.051 20.337 44.2893 18.8801 43.9251 17.5325C43.7794 16.9497 40.465 13.6717 34.9652 8.68188C30.1575 4.31122 26.0054 0.632588 25.7504 0.486899C24.9856 0.0862561 23.8929 -0.0958546 23.0552 0.0498339Z\" fill=\"currentColor\"/><path d=\"M5.49974 29.1147C4.62561 29.3696 3.27799 30.2438 2.51313 30.9722C0.182111 33.2304 0 34.3231 0 45.2133C0 50.9315 0.145689 55.4843 0.400643 56.3949C0.910553 58.3981 2.54955 60.4013 4.47992 61.4575C6.00965 62.2588 6.11892 62.2588 24.0386 62.2588C41.9583 62.2588 42.0676 62.2588 43.5973 61.4575C45.5277 60.4013 47.1667 58.3981 47.6766 56.3949C48.3322 53.8453 48.1865 36.3991 47.5309 34.1409C46.4018 30.3166 42.6503 28.3134 37.3327 28.6412C35.293 28.7505 34.7467 28.9326 34.2004 29.6246C33.3262 30.7173 33.3262 31.9192 34.2004 32.9754C34.8196 33.7403 35.2566 33.8496 38.134 33.9588C40.3557 34.0681 41.5212 34.2866 41.9947 34.6873C42.541 35.1972 42.6139 36.1806 42.6139 45.3954C42.6139 55.0472 42.5775 55.5572 41.8854 56.2492C41.1934 56.9412 40.6835 56.9776 24.0386 56.9776C7.39369 56.9776 6.88378 56.9412 6.19176 56.2492C5.49974 55.5572 5.46332 55.0472 5.46332 45.3954C5.46332 36.1806 5.53616 35.1972 6.11892 34.6873C6.55598 34.2866 7.72149 34.0681 9.94324 33.9588C12.8206 33.8496 13.2577 33.7403 13.8768 32.9754C14.2411 32.5019 14.5689 31.7735 14.5689 31.3C14.5689 30.8265 14.2411 30.0981 13.8768 29.6246C13.2577 28.8597 12.8206 28.7505 10.1254 28.6776C8.26782 28.6048 6.48314 28.7869 5.49974 29.1147Z\" fill=\"currentColor\"/></svg></span>\n        <span class=\"dt-install-step__text\" data-i18n=\"install.modal.step1\"></span>\n      </div>\n      <div class=\"dt-install-step\">\n        <span class=\"dt-install-step__icon\"><svg viewBox=\"0 0 63 63\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\" aria-hidden=\"true\"><path d=\"M31.1284 16.6875C32.277 16.6875 33.2082 17.6187 33.2082 18.7673V29.0496H43.4905C44.6392 29.0497 45.5703 29.9808 45.5703 31.1294C45.5703 32.278 44.6392 33.2091 43.4905 33.2091H33.2082V43.4905C33.2082 44.6392 32.277 45.5703 31.1284 45.5703C29.9798 45.5703 29.0487 44.6392 29.0487 43.4905V33.2091H18.7673C17.6187 33.2091 16.6875 32.278 16.6875 31.1294C16.6875 29.9808 17.6187 29.0497 18.7673 29.0496H29.0487V18.7673C29.0487 17.6187 29.9798 16.6875 31.1284 16.6875Z\" fill=\"currentColor\"/><path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M50.8426 0.00381429C57.1743 0.164152 62.2578 5.34717 62.2578 11.7175V50.5403L62.254 50.8426C62.0962 57.0738 57.0738 62.0962 50.8426 62.254L50.5403 62.2578H11.7175L11.4152 62.254C5.18398 62.0962 0.161608 57.0738 0.00381429 50.8426L0 50.5403V11.7175C0 5.34717 5.08353 0.164148 11.4152 0.00381429L11.7175 0H50.5403L50.8426 0.00381429ZM11.7175 3.90583C7.40324 3.90583 3.90583 7.40324 3.90583 11.7175V50.5403C3.90584 54.8546 7.40324 58.352 11.7175 58.352H50.5403C54.8546 58.352 58.352 54.8546 58.352 50.5403V11.7175C58.352 7.40324 54.8546 3.90584 50.5403 3.90583H11.7175Z\" fill=\"currentColor\"/></svg></span>\n        <span class=\"dt-install-step__text\" data-i18n=\"install.modal.step2\"></span>\n      </div>\n      <div class=\"dt-install-step\">\n        <span class=\"dt-install-step__icon dt-install-step__icon--result\"><svg viewBox=\"0 0 58 75\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\" aria-hidden=\"true\"><rect width=\"58\" height=\"58\" rx=\"16\" fill=\"#0E0E0E\"/><path d=\"M38.044 48.4122L34.0013 44.3331L26.8199 51.5545C25.7919 52.5862 24.1224 52.5898 23.0944 51.5582L21.2389 49.7027L30.2903 40.6002L25.2925 35.555L28.9962 31.8149L41.7623 44.694L38.044 48.4122Z\" fill=\"#FAFAFA\"/><path d=\"M23.6011 44.6283L19.8974 48.3612L18.031 46.4802C17.0139 45.4558 17.0139 43.8008 18.031 42.7765C19.059 41.7376 20.7358 41.7376 21.7638 42.7765L23.6011 44.6283Z\" fill=\"#FAFAFA\"/><path d=\"M26.459 27.4186L14.6153 39.3608C13.5873 40.3997 11.9104 40.3997 10.8824 39.3608C9.86537 38.3365 9.86537 36.6815 10.8824 35.6571L21.2207 25.2314V13.147L14.6189 19.8071C13.5909 20.846 11.9141 20.846 10.8861 19.8071C9.86901 18.7827 9.86901 17.1278 10.8861 16.1034L20.1453 6.764C21.5852 5.31314 23.9292 5.30585 25.3764 6.74941C26.069 7.44203 26.459 8.38254 26.459 9.35949V27.4186Z\" fill=\"#FAFAFA\"/><path d=\"M32.616 6.74941C34.0633 5.30585 36.4072 5.3095 37.8471 6.764L47.1136 16.1071C48.1307 17.1314 48.1307 18.7864 47.1136 19.8107C46.0857 20.8497 44.4088 20.846 43.3808 19.8107L36.779 13.1507V25.2314L47.1173 35.6571C48.1343 36.6815 48.1343 38.3365 47.1173 39.3608C46.0893 40.3997 44.4124 40.3997 43.3844 39.3608L31.5407 27.4186V9.36314C31.5407 8.38254 31.9307 7.44203 32.6233 6.75306L32.616 6.74941Z\" fill=\"#FAFAFA\"/><path d=\"M13.892 72.1364C13.3466 72.1364 12.8651 71.9986 12.4474 71.723C12.0298 71.4446 11.7031 71.0526 11.4673 70.5469C11.2315 70.0384 11.1136 69.4375 11.1136 68.7443C11.1136 68.0568 11.2315 67.4602 11.4673 66.9545C11.7031 66.4489 12.0313 66.0582 12.4517 65.7827C12.8722 65.5071 13.358 65.3693 13.9091 65.3693C14.3352 65.3693 14.6719 65.4403 14.919 65.5824C15.169 65.7216 15.3594 65.8807 15.4901 66.0597C15.6236 66.2358 15.7273 66.3807 15.8011 66.4943H15.8864V63.2727H16.892V72H15.9205V70.9943H15.8011C15.7273 71.1136 15.6222 71.2642 15.4858 71.446C15.3494 71.625 15.1548 71.7855 14.902 71.9276C14.6491 72.0668 14.3125 72.1364 13.892 72.1364ZM14.0284 71.233C14.4318 71.233 14.7727 71.1278 15.0511 70.9176C15.3295 70.7045 15.5412 70.4105 15.6861 70.0355C15.831 69.6577 15.9034 69.2216 15.9034 68.7273C15.9034 68.2386 15.8324 67.8111 15.6903 67.4446C15.5483 67.0753 15.3381 66.7884 15.0597 66.5838C14.7813 66.3764 14.4375 66.2727 14.0284 66.2727C13.6023 66.2727 13.2472 66.3821 12.9631 66.6009C12.6818 66.8168 12.4702 67.1108 12.3281 67.483C12.1889 67.8523 12.1193 68.267 12.1193 68.7273C12.1193 69.1932 12.1903 69.6165 12.3324 69.9972C12.4773 70.375 12.6903 70.6761 12.9716 70.9006C13.2557 71.1222 13.608 71.233 14.0284 71.233ZM21.6179 72.1364C20.9872 72.1364 20.4432 71.9972 19.9858 71.7188C19.5312 71.4375 19.1804 71.0455 18.9332 70.5426C18.6889 70.0369 18.5668 69.4489 18.5668 68.7784C18.5668 68.108 18.6889 67.517 18.9332 67.0057C19.1804 66.4915 19.5241 66.0909 19.9645 65.804C20.4077 65.5142 20.9247 65.3693 21.5156 65.3693C21.8565 65.3693 22.1932 65.4261 22.5256 65.5398C22.858 65.6534 23.1605 65.8381 23.4332 66.0938C23.706 66.3466 23.9233 66.6818 24.0852 67.0994C24.2472 67.517 24.3281 68.0312 24.3281 68.642V69.0682H19.2827V68.1989H23.3054C23.3054 67.8295 23.2315 67.5 23.0838 67.2102C22.9389 66.9205 22.7315 66.6918 22.4616 66.5241C22.1946 66.3565 21.8793 66.2727 21.5156 66.2727C21.1151 66.2727 20.7685 66.3722 20.4759 66.571C20.1861 66.767 19.9631 67.0227 19.8068 67.3381C19.6506 67.6534 19.5724 67.9915 19.5724 68.3523V68.9318C19.5724 69.4261 19.6577 69.8452 19.8281 70.1889C20.0014 70.5298 20.2415 70.7898 20.5483 70.9688C20.8551 71.1449 21.2116 71.233 21.6179 71.233C21.8821 71.233 22.1207 71.196 22.3338 71.1222C22.5497 71.0455 22.7358 70.9318 22.892 70.7812C23.0483 70.6278 23.169 70.4375 23.2543 70.2102L24.2259 70.483C24.1236 70.8125 23.9517 71.1023 23.7102 71.3523C23.4688 71.5994 23.1705 71.7926 22.8153 71.9318C22.4602 72.0682 22.0611 72.1364 21.6179 72.1364ZM29.983 69.3239V65.4545H30.9886V72H29.983V70.892H29.9148C29.7614 71.2244 29.5227 71.5071 29.1989 71.7401C28.875 71.9702 28.4659 72.0852 27.9716 72.0852C27.5625 72.0852 27.1989 71.9957 26.8807 71.8168C26.5625 71.6349 26.3125 71.3622 26.1307 70.9986C25.9489 70.6321 25.858 70.1705 25.858 69.6136V65.4545H26.8636V69.5455C26.8636 70.0227 26.9972 70.4034 27.2642 70.6875C27.5341 70.9716 27.8778 71.1136 28.2955 71.1136C28.5455 71.1136 28.7997 71.0497 29.0582 70.9219C29.3196 70.794 29.5384 70.598 29.7145 70.3338C29.8935 70.0696 29.983 69.733 29.983 69.3239ZM35.6772 65.4545V66.3068H32.2852V65.4545H35.6772ZM33.2738 63.8864H34.2795V70.125C34.2795 70.4091 34.3207 70.6222 34.4031 70.7642C34.4883 70.9034 34.5962 70.9972 34.7269 71.0455C34.8604 71.0909 35.0011 71.1136 35.1488 71.1136C35.2596 71.1136 35.3505 71.108 35.4215 71.0966C35.4925 71.0824 35.5494 71.071 35.592 71.0625L35.7965 71.9659C35.7283 71.9915 35.6332 72.017 35.511 72.0426C35.3888 72.071 35.234 72.0852 35.0465 72.0852C34.7624 72.0852 34.484 72.0241 34.2113 71.902C33.9414 71.7798 33.717 71.5938 33.538 71.3438C33.3619 71.0938 33.2738 70.7784 33.2738 70.3977V63.8864ZM37.19 72V65.4545H38.1616V66.4432H38.2298C38.3491 66.1193 38.565 65.8565 38.8775 65.6548C39.19 65.4531 39.5423 65.3523 39.9343 65.3523C40.0082 65.3523 40.1005 65.3537 40.2113 65.3565C40.3221 65.3594 40.4059 65.3636 40.4627 65.3693V66.392C40.4286 66.3835 40.3505 66.3707 40.2283 66.3537C40.109 66.3338 39.9826 66.3239 39.8491 66.3239C39.5309 66.3239 39.2468 66.3906 38.9968 66.5241C38.7496 66.6548 38.5536 66.8366 38.4087 67.0696C38.2667 67.2997 38.1957 67.5625 38.1957 67.858V72H37.19ZM43.581 72.1534C43.1662 72.1534 42.7898 72.0753 42.4517 71.919C42.1136 71.7599 41.8452 71.5312 41.6463 71.233C41.4474 70.9318 41.348 70.5682 41.348 70.142C41.348 69.767 41.4219 69.4631 41.5696 69.2301C41.7173 68.9943 41.9148 68.8097 42.1619 68.6761C42.4091 68.5426 42.6818 68.4432 42.9801 68.3778C43.2813 68.3097 43.5838 68.2557 43.8878 68.2159C44.2855 68.1648 44.608 68.1264 44.8551 68.1009C45.1051 68.0724 45.2869 68.0256 45.4006 67.9602C45.517 67.8949 45.5753 67.7812 45.5753 67.6193V67.5852C45.5753 67.1648 45.4602 66.8381 45.2301 66.6051C45.0028 66.3722 44.6577 66.2557 44.1946 66.2557C43.7145 66.2557 43.3381 66.3608 43.0653 66.571C42.7926 66.7812 42.6009 67.0057 42.4901 67.2443L41.5355 66.9034C41.706 66.5057 41.9332 66.196 42.2173 65.9744C42.5043 65.75 42.8168 65.5937 43.1548 65.5057C43.4957 65.4148 43.831 65.3693 44.1605 65.3693C44.3707 65.3693 44.6122 65.3949 44.8849 65.446C45.1605 65.4943 45.4261 65.5952 45.6818 65.7486C45.9403 65.902 46.1548 66.1335 46.3253 66.4432C46.4957 66.7528 46.581 67.1676 46.581 67.6875V72H45.5753V71.1136H45.5241C45.456 71.2557 45.3423 71.4077 45.1832 71.5696C45.0241 71.7315 44.8125 71.8693 44.5483 71.983C44.2841 72.0966 43.9616 72.1534 43.581 72.1534ZM43.7344 71.25C44.1321 71.25 44.4673 71.1719 44.7401 71.0156C45.0156 70.8594 45.223 70.6577 45.3622 70.4105C45.5043 70.1634 45.5753 69.9034 45.5753 69.6307V68.7102C45.5327 68.7614 45.4389 68.8082 45.294 68.8509C45.152 68.8906 44.9872 68.9261 44.7997 68.9574C44.6151 68.9858 44.4347 69.0114 44.2585 69.0341C44.0852 69.054 43.9446 69.071 43.8366 69.0852C43.5753 69.1193 43.331 69.1747 43.1037 69.2514C42.8793 69.3253 42.6974 69.4375 42.5582 69.5881C42.4219 69.7358 42.3537 69.9375 42.3537 70.1932C42.3537 70.5426 42.483 70.8068 42.7415 70.9858C43.0028 71.1619 43.3338 71.25 43.7344 71.25Z\" fill=\"currentColor\"/></svg></span>\n        <span class=\"dt-install-step__text\" data-i18n=\"install.modal.step3\"></span>\n      </div>\n    </div>\n    <button class=\"dt-install-modal__ok\" id=\"dt-install-ok\" data-i18n=\"install.modal.ok\">Понятно</button>\n  </div>\n</div>";

  function installModalEl() {
    var el = document.getElementById('dt-install-modal');
    if (el) return el;
    var wrap = document.createElement('div');
    wrap.innerHTML = INSTALL_HTML;
    el = wrap.firstElementChild;
    document.body.appendChild(el);
    el.addEventListener('click', function (e) {
      // клик мимо карточки — закрыть
      if (e.target === el) closeInstallModal();
    });
    el.querySelector('#dt-install-close').addEventListener('click', closeInstallModal);
    el.querySelector('#dt-install-ok').addEventListener('click', closeInstallModal);
    return el;
  }

  function onInstallKey(e) {
    if (e.key === 'Escape' || e.key === 'Esc') closeInstallModal();
  }

  function openInstallModal() {
    var el = installModalEl();
    if (typeof applyI18n === 'function') {
      // модалку добавили в DOM после общего прохода перевода — переводим её
      var pageLang = document.documentElement.getAttribute('lang');
      var fixed = document.documentElement.hasAttribute('data-lang-fixed')
               || pageLang === 'de';
      applyI18n();
      if (fixed) document.documentElement.setAttribute('lang', pageLang);
    }
    el.classList.add('open');
    document.addEventListener('keydown', onInstallKey);
  }

  function closeInstallModal() {
    var el = document.getElementById('dt-install-modal');
    if (el) el.classList.remove('open');
    document.removeEventListener('keydown', onInstallKey);
  }

  // ─── установка приложения ─────────────────────────────────────────
  // beforeinstallprompt приходит только в Chromium и только когда манифест
  // и иконки на месте. Событие надо перехватить и придержать: prompt()
  // можно вызвать один раз и только по жесту пользователя, поэтому вызов
  // живёт в installApp(), а не в обработчике.
  // На iOS события нет вовсе — там установка вручную, открываем модалку.
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
    // MacIntel + тач — это iPad, который с iPadOS 13 представляется десктопом
    var ios = /iphone|ipad|ipod/i.test(navigator.userAgent)
           || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    if (ios) { openInstallModal(); return; }
    showStub(t('toast.install_na'));
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mount);
  } else {
    mount();
  }
})();
