/* Kalimera — menü detay sayfası zenginleştirme: kalori rozetleri + sticky WhatsApp çubuğu.
   Tek dosya, 6 menü sayfasında ortak. Dil: kkLang (tr/en/ru). */
(function () {
  var TIER = { tr: ['Hafif', 'Dengeli', 'Zengin'], en: ['Light', 'Balanced', 'Rich'], ru: ['Лёгкое', 'Сбаланс.', 'Сытное'] };
  var NAMES = {
    tr: { 'turkish-night': 'Türk Gecesi', 'mediterranean-night': 'Akdeniz Gecesi', 'bbq-menu': 'Mangal Menü', 'chicken-bbq': 'Tavuk Mangal', 'fish-bbq': 'Balık Mangal', 'breakfast': 'Türk Kahvaltısı' },
    en: { 'turkish-night': 'Turkish Night', 'mediterranean-night': 'Mediterranean Night', 'bbq-menu': 'BBQ Menu', 'chicken-bbq': 'Chicken BBQ', 'fish-bbq': 'Fish BBQ', 'breakfast': 'Turkish Breakfast' },
    ru: { 'turkish-night': 'Турецкий вечер', 'mediterranean-night': 'Средиземноморский вечер', 'bbq-menu': 'Барбекю', 'chicken-bbq': 'Курица гриль', 'fish-bbq': 'Рыба гриль', 'breakfast': 'Турецкий завтрак' }
  };
  var ASK = { tr: 'Bu Menüyü Sor', en: 'Ask About This Menu', ru: 'Спросить об этом меню' };
  var CAL_NOTE = { tr: 'kişi başı yaklaşık', en: 'approx. per person', ru: 'прибл. на человека' };
  function msg(l, n) {
    return ({
      tr: 'Merhaba, Kalimera Kitchen — ' + n + ' menüsünü sormak istiyorum. Kaç kişi ve hangi tarih için müsaitsiniz?',
      en: 'Hello, Kalimera Kitchen — I would like to ask about the ' + n + ' menu. Availability for my dates and party size?',
      ru: 'Здравствуйте, Kalimera Kitchen — вопрос по меню ' + n + '. Свободны ли даты для моей группы?'
    })[l];
  }
  function lang() { try { var l = localStorage.getItem('kkLang'); return (l && TIER[l]) ? l : 'tr'; } catch (e) { return 'tr'; } }
  function tierOf(k) { return k < 180 ? 0 : k < 400 ? 1 : 2; }
  function norm(s) { return (s || '').toLowerCase().replace(/[\s.,·’'"-]+/g, '').trim(); }
  var SLUG, MENU;

  function injectStyle() {
    if (document.getElementById('kk-enh-style')) return;
    var css = ''
      + '.kcal-badge{display:block;margin-top:.35rem;font-family:Inter,system-ui,sans-serif;font-size:.6rem;letter-spacing:.09em;text-transform:uppercase;color:rgba(242,235,218,.4)}'
      + '.kcal-badge .kb-tier{color:rgba(217,146,59,.9);font-weight:600}'
      + '.kcal-badge .kb-kcal{color:rgba(242,235,218,.55)}'
      + '#kk-sticky{position:fixed;left:0;right:0;bottom:0;z-index:60;display:none;gap:.6rem;padding:.6rem .8rem;padding-bottom:calc(.6rem + env(safe-area-inset-bottom));background:rgba(21,52,43,.96);-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px);border-top:1px solid rgba(242,235,218,.12)}'
      + '#kk-sticky-wa{flex:1;text-align:center;background:linear-gradient(180deg,#E4A04C,#C77A2C);color:#221206;font-weight:700;border-radius:9999px;padding:.8rem;font-size:.82rem;font-family:Inter,sans-serif;text-decoration:none;box-shadow:0 6px 18px -6px rgba(217,146,59,.5)}'
      + '#kk-sticky-call{display:grid;place-items:center;min-width:2.9rem;border-radius:9999px;border:1px solid rgba(242,235,218,.22);color:#F2EBDA;text-decoration:none;font-size:1.1rem}'
      + '@media(max-width:640px){#kk-sticky{display:flex}body{padding-bottom:4.6rem !important}}';
    var s = document.createElement('style'); s.id = 'kk-enh-style'; s.textContent = css; document.head.appendChild(s);
  }

  function addCalories() {
    var l = lang();
    var lines = Array.prototype.slice.call(document.querySelectorAll('.menu-line'));
    var byName = {}; MENU.dishes.forEach(function (d) { byName[norm(d.ad)] = d; });
    // Sıra sayısı verideki yemek sayısıyla birebir ise sıra-bazlı (dil bağımsız), değilse ada göre eşle
    var useOrder = lines.length === MENU.dishes.length;
    lines.forEach(function (line, i) {
      if (line.querySelector('.kcal-badge')) return;
      var nameEl = line.querySelector('.name'); if (!nameEl) return;
      var d = useOrder ? MENU.dishes[i] : byName[norm(nameEl.textContent)];
      if (!d) return;
      var t = tierOf(d.kcal);
      var b = document.createElement('div');
      b.className = 'kcal-badge';
      b.innerHTML = '<span class="kb-tier" data-t="' + t + '">' + TIER[l][t] + '</span> <span class="kb-kcal">· ' + d.kcal + ' kcal</span>';
      line.appendChild(b);
    });
  }

  function buildSticky() {
    if (document.getElementById('kk-sticky')) return;
    var bar = document.createElement('div');
    bar.id = 'kk-sticky';
    bar.innerHTML = '<a id="kk-sticky-wa" target="_blank" rel="noopener"></a>'
      + '<a id="kk-sticky-call" href="tel:+905397430781" aria-label="Ara">☎</a>';
    document.body.appendChild(bar);
    updateSticky();
  }
  function updateSticky() {
    var wa = document.getElementById('kk-sticky-wa'); if (!wa) return;
    var l = lang(), n = (NAMES[l] && NAMES[l][SLUG]) || MENU.name;
    wa.href = 'https://wa.me/905397430781?text=' + encodeURIComponent(msg(l, n));
    wa.textContent = ASK[l] + ' · £' + MENU.price_gbp;
  }

  function relabel() {
    var l = lang();
    document.querySelectorAll('.kcal-badge .kb-tier').forEach(function (el) { el.textContent = TIER[l][+el.dataset.t]; });
    updateSticky();
  }

  function run() {
    var D = window.KALIMERA_DATA; if (!D || !D.menus) return;
    SLUG = location.pathname.replace(/\/$/, '').split('/').pop().replace('.html', '');
    MENU = D.menus[SLUG]; if (!MENU) return;
    injectStyle(); addCalories(); buildSticky();
    document.querySelectorAll('.lang-btn[data-lang]').forEach(function (b) {
      b.addEventListener('click', function () { setTimeout(relabel, 40); });
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run);
  else run();
})();
