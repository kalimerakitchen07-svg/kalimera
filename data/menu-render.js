/* Kalimera — paylaşımlı menü PDF renderer (menu-pdf.html + admin ortak).
   Menüyü 2 dengeli A4 sayfaya böler, forest markalı, sanatsal düzen.
   Kullanım: KalimeraMenu.render(menuObj, 'en', {showPrice:true}) -> HTML (sayfa div'leri) */
window.KalimeraMenu = (function () {
  var TIER = { tr: ['Hafif', 'Dengeli', 'Zengin'], en: ['Light', 'Balanced', 'Rich'], ru: ['Лёгкое', 'Сбаланс.', 'Сытное'] };
  var NAMES = {
    tr: { 'turkish-night': 'Türk Gecesi', 'mediterranean-night': 'Akdeniz Gecesi', 'bbq-menu': 'Mangal Menü', 'chicken-bbq': 'Tavuk Mangal', 'fish-bbq': 'Balık Mangal', 'breakfast': 'Türk Kahvaltısı', 'chefs-table': 'Şefin Sofrası' },
    en: { 'turkish-night': 'Turkish Night', 'mediterranean-night': 'Mediterranean Night', 'bbq-menu': 'BBQ Menu', 'chicken-bbq': 'Chicken BBQ', 'fish-bbq': 'Fish BBQ', 'breakfast': 'Turkish Breakfast', 'chefs-table': "Chef's Table" },
    ru: { 'turkish-night': 'Турецкий вечер', 'mediterranean-night': 'Средиземноморский вечер', 'bbq-menu': 'Барбекю', 'chicken-bbq': 'Курица гриль', 'fish-bbq': 'Рыба гриль', 'breakfast': 'Турецкий завтрак', 'chefs-table': 'Стол шефа' }
  };
  var BOLUM = {
    'Ana Yemek': { tr: 'Ana Yemek', en: 'Main Course', ru: 'Основное блюдо' },
    'Meze': { tr: 'Mezeler', en: 'Starters', ru: 'Закуски' },
    'Salata': { tr: 'Salata', en: 'Salad', ru: 'Салат' },
    'Tatlı': { tr: 'Tatlı', en: 'Dessert', ru: 'Десерт' },
    'Kahvaltı': { tr: 'Kahvaltı', en: 'Breakfast', ru: 'Завтрак' },
    'Izgara': { tr: 'Izgara', en: 'From the Grill', ru: 'С гриля' },
    'Izgara Tavuk': { tr: 'Izgara Tavuk', en: 'Grilled Chicken', ru: 'Курица гриль' },
    'Yan Lezzet': { tr: 'Yan Lezzetler', en: 'Sides', ru: 'Гарниры' },
    'Sos': { tr: 'Soslar', en: 'Sauces', ru: 'Соусы' },
    'Ekmek': { tr: 'Ekmek', en: 'Breads', ru: 'Хлеб' },
    'Sıcak': { tr: 'Sıcak', en: 'Hot Dishes', ru: 'Горячее' },
    'Sofra': { tr: 'Sofra', en: 'The Spread', ru: 'Стол' },
    'İçecek': { tr: 'İçecekler', en: 'Beverages', ru: 'Напитки' }
  };
  var UI = {
    per: { tr: 'Kişi Başı', en: 'Per Person', ru: 'На человека' },
    cont: { tr: 'devamı', en: 'continued', ru: 'продолжение' },
    tagline: { tr: 'Taze, mevsimlik ve açık ateşte · Kalkan sofrası', en: 'Fresh, seasonal & over open flame · a Kalkan table', ru: 'Свежее, сезонное, на открытом огне · стол Калкана' },
    note: { tr: 'Tüm yemekler taze, günlük temin edilen malzemelerle hazırlanır.', en: 'All dishes prepared with fresh, daily-sourced ingredients.', ru: 'Все блюда готовятся из свежих продуктов, поставляемых ежедневно.' },
    calnote: { tr: 'Kalori değerleri kişi başı yaklaşık porsiyon değerleridir.', en: 'Calorie values are approximate per-person portions.', ru: 'Калорийность указана приблизительно на порцию.' }
  };
  function menuName(slug, name, l) { return (NAMES[l] && NAMES[l][slug]) || name; }
  function nameL(d, l) { return l === 'en' ? (d.ad_en || d.ad) : l === 'ru' ? (d.ad_ru || d.ad) : (d.ad_tr || d.ad); }
  function descL(d, l) { return l === 'en' ? (d.desc_en || '') : l === 'ru' ? (d.desc_ru || '') : (d.desc || ''); }
  function tierL(k, l) { return TIER[l][k < 180 ? 0 : k < 400 ? 1 : 2]; }

  function sectionsOf(menu) {
    var groups = {}, order = [];
    menu.dishes.forEach(function (d) { if (!groups[d.bolum]) { groups[d.bolum] = []; order.push(d.bolum); } groups[d.bolum].push(d); });
    return order.map(function (b) { return { bolum: b, dishes: groups[b] }; });
  }
  // Yemek-ağırlığı bazlı dengeli 2'ye bölme (gerekirse bölüm ortasından; sayfa 2'de "devamı")
  function splitSections(sections, l) {
    var flat = [];
    sections.forEach(function (s) { s.dishes.forEach(function (d) { flat.push({ d: d, bolum: s.bolum }); }); });
    var wt = flat.map(function (f) { return 1 + (descL(f.d, l) ? 0.8 : 0); });
    var total = wt.reduce(function (a, b) { return a + b; }, 0);
    var acc = 0, cut = flat.length;
    // Sayfa 1 kapak başlığı daha büyük → hedef %47 (sayfa 1 biraz daha hafif)
    for (var i = 0; i < flat.length; i++) { acc += wt[i]; if (acc >= total * 0.47) { cut = i + 1; break; } }
    cut = Math.max(1, Math.min(flat.length - 1, cut));
    function build(items) {
      var out = [], cur = null;
      items.forEach(function (f) { if (!cur || cur.bolum !== f.bolum) { cur = { bolum: f.bolum, dishes: [] }; out.push(cur); } cur.dishes.push(f.d); });
      return out;
    }
    var p1 = build(flat.slice(0, cut)), p2 = build(flat.slice(cut));
    if (p1.length && p2.length && p1[p1.length - 1].bolum === p2[0].bolum) p2[0].cont = true;
    return [p1, p2];
  }

  function dishHtml(d, l) {
    var desc = descL(d, l);
    return '<div style="margin:0 auto;max-width:560px;text-align:center">'
      + '<div style="font-family:\'Cormorant Garamond\',Georgia,serif;font-size:22px;color:#F2EBDA;line-height:1.25">' + nameL(d, l) + '</div>'
      + (desc ? '<div style="font-style:italic;font-size:12.5px;color:rgba(242,235,218,.62);margin-top:4px;line-height:1.5">' + desc + '</div>' : '')
      + '<div style="font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:rgba(217,146,59,.8);margin-top:5px">' + tierL(d.kcal, l) + ' &middot; ' + d.kcal + ' kcal</div>'
      + '</div>';
  }
  function sectionHtml(sec, l) {
    return '<div>'
      + '<div style="text-align:center;font-size:11px;letter-spacing:.3em;text-transform:uppercase;color:#D9923B;margin-bottom:14px">'
      + ((BOLUM[sec.bolum] || {})[l] || sec.bolum) + (sec.cont ? ' <span style="opacity:.6">· ' + UI.cont[l] + '</span>' : '') + '</div>'
      + '<div style="display:flex;flex-direction:column;gap:16px">' + sec.dishes.map(function (d) { return dishHtml(d, l); }).join('') + '</div>'
      + '</div>';
  }
  function ornament(color) { return '<div style="text-align:center;color:' + (color || '#D9923B') + ';font-size:13px;letter-spacing:.6em;margin:0 0 0 .6em">&#10022;</div>'; }

  function pageOpen() {
    return '<div style="width:794px;height:1123px;box-sizing:border-box;padding:56px 60px;background:#15342B;color:#F2EBDA;font-family:Inter,Arial,sans-serif;display:flex;flex-direction:column;overflow:hidden;position:relative">'
      + '<div style="position:absolute;inset:18px;border:1px solid rgba(217,146,59,.22);pointer-events:none"></div>';
  }

  function render(menu, l, opts) {
    opts = opts || {};
    var slug = menu.slug, title = menuName(slug, menu.name, l);
    var sections = sectionsOf(menu);
    var parts = splitSections(sections, l);
    var priceLine = opts.showPrice !== false
      ? '<div style="text-align:center;font-size:12px;letter-spacing:.14em;color:#D9923B;margin-top:6px">' + UI.per[l] + ' &middot; £' + menu.price_gbp + '</div>' : '';

    // PAGE 1 — kapak başlığı + ilk bölümler
    var p1 = pageOpen()
      + '<div style="text-align:center">'
      + ornament()
      + '<div style="font-family:\'Cormorant Garamond\',serif;font-style:italic;font-size:30px;color:#D9923B;margin-top:6px">Kalimera Kitchen</div>'
      + '<div style="font-size:9px;letter-spacing:.32em;text-transform:uppercase;color:rgba(242,235,218,.5);margin-top:2px">Kalkan &middot; Kaş &middot; Antalya</div>'
      + '<div style="width:56px;height:1px;background:rgba(217,146,59,.5);margin:18px auto 0"></div>'
      + '<h1 style="font-family:\'Cormorant Garamond\',serif;font-style:italic;font-size:38px;color:#F2EBDA;margin:16px 0 0">' + title + '</h1>'
      + priceLine
      + '<div style="font-style:italic;font-size:11.5px;color:rgba(242,235,218,.55);margin-top:8px">' + UI.tagline[l] + '</div>'
      + '</div>'
      + '<div style="flex:1;display:flex;flex-direction:column;justify-content:space-around;padding:22px 0 8px">'
      + parts[0].map(function (s) { return sectionHtml(s, l); }).join('')
      + '</div>'
      + '<div style="text-align:center;font-size:10px;letter-spacing:.3em;color:rgba(217,146,59,.6)">I</div>'
      + '</div>';

    // PAGE 2 — devam bölümleri + kapanış
    var p2 = pageOpen()
      + '<div style="text-align:center">'
      + '<div style="font-family:\'Cormorant Garamond\',serif;font-style:italic;font-size:24px;color:#F2EBDA">' + title + '</div>'
      + ornament('rgba(217,146,59,.7)')
      + '</div>'
      + '<div style="flex:1;display:flex;flex-direction:column;justify-content:space-around;padding:14px 0">'
      + parts[1].map(function (s) { return sectionHtml(s, l); }).join('')
      + '</div>'
      + '<div style="text-align:center">'
      + '<div style="font-style:italic;font-size:11px;color:rgba(242,235,218,.5);max-width:440px;margin:0 auto 6px">' + UI.note[l] + '</div>'
      + '<div style="font-size:9px;color:rgba(242,235,218,.4);margin-bottom:12px">' + UI.calnote[l] + '</div>'
      + '<div style="width:56px;height:1px;background:rgba(217,146,59,.4);margin:0 auto 10px"></div>'
      + '<div style="font-size:10px;color:rgba(242,235,218,.6)">Şef Nurdan Değirmenci &middot; +90 539 743 07 81 &middot; kalimerakitchen.com</div>'
      + '<div style="font-size:10px;letter-spacing:.3em;color:rgba(217,146,59,.6);margin-top:8px">II</div>'
      + '</div>'
      + '</div>';

    return p1 + p2;
  }

  return { render: render, NAMES: NAMES, menuName: menuName };
})();
