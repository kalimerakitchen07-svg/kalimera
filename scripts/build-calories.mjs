// Kalimera — reçetelerden kişi başı kalori hesabı ve menu-data.js'e ekleme
// kcal anlamı: kg/lt malzeme → kcal/100(g|ml); adet → kcal/1 adet; demet → kcal/1 demet
import { readFileSync, writeFileSync } from 'fs';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const { INGREDIENTS, MENUS } = require('/tmp/kdata.js');

// kcal referans değerleri (USDA/TurComp ~ ortalama), id → kcal (birim yukarıdaki gibi)
const KCAL = {
  // Et — kcal/100g
  M001:230, M002:250, M003:180, M004:250, M005:180, M006:120, M007:200, M008:180, M009:380,
  // Deniz — kcal/100g
  M010:100, M011:200, M012:85, M013:92, M014:86, M015:180,
  // Sebze — kcal/100g
  M016:18, M017:18, M018:18, M019:15, M020:40, M021:149, M022:77, M023:77, M024:25, M025:17,
  M026:41, M027:20, M028:31, M029:31, M030:22, M031:25, M032:31,
  // Meyve/Yeşillik
  M033:52, M034:240 /*avokado adet*/, M035:25, M036:15, M037:20,
  // Ot — kcal/1 demet (küçük)
  M038:8, M039:8, M040:8, M041:5, M042:70 /*asma yaprağı kg*/,
  M043:29, M044:30, M045:34, M046:39, M047:69, M048:32, M049:47, M050:83,
  // Bakliyat/Tahıl — kcal/100g (kuru)
  M051:360, M052:360, M053:360, M054:333, M055:364,
  // Kuruyemiş
  M056:579, M057:654, M058:559, M059:283, M060:249,
  // Süt & Peynir
  M061:64, M062:61, M063:96, M064:330, M065:717, M066:195, M067:264, M068:300, M069:380,
  M070:350, M071:350, M072:300, M073:207,
  // Yağ
  M074:884, M075:884,
  // Sos/tatlı
  M076:595, M077:304, M078:293, M079:145, M080:115, M081:250, M082:250, M083:250, M084:82,
  // İçecek
  M085:82, M086:85, M087:327, M088:53, M089:22, M090:88,
  // Baharat — kcal/100g (çok küçük miktar)
  M091:0, M092:251, M093:265, M094:331, M095:375, M096:247, M097:274, M098:282, M099:282, M100:80,
  // Unlu — adet: kcal/1 adet ; yufka kg
  M101:300 /*yufka kg*/, M102:200, M103:250, M104:600, M105:280, M106:250,
  // Diğer
  M107:78 /*yumurta adet*/, M108:387, M109:1, M110:5, M111:5, M112:20,
};

function dishKcal(dish){
  let k = 0;
  for (const it of dish.items){
    const ing = INGREDIENTS.find(i => i.id === it.id);
    if (!ing) continue;
    const kc = KCAL[it.id] ?? 0;
    if (ing.birim === 'kg' || ing.birim === 'lt') k += (kc/100) * it.qty; // qty g/ml
    else k += kc * it.qty; // adet/demet
  }
  return Math.round(k);
}

// zenginlik etiketi
function tier(k){ if (k < 180) return {tr:'Hafif', en:'Light', ru:'Лёгкое'};
  if (k < 400) return {tr:'Dengeli', en:'Balanced', ru:'Сбалансированное'};
  return {tr:'Zengin', en:'Rich', ru:'Сытное'}; }

const slugMap = {'Turkish Night':'turkish-night','Mediterranean Night':'mediterranean-night','BBQ Menu':'bbq-menu','Chicken BBQ':'chicken-bbq','Fish BBQ':'fish-bbq','Breakfast':'breakfast'};
// Canlı site kart fiyatları (menus.html ile hizalı) — excel'deki eski değerlerin yerine
const PRICE_GBP = {'turkish-night':55,'mediterranean-night':50,'bbq-menu':45,'chicken-bbq':35,'fish-bbq':45,'breakfast':25};

const menus = {};
for (const [name, m] of Object.entries(MENUS)){
  let total = 0;
  const slug = slugMap[name];
  const dishes = m.yemekler.map(d => {
    const kcal = dishKcal(d);
    total += kcal;
    return { ad:d.ad, bolum:d.bolum, kcal, items:d.items.map(it=>({id:it.id, qty:it.qty, unit:it.displayUnit})) };
  });
  menus[slug] = { name, slug, price_gbp:PRICE_GBP[slug]??m.fiyat_gbp, kcal_total:total, dishes };
}

// ── ÇAPA MENÜ: Şefin Sofrası / Chef's Table (£90) ──
// Mevcut imza yemeklerden derlenmiş tadım menüsü → gerçek reçete + kalori (Nurdan gerçekten sunabilir)
function findDish(menuName, dishAd) {
  const m = MENUS[menuName]; if (!m) return null;
  return m.yemekler.find(y => y.ad === dishAd) || null;
}
const CHEF_PICKS = [
  ['Fish BBQ', 'Asma Yaprağında Marineli Levrek'],
  ['Turkish Night', 'Özel Soslu Kuzu But'],
  ['Mediterranean Night', 'Hünkar Beğendi'],
  ['Mediterranean Night', 'Saganaki (Deniz Mahsulleri)'],
  ['Mediterranean Night', 'Humuslu Karnabahar'],
  ['Turkish Night', 'Tulum Peynirli Karpuz Roka Salatası'],
  ['Mediterranean Night', 'Zeytinyağlı Fırın Dolma Biber'],
  ['Mediterranean Night', 'İncir Çiçeği Tatlısı'],
  ['Turkish Night', 'Dondurmalı Fırın Şeftali Tatlısı'],
];
{
  const dishes = []; let total = 0;
  for (const [mn, dn] of CHEF_PICKS) {
    const d = findDish(mn, dn);
    if (!d) { console.log('UYARI çapa yemek bulunamadı:', mn, '/', dn); continue; }
    const kcal = dishKcal(d); total += kcal;
    dishes.push({ ad: d.ad, bolum: d.bolum, kcal, items: d.items.map(it => ({ id: it.id, qty: it.qty, unit: it.displayUnit })) });
  }
  menus['chefs-table'] = { name: "Chef's Table", slug: 'chefs-table', price_gbp: 90, kcal_total: total, dishes, signature: true };
}

// kcal ekle ingredients'a da (admin görüntüsü için)
const ingredients = INGREDIENTS.map(i => ({...i, kcal: KCAL[i.id] ?? 0}));

const out = { generated:'2026-07-08', currency_note:'Fiyatlar TL (Türkiye toptan tahmini)',
  kcal_note:'Kişi başı yaklaşık değerler; malzeme miktarlarından hesaplanmıştır (±%15).', ingredients, menus };
writeFileSync('data/menu-data.js', 'window.KALIMERA_DATA = ' + JSON.stringify(out, null, 2) + ';\n');

console.log('MENÜ KALORİLERİ (kişi başı, tüm porsiyon):');
for (const s of Object.keys(menus)){
  const m = menus[s];
  console.log('  ' + m.name.padEnd(22), m.kcal_total + ' kcal', '['+tier(m.kcal_total/m.dishes.length).tr+']');
  m.dishes.forEach(d => console.log('      - ' + d.ad.padEnd(42), d.kcal + ' kcal', tier(d.kcal).tr));
}
