// Her menü detay sayfasındaki minimal Menu JSON-LD'yi kalori/besin + bölüm + offer içeren
// zengin şemayla değiştirir. Kaynak: data/menu-data.js
import { readFileSync, writeFileSync } from 'fs';

const data = eval('(' + readFileSync('data/menu-data.js', 'utf8').replace('window.KALIMERA_DATA = ', '').replace(/;\s*$/, '') + ')');
const files = {
  'turkish-night': 'menu/turkish-night.html',
  'mediterranean-night': 'menu/mediterranean-night.html',
  'bbq-menu': 'menu/bbq-menu.html',
  'chicken-bbq': 'menu/chicken-bbq.html',
  'fish-bbq': 'menu/fish-bbq.html',
  'breakfast': 'menu/breakfast.html',
};
const CUISINE = {
  'turkish-night': ['Turkish', 'Anatolian', 'Mediterranean'],
  'mediterranean-night': ['Mediterranean', 'Aegean', 'Turkish'],
  'bbq-menu': ['Turkish', 'Barbecue', 'Mediterranean'],
  'chicken-bbq': ['Turkish', 'Barbecue', 'Grill'],
  'fish-bbq': ['Seafood', 'Mediterranean', 'Aegean'],
  'breakfast': ['Turkish', 'Breakfast', 'Mediterranean'],
};

function buildLd(slug) {
  const m = data.menus[slug];
  const sections = {};
  m.dishes.forEach(d => { (sections[d.bolum] = sections[d.bolum] || []).push(d); });
  const ld = {
    '@context': 'https://schema.org',
    '@type': 'Menu',
    name: m.name,
    inLanguage: 'tr-TR',
    url: `https://www.kalimerakitchen.com/menu/${slug}`,
    provider: {
      '@type': ['Restaurant', 'CateringService'],
      name: 'Kalimera Kitchen',
      telephone: '+905397430781',
      url: 'https://www.kalimerakitchen.com',
      servesCuisine: CUISINE[slug],
      areaServed: ['Kalkan', 'Kaş', 'Patara', 'Antalya'],
      address: { '@type': 'PostalAddress', addressLocality: 'Kalkan', addressRegion: 'Antalya', addressCountry: 'TR' },
    },
    offers: {
      '@type': 'Offer',
      price: String(m.price_gbp),
      priceCurrency: 'GBP',
      description: 'Kişi başı · min. 10 misafir / per person, minimum 10 guests',
      eligibleQuantity: { '@type': 'QuantitativeValue', minValue: 10, unitText: 'guests' },
    },
    hasMenuSection: Object.keys(sections).map(bolum => ({
      '@type': 'MenuSection',
      name: bolum,
      hasMenuItem: sections[bolum].map(d => ({
        '@type': 'MenuItem',
        name: d.ad,
        nutrition: { '@type': 'NutritionInformation', calories: `${d.kcal} kcal` },
      })),
    })),
  };
  return JSON.stringify(ld);
}

let done = 0;
for (const [slug, f] of Object.entries(files)) {
  let html = readFileSync(f, 'utf8');
  const ld = buildLd(slug);
  const block = `<script type="application/ld+json">\n  ${ld}\n  </script>`;
  // İlk ld+json bloğunu (Menu içeren) değiştir
  const re = /<script type="application\/ld\+json">\s*\{[\s\S]*?"@type":"Menu"[\s\S]*?<\/script>/;
  if (re.test(html)) {
    html = html.replace(re, block);
    writeFileSync(f, html);
    done++;
    console.log('OK  ', f, '(' + data.menus[slug].dishes.length + ' item, kcal dahil)');
  } else {
    console.log('XX  ', f, 'Menu ld+json bulunamadı');
  }
}
console.log('Toplam:', done);
