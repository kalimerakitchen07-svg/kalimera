// i18n denetimi: her sayfada kullanılan data-i18n anahtarları TR/EN/RU bloklarında var mı?
// Eksik anahtar → o dilde Türkçe (fallback) görünür = çeviri hatası.
import { readFileSync } from 'fs';
import { readdirSync } from 'fs';

const files = ['index.html', 'menus.html', 'nurdan.html',
  'menu/turkish-night.html', 'menu/mediterranean-night.html', 'menu/bbq-menu.html',
  'menu/chicken-bbq.html', 'menu/fish-bbq.html', 'menu/breakfast.html', 'menu/build.html'];

function extractBlocks(html) {
  // I18N / TEXTS / translations objelerindeki tr/en/ru bloklarını topla
  const res = { tr: new Set(), en: new Set(), ru: new Set() };
  for (const lang of ['tr', 'en', 'ru']) {
    let from = 0;
    while (true) {
      const m = html.slice(from).match(new RegExp('(?:^|[\\s,{])' + lang + '\\s*:\\s*\\{'));
      if (!m) break;
      const start = from + m.index + m[0].indexOf('{');
      let i = start, depth = 0, end = -1;
      for (; i < html.length; i++) { if (html[i] === '{') depth++; else if (html[i] === '}') { depth--; if (depth === 0) { end = i; break; } } }
      if (end < 0) break;
      const block = html.slice(start + 1, end);
      const re = /(\w+):\s*'/g; let k;
      while ((k = re.exec(block)) !== null) res[lang].add(k[1]);
      from = end + 1;
    }
  }
  return res;
}

let totalIssues = 0;
for (const f of files) {
  let html;
  try { html = readFileSync(f, 'utf8'); } catch (e) { continue; }
  const used = [...new Set([...html.matchAll(/data-i18n="([^"]+)"/g)].map(m => m[1]))];
  const blocks = extractBlocks(html);
  if (blocks.tr.size === 0) { console.log('⚠ ', f, '— i18n bloğu bulunamadı'); continue; }
  const missEn = used.filter(k => blocks.tr.has(k) && !blocks.en.has(k));
  const missRu = used.filter(k => blocks.tr.has(k) && !blocks.ru.has(k));
  const missTr = used.filter(k => !blocks.tr.has(k));
  if (missEn.length || missRu.length || missTr.length) {
    console.log('\n■ ' + f + '  (kullanılan:' + used.length + ' key)');
    if (missTr.length) console.log('   TR eksik:', missTr.join(', '));
    if (missEn.length) console.log('   EN eksik:', missEn.join(', '));
    if (missRu.length) console.log('   RU eksik:', missRu.join(', '));
    totalIssues += missEn.length + missRu.length + missTr.length;
  } else {
    console.log('✓ ' + f + '  (' + used.length + ' key, TR/EN/RU tam)');
  }
}
console.log('\nTOPLAM EKSİK ÇEVİRİ:', totalIssues);
