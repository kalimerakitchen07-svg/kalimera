// Agent'ların ürettiği Türkçe duyusal açıklamaları menü sayfalarının SADECE TR I18N bloğuna uygular.
import { readFileSync, writeFileSync } from 'fs';

const SP = 'C:/Users/socie/AppData/Local/Temp/claude/C--Users-socie/97a8d4e1-b13f-4528-8f64-0931d0249efe/scratchpad';
const map = {
  'turkish-night': 'desc-turkish-night.json',
  'mediterranean-night': 'desc-mediterranean-night.json',
  'bbq-menu': 'desc-bbq-menu.json',
  'chicken-bbq': 'desc-chicken-bbq.json',
  'fish-bbq': 'desc-fish-bbq.json',
  'breakfast': 'desc-breakfast.json',
};
const esc = s => s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
let total = 0;
for (const [slug, jf] of Object.entries(map)) {
  const file = 'menu/' + slug + '.html';
  let html = readFileSync(file, 'utf8');
  const j = JSON.parse(readFileSync(SP + '/' + jf, 'utf8'));
  let applied = 0;
  for (const [k, v] of Object.entries(j)) {
    // İlk occurrence = TR bloğu (obje sırası: tr, en, ru)
    const re = new RegExp(k + ":\\s*'(?:[^'\\\\]|\\\\.)*'");
    if (re.test(html)) {
      html = html.replace(re, k + ":'" + esc(v) + "'");
      applied++; total++;
    } else {
      console.log('  bulunamadı:', slug, k);
    }
  }
  writeFileSync(file, html);
  console.log(slug.padEnd(20), applied + ' TR açıklama güncellendi');
}
console.log('TOPLAM:', total);
