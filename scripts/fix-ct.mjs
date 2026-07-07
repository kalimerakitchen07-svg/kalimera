import { readFileSync, writeFileSync } from 'fs';
let s = readFileSync('menus.html', 'utf8');
// JS I18N değerlerindeki kaçışsız apostrofları düzelt (yalnızca ct_X:'...' JS bağlamı; HTML data-i18n metnine dokunma)
const fixes = [
  ["ct_kicker:'◆ Şefin İmzası · Chef's Signature'", "ct_kicker:'◆ Şefin İmzası · Chef\\'s Signature'"],
  ["ct_kicker:'◆ Chef's Signature'", "ct_kicker:'◆ Chef\\'s Signature'"],
  ["ct_desc:'Nurdan'ın imza tadım menüsü", "ct_desc:'Nurdan\\'ın imza tadım menüsü"],
  ["ct_desc:'Nurdan's signature tasting menu", "ct_desc:'Nurdan\\'s signature tasting menu"],
];
let n = 0;
for (const [a, b] of fixes) { if (s.includes(a)) { s = s.replace(a, b); n++; } else console.log('bulunamadı:', a.slice(0, 40)); }
writeFileSync('menus.html', s);
console.log('düzeltilen:', n);
