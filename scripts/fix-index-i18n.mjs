import { readFileSync, writeFileSync } from 'fs';
let s = readFileSync('index.html', 'utf8');
const fixes = [
  // TR bloğu: hero_title_1 eksik → nav_about (TR) sonrası ekle
  ["nav_about:'Hikayemiz',", "nav_about:'Hikayemiz', hero_title_1:'Akdeniz\\'in en taze',"],
  // EN bloğu: hero_title_1 + rsv_title + f_menu_0 + f_menu_8 eksik → nav_about (EN) sonrası ekle
  ["nav_about:'Story',", "nav_about:'Story', hero_title_1:'The freshest Mediterranean', rsv_title:'Let\\'s plan your table.', f_menu_0:'I haven\\'t decided yet', f_menu_8:'Chef\\'s Choice',"],
];
let n = 0;
for (const [a, b] of fixes) { if (s.includes(a)) { s = s.replace(a, b); n++; } else console.log('bulunamadı:', a); }
writeFileSync('index.html', s);
console.log('index i18n fix:', n);
