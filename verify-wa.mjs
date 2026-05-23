import puppeteer from 'puppeteer';

const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
const page = await browser.newPage();

const cases = [
  ['/menu/turkish-night.html', 'tr'],
  ['/menu/turkish-night.html', 'en'],
  ['/menu/turkish-night.html', 'ru'],
  ['/menu/breakfast.html', 'en'],
  ['/menu/fish-bbq.html', 'ru'],
];

for (const [path, lang] of cases) {
  await page.goto('http://localhost:3010' + path, { waitUntil: 'domcontentloaded' });
  await page.evaluate(l => localStorage.setItem('kkLang', l), lang);
  await page.reload({ waitUntil: 'domcontentloaded' });
  const result = await page.evaluate(() => {
    const a = document.getElementById('menu-reserve-wa');
    const lab = document.getElementById('menu-reserve-label');
    return { href: a?.href, label: lab?.textContent };
  });
  console.log(`[${lang}] ${path}\n  label: ${result.label}\n  href: ${decodeURIComponent(result.href)}\n`);
}
await browser.close();
