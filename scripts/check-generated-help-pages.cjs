const { chromium } = require('playwright');

const baseUrl = process.argv[2] || 'http://127.0.0.1:18080/de/help/views/';
const paths = ['index.html', 'article-edit.html', 'partner-edit.html', 'finance-payment-edit.html'];

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1200 } });

  async function check(path) {
    await page.goto(new URL(path, baseUrl).toString(), { waitUntil: 'networkidle' });
    return page.evaluate(() => {
      const img = document.querySelector('.screenshot img, img');
      return {
        title: document.title,
        h1: document.querySelector('h1')?.textContent?.trim() || '',
        metaDescriptionLength: document.querySelector('meta[name="description"]')?.content.length || 0,
        structuredData: Boolean(document.querySelector('script[type="application/ld+json"]')),
        image: img ? {
          complete: img.complete,
          width: img.naturalWidth,
          height: img.naturalHeight,
          alt: img.getAttribute('alt') || '',
        } : null,
        links: document.querySelectorAll('a').length,
        rows: document.querySelectorAll('tr, .field-row, .button-row').length,
        textSample: document.body.innerText.slice(0, 160),
      };
    });
  }

  const results = {};
  for (const path of paths) {
    results[path] = await check(path);
  }

  await browser.close();
  console.log(JSON.stringify(results, null, 2));
})().catch(error => {
  console.error(error.stack || error.message);
  process.exit(1);
});
