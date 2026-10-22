/**
 * Screenshots the pipeline figure (Fig. 1) from report.html at high resolution,
 * for embedding into the Word (.docx) version of the report.
 */
const { chromium } = require('playwright');
const path = require('path');

const DIR = __dirname;

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ deviceScaleFactor: 4 });
  await page.goto('file://' + path.join(DIR, 'report.html'), { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);

  // Screenshot only the diagram (.flow); the surrounding <figure> also holds the
  // figcaption, which is typeset separately in the .docx build.
  const flow = page.locator('figure .flow').first();
  await flow.screenshot({ path: path.join(DIR, 'fig1_pipeline.png') });

  await browser.close();
  console.log('Wrote fig1_pipeline.png');
})();
