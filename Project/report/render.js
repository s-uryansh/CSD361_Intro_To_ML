/**
 * Renders report.html to a print-ready, IEEE-style two-column PDF using Playwright/Chromium.
 * Usage: node render.js
 */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const DIR = __dirname;
const HTML = path.join(DIR, 'report.html');
const PDF = path.join(DIR, 'CSD361_Preliminary_Report_StructX.pdf');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();

  await page.goto('file://' + HTML, { waitUntil: 'networkidle' });
  await page.emulateMedia({ media: 'print' });
  await page.evaluate(() => document.fonts.ready);

  await page.pdf({
    path: PDF,
    format: 'Letter',
    printBackground: true,
    preferCSSPageSize: true,
    displayHeaderFooter: true,
    headerTemplate: '<div></div>',
    footerTemplate: `
      <div style="width:100%;font-family:'Times New Roman',Times,serif;font-size:7.5pt;color:#000;
                  padding:0 0.60in;display:flex;justify-content:space-between;">
        <span>CSD361 Preliminary Project Report &middot; Struct-X</span>
        <span>Page <span class="pageNumber"></span> of <span class="totalPages"></span></span>
      </div>`,
    margin: { top: '0.66in', right: '0.60in', bottom: '0.70in', left: '0.60in' },
  });

  await browser.close();

  const kb = (fs.statSync(PDF).size / 1024).toFixed(1);
  console.log(`Wrote ${PDF} (${kb} KB)`);
})();
