import puppeteer from 'puppeteer';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const file = 'file:///' + join(__dirname, 'og-preview.html').replace(/\\/g, '/');

const browser = await puppeteer.launch({ headless: true });
const page = await browser.newPage();
await page.setViewport({ width: 1200, height: 630, deviceScaleFactor: 1 });
await page.goto(file, { waitUntil: 'networkidle0' });
await page.screenshot({ path: join(__dirname, 'og2.png'), clip: { x:0, y:0, width:1200, height:630 } });
await browser.close();
console.log('Done: og2.png (1200x630)');
