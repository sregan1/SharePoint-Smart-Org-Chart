'use strict';

const puppeteer    = require('puppeteer-core');
const webpack      = require('webpack');
const http         = require('http');
const fs           = require('fs');
const path         = require('path');

const CHROME       = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const PORT         = 8999;
const BASE_URL     = `http://localhost:${PORT}`;
const SHOTS_DIR    = path.resolve(__dirname, '../docs/screenshots');
const DIST_DIR     = path.resolve(__dirname, 'dist');
const webpackConfig = require('./webpack.demo.config.js');

const VIEWPORT = { width: 1280, height: 950 };

// ── Helpers ──────────────────────────────────────────────────────────────────

function build() {
  return new Promise((resolve, reject) => {
    console.log('Building demo bundle…');
    webpack(webpackConfig, (err, stats) => {
      if (err) return reject(err);
      const info = stats.toJson();
      if (stats.hasErrors()) return reject(new Error(info.errors.map(e => e.message || e).join('\n')));
      console.log('Build complete.');
      resolve();
    });
  });
}

function startServer() {
  return new Promise(resolve => {
    const server = http.createServer((req, res) => {
      let filePath = path.join(DIST_DIR, req.url.split('?')[0]);
      // Serve index.html for directories and missing files (SPA fallback)
      if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        filePath = path.join(DIST_DIR, 'index.html');
      }
      const ext  = path.extname(filePath);
      const mime = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css' }[ext] || 'text/plain';
      res.writeHead(200, { 'Content-Type': mime });
      fs.createReadStream(filePath).pipe(res);
    });
    server.listen(PORT, () => { console.log(`Server at ${BASE_URL}`); resolve(server); });
  });
}

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function screenshot(page, filename, clip) {
  const outPath = path.join(SHOTS_DIR, filename);
  const opts = { path: outPath, type: 'png' };
  if (clip) opts.clip = clip;
  await page.screenshot(opts);
  console.log(`  ✓ ${filename}`);
}

async function goAndWait(page, url, extraMs = 2500) {
  await page.setViewport(VIEWPORT);
  await page.goto(`${BASE_URL}${url}`, { waitUntil: 'networkidle0' });
  // Wait for any loading spinner to disappear, then a bit more for animations
  await page.waitForFunction(
    () => !document.querySelector('[class*="ms-Spinner"]') &&
          !document.querySelector('[class*="isLoading"]'),
    { timeout: 15000 }
  ).catch(() => { /* spinner never appeared — that's fine */ });
  await wait(extraMs);
}

async function clickAndWait(page, selector, ms = 600) {
  await page.click(selector);
  await wait(ms);
}

// The tree canvas centers the whole hierarchy horizontally, so a CEO-rooted
// tree that's wider than the viewport scrolls the root out of view unless we
// explicitly re-center on it (drill mode and narrower subtrees don't need this).
async function centerTree(page) {
  await page.evaluate(() => {
    const el = document.querySelector('[class*="treeScroll"]');
    if (!el) return;
    el.scrollTop = 0;
    el.scrollLeft = Math.max(0, (el.scrollWidth - el.clientWidth) / 2);
  });
  await wait(300);
}

// ── Main ─────────────────────────────────────────────────────────────────────

(async () => {
  fs.mkdirSync(SHOTS_DIR, { recursive: true });

  // 1. Build
  await build();

  // 2. Start server
  const server = await startServer();

  // 3. Launch Chrome
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless:       true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--no-first-run',
      '--no-zygote',
    ],
    defaultViewport: VIEWPORT,
  });

  const page = await browser.newPage();
  page.on('console', m => { if (m.type() === 'error') console.error('Page error:', m.text()); });

  try {
    console.log('\nTaking screenshots…');

    // ── 02 Header bar (clip top 60px) ────────────────────────────────────────
    await goAndWait(page, '/?view=directory');
    await screenshot(page, '02-header-bar.png', { x: 0, y: 0, width: VIEWPORT.width, height: 60 });

    // ── 03 Directory overview ────────────────────────────────────────────────
    await screenshot(page, '03-directory-overview.png');

    // ── 04 Alpha filter active ───────────────────────────────────────────────
    // Click the "S" letter button in the alphabet filter
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const s    = btns.find(b => b.textContent.trim() === 'S');
      if (s) s.click();
    });
    await wait(600);
    await screenshot(page, '04-directory-alpha-filter.png');

    // ── 05 Employee card close-up ────────────────────────────────────────────
    const cardBox = await page.evaluate(() => {
      const card = document.querySelector('[class*="employeeCard"]') ||
                   document.querySelector('[class*="card"]');
      if (!card) return null;
      const r = card.getBoundingClientRect();
      return { x: Math.floor(r.left), y: Math.floor(r.top), width: Math.ceil(r.width) + 40, height: Math.ceil(r.height) + 80 };
    });
    if (cardBox) {
      await screenshot(page, '05-directory-card.png', cardBox);
    } else {
      await screenshot(page, '05-directory-card.png'); // fallback: full page
    }

    // ── 06 Org Chart drill-down ──────────────────────────────────────────────
    await goAndWait(page, '/?view=orgchart&layout=drill');
    await screenshot(page, '06-orgchart-drill.png');

    // ── 07 Drill-down 2 levels deep, with breadcrumb ─────────────────────────
    // Click the first report's expand button twice: CEO -> VP -> Director,
    // so the breadcrumb trail actually has something to show.
    for (let i = 0; i < 2; i++) {
      await page.evaluate(() => {
        const btn = document.querySelector('[class*="expandBtn"]') ||
                    document.querySelector('[class*="nodeCard"]');
        if (btn) btn.click();
      });
      await wait(1000);
    }
    await screenshot(page, '07-orgchart-drill-breadcrumb.png');

    // ── 08 Org Chart vertical (Top Down) tree ────────────────────────────────
    // Fixed zoom + centered scroll — auto-fit shrinks a CEO-rooted 150-person
    // tree down to an unreadable 25%, and the canvas centers on the full
    // hierarchy width, not the viewport, so the root scrolls out of view.
    await goAndWait(page, '/?view=orgchart&layout=vertical&levels=3&zoom=0.65', 3000);
    await centerTree(page);
    await screenshot(page, '08-orgchart-vertical.png');
    // Also the main "overview" hero shot — the tree reads better than the
    // plain directory grid (03-directory-overview.png already covers that).
    fs.copyFileSync(
      path.join(SHOTS_DIR, '08-orgchart-vertical.png'),
      path.join(SHOTS_DIR, '01-overview.png')
    );
    console.log('  ✓ 01-overview.png (copy of 08-orgchart-vertical.png)');

    // ── 09 Toolbar close-up ──────────────────────────────────────────────────
    const toolbarBox = await page.evaluate(() => {
      const tb = document.querySelector('[class*="chartToolbar"]') ||
                 document.querySelector('[class*="toolbar"]');
      if (!tb) return null;
      const r = tb.getBoundingClientRect();
      return { x: 0, y: Math.floor(r.top), width: 1280, height: Math.ceil(r.height) + 4 };
    });
    await screenshot(page, '09-orgchart-toolbar.png', toolbarBox || undefined);

    // ── 10 Stats bar open ────────────────────────────────────────────────────
    // Click the stats toggle button (BarChartVertical icon button)
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button[class*="iconToolBtn"], button[class*="chartActionBtn"]'));
      const stats = btns.find(b => b.title && b.title.toLowerCase().includes('stat'));
      if (stats) stats.click();
    });
    await wait(500);
    await screenshot(page, '10-orgchart-stats.png');

    // ── 11 Department filter open ────────────────────────────────────────────
    // Close stats first so the two panels don't stack in the same shot
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const stats = btns.find(b => b.title && b.title.toLowerCase().includes('stat'));
      if (stats) stats.click();
    });
    await wait(300);
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const dept = btns.find(b => b.title && (b.title.toLowerCase().includes('department') || b.title.toLowerCase().includes('dept')));
      if (dept) dept.click();
    });
    await wait(500);
    await screenshot(page, '11-orgchart-deptfilter.png');

    // ── 12 Person profile card popup ─────────────────────────────────────────
    // Close any open dropdowns first
    await page.keyboard.press('Escape');
    await wait(200);
    await goAndWait(page, '/?view=orgchart&layout=drill');
    // Drill CEO -> VP -> Director so the profile we open has a multi-level
    // "Reports to" chain to show off, then click one of the Director's own
    // reports (a manager) to open their card.
    for (let i = 0; i < 2; i++) {
      await page.evaluate(() => {
        const btn = document.querySelector('[class*="expandBtn"]') ||
                    document.querySelector('[class*="nodeCard"]');
        if (btn) btn.click();
      });
      await wait(1000);
    }
    await page.evaluate(() => {
      const cards = Array.from(document.querySelectorAll('[class*="nodeCard"]'));
      // Cards double as the "drill into" control — click one's avatar/name
      // area isn't separable from drill-in, so instead open the profile via
      // the currently-drilled header, which shows richer chain 2 levels in.
      const profileBtn = document.querySelector('[class*="drillCurrentProfileBtn"]');
      if (profileBtn) { profileBtn.click(); return; }
      if (cards[0]) cards[0].click();
    });
    await wait(800);
    await screenshot(page, '12-person-card.png');

    // ── 13 User Preferences panel ────────────────────────────────────────────
    await goAndWait(page, '/?view=directory');
    // Click the Settings gear (last actionBtn in header)
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('[class*="actionBtn"]'));
      const gear = btns.find(b => b.title && b.title.toLowerCase().includes('pref'));
      if (gear) gear.click();
      else if (btns.length) btns[btns.length - 1].click();
    });
    await wait(800);
    await screenshot(page, '13-preferences-panel.png');

    // ── 14 Property pane — SharePoint only; write a note file instead ─────────
    fs.writeFileSync(
      path.join(SHOTS_DIR, '14-property-pane.txt'),
      'This screenshot must be taken from SharePoint Edit mode (property pane is SharePoint-only).\n' +
      'Capture (or re-capture) the Visual Style group with Chart Theme = Custom, showing the accent-color\n' +
      'picker (swatch, hex, RGB, contrast indicator), and the Custom Attributes group with a couple of\n' +
      'example rows configured.\n'
    );
    console.log('  ✓ 14-property-pane.txt (placeholder note)');

    // ── 15 Theme comparison — five separate shots ─────────────────────────────
    for (const theme of ['modern', 'minimal', 'corporate', 'dark', 'custom']) {
      await goAndWait(page, `/?view=orgchart&layout=drill&theme=${theme}`, 1500);
      await screenshot(page, `15-theme-${theme}.png`);
    }
    // Also save a combined strip if possible (just use modern as the primary 15 shot)
    fs.copyFileSync(
      path.join(SHOTS_DIR, '15-theme-modern.png'),
      path.join(SHOTS_DIR, '15-themes.png')
    );
    console.log('  ✓ 15-themes.png (also saved 15-theme-modern/minimal/corporate/dark.png)');

    // ── 17 Directory list view ────────────────────────────────────────────────
    await goAndWait(page, '/?view=directory');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const listBtn = btns.find(b => b.title && b.title.toLowerCase().includes('list'));
      if (listBtn) listBtn.click();
    });
    await wait(600);
    await screenshot(page, '17-directory-list.png');

    // ── 18 Org Chart horizontal (Left to Right) tree ─────────────────────────
    await goAndWait(page, '/?view=orgchart&layout=horizontal&levels=3&zoom=0.75', 3000);
    await centerTree(page); // no-op for this layout (already left-anchored) but harmless
    await screenshot(page, '18-orgchart-horizontal.png');

    console.log(`\nAll screenshots saved to ${SHOTS_DIR}`);

  } catch (err) {
    console.error('\nScreenshot error:', err.message);
    // Save a debug screenshot to help diagnose
    await page.screenshot({ path: path.join(SHOTS_DIR, '_debug.png') }).catch(() => {});
    process.exitCode = 1;
  } finally {
    await browser.close();
    server.close();
  }
})();
