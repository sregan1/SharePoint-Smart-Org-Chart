/* Run after building demo/webpack.demo.config.js: node tests/tree-navigation.cjs */
const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const http = require('http');
const puppeteer = require('puppeteer-core');
const ts = require('typescript');

const root = path.resolve(__dirname, '../demo/dist');
const canvas = '[class*="__treeCanvas"]';
const content = '[class*="__canvasContent"]';
const zoom = 'input[aria-label="Zoom percentage (1–150)"]';
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function camera(page) {
  return page.$eval(content, el => {
    const m = new DOMMatrix(getComputedStyle(el).transform);
    return { x: m.e, y: m.f, zoom: m.a };
  });
}

async function setZoom(page, value) {
  await page.click(zoom, { clickCount: 3 });
  await page.keyboard.type(String(value));
  await page.keyboard.press('Enter');
  await page.waitForFunction((s, z) => Number(document.querySelector(s).value) === z, {}, zoom, value);
  await delay(60);
}

async function swipe(client, x, y, dx, dy) {
  await client.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y }] });
  for (let i = 1; i <= 12; i++) {
    await client.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: x + dx * i / 12, y: y + dy * i / 12 }] });
    await delay(20);
  }
  await client.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await delay(150);
}

(async () => {
  const server = http.createServer((req, res) => {
    const file = path.join(root, decodeURIComponent(req.url.split('?')[0] === '/' ? '/index.html' : req.url.split('?')[0]));
    if (!file.startsWith(root + path.sep) || !fs.existsSync(file)) { res.writeHead(404).end(); return; }
    res.setHeader('Content-Type', file.endsWith('.js') ? 'application/javascript' : file.endsWith('.html') ? 'text/html' : 'application/octet-stream');
    fs.createReadStream(file).pipe(res);
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const browser = await puppeteer.launch({ headless: true,
    executablePath: process.env.CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' });
  const errors = [];
  try {
    const page = await browser.newPage();
    page.on('pageerror', e => errors.push(e.message));
    await page.setViewport({ width: 1280, height: 950 });
    const base = `http://127.0.0.1:${server.address().port}/?view=orgchart`;
    for (const layout of ['vertical', 'horizontal']) {
      for (const size of [150, 500, 1000]) {
        await page.goto(`${base}&layout=${layout}&mockSize=${size}&levels=5`, { waitUntil: 'networkidle0' });
        await page.waitForSelector(`${content} [data-soc-id]`);
        await delay(250);
        const geometry = await page.$eval(content, el => {
          const a = el.getBoundingClientRect(), b = el.parentElement.getBoundingClientRect();
          return { fits: a.left >= b.left - 1 && a.right <= b.right + 1 && a.top >= b.top - 1 && a.bottom <= b.bottom + 1,
            zoom: new DOMMatrix(getComputedStyle(el).transform).a,
            overflow: getComputedStyle(el.parentElement).overflow, pageOverflow: document.documentElement.scrollWidth > window.innerWidth };
        });
        assert(geometry.fits || Math.abs(geometry.zoom - 0.01) < 0.00001, `${size} ${layout}: fit or minimum 1% zoom`);
        assert.equal(geometry.overflow, 'hidden');
        assert.equal(geometry.pageOverflow, false);
        console.log(`PASS ${size} people / ${layout}: fit and no page overflow`);
      }
    }
    await page.goto(`${base}&layout=vertical&zoom=0.17`, { waitUntil: 'networkidle0' });
    await page.waitForSelector(zoom);
    assert.equal(Number(await page.$eval(zoom, el => el.value)), 17);
    // Exercise the real property-pane renderer without a SharePoint tenant.
    const webPart = ts.transpileModule(fs.readFileSync(path.resolve(__dirname,
      '../src/webparts/smartOrgChart/SmartOrgChartWebPart.ts'), 'utf8'), {
      compilerOptions: { target: ts.ScriptTarget.ES2017, module: ts.ModuleKind.CommonJS }
    }).outputText;
    const propertyResult = await page.evaluate(source => {
      const exported = {};
      const stubRequire = name => name === '@microsoft/sp-webpart-base'
        ? { BaseClientSideWebPart: class {} } : {};
      new Function('require', 'exports', source)(stubRequire, exported);
      const part = new exported.default();
      part.context = { instanceId: 'test' };
      part.properties = { defaultZoom: 0.5 };
      const field = document.createElement('div');
      const changes = [];
      const change = (key, value) => {
        changes.push([part.properties[key], value]);
        part.properties[key] = value;
      };
      part._renderZoomField(field, change);
      const input = field.querySelector('input[type=number]');
      const auto = field.querySelector('input[type=checkbox]');
      const legacy = [0.5, 0.75, 1, 1.25, 1.5].map(value => {
        part.properties.defaultZoom = value;
        part._renderZoomField(field, change);
        return Number(input.value);
      });
      for (const value of ['1', '7', '43', '100']) {
        input.value = value;
        input.dispatchEvent(new Event('input'));
      }
      const count = changes.length;
      for (const value of ['', '0', '101', '12.5']) {
        input.value = value;
        input.dispatchEvent(new Event('input'));
      }
      input.dispatchEvent(new Event('blur'));
      const invalidIgnored = count === changes.length && input.value === '100';
      auto.checked = true;
      auto.dispatchEvent(new Event('change'));
      const autoWorks = part.properties.defaultZoom === 0 && input.disabled;
      auto.checked = false;
      auto.dispatchEvent(new Event('change'));
      return { legacy, values: changes.slice(0, 4).map(pair => pair[1]),
        oldValue: changes[0][0], invalidIgnored, autoWorks,
        fixedWorks: part.properties.defaultZoom === 1 && !input.disabled,
        sameInput: input === field.querySelector('input[type=number]') };
    }, webPart);
    assert.deepEqual(propertyResult, { legacy: [50, 75, 100, 125, 150], values: [0.01, 0.07, 0.43, 1],
      oldValue: 1.5, invalidIgnored: true, autoWorks: true, fixedWorks: true, sameInput: true });
    console.log('PASS property pane: arbitrary defaults, validation, auto mode and legacy settings');
    for (const value of [1, 7, 43, 100, 150]) {
      await setZoom(page, value);
      assert(Math.abs((await camera(page)).zoom - value / 100) < 0.00001);
    }
    await setZoom(page, 43);
    const before = await camera(page);
    const box = await page.$eval(canvas, el => { const r = el.getBoundingClientRect(); return { x: r.left, y: r.top, width: r.width, height: r.height }; });
    const ax = 120, ay = 130;
    await page.mouse.move(box.x + ax, box.y + ay);
    await page.mouse.wheel({ deltaY: -80 });
    await delay(150);
    const after = await camera(page);
    assert(after.zoom > before.zoom);
    assert(Math.abs((ax - before.x) / before.zoom - (ax - after.x) / after.zoom) < 0.05);
    assert(Math.abs((ay - before.y) / before.zoom - (ay - after.y) / after.zoom) < 0.05);
    await page.mouse.move(box.x + 30, box.y + 40);
    await page.mouse.down();
    await page.mouse.move(box.x + 130, box.y + 100, { steps: 10 });
    await page.mouse.up();
    await delay(100);
    const panned = await camera(page);
    assert(Math.abs(panned.x - after.x - 100) < 1);
    assert(Math.abs(panned.y - after.y - 60) < 1);
    assert.equal(await page.$('[class*="__personCardOverlay"]'), null);
    await page.click('button[aria-label="Reset zoom"]');
    await delay(100);
    assert(Math.abs((await camera(page)).zoom - 0.17) < 0.00001);
    await page.click('button[aria-label="Fit entire chart"]');
    await delay(100);
    await page.focus(`${content} [data-soc-id]`);
    const rootFocus = await page.evaluate(() => document.activeElement.getAttribute('data-soc-id'));
    await page.keyboard.press('ArrowDown');
    await delay(100);
    const focusResult = await page.$eval(canvas, el => {
      const card = document.activeElement;
      const a = card.getBoundingClientRect(), b = el.getBoundingClientRect();
      return { id: card.getAttribute('data-soc-id'), visible: a.left >= b.left && a.right <= b.right && a.top >= b.top && a.bottom <= b.bottom,
        scrollLeft: el.scrollLeft, scrollTop: el.scrollTop };
    });
    assert.notEqual(focusResult.id, rootFocus);
    assert(focusResult.visible);
    assert.equal(focusResult.scrollLeft, 0);
    assert.equal(focusResult.scrollTop, 0);
    await delay(350);
    await page.keyboard.press('Enter');
    await page.waitForSelector('[role="dialog"]');
    await page.keyboard.press('Escape');
    await page.waitForSelector('[role="dialog"]', { hidden: true });
    await page.click('button[aria-label="Fit entire chart"]');
    await delay(100);
    await page.screenshot({ path: path.join(__dirname, '../demo/dist/navigation-desktop.png') });
    console.log('PASS arbitrary zoom, cursor anchor, drag, reset, keyboard focus and profile cards');

    const mobile = await browser.newPage();
    mobile.on('pageerror', e => errors.push(e.message));
    await mobile.setViewport({ width: 390, height: 844, isMobile: true, hasTouch: true });
    await mobile.goto(`${base}&layout=vertical&mockSize=1000`, { waitUntil: 'networkidle0' });
    await mobile.waitForSelector(canvas);
    // Model a SharePoint page with content both above and below the web part.
    await mobile.evaluate(() => {
      document.body.style.paddingTop = '100px';
      document.body.style.paddingBottom = '1000px';
    });
    await delay(150);
    assert.equal(await mobile.evaluate(() => document.documentElement.scrollWidth > window.innerWidth), false);
    const client = await mobile.createCDPSession();
    const mobileBox = await mobile.$eval(canvas, el => { const r = el.getBoundingClientRect(); return { x: r.left, y: r.top, height: r.height }; });
    const mobileBefore = await camera(mobile);
    await swipe(client, 170, Math.min(700, mobileBox.y + mobileBox.height - 40), 0, -100);
    assert(await mobile.evaluate(() => window.scrollY > 30), 'inactive canvas must allow page scrolling');
    assert.deepEqual(await camera(mobile), mobileBefore);
    await mobile.click('button::-p-text(Explore chart)');
    await delay(100);
    const activeBox = await mobile.$eval(canvas, el => { const r = el.getBoundingClientRect(); return { y: r.top, height: r.height }; });
    const pageY = await mobile.evaluate(() => window.scrollY);
    const activeBefore = await camera(mobile);
    const sy = Math.min(700, activeBox.y + activeBox.height - 50);
    await swipe(client, 170, sy, 50, -80);
    assert.equal(await mobile.evaluate(() => window.scrollY), pageY);
    const activeAfter = await camera(mobile);
    assert(Math.abs(activeAfter.x - activeBefore.x - 50) < 2);
    assert(Math.abs(activeAfter.y - activeBefore.y + 80) < 2);
    await client.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: 100, y: sy - 60 }, { x: 200, y: sy - 60 }] });
    await client.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: 70, y: sy - 60 }, { x: 230, y: sy - 60 }] });
    await client.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    await delay(100);
    assert((await camera(mobile)).zoom > activeAfter.zoom * 1.4, 'pinch must zoom the chart');
    await mobile.click('button::-p-text(Done)');
    await delay(100);
    await swipe(client, 170, sy, 0, -100);
    assert(await mobile.evaluate(y => window.scrollY > y + 30, pageY));
    await mobile.screenshot({ path: path.join(__dirname, '../demo/dist/navigation-mobile.png') });
    assert.deepEqual(errors, []);
    console.log('PASS mobile inactive scrolling, active pan/pinch, Done restores scrolling; no runtime errors');
  } finally {
    await browser.close();
    server.close();
  }
})().catch(e => { console.error(e); process.exitCode = 1; });
