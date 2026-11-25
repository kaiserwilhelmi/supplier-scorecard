// Drives the real page in headless Chromium for tests/roundtrip.py and prints JSON.
//   node tests/browser_upload.js download <site_dir> <out_dir>
//     -> downloads the default template and a custom-weights template (fill rate off), checks the weights editor
//   node tests/browser_upload.js upload <site_dir> <file>...
// The page reads .xlsx itself (no library), so any request that leaves the test server is blocked and reported as an error.
//     -> resets weights to defaults, uploads each file and reports the message, weights and scores shown
const { chromium } = require("playwright");
const http = require("http"), fs = require("fs"), path = require("path");
const [mode, site, ...rest] = process.argv.slice(2);
const CUSTOM = { otd: 40, fill: 0, quality: 30, invoice: 10, price: 20 };

(async () => {
  const srv = http.createServer((q, r) => {
    const f = path.join(site, decodeURIComponent(q.url.split("?")[0]).replace(/\/$/, "/index.html"));
    fs.readFile(f, (e, b) => { if (e) { r.writeHead(404); r.end(); } else r.end(b); });
  }).listen(0);
  const browser = await chromium.launch();
  const page = await browser.newPage({ acceptDownloads: true });
  const errors = [];
  page.on("pageerror", e => errors.push(String(e)));
  await page.route(u => !u.href.startsWith("http://localhost:"), r => { errors.push("outside request: " + r.request().url()); r.abort(); });
  await page.goto(`http://localhost:${srv.address().port}/`);
  const weights = () => page.$$eval(".w-num", els => Object.fromEntries(els.map(e => [e.dataset.k, +e.value])));
  const status = () => page.$eval("#w-status", e => e.textContent);
  const scores = () => page.$$eval("#table tbody tr[data-n]", trs => Object.fromEntries(trs.map(tr => [tr.dataset.n, +tr.querySelector("b").textContent])));
  const headers = () => page.$$eval("#table thead th", ths => ths.map(t => t.textContent));
  const setW = async (k, v) => { await page.fill(`.w-num[data-k="${k}"]`, String(v)); await page.dispatchEvent(`.w-num[data-k="${k}"]`, "change"); };
  const save = async (btn, file) => { const [d] = await Promise.all([page.waitForEvent("download"), page.click(btn)]); await d.saveAs(file); return d.suggestedFilename(); };
  const out = { errors };

  if (mode === "download") {
    const dir = rest[0];
    await page.click("#use-own");
    out.detailsOpen = await page.$eval("#own-details", d => d.open);
    out.defaultWeights = await weights();
    out.defaultStatus = await status();
    out.names = [await save("#dl-xlsx", path.join(dir, "default.xlsx")), await save("#dl-csv", path.join(dir, "default.csv"))];
    const before = await scores();
    // turning a KPI off rebalances the others to 100
    await page.uncheck('.w-on[data-k="fill"]');
    out.afterUncheck = await weights();
    out.afterUncheckStatus = await status();
    const validScores = await scores();
    // an invalid total is flagged and not applied
    await setW("otd", 90);
    out.invalidStatus = await status();
    out.invalidFixVisible = await page.$eval("#w-fix", b => !b.hidden);
    const mid = await scores();
    await page.click("#w-fix");
    out.afterFix = await weights();
    // now set the custom weights exactly
    for (const [k, v] of Object.entries(CUSTOM)) if (v) await setW(k, v);
    out.customWeights = await weights();
    out.customStatus = await status();
    out.customHeaders = await headers();
    out.customDlNote = await page.$eval("#dl-weights", e => e.textContent);
    out.customScoresSample = await scores();
    out.defaultScoresSample = before;
    out.invalidScoresUnchanged = JSON.stringify(mid) === JSON.stringify(validScores);
    await save("#dl-xlsx", path.join(dir, "custom.xlsx"));
    await save("#dl-csv", path.join(dir, "custom.csv"));
    // weights survive a reload (stored per browser)
    await page.reload();
    out.afterReload = await weights();
    await page.click("#w-reset");
    out.afterReset = await weights();
  } else {
    out.results = [];
    for (const f of rest) {
      await page.click("#use-own");
      if (!(await page.$eval("#w-reset", b => b.disabled))) await page.click("#w-reset");
      await page.click("#reset");
      await page.setInputFiles("#file", f);
      await page.waitForFunction(() => !/^Reading/.test(document.getElementById("csv-msg").textContent));
      const msg = await page.$eval("#csv-msg", m => ({ cls: m.className, text: m.innerText }));
      const r = { file: path.basename(f), msg, weights: await weights(), badge: await page.$eval(".badge-sample", b => b.textContent) };
      if (msg.cls.includes("ok")) {
        await page.click('#period button[data-m="3"]'); r.m3 = await scores();
        await page.click('#period button[data-m="6"]'); r.m6 = await scores();
        r.headers = await headers();
      }
      out.results.push(r);
    }
  }
  console.log(JSON.stringify(out));
  await browser.close(); srv.close();
})();
