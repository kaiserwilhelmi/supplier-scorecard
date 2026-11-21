// Upload each fixture through the real page in headless Chromium and report what the page shows.
// Usage: node tests/browser_upload.js <site_dir> <xlsx.full.min.js> <fixture>...   -> JSON on stdout
const { chromium } = require("playwright");
const http = require("http"), fs = require("fs"), path = require("path");
const [site, xlsxLib, ...fixtures] = process.argv.slice(2);
(async () => {
  const srv = http.createServer((q, r) => {
    const f = path.join(site, decodeURIComponent(q.url.split("?")[0]).replace(/\/$/, "/index.html"));
    fs.readFile(f, (e, b) => { if (e) { r.writeHead(404); r.end(); } else { r.end(b); } });
  }).listen(0);
  const base = `http://localhost:${srv.address().port}/`;
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", e => errors.push(String(e)));
  await page.route("https://cdnjs.cloudflare.com/**", r => r.fulfill({ body: fs.readFileSync(xlsxLib), contentType: "text/javascript" }));
  await page.goto(base);
  const out = { downloads: {} };
  for (const href of ["template/supplier-scorecard-template.xlsx", "template/supplier-scorecard-template.csv"]) {
    const res = await page.request.get(base + href); out.downloads[href] = res.status();
  }
  await page.click("#use-own");
  out.detailsOpen = await page.$eval("#own-details", d => d.open);
  const read = async (months) => {
    if (months) await page.click(`#period button[data-m="${months}"]`);
    return page.$$eval("#table tbody tr[data-n]", trs => trs.map(tr => ({ name: tr.dataset.n, cells: [...tr.children].map(td => td.textContent.trim()) })));
  };
  out.results = [];
  for (const f of fixtures) {
    await page.click("#reset");
    await page.setInputFiles("#file", f);
    await page.waitForFunction(() => !/^Reading/.test(document.getElementById("csv-msg").textContent));
    const msg = await page.$eval("#csv-msg", m => ({ cls: m.className, text: m.innerText }));
    const r = { file: path.basename(f), msg, badge: await page.$eval(".badge-sample", b => b.textContent) };
    if (msg.cls.includes("ok")) { r.m3 = await read(3); r.m6 = await read(6); r.months = await page.$eval("#range-note", e => e.textContent); }
    out.results.push(r);
  }
  out.pageErrors = errors;
  console.log(JSON.stringify(out));
  await browser.close(); srv.close();
})();
