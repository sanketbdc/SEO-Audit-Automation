/**
 * ============================================================
 *  COMPLETE SEO AUDIT TOOL  —  Playwright + ExcelJS
 * ============================================================
 *  Pipeline (runs automatically in order):
 *   1. Crawl all URLs on the domain          → Sheet 1
 *   2. Fetch SEO details (with redirect fix) → Sheet 2
 *   3. Collect image alt tags                → Sheet 3
 *   4. Internal / External link audit        → Sheet 4
 *   5. Email the Excel report
 *
 *  Usage:
 *    node seo_audit.js
 *
 *  Config: edit the CONFIG section below before running.
 * ============================================================
 */

const { chromium }  = require("playwright");
const ExcelJS       = require("exceljs");
const nodemailer    = require("nodemailer");
const cron          = require("node-cron");
const path          = require("path");
const https         = require("https");
const http          = require("http");

// ╔══════════════════════════════════════════════════════════╗
// ║                   USER CONFIG                           ║
// ╚══════════════════════════════════════════════════════════╝
const CONFIG = {

  // ── Websites to audit (add as many as you want) ──────────
  // bypassHttpCheck: true  → use for Cloudflare/security protected sites (slower)
  // bypassHttpCheck: false → default, fast raw HTTP check
  websites: [
    { url: "https://www.delhiredz.com/",             bypassHttpCheck: false },
    // { url: "https://www.bharat-connect.com/",      bypassHttpCheck: true  },
    // { url: "https://www.shapoorjipallonji.com/",   bypassHttpCheck: false },
    // { url: "https://shapoorjirealestate.com/",     bypassHttpCheck: false },
    // { url: "https://www.joyvillehomes.com/",       bypassHttpCheck: false },
    // { url: "https://www.viceroyproperties.in/",   bypassHttpCheck: false },
    // { url: "https://bombaydc.com/",               bypassHttpCheck: false },
  ].filter(s => !process.env.AUDIT_SITE || s.url === process.env.AUDIT_SITE),

  // ── Output Excel file path ────────────────────────────────
  outputFile: "SEO_Audit_Report.xlsx",

  // ── Crawler settings ──────────────────────────────────────
  crawlConcurrency : 5,
  crawlTimeout     : 30_000,

  // ── Email settings ────────────────────────────────────────
  email: {
    enabled : true,
    host    : "smtp.gmail.com",
    port    : 587,
    secure  : false,
    user    : process.env.EMAIL_USER,
    pass    : process.env.EMAIL_PASS,
    to      : "sanket@bombaydc.com",
  },

  // ── Scheduler settings ────────────────────────────────────
  // Runs on the 1st of every month at 8:00 AM
  // Cron format: 'minute hour day month weekday'
  // Examples: '0 8 1 * *' = 1st of month at 8AM
  //           '0 8 * * 1' = every Monday at 8AM
   schedule: "0 2 8 5 *",
};
// ╚══════════════════════════════════════════════════════════╝


// ── Helpers ───────────────────────────────────────────────────────────────────

const MEDIA_EXT = new Set([
  ".jpg",".jpeg",".png",".gif",".svg",".webp",".ico",".bmp",
  ".mp4",".avi",".mov",".mp3",".wav",".pdf",
  ".zip",".tar",".gz",".exe",".css",".js",".woff",".woff2",".ttf",
]);

function normalizeUrl(raw) {
  try {
    const u = new URL(raw);
    u.hash = "";
    if ((u.protocol === "https:" && u.port === "443") ||
        (u.protocol === "http:"  && u.port === "80")) u.port = "";
    let href = u.href;
    if (href.endsWith("/") && href !== u.origin + "/") href = href.slice(0, -1);
    return href;
  } catch { return null; }
}

function shouldCrawl(urlStr, baseDomain) {
  try {
    const u = new URL(urlStr);
    if (!["http:","https:"].includes(u.protocol)) return false;
    if (u.hostname !== baseDomain && !u.hostname.endsWith("." + baseDomain)) return false;
    if (MEDIA_EXT.has(path.extname(u.pathname).toLowerCase())) return false;
    return true;
  } catch { return false; }
}

/** HTTP HEAD/GET status + final redirect URL */
function getStatusAndFinalUrl(inputUrl) {
  return new Promise((resolve) => {
    const doRequest = (url, method, redirectCount) => {
      if (redirectCount > 10) return resolve({ status: -1, finalUrl: url, redirected: true });
      try {
        const parsed = new URL(url);
        const lib = parsed.protocol === "https:" ? https : http;
        const req = lib.request(
          { hostname: parsed.hostname, path: parsed.pathname + parsed.search,
            method, headers: { "User-Agent": "Mozilla/5.0 (SEO-Audit-Bot/1.0)" },
            timeout: 15000, rejectUnauthorized: false },
          (res) => {
            const status = res.statusCode;
            if ([301, 302, 303, 307, 308].includes(status) && res.headers.location) {
              const next = new URL(res.headers.location, url).href;
              return doRequest(next, method, redirectCount + 1);
            }
            // Some servers block HEAD → retry with GET
            if ((status === 403 || status === 405) && method === "HEAD") {
              return doRequest(url, "GET", redirectCount);
            }
            resolve({
              status,
              finalUrl    : url,
              redirected  : redirectCount > 0,
              originalUrl : inputUrl,
            });
          }
        );
        req.on("error", () => resolve({ status: -1, finalUrl: url, redirected: false }));
        req.on("timeout", () => { req.destroy(); resolve({ status: -1, finalUrl: url, redirected: false }); });
        req.end();
      } catch { resolve({ status: -1, finalUrl: inputUrl, redirected: false }); }
    };
    doRequest(inputUrl, "HEAD", 0);
  });
}

// ── Excel styling helpers ─────────────────────────────────────────────────────

function styleHeader(row, bgArgb = "FF1F3864") {
  row.eachCell(cell => {
    cell.font      = { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 11 };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bgArgb } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border    = { bottom: { style: "thin", color: { argb: "FFAAAAAA" } } };
  });
  row.height = 24;
}

function styleDataRow(row, idx) {
  const argb = idx % 2 === 0 ? "FFF5F8FF" : "FFFFFFFF";
  row.eachCell(cell => {
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: argb } };
    cell.font      = { name: "Arial", size: 10 };
    cell.alignment = { vertical: "middle", wrapText: false };
  });
  row.height = 18;
}

function makeHyperlink(cell, url) {
  cell.value = { text: url, hyperlink: url };
  cell.font  = { name: "Arial", size: 10, color: { argb: "FF1155CC" }, underline: true };
}

function freezeAndFilter(ws, colCount) {
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colCount } };
}

// ══════════════════════════════════════════════════════════════════════════════
//  STEP 1 — Crawl all URLs
// ══════════════════════════════════════════════════════════════════════════════
async function crawlAllUrls(startUrl, browser) {
  const startNorm  = normalizeUrl(startUrl);
  const baseDomain = new URL(startNorm).hostname;

  const visited = new Set();
  const found   = new Set();
  const queue   = [startNorm];
  const errors  = [];
  let crawledCount = 0;

  const context = await browser.newContext({
    userAgent: "Mozilla/5.0 (SEO-Audit-Bot/1.0)",
    ignoreHTTPSErrors: true,
  });

  async function crawlOne(url) {
    if (visited.has(url)) return;
    visited.add(url);
    found.add(url);
    const page = await context.newPage();
    try {
      const res = await page.goto(url, { timeout: CONFIG.crawlTimeout, waitUntil: "domcontentloaded" });
      crawledCount++;
      process.stdout.write(`\r  [Crawl] Pages: ${crawledCount} | URLs: ${found.size} | Queue: ${queue.length}   `);
      if ((res?.status() ?? 0) >= 400) { errors.push({ url, status: res.status() }); return; }

      const hrefs = await page.evaluate(() => {
        const s = new Set();
        document.querySelectorAll("a[href],link[href],area[href]")
          .forEach(el => { if (el.href) s.add(el.href); });
        return [...s];
      });

      for (const href of hrefs) {
        const norm = normalizeUrl(href);
        if (norm && shouldCrawl(norm, baseDomain) && !visited.has(norm) && !queue.includes(norm)) {
          found.add(norm);
          queue.push(norm);
        }
      }
    } catch (e) { errors.push({ url, error: e.message }); }
    finally { await page.close(); }
  }

  while (queue.length > 0) {
    const batch = [];
    while (queue.length > 0 && batch.length < CONFIG.crawlConcurrency) {
      const next = queue.shift();
      if (next && !visited.has(next)) batch.push(next);
    }
    if (!batch.length) break;
    await Promise.all(batch.map(crawlOne));
  }

  await context.close();
  console.log(`\n  ✅ Crawl done — ${found.size} URLs found`);
  return { urls: [...found].sort(), errors, baseDomain };
}

// ══════════════════════════════════════════════════════════════════════════════
//  STEP 2 — SEO Details
// ══════════════════════════════════════════════════════════════════════════════
async function fetchSeoDetails(urls, browser, bypassHttpCheck = false) {
  const results = [];
  const context = await browser.newContext({
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/114.0 Safari/537.36",
    ignoreHTTPSErrors: true,
  });

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    process.stdout.write(`\r  [SEO] ${i + 1}/${urls.length}: ${url.slice(0, 70)}   `);

    // ── Redirect / status check ───────────────────────────────────────────────
    let status, finalUrl, redirected;

    if (bypassHttpCheck) {
      // For Cloudflare-protected sites: use Playwright directly, skip raw HTTP check
      finalUrl   = url;
      status     = 200;
      redirected = false;
    } else {
      ({ status, finalUrl, redirected } = await getStatusAndFinalUrl(url));
      if (status !== 200) {
        results.push({ url, finalUrl, status, redirected });
        continue;
      }
    }

    // Use finalUrl for data extraction
    const page = await context.newPage();
    const row  = { url, finalUrl, status, redirected };

    try {
      const res  = await page.goto(finalUrl, { timeout: 30000, waitUntil: bypassHttpCheck ? "networkidle" : "domcontentloaded" });
      if (bypassHttpCheck) { await page.waitForTimeout(2000); row.status = res?.status() ?? -1; }

      // Meta Title
      row.metaTitle = await page.$eval("head > title", el => el.innerText.trim()).catch(() => "Not Found");
      row.metaTitleLen = row.metaTitle !== "Not Found" ? row.metaTitle.length : 0;

      // Meta Description
      row.metaDesc = await page.$eval("meta[name='description']", el => el.getAttribute("content")?.trim() ?? "").catch(() => "Not Found");
      row.metaDescLen = row.metaDesc !== "Not Found" ? row.metaDesc.length : 0;

      // Canonical
      row.canonical = await page.$eval("link[rel='canonical']", el => el.href).catch(() => "Not Found");

      // H1
      const h1s = await page.$$eval("h1", els => els.map(e => e.innerText.trim().replace(/\s+/g, " ")));
      row.h1 = h1s.length ? h1s.join("\n") : "Not Found";
      row.h1Count = h1s.length;

      // H2
      const h2s = await page.$$eval("h2", els => els.map(e => e.innerText.trim().replace(/\s+/g, " ")));
      row.h2 = h2s.length ? h2s.join("\n") : "Not Found";
      row.h2Count = h2s.length;

      // OG Tags
      row.ogTitle       = await page.$eval("meta[property='og:title']",       el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogDescription = await page.$eval("meta[property='og:description']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogImage       = await page.$eval("meta[property='og:image']",       el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogUrl         = await page.$eval("meta[property='og:url']",         el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      // Twitter Tags
      row.twitterCard  = await page.$eval("meta[name='twitter:card'],meta[property='twitter:card']",   el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.twitterTitle = await page.$eval("meta[name='twitter:title'],meta[property='twitter:title']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.twitterDesc  = await page.$eval("meta[name='twitter:description'],meta[property='twitter:description']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      // Robots meta
      row.robotsMeta = await page.$eval("meta[name='robots']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      // Schema
      const schemas = await page.$$eval("script[type='application/ld+json']", els => els.map(e => e.innerText.trim()));
      row.schema = schemas.length ? schemas.join("\n---\n") : "Not Found";

    } catch (e) {
      row.error = e.message;
    } finally {
      await page.close();
    }

    results.push(row);
  }

  await context.close();
  console.log(`\n  ✅ SEO fetch done — ${results.length} pages processed`);
  return results;
}

// ══════════════════════════════════════════════════════════════════════════════
//  STEP 3 — Alt Tag Audit
// ══════════════════════════════════════════════════════════════════════════════
async function fetchAltTags(urls, baseDomain, browser) {
  const results = [];
  const context = await browser.newContext({ ignoreHTTPSErrors: true });

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    if (!url.includes(baseDomain)) continue;
    process.stdout.write(`\r  [Alt] ${i + 1}/${urls.length}: ${url.slice(0, 70)}   `);

    const page = await context.newPage();
    try {
      await page.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });
      const images = await page.$$eval("img", imgs =>
        imgs.map(img => ({
          src : img.getAttribute("src") ?? "",
          alt : img.getAttribute("alt"),
          width : img.naturalWidth,
          height: img.naturalHeight,
        }))
      );

      const MEDIA_RE = /\.(jpg|jpeg|png|gif|svg|bmp|ico|mp4|avi|mpeg|mpg|mov|flv|wmv|webm|webp|ogg|mp3|wav|flac|aac|wma|pdf)$/i;
      for (const img of images) {
        if (img.src && MEDIA_RE.test(img.src)) {
          results.push({
            pageUrl   : url,
            imageSrc  : img.src,
            altTag    : img.alt === null ? "No Alt Attribute" : img.alt.trim() === "" ? "Empty Alt Tag" : img.alt.trim(),
            hasAlt    : img.alt !== null && img.alt.trim() !== "",
            dimensions: img.width && img.height ? `${img.width}x${img.height}` : "Unknown",
          });
        }
      }
    } catch (e) {
      results.push({ pageUrl: url, imageSrc: "Error", altTag: e.message, hasAlt: false, dimensions: "" });
    } finally {
      await page.close();
    }
  }

  await context.close();
  console.log(`\n  ✅ Alt tag done — ${results.length} images found`);
  return results;
}

// ══════════════════════════════════════════════════════════════════════════════
//  STEP 4 — Internal / External Link Collector
// ══════════════════════════════════════════════════════════════════════════════
async function collectLinks(urls, baseDomain, browser) {
  const results    = [];
  const visitedHrefs = new Set();
  const context    = await browser.newContext({ ignoreHTTPSErrors: true });

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    process.stdout.write(`\r  [Links] ${i + 1}/${urls.length}: ${url.slice(0, 70)}   `);

    const page = await context.newPage();
    try {
      await page.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });
      const hrefs = await page.$$eval("a[href]", els =>
        els.map(a => a.href).filter(h => h && h.startsWith("http"))
      );

      for (const href of hrefs) {
        const norm = normalizeUrl(href);
        if (!norm || visitedHrefs.has(norm)) continue;
        visitedHrefs.add(norm);

        let linkHost;
        try { linkHost = new URL(norm).hostname; } catch { continue; }

        const isInternal = linkHost === baseDomain || linkHost.endsWith("." + baseDomain);
        const { status } = await getStatusAndFinalUrl(norm);

        results.push({
          pageUrl  : url,
          linkUrl  : norm,
          type     : isInternal ? "Internal" : "External",
          status,
          statusLabel: status === 200 ? "OK" : status === 301 ? "Redirect" : status === 404 ? "Broken" : status === -1 ? "Error" : String(status),
        });
      }
    } catch (e) {
      results.push({ pageUrl: url, linkUrl: "Error", type: "", status: -1, statusLabel: e.message });
    } finally {
      await page.close();
    }
  }

  await context.close();
  console.log(`\n  ✅ Link collection done — ${results.length} links found`);
  return results;
}

// ══════════════════════════════════════════════════════════════════════════════
//  EXCEL WRITER — 4 sheets + Summary
// ══════════════════════════════════════════════════════════════════════════════
async function buildExcel(siteName, crawlData, seoData, altData, linkData) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "SEO-Audit-Tool";
  wb.created = new Date();

  // ── Summary Sheet FIRST so it appears as tab 1 ───────────────────────────
  const wss = wb.addWorksheet("Summary");
  wss.columns = [{ key: "label", width: 30 }, { key: "value", width: 40 }];

  // ── Sheet 1: All URLs ──────────────────────────────────────────────────────
  const ws1 = wb.addWorksheet("Sheet1 - All URLs");
  ws1.columns = [
    { header: "#",        key: "no",      width: 6  },
    { header: "URL",      key: "url",     width: 80 },
    { header: "Domain",   key: "domain",  width: 30 },
    { header: "Path",     key: "urlPath", width: 50 },
    { header: "Protocol", key: "proto",   width: 12 },
  ];
  styleHeader(ws1.getRow(1), "FF1F3864");
  crawlData.urls.forEach((url, i) => {
    let domain = "", urlPath = "", proto = "";
    try { const u = new URL(url); domain = u.hostname; urlPath = u.pathname; proto = u.protocol.replace(":",""); } catch {}
    const r = ws1.addRow({ no: i + 1, url, domain, urlPath, proto });
    styleDataRow(r, i);
    makeHyperlink(r.getCell("url"), url);
  });
  freezeAndFilter(ws1, 5);

  // ── Sheet 2: SEO Data ─────────────────────────────────────────────────────
  const ws2 = wb.addWorksheet("Sheet2 - SEO Data");
  ws2.columns = [
    { header: "Input URL",       key: "url",          width: 55 },
    { header: "Final URL",       key: "finalUrl",     width: 55 },
    { header: "Status Code",     key: "status",       width: 13 },
    { header: "Redirected?",     key: "redirected",   width: 13 },
    { header: "Meta Title",      key: "metaTitle",    width: 50 },
    { header: "Title Length",    key: "metaTitleLen", width: 13 },
    { header: "Meta Description",key: "metaDesc",     width: 60 },
    { header: "Desc Length",     key: "metaDescLen",  width: 12 },
    { header: "Canonical",       key: "canonical",    width: 55 },
    { header: "H1 Tag(s)",       key: "h1",           width: 50 },
    { header: "H1 Count",        key: "h1Count",      width: 10 },
    { header: "H2 Tags",         key: "h2",           width: 50 },
    { header: "H2 Count",        key: "h2Count",      width: 10 },
    { header: "OG Title",        key: "ogTitle",      width: 40 },
    { header: "OG Description",  key: "ogDescription",width: 50 },
    { header: "OG Image",        key: "ogImage",      width: 55 },
    { header: "OG URL",          key: "ogUrl",        width: 55 },
    { header: "Twitter Card",    key: "twitterCard",  width: 20 },
    { header: "Twitter Title",   key: "twitterTitle", width: 40 },
    { header: "Twitter Desc",    key: "twitterDesc",  width: 50 },
    { header: "Robots Meta",     key: "robotsMeta",   width: 25 },
    { header: "Schema",          key: "schema",       width: 40 },
  ];
  styleHeader(ws2.getRow(1), "FF1F4E79");
  seoData.forEach((d, i) => {
    const r = ws2.addRow({
      url: d.url, finalUrl: d.finalUrl ?? d.url, status: d.status,
      redirected: d.redirected ? "Yes" : "No",
      metaTitle: d.metaTitle ?? "", metaTitleLen: d.metaTitleLen ?? "",
      metaDesc: d.metaDesc ?? "", metaDescLen: d.metaDescLen ?? "",
      canonical: d.canonical ?? "", h1: d.h1 ?? "", h1Count: d.h1Count ?? "",
      h2: d.h2 ?? "", h2Count: d.h2Count ?? "",
      ogTitle: d.ogTitle ?? "", ogDescription: d.ogDescription ?? "",
      ogImage: d.ogImage ?? "", ogUrl: d.ogUrl ?? "",
      twitterCard: d.twitterCard ?? "", twitterTitle: d.twitterTitle ?? "",
      twitterDesc: d.twitterDesc ?? "", robotsMeta: d.robotsMeta ?? "",
      schema: d.schema ?? "",
    });
    styleDataRow(r, i);
    makeHyperlink(r.getCell("url"), d.url);
    if (d.finalUrl && d.finalUrl !== d.url) makeHyperlink(r.getCell("finalUrl"), d.finalUrl);
    r.getCell("h1").alignment = { vertical: "middle", wrapText: true };
    r.getCell("h2").alignment = { vertical: "middle", wrapText: true };

    // Highlight bad title/desc lengths
    const tLen = d.metaTitleLen ?? 0;
    if (tLen > 0 && (tLen < 50 || tLen > 60)) {
      r.getCell("metaTitleLen").font = { name: "Arial", size: 10, color: { argb: "FFCC0000" }, bold: true };
    }
    const dLen = d.metaDescLen ?? 0;
    if (dLen > 0 && (dLen < 150 || dLen > 160)) {
      r.getCell("metaDescLen").font = { name: "Arial", size: 10, color: { argb: "FFCC0000" }, bold: true };
    }
  });
  freezeAndFilter(ws2, ws2.columns.length);

  // ── Sheet 3: Alt Tags ────────────────────────────────────────────────────
  const ws3 = wb.addWorksheet("Sheet3 - Alt Tags");
  ws3.columns = [
    { header: "#",          key: "no",         width: 6  },
    { header: "Page URL",   key: "pageUrl",    width: 60 },
    { header: "Image URL",  key: "imageSrc",   width: 70 },
    { header: "Alt Tag",    key: "altTag",     width: 50 },
    { header: "Has Alt?",   key: "hasAlt",     width: 12 },
    { header: "Dimensions", key: "dimensions", width: 15 },
  ];
  styleHeader(ws3.getRow(1), "FF375623");
  altData.forEach((d, i) => {
    const r = ws3.addRow({ no: i + 1, pageUrl: d.pageUrl, imageSrc: d.imageSrc, altTag: d.altTag, hasAlt: d.hasAlt ? "Yes" : "No", dimensions: d.dimensions });
    styleDataRow(r, i);
    makeHyperlink(r.getCell("pageUrl"), d.pageUrl);
    if (d.imageSrc && d.imageSrc.startsWith("http")) makeHyperlink(r.getCell("imageSrc"), d.imageSrc);
    if (!d.hasAlt) {
      r.getCell("hasAlt").font = { name: "Arial", size: 10, color: { argb: "FFCC0000" }, bold: true };
      r.getCell("altTag").font = { name: "Arial", size: 10, color: { argb: "FFCC0000" } };
    }
  });
  freezeAndFilter(ws3, 6);

  // ── Sheet 4: Internal / External Links ────────────────────────────────────
  const ws4 = wb.addWorksheet("Sheet4 - Links");
  ws4.columns = [
    { header: "#",           key: "no",          width: 6  },
    { header: "Found On",    key: "pageUrl",     width: 60 },
    { header: "Link URL",    key: "linkUrl",     width: 70 },
    { header: "Type",        key: "type",        width: 12 },
    { header: "Status Code", key: "status",      width: 13 },
    { header: "Status",      key: "statusLabel", width: 12 },
  ];
  styleHeader(ws4.getRow(1), "FF7B2C2C");
  linkData.forEach((d, i) => {
    const r = ws4.addRow({ no: i + 1, pageUrl: d.pageUrl, linkUrl: d.linkUrl, type: d.type, status: d.status, statusLabel: d.statusLabel });
    styleDataRow(r, i);
    makeHyperlink(r.getCell("pageUrl"), d.pageUrl);
    if (d.linkUrl && d.linkUrl.startsWith("http")) makeHyperlink(r.getCell("linkUrl"), d.linkUrl);
    if (d.statusLabel === "Broken" || d.status === 404) {
      r.getCell("statusLabel").font = { name: "Arial", size: 10, color: { argb: "FFCC0000" }, bold: true };
    }
    if (d.type === "External") {
      r.getCell("type").font = { name: "Arial", size: 10, color: { argb: "FF7B2C2C" }, bold: true };
    }
  });
  freezeAndFilter(ws4, 6);

  // ── Summary Sheet — fill data ─────────────────────────────────────────────
  const missingAlt    = altData.filter(d => !d.hasAlt).length;
  const brokenLinks   = linkData.filter(d => d.status === 404 || d.status === -1).length;
  const internalLinks = linkData.filter(d => d.type === "Internal").length;
  const externalLinks = linkData.filter(d => d.type === "External").length;
  const redirectedPages = seoData.filter(d => d.redirected).length;

  const rows = [
    ["SEO Audit Summary", ""],
    ["Website",                  siteName],
    ["Audit Date",               new Date().toLocaleString()],
    ["",""],
    ["📄 URLs",                  ""],
    ["  Total URLs Crawled",     crawlData.urls.length],
    ["  Crawl Errors",           crawlData.errors.length],
    ["",""],
    ["🔍 SEO Data",              ""],
    ["  Pages Audited",          seoData.length],
    ["  Redirected Pages",       redirectedPages],
    ["  Missing Meta Title",     seoData.filter(d => !d.metaTitle || d.metaTitle === "Not Found").length],
    ["  Missing Meta Desc",      seoData.filter(d => !d.metaDesc  || d.metaDesc  === "Not Found").length],
    ["",""],
    ["🖼️ Images",                ""],
    ["  Total Images",           altData.length],
    ["  Missing Alt Tags",       missingAlt],
    ["",""],
    ["🔗 Links",                 ""],
    ["  Total Links",            linkData.length],
    ["  Internal Links",         internalLinks],
    ["  External Links",         externalLinks],
    ["  Broken Links (404)",     brokenLinks],
  ];

  rows.forEach(([label, value], i) => {
    const r = wss.addRow({ label, value });
    if (i === 0) {
      r.getCell("label").font = { bold: true, size: 15, name: "Arial", color: { argb: "FF1F3864" } };
      r.eachCell(c => c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } });
      r.height = 30;
    } else if (label.startsWith("📄") || label.startsWith("🔍") || label.startsWith("🖼️") || label.startsWith("🔗")) {
      r.getCell("label").font = { bold: true, size: 12, name: "Arial", color: { argb: "FF1F3864" } };
      r.eachCell(c => c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9EFF8" } });
      r.height = 22;
    } else {
      r.getCell("label").font = { name: "Arial", size: 11 };
      r.getCell("value").font = { name: "Arial", size: 11, bold: true };
      if (i % 2 === 0) r.eachCell(c => c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } });
      r.height = 20;
    }
  });

  return wb;
}

// ══════════════════════════════════════════════════════════════════════════════
//  MAIN — Runs all steps for each website
// ══════════════════════════════════════════════════════════════════════════════
async function runAudit() {
  console.log("\n╔══════════════════════════════════════════╗");
  console.log("║       SEO AUDIT TOOL — Starting          ║");
  console.log(`║  ${new Date().toLocaleString().padEnd(40)}║`);
  console.log("╚══════════════════════════════════════════╝\n");

  const browser = await chromium.launch({ headless: true });
  const attachments = [];

  for (const site of CONFIG.websites) {
    const siteUrl    = site.url;
    const bypass     = site.bypassHttpCheck ?? false;
    const siteName   = new URL(siteUrl).hostname;
    const timestamp  = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputFile = `SEO_Audit_${siteName}_${timestamp}.xlsx`;

    console.log(`\n${"═".repeat(55)}`);
    console.log(`  🌐  Site: ${siteUrl}`);
    console.log(`  🔒  Bypass HTTP Check: ${bypass ? "Yes (Cloudflare mode)" : "No (Fast mode)"}`);
    console.log(`${"═".repeat(55)}\n`);

    console.log("  📡  STEP 1/4 — Crawling all URLs...");
    const crawlData = await crawlAllUrls(siteUrl, browser);

    console.log("\n  🔍  STEP 2/4 — Fetching SEO details...");
    const seoData   = await fetchSeoDetails(crawlData.urls, browser, bypass);

    console.log("\n  🖼️   STEP 3/4 — Auditing image alt tags...");
    const altData   = await fetchAltTags(crawlData.urls, crawlData.baseDomain, browser);

    console.log("\n  🔗  STEP 4/4 — Collecting internal/external links...");
    const linkData  = await collectLinks(crawlData.urls, crawlData.baseDomain, browser);

    console.log("\n  💾  Building Excel report...");
    const workbook  = await buildExcel(siteName, crawlData, seoData, altData, linkData);
    await workbook.xlsx.writeFile(outputFile);
    console.log(`  📊  Saved → ${outputFile}`);

    attachments.push({ filename: path.basename(outputFile), path: outputFile, siteName });
    console.log(`\n  ✅  ${siteName} — COMPLETE!\n`);
  }

  await browser.close();

  // Send one combined email with all site reports attached
  if (CONFIG.email.enabled && attachments.length > 0) {
    console.log("\n  📧  Sending combined email report...");
    await sendCombinedEmail(attachments);
  }

  console.log("\n╔══════════════════════════════════════════╗");
  console.log("║        ALL AUDITS COMPLETE! 🎉           ║");
  console.log("╚══════════════════════════════════════════╝\n");

  // On GitHub Actions (non-TTY) exit cleanly after all audits complete
  if (!process.stdin.isTTY) process.exit(0);
}

async function sendCombinedEmail(attachments) {
  const transporter = nodemailer.createTransport({
    host   : CONFIG.email.host,
    port   : CONFIG.email.port,
    secure : CONFIG.email.secure,
    auth   : { user: CONFIG.email.user, pass: CONFIG.email.pass },
  });

  const siteList = attachments.map(a => `  • ${a.siteName}`).join("\n");
  const date     = new Date().toLocaleDateString("en-IN", { year: "numeric", month: "long", day: "numeric" });

  await transporter.sendMail({
    from       : `"SEO Audit Bot" <${CONFIG.email.user}>`,
    to         : CONFIG.email.to,
    subject    : `Monthly SEO Audit Report — ${date}`,
    text       : `Hello,\n\nPlease find attached the Monthly SEO Audit Reports for:\n\n${siteList}\n\nAudit completed on: ${new Date().toLocaleString()}\n\nRegards,\nSanket Ghadmode`,
    attachments: attachments.map(a => ({ filename: a.filename, path: a.path })),
  });

  console.log(`  ✅ Email sent to ${CONFIG.email.to} with ${attachments.length} report(s) attached`);
}

// ── Schedule + controls ───────────────────────────────────────────────────────
let task = null;

if (CONFIG.schedule) {
  task = cron.schedule(CONFIG.schedule, () => {
    console.log(`\n  ⏰  Scheduled trigger fired at ${new Date().toLocaleString()}`);
    runAudit().catch(err => console.error("  ❌  Audit failed:", err.message));
  });
  console.log(`\n  🕐  Scheduler started — ${CONFIG.schedule}`);
} else {
  console.log(`\n  ⏸️   Scheduler is DISABLED — set CONFIG.schedule to enable.`);
}
console.log(`  📋  Schedule: ${CONFIG.schedule ?? "disabled"}`);
console.log(`  ⌨️   Controls: [P] Pause  [R] Resume  [S] Run now  [Q] Quit\n`);

// Keyboard controls
if (process.stdin.isTTY) {
  process.stdin.setRawMode(true);
  process.stdin.resume();
  process.stdin.on("data", (key) => {
    const k = key.toString().toLowerCase();
    if (k === "p") {
      task?.stop();
      console.log("\n  ⏸️   Scheduler PAUSED — audit will not run on schedule.");
      console.log("  ⌨️   Press [R] to resume, [S] to run now, [Q] to quit.");
    } else if (k === "r") {
      task?.start();
      console.log("\n  ▶️   Scheduler RESUMED — next run: " + CONFIG.schedule);
    } else if (k === "s") {
      console.log("\n  🚀  Manual run triggered...");
      runAudit().catch(err => console.error("  ❌  Audit failed:", err.message));
    } else if (k === "q" || key[0] === 3) {
      console.log("\n  👋  Shutting down scheduler. Goodbye!\n");
      task?.stop();
      process.exit(0);
    }
  });
}

// Run immediately on startup as well
runAudit().catch(err => console.error("  ❌  Audit failed:", err.message));