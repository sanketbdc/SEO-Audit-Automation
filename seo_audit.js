/**
 * ============================================================
 *  COMPLETE SEO AUDIT TOOL  —  Playwright + ExcelJS
 * ============================================================
 *  Pipeline (runs automatically in order):
 *   1. Crawl all URLs on the domain          → Sheet 1
 *   2. Fetch SEO details (with redirect fix) → Sheet 2
 *   3. Collect image alt tags                → Sheet 3
 *   4. Internal / External link audit        → Sheet 4
 *   5. AI Visibility Audit (12-layer)        → Sheet 5  ← NEW
 *   6. Email the Excel report (HTML email)   ← ENHANCED
 *
 *  Quick Run Commands:
 *    node --max-old-space-size=8192 seo_audit.js                  → Full audit (all steps)
 *    node --max-old-space-size=8192 seo_audit.js --only=crawl     → Only crawl URLs (Sheet 1)
 *    node --max-old-space-size=8192 seo_audit.js --only=seo       → Only SEO audit (Sheet 2)
 *    node --max-old-space-size=8192 seo_audit.js --only=alt       → Only image alt tags (Sheet 3)
 *    node --max-old-space-size=8192 seo_audit.js --only=links     → Only internal/external links (Sheet 4)
 *    node --max-old-space-size=8192 seo_audit.js --only=ai        → Only AI visibility audit (Sheet 5)
 *    node --max-old-space-size=8192 seo_audit.js --urls=my_urls.txt --only=seo
 *
 *  audit_urls.txt format (one entry per line):
 *    https://www.example.com                  → website URL (crawls internal links)
 *    https://www.example.com/sitemap.xml      → direct sitemap URL
 *    https://www.example.com/sitemap.xml|true → sitemap + bypass HTTP check (Cloudflare)
 *    # https://www.skipped.com                → commented out = skipped
 *
 *  URL Discovery Logic (per site):
 *    1. If input URL is a sitemap/XML → parse it (supports nested sitemap indexes)
 *    2. If input URL is a website    → crawl the website via internal links
 *
 *  Config: edit the CONFIG section below before running.
 * ============================================================
 */

const { chromium }  = require("playwright");
const ExcelJS       = require("exceljs");
const nodemailer    = require("nodemailer");
const cron          = require("node-cron");
const fs            = require("fs");
const path          = require("path");
const https         = require("https");
const http          = require("http");

function normalizeAuditEntry(line, filePath) {
  if (!line) return null;

  const trimmed = line.trim();
  if (!trimmed || trimmed.startsWith("#") || trimmed.startsWith("//")) return null;

  let cleaned = trimmed;
  const markdownUrlMatch = cleaned.match(/https?:\/\/[^\s)]+/i);
  if (markdownUrlMatch) {
    cleaned = markdownUrlMatch[0];
  }

  cleaned = cleaned.replace(/^[-*]\s*/, "").replace(/^['"`]+|['"`]+$/g, "").trim();
  if (!cleaned) return null;

  const [rawUrl, bypass] = cleaned.split("|").map(value => value.trim());
  if (!rawUrl) return null;

  const urlMatch = rawUrl.match(/https?:\/\/[^\s)]+/i);
  const url = urlMatch ? urlMatch[0].replace(/[),]+$/, "") : rawUrl;

  try { new URL(url); } catch { throw new Error(`Invalid URL in ${filePath}: ${rawUrl}`); }
  return { url, bypassHttpCheck: bypass?.toLowerCase() === "true" };
}

function loadAuditSites() {
  const urlsArg = process.argv.find(arg => arg.startsWith("--urls="));
  const filePath = path.resolve(urlsArg ? urlsArg.slice("--urls=".length) : path.join(__dirname, "audit_urls.txt"));

  if (!fs.existsSync(filePath)) {
    console.warn(`  ⚠ URL list not found: ${filePath}. Using the default site.`);
    return [{ url: "https://www.godrejenterprises.com/", bypassHttpCheck: false }];
  }

  const sites = fs.readFileSync(filePath, "utf8")
    .split(/\r?\n/)
    .map(line => normalizeAuditEntry(line, filePath))
    .filter(Boolean);

  if (!sites.length) throw new Error(`No URLs found in ${filePath}`);
  console.log(`  📄 URL list: ${filePath} (${sites.length} site${sites.length === 1 ? "" : "s"})`);
  return sites;
}

// ── Sitemap / URL Discovery ───────────────────────────────────────────────────

/** Fetch raw text from a URL via http/https */
function fetchText(url, timeoutMs = 15000) {
  return new Promise((resolve) => {
    try {
      const parsed = new URL(url);
      const lib = parsed.protocol === "https:" ? https : http;
      let body = "";
      const req = lib.get(
        { hostname: parsed.hostname, path: parsed.pathname + parsed.search,
          headers: { "User-Agent": "Mozilla/5.0 (SEO-Audit-Bot/1.0)" },
          timeout: timeoutMs, rejectUnauthorized: false },
        (res) => {
          if ([301,302,303,307,308].includes(res.statusCode) && res.headers.location) {
            return fetchText(new URL(res.headers.location, url).href, timeoutMs).then(resolve);
          }
          res.setEncoding("utf8");
          res.on("data", d => { body += d; });
          res.on("end", () => resolve({ ok: true, body, status: res.statusCode, contentType: res.headers["content-type"] || "" }));
        }
      );
      req.on("error", () => resolve({ ok: false, body: "", status: -1, contentType: "" }));
      req.on("timeout", () => { req.destroy(); resolve({ ok: false, body: "", status: -1, contentType: "" }); });
    } catch { resolve({ ok: false, body: "", status: -1, contentType: "" }); }
  });
}

/** Extract all <loc> values from a sitemap XML string */
function extractLocsFromXml(xml) {
  const locs = [];
  const re = /<loc>\s*([^<]+?)\s*<\/loc>/gi;
  let m;
  while ((m = re.exec(xml)) !== null) locs.push(m[1].trim());
  return locs;
}

/** Determine if XML is a sitemap index (contains <sitemapindex>) */
function isSitemapIndex(xml) {
  return /<sitemapindex[\s>]/i.test(xml);
}

/** Determine if XML is a urlset (contains <urlset>) */
function isUrlSet(xml) {
  return /<urlset[\s>]/i.test(xml);
}

/** Determine if content looks like XML/sitemap */
function looksLikeXml(body, contentType) {
  if (/xml/i.test(contentType)) return true;
  const trimmed = body.trimStart();
  return trimmed.startsWith("<?xml") || trimmed.startsWith("<sitemapindex") || trimmed.startsWith("<urlset");
}

/**
 * Recursively parse a sitemap URL.
 * Returns array of page URLs (not sub-sitemap URLs).
 */
async function parseSitemapRecursive(sitemapUrl, baseDomain, visited = new Set(), depth = 0) {
  if (depth > 10 || visited.has(sitemapUrl)) return [];
  if (!isSameHost(sitemapUrl, baseDomain)) {
    console.warn(`  ⚠ Skipping sitemap on a different host: ${sitemapUrl}`);
    return [];
  }
  visited.add(sitemapUrl);

  const { ok, body, status } = await fetchText(sitemapUrl);
  if (!ok || !body) {
    console.warn(`  ⚠ Sitemap fetch failed (${status}): ${sitemapUrl}`);
    return [];
  }

  if (!looksLikeXml(body, "")) {
    console.warn(`  ⚠ Not XML content at: ${sitemapUrl}`);
    return [];
  }

  const locs = extractLocsFromXml(body);

  if (isSitemapIndex(body)) {
    // Nested sitemap — recurse into each child sitemap
    const indent = "  ".repeat(depth + 1);
    console.log(`${indent}📂 Sitemap index: ${sitemapUrl} → ${locs.length} child sitemaps`);
    const pageUrls = [];
    for (const loc of locs) {
      const childUrls = await parseSitemapRecursive(loc, baseDomain, visited, depth + 1);
      pageUrls.push(...childUrls);
    }
    return pageUrls;
  }

  if (isUrlSet(body)) {
    // Direct page URLs — filter to same domain
    const indent = "  ".repeat(depth + 1);
    const filtered = locs.filter(u => {
      return isSameHost(u, baseDomain);
    });
    console.log(`${indent}📄 URL set: ${sitemapUrl} → ${filtered.length} page URLs`);
    return filtered;
  }

  console.warn(`  ⚠ Unrecognised XML structure at: ${sitemapUrl}`);
  return [];
}

/**
 * Master URL discovery function.
 * Auto-detects whether input is a sitemap URL or website URL.
 * Returns { urls[], baseDomain, errors[], discoveryMethod }
 */
async function discoverUrls(inputUrl, browser) {
  let parsed;
  try { parsed = new URL(inputUrl); } catch { throw new Error(`Invalid URL: ${inputUrl}`); }
  const baseDomain = parsed.hostname;

  // Only parse the supplied URL when it is XML; website URLs go straight to crawling.
  const res = await fetchText(inputUrl);
  const isXmlPath = /\.xml(?:$|[?#])/i.test(inputUrl);
  if ((res.ok && res.body && looksLikeXml(res.body, res.contentType)) || isXmlPath) {
    console.log(`  🗺️  Input is a sitemap URL — parsing...`);
    const pageUrls = await parseSitemapRecursive(inputUrl, baseDomain, new Set(), 0);
    if (pageUrls.length > 0) {
      const unique = [...new Set(pageUrls.map(u => normalizeUrl(u)).filter(Boolean))].sort();
      console.log(`  ✅ Sitemap discovery complete — ${unique.length} unique URLs`);
      return { urls: unique, baseDomain, errors: [], discoveryMethod: "sitemap" };
    }
    console.warn(`  ⚠ Sitemap parsed but no URLs found`);
    return { urls: [], baseDomain, errors: [], discoveryMethod: "sitemap" };
  }

  console.log(`  🌐 Input is a website URL — crawling internal links...`);
  console.log(`  🕷️  Starting website crawl from: ${inputUrl}`);
  const crawlResult = await crawlAllUrls(inputUrl, browser);
  return { ...crawlResult, discoveryMethod: "crawl" };
}

const manualUrlMode = process.argv.some(arg => arg.startsWith("--urls=")) ||
  fs.existsSync(path.join(__dirname, "audit_urls.txt"));

function parseEmailList(value) {
  return String(value || "")
    .split(/[\n,;]/)
    .map(email => email.trim())
    .filter(Boolean);
}

// ╔══════════════════════════════════════════════════════════╗
// ║                   USER CONFIG                           ║
// ╚══════════════════════════════════════════════════════════╝
const CONFIG = {

  // ── Websites to audit (add as many as you want) ──────────
  // bypassHttpCheck: true  → use for Cloudflare/security protected sites (slower)
  // bypassHttpCheck: false → default, fast raw HTTP check
  websites: loadAuditSites().filter(s => !process.env.AUDIT_SITE || s.url === process.env.AUDIT_SITE),

  // ── Output Excel file path ────────────────────────────────
  outputFile: "SEO_Audit_Report.xlsx",

  // ── Crawler settings ──────────────────────────────────────
  // Set this high enough to allow a real full-site crawl rather than stopping
  // after a fixed number of URLs.
  crawlMaxUrls     : 20000,
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
    to      : parseEmailList(process.env.EMAIL_TO || process.env.EMAIL_RECIPIENTS || "sanket@bombaydc.com"),
  },

  // ── Scheduler settings ────────────────────────────────────
  // GitHub Actions uses UTC time. This runs every Monday at 06:00 AM IST = 00:30 UTC.
  // Cron format: 'minute hour day month weekday'
  // Example: '30 0 * * 1' = every Monday at 00:30 UTC (06:00 IST)
  schedule: process.env.SEO_AUDIT_SCHEDULE || "30 0 * * 1",
};
// ╚══════════════════════════════════════════════════════════╝


// ── Helpers ───────────────────────────────────────────────────────────────────

const MEDIA_EXT = new Set([
  ".jpg",".jpeg",".png",".gif",".svg",".webp",".ico",".bmp",
  ".mp4",".avi",".mov",".webm",".ogg",".ogv",".mp3",".wav",".m4a",
  ".pdf",".zip",".tar",".gz",".7z",".rar",".exe",
  ".css",".js",".woff",".woff2",".ttf",".eot",".otf",".map",
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
    if (!isSameHost(urlStr, baseDomain)) return false;
    const pathname = u.pathname.split("?")[0].split("#")[0];
    const ext = path.extname(pathname).toLowerCase();
    if (MEDIA_EXT.has(ext)) return false;
    return true;
  } catch { return false; }
}

function isSameHost(urlStr, baseDomain) {
  try {
    const hostname = new URL(urlStr).hostname.toLowerCase();
    const base = baseDomain.toLowerCase();
    return hostname === base || hostname === `www.${base}` || base === `www.${hostname}`;
  } catch { return false; }
}

/** HTTP HEAD/GET status + final redirect URL — optionally returns response headers */
function getStatusAndFinalUrl(inputUrl, returnHeaders = false) {
  return new Promise((resolve) => {
    const doRequest = (url, method, redirectCount) => {
      if (redirectCount > 10) return resolve({ status: -1, finalUrl: url, redirected: true, headers: {} });
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
            if ((status === 403 || status === 405) && method === "HEAD") {
              return doRequest(url, "GET", redirectCount);
            }
            resolve({
              status,
              finalUrl    : url,
              redirected  : redirectCount > 0,
              originalUrl : inputUrl,
              headers     : returnHeaders ? res.headers : {},
            });
          }
        );
        req.on("error", () => resolve({ status: -1, finalUrl: url, redirected: false, headers: {} }));
        req.on("timeout", () => { req.destroy(); resolve({ status: -1, finalUrl: url, redirected: false, headers: {} }); });
        req.end();
      } catch { resolve({ status: -1, finalUrl: inputUrl, redirected: false, headers: {} }); }
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

// ── Score cell colour helper (Red / Amber / Green) ────────────────────────────
function colourScoreCell(cell, score) {
  let argb;
  if      (score >= 75) argb = "FF1E7145"; // Green
  else if (score >= 50) argb = "FF9C6500"; // Amber
  else                  argb = "FF9C0006"; // Red
  cell.font = { name: "Arial", size: 10, bold: true, color: { argb } };
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
          if (found.size >= CONFIG.crawlMaxUrls) break;
          found.add(norm);
          queue.push(norm);
        }
      }
    } catch (e) { errors.push({ url, error: e.message }); }
    finally { await page.close(); }
  }

  while (queue.length > 0 && visited.size < CONFIG.crawlMaxUrls) {
    const batch = [];
    while (queue.length > 0 && batch.length < CONFIG.crawlConcurrency && visited.size + batch.length < CONFIG.crawlMaxUrls) {
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

/**
 * Enhanced status check with redirect type detection
 * Returns: { status, finalUrl, redirected, redirectType, originalUrl, headers }
 */
function getStatusAndFinalUrlWithRedirectType(inputUrl, returnHeaders = false) {
  return new Promise((resolve) => {
    const doRequest = (url, method, redirectCount, chain = []) => {
      chain = Array.isArray(chain) ? chain : [];
      if (redirectCount > 10) return resolve({ status: -1, finalUrl: url, redirected: true, redirectType: "Limit exceeded", originalUrl: inputUrl, headers: {}, chain });
      try {
        const parsed = new URL(url);
        const lib = parsed.protocol === "https:" ? https : http;
        const req = lib.request(
          { hostname: parsed.hostname, path: parsed.pathname + parsed.search,
            method, headers: { "User-Agent": "Mozilla/5.0 (SEO-Audit-Bot/1.0)" },
            timeout: 15000, rejectUnauthorized: false },
          (res) => {
            const status = res.statusCode;
            const redirectCode = [301, 302, 303, 307, 308].includes(status) ? status : null;
            
            if (redirectCode && res.headers.location) {
              const next = new URL(res.headers.location, url).href;
              const redirectType = status === 301 ? "301 Moved Permanently" :
                                  status === 302 ? "302 Found" :
                                  status === 303 ? "303 See Other" :
                                  status === 307 ? "307 Temporary Redirect" :
                                  status === 308 ? "308 Permanent Redirect" : "Unknown";
              const newChain = [...chain, { from: url, to: next, type: redirectType }];
              return doRequest(next, method, redirectCount + 1, newChain);
            }
            if ((status === 403 || status === 405) && method === "HEAD") {
              return doRequest(url, "GET", redirectCount, chain);
            }
            
            const redirectType = chain.length > 0 ? chain.map(c => c.type).join(" → ") : null;
            resolve({
              status,
              finalUrl    : url,
              redirected  : redirectCount > 0,
              redirectType: redirectType,
              originalUrl : inputUrl,
              headers     : returnHeaders ? res.headers : {},
              chain,
            });
          }
        );
        req.on("error", () => resolve({ status: -1, finalUrl: url, redirected: false, redirectType: null, originalUrl: inputUrl, headers: {}, chain }));
        req.on("timeout", () => { req.destroy(); resolve({ status: -1, finalUrl: url, redirected: false, redirectType: null, originalUrl: inputUrl, headers: {}, chain }); });
        req.end();
      } catch { resolve({ status: -1, finalUrl: inputUrl, redirected: false, redirectType: null, originalUrl: inputUrl, headers: {}, chain }); }
    };
    doRequest(inputUrl, "HEAD", 0, []);
  });
}

async function fetchSeoDetails(urls, browser, bypassHttpCheck = false) {
  const results = [];
  const context = await browser.newContext({
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/114.0 Safari/537.36",
    ignoreHTTPSErrors: true,
  });

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    process.stdout.write(`\r  [SEO] ${i + 1}/${urls.length}: ${url.slice(0, 70)}   `);

    let status, finalUrl, redirected, redirectType;

    if (bypassHttpCheck) {
      finalUrl      = url;
      status        = 200;
      redirected    = false;
      redirectType  = null;
    } else {
      ({ status, finalUrl, redirected, redirectType } = await getStatusAndFinalUrlWithRedirectType(url));
      if (status !== 200) {
        results.push({
          seedUrl: url,
          isRedirected: redirected ? "Yes" : "No",
          redirectType: redirectType || "",
          finalUrl,
          finalStatus: status,
          ga4Present: "Not Checked",
          gtmPresent: "Not Checked",
          facebookPixelPresent: "Not Checked",
        });
        continue;
      }
    }

    const page = await context.newPage();
    const row  = { seedUrl: url, isRedirected: redirected ? "Yes" : "No", redirectType: redirectType || "", finalUrl, finalStatus: status };

    try {
      const res  = await page.goto(finalUrl, { timeout: 30000, waitUntil: bypassHttpCheck ? "networkidle" : "domcontentloaded" });
      if (bypassHttpCheck) { await page.waitForTimeout(2000); row.finalStatus = res?.status() ?? -1; }

      // ── Analytics / Tracking Detection ───────────────────────────────────
      // Give deferred / window.onload-injected tags (a common pattern for a
      // *second* GTM container or a gtag script created via createElement on
      // 'load') a brief window to execute before we inspect the page.
      if (!bypassHttpCheck) {
        await page.waitForTimeout(1500).catch(() => {});
      }

      const trackingData = await page.evaluate(() => {
        const ga4Ids    = new Set();
        const gtmIds    = new Set();
        const fbPixelIds = new Set();
        let ga4PropertyId = null;

        // 1) GA4 measurement IDs surfaced via dataLayer config calls that have
        //    actually fired by the time we inspect the page.
        try {
          if (window.dataLayer && Array.isArray(window.dataLayer)) {
            for (const event of window.dataLayer) {
              if (event && event['measurement_id']) ga4Ids.add(event['measurement_id']);
              if (event && event['config']) {
                Object.keys(event.config).forEach(key => {
                  if (key.startsWith('G-')) ga4Ids.add(key);
                });
              }
            }
          }
        } catch (e) { /* ignore */ }

        // 2) Runtime GTM containers that have actually loaded — this catches
        //    containers injected on window 'load' (not just inline at parse time),
        //    including sites that run more than one GTM container simultaneously.
        try {
          if (window.google_tag_manager) {
            Object.keys(window.google_tag_manager).forEach(key => {
              if (/^GTM-/i.test(key)) gtmIds.add(key.toUpperCase());
            });
          }
        } catch (e) { /* ignore */ }

        // 3) Runtime Facebook Pixel IDs (if fbq has initialised by now)
        try {
          const fbInstance = window.fbq || window._fbq;
          if (fbInstance && typeof fbInstance.getState === "function") {
            const state = fbInstance.getState();
            if (state && Array.isArray(state.pixels)) {
              state.pixels.forEach(p => { if (p && p.id) fbPixelIds.add(String(p.id)); });
            }
          }
        } catch (e) { /* ignore */ }

        // 4) Scan every <script> tag (inline body + src) for literal IDs.
        //    This still works even if a script hasn't executed yet, because we
        //    are only pattern-matching source text, not requiring execution —
        //    it also multi-matches, so more than one GTM/GA4 ID on a page is
        //    captured rather than only the last one seen.
        const scripts = Array.from(document.querySelectorAll('script')).map(s => `${s.innerHTML || ''}\n${s.src || ''}`);
        const allScriptText = scripts.join('\n');

        for (const script of scripts) {
          const ga4Matches = script.match(/G-[A-Z0-9]{8,10}/g) || [];
          ga4Matches.forEach(m => ga4Ids.add(m));

          const gtmMatches = script.match(/GTM-[A-Z0-9]+/ig) || [];
          gtmMatches.forEach(m => gtmIds.add(m.toUpperCase()));

          const fbPixelMatches = [...script.matchAll(/fbq\(\s*['"]init['"]\s*,\s*['"](\d+)['"]/ig)];
          fbPixelMatches.forEach(m => { if (m[1]) fbPixelIds.add(m[1]); });
        }

        // 5) Fallback: Facebook Pixel <noscript> tracking pixel
        //    <img src="https://www.facebook.com/tr?id=XXXXXXXXXX&ev=PageView...">
        try {
          document.querySelectorAll('noscript').forEach(ns => {
            const match = ns.innerHTML.match(/facebook\.com\/tr\?id=(\d+)/i);
            if (match && match[1]) fbPixelIds.add(match[1]);
          });
        } catch (e) { /* ignore */ }

        // Check meta tags / attributes for a GA4 property id
        const meta = document.querySelector('meta[property="google-analytics"]');
        if (meta) ga4PropertyId = meta.getAttribute('content');

        const ga4Present         = ga4Ids.size > 0 || scripts.some(s => s.includes('gtag'));
        const gtmPresent         = gtmIds.size > 0;
        const facebookPixelPresent = fbPixelIds.size > 0;

        return {
          ga4Present,
          measurementId: ga4Ids.size ? [...ga4Ids].join(', ') : "Not Found",
          propertyId: ga4PropertyId || "Not Found",
          gtmPresent,
          gtmId: gtmIds.size ? [...gtmIds].join(', ') : "Not Found",
          facebookPixelPresent,
          facebookPixelId: fbPixelIds.size ? [...fbPixelIds].join(', ') : "Not Found",
        };
      });

      row.ga4Present = trackingData.ga4Present ? "Yes" : "No";
      row.ga4MeasurementId = trackingData.measurementId;
      row.ga4PropertyId = trackingData.propertyId;
      row.gtmPresent = trackingData.gtmPresent ? "Yes" : "No";
      row.gtmId = trackingData.gtmId;
      row.facebookPixelPresent = trackingData.facebookPixelPresent ? "Yes" : "No";
      row.facebookPixelId = trackingData.facebookPixelId;

      row.metaTitle    = await page.$eval("head > title", el => el.innerText.trim()).catch(() => "Not Found");
      row.metaTitleLen = row.metaTitle !== "Not Found" ? row.metaTitle.length : 0;

      row.metaDesc    = await page.$eval("meta[name='description']", el => el.getAttribute("content")?.trim() ?? "").catch(() => "Not Found");
      row.metaDescLen = row.metaDesc !== "Not Found" ? row.metaDesc.length : 0;

      row.canonical = await page.$eval("link[rel='canonical']", el => el.href).catch(() => "Not Found");

      const h1s = await page.$$eval("h1", els => els.map(e => e.innerText.trim().replace(/\s+/g, " ")));
      row.h1      = h1s.length ? h1s.join("\n") : "Not Found";
      row.h1Count = h1s.length;

      const h2s = await page.$$eval("h2", els => els.map(e => e.innerText.trim().replace(/\s+/g, " ")));
      row.h2      = h2s.length ? h2s.join("\n") : "Not Found";
      row.h2Count = h2s.length;

      row.ogTitle       = await page.$eval("meta[property='og:title']",       el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogDescription = await page.$eval("meta[property='og:description']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogImage       = await page.$eval("meta[property='og:image']",       el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.ogUrl         = await page.$eval("meta[property='og:url']",         el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      row.twitterCard  = await page.$eval("meta[name='twitter:card'],meta[property='twitter:card']",   el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.twitterTitle = await page.$eval("meta[name='twitter:title'],meta[property='twitter:title']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");
      row.twitterDesc  = await page.$eval("meta[name='twitter:description'],meta[property='twitter:description']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      row.robotsMeta = await page.$eval("meta[name='robots']", el => el.getAttribute("content") ?? "").catch(() => "Not Found");

      const schemas = await page.$$eval("script[type='application/ld+json']", els => els.map(e => e.innerText.trim()));
      row.schema = schemas.length ? schemas.join("\n---\n") : "Not Found";

    } catch (e) {
      row.error = e.message;
      row.ga4Present = "Error";
      row.ga4MeasurementId = "";
      row.ga4PropertyId = "";
      row.gtmPresent = "Error";
      row.gtmId = "";
      row.facebookPixelPresent = "Error";
      row.facebookPixelId = "";
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
    process.stdout.write(`\r  [Alt] ${i + 1}/${urls.length}: ${url.slice(0, 70)}   `);

    const page = await context.newPage();
    try {
      await page.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });
      await page.evaluate(() => {
        document.querySelectorAll("img").forEach(img => img.scrollIntoView({ block: "center" }));
      });
      await page.waitForTimeout(1500);
      const images = await page.$$eval("img", imgs =>
        imgs.map(img => {
          const resolveUrl = value => {
            try { return value ? new URL(value, document.baseURI).href : ""; } catch { return ""; }
          };
          const candidates = [
            img.getAttribute("data-src"),
            img.getAttribute("data-lazy-src"),
            img.getAttribute("data-original"),
            img.currentSrc,
            img.src,
          ].map(resolveUrl).filter(Boolean);
          const realSrc = candidates.find(src => !/(?:^|\/)empty\.webp(?:[?#]|$)/i.test(src)) ?? "";
          return {
            src: realSrc,
            alt: img.getAttribute("alt"),
            width: img.naturalWidth,
            height: img.naturalHeight,
          };
        })
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

        const isInternal = isSameHost(norm, baseDomain);
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
//  STEP 5 — AI Visibility Audit (Layers 1–11)
// ══════════════════════════════════════════════════════════════════════════════

// AI bots to test for robots.txt access (Layer 1)
const AI_BOTS = [
  { name: "GPTBot",           ua: "GPTBot" },
  { name: "Google-Extended",  ua: "Google-Extended" },
  { name: "ClaudeBot",        ua: "anthropic-ai" },
  { name: "PerplexityBot",    ua: "PerplexityBot" },
  { name: "CCBot",            ua: "CCBot" },
];

/**
 * Fetch robots.txt once per domain and return parsed rules.
 * Returns a map: { userAgent (lowercase) → [{ allow, path }] }
 */
async function fetchRobotsTxt(baseUrl) {
  const robotsUrl = new URL("/robots.txt", baseUrl).href;
  return new Promise((resolve) => {
    const parsed = new URL(robotsUrl);
    const lib    = parsed.protocol === "https:" ? https : http;
    let body = "";
    const req = lib.get(
      { hostname: parsed.hostname, path: "/robots.txt",
        headers: { "User-Agent": "SEO-Audit-Bot/1.0" }, timeout: 10000, rejectUnauthorized: false },
      (res) => {
        res.on("data", d => { body += d; });
        res.on("end", () => resolve(parseRobotsTxt(body)));
      }
    );
    req.on("error", () => resolve({}));
    req.on("timeout", () => { req.destroy(); resolve({}); });
  });
}

function parseRobotsTxt(text) {
  const rules = {};
  let currentAgents = [];
  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) { if (!line) currentAgents = []; continue; }
    const [field, ...rest] = line.split(":");
    const key   = field.trim().toLowerCase();
    const value = rest.join(":").trim();
    if (key === "user-agent") { currentAgents.push(value.toLowerCase()); }
    else if (key === "disallow" || key === "allow") {
      for (const agent of currentAgents) {
        if (!rules[agent]) rules[agent] = [];
        // BUGFIX: Empty Disallow: or Allow: means "no restriction", NOT "block everything"
        // Only add rule if path is non-empty; empty path is treated as no rule
        if (value) {
          rules[agent].push({ type: key, path: value });
        }
      }
    }
  }
  return rules;
}

function isBotAllowedByRobots(robotsRules, botUa, urlPath) {
  const ua = botUa.toLowerCase();
  // Check specific UA first, then wildcard
  for (const agent of [ua, "*"]) {
    const agentRules = robotsRules[agent];
    if (!agentRules) continue;
    // Sort: longer (more specific) paths first
    const sorted = [...agentRules].sort((a, b) => b.path.length - a.path.length);
    for (const rule of sorted) {
      const rpath = rule.path.replace(/\*/g, ".*");
      try {
        if (new RegExp("^" + rpath).test(urlPath)) {
          return rule.type === "allow";
        }
      } catch { continue; }
    }
  }
  return true; // No matching rule = allowed
}

/**
 * Layer 2: Fetch HTTP headers for a URL and check X-Robots-Tag + CSP
 */
async function checkHeaders(url) {
  const { headers, status } = await getStatusAndFinalUrl(url, true);
  const xRobots = headers["x-robots-tag"] ?? "";
  const csp     = headers["content-security-policy"] ?? "";
  return {
    httpStatus      : status,
    xRobotsTag      : xRobots || "None",
    xRobotsNoindex  : /noindex/i.test(xRobots),
    xRobotsNofollow : /nofollow/i.test(xRobots),
    cspPresent      : csp !== "",
    cspBlocksScript : /script-src[^;]*(none|'self'[^;]*$)/i.test(csp),
  };
}

/**
 * Layer 3: Compare raw HTML length vs rendered DOM length → JS dependency
 * Layer 4: Compare normal UA vs bot UA content → cloaking detection
 */
async function checkRenderAndParity(url, browser) {
  // Raw HTML fetch (no JS, no browser)
  const rawHtml = await new Promise((resolve) => {
    const parsed = new URL(url);
    const lib    = parsed.protocol === "https:" ? https : http;
    let body = "";
    const req = lib.get(
      { hostname: parsed.hostname, path: parsed.pathname + parsed.search,
        headers: { "User-Agent": "Mozilla/5.0 (SEO-Audit-Bot/1.0)" },
        timeout: 15000, rejectUnauthorized: false },
      (res) => {
        res.setEncoding("utf8");
        res.on("data", d => { body += d; });
        res.on("end", () => resolve(body));
      }
    );
    req.on("error", () => resolve(""));
    req.on("timeout", () => { req.destroy(); resolve(""); });
  });

  // Rendered DOM (full JS via Playwright — normal UA)
  const normalCtx = await browser.newContext({
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/114.0 Safari/537.36",
    ignoreHTTPSErrors: true,
  });
  const page1 = await normalCtx.newPage();
  let renderedText = "";
  try {
    await page1.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });
    renderedText = await page1.evaluate(() => document.body?.innerText ?? "");
  } catch { renderedText = ""; }
  finally { await page1.close(); await normalCtx.close(); }

  // Bot UA fetch (Layer 4 — cloaking detection)
  const botCtx  = await browser.newContext({
    userAgent: "GPTBot/1.0 (+https://openai.com/gptbot)",
    ignoreHTTPSErrors: true,
  });
  const page2 = await botCtx.newPage();
  let botText = "";
  try {
    await page2.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });
    botText = await page2.evaluate(() => document.body?.innerText ?? "");
  } catch { botText = ""; }
  finally { await page2.close(); await botCtx.close(); }

  // JS dependency: raw HTML << rendered
  const rawLen      = rawHtml.length;
  const renderedLen = renderedText.length;
  const jsDependencyRatio = rawLen > 0 ? renderedLen / rawLen : 1;
  const isJsDependent     = jsDependencyRatio > 3; // rendered is 3x+ bigger than raw

  // Cloaking: compare normal UA vs bot UA text similarity
  const similarity  = computeSimilarity(renderedText, botText);
  const isCloaking  = similarity < 0.7 && botText.length > 50;

  return {
    rawHtmlLength    : rawLen,
    renderedLength   : renderedLen,
    jsDependencyRatio: parseFloat(jsDependencyRatio.toFixed(2)),
    isJsDependent,
    botTextLength    : botText.length,
    contentSimilarity: parseFloat(similarity.toFixed(2)),
    isCloaking,
  };
}

/** Simple word-overlap similarity score (0–1) */
function computeSimilarity(textA, textB) {
  if (!textA || !textB) return 0;
  const wordsA = new Set(textA.toLowerCase().split(/\s+/).filter(w => w.length > 3));
  const wordsB = new Set(textB.toLowerCase().split(/\s+/).filter(w => w.length > 3));
  if (wordsA.size === 0 || wordsB.size === 0) return 0;
  let overlap = 0;
  for (const w of wordsA) { if (wordsB.has(w)) overlap++; }
  return overlap / Math.max(wordsA.size, wordsB.size);
}

/**
 * Layers 6–8, 11: Content analysis from rendered page text + DOM
 * Checks: word count, thin content, heading structure, extractability, entities
 */
async function analyzeContent(url, browser, seoRow) {
  const context = await browser.newContext({
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/114.0 Safari/537.36",
    ignoreHTTPSErrors: true,
  });
  const page = await context.newPage();
  let result = {};

  try {
    await page.goto(url, { timeout: 30000, waitUntil: "domcontentloaded" });

    result = await page.evaluate(() => {
      // ── Layer 6: Content presence ──────────────────────────────────────────
      const bodyText    = document.body?.innerText ?? "";
      const words       = bodyText.split(/\s+/).filter(w => w.length > 1);
      const wordCount   = words.length;

      // Estimate boilerplate (nav + header + footer) vs main content
      const main        = document.querySelector("main, article, [role='main'], #content, .content, .main");
      const mainText    = main ? main.innerText : bodyText;
      const mainWords   = mainText.split(/\s+/).filter(w => w.length > 1).length;
      const boilerRatio = wordCount > 0 ? parseFloat(((wordCount - mainWords) / wordCount).toFixed(2)) : 0;

      // ── Layer 7: Extractability ────────────────────────────────────────────
      const hasFaq        = document.querySelectorAll("[itemtype*='FAQPage'], .faq, .FAQ, details, summary").length > 0;
      const hasLists      = document.querySelectorAll("ul li, ol li").length > 5;
      const hasTables     = document.querySelectorAll("table").length > 0;
      const hasDefinitions = document.querySelectorAll("dl, dt, dd").length > 0;
      // Short answer paragraph: first <p> under 200 chars
      const firstPara     = document.querySelector("p");
      const hasShortAnswer = firstPara && firstPara.innerText.trim().length < 200 && firstPara.innerText.trim().length > 30;

      // ── Layer 8: Structure & chunking ──────────────────────────────────────
      const h1Count = document.querySelectorAll("h1").length;
      const h2Count = document.querySelectorAll("h2").length;
      const h3Count = document.querySelectorAll("h3").length;
      const paras   = Array.from(document.querySelectorAll("p")).map(p => p.innerText.trim().split(/\s+/).length);
      const avgParaWords = paras.length ? Math.round(paras.reduce((a, b) => a + b, 0) / paras.length) : 0;
      const hasSemanticHtml = !!(
        document.querySelector("article") ||
        document.querySelector("main") ||
        document.querySelector("section") ||
        document.querySelector("aside")
      );
      const longWallsOfText = paras.filter(w => w > 150).length;

      // ── Layer 11: Entity signals (keyword frequency heuristic) ─────────────
      // Top 5 repeated meaningful words (length > 4, not common stopwords)
      const STOPWORDS = new Set(["about","after","again","also","been","before","being","between","both","could","does","each","from","have","here","into","just","like","more","most","much","need","only","other","over","same","some","such","than","that","their","them","then","there","these","they","this","through","under","until","very","were","what","when","where","which","while","will","with","would","your"]);
      const freq = {};
      for (const w of words) {
        const clean = w.toLowerCase().replace(/[^a-z]/g, "");
        if (clean.length > 4 && !STOPWORDS.has(clean)) {
          freq[clean] = (freq[clean] ?? 0) + 1;
        }
      }
      const topEntities = Object.entries(freq)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([word, count]) => `${word}(${count})`)
        .join(", ");

      return {
        wordCount, mainWords, boilerRatio,
        hasFaq, hasLists, hasTables, hasDefinitions, hasShortAnswer,
        h1Count, h2Count, h3Count,
        avgParaWords, hasSemanticHtml, longWallsOfText,
        topEntities,
        extractabilityScore: (hasFaq ? 1 : 0) + (hasLists || hasTables ? 1 : 0) + (hasShortAnswer ? 1 : 0) + (hasDefinitions ? 1 : 0),
      };
    });

  } catch (e) {
    result = { error: e.message };
  } finally {
    await page.close();
    await context.close();
  }

  return result;
}

/**
 * Master function: run all AI visibility layers for a single URL.
 * Returns a flat result object for use in Sheet 5 and scoring.
 */
async function auditAiVisibility(url, robotsRules, browser, seoRow, linkData) {
  const urlPath = new URL(url).pathname;
  const result  = { url };

  // ── Layer 1: Bot access via robots.txt ──────────────────────────────────────
  const botAccess = {};
  for (const bot of AI_BOTS) {
    botAccess[bot.name] = isBotAllowedByRobots(robotsRules, bot.ua, urlPath) ? "Allowed" : "Blocked";
  }
  result.botGPT        = botAccess["GPTBot"];
  result.botGoogleExt  = botAccess["Google-Extended"];
  result.botClaude     = botAccess["ClaudeBot"];
  result.botPerplexity = botAccess["PerplexityBot"];
  result.botCCBot      = botAccess["CCBot"];
  const blockedCount   = Object.values(botAccess).filter(v => v === "Blocked").length;
  result.accessStatus  = blockedCount === 5 ? "Blocked" : blockedCount > 0 ? "Partial" : "Allowed";

  // ── Layer 2: Headers & CSP ─────────────────────────────────────────────────
  const headers          = await checkHeaders(url);
  result.xRobotsTag      = headers.xRobotsTag;
  result.xRobotsNoindex  = headers.xRobotsNoindex  ? "Yes" : "No";
  result.cspPresent      = headers.cspPresent       ? "Yes" : "No";
  result.cspBlocksScript = headers.cspBlocksScript  ? "⚠ Yes" : "No";

  // ── Layer 3 & 4: Render + Cloaking ─────────────────────────────────────────
  const render              = await checkRenderAndParity(url, browser);
  result.rawHtmlLen         = render.rawHtmlLength;
  result.renderedLen        = render.renderedLength;
  result.jsDependencyRatio  = render.jsDependencyRatio;
  result.isJsDependent      = render.isJsDependent   ? "⚠ Yes" : "No";
  result.contentSimilarity  = render.contentSimilarity;
  result.isCloaking         = render.isCloaking       ? "⚠ Yes" : "No";

  // ── Layer 5: Indexability signals (from existing seoRow data) ───────────────
  const robotsMeta       = (seoRow?.robotsMeta ?? "").toLowerCase();
  const hasNoindex       = /noindex/.test(robotsMeta) || headers.xRobotsNoindex;
  const canonicalMatch   = !seoRow?.canonical || seoRow.canonical === "Not Found" || seoRow.canonical === url;
  result.metaRobotsNoindex = hasNoindex                ? "⚠ Noindex" : "OK";
  result.canonicalOk       = canonicalMatch            ? "OK" : "⚠ Mismatch";

  // ── Layer 6–8 & 11: Content analysis ───────────────────────────────────────
  const content = await analyzeContent(url, browser, seoRow);
  result.wordCount        = content.wordCount    ?? 0;
  result.mainWords        = content.mainWords    ?? 0;
  result.boilerRatio      = content.boilerRatio  ?? 0;
  result.thinContent      = (content.wordCount ?? 0) < 300 ? "⚠ Thin" : "OK";
  result.hasFaq           = content.hasFaq       ? "Yes" : "No";
  result.hasLists         = content.hasLists     ? "Yes" : "No";
  result.hasTables        = content.hasTables    ? "Yes" : "No";
  result.hasShortAnswer   = content.hasShortAnswer ? "Yes" : "No";
  result.extractScore     = content.extractabilityScore ?? 0;
  result.h2Count          = content.h2Count      ?? 0;
  result.h3Count          = content.h3Count      ?? 0;
  result.avgParaWords     = content.avgParaWords  ?? 0;
  result.hasSemanticHtml  = content.hasSemanticHtml ? "Yes" : "No";
  result.longWalls        = content.longWallsOfText ?? 0;
  result.topEntities      = content.topEntities   ?? "";

  // ── Layer 9: Context signals (from existing linkData) ──────────────────────
  const inboundLinks  = linkData.filter(l => l.linkUrl === url && l.type === "Internal").length;
  const outboundLinks = linkData.filter(l => l.pageUrl === url && l.type === "Internal").length;
  result.inboundLinks  = inboundLinks;
  result.outboundLinks = outboundLinks;
  result.isOrphan      = inboundLinks === 0 ? "⚠ Orphan" : "OK";

  // ── Layer 10: Structured data (from existing seoRow data) ──────────────────
  const schema          = seoRow?.schema ?? "Not Found";
  const schemaTypes     = schema !== "Not Found" ? extractSchemaTypes(schema) : [];
  result.schemaTypes    = schemaTypes.join(", ") || "None";
  result.hasFaqSchema   = schemaTypes.some(t => /faq/i.test(t))       ? "Yes" : "No";
  result.hasArticleSchema = schemaTypes.some(t => /article|news|blog/i.test(t)) ? "Yes" : "No";
  result.schemaScore    = schemaTypes.length;

  // ── Final scoring ───────────────────────────────────────────────────────────
  result.accessScore  = scoreAccess(result);
  result.renderScore  = scoreRender(result);
  result.contentScore = scoreContent(result);
  result.structScore  = scoreStructure(result);
  result.entityScore  = scoreEntity(result);
  result.extractScore = scoreExtract(result);
  result.finalScore   = Math.round(
    result.accessScore  * 0.20 +
    result.renderScore  * 0.15 +
    result.contentScore * 0.20 +
    result.structScore  * 0.15 +
    result.entityScore  * 0.10 +
    result.extractScore * 0.20
  );
  result.scoreLabel = result.finalScore >= 75 ? "Good" : result.finalScore >= 50 ? "Needs Work" : "Critical";

  return result;
}

function extractSchemaTypes(schemaStr) {
  const types = [];
  try {
    const blocks = schemaStr.split("\n---\n");
    for (const block of blocks) {
      const parsed = JSON.parse(block);
      const type   = parsed["@type"] ?? parsed.type ?? "";
      if (type) types.push(Array.isArray(type) ? type.join(", ") : type);
    }
  } catch {}
  return types;
}

// ── Scoring functions (each returns 0–100) ────────────────────────────────────

function scoreAccess(r) {
  if (r.accessStatus === "Blocked") return 0;
  if (r.accessStatus === "Partial") return 50;
  if (r.metaRobotsNoindex === "⚠ Noindex") return 30;
  if (r.xRobotsNoindex === "Yes") return 30;
  if (r.canonicalOk !== "OK") return 70;
  return 100;
}

function scoreRender(r) {
  let score = 100;
  if (r.isJsDependent === "⚠ Yes") score -= 50;
  if (r.isCloaking    === "⚠ Yes") score -= 30;
  if (r.cspBlocksScript === "⚠ Yes") score -= 20;
  return Math.max(0, score);
}

function scoreContent(r) {
  let score = 0;
  if (r.wordCount >= 300)  score += 40;
  if (r.wordCount >= 600)  score += 20;
  if (r.wordCount >= 1000) score += 10;
  if (r.boilerRatio < 0.5) score += 20;
  if (r.mainWords >= 200)  score += 10;
  return Math.min(100, score);
}

function scoreStructure(r) {
  let score = 0;
  if (r.h2Count  >= 2)  score += 30;
  if (r.h3Count  >= 1)  score += 10;
  if (r.avgParaWords > 0 && r.avgParaWords <= 100) score += 20;
  if (r.hasSemanticHtml === "Yes") score += 20;
  if (r.longWalls === 0) score += 20;
  return Math.min(100, score);
}

function scoreEntity(r) {
  const entityCount = r.topEntities ? r.topEntities.split(",").length : 0;
  if (entityCount >= 5) return 100;
  if (entityCount >= 3) return 70;
  if (entityCount >= 1) return 40;
  return 0;
}

function scoreExtract(r) {
  let score = 0;
  if (r.hasShortAnswer === "Yes") score += 25;
  if (r.hasFaq  === "Yes")        score += 20;
  if (r.hasLists === "Yes")       score += 20;
  if (r.hasTables === "Yes")      score += 15;
  if (r.schemaScore > 0)          score += 20;
  return Math.min(100, score);
}

/**
 * Orchestrator: runs AI audit for all URLs in a site
 */
async function runAiVisibilityAudit(siteUrl, crawlData, seoData, linkData, browser) {
  console.log(`\n  🤖  STEP 5 — Running AI Visibility Audit (${crawlData.urls.length} pages)...`);

  // Fetch robots.txt once per domain
  const robotsRules = await fetchRobotsTxt(siteUrl);

  const results = [];
  for (let i = 0; i < crawlData.urls.length; i++) {
    const url    = crawlData.urls[i];
    const seoRow = seoData.find(s => s.url === url || s.finalUrl === url) ?? {};
    process.stdout.write(`\r  [AI Audit] ${i + 1}/${crawlData.urls.length}: ${url.slice(0, 60)}   `);
    try {
      const result = await auditAiVisibility(url, robotsRules, browser, seoRow, linkData);
      results.push(result);
    } catch (e) {
      results.push({ url, error: e.message, finalScore: 0, scoreLabel: "Error" });
    }
  }

  console.log(`\n  ✅ AI Visibility audit done — ${results.length} pages scored`);
  return results;
}

// ══════════════════════════════════════════════════════════════════════════════
//  EXCEL WRITER — 5 sheets + Summary
// ══════════════════════════════════════════════════════════════════════════════
async function buildExcel(siteName, crawlData, seoData, altData, linkData, aiData) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "SEO-Audit-Tool";
  wb.created = new Date();

  // ── Summary Sheet FIRST so it appears as tab 1 ────────────────────────────
  const wss = wb.addWorksheet("Summary");
  wss.columns = [{ key: "label", width: 35 }, { key: "value", width: 40 }];

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
    // ── Redirect tracking ──
    { header: "Seed URL",            key: "seedUrl",          width: 55 },
    { header: "Is Redirected?",      key: "isRedirected",     width: 15 },
    { header: "Redirect Type",       key: "redirectType",     width: 25 },
    { header: "Final URL",           key: "finalUrl",         width: 55 },
    { header: "Final Status Code",   key: "finalStatus",      width: 16 },
    // ── GA4 / GTM / Pixel Analytics ──
    { header: "GA4 Present?",          key: "ga4Present",           width: 14 },
    { header: "GA4 Measurement ID",    key: "ga4MeasurementId",     width: 25 },
    { header: "GA4 Property ID",       key: "ga4PropertyId",        width: 25 },
    { header: "GTM Present?",          key: "gtmPresent",           width: 14 },
    { header: "GTM ID",                key: "gtmId",                width: 20 },
    { header: "Facebook Pixel Present?", key: "facebookPixelPresent", width: 20 },
    { header: "Facebook Pixel ID",     key: "facebookPixelId",      width: 22 },
    // ── SEO metadata ──
    { header: "Meta Title",          key: "metaTitle",        width: 50 },
    { header: "Title Length",        key: "metaTitleLen",     width: 13 },
    { header: "Meta Description",    key: "metaDesc",         width: 60 },
    { header: "Desc Length",         key: "metaDescLen",      width: 12 },
    { header: "Canonical",           key: "canonical",        width: 55 },
    { header: "H1 Tag(s)",           key: "h1",               width: 50 },
    { header: "H1 Count",            key: "h1Count",          width: 10 },
    { header: "H2 Tags",             key: "h2",               width: 50 },
    { header: "H2 Count",            key: "h2Count",          width: 10 },
    { header: "OG Title",            key: "ogTitle",          width: 40 },
    { header: "OG Description",      key: "ogDescription",    width: 50 },
    { header: "OG Image",            key: "ogImage",          width: 55 },
    { header: "OG URL",              key: "ogUrl",            width: 55 },
    { header: "Twitter Card",        key: "twitterCard",      width: 20 },
    { header: "Twitter Title",       key: "twitterTitle",     width: 40 },
    { header: "Twitter Desc",        key: "twitterDesc",      width: 50 },
    { header: "Robots Meta",         key: "robotsMeta",       width: 25 },
    { header: "Schema",              key: "schema",           width: 40 },
  ];
  styleHeader(ws2.getRow(1), "FF1F4E79");
  seoData.forEach((d, i) => {
    const r = ws2.addRow({
      seedUrl: d.seedUrl ?? d.url ?? "",
      isRedirected: d.isRedirected ?? "Unknown",
      redirectType: d.redirectType ?? "",
      finalUrl: d.finalUrl ?? d.seedUrl ?? d.url ?? "",
      finalStatus: d.finalStatus ?? d.status ?? "",
      ga4Present: d.ga4Present ?? "Not Checked",
      ga4MeasurementId: d.ga4MeasurementId ?? "",
      ga4PropertyId: d.ga4PropertyId ?? "",
      gtmPresent: d.gtmPresent ?? "Not Checked",
      gtmId: d.gtmId ?? "",
      facebookPixelPresent: d.facebookPixelPresent ?? "Not Checked",
      facebookPixelId: d.facebookPixelId ?? "",
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
    makeHyperlink(r.getCell("seedUrl"), d.seedUrl ?? d.url ?? "");
    if (d.finalUrl && d.finalUrl !== (d.seedUrl ?? d.url)) makeHyperlink(r.getCell("finalUrl"), d.finalUrl);
    
    // Colour-code redirect status
    const redirectCell = r.getCell("isRedirected");
    if (d.isRedirected === "Yes") {
      redirectCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C6500" } };
    }
    
    // Colour-code redirect type
    if (d.redirectType && d.redirectType.includes("301")) {
      r.getCell("redirectType").font = { name: "Arial", size: 10, color: { argb: "FF1E7145" }, bold: true };
    } else if (d.redirectType && d.redirectType.includes("302")) {
      r.getCell("redirectType").font = { name: "Arial", size: 10, color: { argb: "FF9C6500" }, bold: true };
    }
    
    // Colour-code final status
    const statusCell = r.getCell("finalStatus");
    if (d.finalStatus === 200) {
      statusCell.font = { name: "Arial", size: 10, color: { argb: "FF1E7145" }, bold: true };
    } else if ([301, 302, 303, 307, 308].includes(d.finalStatus)) {
      statusCell.font = { name: "Arial", size: 10, color: { argb: "FF9C6500" }, bold: true };
    } else if (d.finalStatus >= 400) {
      statusCell.font = { name: "Arial", size: 10, color: { argb: "FF9C0006" }, bold: true };
    }
    
    // Colour-code GA4 / GTM / Pixel status
    const ga4Cell = r.getCell("ga4Present");
    if (d.ga4Present === "Yes") {
      ga4Cell.font = { name: "Arial", size: 10, color: { argb: "FF1E7145" }, bold: true };
    } else if (d.ga4Present === "No") {
      ga4Cell.font = { name: "Arial", size: 10, color: { argb: "FF9C0006" }, bold: true };
    }

    const gtmCell = r.getCell("gtmPresent");
    if (d.gtmPresent === "Yes") {
      gtmCell.font = { name: "Arial", size: 10, color: { argb: "FF1E7145" }, bold: true };
    } else if (d.gtmPresent === "No") {
      gtmCell.font = { name: "Arial", size: 10, color: { argb: "FF9C0006" }, bold: true };
    }

    const pixelCell = r.getCell("facebookPixelPresent");
    if (d.facebookPixelPresent === "Yes") {
      pixelCell.font = { name: "Arial", size: 10, color: { argb: "FF1E7145" }, bold: true };
    } else if (d.facebookPixelPresent === "No") {
      pixelCell.font = { name: "Arial", size: 10, color: { argb: "FF9C0006" }, bold: true };
    }
    
    r.getCell("h1").alignment = { vertical: "middle", wrapText: true };
    r.getCell("h2").alignment = { vertical: "middle", wrapText: true };

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

  // ── Sheet 3: Alt Tags ─────────────────────────────────────────────────────
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

  // ── Sheet 5: AI Visibility ────────────────────────────────────────────────
  const ws5 = wb.addWorksheet("Sheet5 - AI Visibility");
  ws5.columns = [
    // Identity
    { header: "URL",                  key: "url",             width: 55 },
    // Layer 1 — Bot Access
    { header: "Access Status",        key: "accessStatus",    width: 14 },
    { header: "GPTBot",               key: "botGPT",          width: 12 },
    { header: "Google-Extended",      key: "botGoogleExt",    width: 18 },
    { header: "ClaudeBot",            key: "botClaude",       width: 13 },
    { header: "PerplexityBot",        key: "botPerplexity",   width: 16 },
    { header: "CCBot",                key: "botCCBot",        width: 12 },
    // Layer 2 — Headers
    { header: "X-Robots-Tag",         key: "xRobotsTag",      width: 20 },
    { header: "X-Robots Noindex",     key: "xRobotsNoindex",  width: 18 },
    { header: "CSP Present",          key: "cspPresent",      width: 13 },
    { header: "CSP Blocks Scripts",   key: "cspBlocksScript", width: 20 },
    // Layer 3 — JS Dependency
    { header: "Raw HTML (chars)",     key: "rawHtmlLen",      width: 17 },
    { header: "Rendered (chars)",     key: "renderedLen",     width: 17 },
    { header: "JS Ratio",             key: "jsDependencyRatio", width: 12 },
    { header: "JS Dependent?",        key: "isJsDependent",   width: 15 },
    // Layer 4 — Cloaking
    { header: "Content Similarity",   key: "contentSimilarity", width: 18 },
    { header: "Cloaking Risk?",       key: "isCloaking",      width: 16 },
    // Layer 5 — Indexability
    { header: "Meta Robots",          key: "metaRobotsNoindex", width: 15 },
    { header: "Canonical OK?",        key: "canonicalOk",     width: 15 },
    // Layer 6 — Content
    { header: "Word Count",           key: "wordCount",       width: 13 },
    { header: "Main Content Words",   key: "mainWords",       width: 18 },
    { header: "Boilerplate Ratio",    key: "boilerRatio",     width: 18 },
    { header: "Thin Content?",        key: "thinContent",     width: 15 },
    // Layer 7 — Extractability
    { header: "Has FAQ?",             key: "hasFaq",          width: 12 },
    { header: "Has Lists?",           key: "hasLists",        width: 13 },
    { header: "Has Tables?",          key: "hasTables",       width: 13 },
    { header: "Short Answer?",        key: "hasShortAnswer",  width: 15 },
    // Layer 8 — Structure
    { header: "H2 Count",             key: "h2Count",         width: 12 },
    { header: "H3 Count",             key: "h3Count",         width: 12 },
    { header: "Avg Para Words",       key: "avgParaWords",    width: 16 },
    { header: "Semantic HTML?",       key: "hasSemanticHtml", width: 16 },
    { header: "Long Text Blocks",     key: "longWalls",       width: 18 },
    // Layer 9 — Context
    { header: "Inbound Links",        key: "inboundLinks",    width: 15 },
    { header: "Outbound Internal",    key: "outboundLinks",   width: 18 },
    { header: "Orphan Page?",         key: "isOrphan",        width: 14 },
    // Layer 10 — Schema
    { header: "Schema Types",         key: "schemaTypes",     width: 30 },
    { header: "FAQ Schema?",          key: "hasFaqSchema",    width: 14 },
    { header: "Article Schema?",      key: "hasArticleSchema",width: 16 },
    // Layer 11 — Entities
    { header: "Top Entities",         key: "topEntities",     width: 40 },
    // Scores
    { header: "Access Score",         key: "accessScore",     width: 14 },
    { header: "Render Score",         key: "renderScore",     width: 14 },
    { header: "Content Score",        key: "contentScore",    width: 14 },
    { header: "Structure Score",      key: "structScore",     width: 15 },
    { header: "Entity Score",         key: "entityScore",     width: 14 },
    { header: "Extract Score",        key: "extractScore",    width: 14 },
    { header: "AI Readiness Score",   key: "finalScore",      width: 20 },
    { header: "Rating",               key: "scoreLabel",      width: 14 },
  ];
  styleHeader(ws5.getRow(1), "FF4B0082");

  aiData.forEach((d, i) => {
    const r = ws5.addRow(d);
    styleDataRow(r, i);
    makeHyperlink(r.getCell("url"), d.url);

    // Colour-code access status
    const accessCell = r.getCell("accessStatus");
    if (d.accessStatus === "Blocked") accessCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C0006" } };
    else if (d.accessStatus === "Partial") accessCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C6500" } };

    // Colour-code each bot column
    for (const col of ["botGPT","botGoogleExt","botClaude","botPerplexity","botCCBot"]) {
      const c = r.getCell(col);
      if (d[col] === "Blocked") c.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C0006" } };
    }

    // Flag warning cells
    for (const col of ["isJsDependent","isCloaking","cspBlocksScript","thinContent","isOrphan","metaRobotsNoindex"]) {
      if (String(d[col]).startsWith("⚠")) {
        r.getCell(col).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C6500" } };
      }
    }

    // Colour all score cells
    for (const col of ["accessScore","renderScore","contentScore","structScore","entityScore","extractScore"]) {
      colourScoreCell(r.getCell(col), d[col] ?? 0);
    }

    // Final score — prominent
    const finalCell = r.getCell("finalScore");
    colourScoreCell(finalCell, d.finalScore ?? 0);
    finalCell.font = { ...finalCell.font, size: 11, bold: true };

    const labelCell = r.getCell("scoreLabel");
    if (d.scoreLabel === "Good")       labelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF1E7145" } };
    else if (d.scoreLabel === "Needs Work") labelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C6500" } };
    else                               labelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF9C0006" } };
  });
  freezeAndFilter(ws5, ws5.columns.length);

  // ── Summary Sheet — fill data ─────────────────────────────────────────────
  const missingAlt      = altData.filter(d => !d.hasAlt).length;
  const brokenLinks     = linkData.filter(d => d.status === 404 || d.status === -1).length;
  const internalLinks   = linkData.filter(d => d.type === "Internal").length;
  const externalLinks   = linkData.filter(d => d.type === "External").length;
  const redirectedPages = seoData.filter(d => d.isRedirected === "Yes").length;
  const redirectedCount = seoData.filter(d => d.isRedirected === "Yes" && d.redirectType).length;
  
  // GA4 Analytics stats
  const ga4Present      = seoData.filter(d => d.ga4Present === "Yes").length;
  const ga4Missing      = seoData.filter(d => d.ga4Present === "No").length;

  // AI summary stats
  const aiBlocked     = aiData.filter(d => d.accessStatus === "Blocked").length;
  const aiJsDependent = aiData.filter(d => d.isJsDependent === "⚠ Yes").length;
  const aiCloaking    = aiData.filter(d => d.isCloaking    === "⚠ Yes").length;
  const aiThin        = aiData.filter(d => d.thinContent   === "⚠ Thin").length;
  const aiOrphan      = aiData.filter(d => d.isOrphan      === "⚠ Orphan").length;
  const avgAiScore    = aiData.length ? Math.round(aiData.reduce((s, d) => s + (d.finalScore ?? 0), 0) / aiData.length) : 0;
  const goodPages     = aiData.filter(d => d.finalScore >= 75).length;
  const criticalPages = aiData.filter(d => d.finalScore < 50).length;

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
    ["📊 Analytics (GA4)",        ""],
    ["  Pages with GA4",          ga4Present],
    ["  Pages without GA4",       ga4Missing],
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
    ["",""],
    ["🤖 AI Visibility",         ""],
    ["  Pages Audited",          aiData.length],
    ["  Avg AI Readiness Score", `${avgAiScore}/100`],
    ["  Good Pages (≥75)",       goodPages],
    ["  Critical Pages (<50)",   criticalPages],
    ["  Bot Access Blocked",     aiBlocked],
    ["  JS-Only (AI invisible)", aiJsDependent],
    ["  Cloaking Risk",          aiCloaking],
    ["  Thin Content (<300 words)", aiThin],
    ["  Orphan Pages",           aiOrphan],
  ];

  rows.forEach(([label, value], i) => {
    const r = wss.addRow({ label, value });
    if (i === 0) {
      r.getCell("label").font = { bold: true, size: 15, name: "Arial", color: { argb: "FF1F3864" } };
      r.eachCell(c => c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } });
      r.height = 30;
    } else if (label.startsWith("📄") || label.startsWith("🔍") || label.startsWith("🖼️") || label.startsWith("🔗") || label.startsWith("🤖")) {
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

  return { wb, aiSummary: { avgAiScore, goodPages, criticalPages, aiBlocked, aiJsDependent, aiThin, aiCloaking, aiOrphan, total: aiData.length } };
}

function findPreviousReport(siteName, currentReportPath) {
  const prefix = `SEO_Audit_${siteName}_`;
  const currentPath = path.resolve(currentReportPath);
  const searchDirs = [
    __dirname,
    path.join(__dirname, "report"),
    path.resolve(__dirname, "report"),
  ].filter((dir, index, arr) => dir && arr.indexOf(dir) === index && fs.existsSync(dir));

  const candidates = [];
  for (const dir of searchDirs) {
    const files = fs.existsSync(dir) ? fs.readdirSync(dir) : [];
    for (const file of files) {
      const fullPath = path.resolve(dir, file);
      if (file.startsWith(prefix) && file.endsWith(".xlsx") && fullPath !== currentPath) {
        candidates.push({
          path: fullPath,
          mtime: fs.statSync(fullPath).mtimeMs,
        });
      }
    }
  }

  candidates.sort((a, b) => b.mtime - a.mtime);
  return candidates.length ? candidates[0].path : null;
}

function displayCellValue(cell) {
  const value = cell?.value;
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") return String(value.text ?? value.hyperlink ?? value.result ?? JSON.stringify(value));
  return String(value);
}

function readSummaryMetrics(workbook) {
  const sheet = workbook.getWorksheet("Summary");
  const metrics = new Map();
  if (!sheet) return metrics;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const label = displayCellValue(row.getCell(1));
    const value = displayCellValue(row.getCell(2));
    if (label && value) metrics.set(label, value);
  }
  return metrics;
}

function readReportRecords(workbook, sheetName, keyHeaders) {
  const sheet = workbook.getWorksheet(sheetName);
  const records = new Map();
  if (!sheet || sheet.rowCount < 2) return records;

  const headers = Array.from({ length: sheet.columnCount }, (_, index) =>
    displayCellValue(sheet.getRow(1).getCell(index + 1)));
  const keyColumns = keyHeaders.map(header => headers.indexOf(header) + 1);
  if (keyColumns.some(column => column === 0)) return records;

  const occurrences = new Map();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const identity = keyColumns.map(column => displayCellValue(row.getCell(column)));
    if (identity.every(value => !value)) continue;

    const baseKey = JSON.stringify(identity);
    const occurrence = (occurrences.get(baseKey) ?? 0) + 1;
    occurrences.set(baseKey, occurrence);
    const values = new Map(headers.map((header, index) =>
      [header, displayCellValue(row.getCell(index + 1))]));
    records.set(`${baseKey}#${occurrence}`, { identity, values });
  }
  return records;
}

function parseMetricNumber(value) {
  const match = String(value).match(/^(-?\d+(?:\.\d+)?)(?:\/\d+)?$/);
  return match ? Number(match[1]) : null;
}

function styleComparisonDataRow(row, index, statusColumn) {
  styleDataRow(row, index);
  row.eachCell(cell => {
    cell.alignment = { vertical: "middle", wrapText: true };
  });
  const statusColors = {
    Added: "FF1E7145",
    Removed: "FF9C0006",
    Updated: "FF9C6500",
    Unchanged: "FF666666",
  };
  const statusCell = row.getCell(statusColumn);
  const color = statusColors[statusCell.value];
  if (color) statusCell.font = { name: "Arial", size: 10, bold: true, color: { argb: color } };
}

async function buildComparisonReport(siteName, previousReportPath, currentReportPath, currentWorkbook) {
  const previousWorkbook = new ExcelJS.Workbook();
  if (previousReportPath) await previousWorkbook.xlsx.readFile(previousReportPath);

  const recordSheets = [
    { name: "Sheet1 - All URLs", keys: ["URL"] },
    { name: "Sheet2 - SEO Data", keys: ["Seed URL"] },
    { name: "Sheet3 - Alt Tags", keys: ["Page URL", "Image URL"] },
    { name: "Sheet4 - Links", keys: ["Found On", "Link URL"] },
    { name: "Sheet5 - AI Visibility", keys: ["URL"] },
  ];
  const datasetComparisons = [];
  const comparisonDetails = [];

  for (const { name, keys } of recordSheets) {
    const previousRecords = readReportRecords(previousWorkbook, name, keys);
    const currentRecords = readReportRecords(currentWorkbook, name, keys);
    const counts = { Added: 0, Removed: 0, Updated: 0, Unchanged: 0 };
    let fieldsChanged = 0;

    const compareRecord = (previous, current) => {
      const fields = new Set([
        ...(previous?.values.keys() ?? []),
        ...(current?.values.keys() ?? []),
      ]);
      let status;
      let changes = [];

      if (!previous) {
        status = "Added";
      } else if (!current) {
        status = "Removed";
      } else {
        changes = [...fields]
          .filter(field => field && field !== "#")
          .map(field => ({
            field,
            previous: previous.values.get(field) ?? "",
            current: current.values.get(field) ?? "",
          }))
          .filter(change => change.previous !== change.current);
        status = changes.length ? "Updated" : "Unchanged";
      }

      counts[status]++;

      if (status === "Updated") {
        const identity = current.identity.join(" | ");
        for (const change of changes) {
          comparisonDetails.push({ section: name, status, identity, ...change });
          fieldsChanged++;
        }
      } else if (status === "Removed" || (status === "Added" && previousReportPath)) {
        const record = current ?? previous;
        comparisonDetails.push({
          section: name,
          status,
          identity: record.identity.join(" | "),
          values: record.values,
        });
        fieldsChanged += [...record.values.keys()].filter(field => field && field !== "#").length;
      }
    };

    for (const [recordId, previous] of previousRecords) {
      compareRecord(previous, currentRecords.get(recordId));
    }
    for (const [recordId, current] of currentRecords) {
      if (!previousRecords.has(recordId)) compareRecord(undefined, current);
    }

    datasetComparisons.push({
      section: name,
      previousCount: previousRecords.size,
      currentCount: currentRecords.size,
      ...counts,
      fieldsChanged,
    });
  }

  const comparisonWorkbook = new ExcelJS.Workbook();
  comparisonWorkbook.creator = "SEO-Audit-Tool";
  comparisonWorkbook.created = new Date();

  const summarySheet = comparisonWorkbook.addWorksheet("Summary Changes");
  summarySheet.columns = [
    { key: "metric", width: 30 },
    { key: "previous", width: 36 },
    { key: "current", width: 55 },
    { key: "change", width: 30 },
    { key: "status", width: 50 },
    { key: "added", width: 12 },
    { key: "removed", width: 12 },
    { key: "updated", width: 12 },
    { key: "unchanged", width: 14 },
    { key: "fieldChanges", width: 16 },
  ];
  summarySheet.mergeCells("A1:J1");
  summarySheet.getCell("A1").value = `Audit Comparison Summary: ${siteName}`;
  summarySheet.getCell("A1").font = { bold: true, size: 15, name: "Arial", color: { argb: "FF1F3864" } };
  summarySheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
  summarySheet.getCell("A1").alignment = { vertical: "middle" };
  summarySheet.getRow(1).height = 30;
  summarySheet.addRow(["Previous report", previousReportPath ? path.basename(previousReportPath) : "Initial baseline (no previous report)"]);
  summarySheet.addRow(["Current report", path.basename(currentReportPath)]);
  summarySheet.addRow([]);
  summarySheet.addRow(["Summary metric", "Previous run", "Current run", "Numeric delta", "Status"]);
  styleHeader(summarySheet.getRow(5), "FF1F3864");

  const previousMetrics = readSummaryMetrics(previousWorkbook);
  const currentMetrics = readSummaryMetrics(currentWorkbook);
  const metricNames = [...new Set([...previousMetrics.keys(), ...currentMetrics.keys()])];
  for (const metric of metricNames) {
    const previous = previousMetrics.get(metric) ?? "";
    const current = currentMetrics.get(metric) ?? "";
    const previousNumber = parseMetricNumber(previous);
    const currentNumber = parseMetricNumber(current);
    const delta = previousNumber !== null && currentNumber !== null
      ? currentNumber - previousNumber
      : "";
    const status = !previous && current ? "Added"
      : previous && !current ? "Removed"
      : previous === current ? "Unchanged"
      : "Updated";
    if (status === "Unchanged") continue;
    summarySheet.addRow([metric, previous, current, delta, status]);
  }

  const metricEndRow = summarySheet.rowCount;
  for (let rowNumber = 6; rowNumber <= metricEndRow; rowNumber++) {
    const row = summarySheet.getRow(rowNumber);
    styleComparisonDataRow(row, rowNumber, 5);
    if (row.getCell(5).value === "Updated") {
      for (const column of [2, 3, 4]) {
        row.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
      }
    }
  }

  summarySheet.addRow([]);
  const datasetHeaderRow = summarySheet.rowCount + 1;
  summarySheet.addRow(["Dataset", "Previous records", "Current records", "Added", "Removed", "Updated", "Unchanged", "Field changes"]);
  styleHeader(summarySheet.getRow(datasetHeaderRow), "FF375623");
  for (const dataset of datasetComparisons) {
    const row = summarySheet.addRow([
      dataset.section,
      dataset.previousCount,
      dataset.currentCount,
      dataset.Added,
      dataset.Removed,
      dataset.Updated,
      dataset.Unchanged,
      dataset.fieldsChanged,
    ]);
    styleDataRow(row, row.number);
  }

  summarySheet.addRow([]);
  const detailHeaderRow = summarySheet.rowCount + 1;
  summarySheet.addRow(["Dataset", "Change type", "URL / Item", "Field", "Previous value", "Current value"]);
  styleHeader(summarySheet.getRow(detailHeaderRow), "FF1F3864");
  for (const detail of comparisonDetails) {
    const fieldChanges = detail.values
      ? [...detail.values]
        .filter(([field]) => field && field !== "#")
        .map(([field, value]) => ({
          field,
          previous: detail.status === "Removed" ? value : "",
          current: detail.status === "Added" ? value : "",
        }))
      : [detail];

    for (const change of fieldChanges) {
      const row = summarySheet.addRow([
        detail.section,
        detail.status,
        detail.identity,
        change.field,
        change.previous,
        change.current,
      ]);
      styleComparisonDataRow(row, row.number, 2);
      for (const column of [5, 6]) {
        row.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
      }
    }
  }
  summarySheet.views = [{ state: "frozen", ySplit: 5 }];
  summarySheet.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5, column: 5 } };

  return comparisonWorkbook;
}

// ══════════════════════════════════════════════════════════════════════════════
//  MAIN — Runs all steps for each website
// ══════════════════════════════════════════════════════════════════════════════
async function runAudit() {
  const onlyStep = process.argv.find(a => a.startsWith("--only="))?.split("=")[1];
  const validSteps = new Set(["crawl", "seo", "alt", "links", "ai"]);
  if (onlyStep && !validSteps.has(onlyStep)) {
    throw new Error(`Unknown audit step "${onlyStep}". Use: crawl, seo, alt, links, or ai.`);
  }
  const runSeo   = !onlyStep || onlyStep === "seo" || onlyStep === "ai";
  const runAlt   = !onlyStep || onlyStep === "alt";
  const runLinks = !onlyStep || onlyStep === "links" || onlyStep === "ai";
  const runAi    = !onlyStep || onlyStep === "ai";

  console.log("\n╔══════════════════════════════════════════╗");
  console.log("║       SEO AUDIT TOOL — Starting          ║");
  console.log(`║  ${new Date().toLocaleString().padEnd(40)}║`);
  if (onlyStep) console.log(`║  Mode: --only=${onlyStep.padEnd(27)}║`);
  console.log("╚══════════════════════════════════════════╝\n");

  const sitesToAudit = CONFIG.websites;
  const total = sitesToAudit.length;

  // Print initial queue status
  console.log("  📋  Audit Queue:");
  sitesToAudit.forEach((s, i) => console.log(`       ${i + 1}. ${s.url}  [Waiting]`));
  console.log("");

  const browser     = await chromium.launch({ headless: true });
  const attachments = [];
  const allSiteSummaries = [];
  const auditResults = sitesToAudit.map(s => ({ url: s.url, status: "Waiting" }));

  for (let siteIdx = 0; siteIdx < sitesToAudit.length; siteIdx++) {
    const site    = sitesToAudit[siteIdx];
    const siteUrl = site.url;
    const bypass  = site.bypassHttpCheck ?? false;
    const siteName  = new URL(siteUrl).hostname;
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const outputFile = `SEO_Audit_${siteName}_${timestamp}.xlsx`;

    auditResults[siteIdx].status = "Auditing";

    console.log(`\n${"═".repeat(60)}`);
    console.log(`  🌐  Website ${siteIdx + 1}/${total}: ${siteUrl}`);
    console.log(`  🔒  Bypass HTTP Check: ${bypass ? "Yes (Cloudflare mode)" : "No (Fast mode)"}`);
    console.log(`  📊  Status: Auditing`);
    console.log(`${"═".repeat(60)}\n`);

    let crawlData;
    try {
      console.log("  📡  STEP 1 — Discovering URLs (sitemap / crawl)...");
      crawlData = await discoverUrls(siteUrl, browser);
      console.log(`  ℹ️   Discovery method: ${crawlData.discoveryMethod} | URLs found: ${crawlData.urls.length}`);
    } catch (err) {
      console.error(`  ❌  URL discovery failed for ${siteUrl}: ${err.message}`);
      auditResults[siteIdx].status = `Failed — ${err.message}`;
      continue;
    }

    let seoData = [], altData = [], linkData = [], aiData = [];

    try {
      if (runSeo) {
        console.log("\n  🔍  STEP 2 — Fetching SEO details...");
        seoData = await fetchSeoDetails(crawlData.urls, browser, bypass);
      } else { console.log("\n  ⏭️   STEP 2 — Skipped"); }

      if (runAlt) {
        console.log("\n  🖼️   STEP 3 — Auditing image alt tags...");
        altData = await fetchAltTags(crawlData.urls, crawlData.baseDomain, browser);
      } else { console.log("\n  ⏭️   STEP 3 — Skipped"); }

      if (runLinks) {
        console.log("\n  🔗  STEP 4 — Collecting internal/external links...");
        linkData = await collectLinks(crawlData.urls, crawlData.baseDomain, browser);
      } else { console.log("\n  ⏭️   STEP 4 — Skipped"); }

      if (runAi) {
        aiData = await runAiVisibilityAudit(siteUrl, crawlData, seoData, linkData, browser);
      } else { console.log("\n  ⏭️   STEP 5 — Skipped"); }

      console.log("\n  💾  Building Excel report...");
      const { wb, aiSummary } = await buildExcel(siteName, crawlData, seoData, altData, linkData, aiData);
      const previousReportPath = findPreviousReport(siteName, outputFile);
      await wb.xlsx.writeFile(outputFile);
      console.log(`  📊  Saved → ${outputFile}`);

      attachments.push({ filename: path.basename(outputFile), path: outputFile, siteName });
      const comparisonFile = path.join(path.dirname(outputFile), `SEO_Comparison_${siteName}_${timestamp}.xlsx`);
      try {
        const comparisonWorkbook = await buildComparisonReport(siteName, previousReportPath, outputFile, wb);
        await comparisonWorkbook.xlsx.writeFile(comparisonFile);
        if (previousReportPath) {
          console.log(`  📊  Comparison report saved → ${comparisonFile}`);
        } else {
          console.log(`  📊  Initial comparison report saved → ${comparisonFile}`);
        }
        attachments.push({ filename: path.basename(comparisonFile), path: comparisonFile, siteName });
      } catch (comparisonError) {
        console.warn(`  ⚠ Comparison report failed: ${comparisonError.message}`);
      }
      allSiteSummaries.push({ siteName, siteUrl, aiSummary, crawlCount: crawlData.urls.length });
      auditResults[siteIdx].status = "✓ Completed";
      console.log(`\n  ✅  ${siteName} — COMPLETE!`);
    } catch (err) {
      console.error(`  ❌  Audit failed for ${siteUrl}: ${err.message}`);
      auditResults[siteIdx].status = `✗ Failed — ${err.message}`;
    }

    // Print live queue status after each site
    console.log(`\n  📋  Progress (${siteIdx + 1}/${total} done):`);
    auditResults.forEach((r, i) => {
      const icon = r.status.startsWith("✓") ? "✅" : r.status.startsWith("✗") || r.status.startsWith("Failed") ? "❌" : r.status === "Auditing" ? "🔄" : "⏳";
      console.log(`       ${icon}  ${i + 1}. ${r.url}  [${r.status}]`);
    });
    console.log("");
  }

  await browser.close();

  if (CONFIG.email.enabled && attachments.length > 0) {
    console.log("\n  📧  Sending combined email report...");
    await sendCombinedEmail(attachments, allSiteSummaries);
  }

  // Final summary
  console.log("\n╔══════════════════════════════════════════╗");
  console.log("║        ALL AUDITS COMPLETE! 🎉           ║");
  console.log("╚══════════════════════════════════════════╝");
  console.log("\n  Final Results:");
  auditResults.forEach((r, i) => {
    const icon = r.status.startsWith("✓") ? "✅" : "❌";
    console.log(`    ${icon}  ${i + 1}. ${r.url}  →  ${r.status}`);
  });
  console.log("");

  if (!process.stdin.isTTY) process.exit(0);
}

// ══════════════════════════════════════════════════════════════════════════════
//  EMAIL — HTML with AI Visibility summary table
// ══════════════════════════════════════════════════════════════════════════════
async function sendCombinedEmail(attachments, allSiteSummaries) {
  const transporter = nodemailer.createTransport({
    host   : CONFIG.email.host,
    port   : CONFIG.email.port,
    secure : CONFIG.email.secure,
    auth   : { user: CONFIG.email.user, pass: CONFIG.email.pass },
  });

  const date = new Date().toLocaleDateString("en-IN", { year: "numeric", month: "long", day: "numeric" });

  // Build per-site summary rows for the email table
  const siteRows = allSiteSummaries.map(s => {
    const ai = s.aiSummary;
    const scoreColor = ai.avgAiScore >= 75 ? "#1E7145" : ai.avgAiScore >= 50 ? "#9C6500" : "#9C0006";
    const scoreBg    = ai.avgAiScore >= 75 ? "#E2EFDA" : ai.avgAiScore >= 50 ? "#FFF2CC" : "#FFCCCC";
    return `
      <tr>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;font-weight:600;color:#1F3864">
          <a href="${s.siteUrl}" style="color:#1155CC;text-decoration:none">${s.siteName}</a>
        </td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center">${s.crawlCount}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;font-weight:700;color:${scoreColor};background:${scoreBg}">${ai.avgAiScore}/100</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;color:#1E7145">${ai.goodPages}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;color:${ai.criticalPages > 0 ? "#9C0006" : "#1E7145"}">${ai.criticalPages}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;color:${ai.aiBlocked > 0 ? "#9C0006" : "#1E7145"}">${ai.aiBlocked}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;color:${ai.aiJsDependent > 0 ? "#9C6500" : "#1E7145"}">${ai.aiJsDependent}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #e0e0e0;text-align:center;color:${ai.aiThin > 0 ? "#9C6500" : "#1E7145"}">${ai.aiThin}</td>
      </tr>`;
  }).join("");

  const htmlBody = `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,sans-serif">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:30px 0">
    <tr><td align="center">
      <table width="680" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)">

        <!-- Header -->
        <tr>
          <td style="background:#1F3864;padding:28px 32px">
            <p style="margin:0;font-size:22px;font-weight:700;color:#ffffff">Monthly SEO Audit Report</p>
            <p style="margin:6px 0 0;font-size:14px;color:#B0C4DE">${date} &nbsp;|&nbsp; ${attachments.length} site${attachments.length > 1 ? "s" : ""} audited</p>
          </td>
        </tr>

        <!-- Intro -->
        <tr>
          <td style="padding:24px 32px 8px">
            <p style="margin:0;font-size:15px;color:#333;line-height:1.6">
              Hello,<br><br>
              Please find attached the complete SEO audit reports including the new <strong>AI Visibility analysis</strong>.
              The table below gives you the headline numbers for each site — open the Excel report for the full breakdown.
            </p>
          </td>
        </tr>

        <!-- AI Visibility summary table -->
        <tr>
          <td style="padding:16px 32px 24px">
            <p style="margin:0 0 12px;font-size:16px;font-weight:700;color:#1F3864">🤖 AI Visibility Summary</p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;font-size:13px">
              <tr style="background:#1F3864">
                <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:600">Site</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Pages</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Avg Score</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Good ✓</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Critical ✗</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Blocked</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">JS-Only</th>
                <th style="padding:10px 14px;color:#fff;text-align:center;font-weight:600">Thin</th>
              </tr>
              ${siteRows}
            </table>
            <p style="margin:10px 0 0;font-size:12px;color:#888">
              Score: &ge;75 = Good &nbsp;|&nbsp; 50–74 = Needs Work &nbsp;|&nbsp; &lt;50 = Critical &nbsp;|&nbsp; Blocked = AI bots cannot access &nbsp;|&nbsp; JS-Only = invisible to most AI crawlers
            </p>
          </td>
        </tr>

        <!-- Attachments note -->
        <tr>
          <td style="padding:0 32px 24px">
            <p style="margin:0 0 10px;font-size:15px;font-weight:700;color:#1F3864">📎 Attached Reports</p>
            <ul style="margin:0;padding-left:20px;color:#333;font-size:14px;line-height:2">
              ${attachments.map(a => `<li>${a.filename}</li>`).join("")}
            </ul>
            <p style="margin:12px 0 0;font-size:13px;color:#555">
              Each Excel file contains 5 tabs: All URLs, SEO Data, Alt Tags, Links, and the new <strong>AI Visibility</strong> sheet with per-page scores across all 11 layers.
            </p>
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="background:#f4f6fb;padding:20px 32px;border-top:1px solid #e0e0e0">
            <p style="margin:0;font-size:13px;color:#888">
              Automated report generated by SEO Audit Tool &nbsp;|&nbsp; Bombay DC
            </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;

  await transporter.sendMail({
    from       : `"SEO Audit Bot" <${CONFIG.email.user}>`,
    to         : CONFIG.email.to,
    subject    : `Monthly SEO Audit Report — ${date}`,
    text       : `Hello,\n\nPlease find attached the Monthly SEO Audit Reports.\nAudit completed on: ${new Date().toLocaleString()}\n\nRegards,\nSanket Ghadmode`,
    html       : htmlBody,
    attachments: attachments.map(a => ({ filename: a.filename, path: a.path })),
  });

  console.log(`  ✅ Email sent to ${CONFIG.email.to} with ${attachments.length} report(s) attached`);
}

// ── Schedule + controls ────────────────────────────────────────────────────────
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

runAudit().catch(err => console.error("  ❌  Audit failed:", err.message));