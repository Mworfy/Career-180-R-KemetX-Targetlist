function quoteSheetTitle(title) {
  if (/^[\w]+$/.test(title)) return title;
  return `'${String(title).replace(/'/g, "''")}'`;
}

function columnIndexToA1(colOneBased) {
  let n = colOneBased;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function normalizeSheetRows(values, columnCount) {
  const width = Math.max(columnCount || 0, ...values.map((r) => (r && r.length) || 0));
  return values.map((r) => {
    const row = Array.isArray(r) ? [...r] : [];
    while (row.length < width) row.push('');
    return row;
  });
}

// Statuses worth retrying: rate limits and transient Google/server errors.
const RETRYABLE_STATUS = new Set([408, 429, 500, 502, 503, 504]);
const REQUEST_TIMEOUT_MS = 12000;
const MAX_ATTEMPTS = 4;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function backoffDelay(attempt, retryAfterHeader) {
  const retryAfter = Number(retryAfterHeader);
  if (Number.isFinite(retryAfter) && retryAfter > 0) return Math.min(retryAfter * 1000, 8000);
  const base = 500 * 2 ** (attempt - 1); // 0.5s, 1s, 2s…
  return base + Math.floor(Math.random() * 300);
}

async function sheetsApiGet(path, apiKey) {
  const sep = path.includes('?') ? '&' : '?';
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${path}${sep}key=${encodeURIComponent(apiKey)}`;
  let lastErr;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    let res;
    try {
      res = await fetch(url, { signal: controller.signal });
    } catch (err) {
      // Network failure or timeout: always retryable.
      lastErr = new Error(err.name === 'AbortError' ? 'Google Sheets request timed out' : `Network error: ${err.message}`);
      clearTimeout(timer);
      if (attempt < MAX_ATTEMPTS) await sleep(backoffDelay(attempt));
      continue;
    }
    clearTimeout(timer);

    const text = await res.text().catch(() => '');
    let body = {};
    try { body = text ? JSON.parse(text) : {}; } catch { body = {}; }

    if (res.ok) return body;

    const msg = body.error?.message || res.statusText || `Request failed (${res.status})`;
    lastErr = new Error(msg);
    lastErr.status = res.status;
    if (!RETRYABLE_STATUS.has(res.status)) throw lastErr; // 400/403/404: retrying won't help
    if (attempt < MAX_ATTEMPTS) await sleep(backoffDelay(attempt, res.headers.get('retry-after')));
  }
  throw lastErr;
}

async function fetchSpreadsheetTabs(spreadsheetId, apiKey, tabFilter) {
  const body = await sheetsApiGet(
    `${encodeURIComponent(spreadsheetId)}?fields=sheets(properties(title,gridProperties(rowCount,columnCount)))`,
    apiKey,
  );
  let tabs = (body.sheets || []).map((s) => s.properties).filter(Boolean);
  const filter = (tabFilter || '').trim();
  if (filter) {
    const want = filter.toLowerCase();
    tabs = tabs.filter((t) => t.title.toLowerCase() === want);
    if (!tabs.length) throw new Error(`No tab named "${filter}" in this spreadsheet.`);
  }
  return tabs;
}

function tabRange(tab) {
  const rows = Math.max(1, tab.gridProperties?.rowCount || 1);
  const cols = Math.max(1, tab.gridProperties?.columnCount || 1);
  return `${quoteSheetTitle(tab.title)}!A1:${columnIndexToA1(cols)}${rows}`;
}

function toGrid(tab, values) {
  if (!values || !values.length) return null;
  const cols = Math.max(1, tab.gridProperties?.columnCount || 1);
  return { title: tab.title, values: normalizeSheetRows(values, cols), columnCount: cols };
}

async function fetchTabGrid(spreadsheetId, apiKey, tab) {
  const body = await sheetsApiGet(
    `${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(tabRange(tab))}?majorDimension=ROWS`,
    apiKey,
  );
  return toGrid(tab, body.values);
}

// One request for every tab (cheaper on quota). Fails as a whole if any range is bad.
async function fetchAllTabsBatch(spreadsheetId, apiKey, tabs) {
  const params = tabs.map((t) => `ranges=${encodeURIComponent(tabRange(t))}`).join('&');
  const body = await sheetsApiGet(
    `${encodeURIComponent(spreadsheetId)}/values:batchGet?majorDimension=ROWS&${params}`,
    apiKey,
  );
  const ranges = body.valueRanges || [];
  return tabs.map((tab, i) => toGrid(tab, ranges[i]?.values)).filter(Boolean);
}

// Per-tab fallback: one broken tab doesn't take down the others.
async function fetchTabsIndividually(spreadsheetId, apiKey, tabs) {
  const payloads = [];
  const skipped = [];
  const results = await Promise.allSettled(tabs.map((tab) => fetchTabGrid(spreadsheetId, apiKey, tab)));
  results.forEach((r, i) => {
    if (r.status === 'fulfilled') {
      if (r.value) payloads.push(r.value);
    } else {
      skipped.push({ title: tabs[i].title, message: r.reason?.message || 'Unknown error' });
    }
  });
  return { payloads, skipped };
}

// Last good response, kept while this serverless instance stays warm.
let lastGood = null;

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.GOOGLE_SHEETS_API_KEY;
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  const envRange = process.env.GOOGLE_SHEETS_TAB_FILTER || '';
  const queryRange = typeof req.query.range === 'string' ? req.query.range : '';
  const tabFilter = queryRange || envRange;

  if (!apiKey || !spreadsheetId) {
    return res.status(500).json({
      error: 'Server missing GOOGLE_SHEETS_API_KEY or GOOGLE_SPREADSHEET_ID. Add them in Vercel → Project → Settings → Environment Variables.',
      configured: false,
    });
  }

  try {
    const tabs = await fetchSpreadsheetTabs(spreadsheetId, apiKey, tabFilter);
    let payloads;
    let skipped = [];

    try {
      payloads = await fetchAllTabsBatch(spreadsheetId, apiKey, tabs);
    } catch (batchErr) {
      ({ payloads, skipped } = await fetchTabsIndividually(spreadsheetId, apiKey, tabs));
      if (!payloads.length) throw batchErr;
    }

    const result = {
      configured: true,
      tabs: payloads,
      tabsTotal: tabs.length,
      skipped,
      fetchedAt: Date.now(),
    };
    if (!tabFilter) lastGood = result;

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json(result);
  } catch (err) {
    if (lastGood && !tabFilter) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ ...lastGood, stale: true, staleReason: err.message });
    }
    const status = err.status && !RETRYABLE_STATUS.has(err.status) ? 502 : 503;
    return res.status(status).json({ error: err.message || 'Failed to load Google Sheet', configured: true, retryable: status === 503 });
  }
}
