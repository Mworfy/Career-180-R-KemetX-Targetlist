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

function encodeSheetsRange(range) {
  if (!range.includes('!')) return encodeURIComponent(range);
  const bang = range.indexOf('!');
  const sheet = range.slice(0, bang);
  const cells = range.slice(bang + 1);
  return encodeURIComponent(`${quoteSheetTitle(sheet)}!${cells}`);
}

function normalizeSheetRows(values, columnCount) {
  const width = Math.max(columnCount || 0, ...values.map((r) => (r && r.length) || 0));
  return values.map((r) => {
    const row = Array.isArray(r) ? [...r] : [];
    while (row.length < width) row.push('');
    return row;
  });
}

async function sheetsApiGet(path, apiKey) {
  const sep = path.includes('?') ? '&' : '?';
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${path}${sep}key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(url);
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = body.error?.message || res.statusText || 'Request failed';
    throw new Error(msg);
  }
  return body;
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

async function fetchTabGrid(spreadsheetId, apiKey, tab) {
  const rows = Math.max(1, tab.gridProperties?.rowCount || 1);
  const cols = Math.max(1, tab.gridProperties?.columnCount || 1);
  const lastCol = columnIndexToA1(cols);
  const range = `${quoteSheetTitle(tab.title)}!A1:${lastCol}${rows}`;
  const rangeParam = encodeSheetsRange(range);
  const body = await sheetsApiGet(
    `${encodeURIComponent(spreadsheetId)}/values/${rangeParam}?majorDimension=ROWS`,
    apiKey,
  );
  const values = body.values || [];
  if (!values.length) return null;
  return { title: tab.title, values: normalizeSheetRows(values, cols), columnCount: cols };
}

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
    const payloads = [];
    const skipped = [];

    for (const tab of tabs) {
      try {
        const grid = await fetchTabGrid(spreadsheetId, apiKey, tab);
        if (grid) payloads.push(grid);
      } catch (err) {
        skipped.push({ title: tab.title, message: err.message });
      }
    }

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({
      configured: true,
      tabs: payloads,
      tabsTotal: tabs.length,
      skipped,
    });
  } catch (err) {
    return res.status(502).json({ error: err.message || 'Failed to load Google Sheet', configured: true });
  }
}
