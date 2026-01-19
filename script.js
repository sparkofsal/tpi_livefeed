/**************************************************
 * TPI Live Communication Board (No Backend)
 *
 * RIGHT panel: Live Notes tab (scrolling feed)
 * LEFT panel: New Parts & Projects tab (action items)
 *
 * NEW STATUS VALUES: OPEN / HOLD / DONE
 * Action sorting: HOT first -> OPEN before HOLD -> OLDEST first
 * Highlight: SAMPLES = YES
 **************************************************/

// ===== Spreadsheet + GIDs =====
const sheetID = '1W2Fc0P8Ye9ICLPhbtEIM35n3g_3swuE9HZZ1YU_NqPA';
const GID_LIVE_NOTES = '1591460905';
const GID_NEW_PARTS  = '1564470138';

// ===== Form URLs (buttons open these) =====
const FORM_URL_LIVE_NOTES =
  'https://docs.google.com/forms/d/e/1FAIpQLSeGDsKlB1DcVsFDfbqsHQPU3lxeqtk41LB5Z_OcvuzKgDTzJA/viewform';

const FORM_URL_NEW_PARTS =
  'https://docs.google.com/forms/d/e/1FAIpQLSfHxFmvRXZP4smCSIJkvG1Q83m8W-VhG7Rw7asizmBoXJLLNA/viewform';

// ===== Refresh / Clock =====
const REFRESH_MS = 15 * 1000; // 15 seconds
const CLOCK_MS = 1000;

// ===== Auto-scroll (right panel) =====
const SCROLL_SPEED = 0.35;
const SCROLL_TICK = 25;

// ===== Column header names (match your sheet headers) =====
// Live Notes (A-H): Timestamp, Requested By, Job/Group, Customer, Priority, Notes, Owner, Needed By
const LN_COL_PRIORITY = 'PRIORITY';

// New Parts (A-I): Timestamp, Job/Group, Customer, Requested By, Samples, Priority, Details, Owner, Status
const NP_COL_STATUS   = 'STATUS';
const NP_COL_PRIORITY = 'PRIORITY';
const NP_COL_SAMPLES  = 'SAMPLES';

// ===== Visible columns by index (based on your new layouts) =====
const COLS_LIVE_NOTES = [0,1,2,3,4,5,6,7];     // 8 columns (A-H)
const COLS_NEW_PARTS  = [0,1,2,3,4,5,6,7,8];   // 9 columns (A-I)

// ===== DOM =====
const actionHeaders = document.getElementById('action-headers');
const actionBody = document.getElementById('action-body');
const feedHeaders = document.getElementById('feed-headers');
const feedBody = document.getElementById('feed-body');
const actionCount = document.getElementById('action-count');
const feedCount = document.getElementById('feed-count');
const feedContainer = document.getElementById('feed-container');

// Buttons open forms
document.getElementById('btn-live-notes').href = FORM_URL_LIVE_NOTES;
document.getElementById('btn-new-parts').href = FORM_URL_NEW_PARTS;

// ===== Helpers =====
const normalize = v => String(v ?? '').trim().toUpperCase();
const cellVal = c => c?.v ?? '';

function priorityRank(v) {
  const p = normalize(v);
  if (p === 'HOT') return 0;
  if (p === 'NORMAL') return 1;
  if (p === 'LOW') return 2;
  return 9;
}

// OPEN first, then HOLD, DONE last (DONE is filtered out anyway)
function statusRank(v) {
  const s = normalize(v);
  if (s === 'OPEN') return 0;
  if (s === 'HOLD') return 1;
  if (s === 'DONE') return 9;
  return 5;
}

// gviz date parsing (Date(YYYY,MM,DD,hh,mm,ss)) + safe fallbacks
function formatAnyDate(v) {
  if (!v) return '';

  if (typeof v === 'number') {
    const base = new Date(1899, 11, 30);
    return new Date(base.getTime() + v * 86400000).toLocaleDateString();
  }

  if (typeof v === 'string' && v.startsWith('Date(')) {
    const nums = v.match(/\d+/g)?.map(Number) || [];
    const [y, m, d, hh=0, mm=0, ss=0] = nums;
    const dt = new Date(y, m, d, hh, mm, ss);
    return dt.toLocaleString();
  }

  return String(v);
}

async function fetchGvizTable(gid) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetID}/gviz/tq?gid=${gid}&tqx=out:json&cb=${Date.now()}`;
  const res = await fetch(url);
  const txt = await res.text();

  const start = txt.indexOf('{');
  const end = txt.lastIndexOf('}');
  if (start === -1 || end === -1) {
    throw new Error('No gviz JSON received. Check Google Sheet sharing: Anyone with link = Viewer.');
  }

  const json = JSON.parse(txt.slice(start, end + 1));
  if (!json?.table) throw new Error('gviz JSON parsed but table data missing.');

  return json.table;
}

function buildColIndexMap(cols) {
  const map = {};
  cols.forEach((c, i) => {
    const key = normalize(c.label);
    if (key) map[key] = i;
  });
  return map;
}

function buildHeader(cols, visibleCols) {
  return visibleCols.map(i => `<th>${cols[i]?.label ?? ''}</th>`).join('');
}

function buildRow(row, cols, visibleCols, opts = {}) {
  const tr = document.createElement('tr');

  // Bold HOT
  if (opts.priorityIdx !== undefined) {
    const p = normalize(cellVal(row.c[opts.priorityIdx]));
    if (p === 'HOT') tr.classList.add('row-hot');
  }

  // Highlight SAMPLES = YES
  if (opts.samplesIdx !== undefined) {
    const s = normalize(cellVal(row.c[opts.samplesIdx]));
    if (s === 'YES') tr.classList.add('row-sample');
  }

  // Slight dim for HOLD rows (optional)
  if (opts.statusIdx !== undefined) {
    const st = normalize(cellVal(row.c[opts.statusIdx]));
    if (st === 'HOLD') tr.classList.add('row-hold');
  }

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const headerName = normalize(cols[i]?.label);
    const v = cellVal(row.c[i]);

    // Format Timestamp + Needed By fields
    const isDateField =
      headerName.includes('TIME') ||
      headerName.includes('DATE') ||
      headerName.includes('NEEDED BY');

    td.textContent = isDateField ? formatAnyDate(v) : v;
    tr.appendChild(td);
  });

  return tr;
}

function updateClock() {
  document.getElementById('datetime').textContent = new Date().toLocaleString();
}

// ===== Main load =====
async function loadBoard() {
  try {
    const [live, parts] = await Promise.all([
      fetchGvizTable(GID_LIVE_NOTES),
      fetchGvizTable(GID_NEW_PARTS),
    ]);

    // --------------------------
    // RIGHT: LIVE NOTES
    // --------------------------
    const liveCols = live.cols || [];
    const liveRows = (live.rows || []).slice();
    const liveMap = buildColIndexMap(liveCols);

    feedHeaders.innerHTML = buildHeader(liveCols, COLS_LIVE_NOTES);
    feedBody.innerHTML = '';

    // Newest first (timestamp col 0)
    liveRows.sort((a, b) => {
      const ta = a.c?.[0]?.v ? new Date(a.c[0].v).getTime() : 0;
      const tb = b.c?.[0]?.v ? new Date(b.c[0].v).getTime() : 0;
      return tb - ta;
    });

    const livePriorityIdx = liveMap[normalize(LN_COL_PRIORITY)];
    liveRows.forEach(r => {
      feedBody.appendChild(buildRow(r, liveCols, COLS_LIVE_NOTES, {
        priorityIdx: livePriorityIdx
      }));
    });

    feedCount.textContent = `${liveRows.length} notes`;

    // --------------------------
    // LEFT: NEW PARTS & PROJECTS (Action Items)
    // --------------------------
    const partCols = parts.cols || [];
    const partRowsAll = (parts.rows || []).slice();
    const partMap = buildColIndexMap(partCols);

    actionHeaders.innerHTML = buildHeader(partCols, COLS_NEW_PARTS);
    actionBody.innerHTML = '';

    const statusIdx = partMap[normalize(NP_COL_STATUS)];
    const priorityIdx = partMap[normalize(NP_COL_PRIORITY)];
    const samplesIdx = partMap[normalize(NP_COL_SAMPLES)];

    // Keep items that are NOT DONE
    const partRows = partRowsAll.filter(r => {
      if (statusIdx === undefined) return true;
      const st = normalize(cellVal(r.c[statusIdx]));
      return st !== 'DONE';
    });

    // Sort: HOT first -> OPEN before HOLD -> OLDEST first (accountability)
    partRows.sort((a, b) => {
      const pa = priorityIdx !== undefined ? priorityRank(cellVal(a.c[priorityIdx])) : 9;
      const pb = priorityIdx !== undefined ? priorityRank(cellVal(b.c[priorityIdx])) : 9;
      if (pa !== pb) return pa - pb;

      const sa = statusIdx !== undefined ? statusRank(cellVal(a.c[statusIdx])) : 5;
      const sb = statusIdx !== undefined ? statusRank(cellVal(b.c[statusIdx])) : 5;
      if (sa !== sb) return sa - sb;

      // OLDEST first (prevents cherry-picking)
      const ta = a.c?.[0]?.v ? new Date(a.c[0].v).getTime() : 0;
      const tb = b.c?.[0]?.v ? new Date(b.c[0].v).getTime() : 0;
      return ta - tb;
    });

    partRows.forEach(r => {
      actionBody.appendChild(buildRow(r, partCols, COLS_NEW_PARTS, {
        priorityIdx,
        samplesIdx,
        statusIdx
      }));
    });

    actionCount.textContent = `${partRows.length} items`;

  } catch (err) {
    console.error(err);

    const msg = `⚠️ ${err.message}`;
    actionBody.innerHTML = `<tr><td colspan="100%">${msg}</td></tr>`;
    feedBody.innerHTML = `<tr><td colspan="100%">${msg}</td></tr>`;
    actionCount.textContent = '—';
    feedCount.textContent = '—';
  }
}

// ===== Auto-scroll (right feed only) =====
let pauseScroll = false;
feedContainer.addEventListener('mouseenter', () => pauseScroll = true);
feedContainer.addEventListener('mouseleave', () => pauseScroll = false);

setInterval(() => {
  if (pauseScroll) return;

  const max = feedContainer.scrollHeight - feedContainer.clientHeight;
  if (max <= 0) return;

  feedContainer.scrollTop += SCROLL_SPEED;
  if (feedContainer.scrollTop >= max) feedContainer.scrollTop = 0;
}, SCROLL_TICK);

// ===== Boot =====
loadBoard();
updateClock();
setInterval(loadBoard, REFRESH_MS);
setInterval(updateClock, CLOCK_MS);
