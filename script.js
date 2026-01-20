/**************************************************
 * TPI Live Communication Board (STACKED)
 * Top: New Parts & Projects (Action Items)
 * Bottom: Live Notes (Auto-scroll)
 *
 * Columns (both):
 * A Timestamp (auto)
 * ...
 * NEEDED BY + STATUS exist on both
 *
 * STATUS values: OPEN / HOLD / DONE
 * - DONE is hidden on both
 *
 * SORTING:
 * - Live Notes: newest Timestamp first (true feed)
 * - New Parts: OPEN before HOLD → earliest NEEDED BY → oldest Timestamp
 *
 * VISUAL:
 * - NEEDED BY cell:
 *     RED if due today or overdue
 *     YELLOW if due within 11 days
 * - SAMPLES = YES highlighted (New Parts only)
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
const REFRESH_MS = 15 * 1000;
const CLOCK_MS = 1000;

// ===== Auto-scroll (Live Notes only) =====
const SCROLL_SPEED = 0.35;
const SCROLL_TICK = 25;

// ===== Needed-by warning window =====
const YELLOW_WINDOW_DAYS = 11; // ~week and a half

// ===== Column header names =====
const COL_STATUS    = 'STATUS';
const COL_SAMPLES   = 'SAMPLES';
const COL_PRIORITY  = 'PRIORITY';   // Live Notes only
const COL_NEEDED_BY = 'NEEDED BY';

// ===== Visible columns by index =====
// Both tabs are 9 columns: Timestamp + 8 questions
const COLS_LIVE_NOTES = [0,1,2,3,4,5,6,7,8];
const COLS_NEW_PARTS  = [0,1,2,3,4,5,6,7,8];

// ===== DOM =====
const actionHeaders = document.getElementById('action-headers');
const actionBody = document.getElementById('action-body');
const feedHeaders = document.getElementById('feed-headers');
const feedBody = document.getElementById('feed-body');
const actionCount = document.getElementById('action-count');
const feedCount = document.getElementById('feed-count');
const feedContainer = document.getElementById('feed-container');

// Wire form buttons
document.getElementById('btn-live-notes').href = FORM_URL_LIVE_NOTES;
document.getElementById('btn-new-parts').href = FORM_URL_NEW_PARTS;

// ===== Helpers =====
const normalize = v => String(v ?? '').trim().toUpperCase();
const cellVal = c => c?.v ?? '';

function statusRank(v) {
  const s = normalize(v);
  if (s === 'OPEN') return 0;
  if (s === 'HOLD') return 1;
  if (s === 'DONE') return 9;
  return 5;
}

// Parses "NEEDED BY" to a comparable timestamp (ms)
// - gviz date format: Date(YYYY,MM,DD,hh,mm,ss)
// - other parseable strings like 1/20/2026
// - if empty/invalid: Infinity (sorts last)
function parseNeededByMs(v) {
  if (!v) return Infinity;

  if (typeof v === 'string' && v.startsWith('Date(')) {
    const nums = v.match(/\d+/g)?.map(Number) || [];
    const [y, m, d, hh=0, mm=0, ss=0] = nums;
    const dt = new Date(y, m, d, hh, mm, ss);
    const ms = dt.getTime();
    return Number.isFinite(ms) ? ms : Infinity;
  }

  if (v instanceof Date) {
    const ms = v.getTime();
    return Number.isFinite(ms) ? ms : Infinity;
  }

  const dt = new Date(String(v));
  const ms = dt.getTime();
  return Number.isFinite(ms) ? ms : Infinity;
}

// Formats Timestamp / Needed By for display
function formatAnyDate(v) {
  if (!v) return '';

  if (typeof v === 'number') {
    const base = new Date(1899, 11, 30);
    return new Date(base.getTime() + v * 86400000).toLocaleDateString();
  }

  if (typeof v === 'string' && v.startsWith('Date(')) {
    const nums = v.match(/\d+/g)?.map(Number) || [];
    const [y, m, d, hh=0, mm=0, ss=0] = nums;
    return new Date(y, m, d, hh, mm, ss).toLocaleString();
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
    throw new Error('No gviz JSON returned. Check Sheet sharing (Anyone with link = Viewer).');
  }

  const json = JSON.parse(txt.slice(start, end + 1));
  if (!json?.table) throw new Error('Parsed JSON but missing table.');

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

// Builds a row and applies the "NEEDED BY" cell colors
function buildRow(row, cols, visibleCols, opts = {}) {
  const tr = document.createElement('tr');

  // HOT emphasis (Live Notes only)
  if (opts.priorityIdx !== undefined) {
    const p = normalize(cellVal(row.c[opts.priorityIdx]));
    if (p === 'HOT') tr.classList.add('row-hot');
  }

  // Highlight samples needed (New Parts only)
  if (opts.samplesIdx !== undefined) {
    const s = normalize(cellVal(row.c[opts.samplesIdx]));
    if (s === 'YES') tr.classList.add('row-sample');
  }

  // HOLD dimming
  if (opts.statusIdx !== undefined) {
    const st = normalize(cellVal(row.c[opts.statusIdx]));
    if (st === 'HOLD') tr.classList.add('row-hold');
  }

  // Compare against today at midnight
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayMs = today.getTime();

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const header = normalize(cols[i]?.label);
    const v = cellVal(row.c[i]);

    const isDateField =
      header.includes('TIME') ||
      header.includes('DATE') ||
      header.includes('NEEDED BY');

    td.textContent = isDateField ? formatAnyDate(v) : v;

    // Color ONLY the NEEDED BY cell
    if (opts.neededByIdx !== undefined && i === opts.neededByIdx) {
      const dueMs = parseNeededByMs(v);
      if (Number.isFinite(dueMs) && dueMs !== Infinity) {
        const daysAway = Math.floor((dueMs - todayMs) / 86400000);

        if (daysAway <= 0) {
          td.classList.add('needed-red'); // today or overdue
        } else if (daysAway <= YELLOW_WINDOW_DAYS) {
          td.classList.add('needed-yellow'); // within 11 days
        }
      }
    }

    tr.appendChild(td);
  });

  return tr;
}

function updateClock() {
  document.getElementById('datetime').textContent = new Date().toLocaleString();
}

// ===== MAIN LOAD =====
async function loadBoard() {
  try {
    const [live, parts] = await Promise.all([
      fetchGvizTable(GID_LIVE_NOTES),
      fetchGvizTable(GID_NEW_PARTS),
    ]);

    /* =========================
       LIVE NOTES (BOTTOM FEED)
       ========================= */
    const liveCols = live.cols || [];
    const liveRowsAll = (live.rows || []).slice();
    const liveMap = buildColIndexMap(liveCols);

    feedHeaders.innerHTML = buildHeader(liveCols, COLS_LIVE_NOTES);
    feedBody.innerHTML = '';

    const liveStatusIdx  = liveMap[normalize(COL_STATUS)];
    const livePriorityIdx = liveMap[normalize(COL_PRIORITY)];
    const liveNeededIdx  = liveMap[normalize(COL_NEEDED_BY)];

    // Newest first by Timestamp (col 0)
    liveRowsAll.sort((a, b) => {
      const ta = a.c?.[0]?.v ? new Date(a.c[0].v).getTime() : 0;
      const tb = b.c?.[0]?.v ? new Date(b.c[0].v).getTime() : 0;
      return tb - ta;
    });

    // Hide DONE
    const liveRows = liveRowsAll.filter(r => {
      if (liveStatusIdx === undefined) return true;
      return normalize(cellVal(r.c[liveStatusIdx])) !== 'DONE';
    });

    liveRows.forEach(r => {
      feedBody.appendChild(buildRow(r, liveCols, COLS_LIVE_NOTES, {
        statusIdx: liveStatusIdx,
        priorityIdx: livePriorityIdx,
        neededByIdx: liveNeededIdx
      }));
    });

    feedCount.textContent = `${liveRows.length} notes`;

    /* =========================
       NEW PARTS & PROJECTS (TOP)
       ========================= */
    const partCols = parts.cols || [];
    const partRowsAll = (parts.rows || []).slice();
    const partMap = buildColIndexMap(partCols);

    actionHeaders.innerHTML = buildHeader(partCols, COLS_NEW_PARTS);
    actionBody.innerHTML = '';

    const partStatusIdx  = partMap[normalize(COL_STATUS)];
    const partSamplesIdx = partMap[normalize(COL_SAMPLES)];
    const partNeededIdx  = partMap[normalize(COL_NEEDED_BY)];

    // Hide DONE
    const partRows = partRowsAll.filter(r => {
      if (partStatusIdx === undefined) return true;
      return normalize(cellVal(r.c[partStatusIdx])) !== 'DONE';
    });

    // Sort: OPEN before HOLD → earliest NEEDED BY → oldest Timestamp
    partRows.sort((a, b) => {
      const sa = partStatusIdx !== undefined ? statusRank(cellVal(a.c[partStatusIdx])) : 5;
      const sb = partStatusIdx !== undefined ? statusRank(cellVal(b.c[partStatusIdx])) : 5;
      if (sa !== sb) return sa - sb;

      const na = partNeededIdx !== undefined ? parseNeededByMs(cellVal(a.c[partNeededIdx])) : Infinity;
      const nb = partNeededIdx !== undefined ? parseNeededByMs(cellVal(b.c[partNeededIdx])) : Infinity;
      if (na !== nb) return na - nb;

      const ta = a.c?.[0]?.v ? new Date(a.c[0].v).getTime() : 0;
      const tb = b.c?.[0]?.v ? new Date(b.c[0].v).getTime() : 0;
      return ta - tb;
    });

    partRows.forEach(r => {
      actionBody.appendChild(buildRow(r, partCols, COLS_NEW_PARTS, {
        statusIdx: partStatusIdx,
        samplesIdx: partSamplesIdx,
        neededByIdx: partNeededIdx
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

// ===== Auto-scroll (Live Notes only) =====
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

// ===== BOOT =====
loadBoard();
updateClock();
setInterval(loadBoard, REFRESH_MS);
setInterval(updateClock, CLOCK_MS);
