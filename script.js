/**************************************************
 * TPI Live Communication Board (STACKED)
 *
 * VISUAL / RULES:
 * - NEEDED BY cell only:
 *   - RED    = overdue OR due within 7 days
 *   - YELLOW = due within 10 days (but > 7)
 *   - CLEAR  = due > 10 days (no highlight)
 *
 * DISPLAY:
 * - NEEDED BY shows DATE ONLY (no time ever)
 * - Timestamp shows date/time
 *
 * UX:
 * - Maximize button toggles either panel full view
 * - Auto-scroll:
 *   - Live Notes ALWAYS auto-scrolls (unless user hovers it)
 *   - When Action Items is maximized, Action Items will auto-scroll too
 **************************************************/

// ===== Spreadsheet + GIDs =====
const sheetID = '1UFkn-d_t3DTt1RCHqp4K3HOuTMyrEVBmZnj1in1PoHc';
const GID_LIVE_NOTES = '863386477';
const GID_NEW_PARTS  = '2113651494';

// ===== Form URLs (buttons open these) =====
const FORM_URL_LIVE_NOTES =
  'https://docs.google.com/forms/d/e/1FAIpQLSeGDsKlB1DcVsFDfbqsHQPU3lxeqtk41LB5Z_OcvuzKgDTzJA/viewform';

const FORM_URL_NEW_PARTS =
  'https://docs.google.com/forms/d/e/1FAIpQLSfHxFmvRXZP4smCSIJkvG1Q83m8W-VhG7Rw7asizmBoXJLLNA/viewform';

// ===== Refresh / Clock =====
const REFRESH_MS = 15 * 1000;
const CLOCK_MS = 1000;

// ===== Auto-scroll =====
const SCROLL_SPEED = 0.35;
const SCROLL_TICK = 25;

// ===== Needed-by warning windows =====
const RED_WINDOW_DAYS = 7;
const YELLOW_WINDOW_DAYS = 10;

// ===== Column header names =====
const COL_STATUS    = 'STATUS';
const COL_SAMPLES   = 'SAMPLES';
const COL_PRIORITY  = 'PRIORITY';
const COL_NEEDED_BY = 'NEEDED BY';

// ===== Visible columns by index =====
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
const actionContainer = document.getElementById('action-container');

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

/**
 * Parse Google gviz date formats + regular strings into ms.
 * Returns Infinity when invalid/empty (so sorting pushes to bottom).
 */
function parseAnyDateMs(v) {
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

/**
 * Display ONLY date (no time) for NEEDED BY.
 * This forces date-only even if the input includes "12:00:00 AM".
 */
function formatDateOnly(v) {
  const ms = parseAnyDateMs(v);
  if (!Number.isFinite(ms) || ms === Infinity) return '';
  return new Date(ms).toLocaleDateString();
}

/**
 * Display Timestamp etc. (date/time when present)
 */
function formatTimestamp(v) {
  if (!v) return '';
  if (typeof v === 'string' && v.startsWith('Date(')) {
    const nums = v.match(/\d+/g)?.map(Number) || [];
    const [y, m, d, hh=0, mm=0, ss=0] = nums;
    return new Date(y, m, d, hh, mm, ss).toLocaleString();
  }
  const dt = new Date(String(v));
  if (Number.isFinite(dt.getTime())) return dt.toLocaleString();
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

/**
 * Build a row and apply:
 * - HOLD dimming / HOT bold / SAMPLES highlight
 * - NEEDED BY cell color rules
 * - NEEDED BY display date-only
 */
function buildRow(row, cols, visibleCols, opts = {}) {
  const tr = document.createElement('tr');

  if (opts.priorityIdx !== undefined) {
    const p = normalize(cellVal(row.c[opts.priorityIdx]));
    if (p === 'HOT') tr.classList.add('row-hot');
  }

  if (opts.samplesIdx !== undefined) {
    const s = normalize(cellVal(row.c[opts.samplesIdx]));
    if (s === 'YES') tr.classList.add('row-sample');
  }

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

    // Display rules:
    if (header === 'NEEDED BY') {
      td.textContent = formatDateOnly(v); // FORCE DATE ONLY
    } else if (header.includes('TIME') || header.includes('DATE')) {
      td.textContent = formatTimestamp(v);
    } else {
      td.textContent = v;
    }

    // Color ONLY the NEEDED BY cell
    if (opts.neededByIdx !== undefined && i === opts.neededByIdx) {
      const dueMs = parseAnyDateMs(v);
      if (Number.isFinite(dueMs) && dueMs !== Infinity) {
        const daysAway = Math.floor((dueMs - todayMs) / 86400000);

        if (daysAway <= RED_WINDOW_DAYS) {
          td.classList.add('needed-red');
        } else if (daysAway <= YELLOW_WINDOW_DAYS) {
          td.classList.add('needed-yellow');
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

/* =========================
   MAXIMIZE / RESTORE
   ========================= */
function setMaxMode(mode) {
  document.body.classList.remove('max-action', 'max-feed');
  if (mode === 'action') document.body.classList.add('max-action');
  if (mode === 'feed') document.body.classList.add('max-feed');

  // Update button text
  const btnA = document.getElementById('btn-max-action');
  const btnF = document.getElementById('btn-max-feed');
  if (btnA) btnA.textContent = document.body.classList.contains('max-action') ? '⤡ Restore' : '⤢ Maximize';
  if (btnF) btnF.textContent = document.body.classList.contains('max-feed') ? '⤡ Restore' : '⤢ Maximize';

  // Convenience: when maximizing either panel, make sure its auto-scroll is active
  if (mode === 'feed') {
    pauseScrollFeed = false;
    // optional: restart from top when maximizing
    // feedContainer.scrollTop = 0;
  }
  if (mode === 'action') {
    pauseScrollAction = false;
    // optional: restart from top when maximizing
    // actionContainer.scrollTop = 0;
  }
}

document.getElementById('btn-max-action')?.addEventListener('click', () => {
  setMaxMode(document.body.classList.contains('max-action') ? null : 'action');
});

document.getElementById('btn-max-feed')?.addEventListener('click', () => {
  setMaxMode(document.body.classList.contains('max-feed') ? null : 'feed');
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') setMaxMode(null);
});

// ===== MAIN LOAD =====
async function loadBoard() {
  try {
    const [live, parts] = await Promise.all([
      fetchGvizTable(GID_LIVE_NOTES),
      fetchGvizTable(GID_NEW_PARTS),
    ]);

    /* ===== LIVE NOTES ===== */
    const liveCols = live.cols || [];
    const liveRowsAll = (live.rows || []).slice();
    const liveMap = buildColIndexMap(liveCols);

    feedHeaders.innerHTML = buildHeader(liveCols, COLS_LIVE_NOTES);
    feedBody.innerHTML = '';

    const liveStatusIdx   = liveMap[normalize(COL_STATUS)];
    const livePriorityIdx = liveMap[normalize(COL_PRIORITY)];
    const liveNeededIdx   = liveMap[normalize(COL_NEEDED_BY)];

    // Newest first by Timestamp col 0
    liveRowsAll.sort((a, b) => {
      const ta = a.c?.[0]?.v ? parseAnyDateMs(a.c[0].v) : 0;
      const tb = b.c?.[0]?.v ? parseAnyDateMs(b.c[0].v) : 0;
      return tb - ta;
    });

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

    /* ===== NEW PARTS ===== */
    const partCols = parts.cols || [];
    const partRowsAll = (parts.rows || []).slice();
    const partMap = buildColIndexMap(partCols);

    actionHeaders.innerHTML = buildHeader(partCols, COLS_NEW_PARTS);
    actionBody.innerHTML = '';

    const partStatusIdx  = partMap[normalize(COL_STATUS)];
    const partSamplesIdx = partMap[normalize(COL_SAMPLES)];
    const partNeededIdx  = partMap[normalize(COL_NEEDED_BY)];

    const partRows = partRowsAll.filter(r => {
      if (partStatusIdx === undefined) return true;
      return normalize(cellVal(r.c[partStatusIdx])) !== 'DONE';
    });

    // Sort: OPEN before HOLD → earliest NEEDED BY → oldest Timestamp
    partRows.sort((a, b) => {
      const sa = partStatusIdx !== undefined ? statusRank(cellVal(a.c[partStatusIdx])) : 5;
      const sb = partStatusIdx !== undefined ? statusRank(cellVal(b.c[partStatusIdx])) : 5;
      if (sa !== sb) return sa - sb;

      const na = partNeededIdx !== undefined ? parseAnyDateMs(cellVal(a.c[partNeededIdx])) : Infinity;
      const nb = partNeededIdx !== undefined ? parseAnyDateMs(cellVal(b.c[partNeededIdx])) : Infinity;
      if (na !== nb) return na - nb;

      const ta = a.c?.[0]?.v ? parseAnyDateMs(a.c[0].v) : 0;
      const tb = b.c?.[0]?.v ? parseAnyDateMs(b.c[0].v) : 0;
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

/* =========================
   AUTO-SCROLL
   =========================
   - Feed: always auto-scrolls unless user hovers it
   - Action: auto-scrolls ONLY when maximized (unless user hovers it)
*/
let pauseScrollFeed = false;
let pauseScrollAction = false;

feedContainer.addEventListener('mouseenter', () => pauseScrollFeed = true);
feedContainer.addEventListener('mouseleave', () => pauseScrollFeed = false);

actionContainer.addEventListener('mouseenter', () => pauseScrollAction = true);
actionContainer.addEventListener('mouseleave', () => pauseScrollAction = false);

// Live Notes auto-scroll (always)
setInterval(() => {
  if (pauseScrollFeed) return;

  const max = feedContainer.scrollHeight - feedContainer.clientHeight;
  if (max <= 0) return;

  feedContainer.scrollTop += SCROLL_SPEED;
  if (feedContainer.scrollTop >= max) feedContainer.scrollTop = 0;
}, SCROLL_TICK);

// Action Items auto-scroll (only when maximized)
setInterval(() => {
  if (!document.body.classList.contains('max-action')) return;
  if (pauseScrollAction) return;

  const max = actionContainer.scrollHeight - actionContainer.clientHeight;
  if (max <= 0) return;

  actionContainer.scrollTop += SCROLL_SPEED;
  if (actionContainer.scrollTop >= max) actionContainer.scrollTop = 0;
}, SCROLL_TICK);

// ===== BOOT =====
loadBoard();
updateClock();
setInterval(loadBoard, REFRESH_MS);
setInterval(updateClock, CLOCK_MS);
setMaxMode(null);
