/**************************************************
 * TPI Live Communication Board — Modular UI
 *
 * Home screen shows only buttons.
 * Clicking a module shows only that module.
 *
 * Needed By:
 * - RED <= 7 days / overdue
 * - YELLOW <= 10 days (but > 7)
 * - CLEAR > 10 days
 *
 * NEEDED BY display = date only (no time)
 *
 * Auto-scroll:
 * - Runs ONLY for the module currently visible
 * - requestAnimationFrame + whole pixels for TV smoothness
 **************************************************/

// ===== Spreadsheet + GIDs =====
const sheetID = '1UFkn-d_t3DTt1RCHqp4K3HOuTMyrEVBmZnj1in1PoHc';
const GID_LIVE_NOTES = '863386477';
const GID_NEW_PARTS  = '2113651494';

// Shipping not ready yet: leave blank until you create a tab
const GID_SHIPPING = ''; // e.g. '123456789'

// ===== Form URLs =====
const FORM_URL_LIVE_NOTES =
  'https://docs.google.com/forms/d/e/1FAIpQLSeGDsKlB1DcVsFDfbqsHQPU3lxeqtk41LB5Z_OcvuzKgDTzJA/viewform';

const FORM_URL_NEW_PARTS =
  'https://docs.google.com/forms/d/e/1FAIpQLSfHxFmvRXZP4smCSIJkvG1Q83m8W-VhG7Rw7asizmBoXJLLNA/viewform';

// ===== Refresh / Clock =====
const REFRESH_MS = 15 * 1000;
const CLOCK_MS = 1000;

// ===== Needed-by windows =====
const RED_WINDOW_DAYS = 7;
const YELLOW_WINDOW_DAYS = 10;

// ===== Column header names =====
const COL_STATUS    = 'STATUS';
const COL_SAMPLES   = 'SAMPLES';
const COL_PRIORITY  = 'PRIORITY';
const COL_NEEDED_BY = 'NEEDED BY';

// Shipping column name (later)
const COL_SHIP_DATE = 'SHIP DATE';

// ===== Visible columns by index =====
const COLS_LIVE_NOTES = [0,1,2,3,4,5,6,7,8];
const COLS_NEW_PARTS  = [0,1,2,3,4,5,6,7,8];

// ===== DOM: views =====
const viewHome = document.getElementById('view-home');
const viewAction = document.getElementById('view-action');
const viewFeed = document.getElementById('view-feed');
const viewShipping = document.getElementById('view-shipping');

// ===== DOM: home buttons =====
const btnOpenAction = document.getElementById('btn-open-action');
const btnOpenFeed = document.getElementById('btn-open-feed');
const btnOpenShipping = document.getElementById('btn-open-shipping');

// ===== DOM: back buttons =====
const btnBackAction = document.getElementById('btn-back-action');
const btnBackFeed = document.getElementById('btn-back-feed');
const btnBackShipping = document.getElementById('btn-back-shipping');

// ===== DOM: action =====
const actionHeaders = document.getElementById('action-headers');
const actionBody = document.getElementById('action-body');
const actionCount = document.getElementById('action-count');
const actionContainer = document.getElementById('action-container');

// ===== DOM: feed =====
const feedHeaders = document.getElementById('feed-headers');
const feedBody = document.getElementById('feed-body');
const feedCount = document.getElementById('feed-count');
const feedContainer = document.getElementById('feed-container');

// ===== DOM: shipping =====
const shippingHeaders = document.getElementById('shipping-headers');
const shippingBody = document.getElementById('shipping-body');
const shippingCount = document.getElementById('shipping-count');
const shippingContainer = document.getElementById('shipping-container');

// ===== Wire form buttons =====
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

function parseAnyDateMs(v) {
  if (!v) return Infinity;

  if (typeof v === 'string' && v.startsWith('Date(')) {
    const nums = v.match(/\d+/g)?.map(Number) || [];
    const [y, m, d, hh=0, mm=0, ss=0] = nums;
    const dt = new Date(y, m, d, hh, mm, ss);
    const ms = dt.getTime();
    return Number.isFinite(ms) ? ms : Infinity;
  }

  const dt = new Date(String(v));
  const ms = dt.getTime();
  return Number.isFinite(ms) ? ms : Infinity;
}

function formatDateOnly(v) {
  const ms = parseAnyDateMs(v);
  if (!Number.isFinite(ms) || ms === Infinity) return '';
  return new Date(ms).toLocaleDateString();
}

function formatTimestamp(v) {
  const ms = parseAnyDateMs(v);
  if (!Number.isFinite(ms) || ms === Infinity) return String(v ?? '');
  return new Date(ms).toLocaleString();
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

// Row builder for Action/Feed
function buildRow(row, cols, visibleCols, opts = {}) {
  const tr = document.createElement('tr');

  // HOT emphasis (feed)
  if (opts.priorityIdx !== undefined) {
    const p = normalize(cellVal(row.c[opts.priorityIdx]));
    if (p === 'HOT') tr.classList.add('row-hot');
  }

  // Samples highlight (action)
  if (opts.samplesIdx !== undefined) {
    const s = normalize(cellVal(row.c[opts.samplesIdx]));
    if (s === 'YES') tr.classList.add('row-sample');
  }

  // HOLD dimming
  if (opts.statusIdx !== undefined) {
    const st = normalize(cellVal(row.c[opts.statusIdx]));
    if (st === 'HOLD') tr.classList.add('row-hold');
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayMs = today.getTime();

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const header = normalize(cols[i]?.label);
    const v = cellVal(row.c[i]);

    // Display
    if (header === 'NEEDED BY') td.textContent = formatDateOnly(v);
    else if (header.includes('TIME') || header.includes('DATE')) td.textContent = formatTimestamp(v);
    else td.textContent = v;

    // NEEDED BY colors (cell only)
    if (opts.neededByIdx !== undefined && i === opts.neededByIdx) {
      const dueMs = parseAnyDateMs(v);
      if (Number.isFinite(dueMs) && dueMs !== Infinity) {
        const daysAway = Math.floor((dueMs - todayMs) / 86400000);
        if (daysAway <= RED_WINDOW_DAYS) td.classList.add('needed-red');
        else if (daysAway <= YELLOW_WINDOW_DAYS) td.classList.add('needed-yellow');
      }
    }

    tr.appendChild(td);
  });

  return tr;
}

// Row builder for Shipping (simple)
function buildRowShipping(row, cols, visibleCols) {
  const tr = document.createElement('tr');

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const header = normalize(cols[i]?.label);
    const v = cellVal(row.c[i]);

    if (header.includes('DATE')) td.textContent = formatDateOnly(v);
    else td.textContent = v;

    tr.appendChild(td);
  });

  return tr;
}

function updateClock() {
  document.getElementById('datetime').textContent = new Date().toLocaleString();
}

/* =========================
   VIEW ROUTING (GUARANTEED)
   ========================= */
let currentView = 'home';

function setView(view) {
  currentView = view;

  // HARD hide everything first (so CSS caching cannot break it)
  [viewHome, viewAction, viewFeed, viewShipping].forEach(v => {
    v.classList.remove('active');
    v.style.display = 'none';
  });

  const map = {
    home: viewHome,
    action: viewAction,
    feed: viewFeed,
    shipping: viewShipping
  };

  const target = map[view] || viewHome;
  target.classList.add('active');
  target.style.display = 'block';

  // Load only what we need
  if (view === 'action') {
    actionContainer.scrollTop = 0;
    loadActionOnly();
  }
  if (view === 'feed') {
    feedContainer.scrollTop = 0;
    loadFeedOnly();
  }
  if (view === 'shipping') {
    shippingContainer.scrollTop = 0;
    loadShippingOnly();
  }
}

// Buttons
btnOpenAction.addEventListener('click', () => setView('action'));
btnOpenFeed.addEventListener('click', () => setView('feed'));
btnOpenShipping.addEventListener('click', () => setView('shipping'));

btnBackAction.addEventListener('click', () => setView('home'));
btnBackFeed.addEventListener('click', () => setView('home'));
btnBackShipping.addEventListener('click', () => setView('home'));

// ESC returns home
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') setView('home');
});

// ===== Loaders =====
async function loadActionOnly() {
  try {
    const parts = await fetchGvizTable(GID_NEW_PARTS);
    const cols = parts.cols || [];
    const rowsAll = (parts.rows || []).slice();
    const map = buildColIndexMap(cols);

    actionHeaders.innerHTML = buildHeader(cols, COLS_NEW_PARTS);
    actionBody.innerHTML = '';

    const idxStatus = map[normalize(COL_STATUS)];
    const idxSamples = map[normalize(COL_SAMPLES)];
    const idxNeeded = map[normalize(COL_NEEDED_BY)];

    const rows = rowsAll.filter(r => {
      if (idxStatus === undefined) return true;
      return normalize(cellVal(r.c[idxStatus])) !== 'DONE';
    });

    rows.sort((a, b) => {
      const sa = idxStatus !== undefined ? statusRank(cellVal(a.c[idxStatus])) : 5;
      const sb = idxStatus !== undefined ? statusRank(cellVal(b.c[idxStatus])) : 5;
      if (sa !== sb) return sa - sb;

      const na = idxNeeded !== undefined ? parseAnyDateMs(cellVal(a.c[idxNeeded])) : Infinity;
      const nb = idxNeeded !== undefined ? parseAnyDateMs(cellVal(b.c[idxNeeded])) : Infinity;
      if (na !== nb) return na - nb;

      const ta = a.c?.[0]?.v ? parseAnyDateMs(a.c[0].v) : 0;
      const tb = b.c?.[0]?.v ? parseAnyDateMs(b.c[0].v) : 0;
      return ta - tb;
    });

    rows.forEach(r => {
      actionBody.appendChild(buildRow(r, cols, COLS_NEW_PARTS, {
        statusIdx: idxStatus,
        samplesIdx: idxSamples,
        neededByIdx: idxNeeded
      }));
    });

    actionCount.textContent = `${rows.length} items`;
  } catch (err) {
    console.error(err);
    actionBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    actionCount.textContent = '—';
  }
}

async function loadFeedOnly() {
  try {
    const live = await fetchGvizTable(GID_LIVE_NOTES);
    const cols = live.cols || [];
    const rowsAll = (live.rows || []).slice();
    const map = buildColIndexMap(cols);

    feedHeaders.innerHTML = buildHeader(cols, COLS_LIVE_NOTES);
    feedBody.innerHTML = '';

    const idxStatus = map[normalize(COL_STATUS)];
    const idxPriority = map[normalize(COL_PRIORITY)];
    const idxNeeded = map[normalize(COL_NEEDED_BY)];

    rowsAll.sort((a, b) => {
      const ta = a.c?.[0]?.v ? parseAnyDateMs(a.c[0].v) : 0;
      const tb = b.c?.[0]?.v ? parseAnyDateMs(b.c[0].v) : 0;
      return tb - ta;
    });

    const rows = rowsAll.filter(r => {
      if (idxStatus === undefined) return true;
      return normalize(cellVal(r.c[idxStatus])) !== 'DONE';
    });

    rows.forEach(r => {
      feedBody.appendChild(buildRow(r, cols, COLS_LIVE_NOTES, {
        statusIdx: idxStatus,
        priorityIdx: idxPriority,
        neededByIdx: idxNeeded
      }));
    });

    feedCount.textContent = `${rows.length} notes`;
  } catch (err) {
    console.error(err);
    feedBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    feedCount.textContent = '—';
  }
}

async function loadShippingOnly() {
  try {
    if (!GID_SHIPPING) {
      throw new Error('Shipping module not configured yet (missing GID_SHIPPING).');
    }

    const ship = await fetchGvizTable(GID_SHIPPING);
    const cols = ship.cols || [];
    const rowsAll = (ship.rows || []).slice();
    const map = buildColIndexMap(cols);

    const visible = cols.map((_, i) => i);
    shippingHeaders.innerHTML = buildHeader(cols, visible);
    shippingBody.innerHTML = '';

    const idxShipDate = map[normalize(COL_SHIP_DATE)];
    if (idxShipDate !== undefined) {
      rowsAll.sort((a, b) => {
        const da = parseAnyDateMs(cellVal(a.c[idxShipDate]));
        const db = parseAnyDateMs(cellVal(b.c[idxShipDate]));
        return da - db;
      });
    }

    rowsAll.forEach(r => shippingBody.appendChild(buildRowShipping(r, cols, visible)));
    shippingCount.textContent = `${rowsAll.length} shipments`;
  } catch (err) {
    console.error(err);
    shippingBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    shippingCount.textContent = '—';
  }
}

// Refresh only active module
function refreshActive() {
  if (currentView === 'action') loadActionOnly();
  if (currentView === 'feed') loadFeedOnly();
  if (currentView === 'shipping') loadShippingOnly();
}

/* =========================
   AUTO-SCROLL (TV-SMOOTH)
   ========================= */
let pauseScrollFeed = false;
let pauseScrollAction = false;

feedContainer.addEventListener('mouseenter', () => pauseScrollFeed = true);
feedContainer.addEventListener('mouseleave', () => pauseScrollFeed = false);

actionContainer.addEventListener('mouseenter', () => pauseScrollAction = true);
actionContainer.addEventListener('mouseleave', () => pauseScrollAction = false);

// TV tuning
const TARGET_FPS = 30;
const FRAME_MS = 1000 / TARGET_FPS;
const FEED_PX_PER_FRAME = 1;
const ACTION_PX_PER_FRAME = 1;

let lastFrameTime = 0;

function tick(now) {
  if (now - lastFrameTime >= FRAME_MS) {
    lastFrameTime = now;

    if (currentView === 'feed' && !pauseScrollFeed) {
      const max = feedContainer.scrollHeight - feedContainer.clientHeight;
      if (max > 0) {
        feedContainer.scrollTop += FEED_PX_PER_FRAME;
        if (feedContainer.scrollTop >= max) feedContainer.scrollTop = 0;
      }
    }

    if (currentView === 'action' && !pauseScrollAction) {
      const maxA = actionContainer.scrollHeight - actionContainer.clientHeight;
      if (maxA > 0) {
        actionContainer.scrollTop += ACTION_PX_PER_FRAME;
        if (actionContainer.scrollTop >= maxA) actionContainer.scrollTop = 0;
      }
    }
  }

  requestAnimationFrame(tick);
}

requestAnimationFrame(tick);

// ===== BOOT =====
// Force initial visibility in case CSS is cached/ignored
[viewAction, viewFeed, viewShipping].forEach(v => v.style.display = 'none');
viewHome.style.display = 'block';

setView('home');
updateClock();
setInterval(updateClock, CLOCK_MS);
setInterval(refreshActive, REFRESH_MS);
