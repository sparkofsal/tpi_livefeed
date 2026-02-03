/**************************************************
 * TPI Live Communication Board — Modular UI + Submenus
 **************************************************/

// ===== Spreadsheet + GIDs =====
const sheetID = '1UFkn-d_t3DTt1RCHqp4K3HOuTMyrEVBmZnj1in1PoHc';
const GID_LIVE_NOTES = '863386477';
const GID_NEW_PARTS  = '2113651494';
const GID_SHIPPING = ''; // add later when you have it

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
const COL_SHIP_DATE = 'SHIP DATE';

// ===== Visible columns by index =====
const COLS_LIVE_NOTES = [0,1,2,3,4,5,6,7,8];
const COLS_NEW_PARTS  = [0,1,2,3,4,5,6,7,8];

// ===== DOM: views =====
const viewHome = document.getElementById('view-home');
const viewAction = document.getElementById('view-action');
const viewFeed = document.getElementById('view-feed');
const viewShipping = document.getElementById('view-shipping');
const viewCapacity = document.getElementById('view-capacity');

const viewSchedules = document.getElementById('view-schedules');
const viewCalculators = document.getElementById('view-calculators');

const viewScheduleModule = document.getElementById('view-schedule-module');
const scheduleModuleTitle = document.getElementById('schedule-module-title');
const scheduleModuleSub = document.getElementById('schedule-module-sub');

const viewCalcModule = document.getElementById('view-calculator-module');
const calcModuleTitle = document.getElementById('calc-module-title');
const calcModuleSub = document.getElementById('calc-module-sub');

// ===== DOM: buttons =====
document.getElementById('btn-live-notes').href = FORM_URL_LIVE_NOTES;
document.getElementById('btn-new-parts').href = FORM_URL_NEW_PARTS;

// Home
document.getElementById('btn-open-action').addEventListener('click', () => setView('action'));
document.getElementById('btn-open-feed').addEventListener('click', () => setView('feed'));
document.getElementById('btn-open-shipping').addEventListener('click', () => setView('shipping'));
document.getElementById('btn-open-capacity').addEventListener('click', () => setView('capacity'));
document.getElementById('btn-open-schedules').addEventListener('click', () => setView('schedules'));
document.getElementById('btn-open-calculators').addEventListener('click', () => setView('calculators'));

// Back buttons
document.getElementById('btn-back-action').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-feed').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-shipping').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-capacity').addEventListener('click', () => setView('home'));

document.getElementById('btn-back-schedules').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-calculators').addEventListener('click', () => setView('home'));

document.getElementById('btn-back-schedule-module').addEventListener('click', () => setView('schedules'));
document.getElementById('btn-back-calc-module').addEventListener('click', () => setView('calculators'));

// Schedules submenu
document.getElementById('btn-open-sched-laser').addEventListener('click', () => openScheduleModule('Laser'));
document.getElementById('btn-open-sched-emk').addEventListener('click', () => openScheduleModule('EMK'));
document.getElementById('btn-open-sched-tru1000').addEventListener('click', () => openScheduleModule('TRU1000'));
document.getElementById('btn-open-sched-vaski').addEventListener('click', () => openScheduleModule('VASKI'));
document.getElementById('btn-open-sched-program').addEventListener('click', () => openScheduleModule('Program'));

// Calculators submenu
document.getElementById('btn-open-calc-bending').addEventListener('click', () => openCalcModule('Bending'));
document.getElementById('btn-open-calc-countersink').addEventListener('click', () => openCalcModule('Countersink'));
document.getElementById('btn-open-calc-linear').addEventListener('click', () => openCalcModule('Linear Cutting'));

// ESC always returns to home
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') setView('home');
});

// ===== DOM: tables =====
const actionHeaders = document.getElementById('action-headers');
const actionBody = document.getElementById('action-body');
const actionCount = document.getElementById('action-count');
const actionContainer = document.getElementById('action-container');

const feedHeaders = document.getElementById('feed-headers');
const feedBody = document.getElementById('feed-body');
const feedCount = document.getElementById('feed-count');
const feedContainer = document.getElementById('feed-container');

const shippingHeaders = document.getElementById('shipping-headers');
const shippingBody = document.getElementById('shipping-body');
const shippingCount = document.getElementById('shipping-count');
const shippingContainer = document.getElementById('shipping-container');

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
  if (start === -1 || end === -1) throw new Error('No gviz JSON returned. Check Sheet sharing.');
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

  const today = new Date();
  today.setHours(0,0,0,0);
  const todayMs = today.getTime();

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const header = normalize(cols[i]?.label);
    const v = cellVal(row.c[i]);

    if (header === 'NEEDED BY') td.textContent = formatDateOnly(v);
    else if (header.includes('TIME') || header.includes('DATE')) td.textContent = formatTimestamp(v);
    else td.textContent = v;

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

function buildRowSimple(row, cols, visibleCols) {
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

const ALL_VIEWS = [
  viewHome, viewAction, viewFeed, viewShipping, viewCapacity,
  viewSchedules, viewCalculators, viewScheduleModule, viewCalcModule
];

function setView(view) {
  currentView = view;

  ALL_VIEWS.forEach(v => {
    v.classList.remove('active');
    v.style.display = 'none';
  });

  const map = {
    home: viewHome,
    action: viewAction,
    feed: viewFeed,
    shipping: viewShipping,
    capacity: viewCapacity,
    schedules: viewSchedules,
    calculators: viewCalculators,
    scheduleModule: viewScheduleModule,
    calcModule: viewCalcModule
  };

  const target = map[view] || viewHome;
  target.classList.add('active');
  target.style.display = 'block';

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

// open submenu modules (placeholders for now)
function openScheduleModule(name) {
  scheduleModuleTitle.textContent = `${name} Schedule`;
  scheduleModuleSub.textContent = `Coming soon. This will connect to a Google Sheet tab for ${name}.`;
  setView('scheduleModule');
}

function openCalcModule(name) {
  calcModuleTitle.textContent = `${name} Calculator`;
  calcModuleSub.textContent = `Coming soon. We will build an on-screen calculator for ${name}.`;
  setView('calcModule');
}

/* ===== Loading for existing sheet modules ===== */
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

    const rows = rowsAll.filter(r => idxStatus === undefined || normalize(cellVal(r.c[idxStatus])) !== 'DONE');

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

    const rows = rowsAll.filter(r => idxStatus === undefined || normalize(cellVal(r.c[idxStatus])) !== 'DONE');

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
    if (!GID_SHIPPING) throw new Error('Shipping module not configured yet (missing GID_SHIPPING).');

    const ship = await fetchGvizTable(GID_SHIPPING);
    const cols = ship.cols || [];
    const rowsAll = (ship.rows || []).slice();
    const map = buildColIndexMap(cols);

    const visible = cols.map((_, i) => i);
    shippingHeaders.innerHTML = buildHeader(cols, visible);
    shippingBody.innerHTML = '';

    const idxShipDate = map[normalize(COL_SHIP_DATE)];
    if (idxShipDate !== undefined) {
      rowsAll.sort((a, b) => parseAnyDateMs(cellVal(a.c[idxShipDate])) - parseAnyDateMs(cellVal(b.c[idxShipDate])));
    }

    rowsAll.forEach(r => shippingBody.appendChild(buildRowSimple(r, cols, visible)));
    shippingCount.textContent = `${rowsAll.length} shipments`;
  } catch (err) {
    console.error(err);
    shippingBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    shippingCount.textContent = '—';
  }
}

// Refresh only active table modules
function refreshActive() {
  if (currentView === 'action') loadActionOnly();
  if (currentView === 'feed') loadFeedOnly();
  if (currentView === 'shipping') loadShippingOnly();
}

/* ===== Auto-scroll (TV smooth, only in visible tables) ===== */
let pauseScrollFeed = false;
let pauseScrollAction = false;

feedContainer.addEventListener('mouseenter', () => pauseScrollFeed = true);
feedContainer.addEventListener('mouseleave', () => pauseScrollFeed = false);

actionContainer.addEventListener('mouseenter', () => pauseScrollAction = true);
actionContainer.addEventListener('mouseleave', () => pauseScrollAction = false);

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
ALL_VIEWS.forEach(v => v.style.display = 'none');
viewHome.style.display = 'block';

setView('home');
updateClock();
setInterval(updateClock, CLOCK_MS);
setInterval(refreshActive, REFRESH_MS);
