/**************************************************
 * TPI Live Communication Board — Modular UI + Submenus
 * + Schedules connected to Master Schedule spreadsheet
 * + Schedule STATUS pill colors
 * + Total Estimated Hours shown in schedule count
 **************************************************/

/* =========================
   LIVE FEED SHEET (your existing board)
   ========================= */
const sheetID = '1UFkn-d_t3DTt1RCHqp4K3HOuTMyrEVBmZnj1in1PoHc';
const GID_LIVE_NOTES = '863386477';
const GID_NEW_PARTS  = '2113651494';
const GID_SHIPPING = ''; // add later

/* =========================
   MASTER SCHEDULE SHEET (Schedules)
   ========================= */
const SCHEDULE_SHEET_ID = '1fnOGM1YZ2XB0JbomEcZl5JLt6XZnXjVpDJS0Z-9JsgU';
const GID_MASTER_SCHEDULE = '444063954';

// Columns you want (B,C,D,E,F,G,H,I,J,L,N)
const SCHEDULE_VISIBLE_BY_LETTER_INDEX = [1,2,3,4,5,6,7,8,9,11,13];

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

// Schedules columns
const COL_MACHINE = 'MACHINE';
const COL_DUE_DATE = 'DUE DATE';
const COL_EST_HOURS = 'EST TIME (HOURS)'; // your Column I

// ===== Visible columns by index (your existing live board) =====
const COLS_LIVE_NOTES = [0,1,2,3,4,5,6,7,8];
const COLS_NEW_PARTS  = [0,1,2,3,4,5,6,7,8];

/* =========================
   DOM: views
   ========================= */
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

// calculators placeholder
const viewCalcModule = document.getElementById('view-calculator-module');
const calcModuleTitle = document.getElementById('calc-module-title');
const calcModuleSub = document.getElementById('calc-module-sub');
const calcModuleContent = document.getElementById('calc-module-content');


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

// ===== DOM: schedule table =====
const scheduleHeaders = document.getElementById('schedule-headers');
const scheduleBody = document.getElementById('schedule-body');
const scheduleCount = document.getElementById('schedule-count');
const scheduleContainer = document.getElementById('schedule-container');

// ===== Wire form buttons =====
document.getElementById('btn-live-notes').href = FORM_URL_LIVE_NOTES;
document.getElementById('btn-new-parts').href = FORM_URL_NEW_PARTS;

/* =========================
   Helpers
   ========================= */
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

async function fetchGvizTable(sheetId, gid) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${gid}&tqx=out:json&cb=${Date.now()}`;
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

// Row builder for Action/Feed (needed-by colors)
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

/* =========================
   SCHEDULE: Status classification -> CSS classes
   ========================= */
function scheduleStatusClass(statusText) {
  const s = normalize(statusText);
  if (!s) return '';

  // 🔵 ASSIGNED / OWNERS (blue-ish)
  if (
    s === 'GUILLERMO' ||
    s === 'LALO' ||
    s.includes('MIKE') ||
    s === 'JOEY'
  ) return 'status-assigned';

  // 🟡 MATERIAL / PRE-CUT
  if (s.includes('MATERIAL ON ORDER')) return 'status-material';
  if (s.includes('REQUESTED MATERIAL')) return 'status-material';
  if (s.includes('CHECKING STOCK')) return 'status-material';
  if (s.includes('CUSTOMER MTR')) return 'status-material';

  // 🔴 HOLD / PROBLEM
  if (s.includes('ON HOLD')) return 'status-hold';

  // 🟢 READY / CUTTING
  if (s.includes('READY TO CUT')) return 'status-ready';
  if (s.includes('CURRENTLY CUTTING')) return 'status-cutting';

  // 🟠 PARTIAL / ATTENTION
  if (s.includes('PARTIAL')) return 'status-partial';

  // ⚫ COMPLETE
  if (s.includes('COMPLETE')) return 'status-complete';

  return '';
}


// Row builder for Schedules (date-only for DUE DATE, status pill color)
function buildRowSchedule(row, cols, visibleCols, dueDateIdx, statusIdx) {
  const tr = document.createElement('tr');

  visibleCols.forEach(i => {
    const td = document.createElement('td');
    const v = cellVal(row.c[i]);

    // DUE DATE shown as date-only
    if (i === dueDateIdx) td.textContent = formatDateOnly(v);
    else td.textContent = v;

    // STATUS pill styling
    if (i === statusIdx) {
      td.classList.add('status-pill');
      const css = scheduleStatusClass(v);
      if (css) td.classList.add(css);
    }

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
  if (view === 'scheduleModule') {
    scheduleContainer.scrollTop = 0;
    loadScheduleModule();
  }
}

/* =========================
   Buttons: home + back
   ========================= */
document.getElementById('btn-open-action').addEventListener('click', () => setView('action'));
document.getElementById('btn-open-feed').addEventListener('click', () => setView('feed'));
document.getElementById('btn-open-shipping').addEventListener('click', () => setView('shipping'));
document.getElementById('btn-open-capacity').addEventListener('click', () => setView('capacity'));
document.getElementById('btn-open-schedules').addEventListener('click', () => setView('schedules'));
document.getElementById('btn-open-calculators').addEventListener('click', () => setView('calculators'));

document.getElementById('btn-back-action').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-feed').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-shipping').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-capacity').addEventListener('click', () => setView('home'));

document.getElementById('btn-back-schedules').addEventListener('click', () => setView('home'));
document.getElementById('btn-back-calculators').addEventListener('click', () => setView('home'));

document.getElementById('btn-back-schedule-module').addEventListener('click', () => setView('schedules'));
document.getElementById('btn-back-calc-module').addEventListener('click', () => setView('calculators'));

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') setView('home');
});

/* =========================
   SCHEDULES: filtering setup
   ========================= */
const SCHEDULE_FILTERS = {
  'LASER': ['TL2030', 'TL1030 AUTO', 'MS LASER'],
  'EMK': ['EMK'],
  'TRU1000': ['TRU1000'],
  'VASKI': ['VASKI'],
  'PROGRAM': ['PROGRAM'] // update later if needed
};

let currentScheduleKey = null;

function openScheduleModule(name) {
  currentScheduleKey = normalize(name);

  scheduleModuleTitle.textContent = `${name} Schedule`;

  const machines = SCHEDULE_FILTERS[currentScheduleKey] || [];
  if (machines.length) {
    scheduleModuleSub.textContent = `Machine filter: ${machines.join(', ')} (sorted by DUE DATE)`;
  } else {
    scheduleModuleSub.textContent = `Machine filter not configured yet for "${name}".`;
  }

  setView('scheduleModule');
}

// Hook submenu buttons
document.getElementById('btn-open-sched-laser').addEventListener('click', () => openScheduleModule('Laser'));
document.getElementById('btn-open-sched-emk').addEventListener('click', () => openScheduleModule('EMK'));
document.getElementById('btn-open-sched-tru1000').addEventListener('click', () => openScheduleModule('TRU1000'));
document.getElementById('btn-open-sched-vaski').addEventListener('click', () => openScheduleModule('VASKI'));
document.getElementById('btn-open-sched-program').addEventListener('click', () => openScheduleModule('Program'));

/* =========================
   CALCULATORS (placeholders)
   ========================= */
document.getElementById('btn-open-calc-bending').addEventListener('click', () => openCalcModule('Bending'));
document.getElementById('btn-open-calc-countersink').addEventListener('click', () => openCalcModule('Countersink'));
document.getElementById('btn-open-calc-linear').addEventListener('click', () => openCalcModule('Linear Cutting'));

function openCalcModule(name) {
  const key = normalize(name);

  calcModuleTitle.textContent = `${name} Calculator`;
  calcModuleContent.innerHTML = '';

  if (key === 'COUNTERSINK') {
    calcModuleSub.textContent = 'Matches the Excel module: Min Thru Hole, Pre-punch, and OK/NOT check.';
    renderCountersinkCalculator();
  } else {
    calcModuleSub.textContent = `Coming soon. We will build an on-screen calculator for ${name}.`;
    calcModuleContent.innerHTML = `
      <div class="calc-section">
        <div class="calc-output">No UI yet for ${name}.</div>
      </div>
    `;
  }

  setView('calcModule');
}
/* =========================
   CALCULATOR: COUNTERSINK (Excel-matching)
   Excel formulas:
   MinThru = Major - ( tan(Angle/2 in rad) * (Thickness - 0.004) ) * 2
   PrePunch = Major - (Major - Minor) * 0.75
   OK/NOT: Minor > MinThru => OK, else NOT
   ========================= */
function renderCountersinkCalculator() {
  calcModuleContent.innerHTML = `
 <div class="calc-grid">

  <!-- INPUTS -->
  <div class="calc-section">
    <h3>Inputs</h3>

    <!-- Meaning clarification FIRST -->
    <div class="home-sub" style="margin-bottom:10px;">
      <b>Major</b> = Countersink diameter at surface (CSK OD)<br>
      <b>Minor / Thru</b> = Thru hole diameter from customer print
    </div>

    <!-- Actual inputs -->
    <div class="calc-row">
      <label for="csk-major">Major (in)</label>
      <input id="csk-major" type="number" step="0.0001" placeholder="e.g. 0.2500" />
    </div>

    <div class="calc-row">
      <label for="csk-minor">Minor / Thru (in)</label>
      <input id="csk-minor" type="number" step="0.0001" placeholder="e.g. 0.1900" />
    </div>

    <div class="calc-row">
      <label for="csk-angle">Angle (deg)</label>
      <input id="csk-angle" type="number" step="0.1" value="100" />
    </div>

    <div class="calc-row">
      <label for="csk-thk">Material Thk (in)</label>
      <input id="csk-thk" type="number" step="0.0001" placeholder="e.g. 0.0480" />
    </div>

    <!-- Actions -->
    <div class="calc-actions">
      <button class="btn" id="csk-btn-calc" type="button">Calculate</button>
      <button class="btn" id="csk-btn-reset" type="button">Reset</button>
    </div>

    <!-- Notes LAST -->
    <div class="home-sub" style="margin-top:10px;">
      Notes: Please use Tool type 14 for best results.
    </div>
  </div>

  <!-- OUTPUTS -->
  <div class="calc-section">
    <h3>Outputs</h3>

    <div class="calc-output" id="csk-results">
      Enter values and click <b>Calculate</b>.
    </div>
  </div>

</div>

  `;

  const $major = document.getElementById('csk-major');
  const $minor = document.getElementById('csk-minor');
  const $angle = document.getElementById('csk-angle');
  const $thk   = document.getElementById('csk-thk');

  const $btnCalc = document.getElementById('csk-btn-calc');
  const $btnReset = document.getElementById('csk-btn-reset');
  const $results = document.getElementById('csk-results');

  function n(v) {
    const num = Number(v);
    return Number.isFinite(num) ? num : NaN;
  }

  function fmt4(x) {
    return Number.isFinite(x) ? x.toFixed(4) : '';
  }

  function compute() {
    const major = n($major.value);
    const minor = n($minor.value);
    const angle = n($angle.value);
    const thk   = n($thk.value);

    // Basic validation (keep it simple + operator-friendly)
    if (!Number.isFinite(major) || major <= 0) {
      $results.innerHTML = `⚠️ Please enter a valid <b>Major</b>.`;
      return;
    }
    if (!Number.isFinite(minor) || minor <= 0) {
      $results.innerHTML = `⚠️ Please enter a valid <b>Minor/Thru</b>.`;
      return;
    }
    if (!Number.isFinite(angle) || angle <= 0) {
      $results.innerHTML = `⚠️ Please enter a valid <b>Angle</b>.`;
      return;
    }
    if (!Number.isFinite(thk) || thk <= 0) {
      $results.innerHTML = `⚠️ Please enter a valid <b>Material Thickness</b>.`;
      return;
    }

    // Sanity check: thru hole cannot be bigger than countersink major
    if (minor > major) {
  $results.innerHTML = `
    <div style="margin-bottom:10px;">
      Status: <span class="badge badge-not">NOT</span>
    </div>
    ⚠️ <b>Invalid inputs:</b> Minor/Thru (${minor.toFixed(4)}) is larger than Major (${major.toFixed(4)}).<br/>
    A countersink <b>Major (CSK OD)</b> must be ≥ the <b>Thru Hole</b>.
  `;
  return;
}


    // Excel-matching math
    const halfAngleRad = (angle / 2) * Math.PI / 180;
    const minThru = major - (Math.tan(halfAngleRad) * (thk - 0.004)) * 2;
    const prePunch = major - (major - minor) * 0.75;

    const ok = minor > minThru; // matches your Excel behavior

    $results.innerHTML = `
      <div style="margin-bottom:10px;">
        Status: <span class="badge ${ok ? 'badge-ok' : 'badge-not'}">${ok ? 'OK' : 'NOT'}</span>
      </div>

      <div>Min Thru Hole: <b>${fmt4(minThru)}</b> in</div>
      <div>Pre-punch: <b>${fmt4(prePunch)}</b> in</div>

      <hr style="border:0;border-top:1px solid #2d2d2d;margin:12px 0;" />

      <div style="color:#bbb;">
        Rule used: <b>Minor/Thru (${fmt4(minor)})</b> must be &gt; <b>Min Thru Hole (${fmt4(minThru)})</b>
      </div>
    `;
  }

  function reset() {
    $major.value = '';
    $minor.value = '';
    $angle.value = '100';
    $thk.value = '';
    $results.innerHTML = `Enter values and click <b>Calculate</b>.`;
  }

  // Button events
  $btnCalc.addEventListener('click', compute);
  $btnReset.addEventListener('click', reset);

  // Convenience: press Enter to calculate
  [$major, $minor, $angle, $thk].forEach(el => {
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') compute();
    });
  });
}


/* =========================
   LOADERS: Action / Feed / Shipping / Schedule
   ========================= */
async function loadActionOnly() {
  try {
    const parts = await fetchGvizTable(sheetID, GID_NEW_PARTS);
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
    const live = await fetchGvizTable(sheetID, GID_LIVE_NOTES);
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
    shippingBody.innerHTML = `<tr><td colspan="100%">Shipping module not implemented yet.</td></tr>`;
    shippingCount.textContent = '—';
  } catch (err) {
    console.error(err);
    shippingBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    shippingCount.textContent = '—';
  }
}

async function loadScheduleModule() {
  try {
    if (!currentScheduleKey) throw new Error('No schedule selected.');

    const machines = SCHEDULE_FILTERS[currentScheduleKey] || [];
    if (!machines.length) throw new Error(`No machine filter configured for "${currentScheduleKey}".`);

    const table = await fetchGvizTable(SCHEDULE_SHEET_ID, GID_MASTER_SCHEDULE);
    const cols = table.cols || [];
    const rowsAll = (table.rows || []).slice();
    const map = buildColIndexMap(cols);

    const machineIdx = map[normalize(COL_MACHINE)];
    const dueIdx = map[normalize(COL_DUE_DATE)];
    const statusIdx = map[normalize(COL_STATUS)];
    const estIdx = map[normalize(COL_EST_HOURS)];

    if (machineIdx === undefined) throw new Error(`Could not find "${COL_MACHINE}" column in the schedule sheet.`);

    // Visible columns: B,C,D,E,F,G,H,I,J,L,N
    const visibleCols = SCHEDULE_VISIBLE_BY_LETTER_INDEX.filter(i => i < cols.length);

    scheduleHeaders.innerHTML = buildHeader(cols, visibleCols);
    scheduleBody.innerHTML = '';

    // Filter rows by machine
    const wanted = machines.map(normalize);
    const filtered = rowsAll.filter(r => wanted.includes(normalize(cellVal(r.c[machineIdx]))));

    // Sort by DUE DATE
    if (dueIdx !== undefined) {
      filtered.sort((a, b) => {
        const da = parseAnyDateMs(cellVal(a.c[dueIdx]));
        const db = parseAnyDateMs(cellVal(b.c[dueIdx]));
        return da - db;
      });
    }

    // Sum estimated hours (Column I / Est Time (Hours))
    let totalHours = 0;
    if (estIdx !== undefined) {
      filtered.forEach(r => {
        const v = cellVal(r.c[estIdx]);
        const num = Number(v);
        if (Number.isFinite(num)) totalHours += num;
      });
    }

    filtered.forEach(r => {
      scheduleBody.appendChild(buildRowSchedule(r, cols, visibleCols, dueIdx, statusIdx));
    });

    // Show count + hours
    if (estIdx !== undefined) {
      scheduleCount.textContent = `${filtered.length} jobs • ${totalHours.toFixed(2)} hrs`;
    } else {
      scheduleCount.textContent = `${filtered.length} jobs`;
    }

  } catch (err) {
    console.error(err);
    scheduleBody.innerHTML = `<tr><td colspan="100%">⚠️ ${err.message}</td></tr>`;
    scheduleCount.textContent = '—';
  }
}

/* =========================
   Refresh only active view
   ========================= */
function refreshActive() {
  if (currentView === 'action') loadActionOnly();
  if (currentView === 'feed') loadFeedOnly();
  if (currentView === 'scheduleModule') loadScheduleModule();
}

/* =========================
   Auto-scroll (only in visible tables)
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

/* =========================
   BOOT
   ========================= */
ALL_VIEWS.forEach(v => v.style.display = 'none');
viewHome.style.display = 'block';

setView('home');
updateClock();
setInterval(updateClock, CLOCK_MS);
setInterval(refreshActive, REFRESH_MS);
