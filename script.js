/**************************************************
 * TPI Live Communication Board — Modular UI + Submenus
 * + Schedules connected to Master Schedule spreadsheet
 * + Schedule STATUS pill colors
 * + Total Estimated Hours shown in schedule count
 * + Shipping Schedule connected + sorted by Ship Date
 * + Robust header matching (prevents breakage if a header changes slightly)
 **************************************************/

/* =========================
   LIVE FEED SHEET (your existing board)
   ========================= */
const sheetID = '1UFkn-d_t3DTt1RCHqp4K3HOuTMyrEVBmZnj1in1PoHc';
const GID_LIVE_NOTES = '863386477';
const GID_NEW_PARTS  = '2113651494';

/* =========================
   MASTER SCHEDULE SHEET (Schedules + Shipping)
   ========================= */
const SCHEDULE_SHEET_ID = '1fnOGM1YZ2XB0JbomEcZl5JLt6XZnXjVpDJS0Z-9JsgU';
const GID_MASTER_SCHEDULE = '444063954';
const GID_SHIPPING = '86241306'; // ✅ Shipping Schedule tab gid from your link

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

/* =========================
   SHIPPING: columns + defaults
   ========================= */
const SHIPPING_VISIBLE_COLS = [0,1,2,3,4,5,6,7,8,9,11,13]; // A..N

// Shipping column names (exactly as your headers)
const SH_COL_PO = 'PO #';
const SH_COL_JOB = 'JOB #';
const SH_COL_CUSTOMER = 'CUSTOMER';
const SH_COL_PART = 'PART#';
const SH_COL_DESC = 'DESCRIPTION';
const SH_COL_ORDER_DATE = 'ORDER_DATE';
const SH_COL_DUE_DATE = 'DUE DATE';
const SH_COL_TYPE = 'TYPE';
const SH_COL_QTY_DUE = 'QTY DUE';
const SH_COL_SHIP_DATE = 'SHIP DATE';

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

// ===== DOM: capacity =====
const capacityContent = document.getElementById('capacity-content');
const capacityCount = document.getElementById('capacity-count');

/* =========================
   CAPACITY CONFIG
   ========================= */

// Default planning horizon (workdays)
const CAPACITY_WORKDAYS_AHEAD = 5;

// Define capacity by module (shifts + hrs per shift)
const CAPACITY_MAP = {
  LASER:   { shifts: 2, hoursPerShift: 8 },
  EMK:     { shifts: 2, hoursPerShift: 8 },
  TRU1000: { shifts: 2, hoursPerShift: 8 },
  VASKI:   { shifts: 2, hoursPerShift: 8 },
  PROGRAM: { shifts: 1, hoursPerShift: 8 },
};

function safeNum(v) {
  const raw = String(v ?? '').replace(/,/g, '').trim();
  const n = parseFloat(raw);
  return Number.isFinite(n) ? n : 0;
}

function startOfDayMs(ms) {
  const d = new Date(ms);
  d.setHours(0,0,0,0);
  return d.getTime();
}

function isWorkday(dateObj) {
  const day = dateObj.getDay(); // 0 Sun ... 6 Sat
  return day !== 0 && day !== 6;
}

function getNextWorkdayDates(count) {
  const out = [];
  const d = new Date();
  d.setHours(0,0,0,0);

  while (out.length < count) {
    if (isWorkday(d)) out.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function withinNextWorkdays(dueMs, workdayDates) {
  if (!Number.isFinite(dueMs) || dueMs === Infinity) return false;
  const dueDay = startOfDayMs(dueMs);
  const set = new Set(workdayDates.map(dt => dt.getTime()));
  return set.has(dueDay);
}

async function loadCapacityOnly() {
  try {
    if (!capacityContent) throw new Error('Missing #capacity-content element in HTML.');
    if (!capacityCount) throw new Error('Missing #capacity-count element in HTML.');

    // Pull schedule data once
    const table = await fetchGvizTable(SCHEDULE_SHEET_ID, GID_MASTER_SCHEDULE);
    const cols = table.cols || [];
    const rowsAll = (table.rows || []).slice();
    const map = buildColIndexMap(cols);

    // Find required columns (robust)
    const machineIdx = findColIndex(map, [COL_MACHINE, 'MACHINES']);
    const dueIdx = findColIndex(map, [COL_DUE_DATE, 'DUE', 'DUE_DT', 'DUE DATE ', 'DUE DATE:']);
    const estIdx = findColIndex(map, [
      COL_EST_HOURS,
      'EST TIME',
      'EST TIME (HRS)',
      'EST HOURS',
      'EST. TIME (HOURS)',
      'ESTIMATED HOURS',
      'EST HRS'
    ]);

    if (machineIdx === undefined) throw new Error(`Could not find "${COL_MACHINE}" in Master Schedule.`);
    if (estIdx === undefined) throw new Error(`Could not find "${COL_EST_HOURS}" in Master Schedule.`);
    if (dueIdx === undefined) throw new Error(`Could not find "${COL_DUE_DATE}" in Master Schedule.`);

    // Workday window (next N workdays)
    const workdayDates = getNextWorkdayDates(CAPACITY_WORKDAYS_AHEAD);

    // Aggregate scheduled hours per module (LASER/EMK/etc)
    const results = {};
    Object.keys(SCHEDULE_FILTERS).forEach(k => {
      results[k] = { scheduledHrs: 0, dueSoonHrs: 0, totalJobs: 0 };
    });

    rowsAll.forEach(r => {
      const machineVal = normalize(cellVal(r.c[machineIdx]));
      const dueMs = parseAnyDateMs(cellVal(r.c[dueIdx]));
      const hrs = safeNum(cellVal(r.c[estIdx]));

      if (!Number.isFinite(hrs) || hrs <= 0) return;

      // Assign row to the first matching module key
      for (const key of Object.keys(SCHEDULE_FILTERS)) {
        const wanted = (SCHEDULE_FILTERS[key] || []).map(normalize);
        if (wanted.includes(machineVal)) {
          results[key].scheduledHrs += hrs;
          results[key].totalJobs += 1;

          if (withinNextWorkdays(dueMs, workdayDates)) {
            results[key].dueSoonHrs += hrs;
          }
          break;
        }
      }
    });

    // Render cards
    let html = `<div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap:12px;">`;

    const keysInOrder = ['LASER','EMK','TRU1000','VASKI','PROGRAM'];

    keysInOrder.forEach(key => {
      const cap = CAPACITY_MAP[key] || { shifts: 1, hoursPerShift: 8 };
      const capHrs = cap.shifts * cap.hoursPerShift * CAPACITY_WORKDAYS_AHEAD;

      const sched = results[key]?.scheduledHrs || 0;
      const dueSoon = results[key]?.dueSoonHrs || 0;
      const jobs = results[key]?.totalJobs || 0;

      const util = capHrs > 0 ? (sched / capHrs) : 0;
      const utilPct = Math.round(util * 100);
      const remaining = capHrs - sched;

      const over = sched > capHrs;

      // progress bar
      const barPct = Math.max(0, Math.min(100, util * 100));

      html += `
        <div style="border:1px solid #2d2d2d; background:#111; border-radius:12px; padding:12px;">
          <div style="display:flex; align-items:center; justify-content:space-between; gap:10px;">
            <div style="font-weight:900; font-size:1.05rem;">${key}</div>
            <div style="font-weight:900; padding:4px 8px; border-radius:999px; background:${over ? '#3b0f0f' : '#0f2c14'}; color:${over ? '#ffb3b3' : '#b9ffd0'};">
              ${over ? 'OVER' : 'OK'}
            </div>
          </div>

          <div style="margin-top:8px; color:#bbb; font-weight:800;">
            ${sched.toFixed(1)} / ${capHrs.toFixed(1)} hrs • ${cap.shifts} shift(s)
          </div>

          <div style="margin-top:8px; height:10px; background:#1f1f1f; border-radius:999px; overflow:hidden;">
            <div style="height:10px; width:${barPct}%; background:#7aa6ff;"></div>
          </div>

          <div style="margin-top:8px; display:flex; justify-content:space-between; color:#ddd;">
            <div><b>${utilPct}%</b> utilized</div>
            <div>${remaining >= 0 ? `<b>${remaining.toFixed(1)}</b> hrs left` : `<b>${Math.abs(remaining).toFixed(1)}</b> hrs over`}</div>
          </div>

          <div style="margin-top:10px; color:#bbb;">
            Due in next <b>${CAPACITY_WORKDAYS_AHEAD}</b> workdays: <b>${dueSoon.toFixed(1)}</b> hrs
          </div>

          <div style="margin-top:6px; color:#bbb;">
            Jobs counted: <b>${jobs}</b>
          </div>
        </div>
      `;
    });

    html += `</div>`;

    capacityContent.innerHTML = html;
    capacityCount.textContent = `Next ${CAPACITY_WORKDAYS_AHEAD} workdays • scheduled EST hours vs capacity`;

  } catch (err) {
    console.error(err);
    if (capacityContent) capacityContent.innerHTML = `⚠️ ${err.message}`;
    if (capacityCount) capacityCount.textContent = '—';
  }
}


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

// ✅ Use formatted value when Google provides it (GViz uses .f for pretty display)
const cellDisplay = c => (c?.f ?? c?.v ?? '');

// ✅ Extract first date from text like "shipping 2/12-2/20" or "2/13"
function extractFirstDateMs(text, fallbackYear) {
  const s = String(text ?? '');
  // matches: 2/13, 02/13, 2-13, 2/13/2026, etc.
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?/);
  if (!m) return Infinity;

  const mm = Number(m[1]);
  const dd = Number(m[2]);
  let yy = m[3] ? Number(m[3]) : Number(fallbackYear);

  if (!Number.isFinite(mm) || !Number.isFinite(dd)) return Infinity;
  if (!Number.isFinite(yy)) yy = new Date().getFullYear();
  if (yy < 100) yy = 2000 + yy; // handle 26 -> 2026

  const dt = new Date(yy, mm - 1, dd);
  const ms = dt.getTime();
  return Number.isFinite(ms) ? ms : Infinity;
}

// ✅ Ship date can be a real date OR text notes. This returns a sortable/filterable ms.
function shipCellToMs(shipCell, dueCellForYear) {
  // First: try "real" date value (works for Date(yyyy,mm,dd) and normal date strings)
  const raw = cellVal(shipCell);
  const ms1 = parseAnyDateMs(raw);
  if (Number.isFinite(ms1) && ms1 !== Infinity) return ms1;

  // Second: try formatted display (often "2/13/2026")
  const disp = cellDisplay(shipCell);
  const ms2 = parseAnyDateMs(disp);
  if (Number.isFinite(ms2) && ms2 !== Infinity) return ms2;

  // Third: parse first date inside text like "shipping 2/12-2/20"
  const dueMs = parseAnyDateMs(cellVal(dueCellForYear));
  const fallbackYear = Number.isFinite(dueMs) && dueMs !== Infinity
    ? new Date(dueMs).getFullYear()
    : new Date().getFullYear();

  return extractFirstDateMs(disp, fallbackYear);
}

// ✅ What we SHOW in the SHIP DATE column:
// - if it’s a real date → show mm/dd/yyyy
// - if it’s text → show exactly the text
function formatShipDisplay(shipCell) {
  const raw = cellVal(shipCell);
  const ms = parseAnyDateMs(raw);
  if (Number.isFinite(ms) && ms !== Infinity) return new Date(ms).toLocaleDateString();

  const disp = cellDisplay(shipCell);
  const ms2 = parseAnyDateMs(disp);
  if (Number.isFinite(ms2) && ms2 !== Infinity) return new Date(ms2).toLocaleDateString();

  return String(disp ?? '');
}


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

/**
 * ✅ NEW: Find a column index using multiple acceptable header names.
 * Prevents modules from breaking if a header changes slightly in Google Sheets.
 */
function findColIndex(map, possibleNames = []) {
  for (const name of possibleNames) {
    const idx = map[normalize(name)];
    if (idx !== undefined) return idx;
  }
  return undefined;
}

function buildHeader(cols, visibleCols) {
  return visibleCols.map(i => `<th>${cols[i]?.label ?? ''}</th>`).join('');
}
/* =========================
   SHIPPING FILTERS (per-device)
   - Customer dropdown
   - Type dropdown
   - Text search (PO/Job/Part/Desc/Notes/etc)
   - Due Date range
   - "Open only" (Qty Due > 0)
   ========================= */
let shippingRaw = null;      // { cols, rowsAll, map }
let shippingFiltered = [];   // rows after filter
let shippingFilterEls = null;

const SHIPPING_FILTER_STORAGE_KEY = 'tpi_shipping_filters_v1';

function loadShippingFiltersFromStorage() {
  try {
    const raw = localStorage.getItem(SHIPPING_FILTER_STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function saveShippingFiltersToStorage(state) {
  try {
    localStorage.setItem(SHIPPING_FILTER_STORAGE_KEY, JSON.stringify(state));
  } catch {}
}

function ensureShippingFilterBar() {
  if (shippingFilterEls) return; // already created

  // Build a filter bar and inject it ABOVE the shipping table (inside view-shipping panel)
  const panel = document.getElementById('panel-shipping');
  if (!panel) return;

  // Place under title section (after the .panel-title)
  const title = panel.querySelector('.panel-title');
  if (!title) return;

  const bar = document.createElement('div');
  bar.id = 'shipping-filter-bar';
  bar.style.padding = '10px 12px';
  bar.style.borderBottom = '1px solid #333';
  bar.style.background = '#111';

  bar.innerHTML = `
    <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end;">
      
      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Customer</label>
        <select id="ship-filter-customer" class="btn" style="min-width:220px;"></select>
      </div>

      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Type</label>
        <select id="ship-filter-type" class="btn" style="min-width:160px;"></select>
      </div>

      <div style="display:flex; flex-direction:column; gap:6px; flex:1; min-width:240px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Search (PO / Job / Part / Desc / Notes)</label>
        <input id="ship-filter-search" class="btn" style="width:100%; text-align:left;" placeholder="Type to filter..." />
      </div>

      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Due Date From</label>
        <input id="ship-filter-from" type="date" class="btn" />
      </div>

      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Due Date To</label>
        <input id="ship-filter-to" type="date" class="btn" />
      </div>

      <!-- ✅ NEW: Ship Date range -->
      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Ship Date From</label>
        <input id="ship-filter-ship-from" type="date" class="btn" />
      </div>

      <div style="display:flex; flex-direction:column; gap:6px;">
        <label style="color:#bbb; font-weight:800; font-size:0.85rem;">Ship Date To</label>
        <input id="ship-filter-ship-to" type="date" class="btn" />
      </div>

      <div style="display:flex; gap:10px; align-items:center; padding-bottom:2px;">
        <label style="display:flex; gap:8px; align-items:center; color:#ddd; font-weight:800;">
          <input id="ship-filter-openonly" type="checkbox" />
          Open only (Qty Due &gt; 0)
        </label>
      </div>

      <div style="display:flex; gap:10px; padding-bottom:2px;">
        <button id="ship-filter-clear" class="btn" type="button">Clear</button>
      </div>

    </div>
  `;

  // Insert after title bar
  title.insertAdjacentElement('afterend', bar);

  // ===== DOM refs =====
  const $customer = document.getElementById('ship-filter-customer');
  const $type = document.getElementById('ship-filter-type');
  const $search = document.getElementById('ship-filter-search');
  const $from = document.getElementById('ship-filter-from');
  const $to = document.getElementById('ship-filter-to');

  // ✅ NEW ship date refs
  const $shipFrom = document.getElementById('ship-filter-ship-from');
  const $shipTo = document.getElementById('ship-filter-ship-to');

  const $open = document.getElementById('ship-filter-openonly');
  const $clear = document.getElementById('ship-filter-clear');

  // ✅ Store all refs in one place
  shippingFilterEls = { $customer, $type, $search, $from, $to, $shipFrom, $shipTo, $open, $clear };

  // ===== Restore saved state (per device) =====
  const saved = loadShippingFiltersFromStorage();
  if (saved) {
    $search.value = saved.search ?? '';
    $from.value = saved.from ?? '';
    $to.value = saved.to ?? '';

    // ✅ NEW: restore ship date range
    $shipFrom.value = saved.shipFrom ?? '';
    $shipTo.value = saved.shipTo ?? '';

    $open.checked = !!saved.openOnly;
    // customer/type restored after dropdowns are populated
  }

  function onChange() {
    const state = getShippingFilterState();
    saveShippingFiltersToStorage(state);
    renderShippingFiltered(); // rerender only shipping rows
  }

  // ===== Listeners =====
  $customer.addEventListener('change', onChange);
  $type.addEventListener('change', onChange);
  $search.addEventListener('input', onChange);

  $from.addEventListener('change', onChange);
  $to.addEventListener('change', onChange);

  // ✅ NEW
  $shipFrom.addEventListener('change', onChange);
  $shipTo.addEventListener('change', onChange);

  $open.addEventListener('change', onChange);

  // ===== Clear button =====
  $clear.addEventListener('click', () => {
    $customer.value = '__ALL__';
    $type.value = '__ALL__';
    $search.value = '';

    $from.value = '';
    $to.value = '';

    // ✅ NEW
    $shipFrom.value = '';
    $shipTo.value = '';

    $open.checked = false;
    saveShippingFiltersToStorage(getShippingFilterState());
    renderShippingFiltered();
  });
}


function getShippingFilterState() {
  if (!shippingFilterEls) return {};
  return {
    customer: shippingFilterEls.$customer.value,
    type: shippingFilterEls.$type.value,
    search: shippingFilterEls.$search.value,
    from: shippingFilterEls.$from.value,
    to: shippingFilterEls.$to.value,
    shipFrom: shippingFilterEls.$shipFrom.value,
    shipTo: shippingFilterEls.$shipTo.value,
    openOnly: shippingFilterEls.$open.checked
  };
}


function setShippingDropdownOptions($select, values, savedValue) {
  const cur = savedValue || $select.value;
  $select.innerHTML = '';

  const optAll = document.createElement('option');
  optAll.value = '__ALL__';
  optAll.textContent = 'All';
  $select.appendChild(optAll);

  values.forEach(v => {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = v;
    $select.appendChild(o);
  });

  // restore if exists
  if (cur && Array.from($select.options).some(o => o.value === cur)) {
    $select.value = cur;
  } else {
    $select.value = '__ALL__';
  }
}

function ymdToMs(ymd) {
  if (!ymd) return NaN;
  const [y,m,d] = ymd.split('-').map(Number);
  if (!y || !m || !d) return NaN;
  return new Date(y, m - 1, d).setHours(0,0,0,0);
}

function rowTextBlob(row) {
  const cells = SHIPPING_VISIBLE_COLS.map(i => String(cellVal(row.c[i]) ?? ''));
  return normalize(cells.join(' | '));
}

function numFromCell(v) {
  const n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function renderShippingFiltered() {
  if (!shippingRaw) return;

  const { cols, rowsAll, map } = shippingRaw;

  shippingHeaders.innerHTML = buildHeader(cols, SHIPPING_VISIBLE_COLS);
  shippingBody.innerHTML = '';

  const idxCustomer = map[normalize(SH_COL_CUSTOMER)];
  const idxType = map[normalize(SH_COL_TYPE)];
  const idxDue = map[normalize(SH_COL_DUE_DATE)];
  const idxOrder = map[normalize(SH_COL_ORDER_DATE)];
  const idxQtyDue = map[normalize(SH_COL_QTY_DUE)];
  const idxShip = map[normalize(SH_COL_SHIP_DATE)];


  const state = getShippingFilterState();
  const wantCustomer = normalize(state.customer);
  const wantType = normalize(state.type);
  const q = normalize(state.search);

  const fromMs = ymdToMs(state.from);
  const toMs = ymdToMs(state.to);

  let filtered = rowsAll.slice();

  if (idxCustomer !== undefined && wantCustomer && wantCustomer !== '__ALL__') {
    filtered = filtered.filter(r => normalize(cellVal(r.c[idxCustomer])) === wantCustomer);
  }

  if (idxType !== undefined && wantType && wantType !== '__ALL__') {
    filtered = filtered.filter(r => normalize(cellVal(r.c[idxType])) === wantType);
  }

  if (state.openOnly && idxQtyDue !== undefined) {
    filtered = filtered.filter(r => {
      const v = cellVal(r.c[idxQtyDue]);
      const n = numFromCell(v);
      return Number.isFinite(n) && n > 0;
    });
  }

  if (idxDue !== undefined && (Number.isFinite(fromMs) || Number.isFinite(toMs))) {
    filtered = filtered.filter(r => {
      const dueMs = parseAnyDateMs(cellVal(r.c[idxDue]));
      if (!Number.isFinite(dueMs) || dueMs === Infinity) return false;
      const d0 = new Date(dueMs); d0.setHours(0,0,0,0);
      const ms = d0.getTime();

      if (Number.isFinite(fromMs) && ms < fromMs) return false;
      if (Number.isFinite(toMs) && ms > toMs) return false;
      return true;
    });
  }

  if (q) {
    filtered = filtered.filter(r => rowTextBlob(r).includes(q));
  }

  filtered.sort((a, b) => {
    const da = idxDue !== undefined ? parseAnyDateMs(cellVal(a.c[idxDue])) : Infinity;
    const db = idxDue !== undefined ? parseAnyDateMs(cellVal(b.c[idxDue])) : Infinity;
    if (da !== db) return da - db;

    const oa = idxOrder !== undefined ? parseAnyDateMs(cellVal(a.c[idxOrder])) : Infinity;
    const ob = idxOrder !== undefined ? parseAnyDateMs(cellVal(b.c[idxOrder])) : Infinity;
    if (oa !== ob) return oa - ob;

    return rowTextBlob(a).localeCompare(rowTextBlob(b));
  });

  const idxOrderDate = idxOrder;
  const idxDueDate = idxDue;

  filtered.forEach(r => {
    const tr = document.createElement('tr');

      SHIPPING_VISIBLE_COLS.forEach(i => {
    const td = document.createElement('td');

    // ✅ Order/Due still date-only
    if (i === idxOrderDate || i === idxDueDate) {
      td.textContent = formatDateOnly(cellVal(r.c[i]));

    // ✅ SHIP DATE: show date if it’s a real date, otherwise show text notes
    } else if (idxShip !== undefined && i === idxShip) {
      td.textContent = formatShipDisplay(r.c[i]);

    // ✅ Everything else: use .f when available, else .v
    } else {
      td.textContent = cellDisplay(r.c[i]);
    }

    tr.appendChild(td);
  });


    shippingBody.appendChild(tr);
  });

  shippingCount.textContent = `${filtered.length} rows`;
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
  if (view === 'capacity') {
    loadCapacityOnly();
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
   CALCULATORS
   ========================= */
document.getElementById('btn-open-calc-bending').addEventListener('click', () => openCalcModule('Bending'));
document.getElementById('btn-open-calc-countersink').addEventListener('click', () => openCalcModule('Countersink'));
document.getElementById('btn-open-calc-linear').addEventListener('click', () => openCalcModule('Linear Cutting'));

function openCalcModule(name) {
  const key = normalize(name);

  calcModuleTitle.textContent = `${name} Calculator`;
  calcModuleContent.innerHTML = '';

  if (key === 'BENDING') {
    calcModuleSub.textContent = 'Standard bend math: BA, BD, Setback, and Flat dimension (1 bend) + Multi-bend flat.';
    renderBendingCalculator();

  } else if (key === 'COUNTERSINK') {
    calcModuleSub.textContent = 'Matches my Excel module: Min Thru Hole, Pre-punch, and OK/NOT check.';
    renderCountersinkCalculator();

  } else {
    calcModuleSub.textContent = `Coming soon. ${name}.`;
    calcModuleContent.innerHTML = `
      <div class="calc-section">
        <div class="calc-output">No UI yet for ${name}.</div>
      </div>
    `;
  }

  setView('calcModule');
}

/* =========================
   CALCULATOR: BENDING
   (your original function — unchanged)
   ========================= */
function renderBendingCalculator() {
  // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  // YOUR ORIGINAL BENDING CALCULATOR CODE:
  // (This is exactly what you pasted earlier — kept unchanged.)
  // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  calcModuleContent.innerHTML = `
    <div class="calc-grid">

      <!-- INPUTS -->
      <div class="calc-section">
        <h3>Inputs</h3>

        <div class="home-sub" style="margin-bottom:10px;">
          <b>T</b> = Material thickness (in)<br>
          <b>R</b> = Inside bend radius (in)<br>
          <b>A</b> = Bend angle in degrees (e.g., 90)<br>
          <b>K</b> = K-factor (TPI standard: 0.434)
        </div>

        <!-- Toggle: Single vs Multi -->
        <div class="calc-toggle">
          <button class="btn-toggle active" id="bend-mode-single" type="button">Single Bend</button>
          <button class="btn-toggle" id="bend-mode-multi" type="button">Multi Bend</button>
        </div>

        <!-- Global defaults used for both modes -->
        <div class="calc-row">
          <label for="bend-thk">Material Thk, T (in)</label>
          <input id="bend-thk" type="number" step="0.0001" placeholder="e.g. 0.0480" />
        </div>

        <div class="calc-row">
          <label for="bend-rad">Inside Radius, R (in)</label>
          <input id="bend-rad" type="number" step="0.0001" placeholder="e.g. 0.0625" />
        </div>

        <div class="calc-row">
          <label for="bend-k">K-Factor, K</label>
          <input id="bend-k" type="number" step="0.001" value="0.434" />
        </div>

        <hr style="border:0;border-top:1px solid #2d2d2d;margin:12px 0;" />

        <!-- SINGLE BEND BLOCK -->
        <div id="bend-block-single">
          <h3 style="margin-top:0;">Single Bend</h3>

          <div class="calc-row">
            <label for="bend-ang">Bend Angle, A (deg)</label>
            <input id="bend-ang" type="number" step="0.1" value="90" />
          </div>

          <h3 style="margin-top:0;">Flat (1 bend)</h3>
          <div class="home-sub" style="margin-bottom:10px;">
            Enter outside flange dimensions (optional).<br>
            Flat = L1 + L2 − BD
          </div>

          <div class="calc-row">
            <label for="bend-l1">Flange L1 (in)</label>
            <input id="bend-l1" type="number" step="0.0001" placeholder="e.g. 1.2500" />
          </div>

          <div class="calc-row">
            <label for="bend-l2">Flange L2 (in)</label>
            <input id="bend-l2" type="number" step="0.0001" placeholder="e.g. 0.7500" />
          </div>
        </div>

        <!-- MULTI BEND BLOCK -->
        <div id="bend-block-multi" style="display:none;">
          <h3 style="margin-top:0;">Multi Bend (outside dims)</h3>

          <div class="home-sub" style="margin-bottom:10px;">
            Use this when the part has multiple bends.<br>
            Flat = (sum of outside dimensions) − (sum of BD for each bend)
          </div>

          <div class="calc-row">
            <label for="bend-multi-count">Number of bends</label>
            <input id="bend-multi-count" type="number" step="1" value="2" />
          </div>

          <div class="calc-actions">
            <button class="btn" id="bend-multi-build" type="button">Build Inputs</button>
          </div>

          <div id="bend-multi-inputs" style="margin-top:10px;"></div>

          <div class="home-sub" style="margin-top:10px;">
            Notes: My formula states a global T, R, K for all bends. No Radius or KFactor per bend yet.
          </div>
        </div>

        <div class="calc-actions" style="margin-top:14px;">
          <button class="btn" id="bend-btn-calc" type="button">Calculate</button>
          <button class="btn" id="bend-btn-reset" type="button">Reset</button>
        </div>

        <div class="home-sub" style="margin-top:10px;">
          For 90° bends, this aligns with our shop standard:
          <b>3 × Thickness + Radius × K</b>.<br>
          For critical bends/materials, verify with bend charts or test bends.
        </div>
      </div>

      <!-- OUTPUTS -->
      <div class="calc-section">
        <h3>Outputs</h3>
        <div class="calc-output" id="bend-results">
          Enter values and click <b>Calculate</b>.
        </div>
      </div>

    </div>
  `;

  // ===== DOM refs =====
  const $modeSingle = document.getElementById('bend-mode-single');
  const $modeMulti = document.getElementById('bend-mode-multi');

  const $blockSingle = document.getElementById('bend-block-single');
  const $blockMulti = document.getElementById('bend-block-multi');

  const $t = document.getElementById('bend-thk');
  const $r = document.getElementById('bend-rad');
  const $k = document.getElementById('bend-k');

  const $a = document.getElementById('bend-ang');
  const $l1 = document.getElementById('bend-l1');
  const $l2 = document.getElementById('bend-l2');

  const $multiCount = document.getElementById('bend-multi-count');
  const $multiBuild = document.getElementById('bend-multi-build');
  const $multiInputs = document.getElementById('bend-multi-inputs');

  const $btnCalc = document.getElementById('bend-btn-calc');
  const $btnReset = document.getElementById('bend-btn-reset');
  const $results = document.getElementById('bend-results');

  // ===== state =====
  let bendMode = 'single';

  // ===== helpers =====
  const n = (v) => {
    const num = Number(v);
    return Number.isFinite(num) ? num : NaN;
  };
  const fmt4 = (x) => (Number.isFinite(x) ? x.toFixed(4) : '—');

  function validateGlobals() {
    const T = n($t.value);
    const R = n($r.value);
    const K = n($k.value);

    if (!Number.isFinite(T) || T <= 0) return { ok: false, msg: '⚠️ Please enter a valid <b>Thickness (T)</b>.' };
    if (!Number.isFinite(R) || R < 0) return { ok: false, msg: '⚠️ Please enter a valid <b>Inside Radius (R)</b>.' };
    if (!Number.isFinite(K) || K <= 0 || K >= 1) return { ok: false, msg: '⚠️ Please enter a valid <b>K-Factor</b> (typical: 0.30–0.50).' };

    return { ok: true, T, R, K };
  }

  function calcBendMetrics({ T, R, K, Adeg }) {
    const Arad = (Math.PI / 180) * Adeg;
    const BA = Arad * (R + K * T);
    const SB = (R + T) * Math.tan(Arad / 2);
    const BD = 2 * SB - BA;
    return { BA, SB, BD };
  }

  // ===== mode switching =====
  function setMode(next) {
    bendMode = next;

    if (bendMode === 'single') {
      $modeSingle.classList.add('active');
      $modeMulti.classList.remove('active');
      $blockSingle.style.display = 'block';
      $blockMulti.style.display = 'none';
    } else {
      $modeMulti.classList.add('active');
      $modeSingle.classList.remove('active');
      $blockSingle.style.display = 'none';
      $blockMulti.style.display = 'block';
    }

    $results.innerHTML = `Enter values and click <b>Calculate</b>.`;
  }

  $modeSingle.addEventListener('click', () => setMode('single'));
  $modeMulti.addEventListener('click', () => setMode('multi'));

  // ===== multi input builder =====
  function buildMultiInputs() {
    const count = Math.max(1, Math.min(20, Math.floor(n($multiCount.value) || 1)));
    const segCount = count + 1;

    let html = `
      <div class="home-sub" style="margin-bottom:8px;">
        Enter each bend angle and each outside dimension.
      </div>

      <table class="calc-mini-table">
        <thead>
          <tr>
            <th style="width:50%;">Bend Angles (deg)</th>
            <th style="width:50%;">Outside Dimensions (in)</th>
          </tr>
        </thead>
        <tbody>
    `;

    const rows = Math.max(count, segCount);
    for (let i = 1; i <= rows; i++) {
      html += `<tr>`;

      if (i <= count) {
        html += `
          <td>
            <label style="display:block;color:#bbb;font-weight:800;margin-bottom:6px;">A${i}</label>
            <input class="calc-mini-input" id="bend-a-${i}" type="number" step="0.1" value="90" />
          </td>
        `;
      } else {
        html += `<td></td>`;
      }

      if (i <= segCount) {
        html += `
          <td>
            <label style="display:block;color:#bbb;font-weight:800;margin-bottom:6px;">L${i}</label>
            <input class="calc-mini-input" id="bend-l-${i}" type="number" step="0.0001" placeholder="e.g. 1.2500" />
          </td>
        `;
      } else {
        html += `<td></td>`;
      }

      html += `</tr>`;
    }

    html += `
        </tbody>
      </table>
    `;

    $multiInputs.innerHTML = html;
  }

  $multiBuild.addEventListener('click', buildMultiInputs);
  buildMultiInputs();

  // ===== compute =====
  function computeSingle() {
    const g = validateGlobals();
    if (!g.ok) { $results.innerHTML = g.msg; return; }

    const Adeg = n($a.value);
    if (!Number.isFinite(Adeg) || Adeg <= 0 || Adeg >= 179.9) {
      $results.innerHTML = `⚠️ Please enter a valid <b>Angle (A)</b> between 0 and 180 degrees.`;
      return;
    }

    const { BA, SB, BD } = calcBendMetrics({ ...g, Adeg });

    const L1 = n($l1.value);
    const L2 = n($l2.value);
    const hasFlat = Number.isFinite(L1) && L1 > 0 && Number.isFinite(L2) && L2 > 0;
    const Flat = hasFlat ? (L1 + L2 - BD) : NaN;

    $results.innerHTML = `
      <div>Bend Allowance (BA): <b>${fmt4(BA)}</b> in</div>
      <div>Bend Deduction (BD): <b>${fmt4(BD)}</b> in</div>
      <div>Setback (SB): <b>${fmt4(SB)}</b> in</div>

      <hr style="border:0;border-top:1px solid #2d2d2d;margin:12px 0;" />

      <div>Flat dimension (1 bend): <b>${hasFlat ? fmt4(Flat) : '—'}</b> in</div>
      <div style="color:#bbb;margin-top:6px;">
        ${hasFlat ? `Flat = L1 (${fmt4(L1)}) + L2 (${fmt4(L2)}) − BD (${fmt4(BD)})`
                  : `Enter L1 and L2 to calculate flat dimension.`}
      </div>
    `;
  }

  function computeMulti() {
    const g = validateGlobals();
    if (!g.ok) { $results.innerHTML = g.msg; return; }

    const bendCount = Math.max(1, Math.min(20, Math.floor(n($multiCount.value) || 1)));
    const segCount = bendCount + 1;

    const angles = [];
    for (let i = 1; i <= bendCount; i++) {
      const el = document.getElementById(`bend-a-${i}`);
      const Adeg = n(el?.value);
      if (!Number.isFinite(Adeg) || Adeg <= 0 || Adeg >= 179.9) {
        $results.innerHTML = `⚠️ Invalid angle at <b>A${i}</b>. Use a value between 0 and 180 degrees.`;
        return;
      }
      angles.push(Adeg);
    }

    const segs = [];
    for (let i = 1; i <= segCount; i++) {
      const el = document.getElementById(`bend-l-${i}`);
      const L = n(el?.value);
      if (!Number.isFinite(L) || L <= 0) {
        $results.innerHTML = `⚠️ Please enter a valid outside dimension at <b>L${i}</b>.`;
        return;
      }
      segs.push(L);
    }

    let totalBA = 0;
    let totalBD = 0;

    angles.forEach((Adeg) => {
      const { BA, BD } = calcBendMetrics({ ...g, Adeg });
      totalBA += BA;
      totalBD += BD;
    });

    const sumSegs = segs.reduce((a, b) => a + b, 0);
    const Flat = sumSegs - totalBD;

    $results.innerHTML = `
      <div>Multi Bend Summary</div>
      <div style="margin-top:8px;">Total Bend Allowance (Sum of BA): <b>${fmt4(totalBA)}</b> in</div>
      <div>Total Bend Deduction (Sum of BD): <b>${fmt4(totalBD)}</b> in</div>

      <hr style="border:0;border-top:1px solid #2d2d2d;margin:12px 0;" />

      <div>Sum Outside Dimensions (Total sum of L): <b>${fmt4(sumSegs)}</b> in</div>
      <div>Flat dimension: <b>${fmt4(Flat)}</b> in</div>

      <div style="color:#bbb;margin-top:8px;">
        Flat = (Total sum of L) − (Total sum of BD)
      </div>
    `;
  }

  function compute() {
    if (bendMode === 'single') computeSingle();
    else computeMulti();
  }

  function reset() {
    $t.value = '';
    $r.value = '';
    $k.value = '0.434';

    $a.value = '90';
    $l1.value = '';
    $l2.value = '';

    $multiCount.value = '2';
    buildMultiInputs();

    $results.innerHTML = `Enter values and click <b>Calculate</b>.`;
    setMode('single');
  }

  $btnCalc.addEventListener('click', compute);
  $btnReset.addEventListener('click', reset);

  // Enter key calculates
  [$t, $r, $k, $a, $l1, $l2].forEach(el => {
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') compute();
    });
  });
}

/* =========================
   CALCULATOR: COUNTERSINK (Excel-matching)
   (your original function — unchanged)
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

/**
 * ✅ Shipping schedule loader (implemented)
 * - Reads from the Shipping tab in the Master Schedule spreadsheet
 * - Sorts by Ship Date (oldest → newest) if found
 * - Shows all labeled columns
 */
async function loadShippingOnly() {
  try {
    if (!GID_SHIPPING) throw new Error('Shipping module not configured yet (missing GID_SHIPPING).');

    // ✅ Builds the filter UI once (dropdowns/search/date range)
    ensureShippingFilterBar();

    // ✅ Pull shipping data from the shipping tab
    const table = await fetchGvizTable(SCHEDULE_SHEET_ID, GID_SHIPPING);
    const cols = table.cols || [];
    const rowsAll = (table.rows || []).slice();
    const map = buildColIndexMap(cols);

    // ✅ Build dropdown values from the data
    const idxCustomer = map[normalize(SH_COL_CUSTOMER)];
    const idxType = map[normalize(SH_COL_TYPE)];

    const customers = new Set();
    const types = new Set();

    if (idxCustomer !== undefined) {
      rowsAll.forEach(r => {
        const v = String(cellVal(r.c[idxCustomer]) ?? '').trim();
        if (v) customers.add(v);
      });
    }

    if (idxType !== undefined) {
      rowsAll.forEach(r => {
        const v = String(cellVal(r.c[idxType]) ?? '').trim();
        if (v) types.add(v);
      });
    }

    // ✅ Restore saved dropdown selections (per device)
    const saved = loadShippingFiltersFromStorage() || {};

    if (shippingFilterEls) {
      setShippingDropdownOptions(
        shippingFilterEls.$customer,
        Array.from(customers).sort((a,b) => a.localeCompare(b)),
        saved.customer
      );

      setShippingDropdownOptions(
        shippingFilterEls.$type,
        Array.from(types).sort((a,b) => a.localeCompare(b)),
        saved.type
      );
    }

    // ✅ Cache raw data for filter rendering
    shippingRaw = { cols, rowsAll, map };

    // ✅ Render using filters (includes sorting by DUE DATE)
    renderShippingFiltered();

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

    // ✅ Robust header matching
    const machineIdx = findColIndex(map, [COL_MACHINE, 'MACHINES']);
    const dueIdx = findColIndex(map, [COL_DUE_DATE, 'DUE', 'DUE_DT', 'DUE DATE ', 'DUE DATE:']);
    const statusIdx = findColIndex(map, [COL_STATUS, 'STATUS ', 'JOB STATUS']);
    const estIdx = findColIndex(map, [
      COL_EST_HOURS,
      'EST TIME',
      'EST TIME (HRS)',
      'EST HOURS',
      'EST. TIME (HOURS)',
      'ESTIMATED HOURS',
      'EST HRS'
    ]);

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

    // ✅ Sum estimated hours safely (handles commas/strings)
    let totalHours = 0;
    if (estIdx !== undefined) {
      filtered.forEach(r => {
        const v = cellVal(r.c[estIdx]);
        const raw = String(v ?? '').replace(/,/g, '').trim();
        const num = parseFloat(raw);
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
  if (currentView === 'shipping') loadShippingOnly();     // ✅ NEW: auto refresh shipping
  if (currentView === 'capacity') loadCapacityOnly();
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
