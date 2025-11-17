/******** CONFIG YOU MAY CHANGE ********/
const ROOMS_SHEET_NAME = 'Rooms';   // where your room rows live
const FRIDGE_SPREADSHEET_ID = '1vGZ9Tp7lNqHBIpMgcnk_ZlQr6mdkg7bsQr2l1aYwEbk'; // remote asset tracker
const FRIDGE_SHEET_NAME = 'Fridge'; // tab that holds fridge rows

/******** HEADER HELPERS ********/
function _idxLike_(headers, substrings) {
  const lower = headers.map(h => String(h || '').trim().toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    for (const sub of substrings) if (lower[i].includes(sub.toLowerCase())) return i;
  }
  return -1;
}

function _normalizeStatus_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '';

  const collapsed = raw.replace(/[\s-_]+/g, '');

  if (collapsed.startsWith('avail')) return 'avail';
  if (collapsed.includes('hold')) return 'hold';
  if (/check(?:ed)?in/.test(collapsed)) return 'checkin';
  if (collapsed.includes('soon')) return 'soon';
  if (collapsed.startsWith('reserve') || collapsed.startsWith('reser') || collapsed.startsWith('resrv')) return 'reserved';
  if (collapsed.includes('occup')) return 'reserved';

  return 'reserved';
}

function _emptyStatusTotals_() {
  return { total: 0, avail: 0, hold: 0, reserved: 0, checkin: 0, soon: 0 };
}

function _parkingCapacities_() {
  const fallback = { roofed: 36, open: 22 };
  const cfg = (typeof PARKING_CAPACITY === 'object' && PARKING_CAPACITY) ? PARKING_CAPACITY : fallback;
  const roofed = Number(cfg.roofed || 0);
  const open = Number(cfg.open || 0);
  return { roofed, open, total: roofed + open };
}

function _parkingPlan_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '';

  const cleaned = raw.replace(/[()]/g, '').replace(/\s+/g, ' ').trim();
  const collapsed = cleaned.replace(/\s+/g, '');
  const candidates = [raw, cleaned, collapsed];

  const isNo = ['no','noparking','none','ไม่','ไม่ขอใช้','ไม่มี','ไม่ใช้','ไม่จอด','n/a','na'];
  for (const cand of candidates) {
    if (!cand) continue;
    if (isNo.includes(cand)) return '';
  }

  const roofedTokens = ['roofed','Roofed','covered','car_roof','carroof','roofplan','roofedplan','มีหลังคา','yes'];
  for (const cand of candidates) {
    if (!cand) continue;
    for (const token of roofedTokens) {
      if (cand === token || cand.startsWith(token)) return 'roofed';
    }
  }

  const openTokens = ['open','Open','open-air','outdoor','noroof','no roof','car_noroof','carnoroof','กลางแจ้ง','yes500','yes 500','openplan'];
  for (const cand of candidates) {
    if (!cand) continue;
    for (const token of openTokens) {
      if (cand === token || cand.startsWith(token)) return 'open';
    }
  }

  return '';
}

function _tallyStatuses_(rows) {
  const totals = _emptyStatusTotals_();
  for (const row of rows) {
    const status = row.status || 'reserved';
    totals.total++;
    if (status === 'avail') totals.avail++;
    else if (status === 'hold') totals.hold++;
    else if (status === 'checkin') totals.checkin++;
    else if (status === 'soon') totals.soon++;
    else totals.reserved++;
  }
  return totals;
}

function _readRooms_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ROOMS_SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + ROOMS_SHEET_NAME);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { rows: [], hdr: [] };

  const hdr = values[0].map(String);
  const iRoom    = _idxLike_(hdr, ['room']);
  const iBldg    = _idxLike_(hdr, ['building']);
  const iFloor   = _idxLike_(hdr, ['floor']);
  const iStatus  = _idxLike_(hdr, ['status','สถานะ']);
  const iParking = _idxLike_(hdr, ['parking','ที่จอด']);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const room = String(row[iRoom] ?? '').trim();
    if (!room) continue; // skip blanks

    const statusRaw = String(row[iStatus] ?? '').trim();
    const status = _normalizeStatus_(statusRaw);
    const parkingRaw = String(row[iParking] ?? '').trim();
    const parkingPlan = _parkingPlan_(parkingRaw);
    const parkingYes = !!parkingPlan;
    const building= String(row[iBldg]    ?? '').trim();
    const floor   = Number(row[iFloor] ?? '') || String(row[iFloor] ?? '').trim();

    out.push({
      room,
      building,
      floor,
      status,                         // canonical bucket
      statusRaw,
      parkingYes,
      parkingPlan
    });
  }
  return { rows: out, hdr };
}

function _readFridgeAssets_() {
  if (!FRIDGE_SPREADSHEET_ID) throw new Error('FRIDGE_SPREADSHEET_ID is not configured');
  const ss = SpreadsheetApp.openById(FRIDGE_SPREADSHEET_ID);
  const sh = ss.getSheetByName(FRIDGE_SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + FRIDGE_SHEET_NAME);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const hdr = values[0].map(String);
  const iId = _idxLike_(hdr, ['fridge id','fridgeid','fridge','asset id','assetid']);
  const iStatus = _idxLike_(hdr, ['status']);
  if (iId === -1) throw new Error('Fridge identifier column not found in sheet: ' + FRIDGE_SHEET_NAME);
  if (iStatus === -1) throw new Error('Status column not found in sheet: ' + FRIDGE_SHEET_NAME);

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const fridgeId = String(row[iId] ?? '').trim();
    const status = String(row[iStatus] ?? '').trim();
    if (!fridgeId && !status) continue; // skip empty rows
    rows.push({ fridgeId, status });
  }
  return rows;
}

function _isFridgeInUse_(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return false;

  // Accept only the exact dropdown keywords from Assets_Management sheet.
  const IN_USE_STATUS = ['assigned'];    // treat these as "in use"
  const NOT_IN_USE_STATUS = ['available', 'waiting']; // explicit non-use statuses

  if (NOT_IN_USE_STATUS.includes(normalized)) return false;
  return IN_USE_STATUS.includes(normalized);
}

function _fridgeUsageTotals_() {
  const rows = _readFridgeAssets_();
  return rows.reduce((acc, row) => {
    acc.total++;
    if (_isFridgeInUse_(row.status)) acc.inUse++;
    return acc;
  }, { total: 0, inUse: 0 });
}

function _fridgeSummaryTable_() {
  const totals = _fridgeUsageTotals_();
  const available = Math.max(totals.total - totals.inUse, 0);
  return [
    ['Metric','Value'],
    ['Total fridges', totals.total],
    ['Fridges in use', totals.inUse],
    ['Fridges available', available]
  ];
}

/******** CUSTOM FUNCTIONS (use in cells) ********/
/** =MM_DASHBOARD() -> KPI table */
function MM_DASHBOARD() {
  const { rows } = _readRooms_();

  const statusTotals = _tallyStatuses_(rows);
  const reservedTotal = statusTotals.reserved + statusTotals.checkin + statusTotals.soon;

  return [
    ['Metric','Value'],
    ['Total rooms', statusTotals.total],
    ['Available rooms', statusTotals.avail],
    ['Reserved rooms (incl. check-in & soon)', reservedTotal],
    ['On hold', statusTotals.hold],
    ['Check-in rooms', statusTotals.checkin],
    ['Soon rooms', statusTotals.soon]
  ];
}

/** =MM_ROOMS_BY_BUILDING() */
function MM_ROOMS_BY_BUILDING() {
  const { rows } = _readRooms_();
  const map = new Map();

  for (const r of rows) {
    const k = r.building || '(blank)';
    if (!map.has(k)) map.set(k, { avail:0, hold:0, reserved:0, checkin:0, soon:0, total:0 });
    const obj = map.get(k);
    if (r.status === 'avail') obj.avail++;
    else if (r.status === 'hold') obj.hold++;
    else if (r.status === 'checkin') obj.checkin++;
    else if (r.status === 'soon') obj.soon++;
    else obj.reserved++;
    obj.total++;
  }

  const out = [['Building','avail','hold','reserved','checkin','soon','Total','Reserved%']];
  [...map.entries()]
    .sort(([a],[b]) => String(a).localeCompare(String(b)))
    .forEach(([bldg, o]) => {
      const reservedTotal = o.reserved + o.checkin + o.soon;
      const pct = o.total ? reservedTotal / o.total : 0;
      out.push([bldg, o.avail, o.hold, reservedTotal, o.checkin, o.soon, o.total, pct]);
    });

  return out;
}

/** =MM_ROOMS_BY_FLOOR() */
function MM_ROOMS_BY_FLOOR() {
  const { rows } = _readRooms_();
  const map = new Map(); // key = building|floor

  for (const r of rows) {
    const k = `${r.building || '(blank)'}|${r.floor || '(blank)'}`;
    if (!map.has(k)) map.set(k, { building:r.building||'(blank)', floor:r.floor||'(blank)', avail:0, hold:0, reserved:0, checkin:0, soon:0, total:0 });
    const o = map.get(k);
    if (r.status === 'avail') o.avail++;
    else if (r.status === 'hold') o.hold++;
    else if (r.status === 'checkin') o.checkin++;
    else if (r.status === 'soon') o.soon++;
    else o.reserved++;
    o.total++;
  }

  const out = [['Building','Floor','avail','hold','reserved','checkin','soon','Total','Reserved%']];
  [...map.values()]
    .sort((a,b) => (String(a.building).localeCompare(String(b.building)) || (a.floor>b.floor?1:a.floor<b.floor?-1:0)))
    .forEach(o => {
      const reservedTotal = o.reserved + o.checkin + o.soon;
      const pct = o.total ? reservedTotal / o.total : 0;
      out.push([o.building, o.floor, o.avail, o.hold, reservedTotal, o.checkin, o.soon, o.total, pct]);
    });
  return out;
}

/** =MM_PARKING([capacity]) */
function MM_PARKING(capacity) {
  const caps = _parkingCapacities_();
  const totalCap = Number(capacity) || caps.total;
  const { rows } = _readRooms_();

  const byBldg = new Map();
  for (const r of rows) {
    const b = r.building || '(blank)';
    if (!byBldg.has(b)) byBldg.set(b, { roofed:0, open:0 });
    const obj = byBldg.get(b);
    if (r.parkingPlan === 'roofed') obj.roofed++;
    else if (r.parkingPlan === 'open') obj.open++;
  }

  const totals = [...byBldg.values()].reduce((acc, o) => {
    acc.roofed += o.roofed;
    acc.open += o.open;
    return acc;
  }, { roofed: 0, open: 0 });
  const totalUsed = totals.roofed + totals.open;

  const roofedLeft = Math.max(caps.roofed - totals.roofed, 0);
  const openLeft = Math.max(caps.open - totals.open, 0);
  const totalLeft = totalCap - totalUsed;

  const out = [
    ['Plan','Capacity','Used','Left'],
    ['Roofed', caps.roofed, totals.roofed, roofedLeft],
    ['Open-air', caps.open, totals.open, openLeft],
    ['Total', totalCap, totalUsed, totalLeft],
    ['','','',''],
    ['Building','Roofed Used','Open-Air Used','Total Used']
  ];
  [...byBldg.entries()]
    .sort(([a],[b]) => String(a).localeCompare(String(b)))
    .forEach(([b,o]) => out.push([b, o.roofed, o.open, o.roofed + o.open]));
  out.push(['— TOTAL —', totals.roofed, totals.open, totalUsed]);
  return out;
}

/** =MM_FRIDGE_SUMMARY() */
function MM_FRIDGE_SUMMARY() {
  return _fridgeSummaryTable_();
}

function _writeFridgeSummary_(sheet, anchor = 'E2') {
  if (!sheet) throw new Error('Sheet is required to write fridge summary');
  const table = _fridgeSummaryTable_();
  const start = sheet.getRange(anchor);
  const target = start.offset(0, 0, table.length, table[0].length);
  target.setValues(table);
}

function refreshFridgeSummary_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) throw new Error('Dashboard sheet not found. Run "Build / Refresh sheet" first.');
  _writeFridgeSummary_(sh, 'E2');
}

/******** OPTIONAL: MENU TO BUILD/REFRESH A SHEET WITH CHARTS ********/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard')
    .addItem('Build / Refresh sheet', 'buildDashboardSheet_')
    .addItem('Refresh fridge summary', 'refreshFridgeSummary_')
    .addToUi();
}

function buildDashboardSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Dashboard');
  if (!sh) sh = ss.insertSheet('Dashboard'); else sh.clear();

  // Write formulas so it auto-updates:
  const caps = _parkingCapacities_();
  const totalDefault = caps.total || 0;

  sh.getRange('A1').setValue('KPIs');
  sh.getRange('A2').setFormula('=MM_DASHBOARD()');

  sh.getRange('E1').setValue('Fridges');
  _writeFridgeSummary_(sh, 'E2');

  sh.getRange('A10').setValue('By Building');
  sh.getRange('A11').setFormula('=MM_ROOMS_BY_BUILDING()');

  sh.getRange('H10').setValue('By Floor');
  sh.getRange('H11').setFormula('=MM_ROOMS_BY_FLOOR()');

  sh.getRange('A30').setValue('Parking');
  sh.getRange('A31').setFormula(`=MM_PARKING(${totalDefault})`);

  // Simple chart: Reserved by Floor (from the By Floor table)
  SpreadsheetApp.flush();
  try {
    const range = sh.getRange('H11:L200'); // Building, Floor, avail, hold, reserved_total
    const chart = sh.newChart()
      .asColumnChart()
      .addRange(range)
      .setPosition(30, 8, 0, 0)
      .setOption('title', 'Reserved by Floor (per building)')
      .build();
    sh.insertChart(chart);
  } catch (e) {
    // ignore if table not big enough yet
  }

  sh.autoResizeColumns(1, 20);
}
