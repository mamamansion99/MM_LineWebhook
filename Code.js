/**************** CONFIG ****************/
const BOOKING_DOC_TEMPLATE_ID = '1RYY-YlVhET0YC_LwZgtAzkTOPvGUBESQyx2TpestHH4'; 
const BOOKING_PDF_FOLDER_ID   = '1B8bPFAp0KYxQiO2DdUkF_Vy3ogf2EPyx';
const SHEET_ID   = '1KsimOBXcP2PhZ3Y16DXo7KKcTO9sMNksKJbxc5VEHEQ';
const SHEET_NAME = 'Sheet1';
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
const PAID_MENU_ID = 'richmenu-809f92d6bbba5cc330d0a89f92323a3a';
const PREBOOK_SHEET_NAME   = 'PreBook';
const PREBOOK_CODE_PREFIX  = '#PB';
const SHEET_ROOMS = 'Rooms';
const REVENUE_MASTER_ID   = '1qJU42SUgGgOZY_9X0PedL9yCkLr4d0ni1C4RRRM7g1M'; 
const REVENUE_BILLS_SHEET = 'Horga_Bills';                                    // <-- ชีตปลายทาง
const ASSET_SHEET_ID      = '1vGZ9Tp7lNqHBIpMgcnk_ZlQr6mdkg7bsQr2l1aYwEbk';     // Assets_Management spreadsheet
const ASSET_CAR_SHEET     = 'Car';
const ASSET_SLOT_STATUS_RESERVED  = 'Reserved';
const ASSET_SLOT_STATUS_AVAILABLE = 'Available';
const ROOM_OPENCHAT_LINKS = {
  A: 'https://line.me/ti/g2/dONR8eAdCqgxzVm_5R_rT0OHcVthoguInw74LQ?utm_source=invitation&utm_medium=link_copy&utm_campaign=default'
};
const CHECKIN_PICKER_MAX_DATETIME = '2026-01-14T18:00'; // LINE datetimepicker format (YYYY-MM-DDThhmm)
const CHECKIN_PICKER_TIMEZONE = 'Asia/Bangkok';
const CHECKIN_PICKER_TZ_OFFSET = '+07:00';
const CHECKIN_PICKER_EARLIEST_MINUTES = 10 * 60; // 10:00 น.
const CHECKIN_PICKER_EARLIEST_TIME_LABEL = '10:00';
const CHECKIN_PICKER_LATEST_MINUTES = 18 * 60; // 18:00 น.
const CHECKIN_PICKER_LATEST_TIME_LABEL = '18:00';
const CHECKIN_PICKER_COMMAND_KEYWORDS = [
  'เปลี่ยนวันเช็คอิน',
  'เปลี่ยนวันที่เช็คอิน',
  'เปลี่ยนวันทีเช็คอิน',
  'เปลี่ยนวันเชคอิน',
  'เปลี่ยนเวลาเช็คอิน',
  'เปลี่ยนเวลาเชคอิน',
  'changecheckindate',
  'changecheckintime'
];

const ROOMS_CHECKED_IN_ALIASES = ['checked in', 'check in', 'checked-in', 'check-in'];
const SAFETY_RULE_UPDATE_IMAGE_URL = 'https://drive.google.com/uc?export=view&id=1ctAOSMw22OSyhY4eYoJz08PlWKPH72xT';
const SAFETY_RULE_UPDATE_PREVIEW_URL = SAFETY_RULE_UPDATE_IMAGE_URL;
const SAFETY_RULE_UPDATE_MESSAGE = [
  '📢 แจ้งกฎความปลอดภัยและการใช้กุญแจ/คีย์การ์ด',
  '1) กุญแจเป็นแบบสั่งทำ ซื้อเพิ่มดอกละ 500 บาท คืนวันที่ย้ายออกจะได้รับ 400 บาท',
  '2) ได้คีย์การ์ดฟรี 1 ใบ ต้องคืนเมื่อย้ายออก ซื้อเพิ่มใบละ 100 บาท',
  '⚠️ หากทำหายหรือไม่คืน มีค่าปรับ 1,000 บาท (หักจากเงินประกัน)',
  '🔑 ลืมกุญแจไว้ในห้อง: 08:00-20:00 ช่างเปิดประตู 20 บาท/ครั้ง, หลัง 20:00 บริการเปิด 100 บาท/ครั้ง'
].join('\n');
const SHOE_STORAGE_BROADCAST_MESSAGE = [
  '👟✨ แจ้งเรื่องการเก็บรองเท้าหน่อยนะครับ',
  'ตอนนี้ทางหอมี "ตู้เก็บรองเท้า" หน้าล็อบบี้ให้เรียบร้อยแล้ว',
  'รบกวนลูกหอช่วยกันย้ายรองเท้าที่วางไว้หน้าทางเข้า',
  'มาเก็บในตู้รองเท้าแทนนะครับ',
  'จะได้ทั้งเป็นระเบียบ สะอาด แล้วก็ดูปลอดภัยมากขึ้นด้วย',
  'ขอบคุณที่ช่วยกันดูแลพื้นที่ส่วนกลางของเรานะครับ 🙏🤍'
].join('\n');

var PARKING_CAPACITY = (typeof PARKING_CAPACITY !== 'undefined')
  ? PARKING_CAPACITY
  : { roofed: 36, open: 22 };


const ROOMS_STATUS_OCCUPIED = ['reserved','occupied','จอง','soon','checked in','check in']; // คำที่สื่อว่าอยู่จริง
const ROOMS_MOVEOUT_ALIASES = ['move out date','วันที่ย้ายออก'];

const ROOMS_CURRENT_CODE_ALIASES  = ['current booking code','hg code','mm code','booking code','รหัสการจอง (ปัจจุบัน)'];
const ROOMS_CURRENT_USER_ALIASES = [
  'line id',                        // ✅ ใช้เป็นชื่อหลักใหม่
  'current tenant user id',         // เผื่อใช้ header เก่า
  'ผู้เช่าปัจจุบัน line id'        // รองรับภาษาไทย
];

const ROOMS_MENU_SYNC_AT_ALIASES  = ['menu linked at','menu sync at','เวลาลิงก์เมนู'];
const DAYS_AHEAD_SOON = 90; 
const OCCUPIED_STATUS_KEYWORDS = ROOMS_STATUS_OCCUPIED;
const ROOMS_CODE_ALIASES       = ROOMS_CURRENT_CODE_ALIASES;

function _isFullMonthEra_(checkinDate){
  var cutoff = new Date('2026-01-01T00:00:00+07:00'); // Asia/Bangkok
  var d = new Date(checkinDate.getFullYear(), checkinDate.getMonth(), checkinDate.getDate());
  return d >= cutoff;
}

function normalizeStatus_(s) {
  s = String(s || '').toLowerCase().trim();

  // red
  if (['reserved','จอง','จองแล้ว','checked in','check in','checked-in','checkin','occupied'].includes(s))
    return 'reserved';

  // yellow (optional): treat "soon" as HOLD instead of red
  if (['hold','ถือห้อง','ถือจอง','soon','กำลังจะว่าง'].includes(s))
    return 'hold';

  return 'avail';
}

function readRoomStatus_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) return {};
  const values = sh.getDataRange().getValues();
  const header = values.shift().map(String);
  const iId  = header.findIndex(h => h.toLowerCase().includes('room'));
  const iSt  = header.findIndex(h => h.toLowerCase().includes('status') || h.includes('สถานะ'));
  const out = {};
  values.forEach(r => {
    const id = String(r[iId] || '').trim();
    const st = normalizeStatus_(r[iSt] || '');
    if (id) out[id] = st;
  });
  return out;
}

// --- Parking helpers ---
const PARKING_NO_TOKENS = [
  'no','noparking','none','ไม่','ไม่ขอใช้','ไม่มี','ไม่ใช้','n/a','na','ไม่จอด'
];
const PARKING_ROOFED_TOKENS = [
  'roofed','roof','covered','car_roof','carroof','roof plan','roofed plan','มีหลังคา','roofed(มีหลังคา)','roofed(covered)','yes'
];
const PARKING_OPEN_TOKENS = [
  'open','open air','open-air','outdoor','no roof','noroof','car_noroof','carnoroof','กลางแจ้ง',
  'yes 500','yes500','yes-500','yes(500)','open plan'
];

function _normalizeParkingToken_(txt) {
  const raw = String(txt || '').trim().toLowerCase();
  if (!raw) return '';

  const cleaned = raw.replace(/[()]/g, '').replace(/\s+/g, ' ').trim();
  const collapsed = cleaned.replace(/\s+/g, '');
  const candidates = [raw, cleaned, collapsed];

  for (const cand of candidates) {
    if (!cand) continue;
    if (PARKING_NO_TOKENS.includes(cand)) return '';
  }

  for (const cand of candidates) {
    if (!cand) continue;
    for (const token of PARKING_ROOFED_TOKENS) {
      if (cand === token || cand.startsWith(token + ' ')) return 'roofed';
    }
  }

  for (const cand of candidates) {
    if (!cand) continue;
    for (const token of PARKING_OPEN_TOKENS) {
      if (cand === token || cand.startsWith(token + ' ')) return 'open';
    }
  }

  return '';
}

function _parkingCellValue_(plan, wantsParking) {
  if (!wantsParking) return 'No';
  if (plan === 'roofed') return 'Roofed';
  if (plan === 'open') return 'Open';
  return 'Yes';
}

function _parkingLineLabel_(plan, wantsParking) {
  if (!wantsParking) return 'ไม่ขอใช้';
  if (plan === 'roofed') return 'ขอใช้ (มีหลังคา)';
  if (plan === 'open') return 'ขอใช้ (กลางแจ้ง)';
  return 'ขอใช้';
}

function _parseParkingFromParams_(params) {
  const parkingRaw = String(params.parking || '').trim().toLowerCase();
  const wantsParking = parkingRaw === 'yes';
  const plan = wantsParking
    ? (_normalizeParkingToken_(params.parking_plan) || _normalizeParkingToken_(parkingRaw))
    : '';
  const cell = _parkingCellValue_(plan, wantsParking);
  const label = _parkingLineLabel_(plan, wantsParking);
  return { wantsParking, plan, cell, label };
}

function _parseParkingCellValue_(rawCell) {
  const lower = String(rawCell || '').trim().toLowerCase();
  if (!lower) return { wantsParking: false, plan: '', cell: 'No' };
  if (PARKING_NO_TOKENS.includes(lower)) return { wantsParking: false, plan: '', cell: 'No' };

  const plan = _normalizeParkingToken_(lower);
  if (plan) return { wantsParking: true, plan, cell: _parkingCellValue_(plan, true) };

  if (lower === 'yes') return { wantsParking: true, plan: 'roofed', cell: 'Roofed' };
  if (lower === 'yes 500' || lower === 'yes (500)' || lower === 'yes500') {
    return { wantsParking: true, plan: 'open', cell: 'Open' };
  }

  return { wantsParking: false, plan: '', cell: 'No' };
}

function _parkingCellHasUsage_(rawCell) {
  return _parseParkingCellValue_(rawCell).wantsParking;
}

function _parkingActiveStatuses_() {
  return [
    'pending confirm',
    'awaiting payment',
    'slip received',
    'paid',
    'จ่ายแล้ว'
  ];
}

function _normalizeAssetRoofType_(raw) {
  const token = _normalizeParkingToken_(raw);
  if (token === 'roofed' || token === 'open') return token;
  const txt = String(raw || '').trim().toLowerCase();
  if (['roofed', 'covered'].includes(txt)) return 'roofed';
  if (['open', 'open air', 'open-air', 'outdoor', 'no roof', 'noroof'].includes(txt)) return 'open';
  return '';
}

function _assignAssetParkingSlot_(parkingPlan, info) {
  const plan = _normalizeAssetRoofType_(parkingPlan);
  if (!plan || !ASSET_SHEET_ID || !ASSET_CAR_SHEET) return false;
  info = info || {};

  try {
    const ss = SpreadsheetApp.openById(ASSET_SHEET_ID);
    const sh = ss ? ss.getSheetByName(ASSET_CAR_SHEET) : null;
    if (!sh) return false;

    const headers = _headersRow_(sh);
    const cRoof   = _findCol_(headers, ['roof', 'type']);
    const cStatus = _findCol_(headers, ['status', 'สถานะ']);
    const cRoomId = _findCol_(headers, ['roomid', 'room id', 'room']);
    const cTenant = _findCol_(headers, ['tenantname', 'tenant name', 'tenant', 'name']);
    const cLineId = _findCol_(headers, ['lineuserid', 'line user id', 'line id', 'lineid']);
    const cTel    = _findCol_(headers, ['tel', 'phone', 'โทร', 'เบอร์']);
    if (!cRoof || !cStatus) return false;

    const rowCount = Math.max(sh.getLastRow() - 1, 0);
    if (rowCount === 0) return false;

    const values = sh.getRange(2, 1, rowCount, sh.getLastColumn()).getValues();
    let existingRow = 0;
    let availableRow = 0;

    const desiredRoom = info.roomId ? String(info.roomId || '').trim().toUpperCase() : '';
    const desiredLine = info.userId ? String(info.userId || '').trim() : '';

    for (let i = 0; i < values.length; i++) {
      const rowNum = i + 2;
      const roofType = _normalizeAssetRoofType_(values[i][cRoof - 1]);
      const statusText = String(values[i][cStatus - 1] || '').trim().toLowerCase();

      if (!existingRow) {
        const rowRoom = cRoomId ? String(values[i][cRoomId - 1] || '').trim().toUpperCase() : '';
        const rowLine = cLineId ? String(values[i][cLineId - 1] || '').trim() : '';
        if ((desiredRoom && rowRoom && rowRoom === desiredRoom) ||
            (desiredLine && rowLine && rowLine === desiredLine)) {
          existingRow = rowNum;
        }
      }

      if (!existingRow && !availableRow &&
          roofType === plan &&
          statusText === ASSET_SLOT_STATUS_AVAILABLE.toLowerCase()) {
        availableRow = rowNum;
      }

      if (existingRow && availableRow) break;
    }

    const targetRow = existingRow || availableRow;
    if (!targetRow) {
      Logger.log(`No available ${plan} parking slots for ${info.roomId || 'unknown room'}`);
      try {
        sendLineMessage(`⚠️ ไม่มีที่จอด (${plan}) ว่างสำหรับ ${info.roomId || info.name || 'ลูกค้า'}`);
      } catch (_ignore) {}
      return false;
    }

    const setCell = (col, val) => {
      if (!col) return;
      sh.getRange(targetRow, col).setValue(val || '');
    };

    setCell(cStatus, ASSET_SLOT_STATUS_RESERVED);
    setCell(cRoomId, info.roomId || '');
    setCell(cTenant, info.name || '');
    setCell(cLineId, info.userId || '');
    setCell(cTel, info.phone || '');

    return true;
  } catch (err) {
    Logger.log('_assignAssetParkingSlot_ error: ' + err);
    try { sendLineMessage('❗Parking asset update error: ' + err); } catch (_ignore2) {}
    return false;
  }
}

/** Quick helper to prompt authorization for the asset sheet scopes. */
function testAccessAssetParkingSheet() {
  if (!ASSET_SHEET_ID || !ASSET_CAR_SHEET) {
    throw new Error('ASSET_SHEET_ID / ASSET_CAR_SHEET not configured');
  }
  const ss = SpreadsheetApp.openById(ASSET_SHEET_ID);
  if (!ss) throw new Error('Asset spreadsheet not found');
  const sh = ss.getSheetByName(ASSET_CAR_SHEET);
  if (!sh) throw new Error('Asset sheet not found: ' + ASSET_CAR_SHEET);
  const headers = _headersRow_(sh);
  Logger.log('Asset sheet OK, headers=' + JSON.stringify(headers));
  return 'Asset sheet accessible, columns=' + headers.length;
}

function _parkingCapacities_() {
  const cfg = (typeof PARKING_CAPACITY === 'object' && PARKING_CAPACITY) ? PARKING_CAPACITY : {};
  const roofed = Number(cfg.roofed || 0);
  const open = Number(cfg.open || 0);
  return { roofed, open, total: roofed + open };
}

function _defaultParkingAvailability_() {
  const caps = _parkingCapacities_();
  const nowIso = new Date().toISOString();
  return {
    roofed: { capacity: caps.roofed, used: 0, left: caps.roofed },
    open:   { capacity: caps.open,   used: 0, left: caps.open },
    total:  { capacity: caps.total,  used: 0, left: caps.total },
    ts: nowIso
  };
}

function readParkingAvailability_() {
  const caps = _parkingCapacities_();
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ROOMS);
  if (!sh) return _defaultParkingAvailability_();

  const headers = _headersRow_(sh);
  const cParking = _findCol_(headers, ['parking','ที่จอดรถ']);
  if (!cParking) return _defaultParkingAvailability_();

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return _defaultParkingAvailability_();

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  let roofedUsed = 0;
  let openUsed = 0;

  rows.forEach(row => {
    const info = _parseParkingCellValue_(row[cParking - 1]);
    if (!info.wantsParking) return;

    if (info.plan === 'open') openUsed++;
    else roofedUsed++;
  });

  const totalUsed = roofedUsed + openUsed;
  const nowIso = new Date().toISOString();

  return {
    roofed: {
      capacity: caps.roofed,
      used: roofedUsed,
      left: Math.max(caps.roofed - roofedUsed, 0)
    },
    open: {
      capacity: caps.open,
      used: openUsed,
      left: Math.max(caps.open - openUsed, 0)
    },
    total: {
      capacity: caps.total,
      used: totalUsed,
      left: Math.max(caps.total - totalUsed, 0)
    },
    ts: nowIso
  };
}

/************* RESERVATION FORM HANDLER *************/
function doGet(e) {
  const params = e.parameter || {};
  const action = String(params.action || '').toLowerCase();

  if (action === 'rooms') {
    // Return { rooms: {A101:'avail', A102:'hold', ...}, ts: <iso> }
    const statusMap = readRoomStatus_();
    const out = { rooms: statusMap, ts: new Date().toISOString() };
    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'parking') {
    const availability = readParkingAvailability_();
    return ContentService
      .createTextOutput(JSON.stringify({ parking: availability }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'reserve') {
    // Atomically hold the room and create a reservation row, then push to LINE
    const roomId = String(params.room_id || '').trim();
    if (!roomId) {
      return ContentService.createTextOutput('ERROR: room_id required');
    }

    const parkingInfo = _parseParkingFromParams_(params);
    const wantsParking = parkingInfo.wantsParking;
    const parkingCell = parkingInfo.cell;
    const parkingLabel = parkingInfo.label;

    const lock = LockService.getScriptLock();
    lock.waitLock(30 * 1000);

    try {
      // 1) read room status
      const statusMap = readRoomStatus_();
      const current = statusMap[roomId] || 'avail';
      if (current !== 'avail') {
        return ContentService.createTextOutput('ROOM_TAKEN');
      }

      // 2) flip to HOLD right away
      const ok = setRoomStatus_(roomId, 'hold');
      if (!ok) return ContentService.createTextOutput('ERROR: room not found');

      // 3) create new booking code
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      const lastRow = sheet.getLastRow();
      let newCode;
      if (lastRow <= 1) {
        newCode = "#MM000";
      } else {
        const lastCode = sheet.getRange(lastRow, 1).getValue();
        const lastNum = parseInt(String(lastCode).replace("#MM", ""), 10) || 0;
        const nextNum = lastNum + 1;
        newCode = "#MM" + String(nextNum).padStart(3, '0');
      }

      // 4) build row dynamically with headers
      const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
      let newRow = [
        newCode,                // Code
        new Date(),             // Timestamp
        params.fullname || '',  // Fullname
        params.line_id  || '',  // Line ID
        params.phone    || '',  // Phone
        roomId,                 // Room ID
        params.notes    || '',   // Notes
        parkingCell             // Parking (Roofed/Open/No)
      ];

      // Ensure row length matches header count
      while (newRow.length < headers.length) newRow.push('');

      // ถ้ามีคอลัมน์ Status ให้เติม "Pending Confirm"
      const cStat = headers.findIndex(h => /^(status|สถานะ)$/i.test(h)) + 1;
      if (cStat) newRow[cStat - 1] = 'Pending Confirm';

      // check parking capacity for requested plan
      if (wantsParking) {
        const availability = readParkingAvailability_();
        const planKey = parkingInfo.plan === 'open' ? 'open' : 'roofed';
        const planData = availability[planKey] || { left: 0 };
        const totalData = availability.total || { left: 0 };

        if ((planData.left || 0) <= 0 || (totalData.left || 0) <= 0) {
          return ContentService.createTextOutput('PARKING_FULL');
        }
      }

      // Append the row
      sheet.appendRow(newRow);

      // 5) notify LINE
      const message =
        `📢 มีการจองห้องใหม่!\n` +
        `🆔 รหัส: ${newCode}\n` +
        `👤 ชื่อ: ${params.fullname || '-'}\n` +
        `📞 โทร: ${params.phone || '-'}\n` +
        `🏠 ห้อง: ${roomId}\n` +
        `🚗 ที่จอดรถ: ${parkingLabel}\n` +
        `📝 หมายเหตุ: ${params.notes || '-'}`;
      sendLineMessage(message);

      // 6) return booking code
      return ContentService.createTextOutput(newCode);
    } finally {
      lock.releaseLock();
    }
  }

  if (action === 'prebook') {
    // รับค่าจากฟอร์มจองล่วงหน้า (ไม่มีเลือกห้อง)
    const fullname = String(params.fullname || '').trim();
    const line_id  = String(params.line_id  || '').trim();
    const phone    = String(params.phone    || '').trim();
    const notes    = String(params.notes    || '').trim();
    const parking  = (String(params.parking || '').toLowerCase() === 'yes') ? 'Yes' : 'No';
    const move_in_month = String(params.move_in_month || '').trim(); // รูปแบบ YYYY-MM จาก <input type="month">

    // เขียนลงชีต PreBook
    const sh = _ensurePrebookSheet_();
    const code = _nextPrebookCode_(sh);

    // เตรียมแถวให้เข้ากับเฮดเดอร์
    const row = [
      code, new Date(),
      fullname, line_id, phone,
      parking, move_in_month, notes,
      'New' // Status เริ่มต้น
    ];
    sh.appendRow(row);

    // ส่งแจ้งเตือนเข้า LINE กลุ่มแอดมิน
    const msg =
      '🗓️ มีคำขอ "จองล่วงหน้า"\n' +
      `🆔 รหัส: ${code}\n` +
      `👤 ชื่อ: ${fullname || '-'}\n` +
      `📞 โทร: ${phone || '-'}\n` +
      `💬 Line ID: ${line_id || '-'}\n` +
      `🚗 ที่จอด: ${parking}\n` +
      `📅 เดือนเข้าอยู่: ${move_in_month || '-'}\n` +
      `📝 หมายเหตุ: ${notes || '-'}`;
    sendLineMessage(msg);

    // ตอบกลับเป็นรหัส
    return ContentService.createTextOutput(code);
  }

  // ---- fallback for old links (no action) ----
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const p = e.parameter || {};
  const lastRow = sheet.getLastRow();
  let newCode;
  if (lastRow <= 1) newCode = "#MM000";
  else {
    const lastCode = sheet.getRange(lastRow, 1).getValue();
    const lastNum = parseInt(String(lastCode).replace("#MM", ""), 10) || 0;
    newCode = "#MM" + String(lastNum + 1).padStart(3, '0');
  }
  sheet.appendRow([newCode, new Date(), p.fullname||'', p.line_id||'', p.phone||'', p.room_id||'', p.notes||'']);
  sendLineMessage(`📢 มีการจองห้องใหม่!\n🆔 รหัส: ${newCode}\n👤 ชื่อ: ${p.fullname || '-'}\n📞 โทร: ${p.phone || '-'}\n🏠 ห้อง: ${p.room_id || '-'}\n📝 หมายเหตุ: ${p.notes || '-'}`);
  return ContentService.createTextOutput(newCode);
}

function doPost(e) {
  const secret = PropertiesService.getScriptProperties().getProperty('WORKER_SECRET') || '';
  const bodyText = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
  let body = {};
  try { body = JSON.parse(bodyText || '{}'); }
  catch (_err) { body = {}; }

  const providedSecret = String(body.workerSecret || body.worker_secret || '').trim();
  if (!secret || providedSecret !== secret) {
    return _jsonResponse_({ ok: false, error: 'unauthorized' });
  }

  let handledEvents = false;
  if (Array.isArray(body.events) && body.events.length) {
    try {
      handledEvents = handleForwardedEvents_(body.events) || false;
    } catch (err) {
      Logger.log('handleForwardedEvents_ error: ' + err);
    }
  }

  const act = String(body.act || '').trim();
  if (!act && handledEvents) {
    return _jsonResponse_({ ok: true, handledEvents: true });
  }

  if (act === 'lookup_room_by_line') {
    const lineUserId = String(body.lineUserId || body.line_user_id || '').trim();
    if (!lineUserId) {
      return _jsonResponse_({ ok: false, error: 'missing_line_user' });
    }
    const roomId = _findRoomByUserId_(lineUserId);
    if (!roomId) {
      return _jsonResponse_({ ok: false, error: 'room_not_found' });
    }
    return _jsonResponse_({ ok: true, roomId });
  }

  if (act === 'fridge_rent') {
    const userId =
      String(body.lineUserId || body.line_user_id || '').trim() ||
      String(body.userId || '').trim() ||
      String(body.events && body.events[0] && body.events[0].source && body.events[0].source.userId || '').trim();

    const roomId =
      String(body.roomId || body.room || '').trim() ||
      (userId ? _findRoomByUserId_(userId) : '');

    const comment = String(body.text || body.message || '').trim();
    const via = String(body.via || '').trim() || 'LINE chatbot';

    const lines = [
      '🧊 คำขอเช่าตู้เย็นใหม่',
      `• แหล่งที่มา: ${via}`,
      `• Room: ${roomId || '(ยังไม่ทราบ)'}`,
      userId ? `• LINE UserID: ${userId}` : null,
      comment ? `• ข้อความ: ${comment}` : null
    ].filter(Boolean);

    try { sendLineMessage(lines.join('\n')); }
    catch (err) { Logger.log('sendLineMessage fridge_rent error: ' + err); }

    return _jsonResponse_({ ok: true, roomId: roomId || null });
  }

  return _jsonResponse_({ ok: false, error: 'unknown_act' });
}

function _jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function _parsePostbackData_(raw) {
  const input = String(raw || '').trim();
  if (!input) return {};

  if (input.startsWith('{')) {
    try {
      const parsed = JSON.parse(input);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        return parsed;
      }
    } catch (err) {
      Logger.log('parsePostbackData json error: ' + err);
    }
  }

  const out = {};
  input.split('&').forEach((fragment) => {
    if (!fragment) return;
    const [keyRaw, valRaw = ''] = fragment.split('=');
    const key = decodeURIComponent(String(keyRaw || '')).trim();
    const val = decodeURIComponent(String(valRaw || '')).trim();
    if (!key) return;
    if (Object.prototype.hasOwnProperty.call(out, key)) {
      const prev = out[key];
      out[key] = Array.isArray(prev) ? prev.concat(val) : [prev, val];
    } else {
      out[key] = val;
    }
  });
  return out;
}

function _buildPostbackData_(obj) {
  if (!obj || typeof obj !== 'object') return '';
  return Object.keys(obj)
    .filter(key => Object.prototype.hasOwnProperty.call(obj, key))
    .map(key => {
      const value = obj[key];
      if (value === undefined || value === null || value === '') return null;
      return encodeURIComponent(key) + '=' + encodeURIComponent(String(value));
    })
    .filter(Boolean)
    .join('&');
}

function _lineDatetimeFromParams_(params) {
  if (!params || typeof params !== 'object') return '';
  if (params.datetime) return String(params.datetime);
  if (params.date && params.time) return `${params.date}T${params.time}`;
  return '';
}

function _clockMinutesFromLineDatetime_(raw) {
  const str = String(raw || '').trim();
  if (!str) return NaN;
  const timePart = str.split('T')[1] || str;
  const match = timePart.match(/^(\d{2}):(\d{2})/);
  if (!match) return NaN;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return NaN;
  return hours * 60 + minutes;
}

function _parseLineDatetimeValue_(raw) {
  const str = String(raw || '').trim();
  if (!str) return null;
  const hasSeconds = /:\d{2}(?:[+-]|Z)/.test(str);
  const hasOffset = /[+-]\d{2}:?\d{2}$|Z$/i.test(str);
  let iso = str;
  if (!hasSeconds) iso += ':00';
  if (!hasOffset) iso += CHECKIN_PICKER_TZ_OFFSET;
  const parsed = new Date(iso);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function handleForwardedEvents_(events) {
  if (!Array.isArray(events) || !events.length) return false;
  let handled = false;
  for (let i = 0; i < events.length; i += 1) {
    try {
      if (handleBookingCodeEvent_(events[i]) ||
          handleBookingPostback_(events[i]) ||
          handleCheckinPickerPostback_(events[i]) ||
          handleCheckinPickerTextCommand_(events[i])) {
        handled = true;
      }
    } catch (err) {
      Logger.log('handleForwardedEvents_/item error: ' + err);
    }
  }
  return handled;
}

function handleBookingCodeEvent_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'message') return false;
  const message = event.message || {};
  if (String(message.type || '').toLowerCase() !== 'text') return false;

  const rawText = String(message.text || '').trim();
  if (!rawText) return false;

  const match = rawText.match(/^#?\s*(MM\d{3,})\s*$/i);
  if (!match) return false;

  const bookingCode = '#' + match[1].toUpperCase();
  const targetId =
    (event.source && (event.source.groupId || event.source.roomId || event.source.userId)) || '';

  const info = _lookupReservationByCode_(bookingCode);
  let messages;
  if (info && _shouldShowBookingConfirmCard_(info)) {
    messages = [_buildBookingConfirmTemplate_(info)];
  } else if (info) {
    messages = [{ type: 'text', text: _formatReservationSummary_(info) }];
  } else {
    const notFound = `ไม่พบข้อมูลรหัส ${bookingCode}\nโปรดตรวจสอบว่าพิมพ์ถูกต้อง หรือทักห้องแอดมินเพื่อให้ช่วยตรวจสอบอีกครั้งค่ะ 🙏`;
    messages = [{ type: 'text', text: notFound }];
  }

  if (targetId && messages && messages.length) {
    pushLineMessages_(targetId, messages);
  } else if (!targetId) {
    Logger.log('handleBookingCodeEvent_: missing targetId for ' + bookingCode);
  }

  return true;
}

function handleBookingPostback_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'postback') return false;
  const payload = event.postback || {};
  const data = _parsePostbackData_(payload.data || '');
  const act = String(data.act || '').trim().toLowerCase();
  if (act !== 'booking_confirm') return false;

  const codeRaw = String(data.code || data.bookingCode || '').trim().toUpperCase();
  if (!codeRaw) return false;
  const bookingCode = codeRaw.startsWith('#') ? codeRaw : `#${codeRaw}`;

  const record = _lookupReservationByCode_(bookingCode);
  const targetId =
    (event.source && (event.source.groupId || event.source.roomId || event.source.userId)) || '';

  if (!record) {
    if (targetId) {
      pushLineMessages_(targetId, [{
        type: 'text',
        text: `ไม่พบข้อมูลรหัส ${bookingCode}\nโปรดตรวจสอบอีกครั้ง หรือพิมพ์รหัสใหม่ค่ะ 🙏`
      }]);
    }
    return true;
  }

  _markReservationAwaitingPayment_(record);
  record.status = 'Awaiting Payment';

  if (targetId) {
    pushLineMessages_(targetId, _bookingConfirmationMessages_(record));
  }
  return true;
}

function handleCheckinPickerPostback_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'postback') return false;
  const payload = event.postback || {};
  const data = _parsePostbackData_(payload.data || '');
  if (String(data.act || '').trim().toLowerCase() !== 'checkin_pick') return false;

  const params = payload.params || {};
  const datetimeRaw = _lineDatetimeFromParams_(params);
  const userId = (event.source && event.source.userId) || '';
  const roomId = String(data.room || '').trim() || (userId ? _findRoomByUserId_(userId) : '');
  const pickerMax = CHECKIN_PICKER_MAX_DATETIME ? _parseLineDatetimeValue_(CHECKIN_PICKER_MAX_DATETIME) : null;

  if (!datetimeRaw) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'ระบบไม่พบวัน–เวลาเช็คอินที่เลือก กรุณากดเลือกใหม่อีกครั้งค่ะ 🙏'
      }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  const clockMinutes = _clockMinutesFromLineDatetime_(datetimeRaw);
  const chosenTimeText = (datetimeRaw.split('T')[1] || '').slice(0, 5);
  if (!Number.isFinite(clockMinutes)) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'รูปแบบเวลาไม่ถูกต้อง กรุณาเลือกใหม่อีกครั้งค่ะ 🙏'
      }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  if (clockMinutes < CHECKIN_PICKER_EARLIEST_MINUTES) {
    if (userId) {
      const msg =
        `เวลาที่เลือก (${chosenTimeText || 'ไม่ระบุ'}) ก่อน ${CHECKIN_PICKER_EARLIEST_TIME_LABEL} น.\n` +
        `กรุณาเลือกช่วง ${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้นค่ะ 🙏`;
      pushLineMessages_(userId, [{ type: 'text', text: msg }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  if (clockMinutes > CHECKIN_PICKER_LATEST_MINUTES) {
    if (userId) {
      const msg =
        `เวลาที่เลือก (${chosenTimeText || 'ไม่ระบุ'}) หลัง ${CHECKIN_PICKER_LATEST_TIME_LABEL} น.\n` +
        `กรุณาเลือกช่วง ${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้นค่ะ 🙏`;
      pushLineMessages_(userId, [{ type: 'text', text: msg }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  const selected = _parseLineDatetimeValue_(datetimeRaw);
  if (!selected) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'ไม่สามารถอ่านค่าวันที่/เวลาได้ กรุณาเลือกใหม่ค่ะ 🙏'
      }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  if (pickerMax && selected.getTime() > pickerMax.getTime()) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'วันเช็คอินที่เลือกเกินช่วงที่กำหนดไว้ กรุณาเลือกวันก่อน 15 ม.ค. 2026 ค่ะ 🙏'
      }]);
      if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    }
    return true;
  }

  if (!roomId) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'ระบบไม่พบห้องที่เชื่อมกับบัญชี LINE นี้ กรุณาติดต่อแอดมินเพื่อให้ช่วยบันทึกวันเช็คอินค่ะ 🙏'
      }]);
    }
    return true;
  }

  const dateOnly = new Date(selected.getFullYear(), selected.getMonth(), selected.getDate());
  const timeText = Utilities.formatDate(selected, CHECKIN_PICKER_TIMEZONE, 'HH:mm');
  const saved = _updateRoomCheckinSelection_(roomId, { dateOnly, timeText });

  if (!saved) {
    if (userId) {
      pushLineMessages_(userId, [{
        type: 'text',
        text: 'ระบบบันทึกวันเช็คอินไม่สำเร็จ กรุณาติดต่อแอดมินเพื่อให้ช่วยตรวจสอบค่ะ 🙏'
      }]);
    }
    Logger.log('Check-in picker: failed to write to sheet for room ' + roomId);
    return true;
  }

  const thaiDate = _thaiDate_(dateOnly);
  const ackLines = [
    `บันทึกวันเช็คอินของห้อง ${roomId} แล้วค่ะ 🙏`,
    `🗓️ ${thaiDate} เวลา ${timeText} น.`,
    'หากต้องการปรับเปลี่ยน สามารถกดเลือกใหม่จากปุ่มเดิมได้เลยค่ะ'
  ];
  if (userId) pushLineMessages_(userId, [{ type: 'text', text: ackLines.join('\n') }]);

  Logger.log(`Check-in picker saved for ${roomId}: ${thaiDate} ${timeText}`);
  return true;
}

function handleCheckinPickerTextCommand_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'message') return false;
  const message = event.message || {};
  if (String(message.type || '').toLowerCase() !== 'text') return false;

  const rawText = String(message.text || '').trim();
  if (!rawText) return false;

  const collapsed = rawText
    .toLowerCase()
    .replace(/\s+/g, ''); // ignore spacing differences
  const matched = CHECKIN_PICKER_COMMAND_KEYWORDS.some(keyword => collapsed.includes(keyword));
  if (!matched) return false;

  const source = event.source || {};
  const userId = String(source.userId || '').trim();
  const replyTargetId = String(source.userId || source.groupId || source.roomId || '').trim();

  if (!replyTargetId) return false;

  if (!userId) {
    pushLineMessages_(replyTargetId, [{
      type: 'text',
      text: 'ระบบต้องการข้อมูลบัญชี LINE เพื่อส่งปุ่มเลือกวันเช็คอิน กรุณาเริ่มแชตกับบอทในห้องส่วนตัวก่อนนะคะ 🙏'
    }]);
    return true;
  }

  const roomId = _findRoomByUserId_(userId);
  if (!roomId) {
    pushLineMessages_(replyTargetId, [{
      type: 'text',
      text: 'ระบบไม่พบห้องที่เชื่อมกับบัญชี LINE นี้ กรุณาติดต่อแอดมินเพื่อให้ช่วยตรวจสอบนะคะ 🙏'
    }]);
    return true;
  }

  pushLineMessages_(replyTargetId, [{
    type: 'text',
    text: `ส่งปุ่มเลือกวัน–เวลาเช็คอินของห้อง ${roomId} ให้แล้วค่ะ 🙏\nกดเลือกจากปุ่มล่าสุดได้เลย`
  }]);
  sendCheckinPickerToUser(userId, roomId);
  return true;
}

function _updateRoomCheckinSelection_(roomId, selection) {
  if (!roomId || !selection) return false;
  const { sh, H, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  if (!cRoom) return false;
  const cDate = H.indexOf('CheckinDate') + 1;
  const cTime = H.indexOf('CheckinTime') + 1;
  const cConf = H.indexOf('CheckinConfirmed') + 1;
  if (!cDate && !cTime && !cConf) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const rooms = sh.getRange(2, cRoom, lastRow - 1, 1).getValues();
  const target = String(roomId).trim().toUpperCase();
  for (let i = 0; i < rooms.length; i++) {
    const id = String(rooms[i][0] || '').trim().toUpperCase();
    if (!id || id !== target) continue;
    const row = i + 2;
    if (cDate) sh.getRange(row, cDate).setValue(selection.dateOnly);
    if (cTime) sh.getRange(row, cTime).setValue(selection.timeText);
    if (cConf) sh.getRange(row, cConf).setValue(true);
    return true;
  }
  return false;
}

function _lookupReservationByCode_(code) {
  const normalized = String(code || '').trim().toUpperCase();
  if (!normalized) return null;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return null;

  const headers = _headersRow_(sh);
  const cCode = headers.findIndex(h => /^code$/i.test(h)) + 1 || 1;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const rowCode = String(row[cCode - 1] || '').trim().toUpperCase();
    if (rowCode === normalized) {
      return _extractReservationRecord_(headers, row, i + 2);
    }
  }
  return null;
}

function _extractReservationRecord_(headers, row, rowNumber) {
  const pick = (idx) => (idx ? row[idx - 1] : '');

  const cCode     = headers.findIndex(h => /^code$/i.test(h)) + 1 || 1;
  const cStatus   = _findCol_(headers, ['status', 'สถานะ']);
  const cRoom     = _findCol_(headers, ['room', 'room id', 'ห้อง']);
  const cName     = _findCol_(headers, ['fullname', 'ชื่อ', 'ชื่อ-สกุล']);
  const cPhone    = _findCol_(headers, ['phone', 'โทร']);
  const cParking  = _findCol_(headers, ['parking', 'ที่จอดรถ']);
  const cNotes    = _findCol_(headers, ['notes', 'หมายเหตุ']);
  const cTs       = _findCol_(headers, ['timestamp', 'time', 'วันที่', 'เวลา']) || 2;
  const cExpected = _findCol_(headers, ['expected amount', 'ยอดที่ต้องชำระ', 'amount']);
  const cPaid     = _findCol_(headers, ['paid amount', 'ยอดที่ชำระแล้ว']);
  const cSlipAt   = _findCol_(headers, ['slip received at', 'รับสลิปเมื่อ']);
  const cPaidAt   = _findCol_(headers, ['verified at', 'ยืนยันเมื่อ', 'ตรวจสอบเมื่อ']);

  return {
    row: rowNumber,
    code: String(pick(cCode) || '').trim(),
    status: String(pick(cStatus) || '').trim(),
    room: String(pick(cRoom) || '').trim(),
    name: String(pick(cName) || '').trim(),
    phone: String(pick(cPhone) || '').trim(),
    timestamp: pick(cTs) || null,
    parking: String(pick(cParking) || '').trim(),
    notes: String(pick(cNotes) || '').trim(),
    expectedAmount: _toNumber_(pick(cExpected)),
    paidAmount: _toNumber_(pick(cPaid)),
    slipAt: pick(cSlipAt) || null,
    paidAt: pick(cPaidAt) || null,
  };
}

function _formatReservationSummary_(rec) {
  const parts = [];
  parts.push(`📋 รายละเอียดรหัส ${rec.code || '-'}`);
  if (rec.status) parts.push(`สถานะ: ${rec.status}`);
  if (rec.room) parts.push(`ห้อง: ${rec.room}`);
  if (rec.name) parts.push(`ชื่อผู้จอง: ${rec.name}`);
  if (rec.phone) parts.push(`โทร: ${rec.phone}`);

  const ts = _formatDateTime_(rec.timestamp);
  if (ts) parts.push(`จองเมื่อ: ${ts}`);

  if (rec.parking) parts.push(`ที่จอดรถ: ${rec.parking}`);

  if (rec.expectedAmount) {
    parts.push(`ยอดที่ต้องชำระ: ${_fmtBaht_(rec.expectedAmount)} บ.`);
  }
  if (rec.paidAmount) {
    parts.push(`ยอดที่ชำระแล้ว: ${_fmtBaht_(rec.paidAmount)} บ.`);
  }
  const outstanding = Math.max((rec.expectedAmount || 0) - (rec.paidAmount || 0), 0);
  if (outstanding > 0) {
    parts.push(`ยอดคงเหลือ: ${_fmtBaht_(outstanding)} บ.`);
  }

  const slip = _formatDateTime_(rec.slipAt);
  if (slip) parts.push(`รับสลิปเมื่อ: ${slip}`);
  const verified = _formatDateTime_(rec.paidAt);
  if (verified) parts.push(`ยืนยันยอดเมื่อ: ${verified}`);

  if (rec.notes) parts.push(`หมายเหตุ: ${rec.notes}`);

  parts.push('');
  parts.push('พิมพ์ "ชำระค่าเช่า" เพื่อเริ่มส่งสลิป หรือทักแอดมินได้ตลอดเวลาค่ะ 🙏');
  return parts.filter(Boolean).join('\n');
}

function _shouldShowBookingConfirmCard_(rec) {
  if (!rec) return false;
  const status = String(rec.status || '').trim().toLowerCase();
  return !status || status === 'pending confirm';
}

function _buildBookingConfirmTemplate_(rec) {
  const lines = [
    rec.code ? `รหัส: ${rec.code}` : null,
    rec.room ? `ห้อง: ${rec.room}` : null,
    rec.name ? `ชื่อ: ${rec.name}` : null
  ].filter(Boolean).join('\n') || 'ยืนยันการจองห้องนี้หรือไม่?';

  let dataString = '{}';
  try {
    dataString = JSON.stringify({ act: 'booking_confirm', code: rec.code || '' });
  } catch (err) {
    Logger.log('buildBookingConfirmTemplate stringify error: ' + err);
  }

  return {
    type: 'template',
    altText: `ยืนยันการจอง ${rec.code || ''}`,
    template: {
      type: 'confirm',
      text: lines,
      actions: [
        {
          type: 'postback',
          label: 'ยืนยัน',
          data: dataString,
          displayText: rec.code ? `ยืนยันรหัส ${rec.code}` : 'ยืนยันการจอง'
        },
        {
          type: 'message',
          label: 'ติดต่อแอดมิน',
          text: 'ติดต่อแอดมิน'
        }
      ]
    }
  };
}

function _bookingConfirmationMessages_(rec) {
  const primary = [
    '✅ ยืนยันการจองเรียบร้อย',
    rec.code ? `รหัส: ${rec.code}` : null,
    rec.room ? `ห้อง: ${rec.room}` : null,
    rec.name ? `ชื่อ: ${rec.name}` : null,
    '',
    'กรุณาชำระค่าจอง 2,000 บาทภายใน 24 ชม.',
    'หลังชำระ โปรดส่งสลิปโอน + ภาพบัตรประชาชนของผู้ทำสัญญาเช่าที่แชทนี้เพื่อให้ทีมงานตรวจสอบค่ะ 🙏'
  ].filter(Boolean).join('\n');

  const followUp = 'หากต้องการเปลี่ยนห้องหรือมีคำถามเพิ่มเติม พิมพ์ "ติดต่อแอดมิน" ได้เลยค่ะ';

  return [
    { type: 'text', text: primary },
    { type: 'text', text: followUp }
  ];
}

function _markReservationAwaitingPayment_(rec) {
  if (!rec || !rec.row) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 1000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return;

    const headers = _headersRow_(sh);
    const cStatus = _findCol_(headers, ['status', 'สถานะ']);
    const cConfAt = _findCol_(headers, ['confirmed at', 'ยืนยันเมื่อ']);

    if (cStatus) {
      sh.getRange(rec.row, cStatus).setValue('Awaiting Payment');
    }
    if (cConfAt) {
      sh.getRange(rec.row, cConfAt).setValue(new Date());
    }
  } catch (err) {
    Logger.log('markReservationAwaitingPayment error: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (_e) {}
  }
}

function _formatDateTime_(value) {
  const d = _asDate_(value);
  if (!d) return '';
  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  return Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm');
}

function _asDate_(value) {
  if (value instanceof Date) return value;
  if (value == null || value === '') return null;
  const parsed = Date.parse(value);
  if (isNaN(parsed)) return null;
  return new Date(parsed);
}

function _toNumber_(value) {
  if (typeof value === 'number') return value;
  if (value == null || value === '') return 0;
  const numeric = Number(String(value).replace(/[^\d.-]/g, ''));
  return Number.isFinite(numeric) ? numeric : 0;
}

/************* LINE SENDERS *************/
function testSendLine() {
  const ADMIN_GROUP_ID = "Cdf017804cb8d6f4a8e02c831d700e4b5";
  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: ADMIN_GROUP_ID,
    messages: [{ type: "text", text: "🔔 ทดสอบส่งข้อความจาก Google Apps Script" }]
  };
  const options = {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_TOKEN },
    payload: JSON.stringify(payload)
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    Logger.log("Response: " + res.getContentText());
  } catch (err) {
    Logger.log("Error sending to LINE: " + err);
  }
}

function sendLineMessage(msg) {
  const ADMIN_GROUP_ID = "Cdf017804cb8d6f4a8e02c831d700e4b5";
  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: ADMIN_GROUP_ID,
    messages: [{ type: "text", text: msg }]
  };
  const options = {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_TOKEN },
    payload: JSON.stringify(payload)
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    Logger.log("Response: " + res.getContentText());
  } catch (err) {
    Logger.log("Error sending to LINE: " + err);
  }
}

function pushLineMessages_(targetId, messages) {
  if (!targetId) return false;
  const list = Array.isArray(messages) ? messages.slice() : [];
  if (!list.length) return false;

  const normalized = list
    .map(item => {
      if (!item) return null;
      if (typeof item === 'string') return { type: 'text', text: item };
      if (!item.type) return { type: 'text', text: String(item.text || '') };
      return item;
    })
    .filter(Boolean);

  if (!normalized.length) return false;

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = { to: targetId, messages: normalized };
  const options = {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(url, options);
    return true;
  } catch (err) {
    Logger.log('pushLineMessages_ error: ' + err);
    return false;
  }
}

// Push to a specific LINE userId
function pushLineToUser_(userId, text) {
  if (!userId || !text) return;
  pushLineMessages_(userId, [{ type: 'text', text }]);
}



// ----- Rooms header finder -----
function _roomsHeader_(sh) {
  return sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim().toLowerCase());
}
function _roomsFindCol_(hdr, aliases) {
  for (let i=0;i<hdr.length;i++){
    for (const a of aliases) if (hdr[i].indexOf(String(a).toLowerCase()) !== -1) return i+1;
  }
  return 0;
}

// อ่าน/เขียน "ผู้เช่าปัจจุบัน" ของห้อง
function _getRoomCurrentOccupant_(roomId) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ROOMS);
  if (!sh || !roomId) return {row:0, userId:'', code:''};

  const vals = sh.getDataRange().getValues();
  const hdr  = vals.shift().map(h => String(h||'').trim().toLowerCase());

  const cRoom  = hdr.findIndex(h => h.includes('room')) + 1;
  const cUser  = _roomsFindCol_(hdr, ROOMS_CURRENT_USER_ALIASES);
  const cCode  = _roomsFindCol_(hdr, ROOMS_CURRENT_CODE_ALIASES);

  for (let i=0;i<vals.length;i++){
    const id = String(vals[i][cRoom-1]||'').trim();
    if (id.toUpperCase() === String(roomId).toUpperCase()) {
      return {
        row: i+2,
        userId: cUser ? String(vals[i][cUser-1]||'').trim() : '',
        code:   cCode ? String(vals[i][cCode-1]||'').trim() : ''
      };
    }
  }
  return {row:0, userId:'', code:''};
}

function _setRoomCurrentOccupant_(roomId, opts) {
  // opts = { userId, code, when }
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ROOMS);
  if (!sh || !roomId) return;

  const hdr = _roomsHeader_(sh);
  const cRoom = hdr.findIndex(h => h.includes('room')) + 1;
  const cUser = _roomsFindCol_(hdr, ROOMS_CURRENT_USER_ALIASES);
  const cCode = _roomsFindCol_(hdr, ROOMS_CURRENT_CODE_ALIASES);
  const cAt   = _roomsFindCol_(hdr, ROOMS_MENU_SYNC_AT_ALIASES);

  if (!cUser || !cCode) return;

  const last = sh.getLastRow();
  const vals = sh.getRange(2,1, Math.max(0,last-1), sh.getLastColumn()).getValues();

  for (let i=0;i<vals.length;i++){
    const id = String(vals[i][cRoom-1]||'').trim();
    if (id.toUpperCase() === String(roomId).toUpperCase()) {
      if (opts.code   != null) sh.getRange(i+2, cCode).setValue(opts.code);
      if (opts.userId != null) sh.getRange(i+2, cUser).setValue(opts.userId);
      if (cAt) sh.getRange(i+2, cAt).setValue(opts.when || new Date());
      return;
    }
  }
}

// แกนหลัก: ให้มี “ผู้เช่าคนเดียวต่อห้อง” (unlink คนเก่า, link คนใหม่, อัปเดต Rooms)
function ensureExclusiveMenuForRoom_(roomId, newUserId, newCode) {
  if (!roomId || !newUserId) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 1000);      // กันชนกันถ้ามี event ซ้อน
  try {
    const prev = _getRoomCurrentOccupant_(roomId); // {row, userId, code}

    // 1) ถ้ามีคนเก่าและไม่ใช่คนเดียวกับผู้ชำระล่าสุด → unlink
    if (prev.userId && prev.userId !== newUserId) {
      _unlinkMenu_(prev.userId);
    }

    // 2) ลิงก์เมนูผู้เช่าให้คนใหม่ (ไม่ต้องเช็ค GET ก่อนก็ได้)
    _linkPaidMenu_(newUserId);

    // 3) อัปเดต Rooms: current code/user = ของคนใหม่
    _setRoomCurrentOccupant_(roomId, { userId: newUserId, code: newCode, when: new Date() });

  } finally {
    lock.releaseLock();
  }
}

// ===== Helpers: ตรวจคำว่า "จ่ายแล้ว" =====
function _isPaidText_(txt) {
  const t = String(txt || '').toLowerCase();
  return ['paid','จ่ายแล้ว','verified'].some(k => t.includes(k));
}

// ===== Helpers: ดูว่าห้องนี้เป็น "ผู้เช่าปัจจุบัน" ไหม (อิงชีทรูม) =====
function _isRoomCurrentlyOccupied_(roomId, bookingCodeOpt) {
  if (!roomId) return false;
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ROOMS);
  if (!sh) return false;

  const vals = sh.getDataRange().getValues();
  const hdr  = vals.shift().map(h => String(h || '').trim().toLowerCase());

  const cRoom   = hdr.findIndex(h => h.includes('room'));
  const cStatus = hdr.findIndex(h => h.includes('status') || h.includes('สถานะ'));
  const cMove   = hdr.findIndex(h => ROOMS_MOVEOUT_ALIASES.some(a => h.includes(a)));
  const cCode   = hdr.findIndex(h => ROOMS_CODE_ALIASES.some(a => h.includes(a)));

  for (const r of vals) {
    const rid = String(r[cRoom] || '').trim().toUpperCase();
    if (rid !== String(roomId).trim().toUpperCase()) continue;

    // ถ้ามีโค้ดในชีทรูม ให้เช็คให้ตรง (กันผิดคน)
    if (bookingCodeOpt && cCode > -1) {
      const codeInRoom = String(r[cCode] || '').trim();
      if (codeInRoom && codeInRoom !== bookingCodeOpt) return false;
    }

    const st = String(r[cStatus] || '').toLowerCase();
    const okStatus = OCCUPIED_STATUS_KEYWORDS.some(k => st.includes(k));
    if (!okStatus) return false;

    // ถ้ามีวันย้ายออกและ "เลยวันแล้ว" => ไม่ถือว่า occupied
    if (cMove > -1) {
      const mov = r[cMove];
      if (mov instanceof Date) {
        const endOfDay = new Date(mov.getFullYear(), mov.getMonth(), mov.getDate(), 23,59,59);
        if (endOfDay < new Date()) return false;
      }
    }
    return true;
  }
  return false;
}

// ===== LINE Rich Menu: ลิงก์เมนูจ่ายแล้วให้ user รายคน =====
function _linkPaidMenu_(userId) {
  const url = `https://api.line.me/v2/bot/user/${userId}/richmenu/${PAID_MENU_ID}`;
  const opt = { method:'post', headers:{ Authorization:'Bearer ' + LINE_TOKEN }, muteHttpExceptions:true };
  try { UrlFetchApp.fetch(url, opt); } catch(e){ Logger.log(e); }
}

// ===== LINE Rich Menu: ยกเลิกเมนูของรายคนนั้น (จะกลับ Default เอง) =====
function _unlinkMenu_(userId) {
  const url = `https://api.line.me/v2/bot/user/${userId}/richmenu`;
  const opt = { method:'delete', headers:{ Authorization:'Bearer ' + LINE_TOKEN }, muteHttpExceptions:true };
  try { UrlFetchApp.fetch(url, opt); } catch(e){ Logger.log(e); }
}

// ===== Convenience: unlink menu by room id (clears occupant record too) =====
function unlinkMenuByRoom_(roomId) {
  const room = String(roomId || '').trim();
  if (!room) return 'Room id required';

  const cur = _getRoomCurrentOccupant_(room);
  if (!cur.userId) return 'No linked user for room ' + room;

  _unlinkMenu_(cur.userId);
  _setRoomCurrentOccupant_(room, { userId: '', code: '', when: new Date() });

  return 'Unlinked menu for room ' + room;
}

// ===== Broadcast OpenChat link to tenants who are checked in =====
function broadcastOpenChatToCheckedInTenants() {
  const { sh, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  const cUser = _roomsFindCol_(Hl, ROOMS_CURRENT_USER_ALIASES);
  const cStatus = Hl.findIndex(h => h.includes('status') || h.includes('สถานะ')) + 1;
  if (!cRoom || !cUser || !cStatus) throw new Error('Rooms columns missing (room/current user/status)');

  const rows = Math.max(sh.getLastRow() - 1, 0);
  if (!rows) return 'No tenant rows';

  const vals = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();
  const sentTo = new Set();
  let sent = 0;

  vals.forEach(r => {
    const status = String(r[cStatus - 1] || '').trim().toLowerCase();
    if (!ROOMS_CHECKED_IN_ALIASES.some(t => status === t)) return;

    const roomId = String(r[cRoom - 1] || '').trim();
    const userId = String(r[cUser - 1] || '').trim();
    if (!roomId || !userId || sentTo.has(userId)) return;

    const buildingMatch = roomId.match(/^([A-Z]+)/i);
    const buildingKey = buildingMatch ? buildingMatch[1].toUpperCase() : '';
    const link = ROOM_OPENCHAT_LINKS[buildingKey];
    if (!link) return;

    pushLineToUser_(userId,
      '🎉 ยินดีต้อนรับเข้าสู่ OpenChat สำหรับผู้เช่าอาคาร ' + buildingKey + '\n' +
      'กดเข้าร่วมที่นี่เลย: ' + link);
    sentTo.add(userId);
    sent++;
  });

  const msg = 'broadcastOpenChatToCheckedInTenants: sent=' + sent;
  Logger.log(msg);
  return msg;
}

// ===== Broadcast Safety Rule update (text + image) =====
function broadcastSafetyRuleUpdateToCheckedInTenants_() {
  const { sh, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  const cUser = _roomsFindCol_(Hl, ROOMS_CURRENT_USER_ALIASES);
  const cStatus = Hl.findIndex(h => h.includes('status') || h.includes('สถานะ')) + 1;
  if (!cRoom || !cUser || !cStatus) throw new Error('Rooms columns missing (room/current user/status)');

  const rows = Math.max(sh.getLastRow() - 1, 0);
  if (!rows) return 'No tenant rows';

  const vals = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();
  const sentTo = new Set();
  let sent = 0;

  vals.forEach(r => {
    const status = String(r[cStatus - 1] || '').trim().toLowerCase();
    if (!ROOMS_CHECKED_IN_ALIASES.some(token => status === token)) return;

    const userId = String(r[cUser - 1] || '').trim();
    if (!userId || sentTo.has(userId)) return;

    const messages = [{ type: 'text', text: SAFETY_RULE_UPDATE_MESSAGE }];
    const imageUrl = SAFETY_RULE_UPDATE_IMAGE_URL || '';
    const previewUrl = SAFETY_RULE_UPDATE_PREVIEW_URL || imageUrl;
    if (imageUrl && previewUrl) {
      messages.push({
        type: 'image',
        originalContentUrl: imageUrl,
        previewImageUrl: previewUrl
      });
    }

    pushLineMessages_(userId, messages);
    sentTo.add(userId);
    sent++;
  });

  const msg = 'broadcastSafetyRuleUpdateToCheckedInTenants_: sent=' + sent;
  Logger.log(msg);
  return msg;
}

// ===== Broadcast reminder to place shoes inside the cabinet =====
function broadcastShoeCabinetReminderToCheckedInTenants() {
  const { sh, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  const cUser = _roomsFindCol_(Hl, ROOMS_CURRENT_USER_ALIASES);
  const cStatus = Hl.findIndex(h => h.includes('status') || h.includes('สถานะ')) + 1;
  if (!cRoom || !cUser || !cStatus) throw new Error('Rooms columns missing (room/current user/status)');

  const rows = Math.max(sh.getLastRow() - 1, 0);
  if (!rows) return 'No tenant rows';

  const vals = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();
  const sentTo = new Set();
  let sent = 0;

  vals.forEach(r => {
    const status = String(r[cStatus - 1] || '').trim().toLowerCase();
    if (!ROOMS_CHECKED_IN_ALIASES.some(token => status === token)) return;

    const userId = String(r[cUser - 1] || '').trim();
    if (!userId || sentTo.has(userId)) return;

    pushLineToUser_(userId, SHOE_STORAGE_BROADCAST_MESSAGE);
    sentTo.add(userId);
    sent++;
  });

  const msg = 'broadcastShoeCabinetReminderToCheckedInTenants: sent=' + sent;
  Logger.log(msg);
  return msg;
}

function listRichMenus() {
  const res = fetchL('https://api.line.me/v2/bot/richmenu/list', { method: 'get' });
  const code = res.getResponseCode();
  const body = res.getContentText();
  Logger.log('LIST -> ' + code + ' ' + body);
  if (code < 200 || code >= 300) {
    throw new Error('List rich menus failed: ' + code + ' ' + body);
  }
  return body;
}

// ===== ตัดสินใจจาก "แถวใน Sheet1" แล้วลิงก์/ยกเลิก เมนู =====
function ensureMenuByRow_(row, statusTextNow) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const headers = _headersRow_(sh);

  const cUserId = _findUserIdCol_(headers);
  const cRoomId = _findCol_(headers, ['room','room id','ห้อง']);
  const cCode   = headers.findIndex(h => /^code$/i.test(h)) + 1 || 1;

  const userId = cUserId ? String(sh.getRange(row, cUserId).getValue() || '').trim() : '';
  if (!userId) return;

  const roomId = cRoomId ? String(sh.getRange(row, cRoomId).getValue() || '').trim() : '';
  const code   = String(sh.getRange(row, cCode).getValue() || '').trim();

  const paid = _isPaidText_(statusTextNow);
  const occupied = _isRoomCurrentlyOccupied_(roomId, code);

  if (paid && occupied) _linkPaidMenu_(userId);
  else _unlinkMenu_(userId);
}

function oneTimeRelinkAllCurrentOccupants() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_ROOMS);
  if (!sh) throw new Error('Rooms sheet not found');

  const vals = sh.getDataRange().getValues();
  const hdr  = vals.shift().map(h => String(h||'').trim().toLowerCase());

  const cRoom = hdr.findIndex(h => h.includes('room')) + 1;
  const cUser = _roomsFindCol_(hdr, ROOMS_CURRENT_USER_ALIASES);
  const cStat = hdr.findIndex(h => h.includes('status') || h.includes('สถานะ')) + 1;
  const cMove = hdr.findIndex(h => ROOMS_MOVEOUT_ALIASES.some(a => h.includes(a))) + 1;

  if (!cRoom || !cUser || !cStat) throw new Error('Required columns missing (room/current user/status)');

  let linked = 0, skipped = 0;
  const now = new Date();

  vals.forEach((r, i) => {
    const rowIdx = i + 2;
    const roomId = String(r[cRoom-1]||'').trim();
    const userId = String(r[cUser-1]||'').trim();
    const st     = String(r[cStat-1]||'').toLowerCase();

    if (!roomId || !userId) { skipped++; return; }

    // occupied rule (same as your nightly)
    let occupied = OCCUPIED_STATUS_KEYWORDS.some(a => st.includes(a));
    if (occupied && cMove) {
      const mov = r[cMove-1];
      if (mov instanceof Date) {
        const end = new Date(mov.getFullYear(), mov.getMonth(), mov.getDate(), 23,59,59);
        if (end < new Date()) occupied = false;
      }
    }
    if (!occupied) { skipped++; return; }

    _linkPaidMenu_(userId);
    // write sync time if you have the column
    const cAt = _roomsFindCol_(hdr, ROOMS_MENU_SYNC_AT_ALIASES);
    if (cAt) sh.getRange(rowIdx, cAt).setValue(now);
    linked++;
  });

  const msg = `oneTimeRelinkAllCurrentOccupants(): linked=${linked}, skipped=${skipped}`;
  Logger.log(msg);
  return msg;
}


/************* HEADER ROOMS SHEET FINDER *************/
function _headersRow_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
}
function _findCol_(headers, aliases) {
  const lower = headers.map(h => h.toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    for (const a of aliases) {
      if (lower[i].indexOf(a.toLowerCase()) !== -1) return i + 1;
    }
  }
  return null;
}

function _findUserIdCol_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].trim().toLowerCase();
    if (h === 'line user id') return i + 1;
  }
  // Add broader aliases, including "line id"
  return _findCol_(headers, [
    'line user id', 'line id', 'ไลน์ไอดี', 'ผู้ใช้ไลน์', 'line user'
  ]);
}


function setRoomStatus_(roomId, newStatus) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) throw new Error('Rooms sheet not found');
  const values = sh.getDataRange().getValues();
  const header = values.shift().map(String);
  const iId  = header.findIndex(h => h.toLowerCase().includes('room'));
  const iSt  = header.findIndex(h => h.toLowerCase().includes('status') || h.includes('สถานะ'));
  if (iId < 0 || iSt < 0) throw new Error('Rooms columns not found');
  for (let r = 0; r < values.length; r++) {
    const id = String(values[r][iId] || '').trim();
    if (id === roomId) {
      sh.getRange(r + 2, iSt + 1).setValue(newStatus);
      return true;
    }
  }
  return false;
}

/**
 * Auto-release unconfirmed holds after N hours (default 2h).
 * Looks for Status empty or "Pending Confirm", no Confirmed At, Timestamp older than limit.
 * Actions: Rooms hold->avail, Sheet1.Status="Expired", Expired At=now (if exists).
 */
function releaseUnconfirmedHolds(hours) {
  hours = Number(hours) || 2;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  const cTs      = _findCol_(headers, ['timestamp','time','วันที่','เวลา']) || 2;
  const cRoom    = _findCol_(headers, ['room','room id','ห้อง']) || 6;
  const cStatus  = _findCol_(headers, ['status','สถานะ']);
  const cConfAt  = _findCol_(headers, ['confirmed at','ยืนยันเมื่อ']);
  const cExpired = _findCol_(headers, ['expired at','หมดอายุเมื่อ']);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'No rows';

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const now    = new Date();
  const limit  = hours * 3600 * 1000;

  let scanned = 0, released = 0;

  for (let i = 0; i < values.length; i++) {
    scanned++;
    const row    = values[i];
    const roomId = String(row[cRoom-1] || '').trim();
    if (!roomId) continue;

    const stLower = cStatus ? String(row[cStatus-1] || '').toLowerCase() : '';

    // skip rows that are already confirmed/paid/cancelled/expired/slip
    if (['awaiting payment','paid','จ่ายแล้ว','slip received','ใบเสร็จได้รับ','cancelled','ยกเลิก','expired'].includes(stLower)) {
      continue;
    }

    // if confirmed at exists, this 2h job doesn't apply
    if (cConfAt && (row[cConfAt-1] instanceof Date)) continue;

    const ts = row[cTs-1];
    if (!(ts instanceof Date)) continue;

    if (now - ts > limit) {
      try {
        setRoomStatus_(roomId, 'avail');
        if (cStatus)  sh.getRange(i+2, cStatus).setValue('Expired');
        if (cExpired) sh.getRange(i+2, cExpired).setValue(new Date());
        released++;
      } catch (e) {
        Logger.log('releaseUnconfirmedHolds: failed to release ' + roomId + ': ' + e);
      }
    }
  }

  const msg = `releaseUnconfirmedHolds(): scanned=${scanned}, released=${released}, threshold=${hours}h`;
  Logger.log(msg);
  return msg;
}


/**
 * Release rooms that have been on HOLD longer than maxHours (default 24).
 */
function releaseExpiredHolds(maxHours) {
  maxHours = Number(maxHours) || 24;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  const cTs    = _findCol_(header, ['timestamp','time','วันที่','เวลา']) || 2;
  const cRoom  = _findCol_(header, ['room','room id','ห้อง']) || 6;
  const cStat  = _findCol_(header, ['status','สถานะ']);
  const cExpAt = _findCol_(header, ['expired at','หมดอายุเมื่อ']);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'No rows';

  const rng = sh.getRange(2, 1, lastRow-1, sh.getLastColumn());
  const values = rng.getValues();
  const roomMap = readRoomStatus_();

  let released = 0, examined = 0;
  const now = new Date();
  const msLimit = maxHours * 60 * 60 * 1000;

  for (let i = 0; i < values.length; i++) {
    examined++;
    const row = values[i];
    const ts = row[cTs-1];
    const roomId = String(row[cRoom-1] || '').trim();
    const curStatus = cStat ? String(row[cStat-1] || '').trim() : '';

    if (!roomId || !(ts instanceof Date)) continue;

    const statusLower = (curStatus || '').toLowerCase();
    if (['paid','จ่ายแล้ว','Slip Received','expired','cancelled','ยกเลิก'].includes(statusLower)) {
      continue;
    }

    const ageMs = now - ts;
    if (ageMs < msLimit) continue;

    if (roomMap[roomId] === 'hold') {
      try {
        setRoomStatus_(roomId, 'avail');
        released++;
        if (cStat) sh.getRange(i+2, cStat).setValue('Expired');
        if (cExpAt) sh.getRange(i+2, cExpAt).setValue(new Date());
      } catch (e) {
        Logger.log('Failed to release ' + roomId + ': ' + e);
      }
    }
  }
  const msg = `releaseExpiredHolds(): examined=${examined}, released=${released}, thresholdHours=${maxHours}`;
  Logger.log(msg);
  return msg;
}

/**
 * Single sweep: run both phases safely in one execution.
 * - 2h unconfirmed holds (Pending Confirm)
 * - 24h awaiting-payment holds
 */
function releaseHourlySweep() {
  const lock = LockService.getScriptLock();
  let gotLock = false;
  try {
    gotLock = lock.tryLock(10 * 1000); // up to 10s
    if (!gotLock) return 'releaseHourlySweep: skipped (lock held)';

    const a = releaseUnconfirmedHolds(2); // Phase A
    const b = releaseExpiredHolds(24);    // Phase B

    const msg = `releaseHourlySweep: ${a} | ${b}`;
    Logger.log(msg);
    return msg;
  } catch (err) {
    Logger.log('releaseHourlySweep ERROR: ' + err);
    return 'releaseHourlySweep ERROR: ' + err;
  } finally {
    if (gotLock) lock.releaseLock();
  }
}


function handleStatusEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();

    // 1) Only this spreadsheet
    if (sh.getParent().getId() !== SHEET_ID) return;

    // 2) Only Sheet1 (your bookings sheet)
    if (sh.getName() !== SHEET_NAME) return;

    const row = e.range.getRow();
    if (row === 1) return; // skip header

    // 3) Only when editing the "Status" column
    const headers = _headersRow_(sh);
    const cStatus = _findCol_(headers, ["status", "สถานะ"]);
    if (!cStatus) return;
    if (e.range.getColumn() !== cStatus) return;

    // 4) Ignore no-op changes
    const newVal = String(e.value || "").trim().toLowerCase();
    const oldVal = String(e.oldValue || "").trim().toLowerCase();
    if (newVal === oldVal) return;

    // Column helpers (used below)
    const cUserId   = _findUserIdCol_(headers);
    const cSlipAt   = _findCol_(headers, ["slip received at", "รับสลิปเมื่อ"]);
    const cPaidAt   = _findCol_(headers, ["verified at", "ยืนยันเมื่อ", "ตรวจสอบเมื่อ"]);
    const cRoomId   = _findCol_(headers, ["room","room id","ห้อง"]);
    const cParking  = _findCol_(headers, ["parking","ที่จอดรถ"]); // [ADD]

    const cCode     = headers.findIndex(h => /^code$/i.test(h)) + 1 || 1;
    const cName     = _findCol_(headers, ["fullname","ชื่อ","ชื่อ-สกุล"]) || 3;
    const cPhone    = _findCol_(headers, ["phone","โทร"]) || 5;
    const cExpected = _findCol_(headers, ["expected amount","ยอดที่ต้องชำระ","amount"]);
    const cPaidAmt  = _findCol_(headers, ["paid amount","ยอดที่ชำระแล้ว"]);
    const cRef      = _findCol_(headers, ["ref","reference","เลขอ้างอิง"]);
    const cBookingUrl = _findCol_(headers, [
      "booking pdf url", "confirmation pdf url", "booking note url", "ลิงก์เอกสารยืนยัน"
    ]);

    // Pull common row values once
    const userId = cUserId ? String(sh.getRange(row, cUserId).getValue() || "").trim() : "";
    const roomId = cRoomId ? String(sh.getRange(row, cRoomId).getValue() || "").trim() : "";
    const parkingRaw = cParking ? String(sh.getRange(row, cParking).getValue() || "").trim() : "";
    const parkingInfoCell = _parseParkingCellValue_(parkingRaw);
    if (cParking) {
      const canonicalCell = parkingInfoCell.wantsParking ? parkingInfoCell.cell : 'No';
      if (canonicalCell && canonicalCell !== parkingRaw) {
        sh.getRange(row, cParking).setValue(canonicalCell);
      }
    }
    const code   = String(sh.getRange(row, cCode).getValue() || "").trim();
    const name   = String(sh.getRange(row, cName).getValue() || "").trim();
    const phone  = String(sh.getRange(row, cPhone).getValue() || "").trim();
    const expAmt = cExpected ? Number(sh.getRange(row, cExpected).getValue() || 0) : 0;
    const paidAmt= cPaidAmt  ? Number(sh.getRange(row, cPaidAmt ).getValue() || expAmt) : expAmt;
    const bankRef= cRef ? String(sh.getRange(row, cRef).getValue() || "").trim() : "";

    // =========================
    // A) Slip Received
    // =========================
    if (newVal === "slip received" || newVal === "รับสลิปแล้ว") {
      if (cSlipAt) sh.getRange(row, cSlipAt).setValue(new Date());
      if (userId)  pushLineToUser_(userId, "ได้รับสลิปแล้วค่ะ ทีมงานจะตรวจสอบยอดและยืนยันโดยเร็ว ขอบคุณค่ะ 🙏");
      return;
    }

    // =========================
    // B) Paid
    // =========================
    if (newVal === "paid" || newVal === "จ่ายแล้ว") {
      // timestamp verified if empty
      if (cPaidAt && !(sh.getRange(row, cPaidAt).getValue() instanceof Date)) {
        sh.getRange(row, cPaidAt).setValue(new Date());
      }

      // mark room reserved in Rooms
      if (roomId) {
        try { 
          setRoomStatus_(roomId, "reserved"); 
          const parkingForRooms = parkingInfoCell.wantsParking ? parkingInfoCell.cell : '';
          _setRoomParking_(roomId, parkingForRooms);
        } catch (_e) {}
      }

      if (parkingInfoCell.wantsParking && parkingInfoCell.plan) {
        _assignAssetParkingSlot_(parkingInfoCell.plan, {
          roomId: roomId,
          name: name,
          userId: userId,
          phone: phone
        });
      }

            // [ADD] --- Remaining amount = (room price + 5,000) - paidAmt ---
      // Try to read room price from Rooms sheet (column 'ราคาห้อง' / 'price' aliases).
      let roomPrice = 0;
      try {
        const ssR = SpreadsheetApp.openById(SHEET_ID);
        const shRooms = ssR.getSheetByName('Rooms');
        if (shRooms && roomId) {
          const valsR = shRooms.getDataRange().getValues();
          const hdrR  = valsR.shift().map(h => String(h || '').trim().toLowerCase());

          const iId = hdrR.findIndex(h => h.includes('room'));
          const aliases = ['ราคาห้อง','price','room price','expected amount','amount','monthly'];
          let iPrice = -1;
          for (const a of aliases) {
            const j = hdrR.findIndex(h => h.indexOf(a.toLowerCase()) !== -1);
            if (j > -1) { iPrice = j; break; }
          }

          if (iId > -1 && iPrice > -1) {
            for (const r of valsR) {
              if (String(r[iId] || '').trim().toUpperCase() === String(roomId).trim().toUpperCase()) {
                roomPrice = Number(r[iPrice] || 0) || 0;
                break;
              }
            }
          }
        }
      } catch (_ignore) {}

      // Fallback to Expected Amount if room price not found
      if (!roomPrice && expAmt) roomPrice = Number(expAmt) || 0;

      // Formula: (roomPrice + 5,000) - paidAmt   (never below 0)
      const remainingAmt = Math.max((roomPrice + 5000) - (paidAmt || 0), 0);

      // create / fill booking PDF URL if missing
      let pdfUrl = cBookingUrl ? String(sh.getRange(row, cBookingUrl).getValue() || "") : "";
      if (!pdfUrl) {
        try {
          pdfUrl = createBookingPdf_({
            code: code,
            date: new Date(),
            customer: name,
            phone: phone,
            line_id: userId,
            room_id: roomId,
            paid_amount: paidAmt || expAmt,
            bank_ref: bankRef,
            move_in_month: "",
            move_in_window: "",
            contact_phone: phone,
            contact_line: "@MamaMansion",
            address: "ฉลองกรุง 37 (ใกล้นิคมลาดกระบัง)",
            remaining_amount: remainingAmt, // [CHANGE] use computed remaining
            notes: ""
          });
          if (cBookingUrl) sh.getRange(row, cBookingUrl).setValue(pdfUrl);
        } catch (err) {
          Logger.log("PDF error: " + err);
          sendLineMessage("❗PDF error: " + err);
        }
      }

      // queue “Horganice tasks” on Rooms
      if (roomId) {
        _markRoomNeedsHorganice_(roomId, {
          code: code, name: name, phone: phone, rowIndex: row
        });
      }

      // notify tenant
      if (userId) {
        pushLineToUser_(
          userId,
          "ยืนยันการชำระเงินเรียบร้อย ✅\n" +
          `รหัสการจอง: ${code}\n` +
          `ห้อง: ${roomId}\n` +
          (pdfUrl ? `เอกสารยืนยันการจอง & คู่มือเข้าอยู่: ${pdfUrl}\n` : "") +
          "หมายเหตุ: เอกสารนี้ไม่ใช่ใบเสร็จหรือเอกสารภาษี ใช้เพื่อแจ้งขั้นตอนถัดไปเท่านั้น"
        );

      sendCheckinPickerToUser(userId, roomId);
      }

      // notify admin group
      const rowLink = _rowLink_(row);
      sendLineMessage(
        "✅ Confirmed (manual mark = Paid)\n" +
        `รหัส: ${code} | ห้อง: ${roomId}\n` +
        `ผู้จอง: ${name}\n` +
        (pdfUrl ? `เอกสารยืนยัน: ${pdfUrl}\n` : "") +
        `แถวในชีท: ${rowLink}`
      );

      // ให้มีผู้เช่าคนเดียวต่อห้อง: unlink คนเก่า → link คนใหม่ → อัปเดต Rooms
      ensureExclusiveMenuForRoom_(roomId, userId, code);

      return;
    }

    // =========================
    // C) Cancelled / Expired
    // =========================
    if (newVal === "cancelled" || newVal === "expired" || newVal === "ยกเลิก") {
      if (roomId) {
        try { 
          setRoomStatus_(roomId, "avail"); 
          _setRoomParking_(roomId, ""); // [ADD] clear parking on cancel/expire
        } catch (_e) {}
      }

      // ถ้ามีผู้เช่าปัจจุบันใน Rooms ให้ unlink แล้วเคลียร์คนปัจจุบันออก
      (function(){
        const cur = _getRoomCurrentOccupant_(roomId);
        if (cur.userId) _unlinkMenu_(cur.userId);
        _setRoomCurrentOccupant_(roomId, { userId: '', code: '', when: new Date() });
      })();

      return;
    }

    // Otherwise ignore (we only care about specific Status values)
    return;

  } catch (err) {
    Logger.log("handleStatusEdit error: " + err);
    try { sendLineMessage("❗handleStatusEdit error: " + err); } catch (_e) {}
  }
}



/** Helper: find RoomId from Rooms by Line ID (case-insensitive). */
function _findRoomByUserId_(userId) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) return '';
  const H  = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
  const Hl = H.map(h => h.toLowerCase());
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  const cUser = Hl.findIndex(h => h.includes('line id')) + 1;
  if (!cRoom || !cUser) return '';

  const vals = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const id = String(vals[i][cUser-1]||'').trim();
    if (id && id.toLowerCase() === String(userId||'').trim().toLowerCase()) {
      return String(vals[i][cRoom-1]||'').trim();
    }
  }
  return '';
}

/** Internal: load Rooms headers quickly */
function _roomsHeaders_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  const H  = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  return { sh, H, Hl: H.map(h=>h.toLowerCase()) };
}

/** Push picker to ONE room (for spot tests) */
function sendCheckinPickerToRoom(roomId) {
  const { sh, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h=>h.includes('room')) + 1;
  const cUser = Hl.findIndex(h=>h.includes('line id')) + 1;
  if (!cRoom || !cUser) throw new Error('Rooms: missing Room / Line ID columns');

  const vals = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const id = String(vals[i][cRoom-1]||'').trim();
    if (id.toUpperCase() === String(roomId||'').toUpperCase()) {
      const userId = String(vals[i][cUser-1]||'').trim();
      if (!userId) throw new Error('This room has no LINE ID');
      sendCheckinPickerToUser(userId, id);
      Logger.log('Sent picker to '+id+' -> '+userId);
      return;
    }
  }
  throw new Error('Room not found: '+roomId);
}

/**
 * Push picker to ALL eligible tenants:
 * - has LINE ID
 * - status looks like reserved/occupied/soon
 * - no CheckinDate yet (and not already confirmed)
 */
function sendCheckinPickerToEligibleTenants(limitOpt) {
  const limit = Number(limitOpt)||9999;
  const { sh, H, Hl } = _roomsHeaders_();
  const cRoom  = Hl.findIndex(h => String(h).toLowerCase().includes('room')) + 1;
  const cUser  = Hl.findIndex(h => String(h).toLowerCase().includes('line id')) + 1;
  const cStat  = Hl.findIndex(h => {
    const s = String(h).toLowerCase();
    return s.includes('status') || s.includes('สถานะ');
  }) + 1;
  const cDate  = H.indexOf('CheckinDate') + 1;
  const cTime  = H.indexOf('CheckinTime') + 1;
  const cConf  = H.indexOf('CheckinConfirmed') + 1;

  if (!cRoom || !cUser || !cStat) throw new Error('Rooms: missing Room/LineID/Status columns');

  const rows = sh.getLastRow() - 1;
  if (rows < 1) return 'No data rows';

  const vals = sh.getRange(2,1,rows, sh.getLastColumn()).getValues();

  // --- Added: LINE token fetch once
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) throw new Error('Missing LINE_TOKEN in Script Properties');

  let sent = 0, skipped = 0;
  for (let i=0; i<vals.length; i++){
    if (sent >= limit) break;

    const roomId = String(vals[i][cRoom-1]||'').trim();
    const userId = String(vals[i][cUser-1]||'').trim();
    const status = String(vals[i][cStat-1]||'').trim().toLowerCase();

    if (!userId || !roomId) { skipped++; continue; }

    const okStatus = (OCCUPIED_STATUS_KEYWORDS || []).some(k => st.includes(String(k).toLowerCase()));
    if (!okStatus) { skipped++; continue; }

    const hasDate = cDate && (vals[i][cDate-1] instanceof Date || String(vals[i][cDate-1]||'').trim());
    const hasTime = cTime && String(vals[i][cTime-1]||'').trim();
    const confirmed = cConf && !!vals[i][cConf-1];
    if (hasDate || hasTime || confirmed) { skipped++; continue; }

    try {
      // ----- Added: pre-prompt message before showing the date picker -----
      const promptMsg =
        'เรียนลูกค้า 🙏\n' +
        'กรุณา **วันที่และเวลาเช็คอิน 📅** เพื่อให้เราคำนวณค่าเช่า ' +
        '**ตามจำนวนวันที่เข้าพักจริงสำหรับเดือนแรก ⚖️** อย่างยุติธรรมครับ/ค่ะ\n' +
        'ตัวอย่าง: 4,000/เดือน ≈ 133/วัน 🧮 เช็คอิน 10 → ~2,940 บ. (ปัดขึ้นหลักสิบ)\n' +
        `⏰ เลือกเวลาได้เฉพาะ ${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้น\n` +
        '**กรุณากดยืนยันภายใน 31 ต.ค. ⏰** หากไม่ยืนยันตามกำหนด อาจคิดค่าเช่าเต็มเดือนของเดือนแรก ⚠️\n' +
        'กดปุ่มด้านล่างเพื่อเลือกวันที่เช็คอินได้เลยครับ/ค่ะ 📅';

      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        headers: {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer ' + token
        },
        payload: JSON.stringify({
          to: userId,
          messages: [{ type: 'text', text: promptMsg }]
        }),
        muteHttpExceptions: true
      });
      // -------------------------------------------------------------------

      // ส่งตัวเลือกวันที่ (date picker) ต่อทันที
      sendCheckinPickerToUser(userId, roomId);
      sent++;

      // หน่วงเล็กน้อยกัน rate-limit LINE (120ms ตามเดิม)
      Utilities.sleep(120);
    } catch(e) {
      Logger.log('Failed push to '+roomId+' ('+userId+'): '+e);
    }
  }
  const msg = `sendCheckinPickerToEligibleTenants: sent=${sent}, skipped=${skipped}, scanned=${vals.length}`;
  Logger.log(msg);
  return msg;
}



/** Reusable sender (pushes the datetime picker). */
function sendCheckinPickerToUser(userId, roomId) {
  const url = "https://api.line.me/v2/bot/message/push";
  const pickerData = _buildPostbackData_({ act: 'checkin_pick', room: roomId || '' }) || 'act=checkin_pick';
  const payload = {
    to: userId,
    messages: [{
      type: "template",
      altText: "เลือกวัน–เวลาเช็คอิน",
      template: {
        type: "confirm",
        text: `ห้อง ${roomId || "-"}\nโปรดเลือกวัน–เวลาเช็คอิน (${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้น)`,
        actions: [
          {
            type: "datetimepicker",
            label: "เลือกวัน–เวลา",
            data: pickerData,
            mode: "datetime",
            max: CHECKIN_PICKER_MAX_DATETIME
          },
          { type: "postback",       label: "ยกเลิก",        data: "act=rent_cancel" }
        ]
      }
    }]
  };
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('push -> ' + res.getResponseCode() + ' ' + res.getContentText());
}

/** === TEST: push to *your* LINE ID once === */
function testSendCheckinPickerToMe() {
  const MY_LINE_USER_ID = 'Ue90558b73d62863e2287ac32e69541a3'; // <- yours
  const roomId = _findRoomByUserId_(MY_LINE_USER_ID);          // optional, just for nicer text
  sendCheckinPickerToUser(MY_LINE_USER_ID, roomId);
}

function testCreateBookingPdfFromRow(rowNumber) {
  const rowIdx = Number(rowNumber) || 2;
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);
  const headers = _headersRow_(sh);
  const lastRow = sh.getLastRow();
  if (rowIdx < 2 || rowIdx > lastRow) {
    throw new Error('Row out of range: ' + rowIdx);
  }

  const cCode     = headers.findIndex(h => /^code$/i.test(h)) + 1 || 1;
  const cName     = _findCol_(headers, ['fullname','ชื่อ','ชื่อ-สกุล']) || 3;
  const cPhone    = _findCol_(headers, ['phone','โทร']) || 5;
  const cUserId   = _findUserIdCol_(headers);
  const cRoomId   = _findCol_(headers, ['room','room id','ห้อง']);
  const cPaidAmt  = _findCol_(headers, ['paid amount','ยอดที่ชำระแล้ว']);
  const cExpected = _findCol_(headers, ['expected amount','ยอดที่ต้องชำระ','amount']);
  const cRef      = _findCol_(headers, ['ref','reference','เลขอ้างอิง']);
  const cNotes    = _findCol_(headers, ['notes','หมายเหตุ']);

  const getVal = (col) => col ? String(sh.getRange(rowIdx, col).getValue() || '').trim() : '';

  const paidAmt = cPaidAmt ? Number(sh.getRange(rowIdx, cPaidAmt).getValue() || 0) : 0;
  const expAmt  = cExpected ? Number(sh.getRange(rowIdx, cExpected).getValue() || 0) : 0;

  const payload = {
    code: getVal(cCode),
    date: new Date(),
    customer: getVal(cName),
    phone: getVal(cPhone),
    line_id: getVal(cUserId),
    room_id: getVal(cRoomId),
    paid_amount: paidAmt || expAmt,
    bank_ref: getVal(cRef),
    move_in_month: '',
    move_in_window: '',
    contact_phone: getVal(cPhone),
    contact_line: '@MamaMansion',
    address: 'ฉลองกรุง 37 (ใกล้นิคมลาดกระบัง)',
    remaining_amount: 0,
    notes: getVal(cNotes)
  };

  const url = createBookingPdf_(payload);
  Logger.log('Booking PDF URL: ' + url);
  return url;
}

function createBookingPdf_(row) {
  // row = { code, date, customer, phone, line_id, room_id, paid_amount,
  //         bank_ref, move_in_month, move_in_window, contact_phone,
  //         contact_line, address, remaining_amount, notes }

  // --- 0) Read room price from Rooms sheet ---
  var ssRooms = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  var headR   = ssRooms.getRange(1,1,1, ssRooms.getLastColumn()).getValues()[0].map(String);
  var cRoomId = headR.findIndex(function(h){ return h.toLowerCase().includes('room'); }) + 1;
  var cPrice  = headR.findIndex(function(h){ return /(ราคาห้อง|room\s*price|price)/i.test(h); }) + 1;
  var roomPrice = 0;

  if (cRoomId && cPrice) {
    var vals = ssRooms.getRange(2,1, Math.max(0, ssRooms.getLastRow()-1), ssRooms.getLastColumn()).getValues();
    for (var i=0;i<vals.length;i++){
      var id = String(vals[i][cRoomId-1] || '').trim();
      if (id && id.toUpperCase() === String(row.room_id||'').toUpperCase()) {
        roomPrice = Number(vals[i][cPrice-1] || 0);
        break;
      }
    }
  }

  // --- 1) Calculate amounts ---
  var paid = Number(row.paid_amount || 0);

  // Move-in fees: +5000 - 2000 = 3000 (adjust here if policy changes)
  var feeJoin     = 5000;
  var feeDiscount = 2000;
  var feeNet      = feeJoin - feeDiscount; // = 3000

  var totalBeforePaid   = (roomPrice || 0) + feeNet;
  var remainingComputed = Math.max(totalBeforePaid - paid, 0);

  // ensure the object carries the computed remaining amount
  row.remaining_amount = remainingComputed;

  // human-readable calculation line
  var remainingCalcText = Utilities.formatString(
    '%s + %s - %s = %s บาท',
    (roomPrice||0).toLocaleString(),
    feeNet.toLocaleString(),
    paid.toLocaleString(),
    remainingComputed.toLocaleString()
  );

  // --- 2) Copy Google Doc template ---
  var srcFile  = DriveApp.getFileById(BOOKING_DOC_TEMPLATE_ID); // must be a Google Doc
  var copyName = 'Booking_' + row.code + '_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var copyFile = srcFile.makeCopy(copyName);
  var copyId   = copyFile.getId();

  // open with retry to avoid transient “inaccessible” errors
  var doc = null, tries = 0;
  while (!doc && tries < 5) {
    try { doc = DocumentApp.openById(copyId); }
    catch (e) { tries++; Utilities.sleep(300 * tries); }
  }
  if (!doc) throw new Error('Template copy could not be opened (inaccessible).');

  // --- 3) Replace placeholders in the Doc ---
  var body = doc.getBody();
  var fmt  = function(d){ return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); };

  body.replaceText('{{date}}',             fmt(row.date));
  body.replaceText('{{code}}',             row.code || '-');
  body.replaceText('{{customer}}',         row.customer || '-');
  body.replaceText('{{room_id}}',          row.room_id || '-');

  body.replaceText('{{paid_amount}}',      (paid||0).toLocaleString() + ' บาท');
  body.replaceText('{{bank_ref}}',         row.bank_ref || '-'); // remove from template if unused
  body.replaceText('{{move_in_month}}',    row.move_in_month || '-');
  body.replaceText('{{move_in_window}}',   row.move_in_window || '-');
  body.replaceText('{{contact_phone}}',    row.contact_phone || '-');
  body.replaceText('{{contact_line}}',     row.contact_line || '-');
  body.replaceText('{{address}}',          row.address || '-');
  body.replaceText('{{remaining_amount}}', (remainingComputed||0).toLocaleString() + ' บาท');
  body.replaceText('{{notes}}',            row.notes || '');

  // New calculation placeholders you can place in the template
  body.replaceText('{{room_price}}',            (roomPrice||0).toLocaleString());
  body.replaceText('{{fee_join}}',              feeJoin.toLocaleString());
  body.replaceText('{{fee_discount}}',          feeDiscount.toLocaleString());
  body.replaceText('{{room_price_plus_fee}}',   totalBeforePaid.toLocaleString());
  body.replaceText('{{remaining_calc}}',        remainingCalcText);

  doc.saveAndClose();

  // --- 4) Export PDF and share link ---
  var pdfBlob = DriveApp.getFileById(copyId).getAs(MimeType.PDF);
  var folder  = DriveApp.getFolderById(BOOKING_PDF_FOLDER_ID);
  var pdf     = folder.createFile(pdfBlob).setName(copyName + '.pdf');

  pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Optional cleanup of the working Doc:
  // DriveApp.getFileById(copyId).setTrashed(true);

  return pdf.getUrl();
}

function _rowLink_(row){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const gid = sh.getSheetId();
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${gid}&range=A${row}`;
}


// เมื่อสถานะเป็น Paid → ตั้งคิวที่ Rooms (คอลัมน์ Horganice Done? = FALSE)
function _markRoomNeedsHorganice_(roomId, payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shRooms = ss.getSheetByName('Rooms');
  if (!shRooms) return;

  const head = shRooms.getRange(1,1,1,shRooms.getLastColumn()).getValues()[0].map(String);
  const cId   = head.findIndex(h => h.toLowerCase().includes('room')) + 1;
  const cDone = _findCol_(head, ['horganice done?', 'hg done?', 'สร้างใน horganice แล้ว?', 'done?']);
  const cAt   = _findCol_(head, ['horganice at', 'hg at', 'เวลาบันทึก horganice']);
  const cCode = _findCol_(head, ['hg code','horganice code']);
  const cName = _findCol_(head, ['hg name','horganice name','ชื่อลูกค้า']);
  const cPhone= _findCol_(head, ['hg phone','horganice phone','โทรลูกค้า']);
  const cLink = _findCol_(head, ['hg row link','horganice row link','ลิงก์แถว']);
  if (!cId || !cDone) return;

  const last = shRooms.getLastRow();
  const vals = shRooms.getRange(2,1, last-1, shRooms.getLastColumn()).getValues();
  let r = -1;
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][cId-1]).trim() === roomId) { r = i+2; break; }
  }
  if (r < 2) return;

  shRooms.getRange(r, cDone).setValue(false);   // ยังไม่เสร็จ
  if (cAt)   shRooms.getRange(r, cAt).setValue('');
  if (cCode) shRooms.getRange(r, cCode).setValue(payload.code || '');
  if (cName) shRooms.getRange(r, cName).setValue(payload.name || '');
  if (cPhone)shRooms.getRange(r, cPhone).setValue(payload.phone || '');
  if (cLink) shRooms.getRange(r, cLink).setValue(_rowLink_(payload.rowIndex));
}

// ใส่เวลาอัตโนมัติเมื่อ “Horganice Done?” ถูกติ๊ก
function _timestampWhenRoomHgDone_(shRooms, row) {
  const head = shRooms.getRange(1,1,1,shRooms.getLastColumn()).getValues()[0].map(String);
  const cDone = _findCol_(head, ['horganice done?', 'hg done?', 'สร้างใน horganice แล้ว?', 'done?']);
  const cAt   = _findCol_(head, ['horganice at', 'hg at', 'เวลาบันทึก horganice']);
  if (!cDone || !cAt) return;

  const done = !!shRooms.getRange(row, cDone).getValue();
  shRooms.getRange(row, cAt).setValue(done ? new Date() : '');
}


// [ADD] helper: set value to Rooms.Parking by RoomId
function _setRoomParking_(roomId, value) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Rooms");
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  const header = values.shift().map(String);

  const cId = header.findIndex(h => h.toLowerCase().includes("room"));
  const cParking = header.findIndex(h => h.toLowerCase().includes("parking")); // ชื่อคอลัมน์ใน Rooms = "Parking"
  if (cId < 0 || cParking < 0) return;

  for (let r = 0; r < values.length; r++) {
    const id = String(values[r][cId] || "").trim();
    if (id === roomId) {
      sh.getRange(r + 2, cParking + 1).setValue(value);
      return;
    }
  }
}

// อ่านราคาห้องจากชีต Rooms (รองรับหัวข้อไทย/อังกฤษ)
function _getRoomPriceFromRooms_(roomId) {
  if (!roomId) return 0;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Rooms');
  if (!sh) return 0;

  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(h => String(h || '').trim().toLowerCase());

  const cId = headers.findIndex(h => h.includes('room')); // RoomID
  // หา column ราคาโดยดูได้ทั้ง "ราคาห้อง", "price", "room price", "expected amount"
  const aliases = ['ราคาห้อง', 'price', 'room price', 'expected amount', 'amount', 'monthly'];
  let cPrice = -1;
  for (const a of aliases) {
    const i = headers.findIndex(h => h.indexOf(a.toLowerCase()) !== -1);
    if (i > -1) { cPrice = i; break; }
  }
  if (cId < 0 || cPrice < 0) return 0;

  for (const row of values) {
    const id = String(row[cId] || '').trim().toUpperCase();
    if (id === String(roomId).toUpperCase()) {
      const p = Number(row[cPrice] || 0);
      return isNaN(p) ? 0 : p;
    }
  }
  return 0;
}

function dailyMoveOutSweep() {
  try {
    const result = markSoonStatuses_(DAYS_AHEAD_SOON);  // 90 by default
    Logger.log(result);
    return result;
  } catch (err) {
    Logger.log('dailyMoveOutSweep ERROR: ' + err);
    try { sendLineMessage('❗dailyMoveOutSweep ERROR: ' + err); } catch (_e) {}
    return 'dailyMoveOutSweep ERROR';
  }
}

/** Run daily (e.g., 09:00 Asia/Bangkok) to remind tomorrow's check-ins */
function dailyCheckinReminder() {
  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Rooms');
  if (!sh) return;

  const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const cRoom  = H.findIndex(h=>h.toLowerCase().includes('room')) + 1;
  const cUser  = H.findIndex(h=>h.toLowerCase().includes('line id')) + 1;
  const cDate  = H.indexOf('CheckinDate') + 1;
  const cTime  = H.indexOf('CheckinTime') + 1;
  const cConf  = H.indexOf('CheckinConfirmed') + 1;
  let cSent    = H.indexOf('ReminderSent') + 1;
  let cSentAt  = H.indexOf('ReminderSentAt') + 1;

  if (!cRoom || !cUser || !cDate || !cTime || !cConf) return; // missing required columns

  const last = sh.getLastRow();
  if (last < 2) return;

  const rng = sh.getRange(2,1,last-1,sh.getLastColumn());
  const vals = rng.getValues();

  // helper to compare YYYY-MM-DD only
  const ymd = d => Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), d.getDate()), tz, 'yyyy-MM-dd');
  const today = new Date(); today.setHours(0,0,0,0);
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate()+1);

  const out = vals.map(r => r.slice()); // copy for bulk write

  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const userId = String(row[cUser-1]||'').trim();
    const roomId = String(row[cRoom-1]||'').trim();
    const ckd = row[cDate-1];
    const ckt = String(row[cTime-1]||'').trim();
    const confirmed = !!row[cConf-1];

    if (!userId || !roomId || !confirmed) continue;
    if (!(ckd instanceof Date)) continue;

    // D-1 reminder
    if (ymd(ckd) === ymd(tomorrow)) {
      const already = cSent ? String(row[cSent-1]||'').toUpperCase() : '';
      if (already !== 'D-1') {
        pushLineToUser_(userId,
          "แจ้งเตือนล่วงหน้า 1 วันค่ะ 🗓️\n" +
          `ห้อง: ${roomId}\n` +
          `วันเช็คอิน: ${Utilities.formatDate(ckd, tz, 'dd/MM/yyyy')}\n` +
          (ckt ? `เวลา: ${ckt}\n` : '') +
          "หากต้องการเปลี่ยนเวลา โปรดแจ้งกลับทางแชตค่ะ"
        );
        if (cSent)   out[i][cSent-1] = 'D-1';
        if (cSentAt) out[i][cSentAt-1] = new Date();
      }
    }
  }

  // write back if we have the columns
  if (cSent || cSentAt) rng.setValues(out);
}


function markSoonStatuses_(daysAhead) {
  daysAhead = Number(daysAhead) || 90;

  const ss = SpreadsheetApp.openById(SHEET_ID);           // ✅ use SHEET_ID
  const sh = ss.getSheetByName(SHEET_ROOMS);
  if (!sh) throw new Error('Rooms sheet not found: ' + SHEET_ROOMS);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 'No data rows';

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0]
    .map(h => String(h||'').trim().toLowerCase());

  const findCol = (aliases) => {
    for (let c = 0; c < headers.length; c++) {
      for (const a of aliases) {
        if (headers[c].indexOf(String(a).toLowerCase()) !== -1) return c + 1;
      }
    }
    return 0;
  };

  const cRoom = findCol(['room','room id','ห้อง']);
  const cStat = findCol(['status','สถานะ']);
  // ✅ use your global alias list for move-out date
  const cMove = findCol(ROOMS_MOVEOUT_ALIASES);
  if (!cRoom || !cStat || !cMove) {
    throw new Error('Required columns not found (Room / Status / Move-out Date)');
  }

  const rng    = sh.getRange(2, 1, lastRow-1, lastCol);
  const values = rng.getValues();

  const today = new Date(); today.setHours(0,0,0,0);
  const max   = new Date(today); max.setDate(max.getDate() + daysAhead);

  let setSoon = 0, setAvail = 0;

  // We’ll write back only where needed
  const out = values.map((row) => row.slice()); // shallow clone rows

  for (let i = 0; i < values.length; i++) {
    const mov = values[i][cMove - 1];
    if (!(mov instanceof Date)) continue;   // skip if no valid date

    const movDate = new Date(mov.getFullYear(), mov.getMonth(), mov.getDate());
    const cur     = String(values[i][cStat - 1] || '').trim().toLowerCase();

    if (movDate >= today && movDate <= max) {
      if (cur !== 'soon' && cur !== 'กำลังจะว่าง') {
        out[i][cStat - 1] = 'soon';
        setSoon++;
      }
      continue;
    }

    if (movDate < today) {
      if (cur !== 'avail' && cur !== 'ว่าง') {
        out[i][cStat - 1] = 'avail';
        setAvail++;
      }
    }
  }

  // Bulk write only if there were changes
  if (setSoon || setAvail) {
    rng.setValues(out);
  }

  SpreadsheetApp.flush();
  return `Updated: soon=${setSoon}, avail=${setAvail}`;
}

// สร้างชีต + เฮดเดอร์ (ถ้ายังไม่มี)
function _ensurePrebookSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(PREBOOK_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PREBOOK_SHEET_NAME);
    sh.appendRow([
      'Code', 'Timestamp', 'Fullname', 'Line ID', 'Phone',
      'Parking', 'Move-in Month', 'Notes', 'Status'
    ]);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow([
      'Code', 'Timestamp', 'Fullname', 'Line ID', 'Phone',
      'Parking', 'Move-in Month', 'Notes', 'Status'
    ]);
  }
  return sh;
}

// สร้างโค้ดใหม่ #PBxxx
function _nextPrebookCode_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return PREBOOK_CODE_PREFIX + '000';
  const lastCode = String(sh.getRange(lastRow, 1).getValue() || '').trim();
  const num = parseInt(lastCode.replace(PREBOOK_CODE_PREFIX, ''), 10) || 0;
  return PREBOOK_CODE_PREFIX + String(num + 1).padStart(3, '0');
}







const RICHMENU_ID_RAW = '809f92d6bbba5cc330d0a89f92323a3a';           // your menu
const DRIVE_IMAGE_FILE_ID = '1DZ-uDyKFt9LJplT8ieDNDBbD6LL7LFSz';         // your Drive image (2500×1686, <=1MB)

/**************** CREATE NEW MENU ****************/
function createTenantPaidRichMenu_() {
  const payload = {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: 'Tenant Paid Menu ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmm'),
    chatBarText: 'เมนูผู้เช่า',
    areas: [
      {
        bounds: { x: 0, y: 0, width: 833, height: 843 },
        action: { type: 'message', text: 'ฟีดแบ็ก (ไม่ระบุตัวตน)' }
      },
      {
        bounds: { x: 833, y: 0, width: 834, height: 843 },
        action: { type: 'message', text: 'จ่ายค่าเช่า' }
      },
      {
        bounds: { x: 1667, y: 0, width: 833, height: 843 },
        action: { type: 'message', text: 'แจ้งซ่อม' }
      },
      {
        bounds: { x: 0, y: 843, width: 833, height: 843 },
        action: { type: 'message', text: 'คู่มือผู้เช่า' }
      },
      {
        bounds: { x: 833, y: 843, width: 834, height: 843 },
        action: { type: 'message', text: 'บริการเสริม' }
      },
      {
        bounds: { x: 1667, y: 843, width: 833, height: 843 },
        action: { type: 'message', text: 'ติดต่อ' }
      }
    ]
  };

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  Logger.log('CREATE -> ' + code + ' ' + body);
  if (code < 200 || code >= 300) {
    throw new Error('Create rich menu failed: ' + code + ' ' + body);
  }

  const json = JSON.parse(body || '{}');
  return json.richMenuId || '';
}

/**************** RUN THIS ****************/
function uploadTenantMenuImageTo_(richMenuIdRaw) {
  const richMenuId = RICH(richMenuIdRaw);

  // 1) verify menu & size
  const meta = fetchL('https://api.line.me/v2/bot/richmenu/' + richMenuId, { method: 'get' });
  if (meta.getResponseCode() !== 200) {
    throw new Error('Menu not found: ' + meta.getContentText());
  }
  const m = JSON.parse(meta.getContentText() || '{}');
  if (!(m.size && m.size.width === 2500 && m.size.height === 1686)) {
    throw new Error('Menu must be 2500x1686; got ' + JSON.stringify(m.size));
  }

  // 2) load image
  const file = DriveApp.getFileById(DRIVE_IMAGE_FILE_ID);
  const mime = String(file.getMimeType() || '').toLowerCase();
  const isPng = mime.indexOf('png') !== -1;
  const isJpg = mime.indexOf('jpeg') !== -1 || mime.indexOf('jpg') !== -1;
  if (!isPng && !isJpg) {
    throw new Error('Image must be PNG or JPEG; got ' + mime);
  }

  const blob = file.getBlob().setContentType(isPng ? 'image/png' : 'image/jpeg');
  const bytes = blob.getBytes();
  if (bytes.length > 1000000) {
    throw new Error('Image too large: ' + bytes.length + ' bytes');
  }

  // 3) upload to api-data host (this was the issue)
  const url = 'https://api-data.line.me/v2/bot/richmenu/' + richMenuId + '/content';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: isPng ? 'image/png' : 'image/jpeg',
    payload: bytes,
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    muteHttpExceptions: true
  });
  Logger.log('UPLOAD -> ' + res.getResponseCode() + ' ' + res.getContentText());
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    throw new Error('Upload failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  Logger.log('✅ Upload OK for ' + richMenuId + ' (not linked to anyone)');
  return richMenuId;
}

function uploadTenantMenuImage() {
  return uploadTenantMenuImageTo_(RICHMENU_ID_RAW);
}

function createAndUploadTenantMenu_() {
  const richMenuId = createTenantPaidRichMenu_();
  if (!richMenuId) throw new Error('Rich menu creation failed');
  uploadTenantMenuImageTo_(richMenuId);
  Logger.log('🎯 New tenant menu ready: ' + richMenuId + ' — update PAID_MENU_ID & RICHMENU_ID_RAW');
  return richMenuId;
}

/**************** HELPERS ****************/
function RICH(raw){ const s = String(raw||'').trim(); return s.startsWith('richmenu-') ? s : 'richmenu-' + s; }
function fetchL(url,opt){
  const res = UrlFetchApp.fetch(url, Object.assign({
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    muteHttpExceptions: true
  }, opt||{}));
  Logger.log((opt&&opt.method||'get') + ' ' + url + ' -> ' + res.getResponseCode());
  return res;
}

/**************** CHECK-IN FEE (WEEK-AHEAD NOTICE) ****************/
const CHECKIN_NOTICE = {
  ADVANCE_DAYS: 7,
  ROUND_TO: 10,
  DEPOSIT_NET: 3000,      // 5,000 - 2,000 = 3,000
  FRIDGE_FEE: 200,        // per month
  CAR_PARKING: {          // normalize แล้วจะ map มาที่คีย์พวกนี้
    'car_roof': 800,      // มีหลังคา
    'car_noroof': 500,    // ไม่มีหลังคา
  },
  ROOMS_SHEET: 'Rooms',
  COLS: {
    roomId: 'RoomID',
    lineId: 'Line ID',
    price: 'ราคาห้อง',
    fridge: 'Fridge',
    parking: 'Parking',
    checkinDate: 'CheckinDate',
    checkinTime: 'CheckinTime',
    checkinConfirmed: 'CheckinConfirmed',
    sentWeek: 'ReminderSent_Week',
    sentWeekAt: 'ReminderSentAt_Week',
  }
};

function _findColStrict_(headers, name) {
  const idx = headers.indexOf(name);
  if (idx < 0) throw new Error('Missing column: ' + name);
  return idx + 1;
}
function _daysInMonth_(d){ return new Date(d.getFullYear(), d.getMonth()+1, 0).getDate(); }
function _firstMonthUsedDays_(checkinDate){
  const dim = _daysInMonth_(checkinDate);
  return dim - checkinDate.getDate() + 1; // รวมวันเช็คอิน
}
function _roundUp_(num, to){ return Math.ceil(num / to) * to; }
function _fmtBaht_(n){ return Number(n||0).toLocaleString('th-TH', {maximumFractionDigits:0}); }
function _thaiDate_(d){ return Utilities.formatDate(d, 'Asia/Bangkok', 'dd MMM yyyy'); }

// แปลงค่าที่จอดรถใน Rooms ให้มาอยู่ในชุดคีย์ที่รองรับ
function _normalizeCarParking_(txt) {
  const parsed = _parseParkingCellValue_(txt);
  if (!parsed.wantsParking) return '';
  if (parsed.plan === 'roofed') return 'car_roof';
  if (parsed.plan === 'open') return 'car_noroof';
  return ''; // default = ไม่คิดเงินเพิ่ม
}

// แปลงค่าตู้เย็นให้เป็น boolean
function _hasFridge_(txt) {
  const t = String(txt||'').trim().toLowerCase();
  return ['yes','true','1','มี','y'].includes(t);
}

// คำนวณยอดเดือนแรก (พรอเรต + ประกัน + ที่จอด + ตู้เย็น)
function computeCheckinBill_(rentPerMonth, checkinDate, parkingRaw, fridgeRaw) {
  const dim = _daysInMonth_(checkinDate);
  const used = _firstMonthUsedDays_(checkinDate);
  const daily = rentPerMonth / dim;

  const prorated = _roundUp_(daily * used, CHECKIN_NOTICE.ROUND_TO);
  const deposit  = CHECKIN_NOTICE.DEPOSIT_NET;

  const parkingKey = _normalizeCarParking_(parkingRaw);
  const parkingFee = CHECKIN_NOTICE.CAR_PARKING[parkingKey] || 0;

  const fridgeFee = _hasFridge_(fridgeRaw) ? CHECKIN_NOTICE.FRIDGE_FEE : 0;

  const total = prorated + deposit + parkingFee + fridgeFee;

  return {
    dim, used,
    dailyRoundedHint: Math.ceil(daily),  // สำหรับโชว์ ~บาท/วัน
    prorated, deposit, parkingFee, fridgeFee, total
  };
}

/**
 * First-month charge (silent policy):
 * - Nov–Dec 2025 → prorate (uses computeCheckinBill_)
 * - 2026+        → full month rent + deposit + add-ons (no prorate)
 *
 * Returns a normalized object with fields you already use in messages:
 * { mode, prorated, deposit, parkingFee, fridgeFee, total, used, dim, dailyRoundedHint }
 */
function computeFirstMonthBill_(rentPerMonth, checkinDate, parkingRaw, fridgeRaw) {
  const rentValue = Number(rentPerMonth) || 0;
  const halfRent = _roundUp_(rentValue / 2, CHECKIN_NOTICE.ROUND_TO);
  const deposit = CHECKIN_NOTICE.DEPOSIT_NET;
  const parkingKey = _normalizeCarParking_(parkingRaw);
  const parkingFee = CHECKIN_NOTICE.CAR_PARKING[parkingKey] || 0;
  const fridgeFee  = _hasFridge_(fridgeRaw) ? CHECKIN_NOTICE.FRIDGE_FEE : 0;

  if (!_isFullMonthEra_(checkinDate)) {
    // ก่อน 1 ม.ค. 2026 → เก็บครึ่งเดือนตายตัว
    const totalHalf = halfRent + deposit + parkingFee + fridgeFee;
    return {
      mode: 'half_prorate',
      prorated: halfRent,
      deposit,
      parkingFee,
      fridgeFee,
      total: totalHalf,
      used: null, dim: null, dailyRoundedHint: null
    };
  }

  // 1 ม.ค. 2026 เป็นต้นไป → ปกติเต็มเดือน แต่ถ้าเข้าอยู่หลังวันที่ 15 เก็บครึ่งเดือน
  const lateCheckin = checkinDate.getDate() > 15;
  const rentComponent = lateCheckin ? halfRent : rentValue;
  const total = rentComponent + deposit + parkingFee + fridgeFee;

  return {
    mode: lateCheckin ? 'half_full' : 'full',
    prorated: rentComponent,
    deposit,
    parkingFee,
    fridgeFee,
    total,
    used: null, dim: null, dailyRoundedHint: null
  };
}


function sendWeekAheadCheckinFees() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(CHECKIN_NOTICE.ROOMS_SHEET);
  if (!sh) throw new Error('Rooms not found');

  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const cRoom   = _findColStrict_(headers, CHECKIN_NOTICE.COLS.roomId);
  const cLine   = _findColStrict_(headers, CHECKIN_NOTICE.COLS.lineId);
  const cPrice  = _findColStrict_(headers, CHECKIN_NOTICE.COLS.price);
  const cPark   = _findColStrict_(headers, CHECKIN_NOTICE.COLS.parking);
  const cFridge = _findColStrict_(headers, CHECKIN_NOTICE.COLS.fridge);
  const cDate   = _findColStrict_(headers, CHECKIN_NOTICE.COLS.checkinDate);
  const cTime   = headers.indexOf(CHECKIN_NOTICE.COLS.checkinTime) + 1;         // optional
  const cConf   = headers.indexOf(CHECKIN_NOTICE.COLS.checkinConfirmed) + 1;    // optional
  let   cSent   = headers.indexOf(CHECKIN_NOTICE.COLS.sentWeek) + 1;            // create if missing
  let   cSentAt = headers.indexOf(CHECKIN_NOTICE.COLS.sentWeekAt) + 1;

  // สร้างคอลัมน์ ReminderSent_Week/ReminderSentAt_Week อัตโนมัติถ้าไม่มี
  if (!cSent)   { sh.insertColumnAfter(headers.length); cSent   = headers.length+1; sh.getRange(1,cSent).setValue(CHECKIN_NOTICE.COLS.sentWeek); headers.push(CHECKIN_NOTICE.COLS.sentWeek); }
  if (!cSentAt) { sh.insertColumnAfter(headers.length); cSentAt = headers.length+1; sh.getRange(1,cSentAt).setValue(CHECKIN_NOTICE.COLS.sentWeekAt); headers.push(CHECKIN_NOTICE.COLS.sentWeekAt); }

  const rows = sh.getLastRow() - 1;
  if (rows < 1) return 'No data rows';
  const rng  = sh.getRange(2,1, rows, sh.getLastColumn());
  const vals = rng.getValues();
  const out  = vals.map(r => r.slice());

  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) throw new Error('Missing LINE_TOKEN');

  const tz = 'Asia/Bangkok';
  const today = new Date(); today.setHours(0,0,0,0);
  const ymdToday = _ymd_(today, tz);

  let sent = 0, skipped = 0;

  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const lineId = String(row[cLine-1]||'').trim();
    const roomId = String(row[cRoom-1]||'').trim();
    const price  = Number(row[cPrice-1]||0);
    const ckd    = row[cDate-1];

    if (!lineId || !roomId || !price || !(ckd instanceof Date)) { skipped++; continue; }

    // ถ้ามีคอลัมน์ Confirmed ให้ส่งเฉพาะคนที่ confirm แล้ว
    if (cConf && !row[cConf-1]) { skipped++; continue; }

    // กันยิงซ้ำแบบผูกกับ "CheckinDate ปัจจุบัน"
    const sentMark = String(row[cSent-1]||'').trim(); // expected = 'yyyy-MM-dd'
    const ymdTarget = _ymd_(ckd, tz);
    if (sentMark === ymdTarget) { skipped++; continue; }

    const target = new Date(ckd); target.setHours(0,0,0,0);
    const remind = new Date(target); remind.setDate(remind.getDate() - CHECKIN_NOTICE.ADVANCE_DAYS);
    if (_ymd_(remind, tz) !== ymdToday) { skipped++; continue; }

    // คำนวณ
    const calc = computeFirstMonthBill_(price, target, row[cPark-1], row[cFridge-1]);

    const msg = _weekAheadMessage_(roomId, target, calc);

    try {
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        headers: { 'Content-Type':'application/json; charset=UTF-8', Authorization:'Bearer '+token },
        payload: JSON.stringify({ to: lineId, messages: [{ type:'text', text: msg }] }),
        muteHttpExceptions: true
      });

      // after LINE push succeeded
const monthStr = Utilities.formatDate(target, tz, 'yyyy-MM');
const billId = 'CHKIN-' + Utilities.formatDate(target, tz, 'yyyyMMdd') + '-' + roomId;
const tenant  = _getTenantNameByRoom_(roomId) || '';    // best effort from Rooms
const account = getAccountFromRoom_(roomId);            // your mapping logic
const chargeItems = _chargeItemsFromCalc_(calc);

// Upsert into prediction sheet (safe sandbox for reconciliation)
_upsertHorgaBill_('Horga_Bills', {
  BillID: billId,
  Room: roomId,
  Tenant: tenant,
  Month: monthStr,
  AmountDue: calc.total,
  DueDate: _ymd_(target, tz),
  Status: 'Unpaid',
  PaidAt: '',
  SlipID: '',
  Account: account,
  ChargeItems: chargeItems,
  Notes: 'Week-ahead first-month bill (auto)'
});


      out[i][cSent-1]   = ymdTarget; 
      out[i][cSentAt-1] = new Date();
      sent++;
      Utilities.sleep(120);
    } catch(e) {
      Logger.log('Week-ahead push failed for '+roomId+' -> '+e);
    }
  }

  rng.setValues(out);
  const res = `sendWeekAheadCheckinFees: sent=${sent}, skipped=${skipped}, scanned=${vals.length}`;
  Logger.log(res);
  return res;
}

function _ymd_(d, tz) {
  return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), d.getDate()), tz || 'Asia/Bangkok', 'yyyy-MM-dd');
}

// Floor → Account mapping
function getAccountFromRoom_(roomId) {
  // Normalize like "B504" / "A1203" / "C206" / "B 504" → "B504"
  const roomUpper = String(roomId || '').toUpperCase().replace(/\s+/g, '');
  if (!roomUpper) return '';

  // Optional letters at start, then capture the FIRST digit = floor
  // e.g. B504 → '5', A1203 → '1', 305 → '3'
  const m = roomUpper.match(/^[A-Z]*(\d)/);
  if (!m || !m[1]) return '';

  switch (m[1]) {
    case '1': return 'KKK+';
    case '2': return 'TMK+';
    case '3': return 'KGSI';
    case '4': return 'KBIZ';
    case '5': return 'KBIZ';
    default:  return '';        // floors 0, 6, 7, ... → no account
  }
}


function testWeekAhead_MyRoom() {
  return testWeekAheadForRoom('B504'); // <-- put YOUR RoomID here
}


/** สำหรับเทสต์ห้องเดียวแบบ manual */
function testWeekAheadForRoom(roomId){
  const { sh, H } = _roomsHeaders_();
  const cRoom = H.indexOf('RoomID') + 1;
  const cDate = H.indexOf('CheckinDate') + 1;
  for (let i=2;i<=sh.getLastRow();i++){
    if (String(sh.getRange(i,cRoom).getValue()).trim().toUpperCase() === roomId.toUpperCase()){
      const d = new Date(); d.setDate(d.getDate() + CHECKIN_NOTICE.ADVANCE_DAYS);
      sh.getRange(i, cDate).setValue(d);
      return sendWeekAheadCheckinFees();
    }
  }
  throw new Error('Room not found: '+roomId);
}


// ล้างธง ReminderSent_Week สำหรับห้องเดียว (ใช้ตอนอยากรันทดสอบซ้ำ)
function clearWeekAheadFlagsForRoom(roomId) {
  const { sh, H } = _roomsHeaders_();
  const cRoom   = H.indexOf('RoomID') + 1;
  const cSent   = H.indexOf('ReminderSent_Week') + 1;
  const cSentAt = H.indexOf('ReminderSentAt_Week') + 1;
  if (!cRoom) throw new Error('RoomID col not found');
  if (!cSent && !cSentAt) return 'No week flags found';

  const vals = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][cRoom-1]||'').trim().toUpperCase() === String(roomId).toUpperCase()){
      if (cSent)   sh.getRange(i+2, cSent).setValue('');
      if (cSentAt) sh.getRange(i+2, cSentAt).setValue('');
      return 'Cleared for ' + roomId;
    }
  }
  return 'Room not found: ' + roomId;
}


function _weekAheadRentLine_(calc) {
  const rentValue = _fmtBaht_(calc.prorated);
  switch (calc.mode) {
    case 'full':
      return `• 🏠 ค่าเช่าเต็มเดือน: ${rentValue} บ.`;
    case 'half_full':
      return `• 🏠 ค่าเช่า (ครึ่งเดือน เพราะเช็คอินหลังวันที่ 15): ${rentValue} บ.`;
    case 'half_prorate':
      return `• 🏠 ค่าเช่า (ครึ่งเดือนตามนโยบายปี 2025): ${rentValue} บ.`;
    default: {
      const hasProrateDetail = calc.used && calc.dim && calc.dailyRoundedHint;
      if (hasProrateDetail) {
        return `• 🏠 ค่าเช่าพรอเรต: ${rentValue} บ.  (อยู่ ${calc.used} วัน ~${_fmtBaht_(calc.dailyRoundedHint)}/วัน จาก ${calc.dim} วัน)`;
      }
      return `• 🏠 ค่าเช่า: ${rentValue} บ.`;
    }
  }
}

function _weekAheadMessage_(roomId, targetDate, calc) {
  const GAP = '\u200B'; // keeps line spacing in LINE
  const depositLabel = 'ค่าประกัน (5000 - ค่าจอง2000)';

  const lines = [
    `เรียนลูกค้า ห้อง ${roomId} 🙏`,
    `ใกล้ถึงวันเช็คอินแล้ว 🗓️ (${_thaiDate_(targetDate)})`,
    GAP,

    `ยอดต้องเตรียมสำหรับเดือนแรก (ปัดขึ้นหลักสิบ):`,
    _weekAheadRentLine_(calc),
    `• 🔒 ${depositLabel}: ${_fmtBaht_(calc.deposit)} บ.`,
    (calc.parkingFee ? `• 🚗 ค่าที่จอดรถ: ${_fmtBaht_(calc.parkingFee)} บ.` : null),
    (calc.fridgeFee  ? `• 🧊 ค่าเช่าตู้เย็น: ${_fmtBaht_(calc.fridgeFee)} บ.` : null),
    GAP,

    `💳 รวมที่ต้องชำระวันเข้าอยู่: **${_fmtBaht_(calc.total)} บ.**`,
    GAP,

    `🧾 เอกสารที่ต้องเตรียม: บัตรประชาชนตัวจริงและสำเนา 1 ชุด`,
    `✏️ หากต้องการเปลี่ยนวัน/เวลาเช็คอิน พิมพ์ "เปลี่ยนวันเช็คอิน" ได้เลย`
  ];

  return lines.filter(v => v !== null).join('\n'); // keep GAP/empty lines
}



function forceSendFull2026ToMe(){
  const MY_LINE_USER_ID = 'Ue90558b73d62863e2287ac32e69541a3'; // <- yours
  const roomId = 'B504';
  const d = new Date('2026-01-03T00:00:00+07:00');

  const { sh, H } = _roomsHeaders_();
  const cRoom   = H.indexOf(CHECKIN_NOTICE.COLS.roomId) + 1;
  const cPrice  = H.indexOf(CHECKIN_NOTICE.COLS.price) + 1;
  const cPark   = H.indexOf(CHECKIN_NOTICE.COLS.parking) + 1;
  const cFridge = H.indexOf(CHECKIN_NOTICE.COLS.fridge) + 1;

  const vals = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][cRoom-1]||'').trim().toUpperCase() === roomId.toUpperCase()){
      const price  = Number(vals[i][cPrice-1]||0);
      const park   = vals[i][cPark-1];
      const fridge = vals[i][cFridge-1];
      const calc   = computeFirstMonthBill_(price, d, park, fridge);
      const msg    = _weekAheadMessage_(roomId, d, calc);

      // push to YOU (not the tenant), no flags written
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        headers: { 'Content-Type':'application/json; charset=UTF-8', Authorization:'Bearer '+LINE_TOKEN },
        payload: JSON.stringify({ to: MY_LINE_USER_ID, messages: [{ type:'text', text: msg }] }),
        muteHttpExceptions: true
      });
      return 'Sent preview to me for ' + roomId;
    }
  }
  throw new Error('Room not found: ' + roomId);
}

const HORGA_SCHEMA = [
  'BillID','Room','Tenant','Month','AmountDue','DueDate',
  'Status','PaidAt','SlipID','Account','ChargeItems','Notes'
];

// generic: ensure sheet exists in a specific spreadsheet
function _ensureSheetWithHeaderIn_(spreadsheetId, sheetName, schema) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,schema.length).setValues([schema]);
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,schema.length).setValues([schema]);
  } else {
    const head = sh.getRange(1,1,1, Math.max(schema.length, sh.getLastColumn())).getValues()[0];
    const same = schema.every((h,i)=> String(head[i]||'') === h);
    if (!same) sh.getRange(1,1,1,schema.length).setValues([schema]);
  }
  return sh;
}

// ✅ Single source of truth: always write to Revenue Master / Horga_Bills
function _upsertHorgaBill_(sheetName, billObj) {
  const sh = _ensureSheetWithHeaderIn_(REVENUE_MASTER_ID, sheetName, HORGA_SCHEMA);

  // map BillID -> row
  const last = sh.getLastRow();
  const idCol = 1; // A
  let hit = 0;
  if (last > 1) {
    const ids = sh.getRange(2, idCol, last-1, 1).getValues().map(r => String(r[0]||'').trim());
    const target = String(billObj.BillID||'').trim();
    const idx = target ? ids.findIndex(v => v === target) : -1;
    if (idx >= 0) hit = idx + 2;
  }
  const rowArr = HORGA_SCHEMA.map(k => billObj[k] != null ? billObj[k] : '');
  if (hit) {
    sh.getRange(hit, 1, 1, HORGA_SCHEMA.length).setValues([rowArr]);
    return { action: 'updated', row: hit };
  } else {
    sh.appendRow(rowArr);
    return { action: 'inserted', row: sh.getLastRow() };
  }
}


// Try to get a tenant name from Rooms (best-effort)
function _getTenantNameByRoom_(roomId) {
  const { sh, H, Hl } = _roomsHeaders_();
  const cRoom = Hl.findIndex(h => h.includes('room')) + 1;
  if (!cRoom) return '';

  const nameAliases = [
    'hg name','horganice name','ชื่อลูกค้า','ลูกค้า','tenant',
    'fullname','full name','ชื่อ','ชื่อ-สกุล'
  ];
  let cName = 0;
  for (const a of nameAliases) {
    const j = Hl.findIndex(h => h.indexOf(String(a).toLowerCase()) !== -1) + 1;
    if (j) { cName = j; break; }
  }
  if (!cName) return '';

  const rows = sh.getLastRow() - 1;
  if (rows < 1) return '';
  const vals = sh.getRange(2,1,rows, sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const id = String(vals[i][cRoom-1]||'').trim();
    if (id.toUpperCase() === String(roomId).toUpperCase()) {
      return String(vals[i][cName-1]||'').trim();
    }
  }
  return '';
}

// Turn your calc object into a compact "ChargeItems" string
function _chargeItemsFromCalc_(calc) {
  const parts = [];
  if (calc.mode === 'full') {
    parts.push(`FirstMonthRent ${calc.prorated||0}`);
  } else {
    parts.push(`ProratedRent ${calc.prorated||0}`);
  }
  parts.push(`DepositNet ${calc.deposit||0}`);
  if (calc.parkingFee) parts.push(`Parking ${calc.parkingFee}`);
  if (calc.fridgeFee)  parts.push(`Fridge ${calc.fridgeFee}`);
  return parts.join('; ');
}
