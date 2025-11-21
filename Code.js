/**************************************
 * Mama Mansion LINE BOT (Apps Script)
 * ROLE: Heavy tasks only (push-only)
 **************************************/

const PROPS = PropertiesService.getScriptProperties();
const WORKER_SECRET = PROPS.getProperty('WORKER_SECRET') || '';
const CHANNEL_ACCESS_TOKEN = PROPS.getProperty('CHANNEL_ACCESS_TOKEN');
const LINE_NOTIFY_TOKEN = PROPS.getProperty('LINE_NOTIFY_TOKEN'); // create a LINE Notify token for admin/group

const SHEET_ID            = PROPS.getProperty('SHEET_ID');
const SHEET_NAME          = PROPS.getProperty('SHEET_NAME') || 'Sheet1';

// 🔸 NEW: Payment Slip folders
const SLIP_FOLDER_ID      = PROPS.getProperty('SLIP_FOLDER_ID');
const TEMP_SLIP_FOLDER_ID = PROPS.getProperty('TEMP_SLIP_FOLDER_ID');

// 🔸 NEW: ID Card folders
const ID_FOLDER_ID        = PROPS.getProperty('ID_FOLDER_ID');
const TEMP_ID_FOLDER_ID   = PROPS.getProperty('TEMP_ID_FOLDER_ID');

const QR_IMAGE_FILE_ID    = PROPS.getProperty('QR_IMAGE_FILE_ID');
const QR_IMAGE_URL = `https://drive.google.com/uc?export=view&id=${QR_IMAGE_FILE_ID}`;

const REVENUE_SHEET_ID    = PROPS.getProperty('REVENUE_SHEET_ID');

// Reconcile sheets
const SH_HORGA_BILLS   = 'Horga_Bills';
const SH_PAYMENTS_IN   = 'Payments_Inbox';
const SH_REVIEW_QUEUE  = 'Review_Queue';

// TH months map for OCR parsing
const TH_MONTHS = {
  'ม.ค.':1,'มกราคม':1,
  'ก.พ.':2,'กุมภาพันธ์':2,
  'มี.ค.':3,'มีนาคม':3,
  'เม.ย.':4,'เมษายน':4,
  'พ.ค.':5,'พฤษภาคม':5,
  'มิ.ย.':6,'มิถุนายน':6,
  'ก.ค.':7,'กรกฎาคม':7,
  'ส.ค.':8,'สิงหาคม':8,
  'ก.ย.':9,'กันยายน':9,
  'ต.ค.':10,'ตุลาคม':10,
  'พ.ย.':11,'พฤศจิกายน':11,
  'ธ.ค.':12,'ธันวาคม':12
};

const ADMIN_GROUP_ID = PropertiesService.getScriptProperties().getProperty('ADMIN_GROUP_ID');
const FRONTEND_BASE = PROPS.getProperty('FRONTEND_BASE') || 'https://mama-moveout.pages.dev/';

// Check-in picker configuration
const CHECKIN_PICKER_MAX_DATETIME = '2026-01-14T18:00';
const CHECKIN_PICKER_TIMEZONE = 'Asia/Bangkok';
const CHECKIN_PICKER_TZ_OFFSET = '+07:00';
const CHECKIN_PICKER_EARLIEST_MINUTES = 10 * 60;
const CHECKIN_PICKER_EARLIEST_TIME_LABEL = '10:00';
const CHECKIN_PICKER_LATEST_MINUTES = 18 * 60;
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
const OCCUPIED_STATUS_KEYWORDS = ['reserved','occupied','จอง','soon','checked in','check in'];



function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  if (action === 'rooms') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Rooms');
    const vals = sh.getDataRange().getValues();
    const hdr  = vals.shift().map(h => String(h || '').trim());

    const cRoom = colEq_(hdr, 'RoomId');
    const cSt   = colEq_(hdr, 'Status');
    const cMo   = colEq_(hdr, 'Move-out Date');
    const cNext = colEq_(hdr, 'Next Available');

    const rooms = vals.map(r => ({
      roomId:        cRoom > -1 ? String(r[cRoom] || '').toUpperCase().trim() : '',
      status:        cSt   > -1 ? String(r[cSt] || '').toLowerCase().trim()   : '',
      moveoutDate:   cMo   > -1 ? toIso_(r[cMo])                               : '',
      nextAvailable: cNext > -1 ? toIso_(r[cNext])                             : ''
    }));

    return ContentService
      .createTextOutput(JSON.stringify({ rooms }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('Alive');
}

function doPost(e) {
  try {
    if (!e?.postData?.contents) {
      return ContentService.createTextOutput('OK');
    }

const headers = e?.headers || {};
const body    = JSON.parse(e.postData.contents || '{}');

const secretFromHeader = headers['X-Worker-Secret'] || headers['x-worker-secret'] || '';
const secretFromBody   = body.workerSecret || body.secret || '';
const provided         = String(secretFromHeader || secretFromBody || '');

const isEdgeCall = !!body.act;           // our Worker’s direct calls (e.g. moveout)
const isEvents   = Array.isArray(body.events); // Worker forward of LINE events

// before checking secret
console.log('ACT_WEBHOOK', JSON.stringify({
  act: body?.act, bodyKeys: Object.keys(body||{}), hdr: Object.keys(headers||{})
}));

if (!provided || provided !== WORKER_SECRET) {
  console.error('FORBIDDEN_WORKER_SECRET', {
    where: isEdgeCall ? 'edge' : (isEvents ? 'events' : 'unknown'),
    got: provided ? '(present)' : '(empty)'
  });
  return ContentService
    .createTextOutput(JSON.stringify({ ok:false, error:'forbidden' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- routing after passing the single check ----
if (isEdgeCall) {
  if (body.act === 'moveout') {
    const ok = handleMoveoutFromWorker_(body);
    return ContentService
      .createTextOutput(JSON.stringify({ ok }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok:false, error:'unknown_act' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// LINE events forwarded from Worker
if (isEvents) {
  const events = body.events;
  events.forEach(ev => {
    try {
      if (ev.type === 'message') {
        if (ev.message?.type === 'text')  handleTextPush_(ev); // <— new push-only handler
        if (ev.message?.type === 'image') handleImage(ev);     // already push in rent flow
      } else if (ev.type === 'postback') {
        handlePostback(ev); // your postback paths mostly push already
      }
    } catch (inner) {
      console.error('EVENT_ERROR_FWD', inner, JSON.stringify(ev));
    }
  });
  return ContentService.createTextOutput('OK');
}



    // LINE webhook ปกติ
    const events = Array.isArray(body.events) ? body.events : [];
    events.forEach(ev => {
      try {
        if (ev.type === 'message') {
          if (ev.message?.type === 'text')  handleText(ev);
          if (ev.message?.type === 'image') handleImage(ev);
        } else if (ev.type === 'postback') {
          handlePostback(ev);
        }
      } catch (inner) {
        console.error('EVENT_ERROR', inner, JSON.stringify(ev));
      }
    });

    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('MM_LINE: EVENT_HANDLER_ERROR ' + err);
    return ContentService.createTextOutput('OK');
  }
}

/*************** 3) LINE PUSH / LOADING HELPERS **************/
function pushMessage(to, messages) {
  if (!to) return;
  const url = 'https://api.line.me/v2/bot/message/push';
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({ to, messages }),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  console.log('PUSH_STATUS code=' + res.getResponseCode() + ' body=' + res.getContentText());
}

/*************** LINE REPLY HELPER ***************/
function replyMessage(replyToken, messages) {
  if (!replyToken) return;
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({ replyToken, messages }),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  console.log('REPLY_STATUS code=' + res.getResponseCode() + ' body=' + res.getContentText());
}


function pushWithLoading(to, messages, seconds) {
  if (!to) return;
  safeStartLoading(to, seconds || 5);
  try { Utilities.sleep(300); } catch (e) {}
  return pushMessage(to, messages);
}

function safeStartLoading(userId, seconds) {
  try { if (userId) startLoading(userId, seconds || 5); } catch (e) { Logger.log('LOAD_ERR ' + e); }
}

function startLoading(userId, seconds) {
  const secs = Math.max(5, Math.min(seconds || 5, 60));
  const url = 'https://api.line.me/v2/bot/chat/loading/start';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ chatId: userId, loadingSeconds: secs }),
    muteHttpExceptions: true
  });
  console.log('LOADING status=' + res.getResponseCode() + ' ' + res.getContentText());
}


/*************** 4) MESSAGE ROUTING ****************/

function handleText(event) {
  const userId     = event.source?.userId || '';
  const replyToken = event.replyToken;
  const userText   = (event.message?.text || '').trim();

  if (handleCheckinPickerTextCommand_(event)) return;

  // Rent flow intercepts
  if (rentTextGate_(event)) return;

  // Booking code flow
  const m = userText.match(/^#?\s*MM\d{3,}$/i);
  if (!m) return;

  const codeDisplay = m[0].toUpperCase().replace(/\s/g, '');
  const codeKey     = codeDisplay.replace(/^#/, '');
  const row = findBookingRow(codeKey);
  if (!row) {
    return replyMessage(replyToken, [{ type:'text', text:'ไม่พบข้อมูลรหัสนี้' }]);
  }

  return replyMessage(replyToken, [{
    type:'template',
    altText:'ยืนยันการจอง',
    template:{
      type:'buttons',
      text:([
        `รหัส: ${codeDisplay}`,
        `ห้อง: ${row.roomId || '-'}`,
        `ชื่อ: ${row.name}`
      ].join('\n')).slice(0,160),
      actions:[ { type:'postback', label:'ยืนยัน', data:'act=confirm&code='+codeKey } ]
    }
  }]);
}


function rentTextGate_(event) {
  const userId = event.source?.userId || '';
  const txtRaw = (event.message?.text || '').trim();
  if (!userId || !txtRaw) return false;

  const cache = CacheService.getUserCache();
  let flow = {};
  try { flow = JSON.parse(cache.get(userId + ':rent_flow') || '{}'); } catch (e) {}

  if (!flow.step) return false;

  // cancel
  if (/^(ยกเลิก|cancel)$/i.test(txtRaw)) {
    cache.remove(userId + ':rent_flow');
    pushWithLoading(userId, [{ type:'text', text:'ยกเลิกขั้นตอนชำระค่าเช่าแล้วครับ' }]);
    return true;
  }

  if (flow.step === 'await_room') {
    const room = txtRaw.toUpperCase().replace(/\s/g,'');
    if (!/^[A-Z]?\d{3,4}$/.test(room)) {
      pushWithLoading(userId, [{ type:'text', text:'รูปแบบห้องไม่ถูกต้อง ลองอีกครั้ง เช่น A101' }]);
      return true;
    }
    flow.room = room;
    flow.step = 'await_rent_slip';
    cache.put(userId + ':rent_flow', JSON.stringify(flow), 2 * 60 * 60);

    pushWithLoading(userId, [{ type:'text', text:`ห้อง ${room} ครับ\nโปรดส่ง “รูปสลิปค่าเช่า” 1 รูปได้เลย` }]);
    return true;
  }

  return false;
}

function handleRentSlipImage_(event) {
  const userId = event.source?.userId || '';
  const messageId = event.message?.id;
  const cache = CacheService.getUserCache();
  let flow = {};
  try { flow = JSON.parse(cache.get(userId + ':rent_flow') || '{}'); } catch(e) {}

  if (flow.step !== 'await_rent_slip') {
    return pushWithLoading(userId, [{ type:'text', text:'กรุณาพิมพ์เบอร์ห้องก่อน (เช่น A101)' }]);
  }

  safeStartLoading(userId, 5);
  try {
    const blob = fetchLineContentAsBlob_(messageId);
    const type = (blob.getContentType() || '').toLowerCase();
    const size = blob.getBytes().length;
    if (!/^image\/(jpeg|png)$/.test(type)) return pushWithLoading(userId, [{ type:'text', text:'รองรับเฉพาะภาพ jpg/png ครับ' }]);
    if (size > 10 * 1024 * 1024)         return pushWithLoading(userId, [{ type:'text', text:'ไฟล์ใหญ่เกิน 10MB ครับ' }]);

    const ts  = Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH-mm-ss");
    const ext = type === 'image/png' ? 'png' : 'jpg';
    blob.setName(`${flow.room || 'ROOM'}_SLIP_${ts}.${ext}`);

    const tempFileId = DriveApp.getFolderById(TEMP_SLIP_FOLDER_ID).createFile(blob).getId();
    const publicUrl  = `https://drive.google.com/uc?export=view&id=${tempFileId}`;

    let res = { ok:false };
    try {
      res = tryMatchAndConfirm_({
        room:       flow.room,
        slipUrl:    publicUrl,
        lineUserId: userId,
        fileId:     tempFileId     // ← เพิ่มบรรทัดนี้
      }) || { ok:false };
    } catch (e) { Logger.log('tryMatchAndConfirm_ error: ' + e); }

    cache.remove(userId + ':rent_flow');

    if (res.ok) {
      try { moveFileToFolder_(tempFileId, TEMP_SLIP_FOLDER_ID, SLIP_FOLDER_ID); } catch(e) {}
      return pushWithLoading(userId, [{ type:'text', text:'✅ รับสลิปแล้ว ยืนยันการชำระเรียบร้อย ขอบคุณครับ' }]);
    }

    if (res.reason === 'no_open_bill')   return pushWithLoading(userId, [{ type:'text', text:'⏳ ได้รับสลิปแล้ว แต่ยังไม่พบบิลของเดือนนี้ กำลังตรวจสอบ' }]);
    if (res.reason === 'unknown_room')   return pushWithLoading(userId, [{ type:'text', text:'⏳ รับข้อมูลแล้ว ทีมงานจะตรวจสอบห้องอีกครั้ง' }]);
    return pushWithLoading(userId, [{ type:'text', text:'⏳ ได้รับสลิปแล้ว กำลังตรวจสอบโดยเจ้าหน้าที่' }]);

  } catch (e) {
    console.error('RENT_SLIP_SAVE_ERROR ' + e);
    return pushWithLoading(userId, [{ type:'text', text:'บันทึกไฟล์ไม่สำเร็จ โปรดลองใหม่ครับ' }]);
  }
}


/* 4.2) IMAGE FLOW */
function handleImage(event) {
  const userId = event.source?.userId || '';
  const messageId = event.message?.id;
  if (!userId || !messageId) {
    return pushWithLoading(userId, [{ type:'text', text:'รับไฟล์ไม่สำเร็จ ลองอีกครั้งนะคะ' }]);
  }

  // === RENT image gate (run FIRST) ===
  try {
    const userCache = CacheService.getUserCache();
    let rentFlow = {};
    try { rentFlow = JSON.parse(userCache.get(userId + ':rent_flow') || '{}'); } catch(e) {}
    if (rentFlow.step === 'await_rent_slip') {
      return handleRentSlipImage_(event);
    }
  } catch(e) { Logger.log('rent-image-gate ' + e); }

  // existing booking/ID routing (keep)
  const awaiting = CacheService.getUserCache().get(userId + ':awaiting');
  if (awaiting === 'id') return handleIdImage_(event);
  return handleSlipImage_(event);
}

function handleSlipImage_(event) {
  const userId    = event.source?.userId || '';
  const messageId = event.message?.id;

  if (!userId || !messageId) {
    return send_(event, [{ type:'text', text:'รับไฟล์ไม่สำเร็จ ลองอีกครั้งนะคะ' }], 3);
  }

  try {
    const blob = fetchLineContentAsBlob_(messageId);
    const type = (blob.getContentType() || '').toLowerCase();
    const size = blob.getBytes().length;

    // validate type/size
    if (!/^image\/(jpeg|png)$/.test(type)) {
      return send_(event, [{ type:'text', text:'รองรับเฉพาะภาพ jpg/png ค่ะ' }], 0);
    }
    if (size > 10 * 1024 * 1024) {
      return send_(event, [{ type:'text', text:'ไฟล์ใหญ่เกิน 10MB ค่ะ' }], 0);
    }

    // find the booking row awaiting payment from this user
    const found = findAwaitingRowByUser_(userId);
    if (!found) {
      return send_(event, [{ type:'text', text:'ยังไม่พบรายการ กรุณาพิมพ์รหัส #MM### ก่อนส่งสลิปค่ะ' }], 0);
    }

    // save slip image into temp folder
    const codeDisplay = '#' + found.code;
    const tempFileId  = saveSlipToSpecificFolder_(codeDisplay, blob, TEMP_SLIP_FOLDER_ID);
    const publicUrl   = `https://drive.google.com/uc?export=view&id=${tempFileId}`;

    // write to Payments_Inbox immediately
    recordSlipToInbox_({
      lineUserId: userId,
      room: found.roomId || '',
      slipUrl: publicUrl,
      declaredAmount: null,
      note: 'from booking flow (awaiting user confirm)'
    });

    // keep state for confirm step
    const cache = CacheService.getUserCache();
    cache.put(
      userId,
      JSON.stringify({ fileId: tempFileId, code: found.code, rowIndex: found.rowIndex, kind: 'slip' }),
      2 * 60 * 60
    );

    // ask user to confirm the slip (push-friendly)
    return send_(event, [{
      type: 'template',
      altText: 'ยืนยันสลิป',
      template: {
        type: 'confirm',
        text: `ใช้ภาพนี้เป็นสลิปของรหัส ${codeDisplay} ใช่ไหม?`,
        actions: [
          { type: 'postback', label: 'ใช่',   data: 'act=slip_yes' },
          { type: 'postback', label: 'ไม่ใช่', data: 'act=slip_no'  }
        ]
      }
    }], 5); // show loading briefly then push confirm card

  } catch (e) {
    console.error('SLIP_SAVE_ERROR ' + e);
    return send_(event, [{ type:'text', text:'บันทึกไฟล์ไม่สำเร็จ โปรดลองใหม่อีกครั้งค่ะ' }], 0);
  }
}


function handleIdImage_(event) {
  const userId    = event.source?.userId || '';
  const messageId = event.message?.id;

  if (!userId || !messageId) {
    return send_(event, [{ type: 'text', text: 'รับไฟล์ไม่สำเร็จ ลองอีกครั้งนะคะ' }], 0);
  }

  try {
    const blob = fetchLineContentAsBlob_(messageId);
    const type = (blob.getContentType() || '').toLowerCase();
    const size = blob.getBytes().length;

    // validate type/size
    if (!/^image\/(jpeg|png)$/.test(type)) {
      return send_(event, [{ type: 'text', text: 'รองรับเฉพาะภาพ jpg/png ค่ะ' }], 0);
    }
    if (size > 10 * 1024 * 1024) {
      return send_(event, [{ type: 'text', text: 'ไฟล์ใหญ่เกิน 10MB ค่ะ' }], 0);
    }

    // read state saved from slip confirmation step
    const cache = CacheService.getUserCache();
    let slipInfo = {};
    try { slipInfo = JSON.parse(cache.get(userId) || '{}'); } catch (e) {}

    if (!slipInfo.code || !slipInfo.rowIndex) {
      return send_(event, [{ type: 'text', text: 'ยังไม่พบรายการ กรุณาส่งสลิปใหม่ก่อนค่ะ' }], 0);
    }

    // save ID image into temp folder
    const codeDisplay = '#' + slipInfo.code;
    const tempFileId  = saveIdToSpecificFolder_(codeDisplay, blob, TEMP_ID_FOLDER_ID);

    // keep state for confirm step
    cache.put(
      userId,
      JSON.stringify({ fileId: tempFileId, code: slipInfo.code, rowIndex: slipInfo.rowIndex, kind: 'id' }),
      2 * 60 * 60
    );

    // ask user to confirm the ID photo
    return send_(event, [{
      type: 'template',
      altText: 'ยืนยันรูปบัตร',
      template: {
        type: 'confirm',
        text: `ใช้ภาพนี้เป็น "บัตรประชาชนของผู้ที่จะลงนาม" สำหรับรหัส ${codeDisplay} ใช่ไหม?`,
        actions: [
          { type: 'postback', label: 'ใช่',   data: 'act=id_yes' },
          { type: 'postback', label: 'ไม่ใช่', data: 'act=id_no'  }
        ]
      }
    }], 5); // brief loading then push confirm

  } catch (e) {
    console.error('ID_SAVE_ERROR ' + e);
    return send_(event, [{ type: 'text', text: 'บันทึกไฟล์ไม่สำเร็จ โปรดลองใหม่อีกครั้งค่ะ' }], 0);
  }
}



function handlePostback(event) {
  if (handleCheckinPickerPostback_(event)) return;

  const userId = event.source?.userId || '';
  const data   = parseKv(event.postback?.data || '');

  // === RENT CANCEL ===
  if (data.act === 'rent_cancel') {
    if (userId) {
      const cache = CacheService.getUserCache();
      cache.remove(userId + ':rent_flow');
      cache.remove(userId + ':awaiting');
    }
    return; // nothing to say to user here
  }

  // === RENT FLOW START ===
  if (data.act === 'pay_rent') {
    if (!userId) return;
    const cache = CacheService.getUserCache();
    cache.put(userId + ':rent_flow', JSON.stringify({ step: 'await_room' }), 2 * 60 * 60);
    return send_(event, [{ type:'text', text:'พิมพ์เบอร์ห้องของคุณ (เช่น A101)' }], 3);
  }

  // Worker-answered postbacks → ignore locally
  if (data.act && (data.act.startsWith('ROOM_') || data.act.startsWith('FIX_') || data.act.startsWith('RES_'))) {
    return;
  }

  // === GROUP APPROVAL / REJECT ===
  if (data.act === 'mgr_approve' || data.act === 'mgr_reject') {
    const groupId = event.source?.groupId || '';
    const billId  = String(data.bill || '').trim();
    const slipId  = String(data.slip || '').trim();
    const room    = String(data.room || '').trim();

    if (!billId || !slipId) {
      if (groupId) pushMessage(groupId, [{ type:'text', text:'ไม่พบข้อมูลบิล/สลิปในคำสั่งนี้' }]);
      return;
    }

    if (data.act === 'mgr_approve') {
      try {
        setBillSlipReceived_(billId, slipId, 'Slip Received (Group Approved)');
        const inbox = findInboxRowBySlipId_(slipId);
        if (inbox) {
          const sh  = openRevenueSheetByName_('Payments_Inbox');
          const hdr = getHeaders_(sh);
          const cSt = idxOf_(hdr, 'matchstatus');
          const cCf = idxOf_(hdr, 'confidence');
          const cNt = idxOf_(hdr, 'notes');
          if (cSt>-1) sh.getRange(inbox.rowIndex, cSt+1).setValue('approved_group');
          if (cCf>-1) sh.getRange(inbox.rowIndex, cCf+1).setValue(1.0);
          if (cNt>-1) sh.getRange(inbox.rowIndex, cNt+1).setValue('approved by group');
        }

        const MANAGER_USER_ID = PROPS.getProperty('MANAGER_USER_ID') || '';
        if (MANAGER_USER_ID) {
          pushMessage(MANAGER_USER_ID, [{
            type:'text',
            text:`🧾 ยืนยันการจ่ายเงิน\nห้อง: ${room || '-'}\nบิล: ${billId}\nสลิป: ${slipId}\n\nโปรดกดยืนยันในแอพ Horganice`
          }]);
        }

        if (groupId) pushMessage(groupId, [{ type:'text', text:'✅ บันทึกอนุมัติแล้ว แจ้งผู้จัดการให้ยืนยันใน Horganice' }]);
      } catch (e) {
        if (groupId) pushMessage(groupId, [{ type:'text', text:'เกิดข้อผิดพลาดขณะอนุมัติ โปรดลองอีกครั้ง' }]);
      }
      return;
    }

    if (data.act === 'mgr_reject') {
      enqueueReview_({ room, billId, reason: 'group_reject', slipId, note: 'rejected in group' });
      if (groupId) pushMessage(groupId, [{ type:'text', text:'❌ ปฏิเสธแล้ว — ส่งเข้า Review Queue' }]);
      return;
    }
  }

  // === BOOKING CONFIRM ===
  if (data.act === 'confirm' && data.code) {
    const userCache = CacheService.getUserCache();
    userCache.remove(userId + ':awaiting');

    safeStartLoading(userId, 5);
    try {
      const codeKey = String(data.code || '').toUpperCase();
      const info = getBookingByCode_(codeKey);
      if (!info) {
        return send_(event, [{ type: 'text', text: 'ไม่พบข้อมูลรหัสนี้' }], 0);
      }

      const st = (info.status || '').toLowerCase();

      if (st === 'paid' || st === 'จ่ายแล้ว') {
        return send_(event, [{ type: 'text', text: 'รหัสนี้ชำระเงินเรียบร้อยแล้ว ✅' }], 0);
      }

      if (st === 'expired' || st === 'cancelled' || st === 'ยกเลิก') {
        const current = getRoomStatus_(info.roomId);
        if (current === 'avail') {
          try { setRoomStatus_(info.roomId, 'hold'); } catch (e) {}
          const sh = openSheet_();
          const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());
          const cStatus = hdrIdx(headers, ['status','สถานะ']);
          const cConfAt = hdrIdx(headers, ['confirmed at','ยืนยันเมื่อ']);
          const cUserId = hdrIdx(headers, ['line user id','ผู้ใช้ไลน์']);
          if (cStatus > -1) sh.getRange(info.rowIndex, cStatus+1).setValue('Awaiting Payment');
          if (cConfAt > -1) sh.getRange(info.rowIndex, cConfAt+1).setValue(new Date());
          if (cUserId > -1) sh.getRange(info.rowIndex, cUserId+1).setValue(userId);

          return send_(event, [
            { type: 'text', text: 'รหัสนี้เคยหมดอายุ แต่ห้องยังว่างอยู่ → เปิดให้ยืนยันอีกครั้งแล้วค่ะ กรุณาโอนตาม QR และส่งสลิปกลับมาในแชทนี้' },
            { type: 'image', originalContentUrl: QR_IMAGE_URL, previewImageUrl: QR_IMAGE_URL }
          ], 5);
        } else {
          return send_(event, [{ type: 'text', text: 'ขออภัย รหัสนี้หมดอายุและห้องถูกจองแล้วค่ะ กรุณาจองห้องใหม่ผ่านทางเว็บไซต์' }], 0);
        }
      }

      const ok = updateRowOnConfirm(codeKey, userId);
      if (!ok) throw new Error('Row not found for code ' + codeKey);

      const userCache2 = CacheService.getUserCache();
      userCache2.put(userId + ':awaiting', 'slip', 2 * 60 * 60);

      return send_(event, [
        { type: 'text',
          text: '✅ ยืนยันการจองเรียบร้อย\nกรุณาชำระค่าจอง 2000 บาท และส่งสลิปกลับมาที่แชทนี้\nทีมงานจะตรวจสอบและยืนยันให้ภายใน 24 ชม.' },
        { type: 'image', originalContentUrl: QR_IMAGE_URL, previewImageUrl: QR_IMAGE_URL }
      ], 5);

    } catch (e) {
      console.error('CONFIRM_ERROR ' + e);
      return send_(event, [{ type: 'text', text: 'บันทึกไม่สำเร็จ โปรดลองอีกครั้งค่ะ' }], 0);
    }
  }

  // === SLIP CONFIRM YES/NO ===
  if (data.act === 'slip_yes' || data.act === 'slip_no') {
    const userCache = CacheService.getUserCache();
    const cached = userCache.get(userId);
    if (!cached) {
      return send_(event, [{ type: 'text', text: 'หมดเวลายืนยันสลิป กรุณาส่งภาพสลิปใหม่อีกครั้งค่ะ' }], 0);
    }

    const info = JSON.parse(cached); // { fileId, code, rowIndex, kind }
    const fileId   = info.fileId;
    const rowIndex = info.rowIndex;

    if (data.act === 'slip_no') {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      userCache.remove(userId);
      return send_(event, [{ type: 'text', text: 'ลบรูปแล้วค่ะ โปรดส่งภาพสลิปใหม่อีกครั้ง' }], 0);
    }

    safeStartLoading(userId, 5);
    try {
      moveFileToFolder_(fileId, TEMP_SLIP_FOLDER_ID, SLIP_FOLDER_ID);
      const fileUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
      updateSheetOnSlipReceived_(rowIndex, fileUrl);

      // keep booking info and mark awaiting ID
      const idInfo = { fileId, code: info.code, rowIndex, kind: 'id' };
      userCache.put(userId, JSON.stringify(idInfo), 2 * 60 * 60);
      userCache.put(userId + ':awaiting', 'id', 2 * 60 * 60);

      return send_(event, [
        { type: 'text', text: 'ได้รับสลิปแล้วค่ะ ขอบคุณค่ะ 🙏' },
        { type: 'text', text: '📸 ขั้นตอนถัดไป: โปรดส่ง “ภาพบัตรประชาชนของผู้ที่จะลงนามในสัญญาเช่า” มาในแชทนี้นะคะ' }
      ], 3);
    } catch (e) {
      console.error('SLIP_CONFIRM_ERROR ' + e);
      return send_(event, [{ type: 'text', text: 'มีข้อผิดพลาด โปรดลองใหม่อีกครั้งค่ะ' }], 0);
    }
  }

  // === ID CONFIRM YES/NO ===
  if (data.act === 'id_yes' || data.act === 'id_no') {
    const userCache = CacheService.getUserCache();
    const cached = userCache.get(userId);
    if (!cached) {
      return send_(event, [{ type: 'text', text: 'หมดเวลายืนยันรูป โปรดส่งภาพบัตรประชาชนใหม่อีกครั้งค่ะ' }], 0);
    }

    const info = JSON.parse(cached); // { fileId, code, rowIndex, kind:'id' }
    const fileId   = info.fileId;
    const rowIndex = info.rowIndex;

    if (data.act === 'id_no') {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      userCache.remove(userId);
      return send_(event, [{ type: 'text', text: 'ลบรูปแล้วค่ะ โปรดส่งภาพบัตรประชาชนใหม่อีกครั้ง' }], 0);
    }

    safeStartLoading(userId, 5);
    try {
      moveFileToFolder_(fileId, TEMP_ID_FOLDER_ID, ID_FOLDER_ID);
      const fileUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
      updateSheetOnIdReceived_(rowIndex, fileUrl);

      userCache.remove(userId);
      userCache.remove(userId + ':awaiting');

      return send_(event, [
        { type: 'text', text: 'ได้รับภาพบัตรประชาชนเรียบร้อยค่ะ ✅ เจ้าหน้าที่จะตรวจสอบและยืนยันการจองโดยเร็ว' }
      ], 3);
    } catch (e) {
      console.error('ID_CONFIRM_ERROR ' + e);
      return send_(event, [{ type: 'text', text: 'มีข้อผิดพลาด โปรดลองใหม่อีกครั้งค่ะ' }], 0);
    }
  }
}

function handleCheckinPickerPostback_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'postback') return false;
  const payload = event.postback || {};
  const data = _parsePostbackData_(payload.data || '');
  if (String(data.act || '').trim().toLowerCase() !== 'checkin_pick') return false;

  const params = payload.params || {};
  const datetimeRaw = _lineDatetimeFromParams_(params);
  const source = event.source || {};
  const userId = String(source.userId || '').trim();
  const roomId = String(data.room || '').trim() || (userId ? _findRoomByUserId_(userId) : '');
  const pickerMax = CHECKIN_PICKER_MAX_DATETIME ? _parseLineDatetimeValue_(CHECKIN_PICKER_MAX_DATETIME) : null;
  const maxThaiDate = pickerMax ? _thaiDate_(pickerMax) : '';

  const pushUserText = (txt) => {
    if (userId && txt) {
      pushMessage(userId, [{ type: 'text', text: txt }]);
    }
  };

  if (!datetimeRaw) {
    pushUserText('ระบบไม่พบวัน–เวลาเช็คอินที่เลือก กรุณากดเลือกใหม่อีกครั้งค่ะ 🙏');
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  const clockMinutes = _clockMinutesFromLineDatetime_(datetimeRaw);
  const chosenTimeText = (datetimeRaw.split('T')[1] || '').slice(0, 5);
  if (!Number.isFinite(clockMinutes)) {
    pushUserText('รูปแบบเวลาไม่ถูกต้อง กรุณาเลือกใหม่อีกครั้งค่ะ 🙏');
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  if (clockMinutes < CHECKIN_PICKER_EARLIEST_MINUTES) {
    pushUserText(
      `เวลาที่เลือก (${chosenTimeText || 'ไม่ระบุ'}) ก่อน ${CHECKIN_PICKER_EARLIEST_TIME_LABEL} น.\n` +
      `กรุณาเลือกช่วง ${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้นค่ะ 🙏`
    );
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  if (clockMinutes > CHECKIN_PICKER_LATEST_MINUTES) {
    pushUserText(
      `เวลาที่เลือก (${chosenTimeText || 'ไม่ระบุ'}) หลัง ${CHECKIN_PICKER_LATEST_TIME_LABEL} น.\n` +
      `กรุณาเลือกช่วง ${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้นค่ะ 🙏`
    );
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  const selected = _parseLineDatetimeValue_(datetimeRaw);
  if (!selected) {
    pushUserText('ไม่สามารถอ่านค่าวันที่/เวลาได้ กรุณาเลือกใหม่ค่ะ 🙏');
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  if (pickerMax && selected.getTime() > pickerMax.getTime()) {
    pushUserText(
      `วันเช็คอินที่เลือกเกินช่วงที่กำหนดไว้ กรุณาเลือกวันก่อน ${maxThaiDate || '15 ม.ค. 2026'} ค่ะ 🙏`
    );
    if (userId && roomId) sendCheckinPickerToUser(userId, roomId);
    return true;
  }

  if (!roomId) {
    pushUserText('ระบบไม่พบห้องที่เชื่อมกับบัญชี LINE นี้ กรุณาติดต่อแอดมินเพื่อให้ช่วยบันทึกวันเช็คอินค่ะ 🙏');
    return true;
  }

  const dateOnly = new Date(selected.getFullYear(), selected.getMonth(), selected.getDate());
  const timeText = Utilities.formatDate(selected, CHECKIN_PICKER_TIMEZONE, 'HH:mm');
  const saved = _updateRoomCheckinSelection_(roomId, { dateOnly, timeText });

  if (!saved) {
    pushUserText('ระบบบันทึกวันเช็คอินไม่สำเร็จ กรุณาติดต่อแอดมินเพื่อให้ช่วยตรวจสอบค่ะ 🙏');
    console.log('Check-in picker: failed to write to sheet for room ' + roomId);
    return true;
  }

  const thaiDate = _thaiDate_(dateOnly);
  const ackLines = [
    `บันทึกวันเช็คอินของห้อง ${roomId} แล้วค่ะ 🙏`,
    `🗓️ ${thaiDate} เวลา ${timeText} น.`,
    'หากต้องการปรับเปลี่ยน สามารถกดเลือกใหม่จากปุ่มเดิมได้เลยค่ะ'
  ];
  pushUserText(ackLines.join('\n'));
  console.log(`Check-in picker saved for ${roomId}: ${thaiDate} ${timeText}`);
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
  const cAt   = H.indexOf('ConfirmedAt') + 1;
  if (!cDate && !cTime && !cConf && !cAt) return false;

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
    if (cAt) sh.getRange(row, cAt).setValue(new Date());
    return true;
  }
  return false;
}

function _thaiDate_(date) {
  if (!(date instanceof Date)) return '';
  return Utilities.formatDate(date, CHECKIN_PICKER_TIMEZONE, 'dd MMM yyyy');
}

function handleCheckinPickerTextCommand_(event) {
  if (!event || String(event.type || '').toLowerCase() !== 'message') return false;
  const message = event.message || {};
  if (String(message.type || '').toLowerCase() !== 'text') return false;

  const rawText = String(message.text || '').trim();
  if (!rawText) return false;

  const collapsed = rawText.toLowerCase().replace(/\s+/g, '');
  const matched = CHECKIN_PICKER_COMMAND_KEYWORDS.some(keyword => collapsed.includes(keyword));
  if (!matched) return false;

  const source = event.source || {};
  const userId = String(source.userId || '').trim();
  const replyTargetId = String(source.userId || source.groupId || source.roomId || '').trim();
  if (!replyTargetId) return false;

  if (!userId) {
    pushMessage(replyTargetId, [{
      type: 'text',
      text: 'ระบบต้องการข้อมูลบัญชี LINE เพื่อส่งปุ่มเลือกวันเช็คอิน กรุณาเริ่มแชตกับบอทในห้องส่วนตัวก่อนนะคะ 🙏'
    }]);
    return true;
  }

  const roomId = _findRoomByUserId_(userId);
  if (!roomId) {
    pushMessage(replyTargetId, [{
      type: 'text',
      text: 'ระบบไม่พบห้องที่เชื่อมกับบัญชี LINE นี้ กรุณาติดต่อแอดมินเพื่อให้ช่วยตรวจสอบนะคะ 🙏'
    }]);
    return true;
  }

  pushMessage(replyTargetId, [{
    type: 'text',
    text: `ส่งปุ่มเลือกวัน–เวลาเช็คอินของห้อง ${roomId} ให้แล้วค่ะ 🙏\nกดเลือกจากปุ่มล่าสุดได้เลย`
  }]);
  sendCheckinPickerToUser(userId, roomId);
  return true;
}


/*************** 5) SHEET HELPERS (booking, rooms, hdr) **************/
function openSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);
  return sh;
}

// Rooms helpers
function _normStatus_(s) {
  s = String(s || '').toLowerCase().trim();
  if (['reserved','จองแล้ว'].includes(s)) return 'reserved';
  if (['hold','ถือห้อง','ถือจอง'].includes(s)) return 'hold';
  return 'avail';
}

function getRoomStatus_(roomId) {
  if (!roomId) return '';
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) return '';
  const values = sh.getDataRange().getValues();
  const header = values.shift().map(h => String(h||'').trim());
  const iId  = header.findIndex(h => h.toLowerCase().includes('room'));
  const iSt  = header.findIndex(h => h.toLowerCase().includes('status') || h.includes('สถานะ'));
  for (let r = 0; r < values.length; r++) {
    const id = String(values[r][iId] || '').trim();
    if (id === roomId) {
      return _normStatus_(values[r][iSt] || '');
    }
  }
  return '';
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

// Booking helpers
function getBookingByCode_(codeKey) {
  const sh = openSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => (h || '').toString().trim());
  const idx = {
    code:        hdrIdx(headers, ['รหัสการจอง','code','booking']),
    name:        hdrIdx(headers, ['ชื่อ-นามสกุล','ชื่อ','name']),
    room:        hdrIdx(headers, ['room','room id','ห้อง']),
    status:      hdrIdx(headers, ['status','สถานะ']),
    confirmedAt: hdrIdx(headers, ['confirmed at','ยืนยันเมื่อ']),
    userId:      hdrIdx(headers, ['line user id','ผู้ใช้ไลน์']),
  };
  if (idx.code < 0) idx.code = 0;

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowCode = (row[idx.code] || '').toString().trim().toUpperCase().replace(/^#/, '');
    if (rowCode === codeKey) {
      return {
        rowIndex:    i + 2,
        code:        rowCode,
        name:        (row[idx.name] || '').toString().trim(),
        roomId:      (idx.room > -1 ? (row[idx.room] || '').toString().trim() : ''),
        status:      (idx.status > -1 ? (row[idx.status] || '').toString().trim() : ''),
        confirmedAt: (idx.confirmedAt > -1 && (row[idx.confirmedAt] instanceof Date)) ? row[idx.confirmedAt] : null,
        userId:      (idx.userId > -1 ? (row[idx.userId] || '').toString().trim() : '')
      };
    }
  }
  return null;
}

function findBookingRow(codeKey) {
  const sh = openSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().trim());
  const idx = {
    code: hdrIdx(headers, ['รหัสการจอง','code','booking']),
    name: hdrIdx(headers, ['ชื่อ-นามสกุล','ชื่อ','name']),
    room: hdrIdx(headers, ['room','room id','ห้อง'])
  };
  if (idx.code < 0) idx.code = 0;

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowCode = (row[idx.code] || '').toString().trim().toUpperCase().replace(/^#/, '');
    if (rowCode === codeKey) {
      return {
        rowIndex: i + 2,
        code: rowCode,
        name: (row[idx.name] || '').toString().trim(),
        roomId: (idx.room > -1 ? (row[idx.room] || '').toString().trim() : '')
      };
    }
  }
  return null;
}

function updateRowOnConfirm(codeKey, userId) {
  const sh = openSheet_();
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return false;

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => (h||'').toString().trim());
  const idx = {
    code:        hdrIdx(headers, ['รหัสการจอง','code','booking']),
    status:      hdrIdx(headers, ['status','สถานะ']),
    confirmedAt: hdrIdx(headers, ['confirmed at','ยืนยันเมื่อ']),
    userId:      hdrIdx(headers, ['line user id','ผู้ใช้ไลน์'])
  };
  if (idx.code < 0) idx.code = 0;

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (let r = 0; r < values.length; r++) {
    const rowCode = (values[r][idx.code] || '').toString().trim().toUpperCase().replace(/^#/, '');
    if (rowCode === codeKey) {
      if (idx.status      > -1) sh.getRange(r+2, idx.status+1).setValue('Awaiting Payment');
      if (idx.confirmedAt > -1) sh.getRange(r+2, idx.confirmedAt+1).setValue(new Date());
      if (idx.userId      > -1) sh.getRange(r+2, idx.userId+1).setValue(userId);
      return true;
    }
  }
  return false;
}

function findAwaitingRowByUser_(userId) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  const colCode   = (headers.findIndex(h => h.toLowerCase().includes('รหัสการจอง')) + 1) || 1;
  const colStatus = headers.findIndex(h => h.toLowerCase().includes('status')) + 1;
  const colUserId = headers.findIndex(h => h.trim().toLowerCase() === 'line user id') + 1;
  const colRoom   = headers.findIndex(h => h.toLowerCase().includes('room')) + 1;

  if (!colStatus || !colUserId) return null;

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (let r = values.length - 1; r >= 0; r--) {
    if (String(values[r][colUserId-1]).trim() === userId &&
        String(values[r][colStatus-1]).trim() === 'Awaiting Payment') {
      const codeKey = String(values[r][colCode-1]).trim().toUpperCase().replace(/^#/, '');
      const roomId  = colRoom ? String(values[r][colRoom-1] || '').trim().toUpperCase() : '';
      return { rowIndex: r + 2, code: codeKey, roomId };
    }
  }
  return null;
}

function updateSheetOnSlipReceived_(rowNum, fileUrl) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  const colStatus = headers.findIndex(h => h.toLowerCase().includes('status')) + 1;
  const colSlipAt = headers.findIndex(h => h.toLowerCase().includes('slip received at')) + 1;
  const colLink   = headers.findIndex(h => h.toLowerCase().includes('slip link')) + 1;

  if (colStatus) sh.getRange(rowNum, colStatus).setValue('Slip Received');
  if (colSlipAt) sh.getRange(rowNum, colSlipAt).setValue(new Date());
  if (colLink)   sh.getRange(rowNum, colLink).setValue(fileUrl);
}

function updateSheetOnIdReceived_(rowNum, fileUrl) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  const colIdAt = headers.findIndex(h => h.toLowerCase().includes('id received at')) + 1;
  const colIdLn = headers.findIndex(h => h.toLowerCase().includes('id link')) + 1;
  const colKyc  = headers.findIndex(h => h.toLowerCase().includes('kyc status')) + 1;

  if (colIdAt) sh.getRange(rowNum, colIdAt).setValue(new Date());
  if (colIdLn) sh.getRange(rowNum, colIdLn).setValue(fileUrl);
  if (colKyc)  sh.getRange(rowNum, colKyc).setValue('id_provided'); // optional marker
}

function hdrIdx(headers, keys) {
  const lower = headers.map(h => h.toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    for (const k of keys) if (lower[i].indexOf(k.toLowerCase()) !== -1) return i;
  }
  return -1;
}

function colEq_(headers, name) {
  const want = String(name).trim().toLowerCase();
  return headers.findIndex(h => String(h || '').trim().toLowerCase() === want);
}

function toIso_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}



/*************** 6) LINE & DRIVE HELPERS ***************/
function fetchLineContentAsBlob_(messageId) {
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('LINE_MEDIA_FETCH ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  return res.getBlob();
}

function saveSlipToSpecificFolder_(codeDisplay, blob, folderId) {
  const ts  = Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH-mm-ss");
  const safe= codeDisplay.replace('#', '');
  const ext = blob.getContentType() === 'image/png' ? 'png' : 'jpg';
  blob.setName(`${safe}_SLIP_${ts}.${ext}`);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  return file.getId();
}

function saveIdToSpecificFolder_(codeDisplay, blob, folderId) {
  const ts  = Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH-mm-ss");
  const safe= codeDisplay.replace('#', '');
  const ext = blob.getContentType() === 'image/png' ? 'png' : 'jpg';
  blob.setName(`${safe}_ID_${ts}.${ext}`);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  return file.getId();
}

function moveFileToFolder_(fileId, srcFolderId, destFolderId) {
  if (!fileId || !destFolderId) throw new Error('moveFileToFolder_: missing ids');
  if (srcFolderId === destFolderId) { Logger.log('move: src==dest, skip'); return; }

  try {
    // Preferred path (Shared Drives safe)
    Drive.Files.update(
      {},                              // metadata not changing
      fileId,
      null,                            // media
      {
        addParents: destFolderId,
        removeParents: srcFolderId || '',
        supportsAllDrives: true
      }
    );
    Logger.log('move: Drive.Files.update OK file=%s to=%s (from=%s)', fileId, destFolderId, srcFolderId || '(none)');
  } catch (e) {
    Logger.log('move: Drive.Files.update failed: ' + e);
    // Fallback for My Drive legacy behavior
    try {
      const file = DriveApp.getFileById(fileId);
      DriveApp.getFolderById(destFolderId).addFile(file);
      if (srcFolderId) {
        try { DriveApp.getFolderById(srcFolderId).removeFile(file); } catch (e2) { Logger.log('move: removeFile warn: ' + e2); }
      }
      Logger.log('move: fallback add/remove OK for file=%s', fileId);
    } catch (e3) {
      Logger.log('move: fallback failed: ' + e3);
      throw e3;
    }
  }
}

function parseKv(q) {
  const out = {};
  (q || '').split('&').forEach(p => {
    if (!p) return;
    const eq = p.indexOf('=');
    if (eq === -1) { out[decodeURIComponent(p)] = ''; return; }
    const k = decodeURIComponent(p.slice(0, eq));
    const v = decodeURIComponent(p.slice(eq + 1));
    out[k] = v;
  });
  return out;
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
  input.split('&').forEach(fragment => {
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

function _buildPostbackData_(data) {
  if (!data || typeof data !== 'object') return '';
  return Object.keys(data)
    .filter(key => Object.prototype.hasOwnProperty.call(data, key))
    .map(key => {
      const val = data[key];
      if (val === undefined || val === null || val === '') return null;
      return encodeURIComponent(key) + '=' + encodeURIComponent(String(val));
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

function _tokensSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tokens');
}

function issueToken_(lineId, roomId, minutes, purpose) {
  const sh = _tokensSheet_();
  if (!sh) throw new Error('Tokens sheet not found');
  const tok = Utilities.getUuid().replace(/-/g,'');
  const now = new Date();
  const exp = new Date(now.getTime() + (Number(minutes) || 20)*60*1000);
  sh.appendRow([tok, lineId || '', roomId || '', now, exp, '', purpose || 'moveout']);
  return tok;
}

function findRoomByLineId_(lineId) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) return '';
  const vals = sh.getDataRange().getValues();
  const head = vals.shift().map(h=>String(h||'').trim().toLowerCase());
  const cRoom = head.findIndex(h=>h.includes('room'));
  const cUser = head.findIndex(h=>h.includes('line') && h.includes('id'));
  if (cRoom<0 || cUser<0) return '';
  for (const r of vals) if (String(r[cUser]||'').trim() === lineId) return String(r[cRoom]||'').trim().toUpperCase();
  return '';
}

function _findRoomByUserId_(userId) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  if (!sh) return '';
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim());
  const lower = headers.map(h => h.toLowerCase());
  const cRoom = lower.findIndex(h => h.includes('room')) + 1;
  const cUser = lower.findIndex(h => h.includes('line id')) + 1;
  if (!cRoom || !cUser) return '';

  const rows = sh.getLastRow() - 1;
  if (rows < 1) return '';

  const values = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();
  const target = String(userId || '').trim().toLowerCase();
  for (let i = 0; i < values.length; i++) {
    const id = String(values[i][cUser - 1] || '').trim().toLowerCase();
    if (id && id === target) {
      return String(values[i][cRoom - 1] || '').trim();
    }
  }
  return '';
}

function _roomsHeaders_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rooms');
  const H  = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  return { sh, H, Hl: H.map(h => h.toLowerCase()) };
}


/*************** 7) OCR (Vision API) + THAI PARSERS ***************/
function ocrSlipFromFileId_(fileId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GCV_API_KEY');
  if (!apiKey) throw new Error('Missing GCV_API_KEY');

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const b64  = Utilities.base64Encode(blob.getBytes());

  const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(apiKey);
  const payload = {
    requests: [{
      image: { content: b64 },
      features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
      imageContext: {
        // OCR ไทย + อังกฤษ
        languageHints: ['th', 'en']
      }
    }]
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code >= 300) throw new Error('Vision error ' + code + ' ' + res.getContentText());

  const data = JSON.parse(res.getContentText());
  const text = data.responses?.[0]?.fullTextAnnotation?.text || '';
  return text;
}

function thaiYearToCE_(y) {
  y = Number(y);
  if (y > 2400) return y - 543;
  return y;
}

function parseThaiSlip_(rawText) {
  const text = String(rawText || '').replace(/\u200B/g,'').replace(/\s+/g,' ').trim();

  // 1) Amount (THB)
  let amount = null;
  const amtPatterns = [
    /(?:ยอด(?:เงิน)?|จำนวนเงิน|amount)\s*[:\-]?\s*([0-9,.]+)\s*(?:บาท|thb)?/i,
    /(?:thb|บาท)\s*([0-9,.]+)/i,
    /\b([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{1,2})?)\s*(?:บาท|thb)\b/i
  ];
  for (const re of amtPatterns) {
    const m = text.match(re);
    if (m) { amount = Number(String(m[1]).replace(/,/g,'')); break; }
  }

  // 2) Date & Time
  let txDate = null, txTime = null;

  // dd/mm/yyyy (ไทย/สากล)
  let m = text.match(/\b([0-3]?\d)[\/\-]([01]?\d)[\/\-](\d{4})\b/);
  if (m) {
    const d = Number(m[1]), mo = Number(m[2]), y = thaiYearToCE_(m[3]);
    try { txDate = new Date(y, mo-1, d); } catch(e){}
  }

  // 1 ก.ย. 2568 / 1 กันยายน 2568
  if (!txDate) {
    const m2 = text.match(/\b([0-3]?\d)\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.|มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)\s*(\d{4})\b/);
    if (m2) {
      const d = Number(m2[1]); const monthName = m2[2]; const y = thaiYearToCE_(m2[3]);
      const mo = TH_MONTHS[monthName] || null;
      if (mo) { try { txDate = new Date(y, mo-1, d); } catch(e){} }
    }
  }

  // Time hh:mm
  const t = text.match(/\b([01]?\d|2[0-3])[:\.]([0-5]\d)\b/);
  if (t) {
    txTime = `${t[1].padStart(2,'0')}:${t[2]}`;
  }

  // 3) Reference / Txn ID
  let txId = null;
  const refPatterns = [
    /(หมายเลขอ้างอิง|เลขที่อ้างอิง|reference|ref\.?|transaction id|trace id)\s*[:\-]?\s*([A-Za-z0-9\-]{6,})/i,
    /\bFT[A-Z0-9\-]{6,}\b/i
  ];
  for (const re of refPatterns) {
    const mm = text.match(re);
    if (mm) { txId = (mm[2] || mm[0]).replace(/^(หมายเลขอ้างอิง|เลขที่อ้างอิง|reference|ref\.?|transaction id|trace id)\s*[:\-]?\s*/i,'').trim(); break; }
  }

  // 4) Bank detection (หยาบ ๆ)
  let bank = null;
  if (/ไทยพาณิชย์|SCB/i.test(text)) bank = 'SCB';
  else if (/กสิกร|KBank/i.test(text)) bank = 'KBank';
  else if (/กรุงไทย|Krungthai|KTB/i.test(text)) bank = 'KTB';
  else if (/กรุงเทพ|Bangkok Bank|BBL/i.test(text)) bank = 'BBL';
  else if (/กรุงศรี|Krungsri|BAY/i.test(text)) bank = 'BAY';
  else if (/พร้อมเพย์|PromptPay/i.test(text)) bank = 'PromptPay';

  return {
    rawText: text,
    amount: (amount != null && !isNaN(amount)) ? amount : null,
    txDate: txDate || null,
    txTime: txTime || null,
    txId: txId || null,
    bank: bank || null
  };
}


/*************** 8) RECONCILIATION (Payments/Bills/Review) ***************/
function monthKey_(dateObj) {
  const d = dateObj || new Date();
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
}

function openSheetByName_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
function getHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  return sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
}
function idxOf_(headers, keyLike) {
  const lower = headers.map(h => h.toLowerCase());
  const key   = keyLike.toLowerCase();
  for (let i = 0; i < lower.length; i++) if (lower[i].indexOf(key) !== -1) return i;
  return -1;
}
function genSlipId_() {
  const ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyyMMdd-HHmmss");
  const rnd = Math.floor(Math.random()*9000)+1000;
  return `SLIP-${ts}-${rnd}`;
}

function recordSlipToInbox_({lineUserId, room, slipUrl, declaredAmount, note}) {
  const sh  = openRevenueSheetByName_('Payments_Inbox');  // << เปลี่ยนตรงนี้
  const hdr = getHeaders_(sh);

  const cSlipID   = idxOf_(hdr, 'slipid');
  const cRecvAt   = idxOf_(hdr, 'received');
  const cUser     = idxOf_(hdr, 'lineuserid');
  const cRoom     = idxOf_(hdr, 'room');
  const cAmtDecl  = idxOf_(hdr, 'amountdecl');
  const cUrl      = idxOf_(hdr, 'slipurl');
  const cMatchSt  = idxOf_(hdr, 'matchstatus');
  const cNotes    = idxOf_(hdr, 'notes');

  const slipId = genSlipId_();
  const row    = new Array(hdr.length).fill('');

  if (cSlipID  > -1) row[cSlipID]  = slipId;
  if (cRecvAt  > -1) row[cRecvAt]  = new Date();
  if (cUser    > -1) row[cUser]    = lineUserId || '';
  if (cRoom    > -1) row[cRoom]    = (room || '').toUpperCase();
  if (cAmtDecl > -1) row[cAmtDecl] = (declaredAmount != null ? Number(declaredAmount) : '');
  if (cUrl     > -1) row[cUrl]     = slipUrl || '';
  if (cMatchSt > -1) row[cMatchSt] = 'pending';
  if (cNotes   > -1) row[cNotes]   = note || '';

  const nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, hdr.length).setValues([row]);
  return { slipId, rowIndex: nextRow };
}

function findCandidateBill_({ room, declaredAmount }) {
  if (!room) return null;
  const sh  = openRevenueSheetByName_('Horga_Bills');     // << เปลี่ยนตรงนี้
  const hdr = getHeaders_(sh);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const cBill   = idxOf_(hdr, 'billid');
  const cRoom   = idxOf_(hdr, 'room');
  const cMonth  = idxOf_(hdr, 'month');       // คาดว่าเป็น yyyy-MM
  const cAmtDue = idxOf_(hdr, 'amountdue');
  const cStatus = idxOf_(hdr, 'status');
  const cSlipID = idxOf_(hdr, 'slipid');

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const nowKey = monthKey_(new Date());
  const wantRoom = String(room||'').trim().toUpperCase();

  const candidates = [];
  for (let i=0;i<values.length;i++) {
    const row = values[i];
    const rRoom   = String(row[cRoom] || '').trim().toUpperCase();
    const rStatus = String(row[cStatus] || '').trim().toLowerCase();
    const rMonth  = String(row[cMonth] || '').trim();
    const rSlip   = String(row[cSlipID] || '').trim();
    if (rRoom !== wantRoom) continue;

    // unpaid = ไม่มี slip และ status ไม่ใช่ 'paid'
    const unpaid = !(rStatus === 'paid' || rStatus === 'จ่ายแล้ว' || rSlip);
    if (!unpaid) continue;

    const amtDue = Number(row[cAmtDue] || 0);
    let score = (rMonth === nowKey ? 2 : 1);
    if (declaredAmount != null && !isNaN(declaredAmount)) {
      if (Math.abs(amtDue - Number(declaredAmount)) < 0.5) score += 2; // bonus เมื่อยอดตรง
    }
    candidates.push({ rowIndex: i+2, billId: String(row[cBill]||'').trim(), month: rMonth, amountDue: amtDue, score });
  }

  if (candidates.length === 0) return null;
  candidates.sort((a,b) => b.score - a.score);
  const topScore = candidates[0].score;
  const topGroup = candidates.filter(x => x.score === topScore);
  if (topGroup.length > 1) return { ambiguous: true, candidates: topGroup };
  return { candidate: candidates[0] };
}

function updateBillWithSlip_({ rowIndex, slipId, markStatus }) {
  const sh  = openRevenueSheetByName_('Horga_Bills');     // << เปลี่ยนตรงนี้
  const hdr = getHeaders_(sh);
  const cStatus = idxOf_(hdr, 'status');
  const cPaidAt = idxOf_(hdr, 'paidat');
  const cSlipID = idxOf_(hdr, 'slipid');

  if (cSlipID > -1) sh.getRange(rowIndex, cSlipID+1).setValue(slipId);
  if (cPaidAt > -1) sh.getRange(rowIndex, cPaidAt+1).setValue(new Date());
  if (cStatus > -1) sh.getRange(rowIndex, cStatus+1).setValue(markStatus || 'Slip Received');
}

function updateInboxMatchResult_({ rowIndex, status, matchedBillId, confidence, note }) {
  const sh  = openRevenueSheetByName_('Payments_Inbox');  // << เปลี่ยนตรงนี้
  const hdr = getHeaders_(sh);
  const cMatchSt = idxOf_(hdr, 'matchstatus');
  const cMatchId = idxOf_(hdr, 'matchedbillid');
  const cConf    = idxOf_(hdr, 'confidence');
  const cNotes   = idxOf_(hdr, 'notes');

  if (cMatchSt > -1) sh.getRange(rowIndex, cMatchSt+1).setValue(status || '');
  if (cMatchId > -1) sh.getRange(rowIndex, cMatchId+1).setValue(matchedBillId || '');
  if (cConf    > -1) sh.getRange(rowIndex, cConf+1).setValue(confidence != null ? Number(confidence) : '');
  if (cNotes   > -1 && note) sh.getRange(rowIndex, cNotes+1).setValue(note);
}

function enqueueReview_({ room, billId, declaredAmount, amountDue, reason, slipId, note, lineUserId }) {
  const sh  = openRevenueSheetByName_('Review_Queue');
  const hdr = getHeaders_(sh);
  const row = new Array(hdr.length).fill('');

  const cCreated = idxOf_(hdr, 'createdat');
  const cRoom    = idxOf_(hdr, 'room');
  const cBill    = idxOf_(hdr, 'billid');
  const cAmtDec  = idxOf_(hdr, 'amountdecl');
  const cAmtDue  = idxOf_(hdr, 'amountdue');
  const cReason  = idxOf_(hdr, 'reason');
  const cSlipID  = idxOf_(hdr, 'slipid');
  const cNotes   = idxOf_(hdr, 'notes');
  const cUser    = idxOf_(hdr, 'lineuserid'); // NEW

  if (cCreated > -1) row[cCreated] = new Date();
  if (cRoom    > -1) row[cRoom]    = (room || '').toUpperCase();
  if (cBill    > -1) row[cBill]    = billId || '';
  if (cAmtDec  > -1) row[cAmtDec]  = (declaredAmount != null ? Number(declaredAmount) : '');
  if (cAmtDue  > -1) row[cAmtDue]  = (amountDue != null ? Number(amountDue) : '');
  if (cReason  > -1) row[cReason]  = reason || '';
  if (cSlipID  > -1) row[cSlipID]  = slipId || '';
  if (cNotes   > -1 && note) row[cNotes] = note;
  if (cUser    > -1) row[cUser]    = lineUserId || '';

  const nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, hdr.length).setValues([row]);
}

function tryMatchAndConfirm_(args) {
  const room       = (args.room || '').toUpperCase().trim();
  const slipUrl    = args.slipUrl || '';
  const lineUserId = args.lineUserId || '';
  const fileId     = args.fileId || '';   // สำคัญสำหรับ OCR

  // 1) เขียนเข้ากล่องรับเสมอ
  const inbox = recordSlipToInbox_({
    lineUserId, room, slipUrl, declaredAmount: null,
    note: 'auto-created by LINE bot'
  });

  // 2) OCR จากไฟล์ (ถ้ามี fileId)
  let ocr = null;
  let ocrOk = false;
  if (fileId) {
    try {
      const text = ocrSlipFromFileId_(fileId);
      ocr = parseThaiSlip_(text);
      ocrOk = !!(ocr && (ocr.amount != null || ocr.txDate || ocr.txId || ocr.bank));
      updateInboxMatchResult_({
        rowIndex: inbox.rowIndex,
        status: 'pending_ocr',
        matchedBillId: '',
        confidence: '',
        note: `OCR: amount=${ocr?.amount ?? ''}, date=${ocr?.txDate ? Utilities.formatDate(ocr.txDate,'Asia/Bangkok','yyyy-MM-dd') : ''}, bank=${ocr?.bank ?? ''}, ref=${ocr?.txId ?? ''}`
      });
      Logger.log('OCR_OK for inbox row %s', inbox.rowIndex);
    } catch (e) {
      Logger.log('OCR_FAIL: ' + e);
      updateInboxMatchResult_({
        rowIndex: inbox.rowIndex,
        status: 'ocr_error',
        matchedBillId: '',
        confidence: '',
        note: 'OCR_FAIL: ' + e
      });
    }
  } else {
    // ไม่มี fileId → ไม่มีทาง OCR ได้
    updateInboxMatchResult_({
      rowIndex: inbox.rowIndex,
      status: 'no_ocr_source',
      matchedBillId: '',
      confidence: '',
      note: 'No fileId → skip OCR'
    });
  }

  // 3) หา Bill ผู้ท้าชิง (ถ้า OCR ได้ amount จะช่วยเพิ่มความแม่นยำ)
  const declaredAmount = (ocr && ocr.amount != null) ? Number(ocr.amount) : null;
  const found = findCandidateBill_({ room, declaredAmount });

  if (!found) {
    updateInboxMatchResult_({
      rowIndex: inbox.rowIndex,
      status: 'no_open_bill',
      matchedBillId: '',
      confidence: 0.0,
      note: (ocrOk ? 'no_open_bill (have OCR)' : 'no_open_bill (no/failed OCR)')
    });
    enqueueReview_({ room, declaredAmount, reason: 'no_open_bill', slipId: inbox.slipId, note: 'auto-queued',
    lineUserId });
    return { ok:false, reason:'no_open_bill', slipId: inbox.slipId };
  }

  if (found.ambiguous) {
    updateInboxMatchResult_({
      rowIndex: inbox.rowIndex,
      status: 'ambiguous',
      matchedBillId: '',
      confidence: 0.3,
      note: 'multiple candidates' + (ocrOk ? ' (have OCR)' : ' (no/failed OCR)')
    });
    enqueueReview_({ room, declaredAmount, reason: 'ambiguous_candidates', slipId: inbox.slipId,
    lineUserId });
    return { ok:false, reason:'ambiguous' };
  }

  // 4) จับคู่ได้ → ตรวจ amount ก่อน จากนั้นค่อยอัปเดตบิล + ให้คะแนน
  const cand = found.candidate;

  const billAmt  = Number(cand.amountDue);
  let   conf     = 0.70;   // baseline
  let   amtDelta = null;

  if (ocrOk && ocr.amount != null) {
    amtDelta = Math.abs(Number(ocr.amount) - billAmt);

    // STRICT rule (tune the 3 THB threshold as you like)
    if (amtDelta > 3) {
      updateInboxMatchResult_({
        rowIndex: inbox.rowIndex,
        status: 'amount_mismatch',
        matchedBillId: cand.billId,
        confidence: 0.4,
        note: `OCR amount=${ocr.amount}; bill=${billAmt}; Δ=${amtDelta}`
      });

      enqueueReview_({
        room, billId: cand.billId,
        declaredAmount: ocr.amount,
        amountDue: billAmt,
        reason: 'amount_mismatch',
        slipId: inbox.slipId,
        note: 'blocked auto-match due to amount mismatch',
        lineUserId // 👈 if you want to notify tenant later
      });

      // notify tenant + admin here (optional)
      notifyTenantMismatch_(lineUserId, {
        room,
        billId: cand.billId,
        amountDue: billAmt,
        ocrAmount: ocr.amount,
        delta: amtDelta,
        slipId: inbox.slipId
      });

      // 👇 NEW: notify admin LINE group
      adminNotify_(
        '⚠️ ค่าเช่าไม่ตรงกับบิล\n' +
        `🏠 ห้อง: ${room}\n` +
        `🧾 บิล: ${cand.billId}\n` +
        `💵 ยอดบิล: ${billAmt.toLocaleString()} บาท\n` +
        `📑 จากสลิป: ${ocr.amount.toLocaleString()} บาท\n` +
        `➖ ส่วนต่าง: ${amtDelta.toLocaleString()} บาท\n` +
        `🆔 SlipID: ${inbox.slipId}`
      );


      sendLineNotify_(
        `MM: Amount mismatch\nRoom: ${room}\nBill: ${cand.billId}\nBillAmt: ${billAmt}\nOCR: ${ocr.amount}\nΔ: ${amtDelta}\nSlip: ${inbox.slipId}`
      );

      return { ok:false, reason:'amount_mismatch' };
    }

    // amounts align → high confidence
    conf = amtDelta < 0.5 ? 0.98 : 0.90;
  }

  if (ocrOk && ocr.txDate) {
    const billMonth = String(cand.month || '');
    const txMonth   = Utilities.formatDate(ocr.txDate, 'Asia/Bangkok', 'yyyy-MM');
    if (billMonth === txMonth) conf = Math.min(0.99, conf + 0.02);
  }

  // ✅ Only AFTER passing checks, update the bill:
  updateBillWithSlip_({
    rowIndex: cand.rowIndex,
    slipId: inbox.slipId,
    markStatus: 'Slip Received'
  });

  updateInboxMatchResult_({
    rowIndex: inbox.rowIndex,
    status: 'matched',
    matchedBillId: cand.billId,
    confidence: conf,
    note: (ocrOk ? `OCR OK; bank=${ocr.bank||''}; ref=${ocr.txId||''}` : 'matched (no/failed OCR)')
  });

  // 🔔 NEW: notify admin group with Approve/Reject buttons
  notifyGroupPaymentMatched_({
    room,
    amountDue: cand.amountDue,
    billId: cand.billId,
    ocrAmount: (ocr && ocr.amount != null) ? Number(ocr.amount) : null,
    slipId: inbox.slipId,
    confidence: conf
  });

  return { ok:true, slipId: inbox.slipId, matchedBillId: cand.billId, amountDue: cand.amountDue };
}

function openRevenueSheetByName_(name) {
  const ss = SpreadsheetApp.openById(REVENUE_SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found in Revenue file: ' + name);
  return sh;
}


/*************** 9) ADMIN / DIAGNOSTICS / TRIGGERS / NOTIFY ***************/
function initAuth() {
  const notes = [];

  // Sheets access
  SpreadsheetApp.openById(SHEET_ID).getSheets();
  notes.push('Sheets OK');

  // Slip folders (use the resolved IDs from your config)
  DriveApp.getFolderById(TEMP_SLIP_FOLDER_ID).getName();
  notes.push('TempSlip OK: ' + TEMP_SLIP_FOLDER_ID);

  DriveApp.getFolderById(SLIP_FOLDER_ID).getName();
  notes.push('FinalSlip OK: ' + SLIP_FOLDER_ID);

  // Optional: ID folders if configured
  if (ID_FOLDER_ID) {
    try { DriveApp.getFolderById(ID_FOLDER_ID).getName(); notes.push('ID OK'); } catch (e) { notes.push('ID ERR'); throw e; }
  }
  if (TEMP_ID_FOLDER_ID) {
    try { DriveApp.getFolderById(TEMP_ID_FOLDER_ID).getName(); notes.push('TempID OK'); } catch (e) { notes.push('TempID ERR'); throw e; }
  }

  Logger.log('initAuth success: ' + notes.join(', '));
  return 'ok: ' + notes.join(', ');
}

function driveDiagnostics() {
  Logger.log('SLIP_FOLDER_ID=' + SLIP_FOLDER_ID);
  SpreadsheetApp.openById(SHEET_ID).getSheets();
  Logger.log('Drive root OK: ' + DriveApp.getRootFolder().getName());
  Logger.log('Slip folder OK: ' + DriveApp.getFolderById(SLIP_FOLDER_ID).getName());
  // 🔸 NEW
  try { Logger.log('ID folder OK: ' + DriveApp.getFolderById(ID_FOLDER_ID).getName()); } catch(e) {}
  try { Logger.log('Temp ID folder OK: ' + DriveApp.getFolderById(TEMP_ID_FOLDER_ID).getName()); } catch(e) {}
}

function testOCRSlip() {
  const testFileId = '1M4d4AVVB4EJ8JVxhXz5QPia6XYmYN3iC';  // <— ใส่ fileId จากข้อ 1
  try {
    const text = ocrSlipFromFileId_(testFileId);
    Logger.log("=== OCR RESULT START ===");
    Logger.log(text);
    Logger.log("=== OCR RESULT END ===");
  } catch (e) {
    Logger.log("OCR ERROR: " + e);
  }
}

/** Tell tenant their slip mismatched. */
function notifyTenantMismatch_(userId, { room, billId, amountDue, ocrAmount, delta, slipId }) {
  if (!userId) return;
  const lines = [
    '⚠️ ระบบตรวจพบว่ายอดสลิปไม่ตรงกับบิล',
    room ? `ห้อง: ${room}` : '',
    billId ? `บิล: ${billId}` : '',
    amountDue != null ? `ยอดบิล: ${amountDue.toLocaleString('en-US',{minimumFractionDigits:2})} บาท` : '',
    ocrAmount != null ? `ยอดจากสลิป: ${Number(ocrAmount).toLocaleString('en-US',{minimumFractionDigits:2})} บาท` : '',
    delta != null ? `ต่างกัน: ${delta.toFixed(2)} บาท` : '',
    slipId ? `รหัสสลิป: ${slipId}` : '',
    '',
    'กรุณาตรวจสอบหรือส่งสลิปใหม่อีกครั้งค่ะ'
  ].filter(Boolean).join('\n');

  pushMessage(userId, [{ type:'text', text: lines.slice(0,1200) }]);
}

/** Find Payments_Inbox row by SlipID. */
function findInboxRowBySlipId_(slipId) {
  const sh = openRevenueSheetByName_('Payments_Inbox');
  const hdr = getHeaders_(sh);
  const cSlipID = idxOf_(hdr, 'slipid');
  if (cSlipID < 0) return null;
  const vals = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][cSlipID]).trim() === String(slipId).trim()) return { rowIndex: i+2, headers: hdr };
  }
  return null;
}

function deriveBankMatchStatus_(billAccountCode, ocrAccountCode){
  const bill = String(billAccountCode || '').trim().toUpperCase();
  const ocr  = String(ocrAccountCode || '').trim().toUpperCase();
  if (!bill) return '';
  if (!ocr || ocr === 'NON_MATCH') return 'receiver_non_match';
  if (bill === ocr) return 'receiver_matched';
  return 'receiver_mismatch';
}

function getBillAccountById_(billId){
  const sh  = openRevenueSheetByName_('Horga_Bills');
  const hdr = getHeaders_(sh);
  const cBill   = idxOf_(hdr, 'billid');
  const cAcc    = idxOf_(hdr, 'account');
  if (cBill < 0) throw new Error('Horga_Bills missing BillID col');
  const vals = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][cBill]).trim() === String(billId).trim()) {
      const account = cAcc > -1 ? String(vals[i][cAcc]||'').trim().toUpperCase() : '';
      return { account, rowIndex: i+2 };
    }
  }
  throw new Error('Bill not found: ' + billId);
}

/** Update a bill row to mark slip received (idempotent). */
function setBillSlipReceived_(billId, slipId, statusText, bankMatchStatus) {
  const sh  = openRevenueSheetByName_('Horga_Bills');
  const hdr = getHeaders_(sh);
  const cBill   = idxOf_(hdr, 'billid');
  const cStatus = idxOf_(hdr, 'status');
  const cPaidAt = idxOf_(hdr, 'paidat');
  const cSlipID = idxOf_(hdr, 'slipid');
  const cBank   = idxOf_(hdr, 'bankmatchstatus');
  if (cBill < 0) throw new Error('Horga_Bills missing BillID col');
  const vals = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][cBill]).trim() === String(billId).trim()) {
      if (cSlipID > -1) sh.getRange(i+2, cSlipID+1).setValue(slipId || '');
      if (cPaidAt > -1) sh.getRange(i+2, cPaidAt+1).setValue(new Date());
      if (cStatus > -1) sh.getRange(i+2, cStatus+1).setValue(statusText || 'Slip Received');
      if (cBank > -1 && bankMatchStatus) sh.getRange(i+2, cBank+1).setValue(bankMatchStatus);
      return i+2;
    }
  }
  throw new Error('Bill not found: ' + billId);
}

/** Run once to install the onEdit trigger for the REVENUE_SHEET_ID file. */
function setupReviewTriggers() {
  // Remove old
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onReviewQueueEdit_') ScriptApp.deleteTrigger(t);
  });
  // Create new
  ScriptApp.newTrigger('onReviewQueueEdit_')
    .forSpreadsheet(REVENUE_SHEET_ID)   // <-- same file where Review_Queue lives
    .onEdit()
    .create();
}

/** Trigger target: handles edits in Review_Queue. */
function onReviewQueueEdit_(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getParent().getId() !== REVENUE_SHEET_ID) return;
    if (sheet.getName() !== 'Review_Queue') return;

    const row = e.range.getRow();
    if (row === 1) return; // header

    const hdr = getHeaders_(sheet);
    const get = (key) => {
      const c = idxOf_(hdr, key);
      return c > -1 ? sheet.getRange(row, c+1).getValue() : '';
    };
    const set = (key, val) => {
      const c = idxOf_(hdr, key);
      if (c > -1) sheet.getRange(row, c+1).setValue(val);
    };

    const decision   = String(get('Decision') || '').toUpperCase().trim();
    if (!decision) return; // ignore other edits


    const slipId     = String(get('SlipID') || '').trim();
    const billId     = String(get('ApplyBillID') || get('BillID') || '').trim();
    const adjAmt     = get('AdjustedAmount'); // optional
    const userId     = String(get('LineUserID') || '').trim();

    if (decision === 'APPROVE') {
      let bankMatchStatus = '';
      const inbox = findInboxRowBySlipId_(slipId);
      const billMeta = getBillAccountById_(billId);
      if (inbox) {
        const shIn = openRevenueSheetByName_('Payments_Inbox');
        const hIn  = inbox.headers;
        const cAcc = idxOf_(hIn, 'ocr_accountcode');
        const ocrAccCode = cAcc > -1 ? String(shIn.getRange(inbox.rowIndex, cAcc+1).getValue() || '').trim().toUpperCase() : '';
        bankMatchStatus = deriveBankMatchStatus_(billMeta.account, ocrAccCode);

        const cSt  = idxOf_(hIn, 'matchstatus');
        const cId  = idxOf_(hIn, 'matchedbillid');
        const cCf  = idxOf_(hIn, 'confidence');
        const cNt  = idxOf_(hIn, 'notes');
        if (cSt > -1) shIn.getRange(inbox.rowIndex, cSt+1).setValue('approved_manual');
        if (cId > -1) shIn.getRange(inbox.rowIndex, cId+1).setValue(billId);
        if (cCf > -1) shIn.getRange(inbox.rowIndex, cCf+1).setValue(1.0);
        if (cNt > -1) shIn.getRange(inbox.rowIndex, cNt+1).setValue(
          `approved by admin${adjAmt ? '; adjAmt=' + Number(adjAmt) : ''}`
        );
      }

      // 1) Mark bill as Slip Received
      setBillSlipReceived_(billId, slipId, 'Slip Received (Manual)', bankMatchStatus);

      // 3) Close the review row
      set('ResolvedAt', new Date());
      set('Reason', 'approved_manual');
      sendLineNotify_(`MM: Approved mismatch\nBill: ${billId}\nSlip: ${slipId}${adjAmt?`\nAdjAmt: ${adjAmt}`:''}`);
      if (userId) pushMessage(userId, [{ type:'text', text:'✅ ยืนยันยอดเรียบร้อย ขอบคุณค่ะ' }]);
    }

    if (decision === 'REJECT') {
      // 1) Update Payments_Inbox only
      const inbox = findInboxRowBySlipId_(slipId);
      if (inbox) {
        const shIn = openRevenueSheetByName_('Payments_Inbox');
        const hIn  = inbox.headers;
        const cSt  = idxOf_(hIn, 'matchstatus');
        const cCf  = idxOf_(hIn, 'confidence');
        const cNt  = idxOf_(hIn, 'notes');
        if (cSt > -1) shIn.getRange(inbox.rowIndex, cSt+1).setValue('rejected');
        if (cCf > -1) shIn.getRange(inbox.rowIndex, cCf+1).setValue(0.0);
        if (cNt > -1) shIn.getRange(inbox.rowIndex, cNt+1).setValue('rejected by admin');
      }
      // 2) Keep bill unpaid, close review
      set('ResolvedAt', new Date());
      set('Reason', 'rejected');
      sendLineNotify_(`MM: Rejected mismatch\nBill: ${billId}\nSlip: ${slipId}`);
      if (userId) pushMessage(userId, [{ type:'text', text:'❌ สลิปไม่ถูกต้อง กรุณาส่งสลิปใหม่ค่ะ' }]);
    }
  } catch (err) {
    Logger.log('onReviewQueueEdit_ error: ' + err);
    try { sendLineNotify_('MM error in onReviewQueueEdit_: ' + err); } catch(e){}
  }
}

/** ส่งข้อความเข้า LINE Notify กลุ่ม */
function sendLineNotify_(message) {
  if (!LINE_NOTIFY_TOKEN) { Logger.log('NO LINE_NOTIFY_TOKEN'); return; }
  const url = 'https://notify-api.line.me/api/notify';
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + LINE_NOTIFY_TOKEN },
    payload: { message: String(message || '') },
    muteHttpExceptions: true
  });
}

/** ทดสอบแจ้งเตือน (กด Run ครั้งเดียวเพื่อเทสต์) */
function testNotify() {
  sendLineNotify_('✅ ทดสอบแจ้งเตือนจาก MM_LineWebhook สำเร็จ');
}

function pingGoogle() {
  const r = UrlFetchApp.fetch('https://www.google.com', {muteHttpExceptions:true});
  Logger.log('google.com %s', r.getResponseCode());
}
function pingLineHost() {
  const r = UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {muteHttpExceptions:true});
  Logger.log('notify-api.line.me %s %s', r.getResponseCode(), r.getContentText().slice(0,60));
}

function alertAdminGroup_(text) {
  if (!ADMIN_GROUP_ID) { Logger.log('NO ADMIN_GROUP_ID'); return; }
  // ใช้ pushMessage() เดิมของโปรเจกต์ (มันเรียก /v2/bot/message/push อยู่แล้ว)
  pushMessage(ADMIN_GROUP_ID, [{ type: 'text', text }]);
}

/** Send an admin alert to the OA group (preferred), fallback to LINE Notify if configured. */
function adminNotify_(text) {
  var msg = String(text || '').slice(0, 1200); // keep short
  if (typeof ADMIN_GROUP_ID === 'string' && ADMIN_GROUP_ID) {
    try {
      pushMessage(ADMIN_GROUP_ID, [{ type: 'text', text: msg }]);
      return;
    } catch (e) {
      Logger.log('adminNotify_ push error: ' + e);
    }
  }
  // optional fallback if you also configured LINE_NOTIFY_TOKEN
  try { sendLineNotify_(msg); } catch (e) { Logger.log('adminNotify_ notify error: ' + e); }
}

function testAdminNotify() { adminNotify_('✅ Test from Apps Script'); }

/** Push a Flex card to the admin LINE group with Approve/Reject buttons. */
function notifyGroupPaymentMatched_({ room, amountDue, billId, ocrAmount, slipId, confidence }) {
  if (!ADMIN_GROUP_ID) { Logger.log('NO ADMIN_GROUP_ID'); return; }

  const title = '💵 Payment Received';
  const lines = [
    `Room: ${room || '-'}`,
    `Bill: ${billId || '-'}`,
    `Amount (Bill): ${amountDue != null ? Number(amountDue).toLocaleString() : '-'}`,
    `OCR Amount: ${ocrAmount != null ? Number(ocrAmount).toLocaleString() : '-'}`,
    `Confidence: ${confidence != null ? (Math.round(confidence*100) + '%') : '-'}`,
    slipId ? `SlipID: ${slipId}` : ''
  ].filter(Boolean);

  const approveData = `act=mgr_approve&bill=${encodeURIComponent(billId||'')}&slip=${encodeURIComponent(slipId||'')}&room=${encodeURIComponent(room||'')}`;
  const rejectData  = `act=mgr_reject&bill=${encodeURIComponent(billId||'')}&slip=${encodeURIComponent(slipId||'')}&room=${encodeURIComponent(room||'')}`;

  const flex = {
    type: 'flex',
    altText: `${title} • ${room || ''} • ${billId || ''}`,
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [{ type: 'text', text: title, weight: 'bold', size: 'lg' }]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: lines.map(t => ({ type: 'text', text: t, wrap: true, size: 'sm' }))
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'md',
        contents: [
          { type: 'button', style: 'primary', color: '#22c55e',
            action: { type: 'postback', label: '✅ Approve', data: approveData } },
          { type: 'button', style: 'secondary',
            action: { type: 'postback', label: '❌ Reject', data: rejectData } }
        ]
      }
    }
  };

  pushMessage(ADMIN_GROUP_ID, [flex]);
}

function testGroupApproveCard() {
  notifyGroupPaymentMatched_({
    room: 'A101',
    amountDue: 3800,
    billId: '2025-09-A101',
    ocrAmount: 3800,
    slipId: 'SLIP-TEST-1234',
    confidence: 0.98
  });
}

function sendCheckinPickerToUser(userId, roomId) {
  if (!userId) throw new Error('sendCheckinPickerToUser: missing userId');
  const pickerData = _buildPostbackData_({ act: 'checkin_pick', room: roomId || '' }) || 'act=checkin_pick';
  const payload = {
    to: userId,
    messages: [{
      type: 'template',
      altText: 'เลือกวัน–เวลาเช็คอิน',
      template: {
        type: 'confirm',
        text: `ห้อง ${roomId || '-'}\nโปรดเลือกวัน–เวลาเช็คอิน (${CHECKIN_PICKER_EARLIEST_TIME_LABEL}-${CHECKIN_PICKER_LATEST_TIME_LABEL} น. เท่านั้น)`,
        actions: [
          {
            type: 'datetimepicker',
            label: 'เลือกวัน–เวลา',
            data: pickerData,
            mode: 'datetime',
            max: CHECKIN_PICKER_MAX_DATETIME
          },
          { type: 'postback', label: 'ยกเลิก', data: 'act=rent_cancel' }
        ]
      }
    }]
  };

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('CHECKIN_PICKER_PUSH code=%s body=%s', res.getResponseCode(), res.getContentText());
}

/** === TEST: push to *your* LINE ID once === */
function testSendCheckinPickerToMe() {
  const MY_LINE_USER_ID = 'Ue90558b73d62863e2287ac32e69541a3'; // <- yours
  const roomId = _findRoomByUserId_(MY_LINE_USER_ID);          // optional, just for nicer text
  sendCheckinPickerToUser(MY_LINE_USER_ID, roomId);
}

function handleMoveoutFromWorker_(data) {
  try {
    const roomId  = String(data.roomId || '').toUpperCase().trim();
    const dateISO = String(data.dateISO || '').trim(); // YYYY-MM-DD
    const userId  = String(data.lineUserId || '').trim();

    console.log('MOVEOUT_START', JSON.stringify({ roomId, dateISO, userId }));

    if (!roomId || !dateISO) throw new Error('missing room/date');

    const ss = SpreadsheetApp.openById(SHEET_ID);

    // 1) Append to Moveouts log
    let sh = ss.getSheetByName('Moveouts');
    if (!sh) sh = ss.insertSheet('Moveouts');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Timestamp','Room','MoveoutDate','LineUserId','Status','Note']);
    }
    sh.appendRow([new Date(), roomId, dateISO, userId, 'PENDING', 'via Edge']);
    console.log('MOVEOUT_LOGGED');

    // 2) Update Rooms
    const rooms = ss.getSheetByName('Rooms');
    if (!rooms) throw new Error('Rooms sheet not found');

    const values  = rooms.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());

    // exact-name indices (your sheet matches these)
    const idxRoom    = colEq_(headers, 'RoomId');
    const idxStatus  = colEq_(headers, 'Status');
    const idxMoveout = colEq_(headers, 'Move-out Date');
    const idxNext    = colEq_(headers, 'Next Available');

    console.log('ROOMS_COLIDX', JSON.stringify({ idxRoom, idxStatus, idxMoveout, idxNext, headerCount: headers.length }));

    if (idxRoom === -1 || idxStatus === -1) {
      throw new Error('Missing columns RoomId/Status in Rooms');
    }

    let updated = false;
    for (let r = 1; r < values.length; r++) {
      const id = String(values[r][idxRoom] || '').toUpperCase().trim();
      if (id === roomId) {
        if (idxStatus  > -1) rooms.getRange(r + 1, idxStatus  + 1).setValue('soon');
        if (idxMoveout > -1) rooms.getRange(r + 1, idxMoveout + 1).setValue(dateISO);
        if (idxNext    > -1) {
          const dt = new Date(dateISO);
          dt.setDate(dt.getDate() + 1);
          rooms.getRange(r + 1, idxNext + 1).setValue(dt);
        }
        console.log('ROOMS_UPDATED_EXISTING_ROW', r + 1);
        updated = true;
        break;
      }
    }

    if (!updated) {
      const newRow = new Array(headers.length).fill('');
      if (idxRoom   > -1) newRow[idxRoom]   = roomId;
      if (idxStatus > -1) newRow[idxStatus] = 'soon';
      if (idxMoveout> -1) newRow[idxMoveout]= dateISO;
      if (idxNext   > -1) {
        const dt = new Date(dateISO);
        dt.setDate(dt.getDate() + 1);
        newRow[idxNext] = dt;
      }
      rooms.appendRow(newRow);
      console.log('ROOMS_APPENDED_NEW_ROW');
      updated = true;
    }

    // 3) Optional tenant notify (push)
    if (userId) {
      try {
        pushMessage(userId, [{
          type: 'text',
          text: `✅ รับคำขอแจ้งออกแล้ว\nห้อง ${roomId}\nวันย้ายออก: ${dateISO}`
        }]);
      } catch (e) {
        console.log('TENANT_PUSH_SKIP', e);
      }
    }

    console.log('MOVEOUT_DONE', updated);
    return updated;
  } catch (err) {
    console.error('MOVEOUT_ERR', err);
    return false;
  }
}

function handleTextPush_(event) {
  const userId   = event.source?.userId || '';
  const userText = (event.message?.text || '').trim();
  if (!userId || !userText) return;

  if (handleCheckinPickerTextCommand_(event)) return;

  // (0) Move-out magic link (with loading + button)
  if (/^\s*แจ้งออก\s*$/i.test(userText)) {
    // show the chat loading spinner immediately (7s max per API)
    // pushWithLoading() starts the spinner then sends the message
    const room = findRoomByLineId_(userId);
    const tok  = issueToken_(userId, room, 20, 'moveout');

    // ensure trailing slash before adding ?t=
    const base = /\/$/.test(FRONTEND_BASE) ? FRONTEND_BASE : (FRONTEND_BASE + '/');
    const url  = base + '?t=' + encodeURIComponent(tok);

    return pushWithLoading(userId, [{
      type: 'template',
      altText: 'ฟอร์มแจ้งย้ายออก',
      template: {
        type: 'buttons',
        text: 'กดปุ่มเพื่อเปิดฟอร์มแจ้งย้ายออก\n(ลิงก์ใช้ครั้งเดียว ภายใน 20 นาที)',
        actions: [
          { type: 'uri', label: 'เปิดฟอร์มแจ้งย้ายออก', uri: url }
        ]
      }
    }], 7); // spinner duration (seconds)
  }


  // (A) Pay-rent keywords
  if (/^(ชำระค่าเช่า|จ่ายค่าเช่า|pay\s*rent|ค่าเช่า)$/i.test(userText)) {
    const cache = CacheService.getUserCache();
    cache.put(userId + ':rent_flow', JSON.stringify({ step: 'await_room' }), 2 * 60 * 60);
    return pushWithLoading(userId, [{ type:'text', text:'พิมพ์เบอร์ห้องของคุณ (เช่น A101)' }], 5);
  }

  // (B) Booking code
  const m = userText.match(/^#?\s*MM\d{3,}$/i);
  if (!m) return;

  const codeDisplay = m[0].toUpperCase().replace(/\s/g, '');
  const codeKey     = codeDisplay.replace(/^#/, '');

  // 🔒 Debounce: avoid duplicate “กำลังตรวจสอบ…” for the same user+code within 15s
  const cache = CacheService.getUserCache();
  const debKey = `${userId}:checking:${codeKey}`;
  if (!cache.get(debKey)) {
    pushMessage(userId, [{ type:'text', text:'⏳ กำลังตรวจสอบรหัสจอง…' }]);
    safeStartLoading(userId, 6);
    cache.put(debKey, '1', 15); // 15 seconds TTL
  }

  const row = findBookingRow(codeKey);
  if (!row) {
    return pushMessage(userId, [{ type:'text', text:'ไม่พบข้อมูลรหัสนี้' }]);
  }

  return pushMessage(userId, [{
    type:'template',
    altText:'ยืนยันการจอง',
    template:{
      type:'buttons',
      text:([
        `รหัส: ${codeDisplay}`,
        `ห้อง: ${row.roomId || '-'}`,
        `ชื่อ: ${row.name}`
      ].join('\n')).slice(0,160),
      actions:[ { type:'postback', label:'ยืนยัน', data:'act=confirm&code='+codeKey } ]
    }
  }]);
}

function send_(event, messages, withLoadingSecs) {
  const userId     = event.source?.userId || '';
  const replyToken = event.replyToken;
  const secs       = withLoadingSecs || 5;

  // Try REPLY first (works for native LINE webhooks)
  if (replyToken) {
    try {
      const url = 'https://api.line.me/v2/bot/message/reply';
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
        },
        payload: JSON.stringify({ replyToken, messages }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      Logger.log('REPLY_TRY code=%s body=%s', code, res.getContentText());
      if (code >= 200 && code < 300) return; // success → we're done
    } catch (e) {
      Logger.log('REPLY_TRY error: ' + e);
      // fall through to push
    }
  }

  // Fallback PUSH (works for Worker-forwarded events)
  if (userId) return pushWithLoading(userId, messages, secs);
}
