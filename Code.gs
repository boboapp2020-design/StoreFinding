/*************************************************************
 * ผู้ช่วยคลังพัสดุ — Google Apps Script (Web App API) v2
 * รองรับข้อมูลจริงจาก SAP (Material Master) 40,000+ แถว
 *
 * ★ จุดเด่น: ทนต่อไฟล์ที่ "คอลัมน์ไม่ตรงกันในแต่ละรอบ"
 *   - map ข้อมูลด้วย "ชื่อหัวคอลัมน์" ไม่ใช่ตำแหน่ง → ข้อมูลไม่เพี้ยน
 *   - บังคับแค่ Material + Material Description ที่เหลือ optional
 *   - คอลัมน์แปลกปลอม (เช่นคอลัมน์มูลค่า) จะถูกข้ามอัตโนมัติ
 *   - รองรับชื่อคอลัมน์ทั้งไทย/อังกฤษ/SAP
 *
 * ── วิธีติดตั้ง ────────────────────────────────────────────
 * 1) นำเข้าข้อมูล: เปิด Google Sheet → File → Import → Upload
 *    เลือกไฟล์ MLStoreFinding-xx.xlsx → Import location: "Replace current sheet"
 *    (ให้ชีตข้อมูลชื่อ "Data")  *ไฟล์ .xlsx ของคุณใช้ชื่อชีต Data อยู่แล้ว
 * 2) Extensions → Apps Script → วางไฟล์นี้ทั้งหมด → บันทึก
 * 3) รีเฟรช Sheet → เมนู "📦 คลังพัสดุ" → "ตรวจสอบคอลัมน์ที่ตรวจจับได้"
 *    เพื่อยืนยันว่า map ถูกต้อง
 * 4) Deploy → New deployment → Web app
 *      Execute as: Me | Who has access: Anyone → Deploy → คัดลอก URL
 * 5) วาง URL ในแอป (โหมดผู้ดูแล → เชื่อมต่อ Google Sheet)
 *
 * *แก้โค้ดภายหลังต้อง Deploy → Manage deployments → New version ทุกครั้ง
 *************************************************************/

var DATA_SHEET = 'Data';       // ชื่อชีตที่เก็บข้อมูล
var DEFAULT_LIMIT = 25;        // จำนวนผลลัพธ์สูงสุดต่อการค้นหา

/* ★ กติกากรองข้อมูล ★
 * ตัดรหัส (Material) ที่ขึ้นต้นด้วยตัว N ออกจากฐานข้อมูลค้นหาทั้งหมด
 * ยกเว้นรายการที่ชื่อ (Material Description) ตรงกับรายชื่อใน EXCEPT_NAMES */
var EXCLUDE_PREFIX = 'N';
var EXCEPT_NAMES = ['น้ำมันดีเซลหมุนเร็ว (ขายชาวไร่)'];

function isExcluded_(code, name) {
  if (String(code).charAt(0).toUpperCase() !== EXCLUDE_PREFIX) return false;
  var nm = String(name || '').trim();
  for (var i = 0; i < EXCEPT_NAMES.length; i++) {
    if (nm === EXCEPT_NAMES[i]) return false;   // อยู่ในข้อยกเว้น → ไม่ตัด
  }
  return true;   // ขึ้นต้นด้วย N และไม่ใช่ข้อยกเว้น → ตัดออก
}

/* ============ พจนานุกรมชื่อคอลัมน์ (canonical ← aliases) ============
 * เพิ่มชื่อคอลัมน์ใหม่ได้ที่นี่ ถ้าไฟล์รอบใหม่ใช้ชื่อหัวต่างออกไป      */
var ALIASES = {
  code:      ['material', 'materialnumber', 'matnr', 'รหัสพัสดุ', 'รหัส', 'รหัสวัสดุ', 'รหัสสินค้า', 'code'],
  name:      ['materialdescription', 'materialdesc', 'description', 'ชื่อพัสดุ', 'ชื่อ', 'ชื่อสินค้า', 'รายละเอียด', 'รายการ'],
  unit:      ['baseunitofmeasure', 'baseunit', 'unit', 'uom', 'หน่วยนับ', 'หน่วย'],
  qty:       ['unrestricted', 'unrestrictedstock', 'จำนวน', 'คงเหลือ', 'สต็อก', 'qty', 'quantity', 'stock'],
  group:     ['materialgroup', 'matgroup', 'หมวดหมู่', 'หมวด', 'กลุ่มวัสดุ', 'group'],
  mtype:     ['materialtype', 'ประเภทวัสดุ', 'ประเภท', 'type'],
  plant:     ['plant', 'โรงงาน', 'รหัสโรงงาน'],
  plantName: ['name1', 'ชื่อโรงงาน'],
  sloc:      ['storagelocation', 'sloc', 'ที่เก็บ', 'รหัสที่เก็บ'],
  slocName:  ['descrofstorageloc', 'descriptionofstoragelocation', 'descofstorageloc', 'ที่จัดเก็บ', 'คลัง', 'ชื่อที่เก็บ']
};
var REQUIRED = ['code', 'name'];   // ต้องมี ไม่งั้นถือว่าไฟล์ผิดรูปแบบ

/* ---------- utils ---------- */
function norm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s\._\-\/()]/g, '').trim(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function dataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(DATA_SHEET) || ss.getSheets()[0];
}

/* สร้างตารางจับคู่ canonical → ดัชนีคอลัมน์ (0-based) จากแถวหัว */
function buildMap_(headerRow) {
  var normHeaders = headerRow.map(function (h) { return norm_(h); });
  var map = {};
  Object.keys(ALIASES).forEach(function (field) {
    var aliases = ALIASES[field];
    for (var i = 0; i < normHeaders.length; i++) {
      if (aliases.indexOf(normHeaders[i]) !== -1) { map[field] = i; break; }
    }
  });
  return map;
}

/* ---------- อ่าน + รวมข้อมูล (dedupe by Material, รวม stock หลายที่เก็บ) ---------- */
function readCatalog_() {
  var sh = dataSheet_();
  var last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 2) return { items: [], map: {}, headers: [] };

  var values = sh.getRange(1, 1, last, width).getValues();
  var headers = values[0];
  var map = buildMap_(headers);

  var missing = REQUIRED.filter(function (f) { return map[f] == null; });
  if (missing.length) {
    throw new Error('ไม่พบคอลัมน์บังคับ: ' + missing.join(', ') +
      ' | หัวคอลัมน์ที่เจอ: ' + headers.join(', '));
  }

  var col = function (row, idx) { return idx == null ? '' : String(row[idx]).trim(); };
  var byCode = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var code = col(row, map.code);
    if (!code) continue;
    if (isExcluded_(code, col(row, map.name))) continue;   // ตัดรหัสขึ้นต้นด้วย N (เว้นข้อยกเว้น)
    var qty = map.qty == null ? 0 : (parseFloat(String(row[map.qty]).replace(/[, ]/g, '')) || 0);

    var it = byCode[code];
    if (!it) {
      it = {
        code: code, name: col(row, map.name), unit: col(row, map.unit),
        group: col(row, map.group), type: col(row, map.mtype),
        plant: col(row, map.plant), plantName: col(row, map.plantName),
        qty: 0, locs: []
      };
      byCode[code] = it;
    }
    it.qty += qty;
    var sloc = col(row, map.sloc), sName = col(row, map.slocName);
    if (sloc || sName) it.locs.push({ sloc: sloc, name: sName, qty: qty });
    if (!it.name) it.name = col(row, map.name);
  }

  var arr = [];
  Object.keys(byCode).forEach(function (k) {
    var it = byCode[k];
    it._nn = norm_(it.name);
    it._nc = norm_(it.code);
    it._tk = String(it.name).split(/\s+/).map(function (t) { return norm_(t); }).filter(Boolean);
    arr.push(it);
  });
  return { items: arr, map: map, headers: headers };
}

/* ---------- Levenshtein (สำหรับเดาคำเมื่อพิมพ์ผิด) ---------- */
function lev_(a, b) {
  var m = a.length, n = b.length;
  if (!m) return n; if (!n) return m;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= n; j++) prev[j] = j;
  for (i = 1; i <= m; i++) {
    cur[0] = i;
    for (j = 1; j <= n; j++) {
      var c = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + c);
    }
    for (j = 0; j <= n; j++) prev[j] = cur[j];
  }
  return prev[n];
}

/* ---------- ค้นหา ---------- */
function searchCatalog_(qRaw, limit) {
  var q = norm_(qRaw);
  if (!q) return { mode: 'empty', items: [] };
  limit = limit || DEFAULT_LIMIT;

  var cat = readCatalog_().items;
  var hits = [];
  var i, it, sc;

  // รอบ 1: ตรงตัว / เป็นส่วนหนึ่งของข้อความ (เร็ว ครอบคลุมส่วนใหญ่)
  for (i = 0; i < cat.length; i++) {
    it = cat[i]; sc = 0;
    if (it._nc === q || it._nn === q) sc = 100;
    else if (it._nc.indexOf(q) !== -1) sc = 93;
    else if (it._nn.indexOf(q) !== -1) sc = 85 - Math.min(20, Math.abs(it._nn.length - q.length));
    if (sc > 0) hits.push([sc, it]);
  }
  if (hits.length) {
    hits.sort(function (a, b) { return b[0] - a[0]; });
    return { mode: 'match', items: hits.slice(0, limit).map(pub_) };
  }

  // รอบ 2 (เมื่อไม่เจอเลย): เดาคำด้วย fuzzy แบบเทียบรายคำ (token)
  // รองรับทั้ง "พิมพ์ผิด" (สลับ/ตกตัวอักษร) และ "พิมพ์ไม่ครบคำ" (เทียบกับส่วนหน้าของคำ)
  var best = [];
  var qlen = q.length;
  for (i = 0; i < cat.length; i++) {
    it = cat[i];
    var toks = it._tk || [];
    var localBest = 0;
    for (var ti = 0; ti < toks.length; ti++) {
      var tk = toks[ti];
      if (!tk) continue;
      // ถ้าคำยาวกว่าคำค้นมาก → เทียบเฉพาะส่วนหน้าของคำ (กรณีผู้ใช้พิมพ์ไม่ครบ)
      var cand = (tk.length > qlen + 1) ? tk.substring(0, qlen) : tk;
      if (Math.abs(cand.length - qlen) > 3) continue;
      var sim = 1 - lev_(q, cand) / Math.max(qlen, cand.length);
      if (sim > localBest) localBest = sim;
    }
    if (localBest >= 0.6) best.push([localBest, it]);
  }
  best.sort(function (a, b) { return b[0] - a[0]; });
  return { mode: best.length ? 'fuzzy' : 'none', items: best.slice(0, limit).map(function (x) { return pub_(x[1]); }) };
}

/* ตัวอย่างรายการ (อ่านเฉพาะแถวต้น ๆ เพื่อความเร็วตอนเปิดแอป) */
function sample_(n) {
  var sh = dataSheet_();
  var w = sh.getLastColumn(), rows = Math.min(sh.getLastRow(), 80);
  if (rows < 2) return [];
  var vals = sh.getRange(1, 1, rows, w).getValues();
  var map = buildMap_(vals[0]);
  if (map.code == null || map.name == null) return [];
  var seen = {}, out = [];
  for (var r = 1; r < vals.length && out.length < n; r++) {
    var code = String(vals[r][map.code]).trim();
    if (!code || seen[code]) continue; seen[code] = 1;
    if (isExcluded_(code, map.name != null ? String(vals[r][map.name]).trim() : '')) continue;
    out.push({
      code: code, name: String(vals[r][map.name]).trim(),
      unit: map.unit != null ? String(vals[r][map.unit]).trim() : '',
      qty: map.qty != null ? (parseFloat(vals[r][map.qty]) || 0) : 0,
      group: map.group != null ? String(vals[r][map.group]).trim() : ''
    });
  }
  return out;
}

function pub_(entry) {
  var it = entry[1] || entry;   // รองรับทั้ง [score,item] และ item
  return {
    code: it.code, name: it.name, unit: it.unit, qty: it.qty,
    group: it.group, type: it.type, plant: it.plant, plantName: it.plantName,
    locs: (it.locs || []).slice(0, 8)
  };
}

/* ================= Web App endpoints ================= */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'meta';
  try {
    if (action === 'search') {
      var res = searchCatalog_(p.q || '', parseInt(p.limit) || DEFAULT_LIMIT);
      return json_({ ok: true, mode: res.mode, count: res.items.length, items: res.items, q: p.q || '' });
    }
    if (action === 'sample') {
      return json_({ ok: true, items: sample_(parseInt(p.n) || 6) });
    }
    // meta: ใช้ทดสอบการเชื่อมต่อ
    var info = readCatalog_();
    var detected = {};
    Object.keys(info.map).forEach(function (k) { detected[k] = info.headers[info.map[k]]; });
    return json_({ ok: true, materials: info.items.length, columns: detected, headers: info.headers });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (x) { return json_({ ok: false, error: 'ระบบไม่ว่าง ลองใหม่' }); }
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'upsert';
    if (action === 'replaceAll') return json_(replaceAll_(body.items || []));
    return json_(upsert_(body.items || [], body.mode || 'set'));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* map object ที่ส่งมา (คีย์เป็นชื่อหัวอะไรก็ได้) → canonical ด้วย ALIASES */
function mapIncoming_(obj) {
  var out = {};
  Object.keys(obj).forEach(function (key) {
    var nk = norm_(key);
    Object.keys(ALIASES).forEach(function (field) {
      if (out[field] === undefined && ALIASES[field].indexOf(nk) !== -1) out[field] = obj[key];
    });
  });
  return out;
}

/* อัปเดต/เพิ่มทีละรายการ — เขียนลงตาม header map ของชีตปัจจุบัน (ไม่ทำลายโครงสร้างเดิม) */
function upsert_(items, mode) {
  var sh = dataSheet_();
  var last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 1) return { ok: false, error: 'ชีตยังไม่มีหัวตาราง — นำเข้าข้อมูลก่อน' };

  var values = sh.getRange(1, 1, last, width).getValues();
  var map = buildMap_(values[0]);
  if (map.code == null) return { ok: false, error: 'ไม่พบคอลัมน์ Material ในชีต' };

  var rowOf = {};
  for (var r = 1; r < values.length; r++) {
    var c = String(values[r][map.code]).trim().toLowerCase();
    if (c) rowOf[c] = r;
  }

  var added = 0, updated = 0;
  items.forEach(function (raw) {
    var it = mapIncoming_(raw);
    var code = String(it.code || '').trim();
    if (!code) return;
    var key = code.toLowerCase();
    var set = function (row, field, val) { if (map[field] != null && val !== undefined && val !== '') row[map[field]] = val; };

    if (rowOf.hasOwnProperty(key)) {
      var row = values[rowOf[key]];
      if (map.qty != null && it.qty !== undefined && it.qty !== '') {
        var q = parseFloat(String(it.qty).replace(/[, ]/g, '')) || 0;
        row[map.qty] = (mode === 'add') ? ((parseFloat(row[map.qty]) || 0) + q) : q;
      }
      set(row, 'name', it.name); set(row, 'unit', it.unit); set(row, 'group', it.group);
      updated++;
    } else {
      var nr = new Array(width).fill('');
      nr[map.code] = code;
      set(nr, 'name', it.name); set(nr, 'unit', it.unit); set(nr, 'qty', it.qty); set(nr, 'group', it.group);
      values.push(nr); rowOf[key] = values.length - 1;
      added++;
    }
  });

  sh.getRange(1, 1, values.length, width).setValues(values);
  return { ok: true, added: added, updated: updated, count: values.length - 1 };
}

/* แทนที่ข้อมูลทั้งหมดด้วยหัวตาราง canonical */
function replaceAll_(items) {
  var sh = dataSheet_();
  sh.clearContents();
  var HEAD = ['Material', 'Material Description', 'Material Group', 'Base Unit of Measure', 'Unrestricted', 'Storage Location', 'Descr. of Storage Loc.', 'Plant', 'Name 1'];
  var rows = [HEAD];
  items.forEach(function (raw) {
    var it = mapIncoming_(raw);
    if (!it.code) return;
    rows.push([it.code, it.name || '', it.group || '', it.unit || '', it.qty || 0, it.sloc || '', it.slocName || '', it.plant || '', it.plantName || '']);
  });
  sh.getRange(1, 1, rows.length, HEAD.length).setValues(rows);
  sh.setFrozenRows(1);
  return { ok: true, count: rows.length - 1 };
}

/* ================= เมนูช่วยเหลือในหน้า Sheet ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 คลังพัสดุ')
    .addItem('ตรวจสอบคอลัมน์ที่ตรวจจับได้', 'checkColumns')
    .addItem('นับจำนวนรายการ (unique Material)', 'countItems')
    .addToUi();
}
function checkColumns() {
  var sh = dataSheet_();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = buildMap_(headers);
  var lines = ['หัวคอลัมน์ในชีต:', headers.join(' | '), '', 'จับคู่ได้:'];
  Object.keys(ALIASES).forEach(function (f) {
    lines.push('• ' + f + ' → ' + (map[f] != null ? '"' + headers[map[f]] + '" (คอลัมน์ที่ ' + (map[f] + 1) + ')' : '❌ ไม่พบ'));
  });
  var missing = REQUIRED.filter(function (f) { return map[f] == null; });
  lines.push('', missing.length ? '⚠️ ขาดคอลัมน์บังคับ: ' + missing.join(', ') : '✅ คอลัมน์บังคับครบ พร้อมใช้งาน');
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}
function countItems() {
  try {
    var info = readCatalog_();
    SpreadsheetApp.getUi().alert('มีรายการทั้งหมด (unique Material): ' + info.items.length + ' รายการ');
  } catch (err) {
    SpreadsheetApp.getUi().alert('ผิดพลาด: ' + err.message);
  }
}
