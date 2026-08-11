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
/* ตัดวรรณยุกต์/ไม้ไต่คู้/ทัณฑฆาต ของไทย(U+0E47-4C) และลาว(U+0EC8-CD) เพื่อจับคำพิมพ์ผิดได้ทน (นอต=น็อต, ปั้ม=ปั๊ม) */
function stripTones_(s) {
  s = String(s); var out = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if ((c >= 0x0E47 && c <= 0x0E4C) || (c >= 0x0EC8 && c <= 0x0ECD)) continue; // วรรณยุกต์/ไม้ไต่คู้/ทัณฑฆาต ไทย+ลาว
    out += s.charAt(i);
  }
  return out;
}
function normLoose_(s) { return stripTones_(norm_(s)); }
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

/* ★ รันฟังก์ชันนี้จาก Editor 1 ครั้ง → Google จะเด้งหน้าขออนุญาต "เชื่อมต่ออินเทอร์เน็ต"
 *   (เรียก UrlFetchApp ตรง ๆ ไม่ครอบ try/catch จึงเด้งหน้าอนุญาตจริง)
 *   หลังกด Allow แล้ว ดูผลที่ Execution log — ควรได้ HTTP 200 + คำแปลอังกฤษ */
function testAI() {
  var pr = PropertiesService.getScriptProperties();
  var key = pr.getProperty('AI_API_KEY') || pr.getProperty('GROQ_API_KEY');
  var url = pr.getProperty('AI_API_URL') || 'https://api.groq.com/openai/v1/chat/completions';
  var model = pr.getProperty('AI_MODEL') || 'llama-3.1-8b-instant';
  var res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: model, temperature: 0, max_tokens: 40, messages: [{ role: 'user', content: 'reply with one english word: knife' }] }),
    muteHttpExceptions: true
  });
  Logger.log('HTTP ' + res.getResponseCode() + '  |  ' + res.getContentText());
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
    // ข้ามแถวที่ "Descr. of Storage Loc." มีคำว่า DSV
    if (map.slocName != null && String(row[map.slocName]).toUpperCase().indexOf('DSV') !== -1) continue;
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
    it._tk = String(it.name).split(/[\s,;:\/()\-\.]+/).map(function (t) { return norm_(t); }).filter(Boolean);
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
/* ============ พจนานุกรมคำพ้อง ไทย/ลาว ↔ อังกฤษ (แก้/เพิ่มได้ตามต้องการ) ============ */
/* แต่ละกลุ่มรวมคำที่ความหมายเดียวกัน: ไทย | ลาว | อังกฤษ */
var SYN_GROUPS = [
  ['ปากกา', 'ปากกาลูกลื่น', 'ບິກ', 'ປາກກາ', 'pen', 'ballpoint'],
  ['ดินสอ', 'ສໍ', 'pencil'],
  ['กระดาษ', 'ເຈ້ຍ', 'paper'],
  ['สายไฟ', 'สายเคเบิล', 'ສາຍໄຟ', 'wire', 'cable'],
  ['น็อต', 'นอต', 'สกรู', 'ນັອດ', 'bolt', 'nut', 'screw'],
  ['ถุงมือ', 'ຖົງມື', 'glove', 'gloves'],
  ['หมวก', 'หมวกนิรภัย', 'ໝວກ', 'helmet'],
  ['หลอดไฟ', 'หลอด', 'ຫລອດໄຟ', 'bulb', 'lamp', 'led'],
  ['เทป', 'ເທບ', 'tape'],
  ['น้ำมัน', 'ນ້ຳມັນ', 'oil', 'lubricant'],
  ['จารบี', 'ຈາລະບີ', 'grease'],
  ['แบตเตอรี่', 'ถ่าน', 'ถ่านไฟ', 'ຖ່ານ', 'ຖ່ານໄຟ', 'ໝໍ້ໄຟ', 'battery'],
  ['มอเตอร์', 'ມໍເຕີ', 'motor'],
  ['ปั๊ม', 'ปั้ม', 'ປໍ້າ', 'pump'],
  ['ท่อ', 'ທໍ່', 'pipe', 'tube'],
  ['วาล์ว', 'ประตูน้ำ', 'ວາລ໌', 'valve'],
  ['ลูกปืน', 'ตลับลูกปืน', 'ໝາກປືນ', 'bearing'],
  ['สวิตช์', 'ສະວິດ', 'switch'],
  ['เบรก', 'ເບຣກ', 'brake'],
  ['กรอง', 'ไส้กรอง', 'ໄສ້ກອງ', 'filter'],
  ['แหวน', 'ແຫວນ', 'washer'],
  ['ประเก็น', 'ปะเก็น', 'ปะกำ', 'gasket'],
  ['ข้อต่อ', 'ข้องอ', 'ຂໍ້ຕໍ່', 'elbow', 'joint', 'coupling', 'fitting'],
  ['สายพาน', 'ສາຍພານ', 'belt'],
  ['โซ่', 'ໂສ້', 'chain'],
  ['พัดลม', 'ພັດລົມ', 'fan'],
  ['ฟิวส์', 'ຟິວ', 'fuse'],
  ['รีเลย์', 'relay'],
  ['เซนเซอร์', 'sensor'],
  ['เกจ', 'เกจวัด', 'gauge'],
  ['ประแจ', 'ປະແຈ', 'wrench', 'spanner'],
  ['ค้อน', 'ຄ້ອນ', 'hammer'],
  ['ไขควง', 'ໄຂຄວງ', 'screwdriver'],
  ['คีม', 'ຄີມ', 'plier', 'pliers'],
  ['เลื่อย', 'ເລື່ອຍ', 'saw'],
  ['สว่าน', 'ສະຫວ່ານ', 'drill'],
  ['ตะปู', 'ຕະປູ', 'nail'],
  ['กาว', 'ກາວ', 'glue', 'adhesive'],
  ['แปรง', 'ແປງ', 'brush'],
  ['เชือก', 'ເຊືອກ', 'rope'],
  ['ยาง', 'ຢາງ', 'tire', 'tyre', 'rubber'],
  ['ลูกสูบ', 'piston'],
  ['เพลา', 'shaft'],
  ['เฟือง', 'ເຟືອງ', 'gear'],
  ['หัวฉีด', 'nozzle', 'injector'],
  ['ถัง', 'ຖັງ', 'tank', 'drum'],
  ['ฝา', 'cover', 'cap', 'lid'],
  // ---- เพิ่มเติม: ของใช้ในโรงงาน / ออฟฟิศ / เซฟตี้ ----
  ['ยางลบ', 'eraser'],
  ['ไม้บรรทัด', 'ruler'],
  ['กรรไกร', 'scissors'],
  ['ที่เย็บกระดาษ', 'แม็ก', 'stapler'],
  ['ลวดเย็บ', 'staple'],
  ['แฟ้ม', 'folder', 'file'],
  ['ซองจดหมาย', 'envelope'],
  ['หมึก', 'ink'],
  ['ตลับหมึก', 'cartridge', 'toner'],
  ['แว่นตานิรภัย', 'แว่นตา', 'goggles', 'safety glasses'],
  ['รองเท้าเซฟตี้', 'รองเท้า', 'safety shoes', 'boots'],
  ['หน้ากาก', 'mask', 'respirator'],
  ['ที่อุดหู', 'ปลั๊กอุดหู', 'ear plug', 'earplug'],
  ['เสื้อสะท้อนแสง', 'safety vest'],
  ['สายยาง', 'hose'],
  ['หน้าแปลน', 'flange'],
  ['ซีล', 'seal'],
  ['โอริง', 'oring', 'o-ring'],
  ['ลูกลอย', 'float'],
  ['โซลินอยด์', 'solenoid'],
  ['คอนแทคเตอร์', 'contactor'],
  ['เบรกเกอร์', 'breaker'],
  ['คาปาซิเตอร์', 'คาปา', 'capacitor'],
  ['สตาร์ทเตอร์', 'starter'],
  ['เทอร์มินอล', 'terminal'],
  ['ปลั๊กไฟ', 'ปลั๊ก', 'plug'],
  ['เต้ารับ', 'socket', 'receptacle'],
  ['ตะแกรง', 'screen', 'mesh'],
  ['มู่เล่ย์', 'มูเล่', 'pulley'],
  ['สปริง', 'spring'],
  ['ลวด', 'wire'],
  ['แผ่นเหล็ก', 'steel plate', 'sheet'],
  ['กระบอกลม', 'ลูกสูบลม', 'air cylinder', 'pneumatic'],
  ['เกจวัดแรงดัน', 'เกจวัด', 'pressure gauge'],
  ['เทอร์โมมิเตอร์', 'thermometer'],
  ['ปั๊มน้ำ', 'water pump'],
  ['ใบพัด', 'impeller'],
  ['ฮีตเตอร์', 'heater'],
  ['ผ้าเบรก', 'brake pad', 'brake lining'],
  ['ลูกหมาก', 'ball joint'],
  ['ซีลกันน้ำมัน', 'oil seal'],
  ['บูช', 'bush', 'bushing'],
  ['ปลอก', 'sleeve'],
  ['หมุด', 'rivet'],
  ['กิ๊บ', 'clip'],
  ['เข็มขัดรัดท่อ', 'แคลมป์', 'clamp'],
  ['น้ำมันเครื่อง', 'engine oil'],
  ['น้ำมันไฮดรอลิก', 'hydraulic oil'],
  ['สายพานลำเลียง', 'conveyor belt'],
  ['ตัวกรองอากาศ', 'ไส้กรองอากาศ', 'air filter']
];

/* ขยายคำค้น → array ของคำ (normalize แล้ว) รวมคำพ้องข้ามภาษา */
function expandQuery_(qRaw) {
  var nq = norm_(qRaw), nql = normLoose_(qRaw);
  var set = {}; set[nq] = true;
  for (var g = 0; g < SYN_GROUPS.length; g++) {
    var grp = SYN_GROUPS[g], hit = false;
    for (var t = 0; t < grp.length; t++) {
      var nt = norm_(grp[t]); if (!nt) continue;
      if (nt === nq) { hit = true; break; }
      // เป็นส่วนหนึ่งของกัน แต่ความยาวต่างกันไม่เกิน 3 (กัน "เบรกเกอร์" ไปจับ "เบรก")
      if (nq.length >= 3 && Math.abs(nt.length - nq.length) <= 3 &&
          (nt.indexOf(nq) !== -1 || nq.indexOf(nt) !== -1)) { hit = true; break; }
      // เดาคำพิมพ์ผิดของคำในพจนานุกรม (ไม่สนวรรณยุกต์) เช่น "ปากา"→ปากกา, "นอต"→น็อต
      var ntl = stripTones_(nt);
      if (nql.length >= 3 && Math.abs(ntl.length - nql.length) <= 2) {
        var sim = 1 - lev_(nql, ntl) / Math.max(nql.length, ntl.length);
        if (sim >= 0.72) { hit = true; break; }
      }
    }
    if (hit) for (var k = 0; k < grp.length; k++) { var v = norm_(grp[k]); if (v) set[v] = true; }
  }
  return Object.keys(set);
}

/* ให้คะแนนแบบ ตรงตัว/เป็นส่วนหนึ่ง กับหลายคำค้น (variants) */
function scoreExact_(cat, variants) {
  var hits = [];
  for (var i = 0; i < cat.length; i++) {
    var it = cat[i], toks = it._tk || [], sc = 0;
    for (var v = 0; v < variants.length; v++) {
      var q = variants[v]; if (!q) continue;
      var s = 0;
      if (it._nc === q || it._nn === q) s = 100;               // ตรงทั้งชื่อ/รหัส
      else if (it._nc.indexOf(q) !== -1) s = 95;               // รหัสมีคำนี้
      else {
        for (var t = 0; t < toks.length; t++) {                // เทียบระดับคำ (แม่นกว่า substring กลางคำ)
          var tk = toks[t];
          if (tk === q) { if (s < 92) s = 92; }                // เป็นคำเต็มในชื่อ (เช่น "pen" ใน BALLPOINT PEN)
          else if (q.length >= 3 && tk.indexOf(q) === 0) { if (s < 84) s = 84; } // คำในชื่อขึ้นต้นด้วยคำค้น
        }
        if (s === 0 && q.length >= 5 && it._nn.indexOf(q) !== -1) s = 74; // คำหลายพยางค์ ยอม substring (safety glasses ฯลฯ)
      }
      if (s > sc) sc = s;
    }
    if (sc > 0) hits.push([sc, it]);
  }
  hits.sort(function (a, b) { return b[0] - a[0]; });
  return hits;
}

/* เดาคำพิมพ์ผิด (fuzzy) กับคำเดียว */
function scoreFuzzy_(cat, q) {
  var ql = stripTones_(q), qlen = ql.length;
  var best = [];
  for (var i = 0; i < cat.length; i++) {
    var it = cat[i], toks = it._tk || [], localBest = 0;
    for (var ti = 0; ti < toks.length; ti++) {
      var tk = stripTones_(toks[ti]); if (!tk) continue;
      var cand = (tk.length > qlen + 1) ? tk.substring(0, qlen) : tk;
      if (Math.abs(cand.length - qlen) > 3) continue;
      var sim = 1 - lev_(ql, cand) / Math.max(qlen, cand.length);
      if (sim > localBest) localBest = sim;
    }
    if (localBest >= 0.6) best.push([localBest, it]);
  }
  best.sort(function (a, b) { return b[0] - a[0]; });
  return best;
}

/* ตั้งค่า AI ใน Script Properties (Project Settings → Script Properties)
 *   AI_API_KEY  = คีย์ของผู้ให้บริการ (ต้องมี ไม่งั้นข้าม AI)
 *   AI_API_URL  = endpoint (ค่าเริ่มต้น Groq) — เปลี่ยนได้ตามผู้ให้บริการ
 *   AI_MODEL    = ชื่อโมเดล (ค่าเริ่มต้น llama-3.1-8b-instant)
 * รองรับทุกเจ้าที่ใช้รูปแบบ OpenAI /chat/completions:
 *   Groq        : https://api.groq.com/openai/v1/chat/completions   | llama-3.3-70b-versatile
 *   OpenAI      : https://api.openai.com/v1/chat/completions        | gpt-4o-mini
 *   OpenRouter  : https://openrouter.ai/api/v1/chat/completions     | (หลายโมเดล)
 *   Together    : https://api.together.xyz/v1/chat/completions      | ...
 */
function aiInterpret_(qRaw) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('AI_API_KEY') || props.getProperty('GROQ_API_KEY');
  if (!key) return [];   // ไม่มีคีย์ = ไม่เรียก API เลย (ใช้พจนานุกรม+เดาคำต่อได้)
  // แคชผลแปล 6 ชม. — คำเดิมที่คนค้นซ้ำจะไม่เรียก API อีก (ประหยัดโควต้า)
  var cache = CacheService.getScriptCache();
  var ckey = 'ai:' + norm_(qRaw);
  var cached = cache.get(ckey);
  if (cached !== null) return cached ? cached.split('|') : [];
  var url = props.getProperty('AI_API_URL') || 'https://api.groq.com/openai/v1/chat/completions';
  var model = props.getProperty('AI_MODEL') || 'llama-3.1-8b-instant';
  var out = [];
  try {
    var payload = {
      model: model,
      temperature: 0, max_tokens: 60,
      messages: [
        { role: 'system', content:
          'You are a search assistant for a factory warehouse / spare-parts catalog (sugar mill). ' +
          'Item names in the database are mostly ENGLISH technical terms (e.g. BALLPOINT PEN, BEARING, VALVE, GREASE, BATTERY). ' +
          'The user query may be in THAI, LAO, or English, and may be vague or misspelled. ' +
          'First understand the real MEANING of the item, then output the ENGLISH keyword(s) most likely to appear in the item name. ' +
          'Lao examples: "ຈາລະບີ"=grease; "ຖ່ານໄຟ"=battery; "ຢາງ"=rubber,tire; "ນ້ຳມັນ"=oil; "ສາຍໄຟ"=wire; "ໝວກ"=helmet; "ຖົງມື"=glove; "ໄຂຄວງ"=screwdriver. ' +
          'Reply with ONLY 1-5 comma-separated English keywords, no explanation.' },
        { role: 'user', content: String(qRaw) }
      ]
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    var text = (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    out = text.split(/[,\n]/).map(function (s) { return s.trim(); }).filter(Boolean).slice(0, 5);
  } catch (e) { out = []; }
  try { cache.put(ckey, out.join('|'), 21600); } catch (e) {}   // เก็บผล 6 ชม. (รวมผลว่าง กันเรียกซ้ำ)
  return out;
}

/* ---------- ค้นหา: พจนานุกรม → เดาคำ → Groq ---------- */
function searchCatalog_(qRaw, limit) {
  var q = norm_(qRaw);
  if (!q) return { mode: 'empty', items: [] };
  limit = limit || DEFAULT_LIMIT;
  var cat = readCatalog_().items;

  // 1) พจนานุกรมคำพ้อง (ข้ามภาษา) + ตรงตัว/เป็นส่วนหนึ่ง
  var variants = expandQuery_(qRaw);
  var hits = scoreExact_(cat, variants);
  if (hits.length) return { mode: 'match', items: hits.slice(0, limit).map(pub_) };

  // 2) เดาคำพิมพ์ผิด
  var fz = scoreFuzzy_(cat, q);
  if (fz.length) return { mode: 'fuzzy', items: fz.slice(0, limit).map(pub_) };

  // 3) ให้ AI แปลคำ แล้วค้นซ้ำ (ถ้าตั้ง key ไว้)
  var ai = aiInterpret_(qRaw);
  if (ai.length) {
    var av = [];
    for (var a = 0; a < ai.length; a++) av.push(norm_(ai[a]));
    var ah = scoreExact_(cat, av);
    if (!ah.length) {
      for (var b = 0; b < av.length; b++) ah = ah.concat(scoreFuzzy_(cat, av[b]));
      ah.sort(function (x, y) { return y[0] - x[0]; });
    }
    if (ah.length) return { mode: 'ai', ai: ai.join(', '), items: ah.slice(0, limit).map(pub_) };
  }

  return { mode: 'none', items: [] };
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
      return json_({ ok: true, mode: res.mode, ai: res.ai || '', count: res.items.length, items: res.items, q: p.q || '' });
    }
    if (action === 'sample') {
      return json_({ ok: true, items: sample_(parseInt(p.n) || 6) });
    }
    if (action === 'aidiag') {
      var pr = PropertiesService.getScriptProperties();
      var k = pr.getProperty('AI_API_KEY') || pr.getProperty('GROQ_API_KEY');
      var aurl = pr.getProperty('AI_API_URL') || 'https://api.groq.com/openai/v1/chat/completions';
      var amodel = pr.getProperty('AI_MODEL') || 'llama-3.1-8b-instant';
      if (!k) return json_({ ok: true, hasKey: false, note: 'ไม่พบ AI_API_KEY / GROQ_API_KEY ใน Script Properties' });
      var o = { ok: true, hasKey: true, keyPrefix: String(k).substring(0, 5), model: amodel, url: aurl };
      try {
        var rr = UrlFetchApp.fetch(aurl, {
          method: 'post', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + k },
          payload: JSON.stringify({ model: amodel, temperature: 0, max_tokens: 40, messages: [{ role: 'user', content: 'reply with one english keyword: มีดโกน' }] }),
          muteHttpExceptions: true
        });
        o.httpStatus = rr.getResponseCode();
        o.body = rr.getContentText().substring(0, 500);
      } catch (err) { o.fetchError = String((err && err.message) || err); }
      return json_(o);
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

  if (map.code != null) sh.getRange(1, map.code + 1, values.length, 1).setNumberFormat('@');   // คอลัมน์รหัสเป็น Text เสมอ
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
    rows.push([String(it.code), it.name || '', it.group || '', it.unit || '', it.qty || 0, it.sloc || '', it.slocName || '', it.plant || '', it.plantName || '']);
  });
  // บังคับคอลัมน์รหัส (A) เป็น Text ก่อนเขียน → เลขยาวไม่ถูกย่อเป็น 1.01E+15 และไม่เสียหลัก
  sh.getRange(1, 1, Math.max(rows.length, 2), 1).setNumberFormat('@');
  sh.getRange(1, 1, rows.length, HEAD.length).setValues(rows);
  sh.setFrozenRows(1);
  return { ok: true, count: rows.length - 1 };
}

/* ================= เมนูช่วยเหลือในหน้า Sheet ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 คลังพัสดุ')
    .addItem('ตรวจสอบคอลัมน์ที่ตรวจจับได้', 'checkColumns')
    .addItem('นับจำนวนรายการ (unique Material)', 'countItems')
    .addItem('ซ่อมรหัสให้แสดงเลขเต็ม (แก้ 1.01E+15)', 'fixCodeColumn')
    .addToUi();
}

/* ซ่อมคอลัมน์รหัส: แปลงตัวเลขที่ถูกย่อ (1.01E+15) กลับเป็นข้อความเลขเต็มทุกหลัก */
function fixCodeColumn() {
  var sh = dataSheet_();
  var last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('ไม่มีข้อมูล'); return; }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = buildMap_(headers);
  if (map.code == null) { SpreadsheetApp.getUi().alert('ไม่พบคอลัมน์ Material'); return; }
  var colIdx = map.code + 1;
  var rng = sh.getRange(2, colIdx, last - 1, 1);
  var vals = rng.getValues();
  var fixed = 0;
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    if (typeof v === 'number') { vals[i][0] = v.toFixed(0); fixed++; }   // เลข → ข้อความเลขเต็ม ไม่มี E+
    else vals[i][0] = String(v);
  }
  sh.getRange(1, colIdx, last, 1).setNumberFormat('@');   // ตั้งคอลัมน์เป็น Text ถาวร
  rng.setValues(vals);
  SpreadsheetApp.getUi().alert('ซ่อมเรียบร้อย ✅\nแปลงรหัสที่เป็นตัวเลข ' + fixed + ' แถว ให้เป็นข้อความเลขเต็มแล้ว');
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
