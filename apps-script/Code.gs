/* ระบบเบิกพัสดุห้องยา — แผนกเภสัชกรรม รพ.จุฬารัตน์ 9 แอร์พอร์ต
 * Backend: Google Apps Script + Google Sheets (JSON API)
 * Design & Developed by ภก.วทัญญู โตเจริญนิรัติศัย
 * ใช้งาน: รัน setup() หนึ่งครั้ง แล้ว Deploy > Web app (Execute as: Me, Access: Anyone)
 */
var COLS = ['staff', 'items', 'units', 'requests', 'cycles', 'locks'];
var ADMIN_ONLY = { staff: 1, items: 1, units: 1, cycles: 1, locks: 1 };
var DEFAULT_PW = '12345678';
var TOKEN_MS = 12 * 3600 * 1000;

function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.openById(props_().getProperty('SHEET_ID')); }

/* ---------- setup (รันครั้งเดียวจากตัวแก้ไข) ---------- */
function setup() {
  var p = props_(), ss;
  if (!p.getProperty('SHEET_ID')) {
    ss = SpreadsheetApp.create('ระบบเบิกพัสดุห้องยา - ฐานข้อมูล');
    p.setProperty('SHEET_ID', ss.getId());
  } else ss = ss_();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  var meta = ss.getSheetByName('_meta') || ss.insertSheet('_meta');
  if (meta.getLastRow() < 1) meta.getRange(1, 1, 1, 2).setValues([['collection', 'ver']]);
  COLS.forEach(function (c, i) {
    var sh = ss.getSheetByName(c) || ss.insertSheet(c);
    if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, 3).setValues([['id', 'json', 'ver']]);
    var mr = findRow_(meta, c);
    if (!mr) meta.appendRow([c, 1]);
  });
  var first = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (first) ss.deleteSheet(first);
  var seed = SEED_();
  ['staff', 'items', 'units'].forEach(function (c) {
    var sh = ss.getSheetByName(c);
    if (sh.getLastRow() > 1) return;
    var rows = seed[c].map(function (d) { return [d.id, JSON.stringify(d.data), 1]; });
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
    bump_(c);
  });
  Logger.log('Sheet: ' + ss.getUrl());
}

/* ---------- helpers ---------- */
function findRow_(sh, key) {
  var n = sh.getLastRow(); if (n < 1) return 0;
  var v = sh.getRange(1, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(v[i][0]) === key) return i + 1;
  return 0;
}
function ver_(c) {
  var m = ss_().getSheetByName('_meta'), r = findRow_(m, c);
  return r ? Number(m.getRange(r, 2).getValue()) : 0;
}
function bump_(c) {
  var m = ss_().getSheetByName('_meta'), r = findRow_(m, c);
  if (!r) { m.appendRow([c, 1]); return 1; }
  var v = Number(m.getRange(r, 2).getValue()) + 1;
  m.getRange(r, 2).setValue(v); return v;
}
function readAll_(c) {
  var sh = ss_().getSheetByName(c), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getValues().map(function (r) {
    var d = {}; try { d = JSON.parse(r[1]); } catch (e) {}
    d.__id = String(r[0]); return d;
  });
}
function getDoc_(c, id) {
  var sh = ss_().getSheetByName(c), r = findRow_(sh, id);
  if (!r || r === 1) return null;
  try { return JSON.parse(sh.getRange(r, 2).getValue()); } catch (e) { return null; }
}
function putDoc_(c, id, data) {
  var sh = ss_().getSheetByName(c), r = findRow_(sh, id), j = JSON.stringify(data);
  if (j.length > 49000) throw new Error('too_large');
  if (r > 1) sh.getRange(r, 2, 1, 2).setValues([[j, Date.now()]]);
  else sh.appendRow([id, j, Date.now()]);
}
function delDoc_(c, id) {
  var sh = ss_().getSheetByName(c), r = findRow_(sh, id);
  if (r > 1) sh.deleteRow(r);
}
function hex_(bytes) {
  return bytes.map(function (b) { return ((b < 0 ? b + 256 : b) + 256).toString(16).slice(1); }).join('');
}
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function fp_(h) { return h ? sha_('fp|' + props_().getProperty('SECRET') + '|' + h).slice(0, 12) : null; }
function sign_(s) {
  return hex_(Utilities.computeHmacSha256Signature(s, props_().getProperty('SECRET')));
}
function pub_(id, d) { // ตัดรหัสผ่านออกก่อนส่งให้เบราว์เซอร์
  var o = {}; for (var k in d) if (k !== 'passHash' && k !== '__id') o[k] = d[k];
  o.id = id; o.passHash = fp_(d.passHash || null); return o;
}
function mkToken_(st) {
  var ts = Date.now(), id = String(st.empId);
  return ts + '.' + id + '.' + sign_(ts + '|' + id + '|' + (st.passHash || 'default'));
}
function auth_(token) {
  var t = String(token || '').split('.');
  if (t.length !== 3) return null;
  if (Date.now() - Number(t[0]) > TOKEN_MS) return null;
  var st = getDoc_('staff', t[1]);
  if (!st) return null;
  if (sign_(t[0] + '|' + t[1] + '|' + (st.passHash || 'default')) !== t[2]) return null;
  st.empId = String(st.empId || t[1]); return st;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- entry ---------- */
function doGet() { return out_({ ok: true, name: 'ch9-supply', time: Date.now() }); }
function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, code: 'bad_request' }); }
  try {
    if (req.a === 'login') return out_(login_(req));
    var me = auth_(req.token);
    if (!me) return out_({ ok: false, code: 'auth' });
    if (req.a === 'pull') return out_(pull_(req));
    if (req.a === 'changePw') return out_(withLock_(function () { return changePw_(me, req); }));
    if (req.a === 'write') return out_(withLock_(function () { return write_(me, req); }));
    if (req.a === 'acquire') return out_(withLock_(function () { return acquire_(me, req); }));
    return out_({ ok: false, code: 'bad_request' });
  } catch (err) {
    return out_({ ok: false, code: String(err && err.message || err) });
  }
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- actions ---------- */
function login_(req) {
  var id = String(req.id || '').replace(/\D/g, ''), h = String(req.h || '');
  var cache = CacheService.getScriptCache(), k = 'fail:' + id, n = Number(cache.get(k) || 0);
  if (n >= 8) return { ok: false, code: 'locked' };
  var st = getDoc_('staff', id);
  var good = false;
  if (st) good = st.passHash ? st.passHash === h : h === sha_('ch9sup|' + id + '|' + DEFAULT_PW);
  if (!good) { cache.put(k, String(n + 1), 600); return { ok: false, code: st ? 'bad_pw' : 'no_user' }; }
  cache.remove(k);
  st.empId = id;
  return { ok: true, token: mkToken_(st), user: pub_(id, st) };
}
function changePw_(me, req) {
  var h = String(req.h || '');
  if (!/^[0-9a-f]{64}$/.test(h)) return { ok: false, code: 'bad_request' };
  var st = getDoc_('staff', me.empId);
  st.passHash = h; st.pwChangedAt = Date.now();
  putDoc_('staff', me.empId, st); bump_('staff');
  return { ok: true, token: mkToken_(st), user: pub_(me.empId, st) };
}
function pull_(req) {
  var vers = req.vers || {}, cols = {};
  COLS.forEach(function (c) {
    var v = ver_(c);
    if (vers[c] === v) return;
    var docs = readAll_(c).map(function (d) {
      var id = d.__id; delete d.__id;
      return c === 'staff' ? pub_(id, d) : Object.assign({ id: id }, d);
    });
    cols[c] = { ver: v, docs: docs };
  });
  return { ok: true, cols: cols, now: Date.now() };
}
function write_(me, req) {
  var ops = req.ops || [], touched = {};
  ops.forEach(function (o) {
    var parts = String(o.path || '').split('/'), c = parts[0], id = parts[1];
    if (COLS.indexOf(c) < 0 || !id || parts.length !== 2) throw new Error('bad_path');
    if (ADMIN_ONLY[c] && !me.isAdmin) throw new Error('forbidden');
    var cur = getDoc_(c, id), data = o.data || {};
    if (o.op === 'delete') {
      if (c === 'staff' && cur && cur.isAdmin) {
        var admins = readAll_('staff').filter(function (d) { return d.isAdmin && d.__id !== id; });
        if (!admins.length) throw new Error('last_admin');
      }
      delDoc_(c, id);
    } else {
      var nd;
      if (o.op === 'update') { if (!cur) throw new Error('not_found'); nd = Object.assign({}, cur, data); }
      else nd = Object.assign({}, data);
      if (c === 'staff') {
        var keep = cur ? cur.passHash : null;
        if (o.op === 'update' && data.hasOwnProperty('passHash') && data.passHash === null) nd.passHash = null; // รีเซ็ต
        else nd.passHash = keep || null;
        if (o.op === 'set' && !cur) nd.passHash = null;
        if (id === me.empId && !nd.isAdmin && me.isAdmin) {
          var others = readAll_('staff').filter(function (d) { return d.isAdmin && d.__id !== id; });
          if (!others.length) throw new Error('last_admin');
        }
        nd.empId = id;
      }
      putDoc_(c, id, nd);
    }
    touched[c] = 1;
  });
  Object.keys(touched).forEach(bump_);
  return { ok: true };
}
function acquire_(me, req) {
  if (!me.isAdmin) throw new Error('forbidden');
  var id = String(req.id || ''), holder = String(req.holder || me.empId), ttl = Math.min(Number(req.ttlMs) || 20000, 120000);
  if (!id) throw new Error('bad_path');
  var cur = getDoc_('locks', id), now = Date.now();
  if (cur && cur.until > now && cur.holderId !== holder) return { ok: true, acquired: false };
  var d = Object.assign({}, req.data || {}, { holderId: holder, until: now + ttl });
  putDoc_('locks', id, d); bump_('locks');
  return { ok: true, acquired: true };
}

/* ---------- ข้อมูลตั้งต้น ----------
 * ใส่แอดมินคนแรกไว้เป็นตัวอย่าง (รหัส 12345678) ให้แก้รหัส/ชื่อให้ตรงก่อนรัน setup()
 * พนักงานที่เหลือเพิ่มได้ในแอป: ตั้งค่า > พนักงาน
 */
function SEED_() { return {"staff":[{"id":"12345678","data":{"empId":"12345678","first":"ตัวอย่าง","last":"แอดมิน","isAdmin":true,"passHash":null,"createdAt":0}}],"items":[{"id":"i001","data":{"name":"กระดาษถ่ายเอกสาร A4 70แกรม","cat":"กระดาษ","order":1}},{"id":"i002","data":{"name":"กระดาษการ์ดสี สีเหลือง 180 แกรม (200แผ่น/รีม)","cat":"กระดาษ","order":2}},{"id":"i003","data":{"name":"กระดาษการ์ดสีเขียว 180แกรม (250 แผ่น/รีม)","cat":"กระดาษ","order":3}},{"id":"i004","data":{"name":"กระดาษการ์ดสี สีฟ้า 180 แกรม (200แผ่น/รีม)","cat":"กระดาษ","order":4}},{"id":"i005","data":{"name":"กระดาษการ์ดสี สีขาว 180 แกรม (200แผ่น/รีม)","cat":"กระดาษ","order":5}},{"id":"i006","data":{"name":"ปากกาหมึกแห้งสีน้ำเงิน","cat":"เครื่องเขียน","order":6}},{"id":"i007","data":{"name":"ถุงซิปใส18x23ซม","cat":"ถุง","order":7}},{"id":"i008","data":{"name":"ที่เจาะกระดาษ No.100XL","cat":"อื่นๆ","order":8}},{"id":"i009","data":{"name":"ถุงร้อน6x9นิ้ว","cat":"ถุง","order":9}},{"id":"i010","data":{"name":"ตัวคัตเตอร์ใหญ่","cat":"อื่นๆ","order":10}},{"id":"i011","data":{"name":"ถ่านไฟฉายจิ๋ว1.5V.รุ่นAAA Panasonic (แพ็ค 4ก้อน)","cat":"อื่นๆ","order":11}},{"id":"i012","data":{"name":"หนังยางวงใหญ่","cat":"อื่นๆ","order":12}},{"id":"i013","data":{"name":"ปากกาเน้นข้อความ สีเหลือง","cat":"เครื่องเขียน","order":13}},{"id":"i014","data":{"name":"ตัวแม็กซ์ No.35","cat":"อื่นๆ","order":14}},{"id":"i015","data":{"name":"ซองเอกสารน้ำตาล 11x17 แพ็ค 50 ซอง","cat":"แฟ้ม/ซอง/สมุด","order":15}},{"id":"i016","data":{"name":"น้ำหมึกประทับตรายาง สีแดง(28ซีซี.)","cat":"เครื่องเขียน","order":16}},{"id":"i017","data":{"name":"แฟ้ม 2 ห่วงตราช้างสีน้ำเงิน A4 No.420","cat":"แฟ้ม/ซอง/สมุด","order":17}},{"id":"i018","data":{"name":"แผ่นสติ๊กเกอร์ใส (PVC 53cm x 70cm)","cat":"เทป/สติ๊กเกอร์","order":18}},{"id":"i019","data":{"name":"แผ่นสติ๊กเกอร์สี คละสี (PVC 53cm x 70cm)","cat":"เทป/สติ๊กเกอร์","order":19}},{"id":"i020","data":{"name":"ถุงกระดาษใส่ยา (ใบใหญ่)(กระดาษคราฟน้ำตาล)","cat":"ถุง","order":20}},{"id":"i021","data":{"name":"แท่นตัดเทป No.25","cat":"เทป/สติ๊กเกอร์","order":21}},{"id":"i022","data":{"name":"ป้ายสติ๊กเกอร์A5(13x38มม.)","cat":"เทป/สติ๊กเกอร์","order":22}},{"id":"i023","data":{"name":"ใบมีดคัตเตอร์ใหญ่","cat":"อื่นๆ","order":23}},{"id":"i024","data":{"name":"ใบมีดคัตเตอร์เล็ก","cat":"อื่นๆ","order":24}},{"id":"i025","data":{"name":"ตัวคัตเตอร์เล็ก","cat":"อื่นๆ","order":25}},{"id":"i026","data":{"name":"แท่นตัดเทปใสใหญ่ No.50","cat":"เทป/สติ๊กเกอร์","order":26}},{"id":"i027","data":{"name":"สมุดบันทึกขนาด5/100(กลาง)","cat":"แฟ้ม/ซอง/สมุด","order":27}},{"id":"i028","data":{"name":"สมุดปกอ่อน","cat":"แฟ้ม/ซอง/สมุด","order":28}},{"id":"i029","data":{"name":"ไส้แฟ้มโชว์A4(แบบ11รู)","cat":"แฟ้ม/ซอง/สมุด","order":29}},{"id":"i030","data":{"name":"แฟ้มสันกว้างNo.120F F4","cat":"แฟ้ม/ซอง/สมุด","order":30}},{"id":"i031","data":{"name":"ไม้บรรทัดพลาสติก12นิ้ว","cat":"เครื่องเขียน","order":31}},{"id":"i032","data":{"name":"ยางลบดินสอ","cat":"เครื่องเขียน","order":32}},{"id":"i033","data":{"name":"ดินสอ","cat":"เครื่องเขียน","order":33}},{"id":"i034","data":{"name":"ปากกาหมึกแห้งสีแดง","cat":"เครื่องเขียน","order":34}},{"id":"i035","data":{"name":"ปากกาหมึกแห้งสีดำ","cat":"เครื่องเขียน","order":35}},{"id":"i036","data":{"name":"ปากกาเน้นข้อความสีส้ม","cat":"เครื่องเขียน","order":36}},{"id":"i037","data":{"name":"ปากกาเน้นข้อความสีชมพู","cat":"เครื่องเขียน","order":37}},{"id":"i038","data":{"name":"ปากกาเน้นข้อความสีเขียว","cat":"เครื่องเขียน","order":38}},{"id":"i039","data":{"name":"เปเปอร์คลิบNo.1","cat":"อื่นๆ","order":39}},{"id":"i040","data":{"name":"กรรไกร 8 นิ้ว","cat":"อื่นๆ","order":40}},{"id":"i041","data":{"name":"น้ำหมึกประทับตรายางสีน้ำเงิน(28ซีซี.)","cat":"เครื่องเขียน","order":41}},{"id":"i042","data":{"name":"ไส้แม็กซ์No.35","cat":"อื่นๆ","order":42}},{"id":"i043","data":{"name":"ไส้แม็กซ์No.10","cat":"อื่นๆ","order":43}},{"id":"i044","data":{"name":"ตัวแม็กซ์No.10","cat":"อื่นๆ","order":44}},{"id":"i045","data":{"name":"เทปกาว2หน้าเยื่อเทป(บาง)1/2\"x20หลา","cat":"เทป/สติ๊กเกอร์","order":45}},{"id":"i046","data":{"name":"เทปใส(กว้าง1/2นิ้วxยาว72หลา)","cat":"เทป/สติ๊กเกอร์","order":46}},{"id":"i047","data":{"name":"โอพีพีเทปสีน้ำตาล(กว้าง2นิ้วxยาว45หลา)","cat":"เทป/สติ๊กเกอร์","order":47}},{"id":"i048","data":{"name":"เทปกาวย่น(1นิ้วx40หลา)","cat":"เทป/สติ๊กเกอร์","order":48}},{"id":"i049","data":{"name":"เทปกาว2หน้าโฟม(หนา)21mm.x50cm.","cat":"เทป/สติ๊กเกอร์","order":49}},{"id":"i050","data":{"name":"กระดาษคาร์บอนสีน้ำเงิน","cat":"กระดาษ","order":50}},{"id":"i051","data":{"name":"กระดาษแบงค์สีชมพูปริ้นใบสำคัญรับ(55แกรม)","cat":"กระดาษ","order":51}},{"id":"i052","data":{"name":"ถ่านไฟฉายอัลคาไลน์(รุ่นAA) Panasonic (48 ก้อน/กล่อง)","cat":"อื่นๆ","order":52}},{"id":"i053","data":{"name":"ปากกาเมจิกสีแดง","cat":"เครื่องเขียน","order":53}},{"id":"i054","data":{"name":"กระดาษ A5(500แผ่น/รีม)","cat":"กระดาษ","order":54}},{"id":"i055","data":{"name":"กระดาษการ์ดสีม่วง 80 แกรม (500แผ่น/รีม)","cat":"กระดาษ","order":55}},{"id":"i056","data":{"name":"ลิควิดเปเปอร์","cat":"เครื่องเขียน","order":56}},{"id":"i057","data":{"name":"เทปใสแกนใหญ่ 3 นิ้ว 3/4\" 18x33 36 หลา","cat":"เทป/สติ๊กเกอร์","order":57}},{"id":"i058","data":{"name":"ถุงซิปสีน้ำตาล4x5นิ้ว(10x12.5ซม.)","cat":"ถุง","order":58}},{"id":"i059","data":{"name":"กระดาษชำระม้วนเล็ก","cat":"กระดาษ","order":59}},{"id":"i060","data":{"name":"คลิปบอร์ดหนังเทียม A4","cat":"แฟ้ม/ซอง/สมุด","order":60}},{"id":"i061","data":{"name":"กระดาษโปสเตอร์บาง 2 หน้า","cat":"กระดาษ","order":61}},{"id":"i062","data":{"name":"ถุงกระดาษใส่ยา(ใบเล็ก) กระดาษคราฟท์น้ำตาล","cat":"ถุง","order":62}},{"id":"i063","data":{"name":"ถ่านอัลคาไลน์ AAA PANASONIC","cat":"อื่นๆ","order":63}},{"id":"i064","data":{"name":"ถุงซิปสีน้ำตาล8.5x10ซม.","cat":"ถุง","order":64}},{"id":"i065","data":{"name":"ถุงซิปใส 8.5x10ซม","cat":"ถุง","order":65}},{"id":"i066","data":{"name":"กระดาษสติ๊กเกอร์ ขาวด้าน A4","cat":"กระดาษ","order":66}},{"id":"i067","data":{"name":"ซองเอกสารน้ำตาล A4 9x12(1 แพ็ค 50 ซอง)","cat":"แฟ้ม/ซอง/สมุด","order":67}},{"id":"i068","data":{"name":"Logitech Keyboard _K120","cat":"อื่นๆ","order":68}},{"id":"i069","data":{"name":"Logitech M100R Mouse","cat":"อื่นๆ","order":69}},{"id":"i070","data":{"name":"ถุงซิปน้ำตาล 28x36ซม.","cat":"ถุง","order":70}},{"id":"i071","data":{"name":"เทปใสแกนเล็ก 1 นิ้ว 3/4\"18x20 มม 36 หลา","cat":"เทป/สติ๊กเกอร์","order":71}},{"id":"i072","data":{"name":"แท่นตัดเทปเล็ก No.15","cat":"เทป/สติ๊กเกอร์","order":72}},{"id":"i073","data":{"name":"ถุงซิปใส 28x35ซม.","cat":"ถุง","order":73}},{"id":"i074","data":{"name":"ถุงซิปน้ำตาล 12.5x20.5ซม.","cat":"ถุง","order":74}},{"id":"i075","data":{"name":"ถุงซิปใส10x12ซม.","cat":"ถุง","order":75}},{"id":"i076","data":{"name":"กระดาษ Thermal ขนาด 80x80 mm (บัตรคิว)","cat":"กระดาษ","order":76}},{"id":"i077","data":{"name":"ที่ถอนลวดเย็บก้ามปู","cat":"อื่นๆ","order":77}},{"id":"i078","data":{"name":"แฟ้มแขวนสีเขียวแก่F4","cat":"แฟ้ม/ซอง/สมุด","order":78}},{"id":"i079","data":{"name":"ถุงร้อน10x15นิ้ว","cat":"ถุง","order":79}},{"id":"i080","data":{"name":"ถุงซิปใส13x18ซม.","cat":"ถุง","order":80}},{"id":"i081","data":{"name":"ถุงซิปใส11x15ซม.","cat":"ถุง","order":81}},{"id":"i082","data":{"name":"ถุงซิปน้ำตาล 10x15ซม","cat":"ถุง","order":82}},{"id":"i083","data":{"name":"ปากกาเคมี 2 หัว Horse น้ำเงิน","cat":"เครื่องเขียน","order":83}},{"id":"i084","data":{"name":"ถุงซิปใส 6x8ซม.","cat":"ถุง","order":84}},{"id":"i085","data":{"name":"ปากกาเคมี 2 หัว Horse ดำ","cat":"เครื่องเขียน","order":85}},{"id":"i086","data":{"name":"แผ่นฟิวเจอร์บอร์ด(ลูกฟูก) 3mm. 65x80cm.","cat":"อื่นๆ","order":86}},{"id":"i087","data":{"name":"แฟ้มตราช้าง No.925 สีเขียว","cat":"แฟ้ม/ซอง/สมุด","order":87}},{"id":"i088","data":{"name":"ถุงซิปใส 12x18ซม.","cat":"ถุง","order":88}},{"id":"i089","data":{"name":"แฟ้มตราช้าง No.926 สีเขียว","cat":"แฟ้ม/ซอง/สมุด","order":89}},{"id":"i090","data":{"name":"เคเบิ้ลไทด์ ขนาด 3.6x150 มม. 1:100 เส้น สีขาว","cat":"อื่นๆ","order":90}},{"id":"i091","data":{"name":"สติ๊กเกอร์กลุ่มโรงพยาบาลจุฬารัตน์ สีแดง","cat":"เทป/สติ๊กเกอร์","order":91}},{"id":"i092","data":{"name":"หมึก RIBBON ARMOR AWR8 102*360","cat":"อื่นๆ","order":92}},{"id":"i093","data":{"name":"แฟ้มตราช้าง No.905 สีเขียว","cat":"แฟ้ม/ซอง/สมุด","order":93}},{"id":"i094","data":{"name":"แฟ้มตราช้าง No.925 สีเหลือง","cat":"แฟ้ม/ซอง/สมุด","order":94}},{"id":"i095","data":{"name":"แฟ้มตราช้าง No.926 สีชมพู","cat":"แฟ้ม/ซอง/สมุด","order":95}},{"id":"i096","data":{"name":"แฟ้มตราช้าง No.926 สีม่วง","cat":"แฟ้ม/ซอง/สมุด","order":96}},{"id":"i097","data":{"name":"แฟ้มตราช้าง No.905 สีส้ม","cat":"แฟ้ม/ซอง/สมุด","order":97}},{"id":"i098","data":{"name":"แฟ้มตราช้าง No.925 สีแดง","cat":"แฟ้ม/ซอง/สมุด","order":98}},{"id":"i099","data":{"name":"สติ๊กเกอร์ติดซองยา Label Sticker Art W8*L6","cat":"เทป/สติ๊กเกอร์","order":99}},{"id":"i100","data":{"name":"ถุงร้อน 12x14 นิ้ว","cat":"ถุง","order":100}},{"id":"i101","data":{"name":"เคเบิ้ลไทด์ ขนาด 4.8x300 มม. 1:100 เส้น สีขาว","cat":"อื่นๆ","order":101}},{"id":"i102","data":{"name":"แฟ้มตราช้าง No.925 สีฟ้า","cat":"แฟ้ม/ซอง/สมุด","order":102}},{"id":"i103","data":{"name":"สติ๊กเกอร์ BARCODE 8.5 x 4.5cm.","cat":"เทป/สติ๊กเกอร์","order":103}},{"id":"i104","data":{"name":"แฟ้มตราช้าง No.926 สีฟ้า","cat":"แฟ้ม/ซอง/สมุด","order":104}},{"id":"i105","data":{"name":"ถุงร้อน 14x18 นิ้ว","cat":"ถุง","order":105}},{"id":"i106","data":{"name":"ถุงร้อน 16x20 นิ้ว","cat":"ถุง","order":106}},{"id":"i107","data":{"name":"ถุงซิปน้ำตาล 18x23ซม.","cat":"ถุง","order":107}},{"id":"i108","data":{"name":"ถุงใส่ศพ","cat":"ถุง","order":108}},{"id":"i109","data":{"name":"PVC STICKER OUTDOOR FUZUN ขาวใส 127cm*50M","cat":"เทป/สติ๊กเกอร์","order":109}},{"id":"i110","data":{"name":"กระดาษ Thermal (80*20)","cat":"กระดาษ","order":110}},{"id":"i111","data":{"name":"แผ่นลูกฟูกสีขาวหนา 3 มิล","cat":"อื่นๆ","order":111}}],"units":[{"id":"u1","data":{"name":"ห้องยาชั้น 1A (M1)","order":1}},{"id":"u2","data":{"name":"ห้องยาชั้น 1C (M2)","order":2}},{"id":"u3","data":{"name":"ห้องยาชั้น 2A (ประกันสังคม)","order":3}},{"id":"u4","data":{"name":"ห้องยาชั้น 2C (M4)","order":4}},{"id":"u5","data":{"name":"Sub-stock","order":5}}]}; }
