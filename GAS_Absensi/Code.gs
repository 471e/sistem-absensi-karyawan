function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : '';

  if (action === 'export') {
    var employees = getSheetRows('Employees');
    var attendance = getSheetRows('Attendance');
    var settings = getSettings();
    var result = {
      status: 'success',
      employees: employees,
      attendance: attendance,
      settings: settings,
      serverTime: new Date().toISOString()
    };
    var json = JSON.stringify(result);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
  }

  ensureAllSheets('safe');
  ensureFolder('Absensi_Photos');
  ensureFolder('Absensi_Exports');
  var t = HtmlService.createTemplateFromFile('index');
  return t.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  var parsed = {};
  try { parsed = JSON.parse(raw); } catch (err) {}
  var action = parsed && parsed.action ? String(parsed.action) : '';
  var payload = parsed && parsed.payload ? parsed.payload : null;
  var mode = parsed && parsed.mode ? String(parsed.mode) : 'replace';
  if (action === 'sync_employees') {
    writeEmployees(payload, mode);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'sync_attendance') {
    if (String(mode) === 'upsert') upsertAttendance(payload);
    else writeAttendance(payload, mode);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'sync_settings') {
    writeSettings(payload);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, reason: 'unknown-action' })).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheet(name, headers, mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  var h = Array.isArray(headers) ? headers : [];
  if (h.length > 0) {
    var doReplace = String(mode||'replace') === 'replace';
    if (doReplace) {
      sh.clearContents();
      sh.getRange(1, 1, 1, h.length).setValues([h]);
      sh.getRange(1, 1, 1, h.length).setFontWeight('bold');
    } else {
      var lastCol = Math.max(sh.getLastColumn(), h.length);
      var current = sh.getRange(1,1,1,lastCol).getValues()[0];
      var same = true;
      for (var i=0;i<h.length;i++) { if (String(current[i]||'') !== String(h[i]||'')) { same = false; break; } }
      if (!same) {
        sh.getRange(1, 1, 1, h.length).setValues([h]);
        sh.getRange(1, 1, 1, h.length).setFontWeight('bold');
      }
    }
  }
  return sh;
}

function writeEmployees(items, mode) {
  var headers = ["id","name","workUnit","role","password","gender","phone","email","address","joinDate","nik","npp","golongan","pangkat","masaKerja","birthPlace","birthDate","religion","maritalStatus","status","photoUrl"];
  var sh = ensureSheet('Employees', headers, mode);
  var arr = (Array.isArray(items) ? items : []).filter(function(e){return e && String(e.id||'').length>0;}).map(function(e){
    return [
      String(e.id||""), String(e.name||""), String(e.workUnit||""), String(e.role||""),
      String(e.password||"123"), String(e.gender||""), String(e.phone||""), String(e.email||""),
      String(e.address||""), String(e.joinDate||""), String(e.nik||""), String(e.npp||""),
      String(e.golongan||""), String(e.pangkat||""), String(e.masaKerja||""), String(e.birthPlace||""),
      String(e.birthDate||""), String(e.religion||""), String(e.maritalStatus||""), String(e.status||""),
      String(e.photoUrl||"")
    ];
  });
  if (arr.length > 0) {
    var doReplace = String(mode||'replace') === 'replace';
    if (doReplace) {
      sh.getRange(2,1,arr.length,headers.length).setValues(arr);
    } else {
      var startRow = Math.max(2, sh.getLastRow()+1);
      sh.getRange(startRow,1,arr.length,headers.length).setValues(arr);
    }
  }
}

function getHeaderIndex(sh, name) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  for (var i=0;i<headers.length;i++) {
    if (String(headers[i]).trim() === String(name).trim()) return i+1; // 1-based
  }
  return -1;
}

function upsertEmployee(emp) {
  var headers = ["id","name","workUnit","role","password","gender","phone","email","address","joinDate","nik","npp","golongan","pangkat","masaKerja","birthPlace","birthDate","religion","maritalStatus","status","photoUrl"];
  var sh = ensureSheet('Employees', headers, 'append');
  var idCol = getHeaderIndex(sh, 'id');
  if (idCol < 1) throw new Error('Header id tidak ditemukan');
  var lastRow = sh.getLastRow();
  var targetRow = -1;
  if (lastRow >= 2) {
    var ids = sh.getRange(2, idCol, lastRow-1, 1).getValues();
    for (var i=0;i<ids.length;i++) {
      if (String(ids[i][0]) === String(emp.id||'')) { targetRow = 2+i; break; }
    }
  }
  var row = [
    String(emp.id||""), String(emp.name||""), String(emp.workUnit||""), String(emp.role||""),
    String(emp.password||"123"), String(emp.gender||""), String(emp.phone||""), String(emp.email||""),
    String(emp.address||""), String(emp.joinDate||""), String(emp.nik||""), String(emp.npp||""),
    String(emp.golongan||""), String(emp.pangkat||""), String(emp.masaKerja||""), String(emp.birthPlace||""),
    String(emp.birthDate||""), String(emp.religion||""), String(emp.maritalStatus||""), String(emp.status||""),
    String(emp.photoUrl||"")
  ];
  if (targetRow > -1) {
    sh.getRange(targetRow,1,1,headers.length).setValues([row]);
    return { ok:true, updated:true, row: targetRow };
  } else {
    var startRow = Math.max(2, sh.getLastRow()+1);
    sh.getRange(startRow,1,1,headers.length).setValues([row]);
    return { ok:true, inserted:true, row: startRow };
  }
}

function bulkUpsertEmployees(items) {
  if (!Array.isArray(items) || items.length === 0) return { ok:true, count:0 };
  var headers = ["id","name","workUnit","role","password","gender","phone","email","address","joinDate","nik","npp","golongan","pangkat","masaKerja","birthPlace","birthDate","religion","maritalStatus","status","photoUrl"];
  var sh = ensureSheet('Employees', headers, 'append');
  var idCol = getHeaderIndex(sh, 'id');
  if (idCol < 1) throw new Error('Header id tidak ditemukan');
  
  var lastRow = sh.getLastRow();
  var idMap = {}; 
  var existingData = [];
  
  if (lastRow >= 2) {
    existingData = sh.getRange(2, 1, lastRow-1, headers.length).getValues();
    var idIdx = idCol - 1; 
    for (var i=0; i<existingData.length; i++) {
      var eid = String(existingData[i][idIdx]||'');
      if (eid) idMap[eid] = i;
    }
  }
  
  var newRows = [];
  var updatesCount = 0;
  var insertsCount = 0;
  
  items.forEach(function(emp){
    var eid = String(emp.id||'');
    if (!eid) return;
    
    if (idMap.hasOwnProperty(eid)) {
      // Update existing - support partial update
      var idx = idMap[eid];
      var current = existingData[idx];
      var updated = current.slice();
      headers.forEach(function(h, col){
        if (emp[h] !== undefined) {
           updated[col] = String(emp[h]);
        }
      });
      existingData[idx] = updated;
      updatesCount++;
    } else {
      // Insert new
      var row = headers.map(function(h){
        if (h === 'password') return String(emp[h] !== undefined ? emp[h] : "123");
        return String(emp[h] !== undefined ? emp[h] : "");
      });
      newRows.push(row);
      insertsCount++;
    }
  });
  
  if (updatesCount > 0) {
    sh.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  }
  
  if (newRows.length > 0) {
    var startRow = 2;
    if (lastRow >= 2) startRow = 2 + existingData.length; 
    sh.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }
  
  return { ok:true, updated: updatesCount, inserted: insertsCount };
}

function deleteEmployee(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  if (!sh) return { ok:false, reason:'no-sheet' };
  var idCol = getHeaderIndex(sh, 'id');
  if (idCol < 1) return { ok:false, reason:'no-id-header' };
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:false, reason:'empty' };
  var ids = sh.getRange(2, idCol, lastRow-1, 1).getValues();
  for (var i=0;i<ids.length;i++) {
    if (String(ids[i][0]) === String(id||'')) {
      sh.deleteRow(2+i);
      return { ok:true, deleted:true };
    }
  }
  return { ok:false, reason:'not-found' };
}

function writeAttendance(items, mode) {
  var headers = ["recordId","employeeId","employeeName","workUnit","timestamp","jamMasuk","typeMasuk","statusMasuk","jamKeluar","typeKeluar","statusKeluar","shift","workLocation","notes","method","leaveType","latitude","longitude"];
  var sh = ensureSheet('Attendance', headers, mode);
  var arr = (Array.isArray(items) ? items : []).filter(function(r){return r && (String(r.id||r.recordId||'').length>0);}).map(function(r){
    return [
      String(r.id||r.recordId||""), String(r.employeeId||""), String(r.employeeName||""), String(r.workUnit||""),
      String(r.timestamp||""), String(r.jamMasuk||""), String(r.typeMasuk||""), String(r.statusMasuk||""),
      String(r.jamKeluar||""), String(r.typeKeluar||""), String(r.statusKeluar||""), String(r.shift||""),
      String(r.workLocation||""), String(r.notes||""), String(r.method||""), String(r.leaveType||""),
      String(r.latitude||""), String(r.longitude||"")
    ];
  });
  if (arr.length > 0) {
    var doReplace = String(mode||'replace') === 'replace';
    if (doReplace) {
      sh.getRange(2,1,arr.length,headers.length).setValues(arr);
    } else {
      var recCol = getHeaderIndex(sh, 'recordId');
      var existing = {};
      if (sh.getLastRow() >= 2 && recCol > 0) {
        var vals = sh.getRange(2, recCol, sh.getLastRow()-1, 1).getValues();
        for (var i=0;i<vals.length;i++) existing[String(vals[i][0])] = true;
      }
      var toWrite = [];
      for (var j=0;j<arr.length;j++) {
        var recId = String(arr[j][0]||'');
        if (!existing[recId]) toWrite.push(arr[j]);
      }
      if (toWrite.length > 0) {
        var startRow = Math.max(2, sh.getLastRow()+1);
        sh.getRange(startRow,1,toWrite.length,headers.length).setValues(toWrite);
      }
    }
  }
}

function upsertAttendance(items) {
  var headers = ["recordId","employeeId","employeeName","workUnit","timestamp","jamMasuk","typeMasuk","statusMasuk","jamKeluar","typeKeluar","statusKeluar","shift","workLocation","notes","method","leaveType","latitude","longitude"];
  var sh = ensureSheet('Attendance', headers, 'append');
  var recCol = getHeaderIndex(sh, 'recordId');
  if (recCol < 1) throw new Error('Header recordId tidak ditemukan');
  var lastRow = sh.getLastRow();
  var existingMap = {};
  if (lastRow >= 2) {
    var ids = sh.getRange(2, recCol, lastRow-1, 1).getValues();
    for (var i=0;i<ids.length;i++) {
      var rid = String(ids[i][0]||'');
      if (rid) existingMap[rid] = 2+i;
    }
  }
  var arr = (Array.isArray(items) ? items : []).filter(function(r){return r && (String(r.id||r.recordId||'').length>0);}).map(function(r){
    return [
      String(r.id||r.recordId||""), String(r.employeeId||""), String(r.employeeName||""), String(r.workUnit||""),
      String(r.timestamp||""), String(r.jamMasuk||""), String(r.typeMasuk||""), String(r.statusMasuk||""),
      String(r.jamKeluar||""), String(r.typeKeluar||""), String(r.statusKeluar||""), String(r.shift||""),
      String(r.workLocation||""), String(r.notes||""), String(r.method||""), String(r.leaveType||""),
      String(r.latitude||""), String(r.longitude||"")
    ];
  });
  var toAppend = [];
  for (var j=0;j<arr.length;j++) {
    var rid = String(arr[j][0]||'');
    var rowIndex = existingMap[rid];
    if (rowIndex) {
      sh.getRange(rowIndex,1,1,headers.length).setValues([arr[j]]);
    } else {
      toAppend.push(arr[j]);
    }
  }
  if (toAppend.length > 0) {
    var startRow = Math.max(2, sh.getLastRow()+1);
    sh.getRange(startRow,1,toAppend.length,headers.length).setValues(toAppend);
  }
}

function exportCSV(name) {
  var rows = getSheetRows(name);
  if (!rows || rows.length === 0) return { ok:false, reason:'empty' };
  var headers = Object.keys(rows[0]);
  var csv = [headers.join(',')].concat(rows.map(function(r){
    return headers.map(function(h){
      var v = r[h] == null ? '' : String(r[h]).replace(/"/g,'""');
      if (v.indexOf(',')>=0 || v.indexOf('"')>=0) v = '"'+v+'"';
      return v;
    }).join(',');
  })).join('\n');
  var folder = ensureFolder('Absensi_Exports');
  var file = folder.createFile(name+'_'+Date.now()+'.csv', csv, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, fileId:file.getId(), url:file.getUrl() };
}

function exportJSON(name) {
  var rows = getSheetRows(name);
  var json = JSON.stringify(rows||[], null, 2);
  var folder = ensureFolder('Absensi_Exports');
  var file = folder.createFile(name+'_'+Date.now()+'.json', json, MimeType.PLAIN_TEXT);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, fileId:file.getId(), url:file.getUrl() };
}

function ensureAllSheets(mode) {
  var m = String(mode||'safe');
  var employeesHeaders = ["id","name","workUnit","role","password","gender","phone","email","address","joinDate","nik","npp","golongan","pangkat","masaKerja","birthPlace","birthDate","religion","maritalStatus","status","photoUrl"];
  var attendanceHeaders = ["recordId","employeeId","employeeName","workUnit","timestamp","jamMasuk","typeMasuk","statusMasuk","jamKeluar","typeKeluar","statusKeluar","shift","workLocation","notes","method","leaveType","latitude","longitude","Hitung_Keterlambatan","Hitung_PulangCepat"];
  var settingsHeaders = ["key","value"];
  ensureSheet('Employees', employeesHeaders, m==='reset' ? 'replace' : 'safe');
  ensureSheet('Attendance', attendanceHeaders, m==='reset' ? 'replace' : 'safe');
  ensureSheet('Settings', settingsHeaders, m==='reset' ? 'replace' : 'safe');
}

/**
 * --- APPSHEET INTEGRATION HELPER ---
 * Jalankan fungsi ini sekali untuk menyiapkan Sheet agar "Ramah AppSheet"
 * dan memiliki logika otomatis (ArrayFormula) yang sinkron untuk kedua aplikasi.
 */
function setupAppSheetCompatibility() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Absensi dengan Rumus Otomatis
  var shAtt = ss.getSheetByName('Attendance');
  if (shAtt) {
    // Pastikan header ada
    var headers = ["recordId","employeeId","employeeName","workUnit","timestamp","jamMasuk","typeMasuk","statusMasuk","jamKeluar","typeKeluar","statusKeluar","shift","workLocation","notes","method","leaveType","latitude","longitude", "Hitung_Keterlambatan", "Hitung_PulangCepat"];
    
    // Cek apakah kolom rumus sudah ada, jika belum tambahkan
    var currentHeaders = shAtt.getRange(1, 1, 1, shAtt.getLastColumn()).getValues()[0];
    if (currentHeaders.indexOf("Hitung_Keterlambatan") === -1) {
      shAtt.getRange(1, currentHeaders.length + 1).setValue("Hitung_Keterlambatan");
      shAtt.getRange(1, currentHeaders.length + 2).setValue("Hitung_PulangCepat");
    }

    // INJEKSI ARRAY FORMULA DI BARIS 2
    // Rumus ini akan otomatis menghitung keterlambatan untuk SEMUA baris, baik input dari Web App maupun AppSheet
    // Asumsi: Kolom F adalah Jam Masuk (index 6), Kolom I adalah Jam Keluar (index 9)
    // Format Jam di Sheet harus HH:mm
    
    // Formula Keterlambatan (Contoh sederhana: Lewat 08:00 dihitung terlambat)
    // Logika: Jika JamMasuk > 08:00, maka (JamMasuk - 08:00). Jika tidak, 0.
    var colJamMasukLetter = "F"; 
    var formulaLate = '={"Hitung_Keterlambatan";ARRAYFORMULA(IF(LEN(' + colJamMasukLetter + '2:' + colJamMasukLetter + '), IF(' + colJamMasukLetter + '2:' + colJamMasukLetter + ' > TIME(8,0,0), TEXT(' + colJamMasukLetter + '2:' + colJamMasukLetter + ' - TIME(8,0,0), "HH:mm:ss"), "-"), ""))}';
    
    // Formula Pulang Cepat
    var colJamKeluarLetter = "I";
    var formulaEarly = '={"Hitung_PulangCepat";ARRAYFORMULA(IF(LEN(' + colJamKeluarLetter + '2:' + colJamKeluarLetter + '), IF(' + colJamKeluarLetter + '2:' + colJamKeluarLetter + ' < TIME(17,0,0), TEXT(TIME(17,0,0) - ' + colJamKeluarLetter + '2:' + colJamKeluarLetter + ', "HH:mm:ss"), "-"), ""))}';

    // Set Formula di Header Row (Teknik ArrayFormula Header) agar tidak terhapus saat insert row
    // Atau kita taruh di Baris 2 jika user preferensi manual.
    // Di sini kita gunakan teknik "Header Injection" agar lebih aman.
    
    // Namun untuk keamanan AppSheet, lebih baik formula ditaruh di Row 2, dan AppSheet di-setting "Read Only" untuk kolom tersebut.
    var lastCol = shAtt.getLastColumn();
    // Cari index kolom
    var idxLate = getHeaderIndex(shAtt, "Hitung_Keterlambatan");
    var idxEarly = getHeaderIndex(shAtt, "Hitung_PulangCepat");
    
    if (idxLate > 0) shAtt.getRange(1, idxLate).setFormula(formulaLate);
    if (idxEarly > 0) shAtt.getRange(1, idxEarly).setFormula(formulaEarly);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Absensi Setup')
    .addItem('Initialize Sheets', 'ensureAllSheets')
    .addItem('Create Folders', 'setupFolders')
    .addItem('Setup AppSheet Formulas', 'setupAppSheetCompatibility')
    .addToUi();
}

function setupFolders() {
  ensureFolder('Absensi_Photos');
  ensureFolder('Absensi_Exports');
  return true;
}
function writeSettings(obj) {
  var headers = ["key","value"];
  var sh = ensureSheet('Settings', headers, 'safe');
  var rows = [["json", JSON.stringify(obj||{})]];
  sh.getRange(2,1,rows.length,headers.length).setValues(rows);
}
function getSettings() {
  var rows = getSheetRows('Settings');
  var found = null;
  for (var i=0;i<rows.length;i++) {
    if (String(rows[i].key||'') === 'json') { found = rows[i]; break; }
  }
  var v = found && found.value ? String(found.value) : '';
  var obj = {};
  try { obj = v ? JSON.parse(v) : {}; } catch (e) { obj = {}; }
  if (!obj.jamMasuk) obj.jamMasuk = '08:00';
  if (!obj.jamPulang) obj.jamPulang = '17:00';
  if (obj.toleransiTerlambatMenit == null) obj.toleransiTerlambatMenit = 0;
  if (obj.toleransiPulangCepatMenit == null) obj.toleransiPulangCepatMenit = 0;
  return obj;
}
function resetSheets() {
  ensureAllSheets('reset');
  return { ok:true };
}

function getSheetRows(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  var values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  return values.map(function(row){
    var o = {};
    headers.forEach(function(h, i){ o[String(h)] = row[i]; });
    return o;
  });
}

function doRelay(obj) {
  var action = obj && obj.action ? String(obj.action) : '';
  var payload = obj && obj.payload ? obj.payload : null;
  var mode = obj && obj.mode ? String(obj.mode) : 'replace';
  if (action === 'sync_employees') { writeEmployees(payload); }
  else if (action === 'sync_attendance') { writeAttendance(payload, mode); }
  else if (action === 'sync_settings') { writeSettings(payload); }
  return true;
}
function include(name) {
  try {
    return HtmlService.createHtmlOutputFromFile(String(name||'')).getContent();
  } catch (e) {
    return '';
  }
}

function ensureFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function savePhotoBase64(dataUrl, preferName) {
  var folder = ensureFolder('Absensi_Photos');
  var parts = String(dataUrl||'').split(',');
  var meta = parts[0] || '';
  var b64 = parts[1] || '';
  var contentType = 'image/png';
  if (meta.indexOf('image/jpeg') >= 0) contentType = 'image/jpeg';
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, preferName||('absensi_'+Date.now()));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function listEmployees() {
  return getSheetRows('Employees');
}
function listUnits() {
  var emps = getSheetRows('Employees');
  var set = {};
  var out = [];
  for (var i=0;i<emps.length;i++) {
    var u = String(emps[i].workUnit||'').trim();
    if (u && !set[u]) { set[u]=true; out.push(u); }
  }
  out.sort();
  return out;
}

function listAttendanceRange(startIso, endIso, unitFilter, empFilter) {
  var rows = getSheetRows('Attendance');
  var s = String(startIso||'');
  var e = String(endIso||'');
  var uf = String(unitFilter||'').trim();
  var ef = String(empFilter||'').trim();
  if (!s && !e) return rows;
  var sd = s ? new Date(s) : null;
  var ed = e ? new Date(e) : null;
  return rows.filter(function(r){
    var ts = new Date(String(r.timestamp||''));
    if (sd && ts < sd) return false;
    if (ed && ts > ed) return false;
    if (uf && String(r.workUnit||'').trim() !== uf) return false;
    if (ef && String(r.employeeId||'').trim() !== ef) return false;
    return true;
  });
}
function computeLateEarlyReport(startIso, endIso, unitFilter, empFilter) {
  var cfg = getSettings();
  var jm = String(cfg.jamMasuk||'08:00');
  var jp = String(cfg.jamPulang||'17:00');
  var tolL = Number(cfg.toleransiTerlambatMenit||0);
  var tolE = Number(cfg.toleransiPulangCepatMenit||0);
  function toMin(s) {
    var p = String(s||'').split(':');
    var h = parseInt(p[0]||'0',10);
    var m = parseInt(p[1]||'0',10);
    return h*60+m;
  }
  var stdIn = toMin(jm);
  var stdOut = toMin(jp);
  var rows = listAttendanceRange(startIso, endIso, unitFilter, empFilter);
  return rows.map(function(r){
    var ci = toMin(String(r.jamMasuk||''));
    var co = toMin(String(r.jamKeluar||''));
    var late = 0;
    var early = 0;
    if (String(r.jamMasuk||'').length>0 && ci > stdIn) late = ci - stdIn;
    if (String(r.jamKeluar||'').length>0 && co < stdOut) early = stdOut - co;
    if (late > 0 && tolL > 0) late = Math.max(0, late - tolL);
    if (early > 0 && tolE > 0) early = Math.max(0, early - tolE);
    var photoId = String(r.photoId||'');
    var photoUrl = '';
    if (photoId) {
      try { photoUrl = DriveApp.getFileById(photoId).getUrl(); } catch (e) { photoUrl = ''; }
    }
    return {
      recordId: r.recordId||'',
      employeeId: r.employeeId||'',
      employeeName: r.employeeName||'',
      workUnit: r.workUnit||'',
      timestamp: r.timestamp||'',
      jamMasuk: r.jamMasuk||'',
      jamKeluar: r.jamKeluar||'',
      terlambatMenit: late,
      pulangCepatMenit: early,
      photoId: photoId,
      photoUrl: photoUrl
    };
  });
}
function computeLateEarlySummary(startIso, endIso, unitFilter, empFilter) {
  var detail = computeLateEarlyReport(startIso, endIso, unitFilter, empFilter);
  var agg = {};
  for (var i=0;i<detail.length;i++) {
    var r = detail[i];
    var id = String(r.employeeId||'');
    if (!agg[id]) agg[id] = { employeeId:id, employeeName:String(r.employeeName||''), totalTerlambatMenit:0, totalPulangCepatMenit:0, totalCheckIn:0, totalCheckOut:0 };
    agg[id].totalTerlambatMenit += Number(r.terlambatMenit||0);
    agg[id].totalPulangCepatMenit += Number(r.pulangCepatMenit||0);
    if (String(r.jamMasuk||'').length>0) agg[id].totalCheckIn += 1;
    if (String(r.jamKeluar||'').length>0) agg[id].totalCheckOut += 1;
  }
  var out = [];
  for (var k in agg) { if (Object.prototype.hasOwnProperty.call(agg,k)) out.push(agg[k]); }
  return out;
}
function exportLateEarlySummary(startIso, endIso) {
  var rows = computeLateEarlySummary(startIso, endIso);
  if (!rows || rows.length === 0) return { ok:false, reason:'empty' };
  var headers = Object.keys(rows[0]);
  var csv = [headers.join(',')].concat(rows.map(function(r){
    return headers.map(function(h){
      var v = r[h] == null ? '' : String(r[h]).replace(/"/g,'""');
      if (v.indexOf(',')>=0 || v.indexOf('"')>=0) v = '"'+v+'"';
      return v;
    }).join(',');
  })).join('\n');
  var folder = ensureFolder('Absensi_Exports');
  var name = 'LateEarlySummary_'+(startIso||'')+'_'+(endIso||'')+'_'+Date.now();
  var file = folder.createFile(name+'.csv', csv, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, fileId:file.getId(), url:file.getUrl() };
}
function exportLateEarlyDetailCSV(startIso, endIso, unitFilter, empFilter) {
  var rows = computeLateEarlyReport(startIso, endIso, unitFilter, empFilter);
  if (!rows || rows.length === 0) return { ok:false, reason:'empty' };
  var headers = Object.keys(rows[0]);
  var csv = [headers.join(',')].concat(rows.map(function(r){
    return headers.map(function(h){
      var v = r[h] == null ? '' : String(r[h]).replace(/"/g,'""');
      if (v.indexOf(',')>=0 || v.indexOf('"')>=0) v = '"'+v+'"';
      return v;
    }).join(',');
  })).join('\n');
  var folder = ensureFolder('Absensi_Exports');
  var name = 'LateEarlyDetail_'+(startIso||'')+'_'+(endIso||'')+'_'+Date.now();
  var file = folder.createFile(name+'.csv', csv, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, fileId:file.getId(), url:file.getUrl() };
}

function recordAttendanceWithPhoto(opts) {
  var o = opts || {};
  var empId = String(o.employeeId||'');
  var type = String(o.type||'').toLowerCase();
  var status = String(o.status||'');
  var notes = String(o.notes||'');
  var imageDataUrl = String(o.imageDataUrl||'');
  var empName = String(o.employeeName||'');
  var workUnit = String(o.workUnit||'');
  if (!empName || !workUnit) {
    var emps = listEmployees();
    for (var i=0;i<emps.length;i++) {
      if (String(emps[i].id||'') === empId) {
        empName = String(emps[i].name||'');
        workUnit = String(emps[i].workUnit||'');
        break;
      }
    }
  }
  var photoId = imageDataUrl ? savePhotoBase64(imageDataUrl, empId+'_'+Date.now()) : '';
  var now = new Date();
  var iso = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  var jam = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
  var typeMasuk = type === 'check-in' || type === 'masuk' ? 'Masuk' : '';
  var typeKeluar = type === 'check-out' || type === 'pulang' ? 'Pulang' : '';
  var recId = empId+'_'+iso+'_'+(typeMasuk ? 'check-in' : (typeKeluar ? 'check-out' : ''));
  var headers = ["recordId","employeeId","employeeName","workUnit","timestamp","jamMasuk","typeMasuk","statusMasuk","jamKeluar","typeKeluar","statusKeluar","shift","workLocation","notes","method","leaveType","latitude","longitude","photoId"];
  var sh = ensureSheet('Attendance', headers, 'append');
  var row = [
    recId, empId, empName, workUnit, iso,
    typeMasuk ? jam : '', typeMasuk, typeMasuk ? status : '',
    typeKeluar ? jam : '', typeKeluar, typeKeluar ? status : '',
    '', '', notes, 'Camera WebApp', '', '', '', photoId
  ];
  var startRow = Math.max(2, sh.getLastRow()+1);
  sh.getRange(startRow,1,1,headers.length).setValues([row]);
  return { ok:true, recordId: recId, photoId: photoId };
}
