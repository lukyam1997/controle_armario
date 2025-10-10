/** ========= GLOBALS ========= **/
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SPREADSHEET_ID = SS.getId();

const PROFILES = {
  VISITOR: 'armario_visitante',
  COMPANION: 'guarda_volume',
  ADMIN: 'admin'
};

/** ========= WEB APP ========= **/
function doGet(e) {
  // Garante setup mínimo antes de renderizar
  safeSetup_();
  const template = HtmlService.createTemplateFromFile('Index');
  return template
    .evaluate()
    .setTitle('Hospital Storage Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ========= SETUP ========= **/
function safeSetup_() {
  const sheets = [
    'VisitorLockers',
    'VisitorRegistrations','CompanionRegistrations',
    'Users','AuditLog','Settings','Units'
  ];
  sheets.forEach(name => { if (!SS.getSheetByName(name)) SS.insertSheet(name); });

  ensureHeaders_('VisitorLockers', ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
  ensureHeaders_('VisitorRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit','Stored Items','Expected End']);
  ensureHeaders_('CompanionRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit','Stored Items','Expected End']);
  ensureHeaders_('Users', ['Username','Password','Profile','Email','Unit']);
  ensureHeaders_('AuditLog', ['Timestamp','User','Action','Details']);
  ensureHeaders_('Settings', ['Key','Value']);

  ensureHeaders_('Units', ['Unit','Display Name','Active']);
  const unitSheet = SS.getSheetByName('Units');
  if (unitSheet.getLastRow() === 1) {
    unitSheet.appendRow(['NAC','NAC', true]);
    unitSheet.appendRow(['UIB','UIB', true]);
  }

  const defaults = [
    ['NUM_ARMARIOS_VISITANTE', 20],
    ['NUM_ARMARIOS_VISITANTE_ROWS', 4],
    ['NUM_ARMARIOS_VISITANTE_COLS', 5],
  ];
  defaults.forEach(([k,v])=>{ if (getSetting(k) === null) updateSetting(k, v); });

  // Gera planta se estiver vazia
  generateLockers('VisitorLockers', getSetting('NUM_ARMARIOS_VISITANTE'), '', getSetting('NUM_ARMARIOS_VISITANTE_ROWS'), getSetting('NUM_ARMARIOS_VISITANTE_COLS'));
  const activeUnits = listUnits();
  activeUnits.forEach(unit => {
    const visitorInfo = ensureVisitorDefaults_(unit.name);
    const visitorSheet = getVisitorSheetName_(unit.name);
    if (!SS.getSheetByName(visitorSheet)) SS.insertSheet(visitorSheet);
    ensureHeaders_(visitorSheet, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
    generateLockers(visitorSheet, getSetting(`NUM_ARMARIOS_VISITOR_${visitorInfo.key}`), '', getSetting(`NUM_ARMARIOS_VISITOR_${visitorInfo.key}_ROWS`), getSetting(`NUM_ARMARIOS_VISITOR_${visitorInfo.key}_COLS`));

    ensureCompanionDefaults_(unit.name);
    const sheetName = getCompanionSheetName_(unit.name);
    if (!SS.getSheetByName(sheetName)) SS.insertSheet(sheetName);
    ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
    generateLockers(sheetName, getSetting(`NUM_ARMARIOS_COMPANION_${unit.key}`), formatUnitKey_(unit.name), getSetting(`NUM_ARMARIOS_COMPANION_${unit.key}_ROWS`), getSetting(`NUM_ARMARIOS_COMPANION_${unit.key}_COLS`));
  });

  // Se não existir usuário, cria admin padrão (pode apagar depois)
  const usersSheet = SS.getSheetByName('Users');
  if (usersSheet.getLastRow() === 1) {
    const defaultUnit = activeUnits.length ? activeUnits[0].name : '';
    addUser('admin', 'admin', PROFILES.ADMIN, Session.getActiveUser().getEmail() || 'admin@example.com', defaultUnit);
  }
}

function ensureHeaders_(sheetName, headers) {
  const sh = SS.getSheetByName(sheetName) || SS.insertSheet(sheetName);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    const existing = sh.getRange(1,1,1,headers.length).getValues()[0];
    const need = headers.some((h,i)=>existing[i] !== h);
    if (need) {
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
}

function generateLockers(sheetName, count, prefix, rows, cols) {
  const sheet = SS.getSheetByName(sheetName);
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  const data = sheet.getDataRange().getValues();
  const have = new Set(data.slice(1).map(r => r[0])); // números já existentes

  const rowsInt = Number(rows) || 4;
  const colsInt = Number(cols) || 5;
  const target = Math.min(Number(count) || rowsInt*colsInt, rowsInt*colsInt);
  const toAppend = [];

  for (let i = 0; i < target; i++) {
    const r = Math.floor(i / colsInt);
    const c = (i % colsInt) + 1;
    const letter = letters[r % 26];
    const num = prefix ? `${prefix}-${letter}${c}` : `${letter}${c}`;
    if (!have.has(num)) {
      toAppend.push([num, 'Free', '', '', '', '', '', '', '', '']);
    }
  }
  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow()+1, 1, toAppend.length, 10).setValues(toAppend);
  }
}

function computeGridLayout_(count) {
  const safe = Math.max(1, Number(count) || 1);
  const cols = Math.ceil(Math.sqrt(safe));
  const rows = Math.ceil(safe / cols);
  return { rows, cols };
}

/** ========= SETTINGS ========= **/
function setLockerConfig(type, total, unit='') {
  let sheetName, prefix, key;
  if (type === 'visitor') {
    if (!unit) return { success:false, message:'Unidade obrigatória' };
    const info = ensureVisitorDefaults_(unit);
    sheetName = getVisitorSheetName_(unit);
    prefix = '';
    key = `NUM_ARMARIOS_VISITOR_${info.key}`;
  }
  else if (type === 'companion') {
    if (!unit) return { success:false, message:'Unidade obrigatória' };
    const companion = ensureCompanionDefaults_(unit);
    sheetName = getCompanionSheetName_(unit);
    prefix = formatUnitKey_(unit);
    key = `NUM_ARMARIOS_COMPANION_${companion.key}`;
  } else return { success:false, message:'Tipo inválido' };

  const totalN = Number(total);
  if (!totalN || totalN < 1) return { success:false, message:'Quantidade inválida' };
  const layout = computeGridLayout_(totalN);
  const rowsN = layout.rows;
  const colsN = layout.cols;
  const count = totalN;
  let sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    sheet = SS.insertSheet(sheetName);
    ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
  }
  const currentCount = sheet.getLastRow() - 1;
  if (count < currentCount) return { success:false, message:'Não é possível reduzir armários (existem registros).' };

  updateSetting(key, count);
  updateSetting(key+'_ROWS', rowsN);
  updateSetting(key+'_COLS', colsN);

  generateLockers(sheetName, count, prefix, rowsN, colsN);
  return { success:true, count, layout };
}

function getLockerConfig(type, unit='') {
  let key;
  if (type === 'visitor') {
    if (!unit) return { rows:0, cols:0, count:0 };
    const visitor = ensureVisitorDefaults_(unit);
    key = `NUM_ARMARIOS_VISITOR_${visitor.key}`;
  }
  else if (type === 'companion') {
    if (!unit) return { rows:0, cols:0, count:0 };
    const companion = ensureCompanionDefaults_(unit);
    key = `NUM_ARMARIOS_COMPANION_${companion.key}`;
  }

  return {
    rows: Number(getSetting(key+'_ROWS')) || 4,
    cols: Number(getSetting(key+'_COLS')) || 5,
    count: Number(getSetting(key)) || 20
  };
}

function updateSetting(key, value) {
  let sh = SS.getSheetByName('Settings');
  if (!sh) sh = SS.insertSheet('Settings');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i+1,2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}
function getSetting(key) {
  const sh = SS.getSheetByName('Settings');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0] === key) return data[i][1];
  return null;
}
function getSettings() {
  return SS.getSheetByName('Settings').getDataRange().getValues();
}

/** ========= LOGIN & USERS ========= **/
function login(username, password) {
  try {
    const usersSheet = SS.getSheetByName('Users');
    if (usersSheet.getLastRow() === 0) ensureHeaders_('Users', ['Username','Password','Profile','Email','Unit']);
    const data = usersSheet.getDataRange().getValues();
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
    for (let i=1; i<data.length; i++) {
      if (data[i][0] === username && data[i][1] === hash) {
        logAudit(username, 'Login', 'Usuário autenticado');
        return { success:true, profile:data[i][2], username, email: data[i][3] || '', unit: data[i][4] || '' };
      }
    }
    return { success:false, message:'Credenciais inválidas' };
  } catch (err) {
    return { success:false, message:'Erro no login: '+err };
  }
}

function addUser(username, password, profile, email, unit) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0] === username) return { success:false, message:'Usuário já existe' };
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
  sh.appendRow([username, hash, profile, email, unit || '']);
  logAudit('Admin','Add User', username);
  return { success:true };
}

function resetPassword(username, newPassword) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, newPassword));
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === username) {
      sh.getRange(i+1,2).setValue(hash);
      logAudit('Admin','Reset Password', username);
      return { success:true };
    }
  }
  return { success:false, message:'Usuário não encontrado' };
}

function deleteUser(username) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === username) {
      sh.deleteRow(i+1);
      logAudit('Admin','Delete User', username);
      return { success:true };
    }
  }
  return { success:false, message:'Usuário não encontrado' };
}

function listUsers() {
  return SS.getSheetByName('Users').getDataRange().getValues();
}

/** ========= ARMÁRIOS ========= **/
function getLockersData(type, unit='') {
  const sheetName = resolveLockerSheet_(type, unit);
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const updated = recalcStatuses_(sheet, data, type);
  return updated;
}

function getLockerStats(type, unit='') {
  const sheetName = resolveLockerSheet_(type, unit);
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues().slice(1);
  let free=0, inUse=0, dueSoon=0, overdue=0;
  data.forEach(r=>{
    switch(r[1]) {
      case 'Free': free++; break;
      case 'InUse': inUse++; break;
      case 'DueSoon': dueSoon++; break;
      case 'Overdue': overdue++; break;
    }
  });
  return { free, inUse, dueSoon, overdue, total:data.length };
}

function registerVisitor(patientName, bed, visitorName, phone, storedItems, expectedEnd, type, unit, lockerNum) {
  const sheetName = resolveLockerSheet_(type, unit);
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  for (let i=1; i<data.length; i++) {
    if (data[i][0] == lockerNum && data[i][1] === 'Free') {
      const start = new Date();
      const end = expectedEnd ? parseDateTimeLocal_(expectedEnd) : '';
      if (type === 'visitor' && !end) {
        return { success:false, message:'Horário de saída obrigatório' };
      }
      sheet.getRange(i+1,2).setValue('InUse');
      sheet.getRange(i+1,3).setValue(visitorName);
      sheet.getRange(i+1,4).setValue(phone);
      sheet.getRange(i+1,5).setValue(patientName);
      sheet.getRange(i+1,6).setValue(bed);
      sheet.getRange(i+1,7).setValue(start);
      sheet.getRange(i+1,8).setValue(storedItems || '');
      sheet.getRange(i+1,9).setValue(end || '');
      sheet.getRange(i+1,10).setValue(new Date());

      const regSheet = SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
      regSheet.appendRow([new Date(), visitorName, phone, patientName, bed, lockerNum, 'Registrar', unit, storedItems || '', end || '']);
      logAudit(Session.getActiveUser().getEmail() || 'anonymous', 'Register', `${visitorName} (${lockerNum})`);
      return { success:true, locker:lockerNum };
    }
  }
  return { success:false, message:'Armário não disponível' };
}

function checkoutLocker(lockerNum, type, unit) {
  const sheetName = resolveLockerSheet_(type, unit);
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] == lockerNum && data[i][1] !== 'Free') {
      const vname = data[i][2], phone=data[i][3];
      const stored = data[i][7];
      const end = data[i][8];
      // Limpa colunas 2..10 e preserva número
      sheet.getRange(i+1,2,1,9).clearContent();
      sheet.getRange(i+1,1).setValue(lockerNum);
      sheet.getRange(i+1,2).setValue('Free');

      const regSheet = SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
      regSheet.appendRow([new Date(), vname, phone, '', '', lockerNum, 'Baixa', unit, stored || '', end || '']);
      logAudit(Session.getActiveUser().getEmail() || 'anonymous','Checkout', `${vname} (${lockerNum})`);
      return { success:true };
    }
  }
  return { success:false, message:'Armário não encontrado' };
}

/** ========= RELATÓRIOS ========= **/
function exportReport(type) {
  const sh = SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
  const rows = sh.getDataRange().getValues();
  const out = rows.map(r => r.map(v => typeof v === 'string' && v.includes(',') ? `"${v.replace(/"/g,'""')}"` : v).join(',')).join('\n');
  return out;
}
function getRegistrations(type) {
  return SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations').getDataRange().getValues();
}

/** ========= AUDITORIA ========= **/
function logAudit(user, action, details) {
  SS.getSheetByName('AuditLog').appendRow([new Date(), user, action, details]);
}
function getAuditLog() {
  return SS.getSheetByName('AuditLog').getDataRange().getValues();
}

/** ========= UNITS ========= **/
function listUnits() {
  const sh = SS.getSheetByName('Units');
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i=1; i<data.length; i++) {
    if (String(data[i][2]) === 'true' || data[i][2] === true || data[i][2] === 'TRUE' || data[i][2] === 1) {
      const name = data[i][0];
      list.push({ name, display:data[i][1] || name, key: formatUnitKey_(name) });
    }
  }
  return list;
}

function addUnit(unit, display) {
  if (!unit) return { success:false, message:'Nome obrigatório' };
  const sh = SS.getSheetByName('Units');
  const normalized = unit.trim();
  const unitKey = formatUnitKey_(normalized);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]).toLowerCase() === normalized.toLowerCase()) {
      sh.getRange(i+1,2).setValue(display || normalized);
      sh.getRange(i+1,3).setValue(true);
      ensureVisitorDefaults_(normalized);
      const visitorSheet = getVisitorSheetName_(normalized);
      if (!SS.getSheetByName(visitorSheet)) {
        SS.insertSheet(visitorSheet);
        ensureHeaders_(visitorSheet, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
      }
      generateLockers(visitorSheet, getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}`), '', getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}_ROWS`), getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}_COLS`));
      ensureCompanionDefaults_(normalized);
      const sheetName = getCompanionSheetName_(normalized);
      if (!SS.getSheetByName(sheetName)) {
        SS.insertSheet(sheetName);
        ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
      }
      return { success:true, updated:true };
    }
  }
  sh.appendRow([normalized, display || normalized, true]);
  ensureVisitorDefaults_(normalized);
  const visitorSheet = getVisitorSheetName_(normalized);
  if (!SS.getSheetByName(visitorSheet)) {
    SS.insertSheet(visitorSheet);
    ensureHeaders_(visitorSheet, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
  }
  generateLockers(visitorSheet, getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}`), '', getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}_ROWS`), getSetting(`NUM_ARMARIOS_VISITOR_${unitKey}_COLS`));
  ensureCompanionDefaults_(normalized);
  const sheetName = getCompanionSheetName_(normalized);
  if (!SS.getSheetByName(sheetName)) {
    SS.insertSheet(sheetName);
    ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
  }
  generateLockers(sheetName, getSetting(`NUM_ARMARIOS_COMPANION_${unitKey}`), unitKey, getSetting(`NUM_ARMARIOS_COMPANION_${unitKey}_ROWS`), getSetting(`NUM_ARMARIOS_COMPANION_${unitKey}_COLS`));
  return { success:true };
}

function removeUnit(unit) {
  if (!unit) return { success:false, message:'Nome obrigatório' };
  const sh = SS.getSheetByName('Units');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]).toLowerCase() === unit.toLowerCase()) {
      sh.getRange(i+1,3).setValue(false);
      return { success:true };
    }
  }
  return { success:false, message:'Unidade não encontrada' };
}

/** ========= INCLUDE (para HTML) ========= **/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ========= HELPERS ========= **/
function resolveLockerSheet_(type, unit) {
  if (type === 'visitor') {
    if (!unit) throw new Error('Unidade obrigatória');
    const info = ensureVisitorDefaults_(unit);
    const sheetName = getVisitorSheetName_(unit);
    if (!SS.getSheetByName(sheetName)) {
      SS.insertSheet(sheetName);
      ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
      generateLockers(sheetName, getSetting(`NUM_ARMARIOS_VISITOR_${info.key}`), '', getSetting(`NUM_ARMARIOS_VISITOR_${info.key}_ROWS`), getSetting(`NUM_ARMARIOS_VISITOR_${info.key}_COLS`));
    }
    return sheetName;
  }
  if (type === 'companion') {
    if (!unit) throw new Error('Unidade obrigatória');
    const info = ensureCompanionDefaults_(unit);
    const sheetName = getCompanionSheetName_(unit);
    if (!SS.getSheetByName(sheetName)) {
      SS.insertSheet(sheetName);
      ensureHeaders_(sheetName, ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
      generateLockers(sheetName, getSetting(`NUM_ARMARIOS_COMPANION_${info.key}`), unit, getSetting(`NUM_ARMARIOS_COMPANION_${info.key}_ROWS`), getSetting(`NUM_ARMARIOS_COMPANION_${info.key}_COLS`));
    }
    return sheetName;
  }
  throw new Error('Tipo inválido');
}

function ensureVisitorDefaults_(unit) {
  const key = formatUnitKey_(unit);
  const defaults = [
    [`NUM_ARMARIOS_VISITOR_${key}`, getSetting('NUM_ARMARIOS_VISITANTE') || 20],
    [`NUM_ARMARIOS_VISITOR_${key}_ROWS`, getSetting('NUM_ARMARIOS_VISITANTE_ROWS') || 4],
    [`NUM_ARMARIOS_VISITOR_${key}_COLS`, getSetting('NUM_ARMARIOS_VISITANTE_COLS') || 5]
  ];
  defaults.forEach(([k,v])=>{ if (getSetting(k) === null) updateSetting(k, v); });
  return { name: unit, key };
}

function getVisitorSheetName_(unit) {
  const key = formatUnitKey_(unit);
  const modern = `VisitorLockers_${key}`;
  const legacy = `VisitorLockers${key}`;
  if (SS.getSheetByName(modern)) return modern;
  if (SS.getSheetByName(legacy)) return legacy;
  return modern;
}

function ensureCompanionDefaults_(unit) {
  const key = formatUnitKey_(unit);
  const defaults = [
    [`NUM_ARMARIOS_COMPANION_${key}`, 20],
    [`NUM_ARMARIOS_COMPANION_${key}_ROWS`, 4],
    [`NUM_ARMARIOS_COMPANION_${key}_COLS`, 5]
  ];
  const legacy = [
    [`NUM_ARMARIOS_${key}`, null],
    [`NUM_ARMARIOS_${key}_ROWS`, null],
    [`NUM_ARMARIOS_${key}_COLS`, null]
  ];
  legacy.forEach(([k])=>{
    const val = getSetting(k);
    if (val !== null) {
      const newKey = k.replace(`NUM_ARMARIOS_${key}`, `NUM_ARMARIOS_COMPANION_${key}`);
      if (getSetting(newKey) === null) updateSetting(newKey, val);
    }
  });
  defaults.forEach(([k,v])=>{ if (getSetting(k) === null) updateSetting(k, v); });
  return { name: unit, key };
}

function formatUnitKey_(unit) {
  return unit.toUpperCase().replace(/[^A-Z0-9]/g,'_');
}

function getCompanionSheetName_(unit) {
  const key = formatUnitKey_(unit);
  const modern = `CompanionLockers_${key}`;
  const legacy = `CompanionLockers${key}`;
  if (SS.getSheetByName(modern)) return modern;
  if (SS.getSheetByName(legacy)) return legacy;
  return modern;
}

function parseDateTimeLocal_(value) {
  if (!value) return null;
  try {
    return new Date(value);
  } catch (err) {
    return null;
  }
}

function recalcStatuses_(sheet, data, type) {
  if (!data || data.length <= 1) return data;
  let needsUpdate = false;
  const now = new Date();
  for (let i=1; i<data.length; i++) {
    const row = data[i];
    let status = row[1];
    if (status === 'Free' || !row[0]) continue;
    let newStatus = 'InUse';
    const end = row[8];
    if (type === 'visitor' && end) {
      const endDate = (end instanceof Date) ? end : new Date(end);
      const diff = endDate.getTime() - now.getTime();
      if (diff < 0) newStatus = 'Overdue';
      else if (diff <= 30*60*1000) newStatus = 'DueSoon';
      else newStatus = 'InUse';
    } else {
      newStatus = 'InUse';
    }
    if (status !== newStatus) {
      data[i][1] = newStatus;
      sheet.getRange(i+1,2).setValue(newStatus);
      sheet.getRange(i+1,10).setValue(new Date());
      needsUpdate = true;
    }
  }
  return needsUpdate ? sheet.getDataRange().getValues() : data;
}
