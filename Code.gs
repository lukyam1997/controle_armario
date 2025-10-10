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
    'Users','AuditLog','Settings','Units','Logs'
  ];
  sheets.forEach(name => { if (!SS.getSheetByName(name)) SS.insertSheet(name); });

  ensureHeaders_('VisitorLockers', ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Stored Items','Expected End','Status Updated']);
  ensureHeaders_('VisitorRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit','Stored Items','Expected End']);
  ensureHeaders_('CompanionRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit','Stored Items','Expected End']);
  ensureHeaders_('Users', ['Username','Password','Profile','Email','Units']);
  ensureHeaders_('AuditLog', ['Timestamp','User','Action','Details']);
  ensureHeaders_('Logs', ['Timestamp','User','Entity','Action','Details','Metadata']);

  migrateUserUnits_();
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

function parseUnits_(value) {
  if (!value && value !== 0) return [];
  if (Array.isArray(value)) {
    return value.map(String).map(v => v.trim()).filter(Boolean);
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) {
        return parsed.map(String).map(v => v.trim()).filter(Boolean);
      }
    } catch (err) {
      // fallback to delimiter-based parsing
    }
    return trimmed.split(/[;,]/).map(v => v.trim()).filter(Boolean);
  }
  return [];
}

function stringifyUnits_(units) {
  const unique = Array.from(new Set(parseUnits_(units)));
  return unique.length ? JSON.stringify(unique) : '';
}

function parseMetadata_(value) {
  if (!value && value !== 0) return {};
  if (typeof value === 'object' && !Array.isArray(value)) return value || {};
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return {};
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (err) {
      return { raw: trimmed };
    }
    return { raw: trimmed };
  }
  return { raw: value };
}

function normalizeDate_(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function migrateUserUnits_() {
  const sh = SS.getSheetByName('Users');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;
  const range = sh.getRange(2, 5, lastRow - 1, 1);
  const values = range.getValues();
  let changed = false;
  const newValues = values.map(row => {
    const current = row[0];
    const normalized = stringifyUnits_(current);
    if (current !== normalized) changed = true;
    return [normalized];
  });
  if (changed) {
    range.setValues(newValues);
  }
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
  const actor = Session.getActiveUser().getEmail() || 'system';
  const metadata = { type, unit, count, rows: rowsN, cols: colsN };
  logAudit(actor, 'Update Locker Config', `${type} · ${unit || 'global'}`);
  logEvent(actor, 'Configuration', 'LockerLayout', sheetName, metadata);
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
  const actor = Session.getActiveUser().getEmail() || 'system';
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === key) {
      if (data[i][1] !== value) {
        sh.getRange(i+1,2).setValue(value);
        logEvent(actor, 'Configuration', 'SettingUpdated', key, { previous: data[i][1], value });
      }
      return;
    }
  }
  sh.appendRow([key, value]);
  logEvent(actor, 'Configuration', 'SettingCreated', key, { value });
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
        const units = parseUnits_(data[i][4]);
        logAudit(username, 'Login', 'Usuário autenticado');
        logEvent(username, 'Session', 'Login', 'Usuário autenticado', { units });
        return {
          success:true,
          profile:data[i][2],
          username,
          email: data[i][3] || '',
          units,
          defaultUnit: units[0] || ''
        };
      }
    }
    return { success:false, message:'Credenciais inválidas' };
  } catch (err) {
    return { success:false, message:'Erro no login: '+err };
  }
}

function addUser(username, password, profile, email, units) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0] === username) return { success:false, message:'Usuário já existe' };
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
  const normalizedUnits = parseUnits_(units);
  const storedUnits = stringifyUnits_(normalizedUnits);
  sh.appendRow([username, hash, profile, email, storedUnits]);
  const actor = Session.getActiveUser().getEmail() || 'Admin';
  logAudit(actor,'Add User', username);
  logEvent(actor, 'User', 'Create', username, { profile, units: normalizedUnits });
  return { success:true };
}

function resetPassword(username, newPassword) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, newPassword));
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === username) {
      sh.getRange(i+1,2).setValue(hash);
      const actor = Session.getActiveUser().getEmail() || 'Admin';
      logAudit(actor,'Reset Password', username);
      logEvent(actor, 'User', 'ResetPassword', username, {});
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
      const actor = Session.getActiveUser().getEmail() || 'Admin';
      logAudit(actor,'Delete User', username);
      logEvent(actor, 'User', 'Delete', username, {});
      return { success:true };
    }
  }
  return { success:false, message:'Usuário não encontrado' };
}

function listUsers() {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const users = [];
  for (let i=1; i<data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const units = parseUnits_(row[4]);
    users.push({
      username: row[0],
      profile: row[2],
      email: row[3] || '',
      units,
      unitsRaw: stringifyUnits_(units)
    });
  }
  return users;
}

function updateUserUnits(username, units) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const normalized = parseUnits_(units);
  const stored = stringifyUnits_(normalized);
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === username) {
      sh.getRange(i+1,5).setValue(stored);
      const actor = Session.getActiveUser().getEmail() || 'Admin';
      logAudit(actor, 'Update User Units', username);
      logEvent(actor, 'User', 'UpdateUnits', username, { units: normalized });
      return { success:true, units: normalized };
    }
  }
  return { success:false, message:'Usuário não encontrado' };
}

/** ========= ARMÁRIOS ========= **/
function getLockersData(type, unit='') {
  const sheetName = resolveLockerSheet_(type, unit);
  const sheet = SS.getSheetByName(sheetName);
  ensureLockerInventory_(sheet, type, unit);
  const data = sheet.getDataRange().getValues();
  const updated = recalcStatuses_(sheet, data, type);
  return updated;
}

function ensureLockerInventory_(sheet, type, unit) {
  if (!sheet || !unit) return;
  const lastRow = sheet.getLastRow();
  let hasNumbers = false;
  if (lastRow > 1) {
    const numbers = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    hasNumbers = numbers.some(row => {
      const value = row && row[0];
      return value !== null && value !== undefined && String(value).trim() !== '';
    });
  }
  if (hasNumbers) return;
  if (type === 'visitor') {
    const info = ensureVisitorDefaults_(unit);
    generateLockers(sheet.getName(), getSetting(`NUM_ARMARIOS_VISITOR_${info.key}`), '', getSetting(`NUM_ARMARIOS_VISITOR_${info.key}_ROWS`), getSetting(`NUM_ARMARIOS_VISITOR_${info.key}_COLS`));
  } else if (type === 'companion') {
    const info = ensureCompanionDefaults_(unit);
    generateLockers(sheet.getName(), getSetting(`NUM_ARMARIOS_COMPANION_${info.key}`), unit, getSetting(`NUM_ARMARIOS_COMPANION_${info.key}_ROWS`), getSetting(`NUM_ARMARIOS_COMPANION_${info.key}_COLS`));
  }
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
      const actor = Session.getActiveUser().getEmail() || 'anonymous';
      const metadata = { type, unit, locker: lockerNum, visitorName, patientName, phone, storedItems: storedItems || '', expectedEnd: end || '' };
      logAudit(actor, 'Register', `${visitorName} (${lockerNum})`);
      logEvent(actor, 'Locker', 'CheckIn', lockerNum, metadata);
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
      const actor = Session.getActiveUser().getEmail() || 'anonymous';
      const metadata = { type, unit, locker: lockerNum, visitorName: vname, phone, storedItems: stored || '', expectedEnd: end || '' };
      logAudit(actor,'Checkout', `${vname} (${lockerNum})`);
      logEvent(actor, 'Locker', 'Checkout', lockerNum, metadata);
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

function collectRegistrations_(type) {
  const sheetName = type === 'visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations';
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < values.length; i++) {
    records.push(mapRegistrationRow_(type, values[i]));
  }
  return records;
}

function mapRegistrationRow_(type, row) {
  const timestamp = normalizeDate_(row[0]);
  const expectedEnd = normalizeDate_(row[9]);
  return {
    timestamp,
    type,
    visitorName: row[1] || '',
    phone: row[2] || '',
    patientName: row[3] || '',
    bed: row[4] || '',
    locker: row[5] || '',
    action: row[6] || '',
    unit: row[7] || '',
    storedItems: row[8] || '',
    expectedEnd,
    raw: row
  };
}

function getAdvancedReport(filters) {
  const tz = Session.getScriptTimeZone();
  const f = filters || {};
  const typeFilter = f.type && f.type !== 'all' ? String(f.type) : 'all';
  const unitFilter = f.unit && f.unit !== 'all' ? String(f.unit) : '';
  const unitList = Array.isArray(f.units) ? f.units.map(String).filter(Boolean) : [];
  const unitSet = new Set(unitList);
  const actionFilter = f.action && f.action !== 'all' ? String(f.action) : '';
  const statusFilter = f.status && f.status !== 'all' ? String(f.status) : 'all';
  const query = (f.query || '').toString().trim().toLowerCase();
  const start = normalizeDate_(f.start);
  const end = normalizeDate_(f.end);
  if (end) end.setHours(23, 59, 59, 999);
  const types = typeFilter === 'all' ? ['visitor', 'companion'] : [typeFilter];

  const rows = [];
  types.forEach(type => {
    collectRegistrations_(type).forEach(record => rows.push(record));
  });

  const filtered = rows.filter(record => {
    if (unitSet.size) {
      if (!record.unit || !unitSet.has(record.unit)) return false;
    } else if (unitFilter && record.unit !== unitFilter) return false;
    if (actionFilter && record.action !== actionFilter) return false;
    if (statusFilter !== 'all') {
      const status = record.action === 'Registrar' ? 'Ativo' : 'Concluído';
      if (statusFilter === 'ativo' && status !== 'Ativo') return false;
      if (statusFilter === 'concluido' && status !== 'Concluído') return false;
    }
    if (start && record.timestamp && record.timestamp < start) return false;
    if (end && record.timestamp && record.timestamp > end) return false;
    if (query) {
      const searchable = [
        record.visitorName,
        record.patientName,
        record.bed,
        record.locker,
        record.unit,
        record.phone,
        record.storedItems
      ].join(' ').toLowerCase();
      if (!searchable.includes(query)) return false;
    }
    return true;
  });

  filtered.sort((a, b) => {
    const at = a.timestamp ? a.timestamp.getTime() : 0;
    const bt = b.timestamp ? b.timestamp.getTime() : 0;
    return bt - at;
  });

  const headers = ['Data', 'Tipo', 'Unidade', 'Visitante', 'Paciente', 'Leito', 'Armário', 'Ação', 'Itens', 'Previsão', 'Telefone'];
  const summary = {
    total: filtered.length,
    byAction: {},
    byType: {},
    byUnit: {},
    byDay: {}
  };

  const rowsFormatted = filtered.map(record => {
    const timestampDisplay = record.timestamp ? Utilities.formatDate(record.timestamp, tz, 'dd/MM/yyyy HH:mm') : '';
    const expectedDisplay = record.expectedEnd ? Utilities.formatDate(record.expectedEnd, tz, 'dd/MM/yyyy HH:mm') : '';
    const dayKey = record.timestamp ? Utilities.formatDate(record.timestamp, tz, 'yyyy-MM-dd') : 'sem-data';
    summary.byAction[record.action] = (summary.byAction[record.action] || 0) + 1;
    summary.byType[record.type] = (summary.byType[record.type] || 0) + 1;
    summary.byUnit[record.unit || 'Sem unidade'] = (summary.byUnit[record.unit || 'Sem unidade'] || 0) + 1;
    summary.byDay[dayKey] = (summary.byDay[dayKey] || 0) + 1;
    return {
      timestamp: record.timestamp ? record.timestamp.toISOString() : '',
      timestampDisplay,
      type: record.type,
      unit: record.unit,
      visitorName: record.visitorName,
      patientName: record.patientName,
      bed: record.bed,
      locker: record.locker,
      action: record.action,
      storedItems: record.storedItems,
      expectedEnd: record.expectedEnd ? record.expectedEnd.toISOString() : '',
      expectedDisplay,
      phone: record.phone
    };
  });

  const filtersApplied = {
    type: typeFilter,
    unit: unitFilter,
    units: unitList,
    action: actionFilter,
    status: statusFilter,
    query,
    start: start ? start.toISOString() : '',
    end: end ? end.toISOString() : ''
  };

  return {
    success: true,
    headers,
    rows: rowsFormatted,
    summary,
    filters: filtersApplied
  };
}

function exportAdvancedReport(format, filters) {
  const report = getAdvancedReport(filters || {});
  if (!report.success) return report;
  const actor = Session.getActiveUser().getEmail() || 'anonymous';
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const filenameBase = `relatorio_armarios_${timestamp}`;
  const headers = report.headers;

  if (format === 'csv') {
    const rows = [headers.join(',')];
    report.rows.forEach(row => {
      const data = [
        row.timestampDisplay,
        row.type,
        row.unit,
        row.visitorName,
        row.patientName,
        row.bed,
        row.locker,
        row.action,
        row.storedItems,
        row.expectedDisplay,
        row.phone
      ];
      rows.push(data.map(v => toCsv_(v)).join(','));
    });
    const csv = rows.join('\n');
    logAudit(actor, 'Export Report', 'CSV');
    logEvent(actor, 'Report', 'Export', 'CSV', { total: report.summary.total, filters: report.filters });
    return { success: true, format: 'csv', filename: `${filenameBase}.csv`, data: csv };
  }

  if (format === 'pdf') {
    const tableRows = report.rows.map(row => `
      <tr>
        <td>${sanitizeHtml_(row.timestampDisplay)}</td>
        <td>${sanitizeHtml_(row.type)}</td>
        <td>${sanitizeHtml_(row.unit)}</td>
        <td>${sanitizeHtml_(row.visitorName)}</td>
        <td>${sanitizeHtml_(row.patientName)}</td>
        <td>${sanitizeHtml_(row.bed)}</td>
        <td>${sanitizeHtml_(row.locker)}</td>
        <td>${sanitizeHtml_(row.action)}</td>
        <td>${sanitizeHtml_(row.storedItems)}</td>
        <td>${sanitizeHtml_(row.expectedDisplay)}</td>
        <td>${sanitizeHtml_(row.phone)}</td>
      </tr>`).join('');
    const unitLabel = (report.filters.units && report.filters.units.length)
      ? report.filters.units.join(', ')
      : (report.filters.unit || 'Todas');
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><style>
      body { font-family: Arial, sans-serif; font-size: 12px; color: #111; }
      h1 { font-size: 18px; margin-bottom: 6px; }
      table { width: 100%; border-collapse: collapse; margin-top: 12px; }
      th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
      th { background: #f3f4f6; text-transform: uppercase; font-size: 10px; }
      .summary { margin-top: 16px; }
      .summary span { display: inline-block; margin-right: 12px; }
    </style></head><body>
      <h1>Relatório de Armários</h1>
      <div class="summary">
        <span><strong>Total:</strong> ${report.summary.total}</span>
        <span><strong>Filtro tipo:</strong> ${sanitizeHtml_(report.filters.type)}</span>
        <span><strong>Unidades:</strong> ${sanitizeHtml_(unitLabel)}</span>
      </div>
      <table>
        <thead><tr>${headers.map(h => `<th>${sanitizeHtml_(h)}</th>`).join('')}</tr></thead>
        <tbody>${tableRows || '<tr><td colspan="11">Sem registros</td></tr>'}</tbody>
      </table>
    </body></html>`;
    const blob = Utilities.newBlob(html, 'text/html', `${filenameBase}.html`).getAs('application/pdf');
    logAudit(actor, 'Export Report', 'PDF');
    logEvent(actor, 'Report', 'Export', 'PDF', { total: report.summary.total, filters: report.filters });
    return { success: true, format: 'pdf', filename: `${filenameBase}.pdf`, data: Utilities.base64Encode(blob.getBytes()) };
  }

  return { success: false, message: 'Formato não suportado' };
}

function sanitizeHtml_(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;' }[c]));
}

function toCsv_(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  if (/[",\n]/.test(str)) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function getIndicatorsData(filters) {
  const f = filters || {};
  const tz = Session.getScriptTimeZone();
  const typeFilter = f.type && f.type !== 'all' ? String(f.type) : 'all';
  const types = typeFilter === 'all' ? ['visitor', 'companion'] : [typeFilter];
  const requestedUnits = Array.isArray(f.units) ? f.units.map(String) : [];
  const availableUnits = listUnits();
  const unitNames = requestedUnits.length ? requestedUnits : availableUnits.map(u => u.name);
  const unitSet = new Set(unitNames);
  const days = Number(f.days) > 0 ? Number(f.days) : 30;
  const now = new Date();
  const end = normalizeDate_(f.end) || now;
  const start = normalizeDate_(f.start) || new Date(end.getTime() - days * 24 * 60 * 60 * 1000);
  start.setHours(0,0,0,0);
  end.setHours(23,59,59,999);

  const usageByDay = {};
  const summaryByUnit = {};
  const summaryByType = {};
  const expectedDurations = [];
  const recent = [];

  types.forEach(type => { summaryByType[type] = { entries:0, checkouts:0 }; });

  types.forEach(type => {
    collectRegistrations_(type).forEach(record => {
      const timestamp = record.timestamp;
      if (!timestamp) return;
      if (timestamp < start || timestamp > end) return;
      if (unitSet.size && record.unit && !unitSet.has(record.unit)) return;
      const unitKey = record.unit || 'Sem unidade';
      if (!summaryByUnit[unitKey]) summaryByUnit[unitKey] = { entries:0, checkouts:0, typeBreakdown:{} };
      if (!summaryByUnit[unitKey].typeBreakdown[type]) summaryByUnit[unitKey].typeBreakdown[type] = { entries:0, checkouts:0 };
      const dayKey = Utilities.formatDate(timestamp, tz, 'yyyy-MM-dd');
      if (!usageByDay[dayKey]) usageByDay[dayKey] = { date: dayKey, entries:0, checkouts:0 };
      const formatted = {
        timestamp: timestamp.toISOString(),
        timestampDisplay: Utilities.formatDate(timestamp, tz, 'dd/MM/yyyy HH:mm'),
        type,
        unit: record.unit,
        visitorName: record.visitorName,
        patientName: record.patientName,
        locker: record.locker,
        action: record.action
      };
      recent.push(formatted);
      if (record.action === 'Registrar') {
        usageByDay[dayKey].entries += 1;
        summaryByUnit[unitKey].entries += 1;
        summaryByUnit[unitKey].typeBreakdown[type].entries += 1;
        summaryByType[type].entries += 1;
        if (record.expectedEnd && record.timestamp) {
          const diff = record.expectedEnd.getTime() - record.timestamp.getTime();
          if (!isNaN(diff) && diff > 0) {
            expectedDurations.push(diff / (1000 * 60 * 60));
          }
        }
      } else if (record.action === 'Baixa') {
        usageByDay[dayKey].checkouts += 1;
        summaryByUnit[unitKey].checkouts += 1;
        summaryByUnit[unitKey].typeBreakdown[type].checkouts += 1;
        summaryByType[type].checkouts += 1;
      }
    });
  });

  recent.sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || ''));
  const recentLimited = recent.slice(0, 12);

  const occupancy = [];
  unitNames.forEach(unit => {
    types.forEach(type => {
      try {
        const stats = getLockerStats(type, unit);
        occupancy.push({
          unit,
          type,
          total: stats.total,
          free: stats.free,
          inUse: stats.inUse,
          dueSoon: stats.dueSoon,
          overdue: stats.overdue,
          occupancyRate: stats.total ? (stats.total - stats.free) / stats.total : 0
        });
      } catch (err) {
        // ignore units sem configuração
      }
    });
  });

  const averageExpectedHours = expectedDurations.length
    ? expectedDurations.reduce((sum, hours) => sum + hours, 0) / expectedDurations.length
    : 0;

  const activitySeries = Object.values(usageByDay)
    .map(entry => ({
      date: entry.date,
      entries: entry.entries,
      checkouts: entry.checkouts
    }))
    .sort((a, b) => a.date.localeCompare(b.date));

  const unitSeries = Object.keys(summaryByUnit).map(unit => ({
    unit,
    entries: summaryByUnit[unit].entries,
    checkouts: summaryByUnit[unit].checkouts,
    typeBreakdown: summaryByUnit[unit].typeBreakdown
  })).sort((a, b) => (b.entries + b.checkouts) - (a.entries + a.checkouts));

  const totals = {
    entries: unitSeries.reduce((sum, item) => sum + item.entries, 0),
    checkouts: unitSeries.reduce((sum, item) => sum + item.checkouts, 0)
  };

  return {
    success: true,
    period: {
      start: start.toISOString(),
      end: end.toISOString(),
      days
    },
    occupancy,
    activity: activitySeries,
    summaryByType,
    summaryByUnit: unitSeries,
    totals,
    expected: {
      averageHours: Number(averageExpectedHours.toFixed(2)),
      samples: expectedDurations.length
    },
    recent: recentLimited
  };
}

/** ========= AUDITORIA ========= **/
function logEvent(user, entity, action, details, metadata) {
  const sheet = SS.getSheetByName('Logs');
  if (!sheet) return;
  let metaString = '';
  if (metadata !== undefined) {
    try {
      metaString = JSON.stringify(metadata);
    } catch (err) {
      metaString = JSON.stringify({ raw: metadata });
    }
  }
  sheet.appendRow([new Date(), user || '', entity || '', action || '', details || '', metaString]);
}

function logAudit(user, action, details) {
  SS.getSheetByName('AuditLog').appendRow([new Date(), user, action, details]);
}
function getAuditLog() {
  return SS.getSheetByName('AuditLog').getDataRange().getValues();
}

function getLogs(limit) {
  const sheet = SS.getSheetByName('Logs');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  const startIndex = limit && limit > 0 && values.length - 1 > limit ? values.length - limit : 1;
  const rows = [];
  for (let i = startIndex; i < values.length; i++) {
    const row = values[i];
    rows.push({
      timestamp: row[0],
      user: row[1],
      entity: row[2],
      action: row[3],
      details: row[4],
      metadata: parseMetadata_(row[5])
    });
  }
  return rows;
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
  const actor = Session.getActiveUser().getEmail() || 'system';
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
      logAudit(actor, 'Update Unit', normalized);
      logEvent(actor, 'Unit', 'Reactivate', normalized, { display: display || normalized });
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
  logAudit(actor, 'Add Unit', normalized);
  logEvent(actor, 'Unit', 'Create', normalized, { display: display || normalized });
  return { success:true };
}

function removeUnit(unit) {
  if (!unit) return { success:false, message:'Nome obrigatório' };
  const sh = SS.getSheetByName('Units');
  const data = sh.getDataRange().getValues();
  const actor = Session.getActiveUser().getEmail() || 'system';
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]).toLowerCase() === unit.toLowerCase()) {
      sh.getRange(i+1,3).setValue(false);
      logAudit(actor, 'Remove Unit', unit);
      logEvent(actor, 'Unit', 'Deactivate', unit, {});
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
