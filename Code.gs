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
    'VisitorLockers','CompanionLockersNAC','CompanionLockersUIB',
    'VisitorRegistrations','CompanionRegistrations',
    'Users','AuditLog','Settings'
  ];
  sheets.forEach(name => { if (!SS.getSheetByName(name)) SS.insertSheet(name); });

  ensureHeaders_('VisitorLockers', ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Welcome Sent','Alert Sent','End Time']);
  ensureHeaders_('CompanionLockersNAC', ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Welcome Sent','Alert Sent','End Time']);
  ensureHeaders_('CompanionLockersUIB', ['Number','Status','Visitor Name','Phone','Patient Name','Bed','Start Time','Welcome Sent','Alert Sent','End Time']);
  ensureHeaders_('VisitorRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit']);
  ensureHeaders_('CompanionRegistrations', ['Timestamp','VisitorName','Phone','PatientName','Bed','Locker','Action','Unit']);
  ensureHeaders_('Users', ['Username','Password','Profile','Email']);
  ensureHeaders_('AuditLog', ['Timestamp','User','Action','Details']);
  ensureHeaders_('Settings', ['Key','Value']);

  const defaults = [
    ['NUM_ARMARIOS_VISITANTE', 20],
    ['NUM_ARMARIOS_VISITANTE_ROWS', 4],
    ['NUM_ARMARIOS_VISITANTE_COLS', 5],
    ['NUM_ARMARIOS_NAC', 20],
    ['NUM_ARMARIOS_NAC_ROWS', 4],
    ['NUM_ARMARIOS_NAC_COLS', 5],
    ['NUM_ARMARIOS_UIB', 20],
    ['NUM_ARMARIOS_UIB_ROWS', 4],
    ['NUM_ARMARIOS_UIB_COLS', 5]
  ];
  defaults.forEach(([k,v])=>{ if (getSetting(k) === null) updateSetting(k, v); });

  // Gera planta se estiver vazia
  generateLockers('VisitorLockers', getSetting('NUM_ARMARIOS_VISITANTE'), '', getSetting('NUM_ARMARIOS_VISITANTE_ROWS'), getSetting('NUM_ARMARIOS_VISITANTE_COLS'));
  generateLockers('CompanionLockersNAC', getSetting('NUM_ARMARIOS_NAC'), 'NAC', getSetting('NUM_ARMARIOS_NAC_ROWS'), getSetting('NUM_ARMARIOS_NAC_COLS'));
  generateLockers('CompanionLockersUIB', getSetting('NUM_ARMARIOS_UIB'), 'UIB', getSetting('NUM_ARMARIOS_UIB_ROWS'), getSetting('NUM_ARMARIOS_UIB_COLS'));

  // Se não existir usuário, cria admin padrão (pode apagar depois)
  const usersSheet = SS.getSheetByName('Users');
  if (usersSheet.getLastRow() === 1) {
    addUser('admin', 'admin', PROFILES.ADMIN, Session.getActiveUser().getEmail() || 'admin@example.com');
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

/** ========= SETTINGS ========= **/
function setLockerConfig(type, rows, cols) {
  let sheetName, prefix, key;
  if (type === 'visitor') { sheetName = 'VisitorLockers'; prefix = ''; key = 'NUM_ARMARIOS_VISITANTE'; }
  else if (type === 'nac') { sheetName = 'CompanionLockersNAC'; prefix = 'NAC'; key = 'NUM_ARMARIOS_NAC'; }
  else if (type === 'uib') { sheetName = 'CompanionLockersUIB'; prefix = 'UIB'; key = 'NUM_ARMARIOS_UIB'; }
  else return { success:false, message:'Tipo inválido' };

  const rowsN = Number(rows), colsN = Number(cols);
  if (!rowsN || !colsN) return { success:false, message:'Rows/Cols inválidos' };

  const count = rowsN * colsN;
  const sheet = SS.getSheetByName(sheetName);
  const currentCount = sheet.getLastRow() - 1;
  if (count < currentCount) return { success:false, message:'Não é possível reduzir armários (existem registros).' };

  updateSetting(key, count);
  updateSetting(key+'_ROWS', rowsN);
  updateSetting(key+'_COLS', colsN);

  generateLockers(sheetName, count, prefix, rowsN, colsN);
  return { success:true, count };
}

function getLockerConfig(type, unit='') {
  let key;
  if (type === 'visitor') key = 'NUM_ARMARIOS_VISITANTE';
  else if (unit === 'NAC') key = 'NUM_ARMARIOS_NAC';
  else if (unit === 'UIB') key = 'NUM_ARMARIOS_UIB';

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
    if (usersSheet.getLastRow() === 0) ensureHeaders_('Users', ['Username','Password','Profile','Email']);
    const data = usersSheet.getDataRange().getValues();
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
    for (let i=1; i<data.length; i++) {
      if (data[i][0] === username && data[i][1] === hash) {
        logAudit(username, 'Login', 'Usuário autenticado');
        return { success:true, profile:data[i][2], username, email: data[i][3] || '' };
      }
    }
    return { success:false, message:'Credenciais inválidas' };
  } catch (err) {
    return { success:false, message:'Erro no login: '+err };
  }
}

function addUser(username, password, profile, email) {
  const sh = SS.getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0] === username) return { success:false, message:'Usuário já existe' };
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
  sh.appendRow([username, hash, profile, email]);
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
  const sheetName = type==='visitor' ? 'VisitorLockers' : (unit==='NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  return SS.getSheetByName(sheetName).getDataRange().getValues();
}

function getLockerStats(type, unit='') {
  const sheetName = type==='visitor' ? 'VisitorLockers' : (unit==='NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  const data = SS.getSheetByName(sheetName).getDataRange().getValues().slice(1);
  let free=0, occupied=0, orange=0, red=0;
  data.forEach(r=>{
    switch(r[1]) {
      case 'Free': free++; break;
      case 'Occupied': occupied++; break;
      case 'Orange': orange++; break;
      case 'Red': red++; break;
    }
  });
  return { free, occupied, orange, red, total:data.length };
}

function registerVisitor(patientName, bed, visitorName, phone, type, unit, lockerNum) {
  const sheetName = type==='visitor' ? 'VisitorLockers' : (unit==='NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  for (let i=1; i<data.length; i++) {
    if (data[i][0] == lockerNum && data[i][1] === 'Free') {
      const start = new Date();
      const end = new Date(start.getTime() + 60*60*1000);
      sheet.getRange(i+1,2).setValue('Occupied');
      sheet.getRange(i+1,3).setValue(visitorName);
      sheet.getRange(i+1,4).setValue(phone);
      sheet.getRange(i+1,5).setValue(patientName);
      sheet.getRange(i+1,6).setValue(bed);
      sheet.getRange(i+1,7).setValue(start);
      sheet.getRange(i+1,10).setValue(end);

      const regSheet = SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
      regSheet.appendRow([new Date(), visitorName, phone, patientName, bed, lockerNum, 'Register', unit]);
      logAudit(Session.getActiveUser().getEmail() || 'anonymous', 'Register', `${visitorName} (${lockerNum})`);
      return { success:true, locker:lockerNum };
    }
  }
  return { success:false, message:'Armário não disponível' };
}

function checkoutLocker(lockerNum, type, unit) {
  const sheetName = type==='visitor' ? 'VisitorLockers' : (unit==='NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] == lockerNum && data[i][1] !== 'Free') {
      const vname = data[i][2], phone=data[i][3];
      // Limpa colunas 2..10 e preserva número
      sheet.getRange(i+1,2,1,9).clearContent();
      sheet.getRange(i+1,1).setValue(lockerNum);
      sheet.getRange(i+1,2).setValue('Free');

      const regSheet = SS.getSheetByName(type==='visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
      regSheet.appendRow([new Date(), vname, phone, '', '', lockerNum, 'Checkout', unit]);
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

/** ========= INCLUDE (para HTML) ========= **/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
