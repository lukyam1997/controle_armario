const SS = SpreadsheetApp.getActiveSpreadsheet();
const SPREADSHEET_ID = SS.getId(); // Ou cole o ID manual.

const PROFILES = {
  VISITOR: 'armario_visitante',
  COMPANION: 'guarda_volume',
  ADMIN: 'admin'
};

// Função principal para web app
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate().setTitle('Hospital Storage Manager').setFaviconUrl('https://example.com/favicon.ico');
}

// Incluir arquivos HTML (crie arquivos separados no Apps Script ou use inline)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Setup inicial: Rode uma vez
function setupSystem() {
  // Gerar armários automaticamente
  generateLockers('VisitorLockers', getSetting('NUM_ARMARIOS_VISITANTE'), '');
  generateLockers('CompanionLockersNAC', getSetting('NUM_ARMARIOS_NAC'), 'NAC');
  generateLockers('CompanionLockersUIB', getSetting('NUM_ARMARIOS_UIB'), 'UIB');
  
  // Configurar formatação condicional nos painéis (para visualização na planilha)
  setConditionalFormatting('VisitorLockers');
  setConditionalFormatting('CompanionLockersNAC');
  setConditionalFormatting('CompanionLockersUIB');
  
  // Criar trigger para verificações a cada 5 min
  ScriptApp.newTrigger('checkLockersAndSendAlerts')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  // Hash senha admin inicial se não existir
  const usersSheet = SS.getSheetByName('Users');
  if (usersSheet.getLastRow() === 1) {
    usersSheet.appendRow(['admin', Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'senha123')), PROFILES.ADMIN, 'admin@hospital.com']);
  }
}

function generateLockers(sheetName, num, unit) {
  const sheet = SS.getSheetByName(sheetName) || SS.insertSheet(sheetName);
  if (sheet.getLastRow() === 1) {
    sheet.appendRow(['LockerNumber', 'Status', 'VisitorName', 'Phone', 'PatientName', 'Bed', 'StartTime', 'MessageWelcomeSent', 'MessageAlertSent', 'EndTime', 'Unit']);
    for (let i = 1; i <= num; i++) {
      sheet.appendRow([i, 'Free', '', '', '', '', '', 'No', 'No', '', unit]);
    }
  }
}

function setConditionalFormatting(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  const range = sheet.getDataRange();
  const rules = [];
  
  // Laranja se Status = Orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Orange')
    .setBackground('#FFA500')
    .setRanges([range])
    .build());
  
  // Vermelho se Status = Red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Red')
    .setBackground('#FF0000')
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
}

function getSetting(key) {
  const sheet = SS.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for (let row of data) {
    if (row[0] === key) return row[1];
  }
  return null;
}

// Login: Chamado do client-side
function login(username, password) {
  const usersSheet = SS.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === hash) {
      return { success: true, profile: data[i][2], username: username };
    }
  }
  return { success: false };
}

// Reset senha (admin ou via e-mail)
function resetPassword(username, newPassword) {
  // Verifique se caller é admin (implemente checagem)
  const usersSheet = SS.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, newPassword));
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      usersSheet.getRange(i+1, 2).setValue(hash);
      MailApp.sendEmail(data[i][3], 'Senha Resetada', 'Sua nova senha: ' + newPassword);
      logAudit('Admin', 'Reset Password', username);
      return true;
    }
  }
  return false;
}

// Adicionar usuário (admin)
function addUser(username, password, profile, email) {
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
  SS.getSheetByName('Users').appendRow([username, hash, profile, email]);
  logAudit('Admin', 'Add User', username);
}

// Deletar usuário (admin)
function deleteUser(username) {
  const usersSheet = SS.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      usersSheet.deleteRow(i+1);
      logAudit('Admin', 'Delete User', username);
      return true;
    }
  }
  return false;
}

// Registro de visitante/acompanhante
function registerVisitor(patientName, bed, visitorName, phone, type, unit = '') {
  let sheetName = type === 'visitor' ? 'VisitorLockers' : (unit === 'NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  let lockerNum = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === 'Free') {
      lockerNum = data[i][0];
      const startTime = new Date();
      const endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // +1 hora
      sheet.getRange(i+1, 2).setValue('Occupied'); // Status
      sheet.getRange(i+1, 3).setValue(visitorName);
      sheet.getRange(i+1, 4).setValue(phone);
      sheet.getRange(i+1, 5).setValue(patientName);
      sheet.getRange(i+1, 6).setValue(bed);
      sheet.getRange(i+1, 7).setValue(startTime);
      sheet.getRange(i+1, 8).setValue('No'); // Welcome sent
      sheet.getRange(i+1, 9).setValue('No'); // Alert sent
      sheet.getRange(i+1, 10).setValue(endTime);
      break;
    }
  }
  if (!lockerNum) throw 'Sem armários disponíveis!';
  
  // Enviar boas-vindas se phone
  if (phone) sendWhatsApp(phone, getSetting('WHATSAPP_TEMPLATE_BOASVINDAS'), [visitorName, lockerNum]);
  
  // Log registro
  const regSheet = SS.getSheetByName(type === 'visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
  regSheet.appendRow([new Date(), visitorName, phone, patientName, bed, lockerNum, 'Register', unit]);
  logAudit(Session.getActiveUser().getEmail(), 'Register', visitorName);
  
  return lockerNum;
}

// Checkout / Liberar armário
function checkoutLocker(lockerNum, type, unit = '') {
  let sheetName = type === 'visitor' ? 'VisitorLockers' : (unit === 'NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  const sheet = SS.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == lockerNum && data[i][1] !== 'Free') {
      const visitorName = data[i][2];
      const phone = data[i][3];
      sheet.getRange(i+1, 2).setValue('Free');
      sheet.getRange(i+1, 3, 10).clearContent(); // Limpa dados
      // Log
      const regSheet = SS.getSheetByName(type === 'visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
      regSheet.appendRow([new Date(), visitorName, phone, '', '', lockerNum, 'Checkout', unit]);
      logAudit(Session.getActiveUser().getEmail(), 'Checkout', visitorName);
      return true;
    }
  }
  return false;
}

// Verificar armários e enviar alertas (trigger)
function checkLockersAndSendAlerts() {
  const now = new Date();
  ['VisitorLockers', 'CompanionLockersNAC', 'CompanionLockersUIB'].forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === 'Occupied' || data[i][1] === 'Orange') {
        const startTime = new Date(data[i][6]);
        const endTime = new Date(data[i][9]);
        const timeLeft = endTime - now;
        const visitorName = data[i][2];
        const phone = data[i][3];
        const lockerNum = data[i][0];
        
        // 10 min antes: Alerta
        if (timeLeft <= 10 * 60 * 1000 && timeLeft > 0 && data[i][8] === 'No') {
          if (phone) sendWhatsApp(phone, getSetting('WHATSAPP_TEMPLATE_ALERTA'), [visitorName, lockerNum]);
          sheet.getRange(i+1, 9).setValue('Yes');
        }
        
        // Vencido: Laranja
        if (now > endTime && data[i][1] !== 'Orange' && data[i][1] !== 'Red') {
          sheet.getRange(i+1, 2).setValue('Orange');
        }
        
        // +10 min vencido: Vermelho + alerta e-mail
        if (now > endTime + 10 * 60 * 1000 && data[i][1] !== 'Red') {
          sheet.getRange(i+1, 2).setValue('Red');
          MailApp.sendEmail(getSetting('EMAIL_ALERTA'), 'Alerta Armário Vencido', `Armário ${lockerNum} em ${sheetName} está vencido há +10 min. Verificar com ${visitorName}.`);
        }
      }
    }
  });
  
  // Limpeza automática: Remover registros > 1 dia
  cleanOldRegistrations();
}

function cleanOldRegistrations() {
  ['VisitorRegistrations', 'CompanionRegistrations'].forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    for (let i = data.length - 1; i > 0; i--) {
      if (new Date(data[i][0]) < now - 24 * 60 * 60 * 1000) {
        sheet.deleteRow(i+1);
      }
    }
  });
}

// Enviar WhatsApp
function sendWhatsApp(phone, templateName, parameters) {
  const apiUrl = `https://graph.facebook.com/v13.0/${getSetting('WHATSAPP_PHONE_ID')}/messages`;
  const payload = {
    messaging_product: 'whatsapp',
    to: phone.replace(/[^\d]/g, ''),
    type: 'template',
    template: {
      name: templateName,
      language: { code: 'pt_BR' }, // Português Brasil
      components: [{ type: 'body', parameters: parameters.map(p => ({ type: 'text', text: p })) }]
    }
  };
  UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${getSetting('WHATSAPP_TOKEN')}`, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload)
  });
}

// Log auditoria
function logAudit(user, action, details) {
  SS.getSheetByName('AuditLog').appendRow([new Date(), user, action, details]);
}

// Obter dados para painel (chamado do client para HTML table)
function getLockersData(type, unit = '') {
  let sheetName = type === 'visitor' ? 'VisitorLockers' : (unit === 'NAC' ? 'CompanionLockersNAC' : 'CompanionLockersUIB');
  return SS.getSheetByName(sheetName).getDataRange().getValues();
}

// Busca registros (extra)
function searchRegistrations(query, type) {
  const sheet = SS.getSheetByName(type === 'visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
  const data = sheet.getDataRange().getValues();
  return data.filter(row => row.join('').toLowerCase().includes(query.toLowerCase()));
}

// Export relatório (extra, gera CSV)
function exportReport(type) {
  const sheet = SS.getSheetByName(type === 'visitor' ? 'VisitorRegistrations' : 'CompanionRegistrations');
  const csv = sheet.getDataRange().getValues().map(row => row.join(',')).join('\n');
  return csv;
}
