const SS = SpreadsheetApp.getActiveSpreadsheet();
const SPREADSHEET_ID = SS.getId();

const PROFILES = Object.freeze({
  VISITOR: 'armario_visitante',
  COMPANION: 'guarda_volume',
  ADMIN: 'admin'
});

const STATUS = Object.freeze({
  FREE: 'Free',
  OCCUPIED: 'Occupied',
  ORANGE: 'Orange',
  RED: 'Red'
});

const LOCKER_SHEETS = Object.freeze({
  visitor: {
    default: { lockers: 'VisitorLockers', registrations: 'VisitorRegistrations' }
  },
  companion: {
    NAC: { lockers: 'CompanionLockersNAC', registrations: 'CompanionRegistrations' },
    UIB: { lockers: 'CompanionLockersUIB', registrations: 'CompanionRegistrations' },
    default: { lockers: 'CompanionLockersUIB', registrations: 'CompanionRegistrations' }
  }
});

const SETTINGS_SHEET = 'Settings';
const USERS_HEADERS = ['Username', 'PasswordHash', 'Profile', 'Email'];
const LOCKER_HEADERS = [
  'LockerNumber', 'Status', 'VisitorName', 'Phone', 'PatientName', 'Bed',
  'StartTime', 'MessageWelcomeSent', 'MessageAlertSent', 'EndTime', 'Unit'
];
const REGISTRATION_HEADERS = [
  'Timestamp', 'VisitorName', 'Phone', 'PatientName', 'Bed',
  'LockerNumber', 'Action', 'Unit'
];
const AUDIT_HEADERS = ['Timestamp', 'User', 'Action', 'Details'];

const TYPE_ALIASES = Object.freeze({
  visitor: 'visitor',
  armario_visitante: 'visitor',
  companion: 'companion',
  guarda_volume: 'companion'
});

const DEFAULT_LOCKER_DURATION_MINUTES = 60;
const DEFAULT_ALERT_LEAD_TIME_MINUTES = 10;
const DEFAULT_RED_ESCALATION_MINUTES = 10;
const DEFAULT_REGISTRATION_RETENTION_HOURS = 24;

let cachedSettings = null;

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template
    .evaluate()
    .setTitle('Hospital Storage Manager')
    .setFaviconUrl('https://example.com/favicon.ico');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSystem() {
  cachedSettings = null;
  generateLockers('VisitorLockers', getNumericSetting('NUM_ARMARIOS_VISITANTE', 0), '');
  generateLockers('CompanionLockersNAC', getNumericSetting('NUM_ARMARIOS_NAC', 0), 'NAC');
  generateLockers('CompanionLockersUIB', getNumericSetting('NUM_ARMARIOS_UIB', 0), 'UIB');

  ['VisitorLockers', 'CompanionLockersNAC', 'CompanionLockersUIB'].forEach(setConditionalFormatting);

  ensureSheetExists('VisitorRegistrations', REGISTRATION_HEADERS);
  ensureSheetExists('CompanionRegistrations', REGISTRATION_HEADERS);
  ensureSheetExists('AuditLog', AUDIT_HEADERS);
  const usersSheet = ensureSheetExists('Users', USERS_HEADERS);

  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow([
      'admin',
      hashPassword('senha123'),
      PROFILES.ADMIN,
      'admin@hospital.com'
    ]);
  }

  const hasExistingTrigger = ScriptApp.getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === 'checkLockersAndSendAlerts');

  if (!hasExistingTrigger) {
    ScriptApp.newTrigger('checkLockersAndSendAlerts')
      .timeBased()
      .everyMinutes(5)
      .create();
  }
}

function generateLockers(sheetName, quantity, unit) {
  const sheet = ensureSheetExists(sheetName, LOCKER_HEADERS);
  const targetQuantity = Math.max(0, Math.floor(Number(quantity) || 0));
  const currentLockers = Math.max(sheet.getLastRow() - 1, 0);

  for (let i = currentLockers + 1; i <= targetQuantity; i++) {
    sheet.appendRow([
      i,
      STATUS.FREE,
      '', '', '', '', '',
      'No',
      'No',
      '',
      unit
    ]);
  }

  const totalRows = Math.max(sheet.getLastRow() - 1, 0);
  if (totalRows > 0) {
    const unitValues = Array.from({ length: totalRows }, () => [unit || '']);
    sheet.getRange(2, 11, totalRows, 1).setValues(unitValues);
  }
}

function setConditionalFormatting(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    return;
  }

  const range = sheet.getDataRange();
  const rules = [];

  const statusColors = {
    [STATUS.ORANGE]: '#FFA500',
    [STATUS.RED]: '#FF0000'
  };

  Object.keys(statusColors).forEach(status => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(statusColors[status])
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
}

function getSetting(key, defaultValue = null) {
  if (!cachedSettings) {
    const sheet = SS.getSheetByName(SETTINGS_SHEET);
    cachedSettings = {};
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      data.forEach(row => {
        const settingKey = row[0];
        if (settingKey) {
          cachedSettings[settingKey] = row[1];
        }
      });
    }
  }

  return cachedSettings.hasOwnProperty(key) ? cachedSettings[key] : defaultValue;
}

function getNumericSetting(key, fallback) {
  const value = Number(getSetting(key));
  return Number.isFinite(value) ? value : fallback;
}

function login(username, password) {
  const usersSheet = ensureSheetExists('Users', USERS_HEADERS);
  const data = usersSheet.getDataRange().getValues();
  const providedHash = hashPassword(password);
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === providedHash) {
      return {
        success: true,
        profile: data[i][2],
        username: username
      };
    }
  }
  return { success: false };
}

function resetPassword(username, newPassword) {
  const usersSheet = ensureSheetExists('Users', USERS_HEADERS);
  const data = usersSheet.getDataRange().getValues();
  const newHash = hashPassword(newPassword);
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      usersSheet.getRange(i + 1, 2).setValue(newHash);
      const email = data[i][3];
      if (email) {
        MailApp.sendEmail(email, 'Senha redefinida', 'Sua nova senha: ' + newPassword);
      }
      logAudit(getCurrentUserEmail(), 'Reset Password', username);
      return true;
    }
  }
  return false;
}

function addUser(username, password, profile, email) {
  const usersSheet = ensureSheetExists('Users', USERS_HEADERS);
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      throw new Error('Usuário já existe.');
    }
  }
  usersSheet.appendRow([username, hashPassword(password), profile, email]);
  logAudit(getCurrentUserEmail(), 'Add User', username);
}

function deleteUser(username) {
  const usersSheet = ensureSheetExists('Users', USERS_HEADERS);
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      usersSheet.deleteRow(i + 1);
      logAudit(getCurrentUserEmail(), 'Delete User', username);
      return true;
    }
  }
  return false;
}

function registerVisitor(patientName, bed, visitorName, phone, type, unit) {
  const context = resolveLockerContext(type, unit);
  const sheet = SS.getSheetByName(context.lockers);
  if (!sheet) {
    throw new Error('Planilha de armários não encontrada.');
  }

  const data = sheet.getDataRange().getValues();
  let lockerRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === STATUS.FREE) {
      lockerRow = i + 1;
      break;
    }
  }

  if (lockerRow === -1) {
    throw new Error('Sem armários disponíveis.');
  }

  const lockerNum = sheet.getRange(lockerRow, 1).getValue();
  const startTime = new Date();
  const endTime = new Date(startTime.getTime() + minutesToMillis(getNumericSetting('DURACAO_PADRAO_MIN', DEFAULT_LOCKER_DURATION_MINUTES)));
  const sanitizedPhone = String(phone || '').trim();

  sheet.getRange(lockerRow, 2, 1, 9).setValues([[
    STATUS.OCCUPIED,
    visitorName,
    sanitizedPhone,
    patientName,
    bed,
    startTime,
    'No',
    'No',
    endTime
  ]]);

  if (typeof unit === 'string') {
    sheet.getRange(lockerRow, 11).setValue(unit);
  }

  const welcomeSent = sanitizedPhone && sendWhatsApp(sanitizedPhone, getSetting('WHATSAPP_TEMPLATE_BOASVINDAS'), [visitorName, lockerNum]);
  sheet.getRange(lockerRow, 8).setValue(welcomeSent ? 'Yes' : 'No');

  const regSheet = ensureSheetExists(context.registrations, REGISTRATION_HEADERS);
  regSheet.appendRow([
    new Date(),
    visitorName,
    sanitizedPhone,
    patientName,
    bed,
    lockerNum,
    'Register',
    unit || ''
  ]);

  logAudit(getCurrentUserEmail(), 'Register', visitorName);
  return lockerNum;
}

function checkoutLocker(lockerNum, type, unit) {
  const context = resolveLockerContext(type, unit);
  const sheet = SS.getSheetByName(context.lockers);
  if (!sheet) {
    throw new Error('Planilha de armários não encontrada.');
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(lockerNum) && data[i][1] !== STATUS.FREE) {
      const row = i + 1;
      const visitorName = data[i][2];
      const phone = data[i][3];

      sheet.getRange(row, 2, 1, 9).setValues([[
        STATUS.FREE,
        '', '', '', '', '',
        'No',
        'No',
        ''
      ]]);
      sheet.getRange(row, 11).setValue(unit || '');

      const regSheet = ensureSheetExists(context.registrations, REGISTRATION_HEADERS);
      regSheet.appendRow([
        new Date(),
        visitorName,
        phone,
        '',
        '',
        lockerNum,
        'Checkout',
        unit || ''
      ]);

      logAudit(getCurrentUserEmail(), 'Checkout', visitorName);
      return true;
    }
  }
  return false;
}

function checkLockersAndSendAlerts() {
  const now = new Date();
  const alertLeadTime = minutesToMillis(getNumericSetting('ALERTA_MINUTOS', DEFAULT_ALERT_LEAD_TIME_MINUTES));
  const redEscalationTime = minutesToMillis(getNumericSetting('ATRASO_VERMELHO_MIN', DEFAULT_RED_ESCALATION_MINUTES));

  ['VisitorLockers', 'CompanionLockersNAC', 'CompanionLockersUIB'].forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const status = data[i][1];
      if (status === STATUS.FREE) {
        continue;
      }

      const startTime = new Date(data[i][6]);
      const endTime = new Date(data[i][9]);
      const visitorName = data[i][2];
      const phone = data[i][3];
      const lockerNum = data[i][0];
      const alertSent = data[i][8];
      const row = i + 1;
      const timeLeft = endTime.getTime() - now.getTime();

      if (timeLeft <= alertLeadTime && timeLeft > 0 && alertSent !== 'Yes') {
        const alertWasSent = phone && sendWhatsApp(phone, getSetting('WHATSAPP_TEMPLATE_ALERTA'), [visitorName, lockerNum]);
        if (alertWasSent) {
          sheet.getRange(row, 9).setValue('Yes');
        }
      }

      if (now > endTime && status !== STATUS.ORANGE && status !== STATUS.RED) {
        sheet.getRange(row, 2).setValue(STATUS.ORANGE);
      }

      if (now.getTime() > endTime.getTime() + redEscalationTime && status !== STATUS.RED) {
        sheet.getRange(row, 2).setValue(STATUS.RED);
        const alertEmail = getSetting('EMAIL_ALERTA');
        if (alertEmail) {
          MailApp.sendEmail(
            alertEmail,
            'Alerta Armário Vencido',
            'Armário ' + lockerNum + ' em ' + sheetName + ' está vencido há mais de ' + getNumericSetting('ATRASO_VERMELHO_MIN', DEFAULT_RED_ESCALATION_MINUTES) + ' minutos. Visitante: ' + visitorName + '.'
          );
        }
      }
    }
  });

  cleanOldRegistrations();
}

function cleanOldRegistrations() {
  const retentionMillis = hoursToMillis(getNumericSetting('RETENCAO_REGISTROS_H', DEFAULT_REGISTRATION_RETENTION_HOURS));
  if (!Number.isFinite(retentionMillis) || retentionMillis <= 0) {
    return;
  }
  const threshold = Date.now() - retentionMillis;
  ['VisitorRegistrations', 'CompanionRegistrations'].forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) {
      const timestamp = new Date(data[i][0]).getTime();
      if (!isNaN(timestamp) && timestamp < threshold) {
        sheet.deleteRow(i + 1);
      }
    }
  });
}

function getLockersData(type, unit) {
  const context = resolveLockerContext(type, unit);
  const sheet = SS.getSheetByName(context.lockers);
  if (!sheet) {
    throw new Error('Planilha de armários não encontrada.');
  }
  return sheet.getDataRange().getValues();
}

function searchRegistrations(query, type) {
  const context = resolveLockerContext(type, '');
  const sheet = SS.getSheetByName(context.registrations);
  if (!sheet) {
    return [];
  }
  const lowercaseQuery = String(query || '').toLowerCase();
  if (!lowercaseQuery) {
    return sheet.getDataRange().getValues();
  }
  return sheet
    .getDataRange()
    .getValues()
    .filter(row => row.join(' ').toLowerCase().includes(lowercaseQuery));
}

function exportReport(type) {
  const context = resolveLockerContext(type, '');
  const sheet = SS.getSheetByName(context.registrations);
  if (!sheet) {
    throw new Error('Planilha não encontrada.');
  }
  return sheet
    .getDataRange()
    .getValues()
    .map(row => row.join(','))
    .join('\n');
}

function sendWhatsApp(phone, templateName, parameters) {
  if (!phone || !templateName) {
    return false;
  }

  const phoneId = getSetting('WHATSAPP_PHONE_ID');
  const token = getSetting('WHATSAPP_TOKEN');
  if (!phoneId || !token) {
    return false;
  }

  const apiUrl = 'https://graph.facebook.com/v13.0/' + phoneId + '/messages';
  const payload = {
    messaging_product: 'whatsapp',
    to: String(phone).replace(/[^\d]/g, ''),
    type: 'template',
    template: {
      name: templateName,
      language: { code: 'pt_BR' },
      components: [
        {
          type: 'body',
          parameters: (parameters || []).map(text => ({ type: 'text', text: text }))
        }
      ]
    }
  };

  try {
    UrlFetchApp.fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return true;
  } catch (error) {
    logAudit('system', 'WhatsApp Error', error && error.message ? error.message : String(error));
    return false;
  }
}

function logAudit(user, action, details) {
  const sheet = ensureSheetExists('AuditLog', AUDIT_HEADERS);
  sheet.appendRow([new Date(), user, action, details]);
}

function resolveLockerContext(type, unit) {
  const normalizedTypeInput = String(type || '').toLowerCase();
  const normalizedType = TYPE_ALIASES[normalizedTypeInput] || normalizedTypeInput;
  const normalizedUnit = String(unit || '').toUpperCase();
  const typeConfig = LOCKER_SHEETS[normalizedType];
  if (!typeConfig) {
    throw new Error('Tipo de armário inválido: ' + type);
  }
  const unitConfig = typeConfig[normalizedUnit] || typeConfig.default;
  if (!unitConfig) {
    throw new Error('Unidade inválida para o tipo informado.');
  }
  return unitConfig;
}

function ensureSheetExists(sheetName, headers) {
  const sheet = SS.getSheetByName(sheetName) || SS.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const currentHeaders = headerRange.getValues()[0];
    const needsUpdate = headers.some((header, index) => currentHeaders[index] !== header);
    if (needsUpdate) {
      headerRange.setValues([headers]);
    }
  }
  return sheet;
}

function hashPassword(password) {
  const normalizedPassword = password == null ? '' : String(password);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalizedPassword);
  return Utilities.base64Encode(digest);
}

function getCurrentUserEmail() {
  const user = Session.getActiveUser();
  return user ? user.getEmail() : 'system';
}

function minutesToMillis(minutes) {
  const value = Number(minutes);
  return Number.isFinite(value) ? value * 60 * 1000 : 0;
}

function hoursToMillis(hours) {
  const value = Number(hours);
  return Number.isFinite(value) ? value * 60 * 60 * 1000 : 0;
}
