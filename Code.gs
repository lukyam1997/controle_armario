const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAME = 'ControleArmarios';
const LOCKERS_SHEET_NAME = 'ArmariosMonitor';
const HEADERS = [
  'Registrado em',
  'Armario',
  'Unidade',
  'Perfil',
  'Responsavel',
  'Paciente',
  'Contato',
  'Itens Guardados',
  'Status',
  'Encerrado em',
  'Observacoes'
];
const LOCKERS_HEADERS = ['Unidade', 'Armario', 'Descricao', 'Capacidade', 'Observacoes'];

function doGet() {
  ensureSetup_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Controle de Armarios')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSetup_() {
  const sheet = SS.getSheetByName(SHEET_NAME) || SS.insertSheet(SHEET_NAME);
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const current = headerRange.getValues()[0];
  const needHeader = HEADERS.some((header, index) => current[index] !== header);
  if (sheet.getLastRow() === 0 || needHeader) {
    headerRange.setValues([HEADERS]);
  }
  ensureLockersSheet_();
}

function ensureLockersSheet_() {
  const sheet = SS.getSheetByName(LOCKERS_SHEET_NAME) || SS.insertSheet(LOCKERS_SHEET_NAME);
  const headerRange = sheet.getRange(1, 1, 1, LOCKERS_HEADERS.length);
  const current = headerRange.getValues()[0];
  const needHeader = LOCKERS_HEADERS.some((header, index) => current[index] !== header);
  if (sheet.getLastRow() === 0 || needHeader) {
    headerRange.setValues([LOCKERS_HEADERS]);
  }
  if (sheet.getLastRow() === 1) {
    const samples = [
      ['Acolhimento', 'A-01', 'Armário principal da recepção', '', ''],
      ['Acolhimento', 'A-02', '', '', ''],
      ['UTI Adulto', 'U-01', '', '', ''],
      ['UTI Adulto', 'U-02', '', '', '']
    ];
    sheet.getRange(2, 1, samples.length, LOCKERS_HEADERS.length).setValues(samples);
  }
}

function getDashboardData() {
  ensureSetup_();
  const entries = readEntries_();
  syncEntriesIntoLockers_(entries);
  const lockers = getLockers_();
  return buildResponse_(entries, lockers);
}

function getEntries() {
  return getDashboardData();
}

function readEntries_() {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    return [];
  }
  return values.slice(1).map((row, index) => rowToEntry_(row, index + 2));
}

function getLockers_() {
  const sheet = SS.getSheetByName(LOCKERS_SHEET_NAME);
  if (!sheet) {
    return [];
  }
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    return [];
  }
  return values
    .slice(1)
    .map(row => ({
      unidade: cleanString_(row[0]),
      armario: cleanString_(row[1]).toUpperCase(),
      descricao: cleanString_(row[2]),
      capacidade: cleanString_(row[3]),
      observacoes: cleanString_(row[4])
    }))
    .filter(locker => locker.unidade && locker.armario);
}

function syncEntriesIntoLockers_(entries) {
  if (!entries || !entries.length) {
    return;
  }
  const sheet = SS.getSheetByName(LOCKERS_SHEET_NAME);
  if (!sheet) {
    return;
  }
  const lastRow = sheet.getLastRow();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, LOCKERS_HEADERS.length).getValues() : [];
  const existingSet = existingValues.reduce((acc, row) => {
    const key = toLockerKey_(row[0], row[1]);
    if (key) {
      acc[key] = true;
    }
    return acc;
  }, {});
  const rowsToAppend = [];
  entries.forEach(entry => {
    const key = toLockerKey_(entry.unidade, entry.armario);
    if (key && !existingSet[key]) {
      rowsToAppend.push([entry.unidade, entry.armario.toUpperCase(), '', '', '']);
      existingSet[key] = true;
    }
  });
  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, LOCKERS_HEADERS.length).setValues(rowsToAppend);
  }
}

function addEntry(payload) {
  ensureSetup_();
  const sheet = SS.getSheetByName(SHEET_NAME);
  const cleaned = sanitizePayload_(payload);
  if (!cleaned.armario || !cleaned.unidade || !cleaned.perfil || !cleaned.responsavel) {
    throw new Error('Preencha os campos obrigatórios.');
  }
  cleaned.armario = cleaned.armario.toUpperCase();
  ensureLockerExists_(cleaned.unidade, cleaned.armario);
  const row = [
    new Date(),
    cleaned.armario,
    cleaned.unidade,
    cleaned.perfil,
    cleaned.responsavel,
    cleaned.paciente,
    cleaned.contato,
    cleaned.itens,
    'Ativo',
    '',
    cleaned.observacoes
  ];
  sheet.appendRow(row);
  return true;
}

function ensureLockerExists_(unidade, armario) {
  const sheet = SS.getSheetByName(LOCKERS_SHEET_NAME);
  if (!sheet || !unidade || !armario) {
    return;
  }
  const keyToFind = toLockerKey_(unidade, armario);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    sheet.appendRow([unidade, armario, '', '', '']);
    return;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, LOCKERS_HEADERS.length).getValues();
  const exists = values.some(row => toLockerKey_(row[0], row[1]) === keyToFind);
  if (!exists) {
    sheet.appendRow([unidade, armario, '', '', '']);
  }
}

function finalizeEntry(options) {
  ensureSetup_();
  const rowNumber = Number(options && options.id);
  if (!rowNumber || rowNumber < 2) {
    throw new Error('Registro inválido.');
  }
  const sheet = SS.getSheetByName(SHEET_NAME);
  if (rowNumber > sheet.getLastRow()) {
    throw new Error('Registro não encontrado.');
  }
  const range = sheet.getRange(rowNumber, 1, 1, HEADERS.length);
  const row = range.getValues()[0];
  if (!row || !row.length) {
    throw new Error('Registro não encontrado.');
  }
  row[8] = 'Finalizado';
  row[9] = new Date();
  const extraNote = cleanString_(options && options.nota);
  if (extraNote) {
    const existing = cleanString_(row[10]);
    const formatted = 'Encerramento: ' + extraNote;
    row[10] = existing ? existing + '\n' + formatted : formatted;
  }
  range.setValues([row]);
  return true;
}

function reopenEntry(options) {
  ensureSetup_();
  const rowNumber = Number(options && options.id);
  if (!rowNumber || rowNumber < 2) {
    throw new Error('Registro inválido.');
  }
  const sheet = SS.getSheetByName(SHEET_NAME);
  if (rowNumber > sheet.getLastRow()) {
    throw new Error('Registro não encontrado.');
  }
  const range = sheet.getRange(rowNumber, 1, 1, HEADERS.length);
  const row = range.getValues()[0];
  if (!row || !row.length) {
    throw new Error('Registro não encontrado.');
  }
  row[8] = 'Ativo';
  row[9] = '';
  range.setValues([row]);
  return true;
}

function deleteEntry(options) {
  ensureSetup_();
  const rowNumber = Number(options && options.id);
  if (!rowNumber || rowNumber < 2) {
    throw new Error('Registro inválido.');
  }
  const sheet = SS.getSheetByName(SHEET_NAME);
  if (rowNumber > sheet.getLastRow()) {
    throw new Error('Registro não encontrado.');
  }
  sheet.deleteRow(rowNumber);
  return true;
}

function buildResponse_(entries, lockers) {
  const summary = {
    total: entries.length,
    ativos: 0,
    finalizados: 0,
    disponiveis: 0,
    ocupados: 0,
    porPerfil: {},
    unidades: [],
    perfis: [],
    ocupacaoPercentual: 0
  };
  const unidadesSet = {};
  const perfisSet = {};
  const ativosPorArmario = {};
  entries.forEach(entry => {
    const status = (entry.status || '').toLowerCase();
    if (status === 'finalizado') {
      summary.finalizados++;
    } else {
      summary.ativos++;
      const key = toLockerKey_(entry.unidade, entry.armario);
      if (key && (!ativosPorArmario[key] || entry.id > ativosPorArmario[key].id)) {
        ativosPorArmario[key] = entry;
      }
    }
    const perfil = entry.perfil || 'Não informado';
    summary.porPerfil[perfil] = (summary.porPerfil[perfil] || 0) + 1;
    if (entry.unidade) {
      unidadesSet[entry.unidade] = true;
    }
    if (entry.perfil) {
      perfisSet[entry.perfil] = true;
    }
  });

  const decoratedLockers = lockers.map(locker => {
    unidadesSet[locker.unidade] = true;
    const key = toLockerKey_(locker.unidade, locker.armario);
    const registro = key ? ativosPorArmario[key] || null : null;
    const status = registro ? 'Ocupado' : 'Livre';
    if (status === 'Ocupado') {
      summary.ocupados++;
      if (registro && registro.perfil) {
        perfisSet[registro.perfil] = true;
      }
    } else {
      summary.disponiveis++;
    }
    return Object.assign({}, locker, {
      status,
      registro
    });
  });

  if (!decoratedLockers.length && Object.keys(ativosPorArmario).length) {
    Object.values(ativosPorArmario).forEach(entry => {
      const key = toLockerKey_(entry.unidade, entry.armario);
      if (!key) {
        return;
      }
      summary.ocupados++;
      decoratedLockers.push({
        unidade: entry.unidade,
        armario: entry.armario,
        descricao: '',
        capacidade: '',
        observacoes: '',
        status: 'Ocupado',
        registro: entry
      });
    });
  }

  if (decoratedLockers.length) {
    summary.ocupacaoPercentual = Math.round((summary.ocupados / decoratedLockers.length) * 1000) / 10;
  }

  summary.unidades = Object.keys(unidadesSet).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  summary.perfis = Object.keys(perfisSet).sort((a, b) => a.localeCompare(b, 'pt-BR'));

  return { entries, lockers: decoratedLockers, summary };
}

function rowToEntry_(row, rowNumber) {
  const createdAt = normalizeDate_(row[0]);
  const closedAt = normalizeDate_(row[9]);
  return {
    id: rowNumber,
    armario: cleanString_(row[1]),
    unidade: cleanString_(row[2]),
    perfil: cleanString_(row[3]),
    responsavel: cleanString_(row[4]),
    paciente: cleanString_(row[5]),
    contato: cleanString_(row[6]),
    itens: cleanString_(row[7]),
    status: cleanString_(row[8]) || 'Ativo',
    observacoes: cleanString_(row[10]),
    criadoEm: createdAt ? formatDate_(createdAt) : '',
    encerradoEm: closedAt ? formatDate_(closedAt) : ''
  };
}

function toLockerKey_(unidade, armario) {
  if (!unidade || !armario) {
    return '';
  }
  return [cleanString_(unidade).toLowerCase(), cleanString_(armario).toLowerCase()].join('::');
}

function sanitizePayload_(payload) {
  return {
    armario: cleanString_(payload && payload.armario),
    unidade: cleanString_(payload && payload.unidade),
    perfil: cleanString_(payload && payload.perfil),
    responsavel: cleanString_(payload && payload.responsavel),
    paciente: cleanString_(payload && payload.paciente),
    contato: cleanString_(payload && payload.contato),
    itens: cleanString_(payload && payload.itens),
    observacoes: cleanString_(payload && payload.observacoes)
  };
}

function cleanString_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  return String(value).trim();
}

function normalizeDate_(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDate_(date) {
  if (!date) {
    return '';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
