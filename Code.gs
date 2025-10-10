const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAME = 'ControleArmarios';
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
}

function getEntries() {
  ensureSetup_();
  const sheet = SS.getSheetByName(SHEET_NAME);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    return buildResponse_([]);
  }
  const data = values.slice(1).map((row, index) => rowToEntry_(row, index + 2));
  return buildResponse_(data);
}

function addEntry(payload) {
  ensureSetup_();
  const sheet = SS.getSheetByName(SHEET_NAME);
  const cleaned = sanitizePayload_(payload);
  if (!cleaned.armario || !cleaned.unidade || !cleaned.perfil || !cleaned.responsavel) {
    throw new Error('Preencha os campos obrigatórios.');
  }
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

function buildResponse_(entries) {
  const summary = {
    total: entries.length,
    ativos: 0,
    finalizados: 0,
    porPerfil: {},
    unidades: []
  };
  const unidadesSet = {};
  entries.forEach(entry => {
    if (entry.status === 'Ativo') {
      summary.ativos++;
    } else {
      summary.finalizados++;
    }
    const perfil = entry.perfil || 'Não informado';
    summary.porPerfil[perfil] = (summary.porPerfil[perfil] || 0) + 1;
    if (entry.unidade) {
      unidadesSet[entry.unidade] = true;
    }
  });
  summary.unidades = Object.keys(unidadesSet).sort();
  return { entries, summary };
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
