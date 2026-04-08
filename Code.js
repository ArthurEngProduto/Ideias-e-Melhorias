/**
 * Portal de visualização de respostas do Google Forms em formato de cards.
 *
 * IMPORTANTE:
 * - Se o projeto estiver VINCULADO à planilha, pode deixar SPREADSHEET_ID vazio.
 * - Se o projeto estiver INDEPENDENTE, preencha SPREADSHEET_ID com o ID da planilha.
 */
const SPREADSHEET_ID = ''; // Ex: '1AbC...xyz'
const SHEET_NAME = ''; // vazio = primeira aba
const DEFAULT_PAGE_SIZE = 20;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal de Ideias e Melhorias')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPortalBootstrap() {
  const rows = getRows_();

  return {
    totalRecords: rows.length,
    fields: rows.length ? Object.keys(rows[0]) : [],
    options: getFilterOptions_(),
  };
}

function getPortalData(filters, page, pageSize) {
  const rows = getRows_();
  const normalizedFilters = normalizeFilters_(filters || {});

  const filteredRows = rows.filter((row) => {
    const rowDate = normalizeDateFilter_(getFieldValue_(row, 'Carimbo de data/hora'));
    if (normalizedFilters.timestampStart && (!rowDate || rowDate < normalizedFilters.timestampStart)) return false;
    if (normalizedFilters.timestampEnd && (!rowDate || rowDate > normalizedFilters.timestampEnd)) return false;
    if (normalizedFilters.name && !getFieldValue_(row, 'Digite seu nome:').toLowerCase().includes(normalizedFilters.name)) return false;
    if (normalizedFilters.sector && getFieldValue_(row, 'Selecione o seu setor:') !== normalizedFilters.sector) return false;
    if (normalizedFilters.reference && getFieldValue_(row, 'Este registro se refere a:') !== normalizedFilters.reference) return false;
  
    return true;
  });

  const safePageSize = Number(pageSize) > 0 ? Number(pageSize) : DEFAULT_PAGE_SIZE;
  const requestedPage = Number(page) > 0 ? Number(page) : 1;
  const total = filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(total / safePageSize));
  const safePage = Math.min(requestedPage, totalPages);
  const pagedRows = filteredRows.slice((safePage - 1) * safePageSize, safePage * safePageSize);

  return {
    page: safePage,
    pageSize: safePageSize,
    total,
    totalPages,
    rows: pagedRows,
  };
}

function getRows_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];

  const rawHeaders = values[0];
  const headers = rawHeaders.map((h) => normalizeHeader_(h));
  const dataRows = values.slice(1);

  return dataRows
    .filter((row) => row.some((cell) => String(cell || '').trim() !== ''))
    .map((row, index) => {
      const obj = {};
      headers.forEach((header, headerIndex) => {
        obj[header] = row[headerIndex] || '';
      });
      obj.__rowNumber = index + 2;
      return obj;
    });
}

function updateConcludedStatus(rowNumber, concluded) {
  const numericRow = Number(rowNumber);
  if (!numericRow || numericRow < 2) {
    throw new Error('Linha inválida para atualização.');
  }

  const sheet = getSheet_();
  const columnR = 18;
  const value = concluded ? 'OK' : '';
  sheet.getRange(numericRow, columnR).setValue(value);

  return { success: true, rowNumber: numericRow, concluded: value };
}

function deletePortalRow(rowNumber) {
  const numericRow = Number(rowNumber);
  if (!numericRow || numericRow < 2) {
    throw new Error('Linha inválida para exclusão.');
  }

  const sheet = getSheet_();
  if (numericRow > sheet.getLastRow()) {
    throw new Error('A linha informada não existe mais na planilha.');
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map((header) => normalizeHeader_(header));
  const rowValues = sheet.getRange(numericRow, 1, 1, headers.length).getDisplayValues()[0];
  const rowObject = {};
  headers.forEach((header, index) => {
    rowObject[header] = rowValues[index] || '';
  });

  const deletionSummary = deleteFormResponseForRow_(rowObject);
  sheet.deleteRow(numericRow);

  return {
    success: true,
    rowNumber: numericRow,
    deletedFromForm: deletionSummary.deletedFromForm,
    deletedFromFormResponsesSheet: deletionSummary.deletedFromResponsesSheet,
  };
}

function getSheet_() {
  const ss = resolveSpreadsheet_();
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

  if (!sheet) {
    throw new Error('A aba configurada não foi encontrada. Revise SHEET_NAME em Code.gs.');
  }

  return sheet;
}

function resolveSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  throw new Error(
    'Não foi possível localizar a planilha ativa. Se o script for independente, preencha SPREADSHEET_ID em Code.gs.'
  );
}

function deleteFormResponseForRow_(row) {
  const form = getLinkedForm_();
  if (!form) {
    return {
      deletedFromForm: false,
      deletedFromResponsesSheet: false,
    };
  }

  let deletedFromForm = false;

  const responseId = findResponseId_(row);
  if (responseId) {
    form.deleteResponse(responseId);
    deletedFromForm = true;
  } else {
    const response = findFormResponseByTimestamp_(form, row);
    if (response) {
      form.deleteResponse(response.getId());
      deletedFromForm = true;
    }
  }

  const deletedFromResponsesSheet = deleteFromLinkedResponsesSheet_(form, row);

  return {
    deletedFromForm,
    deletedFromResponsesSheet,
  };
}

function getLinkedForm_() {
  const spreadsheet = resolveSpreadsheet_();
  const formUrl = spreadsheet.getFormUrl();
  if (!formUrl) return null;
  return FormApp.openByUrl(formUrl);
}

function findResponseId_(row) {
  const candidateKeys = [
    'ID da resposta',
    'ID da resposta do formulário',
    'Response ID',
    'Response Id',
    'responseId',
    'response id',
  ];

  for (let i = 0; i < candidateKeys.length; i += 1) {
    const value = getFieldValue_(row, candidateKeys[i]);
    if (value) return value;
  }

  return '';
}

function findFormResponseByTimestamp_(form, row) {
  const rawTimestamp = getFieldValue_(row, 'Carimbo de data/hora');
  if (!rawTimestamp) return null;

  const parsedTimestamp = parseTimestamp_(rawTimestamp);
  if (!parsedTimestamp) return null;

  const responses = form.getResponses(parsedTimestamp);
  if (!responses || !responses.length) return null;

  const rowName = getFieldValue_(row, 'Digite seu nome:').toLowerCase();

  for (let i = 0; i < responses.length; i += 1) {
    const response = responses[i];
    const responseTimestamp = response.getTimestamp();
    if (!responseTimestamp || responseTimestamp.getTime() !== parsedTimestamp.getTime()) {
      continue;
    }

    if (!rowName) return response;

    const answers = response.getItemResponses();
    const matchedName = answers.some((itemResponse) => {
      const title = normalizeHeader_(itemResponse.getItem().getTitle());
      if (title !== 'Digite seu nome:') return false;
      return String(itemResponse.getResponse() || '').trim().toLowerCase() === rowName;
    });

    if (matchedName) return response;
  }

  Logger.log('Não foi possível identificar resposta pelo carimbo/nome. Timestamp: %s', rawTimestamp);
  return null;
}

function deleteFromLinkedResponsesSheet_(form, row) {
  try {
    const destinationId = form.getDestinationId();
    if (!destinationId) return false;

    const responsesSpreadsheet = SpreadsheetApp.openById(destinationId);
    const responsesSheet = getResponsesSheet_(responsesSpreadsheet);
    if (!responsesSheet) return false;

    return deleteRowByTimestampAndName_(responsesSheet, row);
  } catch (error) {
    Logger.log('Falha ao excluir na aba de respostas: %s', error);
    return false;
  }
}

function getResponsesSheet_(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i += 1) {
    const sheet = sheets[i];
    const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const normalizedHeaders = headers.map((h) => normalizeHeader_(h));
    if (normalizedHeaders.includes('Carimbo de data/hora')) {
      return sheet;
    }
  }
  return null;
}

function deleteRowByTimestampAndName_(sheet, row) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return false;

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map((h) => normalizeHeader_(h));
  const timestampIndex = headers.indexOf('Carimbo de data/hora');
  if (timestampIndex === -1) return false;

  const nameIndex = headers.indexOf('Digite seu nome:');
  const expectedTimestamp = normalizeDateTimeText_(getFieldValue_(row, 'Carimbo de data/hora'));
  const expectedName = getFieldValue_(row, 'Digite seu nome:').toLowerCase();

  for (let line = 1; line < values.length; line += 1) {
    const candidateTimestamp = normalizeDateTimeText_(values[line][timestampIndex]);
    if (!candidateTimestamp || candidateTimestamp !== expectedTimestamp) continue;

    if (nameIndex !== -1 && expectedName) {
      const candidateName = String(values[line][nameIndex] || '').trim().toLowerCase();
      if (candidateName !== expectedName) continue;
    }

    sheet.deleteRow(line + 1);
    return true;
  }

  return false;
}

function parseTimestamp_(value) {
  const text = String(value || '').trim();
  if (!text) return null;

  const brDateTimeMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (brDateTimeMatch) {
    const [, day, month, year, hour = '0', minute = '0', second = '0'] = brDateTimeMatch;
    return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function normalizeDateTimeText_(value) {
  const parsed = parseTimestamp_(value);
  if (!parsed) return '';

  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  const hour = String(parsed.getHours()).padStart(2, '0');
  const minute = String(parsed.getMinutes()).padStart(2, '0');
  const second = String(parsed.getSeconds()).padStart(2, '0');

  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function getFilterOptions_() {
  const rows = getRows_();
  return {
    sectors: uniqueByKey_(rows, 'Selecione o seu setor:'),
    refs: uniqueByKey_(rows, 'Este registro se refere a:'),
  };
}

function uniqueByKey_(rows, key) {
  const set = new Set();
  rows.forEach((row) => {
    const value = getFieldValue_(row, key);
    if (value) set.add(value);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function normalizeFilters_(filters) {
  const timestampStart = String(filters.timestampStart || filters.timestamp || '').trim().toLowerCase();
  const timestampEnd = String(filters.timestampEnd || filters.timestamp || '').trim().toLowerCase();

  return {
    timestampStart,

    timestampEnd,
    name: String(filters.name || '').trim().toLowerCase(),
    sector: String(filters.sector || '').trim(),
    reference: String(filters.reference || '').trim(),
  };
}

function normalizeHeader_(header) {
  return String(header || '').replace(/\s+/g, ' ').trim();
}

function getFieldValue_(row, key) {
  const normalizedKey = normalizeHeader_(key);
  return String(row[normalizedKey] || '').trim();
}

function normalizeDateFilter_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const brMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (brMatch) {
    const [, day, month, year] = brMatch;
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return `${year}-${month}-${day}`;
  }

  return '';
}
