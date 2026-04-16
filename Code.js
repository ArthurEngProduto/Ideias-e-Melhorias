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
const CONCLUSION_DRIVE_FOLDER_URL = 'https://drive.google.com/drive/folders/1HwupfZk7S3VwRyRhNTUsc3j5UMeY8WAs';
const CONCLUSION_IMAGES_COLUMN_INDEX = 20; // Coluna T
const CONCLUSION_IMAGES_HEADER = 'Imagens de conclusão';
const CONTRIBUTION_TYPE_FIELDS = {
  'Qual é o tipo de contribuição neste produto': [
    'Ideia de melhoria em um produto',
    'Problema identificado',
    'Problema que acontece com frequência',
  ],
  'Qual é o tipo de contribuição neste processo?': [
    'Ideia de melhoria em um processo',
    'Problema identificado',
    'Problema que acontece com frequência',
  ],
};

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
  const filteredRows = getFilteredRows_(filters || {});

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

function getPortalStats(filters) {
  const filteredRows = getFilteredRows_(filters || {});
  const fields = [
    'Selecione o seu setor:',
    'Este registro se refere a:',
    'Qual é o tipo de contribuição neste produto',
    'Qual é o tipo de contribuição neste processo?',
    'Qual é a recorrência ou necessidade desta melhoria?',
  ];

  return {
    total: filteredRows.length,
    charts: fields
      .map((field) => ({
        field,
        values: countByField_(filteredRows, field, {
          allowedValues: CONTRIBUTION_TYPE_FIELDS[field] || null,
        }),
      }))
      .concat({
        field: 'Status da conclusão',
        values: getConclusionStatusCounts_(filteredRows),
      }),
  
  };
}

function getConclusionStatusCounts_(rows) {
  let concluded = 0;
  let inProgress = 0;
  let queued = 0;

  rows.forEach((row) => {
    const statusValue = getRowStatusValue_(row);
    if (statusValue === 'Concluído') {
      concluded += 1;
      return;
    }
    if (statusValue === 'Em andamento') {
      inProgress += 1;
      return;
    }
    queued += 1;
  });

  return [
    { label: 'Concluídos', count: concluded },
    { label: 'Em andamento', count: inProgress },
    { label: 'Na fila', count: queued },
  ];
}

function getRowStatusValue_(row) {
  const rawStatus = String(row['Status'] || row['Concluído'] || '').trim();
  return normalizeRowStatus_(rawStatus);
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

function getFilteredRows_(filters) {
  const rows = getRows_();
  const normalizedFilters = normalizeFilters_(filters || {});

  return rows.filter((row) => {
    const rowDate = normalizeDateFilter_(getFieldValue_(row, 'Carimbo de data/hora'));
    if (normalizedFilters.timestampStart && (!rowDate || rowDate < normalizedFilters.timestampStart)) return false;
    if (normalizedFilters.timestampEnd && (!rowDate || rowDate > normalizedFilters.timestampEnd)) return false;
    if (normalizedFilters.name && !getFieldValue_(row, 'Digite seu nome:').toLowerCase().includes(normalizedFilters.name)) return false;
    const rowSector = getFieldValue_(row, 'Selecione o seu setor:');
    if (normalizedFilters.sectors.length && !normalizedFilters.sectors.includes(rowSector)) return false;
    if (normalizedFilters.sector && rowSector !== normalizedFilters.sector) return false;
    if (normalizedFilters.reference && getFieldValue_(row, 'Este registro se refere a:') !== normalizedFilters.reference) return false;
    if (normalizedFilters.recurrence && getFieldValue_(row, 'Qual é a recorrência ou necessidade desta melhoria?') !== normalizedFilters.recurrence) return false;
    if (normalizedFilters.status && getRowStatusValue_(row) !== normalizedFilters.status) return false;
    return true;
  });
}

function updateRowStatus(rowNumber, status) {
  const numericRow = Number(rowNumber);
  if (!numericRow || numericRow < 2) {
    throw new Error('Linha inválida para atualização.');
  }

  const sheet = getSheet_();
  const statusColumn = getStatusColumnIndex_(sheet);
  const normalizedStatus = normalizeRowStatus_(status);
  sheet.getRange(numericRow, statusColumn).setValue(normalizedStatus);

  return { success: true, rowNumber: numericRow, status: normalizedStatus };
}

function saveConclusionReport(rowNumber, reportText) {
  return saveConclusionReportWithAttachments(rowNumber, reportText, []);
}

function saveConclusionReportWithAttachments(rowNumber, reportText, attachments) {
  const numericRow = Number(rowNumber);
  if (!numericRow || numericRow < 2) {
    throw new Error('Linha inválida para salvar o relatório.');
  }

  const normalizedReport = String(reportText || '').trim();
  if (!normalizedReport) {
    throw new Error('O relatório de conclusão é obrigatório.');
  }

  const sheet = getSheet_();
  if (numericRow > sheet.getLastRow()) {
    throw new Error('A linha informada não existe mais na planilha.');
  }

  const reportColumn = getConclusionReportColumnIndex_(sheet);
  sheet.getRange(numericRow, reportColumn).setValue(normalizedReport);

  const uploadedLinks = uploadConclusionAttachments_(attachments || []);
  if (uploadedLinks.length) {
    const imagesColumn = getConclusionImagesColumnIndex_(sheet);
    const currentValue = String(sheet.getRange(numericRow, imagesColumn).getDisplayValue() || '').trim();
    const mergedLinks = mergeAttachmentLinks_(currentValue, uploadedLinks);
    sheet.getRange(numericRow, imagesColumn).setValue(mergedLinks);
  }

  return { success: true, rowNumber: numericRow, report: normalizedReport, attachmentLinks: uploadedLinks };
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

  const deletionSummary = deleteFormResponseForRow_(rowObject, sheet);
  if (!(deletionSummary.deletedFromForm && deletionSummary.formLinkedToPortalSheet)) {
    sheet.deleteRow(numericRow);
  }

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

function getStatusColumnIndex_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map((header) => normalizeHeader_(header));
  const normalizedCandidates = ['Status', 'Concluído'].map((name) => normalizeHeader_(name));

  for (let i = 0; i < headers.length; i += 1) {
    if (normalizedCandidates.includes(headers[i])) {
      return i + 1;
    }
  }

  return 18;
}

function getConclusionReportColumnIndex_(sheet) {
  const reportHeader = 'Relatório de conclusão';
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map((header) => normalizeHeader_(header));
  const normalizedReportHeader = normalizeHeader_(reportHeader);

  for (let i = 0; i < headers.length; i += 1) {
    if (headers[i] === normalizedReportHeader) {
      return i + 1;
    }
  }

  const newColumn = headers.length + 1;
  sheet.getRange(1, newColumn).setValue(reportHeader);
  return newColumn;
}

function getConclusionImagesColumnIndex_(sheet) {
  const currentHeader = normalizeHeader_(sheet.getRange(1, CONCLUSION_IMAGES_COLUMN_INDEX).getDisplayValue());
  const expectedHeader = normalizeHeader_(CONCLUSION_IMAGES_HEADER);

  if (currentHeader !== expectedHeader) {
    sheet.getRange(1, CONCLUSION_IMAGES_COLUMN_INDEX).setValue(CONCLUSION_IMAGES_HEADER);
  }

  return CONCLUSION_IMAGES_COLUMN_INDEX;
}

function uploadConclusionAttachments_(attachments) {
  if (!Array.isArray(attachments) || !attachments.length) return [];

  const folder = DriveApp.getFolderById(getDriveFolderIdFromUrl_(CONCLUSION_DRIVE_FOLDER_URL));

  return attachments
    .map((file) => normalizeAttachmentPayload_(file))
    .filter((file) => file && file.base64Content)
    .map((file) => {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(file.base64Content),
        file.mimeType || 'application/octet-stream',
        file.fileName || `anexo-${new Date().getTime()}`
      );
      const createdFile = folder.createFile(blob);
      return createdFile.getUrl();
    });
}

function normalizeAttachmentPayload_(file) {
  if (!file || typeof file !== 'object') return null;
  const fileName = String(file.fileName || file.name || '').trim();
  const mimeType = String(file.mimeType || file.type || '').trim();
  const base64Content = String(file.base64Content || file.base64 || '').trim();
  if (!base64Content) return null;

  return {
    fileName,
    mimeType,
    base64Content,
  };
}

function getDriveFolderIdFromUrl_(folderUrl) {
  const url = String(folderUrl || '').trim();
  const match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!match || !match[1]) {
    throw new Error('Não foi possível identificar o ID da pasta de destino no Google Drive.');
  }
  return match[1];
}

function mergeAttachmentLinks_(existingValue, newLinks) {
  const existingLinks = String(existingValue || '')
    .split('@')
    .map((item) => String(item || '').trim())
    .filter((item) => item);
  const additionalLinks = (newLinks || [])
    .map((item) => String(item || '').trim())
    .filter((item) => item);

  return existingLinks.concat(additionalLinks).join('@');
}
function normalizeRowStatus_(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return 'Na fila';
  if (normalized === 'concluído' || normalized === 'concluido' || normalized === 'ok') return 'Concluído';
    if (normalized === 'em andamento' || normalized === 'andamento' || normalized === 'in progress') return 'Em andamento';
  if (normalized === 'na fila' || normalized === 'fila' || normalized === 'parado' || normalized === 'pausado' || normalized === 'stopped') return 'Na fila';
  return 'Em andamento';
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

function deleteFormResponseForRow_(row, portalSheet) {
  const form = getLinkedForm_();
  if (!form) {
    return {
      deletedFromForm: false,
      deletedFromResponsesSheet: false,
      formLinkedToPortalSheet: false,
    };
  }

  let deletedFromForm = false;
  const formLinkedToPortalSheet = isFormLinkedToPortalSheet_(form, portalSheet);

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

  const deletedFromResponsesSheet = deleteFromLinkedResponsesSheet_(form, row, portalSheet);

  return {
    deletedFromForm,
    deletedFromResponsesSheet,
    formLinkedToPortalSheet,
  };
}

function isFormLinkedToPortalSheet_(form, portalSheet) {
  if (!portalSheet) return false;

  try {
    const destinationId = form.getDestinationId();
    if (!destinationId) return false;

    const responsesSpreadsheet = SpreadsheetApp.openById(destinationId);
    const responsesSheet = getResponsesSheet_(responsesSpreadsheet);
    if (!responsesSheet) return false;

    return (
      responsesSpreadsheet.getId() === portalSheet.getParent().getId() &&
      responsesSheet.getSheetId() === portalSheet.getSheetId()
    );
  } catch (error) {
    Logger.log('Falha ao validar vínculo da aba de respostas: %s', error);
    return false;
  }
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

function deleteFromLinkedResponsesSheet_(form, row, portalSheet) {
  try {
    const destinationId = form.getDestinationId();
    if (!destinationId) return false;

    const responsesSpreadsheet = SpreadsheetApp.openById(destinationId);
    const responsesSheet = getResponsesSheet_(responsesSpreadsheet);
    if (!responsesSheet) return false;
    if (
      portalSheet &&
      responsesSpreadsheet.getId() === portalSheet.getParent().getId() &&
      responsesSheet.getSheetId() === portalSheet.getSheetId()
    ) {
      return false;
    }

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
    recurrences: uniqueByKey_(rows, 'Qual é a recorrência ou necessidade desta melhoria?'),
    statuses: uniqueByRowStatus_(rows),
  };
}

function uniqueByRowStatus_(rows) {
  const defaultStatuses = ['Na fila', 'Em andamento', 'Concluído'];
  const set = new Set(defaultStatuses);
  rows.forEach((row) => {
    set.add(getRowStatusValue_(row));
  });
  
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function uniqueByKey_(rows, key) {
  const set = new Set();
  rows.forEach((row) => {
    const value = getFieldValue_(row, key);
    if (value) set.add(value);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function countByField_(rows, key, options) {
  const safeOptions = options || {};
  const allowedValues = Array.isArray(safeOptions.allowedValues) ? safeOptions.allowedValues : null;
  const allowedValueByNormalizedKey = allowedValues
    ? allowedValues.reduce((acc, value) => {
        acc[normalizeOptionValue_(value)] = value;
        return acc;
      }, {})
    : null;
  const counts = {};
    if (allowedValues) {
    allowedValues.forEach((value) => {
      counts[value] = 0;
    });
  }

  rows.forEach((row) => {
    const rawValue = getFieldValue_(row, key);

    if (allowedValueByNormalizedKey) {
      const normalizedValue = normalizeOptionValue_(rawValue);
      const canonicalValue = allowedValueByNormalizedKey[normalizedValue];
      if (!canonicalValue) return;
      counts[canonicalValue] = (counts[canonicalValue] || 0) + 1;
      return;
    }

    const value = rawValue || 'Não informado';
    counts[value] = (counts[value] || 0) + 1;
  });

  return Object.keys(counts)
    .sort((a, b) => counts[b] - counts[a] || a.localeCompare(b, 'pt-BR'))
    .map((label) => ({ label, count: counts[label] }));
}


function normalizeOptionValue_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeFilters_(filters) {
  const timestampStart = String(filters.timestampStart || filters.timestamp || '').trim().toLowerCase();
  const timestampEnd = String(filters.timestampEnd || filters.timestamp || '').trim().toLowerCase();

  const normalizedSectors = Array.isArray(filters.sectors)
    ? filters.sectors
        .map((value) => String(value || '').trim())
        .filter((value, index, array) => value && array.indexOf(value) === index)
    : [];

  return {
    timestampStart,

    timestampEnd,
    name: String(filters.name || '').trim().toLowerCase(),
    sectors: normalizedSectors,
    sector: String(filters.sector || '').trim(),
    reference: String(filters.reference || '').trim(),
    recurrence: String(filters.recurrence || '').trim(),
    status: String(filters.status || '').trim() ? normalizeRowStatus_(filters.status) : '',
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

