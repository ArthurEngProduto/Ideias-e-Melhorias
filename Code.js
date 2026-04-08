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
        if (normalizedFilters.timestamp && !getFieldValue_(row, 'Carimbo de data/hora').toLowerCase().includes(normalizedFilters.timestamp)) return false;
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
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
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
  return {
    timestamp: String(filters.timestamp || '').trim().toLowerCase(),
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
