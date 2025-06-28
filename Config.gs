// Config.gs
// Este arquivo contém configurações globais para a aplicação.

// ID da sua planilha Google. Certifique-se de que este ID está correto!
const SPREADSHEET_ID = '1Ri0_k-8gBiXYjGiNgMBqmyqrZ_5en6ZhMzZeXGRDbX4';

// Nomes das abas da sua planilha para fácil referência.
const SHEETS = {
  TRANSACOES: 'Transacoes',
  CONTAS: 'Contas',
  CATEGORIAS: 'Categorias',
  PESSOAS: 'Pessoas',
  DIVIDAS: 'Dividas',
  INVESTIMENTOS: 'Investimentos',
  APORTES_INVESTIMENTOS: 'AportesInvestimentos', // NOVA ABA
  ORCAMENTO: 'Orcamento',
  METAS: 'Metas',
  DASHBOARD: 'Dashboard'
};

// Função utilitária para obter a planilha por ID
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Função utilitária para obter uma aba específica pelo nome
function getSheet(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    console.error(`Erro: Aba '${sheetName}' não encontrada na planilha ID: ${SPREADSHEET_ID}`);
    throw new Error(`Aba '${sheetName}' não encontrada.`);
  }
  return sheet;
}

// Função utilitária para logar mensagens no console do Apps Script
function log(message) {
  console.log(message);
}