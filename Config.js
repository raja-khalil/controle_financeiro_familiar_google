// Config.gs
// Este arquivo contém configurações globais para a aplicação.

/**
 * ID da planilha Google que serve como banco de dados.
 * Você pode encontrar este ID na URL da sua planilha:
 * https://docs.google.com/spreadsheets/d/SEU_ID_DA_PLANILHA_AQUI/edit
 */
const SPREADSHEET_ID = '1Ri0_k-8gBiXYjGiNgMBqmyqrZ_5en6ZhMzZeXGRDbX4';

/**
 * Objeto contendo os nomes das abas (sheets) da sua planilha.
 * Certifique-se de que estes nomes correspondem exatamente aos nomes das abas na sua planilha.
 */
const SHEETS = {
  DASHBOARD: 'Dashboard',
  TRANSACOES: 'Transacoes',
  CATEGORIAS: 'Categorias',
  CONTAS: 'Contas',
  PESSOAS: 'Pessoas',
  DIVIDAS: 'Dividas',
  INVESTIMENTOS: 'Investimentos',
  APORTES_INVESTIMENTOS: 'AportesInvestimentos',
  ORCAMENTO: 'Orcamento',
  METAS: 'Metas'
};

/**
 * Função auxiliar para obter uma aba específica da planilha.
 * @param {string} sheetName O nome da aba a ser obtida.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} O objeto da aba.
 * @throws {Error} Se a aba não for encontrada.
 */
function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Aba "${sheetName}" não encontrada na planilha. Verifique o nome da aba.`);
  }
  return sheet;
}

/**
 * Função auxiliar para logar mensagens no Cloud Logs.
 * @param {string} message A mensagem a ser logada.
 */
function log(message) {
  console.log(message);
}