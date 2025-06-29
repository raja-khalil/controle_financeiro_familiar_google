// Web_App_Functions.gs
// Este arquivo contém as funções de backend chamadas pela Web App.
// Inclui todas as lógicas de interação com o Google Sheets, cálculos e notificações.

// Certifique-se de que o arquivo Config.gs esteja no mesmo projeto e contenha
// as constantes SPREADSHEET_ID e SHEETS corretamente definidas.

/**
 * Função principal para servir a Web App.
 * Executada quando a URL da Web App é acessada via GET.
 * Renderiza o arquivo HTML correspondente ao parâmetro 'page' da URL.
 * @param {GoogleAppsScript.Events.AppsScriptHttpRequestEvent} e Objeto de evento HTTP.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Conteúdo HTML da página solicitada.
 */
function doGet(e) {
  const page = e.parameter.page;
  let template;

  if (page) {
    switch (page) {
      case 'dashboard':
        template = HtmlService.createTemplateFromFile('Dashboard');
        break;
      case 'transacoes':
        template = HtmlService.createTemplateFromFile('Transacoes');
        break;
      case 'orcamento':
        template = HtmlService.createTemplateFromFile('Orcamento');
        break;
      case 'metas':
        template = HtmlService.createTemplateFromFile('Metas');
        break;
      case 'dividas':
        template = HtmlService.createTemplateFromFile('Dividas');
        break;
      case 'investimentos':
        template = HtmlService.createTemplateFromFile('Investimentos');
        break;
      case 'analises':
        template = HtmlService.createTemplateFromFile('Analises');
        break;
      case 'gerenciarCategorias':
        template = HtmlService.createTemplateFromFile('GerenciarCategorias');
        break;
      case 'gerenciarContas':
        template = HtmlService.createTemplateFromFile('GerenciarContas');
        break;
      case 'gerenciarPessoas':
        template = HtmlService.createTemplateFromFile('GerenciarPessoas');
        break;
      default:
        template = HtmlService.createTemplateFromFile('Index'); // Página padrão se inválido
    }
  } else {
    template = HtmlService.createTemplateFromFile('Index'); // Página inicial
  }

  return template
      .evaluate()
      .setTitle('Controle Financeiro Familiar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // Modo de segurança recomendado
}

/**
 * Função auxiliar para incluir outros arquivos HTML (CSS, JS) dentro de templates HTML.
 * (Não é mais usada para carregar CSS/JS embutido, mas pode ser útil para modularizar HTML complexo).
 * @param {string} filename O nome do arquivo HTML a ser incluído (sem extensão .html).
 * @returns {string} O conteúdo do arquivo HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtém todos os dados de uma aba específica da planilha.
 * A primeira linha é considerada o cabeçalho.
 * @param {string} sheetName Nome da aba (Ex: 'Transacoes', 'Contas').
 * @returns {Array<Array<any>>} Array de arrays com os dados da aba, incluindo cabeçalho.
 */
function getSheetData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    const range = sheet.getDataRange();
    const data = range.getValues();
    
    if (data.length === 0) {
      log(`getSheetData: Aba '${sheetName}' está vazia.`);
      return [];
    }

    const headers = data[0];
    log(`getSheetData: Aba '${sheetName}' Headers: [${headers.join(', ')}]`);
    log(`getSheetData: Aba '${sheetName}' Total de linhas (incluindo header): ${data.length}`);
    
    return data;
  } catch (e) {
    log(`getSheetData: Erro ao obter dados da aba '${sheetName}': ${e.message}`);
    return [];
  }
}

/**
 * Obtém dados de múltiplas abas da planilha em uma única chamada.
 * Isso é mais eficiente do que fazer várias chamadas separadas do frontend.
 * @param {Object} sheetsToFetch Objeto onde as chaves são nomes de referência (e.g., 'categorias', 'contas')
 * e os valores são os nomes reais das abas na planilha (e.g., 'Categorias', 'Contas').
 * @returns {Object} Um objeto com os dados de cada aba, usando as chaves de referência.
 */
function getSheetDataBatch(sheetsToFetch) {
  const result = {};
  log("getSheetDataBatch: Iniciando carregamento de dados em lote.");
  for (const key in sheetsToFetch) {
    if (Object.prototype.hasOwnProperty.call(sheetsToFetch, key)) {
      result[key] = getSheetData(sheetsToFetch[key]);
      log(`getSheetDataBatch: Dados para '${key}' (${sheetsToFetch[key]}) processados. Retornou ${result[key].length} linhas.`);
    }
  }
  log("getSheetDataBatch: Carregamento de dados em lote concluído.");
  return result;
}

/**
 * Atualiza o saldo de uma conta específica na aba 'Contas'.
 * Procura a conta pelo nome e soma o valor ao saldo atual.
 * @param {string} accountName O nome da conta a ser atualizada.
 * @param {number} amount O valor a ser adicionado (positivo para entrada, negativo para saída).
 * @returns {boolean} true se o saldo for atualizado, false caso contrário.
 */
function updateAccountBalance(accountName, amount) {
  try {
    const contasSheet = getSheet(SHEETS.CONTAS);
    const data = contasSheet.getDataRange().getValues();

    const headerRow = data[0];
    const accountNameColIndex = headerRow.indexOf('Nome da Conta');
    const saldoAtualColIndex = headerRow.indexOf('Saldo Atual');

    if (accountNameColIndex === -1 || saldoAtualColIndex === -1) {
      throw new Error('Colunas "Nome da Conta" ou "Saldo Atual" não encontradas na aba Contas.');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][accountNameColIndex] === accountName) {
        let currentBalance = parseFloat(data[i][saldoAtualColIndex] || 0);
        const newBalance = currentBalance + amount;
        contasSheet.getRange(i + 1, saldoAtualColIndex + 1).setValue(newBalance);
        log(`updateAccountBalance: Conta '${accountName}' atualizada para R$ ${newBalance.toFixed(2)}.`);
        return true;
      }
    }
    log(`updateAccountBalance: Erro: Conta '${accountName}' não encontrada para atualização de saldo.`);
    return false;
  } catch (e) {
    log(`updateAccountBalance: Erro ao atualizar saldo da conta: ${e.message}`);
    return false;
  }
}

// --- Funções para Transações ---

/**
 * Salva uma nova transação na aba 'Transacoes' e atualiza o saldo da conta.
 * @param {Object} transaction Objeto com os dados da transação.
 * @returns {boolean} true se a transação for salva, false caso contrário.
 */
function saveTransaction(transaction) {
  try {
    const transacoesSheet = getSheet(SHEETS.TRANSACOES);
    const headers = transacoesSheet.getDataRange().getValues()[0]; // Obter cabeçalhos atualizados

    if (!transaction.data || !transaction.tipo || !transaction.valor || !transaction.conta || !transaction.descricao || !transaction.categoria || !transaction.pessoa || !transaction.tipoPagamento) {
      throw new Error('Dados da transação incompletos. Verifique Data, Tipo, Valor, Descrição, Categoria, Conta, Tipo de Pagamento e Pessoa.');
    }
    const valorNumerico = parseFloat(transaction.valor);
    if (isNaN(valorNumerico) || valorNumerico <= 0) {
      throw new Error('Valor da transação inválido. Deve ser um número positivo.');
    }

    const nextId = `TR${transacoesSheet.getLastRow() + 1}`; 
    const valorParaContas = transaction.tipo === 'Saída' ? -valorNumerico : valorNumerico;

    // Mapear os dados do objeto transaction para a ordem das colunas na planilha
    const rowData = new Array(headers.length).fill(''); // Cria uma linha vazia com o tamanho dos cabeçalhos
    
    rowData[headers.indexOf('ID')] = nextId;
    rowData[headers.indexOf('Data')] = transaction.data;
    rowData[headers.indexOf('Tipo')] = transaction.tipo;
    rowData[headers.indexOf('Valor (R$)')] = valorNumerico;
    rowData[headers.indexOf('Descricao')] = transaction.descricao;
    rowData[headers.indexOf('Categoria')] = transaction.categoria;
    rowData[headers.indexOf('Conta')] = transaction.conta;
    rowData[headers.indexOf('Pessoa')] = transaction.pessoa;
    rowData[headers.indexOf('Observacoes')] = transaction.observacoes || '';
    
    const tipoPagamentoColIndex = headers.indexOf('Tipo de Pagamento');
    if (tipoPagamentoColIndex !== -1) {
        rowData[tipoPagamentoColIndex] = transaction.tipoPagamento;
    } else {
        log("saveTransaction: Aviso: Coluna 'Tipo de Pagamento' não encontrada. Verifique se executou a função addPaymentTypeColumn().");
    }

    transacoesSheet.appendRow(rowData);
    log(`saveTransaction: Transação '${transaction.descricao}' (${transaction.tipo}) salva.`);

    updateAccountBalance(transaction.conta, valorParaContas);
    log(`saveTransaction: Saldo da conta '${transaction.conta}' atualizado.`);

    return true;
  } catch (e) {
    log(`saveTransaction: Erro ao salvar transação: ${e.message}`);
    return false;
  }
}

/**
 * Atualiza um registro existente na aba 'Transacoes'.
 * @param {Object} transactionData Objeto com os dados da transação a ser atualizada. Deve incluir o ID.
 * @returns {boolean} true se a transação for atualizada, false caso contrário.
 */
function updateTransaction(transactionData) {
  try {
    const sheet = getSheet(SHEETS.TRANSACOES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const idColIndex = headers.indexOf('ID');
    const dataColIndex = headers.indexOf('Data');
    const tipoColIndex = headers.indexOf('Tipo');
    const valorColIndex = headers.indexOf('Valor (R$)');
    const descricaoColIndex = headers.indexOf('Descricao');
    const categoriaColIndex = headers.indexOf('Categoria');
    const contaColIndex = headers.indexOf('Conta');
    const pessoaColIndex = headers.indexOf('Pessoa');
    const observacoesColIndex = headers.indexOf('Observacoes');
    const tipoPagamentoColIndex = headers.indexOf('Tipo de Pagamento'); 

    if (idColIndex === -1 || dataColIndex === -1 || tipoColIndex === -1 || valorColIndex === -1 ||
        descricaoColIndex === -1 || categoriaColIndex === -1 || contaColIndex === -1 ||
        pessoaColIndex === -1 || observacoesColIndex === -1) {
      throw new Error('Colunas necessárias não encontradas na aba Transacoes para atualização.');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === transactionData.id) {
        const oldTipo = data[i][tipoColIndex];
        const oldValor = parseFloat(data[i][valorColIndex] || 0);
        const oldConta = data[i][contaColIndex];

        const newValor = parseFloat(transactionData.valor);
        const newTipo = transactionData.tipo;
        const newConta = transactionData.conta;

        // Estornar o valor antigo da conta antiga (se mudou)
        if (oldConta && (oldConta !== newConta || oldTipo !== newTipo || oldValor !== newValor)) {
          const valorParaEstornar = oldTipo === 'Saída' ? oldValor : -oldValor;
          updateAccountBalance(oldConta, valorParaEstornar);
        }

        const rowToUpdate = data[i]; 

        rowToUpdate[dataColIndex] = transactionData.data;
        rowToUpdate[tipoColIndex] = newTipo;
        rowToUpdate[valorColIndex] = newValor;
        rowToUpdate[descricaoColIndex] = transactionData.descricao;
        rowToUpdate[categoriaColIndex] = transactionData.categoria;
        rowToUpdate[contaColIndex] = newConta;
        rowToUpdate[pessoaColIndex] = transactionData.pessoa;
        rowToUpdate[observacoesColIndex] = transactionData.observacoes || '';
        
        if (tipoPagamentoColIndex !== -1) {
            rowToUpdate[tipoPagamentoColIndex] = transactionData.tipoPagamento || '';
        }

        sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);

        // Aplica o novo valor na nova conta (se mudou)
        const valorParaAplicar = newTipo === 'Saída' ? -newValor : newValor;
        updateAccountBalance(newConta, valorParaAplicar);

        log(`updateTransaction: Transação '${transactionData.id}' atualizada.`);
        return true;
      }
    }
    log(`updateTransaction: Erro: Transação com ID '${transactionData.id}' não encontrada para atualização.`);
    return false;
  } catch (e) {
    log(`updateTransaction: Erro ao atualizar transação: ${e.message}`);
    return false;
  }
}

/**
 * Exclui uma transação da aba 'Transacoes' e ajusta o saldo da conta.
 * @param {string} transactionId ID da transação a ser excluída.
 * @returns {boolean} true se a transação for excluída, false caso contrário.
 */
function deleteTransaction(transactionId) {
  try {
    const sheet = getSheet(SHEETS.TRANSACOES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const idColIndex = headers.indexOf('ID');
    const tipoColIndex = headers.indexOf('Tipo');
    const valorColIndex = headers.indexOf('Valor (R$)');
    const contaColIndex = headers.indexOf('Conta');

    if (idColIndex === -1 || tipoColIndex === -1 || valorColIndex === -1 || contaColIndex === -1) {
      throw new Error('Colunas necessárias não encontradas na aba Transacoes para exclusão.');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === transactionId) {
        const tipo = data[i][tipoColIndex];
        const valor = parseFloat(data[i][valorColIndex] || 0);
        const conta = data[i][contaColIndex];

        // Estorna o valor da conta (o oposto do que foi registrado)
        const valorParaEstornar = tipo === 'Saída' ? valor : -valor;
        updateAccountBalance(conta, valorParaEstornar);

        sheet.deleteRow(i + 1);
        log(`deleteTransaction: Transação '${transactionId}' excluída.`);
        return true;
      }
    }
    log(`deleteTransaction: Erro: Transação com ID '${transactionId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`deleteTransaction: Erro ao excluir transação: ${e.message}`);
    return false;
  }
}

/**
 * Registra uma movimentação (transferência) entre duas contas.
 * Decrementa o saldo da conta de origem e incrementa o da conta de destino.
 * Registra duas transações (Saída e Entrada) na aba 'Transacoes'.
 * @param {Object} transferData Objeto com os dados da transferência.
 * @returns {boolean} true se a transferência for registrada, false caso contrário.
 */
function recordTransfer(transferData) {
  try {
    if (!transferData.data || !transferData.fromAccount || !transferData.toAccount || !transferData.value || !transferData.person) {
      throw new Error('Dados da transferência incompletos. Verifique Data, Contas, Valor e Pessoa.');
    }
    const valueNumeric = parseFloat(transferData.value);
    if (isNaN(valueNumeric) || valueNumeric <= 0) {
      throw new Error('Valor da transferência inválido. Deve ser um número positivo.');
    }
    if (transferData.fromAccount === transferData.toAccount) {
        throw new Error('Conta de origem e conta de destino não podem ser as mesmas para uma transferência.');
    }

    // 1. Debitar da conta de origem
    const debitSuccess = updateAccountBalance(transferData.fromAccount, -valueNumeric);
    if (!debitSuccess) {
        throw new Error(`recordTransfer: Falha ao debitar da conta de origem: ${transferData.fromAccount}.`);
    }

    // 2. Creditar na conta de destino
    const creditSuccess = updateAccountBalance(transferData.toAccount, valueNumeric);
    if (!creditSuccess) {
        // Se o crédito falhar, tentar reverter o débito (opcional, mas boa prática)
        updateAccountBalance(transferData.fromAccount, valueNumeric); 
        throw new Error(`recordTransfer: Falha ao creditar na conta de destino: ${transferData.toAccount}.`);
    }

    // 3. Registrar transação de SAÍDA na aba 'Transacoes'
    saveTransaction({
      data: transferData.data,
      tipo: 'Saída',
      valor: valueNumeric,
      descricao: `Transferência para: ${transferData.toAccount}`,
      categoria: 'Transferência', 
      conta: transferData.fromAccount,
      tipoPagamento: 'Transferência entre Contas', 
      pessoa: transferData.person,
      observacoes: transferData.observations || `Transferência de ${transferData.fromAccount} para ${transferData.toAccount}`
    });

    // 4. Registrar transação de ENTRADA na aba 'Transacoes'
    saveTransaction({
      data: transferData.data,
      tipo: 'Entrada',
      valor: valueNumeric,
      descricao: `Transferência de: ${transferData.fromAccount}`,
      categoria: 'Transferência', 
      conta: transferData.toAccount,
      pessoa: transferData.person,
      tipoPagamento: 'Transferência entre Contas', 
      observacoes: transferData.observations || `Transferência de ${transferData.fromAccount} para ${transferData.toAccount}`
    });

    log(`recordTransfer: Transferência de R$ ${valueNumeric.toFixed(2)} de '${transferData.fromAccount}' para '${transferData.toAccount}' registrada.`);
    return true;
  } catch (e) {
    log(`recordTransfer: Erro ao registrar transferência: ${e.message}`);
    return false;
  }
}


// --- Funções para Orçamento ---

/**
 * Salva ou atualiza um item de orçamento na aba 'Orcamento'.
 * @param {Object} budgetData Objeto com os dados do orçamento: { anoMes, categoria, produtoServico, tipo, valorOrcado }
 * @returns {boolean} true se o orçamento for salvo/atualizado, false caso contrário.
 */
function saveBudget(budgetData) {
  try {
    const sheet = getSheet(SHEETS.ORCAMENTO);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const anoMesColIndex = headers.indexOf('AnoMes');
    const categoriaColIndex = headers.indexOf('Categoria');
    const produtoServicoColIndex = headers.indexOf('Produto/Serviço');
    const tipoColIndex = headers.indexOf('Tipo');
    const valorOrcadoColIndex = headers.indexOf('Valor Orcado');

    if (anoMesColIndex === -1 || categoriaColIndex === -1 || produtoServicoColIndex === -1 || tipoColIndex === -1 || valorOrcadoColIndex === -1) {
      throw new Error('Colunas de orçamento (AnoMes, Categoria, Produto/Serviço, Tipo, Valor Orcado) não encontradas. Verifique os cabeçalhos da aba Orcamento.');
    }

    if (!budgetData.anoMes || !budgetData.categoria || !budgetData.produtoServico || !budgetData.tipo || isNaN(parseFloat(budgetData.valorOrcado)) || parseFloat(budgetData.valorOrcado) < 0) {
        throw new Error('Dados do orçamento incompletos ou inválidos.');
    }

    let found = false;
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (row[anoMesColIndex] === budgetData.anoMes && 
          row[categoriaColIndex] === budgetData.categoria &&
          row[produtoServicoColIndex] === budgetData.produtoServico &&
          row[tipoColIndex] === budgetData.tipo) {
        
        sheet.getRange(i + 1, valorOrcadoColIndex + 1).setValue(parseFloat(budgetData.valorOrcado));
        found = true;
        log(`saveBudget: Orçamento para '${budgetData.produtoServico}' em '${budgetData.categoria}' (${budgetData.tipo}, ${budgetData.anoMes}) atualizado.`);
        break;
      }
    }
    if (!found) {
      sheet.appendRow([budgetData.anoMes, budgetData.categoria, budgetData.produtoServico, budgetData.tipo, parseFloat(budgetData.valorOrcado)]);
      log(`saveBudget: Novo orçamento para '${budgetData.produtoServico}' em '${budgetData.categoria}' (${budgetData.tipo}, ${budgetData.anoMes}) salvo.`);
    }
    return true;
  } catch (e) {
    log(`saveBudget: Erro ao salvar orçamento: ${e.message}`);
    return false;
  }
}

/**
 * Retorna dados para análise de orçamento (gastos reais vs. valor orçado) para um dado mês/ano.
 * Agrega os gastos das transações e compara com o orçamento definido, somando por categoria.
 * Inclui "Receita Estimada" no relatório.
 * @param {string} anoMes Ex: '2025-06'
 * @returns {Object} Objeto com { receitas: Array, despesas: Array }.
 */
function getBudgetAnalysis(anoMes) {
  try {
    const transacoesSheet = getSheet(SHEETS.TRANSACOES);
    const orcamentoSheet = getSheet(SHEETS.ORCAMENTO);

    const transacoes = transacoesSheet.getDataRange().getValues();
    const orcamentos = orcamentoSheet.getDataRange().getValues();

    const gastosPorCategoria = {};
    const receitaRealPorCategoria = {};
    const orcamentosPorCategoria = { receitas: {}, despesas: {} };

    // Processa orçamento (pulando cabeçalho)
    const orcamentoHeaders = orcamentos[0];
    const orcamentoAnoMesCol = orcamentoHeaders.indexOf('AnoMes');
    const orcamentoCategoriaCol = orcamentoHeaders.indexOf('Categoria');
    const orcamentoTipoCol = orcamentoHeaders.indexOf('Tipo');
    const orcamentoValorCol = orcamentoHeaders.indexOf('Valor Orcado');

    if (orcamentoAnoMesCol === -1 || orcamentoCategoriaCol === -1 || orcamentoTipoCol === -1 || orcamentoValorCol === -1) {
      throw new Error('Colunas da aba Orcamento não encontradas (AnoMes, Categoria, Tipo, Valor Orcado).');
    }

    for (let i = 1; i < orcamentos.length; i++) {
      const row = orcamentos[i];
      const sheetAnoMes = row[orcamentoAnoMesCol];
      const categoria = row[orcamentoCategoriaCol];
      const tipoOrcamento = row[orcamentoTipoCol];
      const valorOrcado = row[orcamentoValorCol];

      if (sheetAnoMes === anoMes && categoria && tipoOrcamento) {
        if (tipoOrcamento === 'Receita') {
          orcamentosPorCategoria.receitas[categoria] = (orcamentosPorCategoria.receitas[categoria] || 0) + parseFloat(valorOrcado || 0);
        } else if (tipoOrcamento === 'Despesa') {
          orcamentosPorCategoria.despesas[categoria] = (orcamentosPorCategoria.despesas[categoria] || 0) + parseFloat(valorOrcado || 0);
        }
      }
    }

    // Processa transações (pulando cabeçalho)
    const transacoesHeaders = transacoes[0];
    const transacaoDataCol = transacoesHeaders.indexOf('Data');
    const transacaoTipoCol = transacoesHeaders.indexOf('Tipo');
    const transacaoValorCol = transacoesHeaders.indexOf('Valor (R$)');
    const transacaoCategoriaCol = transacoesHeaders.indexOf('Categoria');

    if (transacaoDataCol === -1 || transacaoTipoCol === -1 || transacaoValorCol === -1 || transacaoCategoriaCol === -1) {
        throw new Error('Colunas da aba Transacoes não encontradas para análise de orçamento.');
    }

    for (let i = 1; i < transacoes.length; i++) {
      const row = transacoes[i];
      const transacaoDataStr = Utilities.formatDate(new Date(row[transacaoDataCol]), Session.getScriptTimeZone(), 'yyyy-MM');
      const tipoTransacao = row[transacaoTipoCol];
      const valor = parseFloat(row[transacaoValorCol] || 0);
      const categoria = row[transacaoCategoriaCol];
      
      if (transacaoDataStr === anoMes && categoria) {
        if (tipoTransacao === 'Saída') {
          gastosPorCategoria[categoria] = (gastosPorCategoria[categoria] || 0) + Math.abs(valor);
        } else if (tipoTransacao === 'Entrada') {
          receitaRealPorCategoria[categoria] = (receitaRealPorCategoria[categoria] || 0) + Math.abs(valor);
        }
      }
    }

    const despesasResults = [];
    for (const categoria in orcamentosPorCategoria.despesas) {
      despesasResults.push({
        categoria: categoria,
        orcado: orcamentosPorCategoria.despesas[categoria],
        gasto: gastosPorCategoria[categoria] || 0
      });
    }
    for (const categoria in gastosPorCategoria) {
        if (!orcamentosPorCategoria.despesas.hasOwnProperty(categoria) && gastosPorCategoria.hasOwnProperty(categoria)) { 
            despesasResults.push({
                categoria: categoria,
                orcado: 0,
                gasto: gastosPorCategoria[categoria]
            });
        }
    }

    const receitasResults = [];
    for (const categoria in orcamentosPorCategoria.receitas) {
      receitasResults.push({
        categoria: categoria,
        estimado: orcamentosPorCategoria.receitas[categoria],
        realizado: receitaRealPorCategoria[categoria] || 0
      });
    }
    for (const categoria in receitaRealPorCategoria) {
        if (!orcamentosPorCategoria.receitas.hasOwnProperty(categoria) && receitaRealPorCategoria.hasOwnProperty(categoria)) { 
            receitasResults.push({
                categoria: categoria,
                estimado: 0,
                realizado: receitaRealPorCategoria[categoria]
            });
        }
    }

    return { receitas: receitasResults, despesas: despesasResults };
  } catch (e) {
    log(`getBudgetAnalysis: Erro ao obter análise de orçamento: ${e.message}`);
    return { receitas: [], despesas: [] };
  }
}

// --- Funções para Metas ---

/**
 * Salva ou atualiza uma meta financeira na aba 'Metas'.
 * @param {Object} goalData Objeto com os dados da meta.
 * @returns {boolean} true se a meta for salva/atualizada, false caso contrário.
 */
function saveGoal(goalData) {
  try {
    const sheet = getSheet(SHEETS.METAS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const colIndices = {
        id: headers.indexOf('ID'),
        nome: headers.indexOf('Nome da Meta'),
        tipo: headers.indexOf('Tipo'),
        valorAlvo: headers.indexOf('Valor Alvo'),
        valorContribuido: headers.indexOf('Valor Contribuido'),
        dataInicio: headers.indexOf('Data Inicio'),
        dataAlvo: headers.indexOf('Data Alvo'),
        status: headers.indexOf('Status'),
        prioridade: headers.indexOf('Prioridade'),
        observacoes: headers.indexOf('Observacoes')
    };

    if (Object.values(colIndices).some(idx => idx === -1)) {
        throw new Error(`Colunas da aba Metas não encontradas. Verifique os cabeçalhos.`);
    }

    if (!goalData.nome || !goalData.tipo || isNaN(parseFloat(goalData.valorAlvo)) || parseFloat(goalData.valorAlvo) < 0 || !goalData.dataInicio || !goalData.status) {
        throw new Error('Dados da meta incompletos ou inválidos (Nome, Tipo, Valor Alvo, Data Início, Status são obrigatórios).');
    }

    if (goalData.id && goalData.id.startsWith('M')) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][colIndices.id] === goalData.id) {
          const rowToUpdate = new Array(headers.length);
          rowToUpdate[colIndices.id] = goalData.id;
          rowToUpdate[colIndices.nome] = goalData.nome;
          rowToUpdate[colIndices.tipo] = goalData.tipo;
          rowToUpdate[colIndices.valorAlvo] = parseFloat(goalData.valorAlvo);
          rowToUpdate[colIndices.valorContribuido] = parseFloat(goalData.valorContribuido || 0);
          rowToUpdate[colIndices.dataInicio] = goalData.dataInicio;
          rowToUpdate[colIndices.dataAlvo] = goalData.dataAlvo;
          rowToUpdate[colIndices.status] = goalData.status;
          rowToUpdate[colIndices.prioridade] = goalData.prioridade || '';
          rowToUpdate[colIndices.observacoes] = goalData.observacoes || '';

          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`saveGoal: Meta '${goalData.nome}' (ID: ${goalData.id}) atualizada.`);
          return true;
        }
      }
    }

    const nextId = `M${sheet.getLastRow() + 1}`;
    sheet.appendRow([
      nextId,
      goalData.nome,
      goalData.tipo,
      parseFloat(goalData.valorAlvo),
      parseFloat(goalData.valorContribuido || 0),
      goalData.dataInicio,
      goalData.dataAlvo,
      goalData.status,
      goalData.prioridade || '',
      goalData.observacoes || ''
    ]);
    log(`saveGoal: Nova meta '${goalData.nome}' (ID: ${nextId}) salva.`);
    return true;
  }
   catch (e) {
    log(`saveGoal: Erro ao salvar meta: ${e.message}`);
    return false;
  }
}

/**
 * Adiciona uma contribuição (valor monetário) a uma meta existente.
 * @param {string} goalId ID da meta a ser atualizada.
 * @param {number} amount Valor da contribuição a ser adicionado.
 * @returns {boolean} true se a contribuição for adicionada com sucesso, false caso contrário.
 */
function contributeToGoal(goalId, amount) {
  try {
    const sheet = getSheet(SHEETS.METAS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const idColIndex = headers.indexOf('ID');
    const valorContribuidoColIndex = headers.indexOf('Valor Contribuido');
    const valorAlvoColIndex = headers.indexOf('Valor Alvo');
    const statusColIndex = headers.indexOf('Status');
    const nomeMetaColIndex = headers.indexOf('Nome da Meta');

    if (idColIndex === -1 || valorContribuidoColIndex === -1 || valorAlvoColIndex === -1 || statusColIndex === -1 || nomeMetaColIndex === -1) {
        throw new Error(`Colunas da aba Metas não encontradas para adicionar contribuição.`);
    }
    if (isNaN(amount) || amount <= 0) {
        throw new Error('Valor de contribuição inválido. Deve ser um número positivo.');
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idColIndex] === goalId) {
        let currentContributed = parseFloat(allData[i][valorContribuidoColIndex] || 0);
        let valorAlvo = parseFloat(allData[i][valorAlvoColIndex] || 0);
        let newContributed = currentContributed + amount;

        sheet.getRange(i + 1, valorContribuidoColIndex + 1).setValue(newContributed);

        if (newContributed >= valorAlvo && allData[i][statusColIndex] !== 'Alcancada') {
          sheet.getRange(i + 1, statusColIndex + 1).setValue('Alcancada');
          sendGoalReachedEmail(allData[i][nomeMetaColIndex]);
          log(`contributeToGoal: Meta '${allData[i][nomeMetaColIndex]}' (ID: ${goalId}) alcançada.`);
        }
        log(`contributeToGoal: Contribuição de R$ ${amount.toFixed(2)} adicionada à meta '${allData[i][nomeMetaColIndex]}'.`);
        return true;
      }
    }
    log(`contributeToGoal: Erro: Meta com ID '${goalId}' não encontrada para adicionar contribuição.`);
    return false;
  } catch (e) {
    log(`contributeToGoal: Erro ao contribuir para meta: ${e.message}`);
    return false;
  }
}

/**
 * Atualiza um registro existente na aba 'Metas'.
 * @param {Object} goalData Objeto com os dados da meta a ser atualizada. Deve incluir o ID.
 * @returns {boolean} true se a meta for atualizada, false caso contrário.
 */
function updateGoal(goalData) {
    // Reutilizamos saveGoal, que já tem a lógica de atualização se o ID for fornecido
    return saveGoal(goalData);
}

/**
 * Exclui uma meta da aba 'Metas'.
 * @param {string} goalId ID da meta a ser excluída.
 * @returns {boolean} true se a meta for excluída, false caso contrário.
 */
function deleteGoal(goalId) {
  try {
    const sheet = getSheet(SHEETS.METAS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID');

    if (idColIndex === -1) throw new Error('Coluna ID não encontrada na aba Metas.');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === goalId) {
        sheet.deleteRow(i + 1);
        log(`deleteGoal: Meta '${goalId}' excluída.`);
        return true;
      }
    }
    log(`deleteGoal: Erro: Meta com ID '${goalId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`deleteGoal: Erro ao excluir meta: ${e.message}`);
    return false;
  }
}


// --- Funções para Dívidas ---

/**
 * Salva ou atualiza uma dívida na aba 'Dividas'.
 * @param {Object} debtData Objeto com os dados da dívida.
 * @returns {boolean} true se a dívida for salva/atualizada, false caso contrário.
 */
function saveDebt(debtData) {
  try {
    const sheet = getSheet(SHEETS.DIVIDAS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const colIndices = {
      id: headers.indexOf('ID'),
      nomeDivida: headers.indexOf('Nome da Dívida'),
      credor: headers.indexOf('Credor'),
      valorTotal: headers.indexOf('Valor Total'),
      valorPago: headers.indexOf('Valor Pago'),
      dataInicio: headers.indexOf('Data Inicio'),
      dataVencimento: headers.indexOf('Data Vencimento'),
      status: headers.indexOf('Status'),
      observacoes: headers.indexOf('Observacoes'),
      quantidadeParcelas: headers.indexOf('Quantidade de Parcelas'), 
      periodicidade: headers.indexOf('Periodicidade') 
    };

    if (Object.values(colIndices).some(idx => idx === -1 && idx !== colIndices.quantidadeParcelas && idx !== colIndices.periodicidade)) {
        throw new Error('Colunas obrigatórias da aba Dividas não encontradas. Verifique os cabeçalhos.');
    }
    if (colIndices.quantidadeParcelas === -1 || colIndices.periodicidade === -1) {
        log("saveDebt: Aviso: Colunas 'Quantidade de Parcelas' ou 'Periodicidade' não encontradas. Verifique se executou a função addInstallmentColumns().");
    }


    if (!debtData.nomeDivida || !debtData.credor || isNaN(parseFloat(debtData.valorTotal)) || parseFloat(debtData.valorTotal) <= 0 || !debtData.dataInicio || !debtData.dataVencimento || !debtData.status) {
      throw new Error('Dados da dívida incompletos ou inválidos.');
    }
    if (isNaN(parseInt(debtData.quantidadeParcelas)) || parseInt(debtData.quantidadeParcelas) < 1) {
      throw new Error('Quantidade de parcelas inválida. Deve ser um número inteiro maior ou igual a 1.');
    }
    if (!debtData.periodicidade) {
      throw new Error('Periodicidade das parcelas é obrigatória.');
    }


    if (debtData.id && debtData.id.startsWith('DIV')) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][colIndices.id] === debtData.id) {
          const rowToUpdate = allData[i]; 

          rowToUpdate[colIndices.nomeDivida] = debtData.nomeDivida;
          rowToUpdate[colIndices.credor] = debtData.credor;
          rowToUpdate[colIndices.valorTotal] = parseFloat(debtData.valorTotal);
          rowToUpdate[colIndices.valorPago] = parseFloat(debtData.valorPago || 0);
          rowToUpdate[colIndices.dataInicio] = debtData.dataInicio;
          rowToUpdate[colIndices.dataVencimento] = debtData.dataVencimento;
          rowToUpdate[colIndices.status] = debtData.status;
          rowToUpdate[colIndices.observacoes] = debtData.observacoes || '';
          
          if (colIndices.quantidadeParcelas !== -1) {
            rowToUpdate[colIndices.quantidadeParcelas] = parseInt(debtData.quantidadeParcelas);
          }
          if (colIndices.periodicidade !== -1) {
            rowToUpdate[colIndices.periodicidade] = debtData.periodicidade;
          }

          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`saveDebt: Dívida '${debtData.nomeDivida}' (ID: ${debtData.id}) atualizada.`);
          return true;
        }
      }
    }

    const nextId = `DIV${sheet.getLastRow() + 1}`;
    const newRow = new Array(headers.length).fill(''); 

    newRow[colIndices.id] = nextId;
    newRow[colIndices.nomeDivida] = debtData.nomeDivida;
    newRow[colIndices.credor] = debtData.credor;
    newRow[colIndices.valorTotal] = parseFloat(debtData.valorTotal);
    newRow[colIndices.valorPago] = parseFloat(debtData.valorPago || 0);
    newRow[colIndices.dataInicio] = debtData.dataInicio;
    newRow[colIndices.dataVencimento] = debtData.dataVencimento;
    newRow[colIndices.status] = debtData.status;
    newRow[colIndices.observacoes] = debtData.observacoes || '';
    
    if (colIndices.quantidadeParcelas !== -1) {
        newRow[colIndices.quantidadeParcelas] = parseInt(debtData.quantidadeParcelas);
    }
    if (colIndices.periodicidade !== -1) {
        newRow[colIndices.periodicidade] = debtData.periodicidade;
    }


    sheet.appendRow(newRow);
    log(`saveDebt: Nova dívida '${debtData.nomeDivida}' (ID: ${nextId}) salva.`);
    return true;
  } catch (e) {
    log(`saveDebt: Erro ao salvar dívida: ${e.message}`);
    return false;
  }
}

/**
 * Registra um pagamento para uma dívida existente.
 * Atualiza 'Valor Pago' e o 'Status' da dívida se for quitada.
 * Também registra a transação de saída correspondente na aba 'Transacoes'.
 * @param {string} debtId ID da dívida.
 * @param {number} paymentAmount Valor do pagamento.
 * @param {string} paymentDate Data do pagamento (formatoISO).
 * @param {string} paymentAccount Conta de origem do pagamento.
 * @param {string} paymentPerson Pessoa que fez o pagamento.
 * @returns {boolean} true se o pagamento for registrado, false caso contrário.
 */
function recordDebtPayment(debtId, paymentAmount, paymentDate, paymentAccount, paymentPerson) {
  try {
    const sheet = getSheet(SHEETS.DIVIDAS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const idColIndex = headers.indexOf('ID');
    const valorPagoColIndex = headers.indexOf('Valor Pago');
    const valorTotalColIndex = headers.indexOf('Valor Total');
    const statusColIndex = headers.indexOf('Status');
    const nomeDividaColIndex = headers.indexOf('Nome da Dívida');

    if (idColIndex === -1 || valorPagoColIndex === -1 || valorTotalColIndex === -1 || statusColIndex === -1 || nomeDividaColIndex === -1) {
        throw new Error('Colunas de dívida não encontradas para registro de pagamento.');
    }
    if (isNaN(paymentAmount) || paymentAmount <= 0) {
        throw new Error('Valor de pagamento inválido. Deve ser um número positivo.');
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idColIndex] === debtId) {
        let currentPaid = parseFloat(allData[i][valorPagoColIndex] || 0);
        const totalDebt = parseFloat(allData[i][valorTotalColIndex] || 0);
        const newPaid = currentPaid + paymentAmount;
        
        sheet.getRange(i + 1, valorPagoColIndex + 1).setValue(newPaid);
        log(`recordDebtPayment: Pagamento de R$ ${paymentAmount.toFixed(2)} registrado para dívida '${allData[i][nomeDividaColIndex]}'.`);

        if (newPaid >= totalDebt) {
          sheet.getRange(i + 1, statusColIndex + 1).setValue('Paga');
          log(`recordDebtPayment: Dívida '${allData[i][nomeDividaColIndex]}' quitada!`);
        } else if (allData[i][statusColIndex] === 'Aguardando Início' && newPaid > 0) {
          sheet.getRange(i + 1, statusColIndex + 1).setValue('Ativa');
          log(`recordDebtPayment: Dívida '${allData[i][nomeDividaColIndex]}' ativada pelo pagamento.`);
        }

        saveTransaction({
          data: paymentDate,
          tipo: 'Saída',
          valor: paymentAmount,
          descricao: `Pagamento de Dívida: ${allData[i][nomeDividaColIndex]}`,
          categoria: 'Dívidas',
          conta: paymentAccount,
          pessoa: paymentPerson,
          tipoPagamento: 'Débito Automático', 
          observacoes: `Pagamento para ${allData[i][nomeDividaColIndex]}`
        });
        log(`recordDebtPayment: Transação de saída para pagamento de dívida registrada.`);
        return true;
      }
    }
    log(`recordDebtPayment: Erro: Dívida com ID '${debtId}' não encontrada para registro de pagamento.`);
    return false;
  } catch (e) {
    log(`recordDebtPayment: Erro ao registrar pagamento de dívida: ${e.message}`);
    return false;
  }
}

/**
 * Exclui uma dívida da aba 'Dividas'.
 * @param {string} debtId ID da dívida a ser excluída.
 * @returns {boolean} true se a dívida for excluída, false caso contrário.
 */
function deleteDebt(debtId) {
  try {
    const sheet = getSheet(SHEETS.DIVIDAS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID');

    if (idColIndex === -1) throw new Error('Coluna ID não encontrada na aba Dividas.');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === debtId) {
        sheet.deleteRow(i + 1);
        log(`deleteDebt: Dívida '${debtId}' excluída.`);
        return true;
      }
    }
    log(`deleteDebt: Erro: Dívida com ID '${debtId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`deleteDebt: Erro ao excluir dívida: ${e.message}`);
    return false;
  }
}


// --- Funções para Investimentos ---

/**
 * Salva ou atualiza um investimento na aba 'Investimentos'.
 * @param {Object} investData Objeto com os dados do investimento.
 * @returns {boolean} true se o investimento for salvo/atualizado, false caso contrário.
 */
function saveInvestment(investData) {
  try {
    const sheet = getSheet(SHEETS.INVESTIMENTOS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const colIndices = {
      id: headers.indexOf('ID'),
      nomeInvestimento: headers.indexOf('Nome do Investimento'),
      instituicao: headers.indexOf('Instituição'),
      valorInicial: headers.indexOf('Valor Inicial'),
      valorAtual: headers.indexOf('Valor Atual'),
      tipo: headers.indexOf('Tipo'),
      rentabilidade: headers.indexOf('Rentabilidade %'),
      dataAporteInicial: headers.indexOf('Data Aporte Inicial'),
      observacoes: headers.indexOf('Observacoes'),
      tipoAporte: headers.indexOf('Tipo de Aporte'), 
      tipoMovimentacao: headers.indexOf('Tipo de Movimentação') 
    };

    if (Object.values(colIndices).some(idx => idx === -1 && idx !== colIndices.tipoAporte && idx !== colIndices.tipoMovimentacao)) {
        throw new Error('Colunas obrigatórias da aba Investimentos não encontradas. Verifique os cabeçalhos.');
    }
    if (colIndices.tipoAporte === -1 || colIndices.tipoMovimentacao === -1) {
        log("saveInvestment: Aviso: Colunas 'Tipo de Aporte' ou 'Tipo de Movimentação' não encontradas. Verifique se executou a função addInvestmentPlanColumns().");
    }

    if (!investData.nomeInvestimento || !investData.instituicao || isNaN(parseFloat(investData.valorInicial)) || parseFloat(investData.valorInicial) <= 0 || !investData.tipo || !investData.dataAporteInicial) {
      throw new Error('Dados do investimento incompletos ou inválidos. Nome, Instituição, Valor Inicial, Tipo e Data de Aporte são obrigatórios.');
    }
    if (!investData.tipoAporte || !investData.tipoMovimentacao) {
        throw new Error('Tipo de Aporte e Tipo de Movimentação são obrigatórios.');
    }

    if (investData.id && investData.id.startsWith('INV')) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][colIndices.id] === investData.id) {
          const rowToUpdate = allData[i]; 

          rowToUpdate[colIndices.nomeInvestimento] = investData.nomeInvestimento;
          rowToUpdate[colIndices.instituicao] = investData.instituicao;
          rowToUpdate[colIndices.valorInicial] = parseFloat(investData.valorInicial);
          rowToUpdate[colIndices.valorAtual] = parseFloat(investData.valorAtual || investData.valorInicial);
          rowToUpdate[colIndices.tipo] = investData.tipo;
          rowToUpdate[colIndices.rentabilidade] = parseFloat(investData.rentabilidade || 0);
          rowToUpdate[colIndices.dataAporteInicial] = investData.dataAporteInicial;
          rowToUpdate[colIndices.observacoes] = investData.observacoes || '';

          if (colIndices.tipoAporte !== -1) {
            rowToUpdate[colIndices.tipoAporte] = investData.tipoAporte;
          }
          if (colIndices.tipoMovimentacao !== -1) {
            rowToUpdate[colIndices.tipoMovimentacao] = investData.tipoMovimentacao;
          }

          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`saveInvestment: Investimento '${investData.nomeInvestimento}' (ID: ${investData.id}) atualizado.`);
          return true;
        }
      }
    }

    const nextId = `INV${sheet.getLastRow() + 1}`;
    const newRow = new Array(headers.length).fill(''); 
    
    newRow[colIndices.id] = nextId;
    newRow[colIndices.nomeInvestimento] = investData.nomeInvestimento;
    newRow[colIndices.instituicao] = investData.instituicao;
    newRow[colIndices.valorInicial] = parseFloat(investData.valorInicial);
    newRow[colIndices.valorAtual] = parseFloat(investData.valorInicial); 
    newRow[colIndices.tipo] = investData.tipo;
    newRow[colIndices.rentabilidade] = 0; 
    newRow[colIndices.dataAporteInicial] = investData.dataAporteInicial;
    newRow[colIndices.observacoes] = investData.observacoes || '';
    
    if (colIndices.tipoAporte !== -1) {
        newRow[colIndices.tipoAporte] = investData.tipoAporte;
    }
    if (colIndices.tipoMovimentacao !== -1) {
        newRow[colIndices.tipoMovimentacao] = investData.tipoMovimentacao;
    }

    sheet.appendRow(newRow);
    log(`saveInvestment: Novo investimento '${investData.nomeInvestimento}' (ID: ${nextId}) salvo.`);
    return true;
  } catch (e) {
    log(`saveInvestment: Erro ao salvar investimento: ${e.message}`);
    return false;
  }
}

/**
 * Atualiza o valor atual e/o u rentabilidade de um investimento.
 * @param {string} investId ID do investimento.
 * @param {number} newCurrentValue Novo valor atual.
 * @param {number} newRentability Nova rentabilidade em percentual (opcional).
 * @returns {boolean} true se atualizado, false caso contrário.
 */
function updateInvestmentValue(investId, newCurrentValue, newRentability) {
  try {
    const sheet = getSheet(SHEETS.INVESTIMENTOS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const idColIndex = headers.indexOf('ID');
    const valorAtualColIndex = headers.indexOf('Valor Atual');
    const rentabilidadeColIndex = headers.indexOf('Rentabilidade %');
    const valorInicialColIndex = headers.indexOf('Valor Inicial');
    const nomeInvestimentoColIndex = headers.indexOf('Nome do Investimento');

    if (idColIndex === -1 || valorAtualColIndex === -1 || rentabilidadeColIndex === -1 || valorInicialColIndex === -1 || nomeInvestimentoColIndex === -1) {
        throw new Error('Colunas de investimento não encontradas para atualização de valor.');
    }
    if (isNaN(newCurrentValue) || newCurrentValue < 0) {
        throw new Error('Novo valor atual inválido. Deve ser um número não negativo.');
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idColIndex] === investId) {
        sheet.getRange(i + 1, valorAtualColIndex + 1).setValue(parseFloat(newCurrentValue));
        log(`updateInvestmentValue: Valor atual de '${allData[i][nomeInvestimentoColIndex]}' atualizado para R$ ${newCurrentValue.toFixed(2)}.`);

        const initialValue = parseFloat(allData[i][valorInicialColIndex] || 0);
        if (initialValue > 0) {
            const calculatedRentability = ((newCurrentValue - initialValue) / initialValue) * 100;
            sheet.getRange(i + 1, rentabilidadeColIndex + 1).setValue(calculatedRentability);
            log(`updateInvestmentValue: Rentabilidade de '${allData[i][nomeInvestimentoColIndex]}' recalculada para ${calculatedRentability.toFixed(2)}%.`);
        } else if (newRentability !== undefined && !isNaN(newRentability)) {
             sheet.getRange(i + 1, rentabilidadeColIndex + 1).setValue(parseFloat(newRentability));
             log(`updateInvestmentValue: Rentabilidade de '${allData[i][nomeInvestimentoColIndex]}' definida para ${newRentability.toFixed(2)}%.`);
        }
        return true;
      }
    }
    log(`updateInvestmentValue: Erro: Investimento com ID '${investId}' não encontrado para atualização de valor.`);
    return false;
  } catch (e) {
    log(`updateInvestmentValue: Erro ao atualizar valor do investimento: ${e.message}`);
    return false;
  }
}

/**
 * Registra um aporte/resgate para um investimento.
 * Atualiza o 'Valor Atual' do investimento e registra a transação na aba 'Transacoes'.
 * @param {Object} aporteData Dados do aporte/resgate: { investId, data, tipoTransacao, valor, conta, pessoa, observacoes }
 * @returns {boolean} true se o aporte for registrado, false caso contrário.
 */
function recordInvestmentMovement(aporteData) {
  try {
    const aportesSheet = getSheet(SHEETS.APORTES_INVESTIMENTOS);
    const investimentosSheet = getSheet(SHEETS.INVESTIMENTOS);
    const investData = investimentosSheet.getDataRange().getValues();
    const investHeaders = investData[0];

    const investIdCol = investHeaders.indexOf('ID');
    const investNomeCol = investHeaders.indexOf('Nome do Investimento');
    const investValorAtualCol = investHeaders.indexOf('Valor Atual');
    const investValorInicialCol = investHeaders.indexOf('Valor Inicial');
    const investRentabilidadeCol = investHeaders.indexOf('Rentabilidade %');

    if (investIdCol === -1 || investNomeCol === -1 || investValorAtualCol === -1 || investValorInicialCol === -1 || investRentabilidadeCol === -1) {
      throw new Error('Colunas da aba Investimentos não encontradas para registrar aporte.');
    }

    let currentInvestmentRow = -1;
    let currentInvestmentName = '';
    let currentInvestmentValue = 0;
    let initialInvestmentValue = 0;

    for (let i = 1; i < investData.length; i++) {
      if (investData[i][investIdCol] === aporteData.investId) {
        currentInvestmentRow = i + 1;
        currentInvestmentName = investData[i][investNomeCol];
        currentInvestmentValue = parseFloat(investData[i][investValorAtualCol] || 0);
        initialInvestmentValue = parseFloat(investData[i][investValorInicialCol] || 0);
        break;
      }
    }

    if (currentInvestmentRow === -1) {
      throw new Error(`recordInvestmentMovement: Investimento com ID '${aporteData.investId}' não encontrado.`);
    }

    if (isNaN(parseFloat(aporteData.valor)) || parseFloat(aporteData.valor) <= 0) {
      throw new Error('Valor do aporte/resgate inválido.');
    }

    const nextAporteId = `AP${aportesSheet.getLastRow() + 1}`;
    aportesSheet.appendRow([
      nextAporteId,
      aporteData.investId,
      aporteData.data,
      aporteData.tipoTransacao,
      parseFloat(aporteData.valor),
      aporteData.conta,
      aporteData.pessoa,
      aporteData.observacoes || ''
    ]);
    log(`recordInvestmentMovement: Aporte/Resgate '${aporteData.tipoTransacao}' de R$ ${parseFloat(aporteData.valor).toFixed(2)} para investimento '${currentInvestmentName}' salvo.`);

    // Atualiza o valor atual do investimento
    let newInvestmentValue = currentInvestmentValue;
    if (aporteData.tipoTransacao === 'Aporte') {
      newInvestmentValue += parseFloat(aporteData.valor);
    } else if (aporteData.tipoTransacao === 'Resgate') {
      newInvestmentValue -= parseFloat(aporteData.valor);
    }
    investimentosSheet.getRange(currentInvestmentRow, investValorAtualCol + 1).setValue(newInvestmentValue);

    if (initialInvestmentValue > 0) {
      const calculatedRentability = ((newInvestmentValue - initialInvestmentValue) / initialInvestmentValue) * 100;
      investimentosSheet.getRange(currentInvestmentRow, investRentabilidadeCol + 1).setValue(calculatedRentability);
    }

    saveTransaction({
      data: aporteData.data,
      tipo: aporteData.tipoTransacao === 'Aporte' ? 'Saída' : 'Entrada', 
      valor: parseFloat(aporteData.valor),
      descricao: `${aporteData.tipoTransacao} em ${currentInvestmentName}`,
      categoria: 'Investimentos', 
      conta: aporteData.conta,
      pessoa: aporteData.pessoa,
      tipoPagamento: 'Transferência Bancária', 
      observacoes: `${aporteData.tipoTransacao} em ${currentInvestmentName}`
    });
    
    return true;
  } catch (e) {
    log(`recordInvestmentMovement: Erro ao registrar aporte/resgate: ${e.message}`);
    return false;
  }
}

/**
 * Exclui um investimento da aba 'Investimentos'.
 * @param {string} investId ID do investimento a ser excluído.
 * @returns {boolean} true se o investimento for excluído, false caso contrário.
 */
function deleteInvestment(investId) {
  try {
    const sheet = getSheet(SHEETS.INVESTIMENTOS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID');

    if (idColIndex === -1) throw new Error('Coluna ID não encontrada na aba Investimentos.');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === investId) {
        sheet.deleteRow(i + 1);
        log(`deleteInvestment: Investimento '${investId}' excluído.`);
        return true;
      }
    }
    log(`deleteInvestment: Erro: Investimento com ID '${investId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`deleteInvestment: Erro ao excluir investimento: ${e.message}`);
    return false;
  }
}


// --- Funções para Análises (Expandidas) ---

/**
 * Retorna dados para gráficos de fluxo de caixa (Entradas vs Saídas) por ano ou mês.
 * @param {number} year O ano para o qual a análise deve ser feita.
 * @param {string} type 'monthly' para análise mensal, 'annual' para análise anual.
 * @returns {Array<Object>} Array de objetos com { period, revenues, expenses, balance }.
 */
function getFinancialFlowAnalysis(year, type) {
  try {
    const transacoesSheet = getSheet(SHEETS.TRANSACOES);
    const transacoes = transacoesSheet.getDataRange().getValues();
    const headers = transacoes[0];

    const dataColIndex = headers.indexOf('Data');
    const tipoColIndex = headers.indexOf('Tipo');
    const valorColIndex = headers.indexOf('Valor (R$)');

    if ([dataColIndex, tipoColIndex, valorColIndex].some(idx => idx === -1)) {
      throw new Error('Colunas de transação não encontradas para fluxo de caixa.');
    }

    const flowData = {}; 

    for (let i = 1; i < transacoes.length; i++) {
      const row = transacoes[i];
      const transactionDate = new Date(row[dataColIndex]);
      const transactionType = row[tipoColIndex];
      const transactionValue = parseFloat(row[valorColIndex] || 0);

      if (!isNaN(transactionDate.getTime()) && transactionDate.getFullYear() === year && !isNaN(transactionValue)) {
        let periodKey;
        if (type === 'monthly') {
          periodKey = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM');
        } else { 
          periodKey = String(transactionDate.getFullYear());
        }

        if (!flowData[periodKey]) {
          flowData[periodKey] = { revenues: 0, expenses: 0 };
        }

        if (transactionType === 'Entrada') {
          flowData[periodKey].revenues += transactionValue;
        } else if (transactionType === 'Saída') {
          flowData[periodKey].expenses += transactionValue;
        }
      }
    }

    const results = Object.keys(flowData).map(period => ({
      period: period,
      revenues: flowData[period].revenues,
      expenses: flowData[period].expenses,
      balance: flowData[period].revenues - flowData[period].expenses
    }));

    results.sort((a, b) => a.period.localeCompare(b.period));

    return results;

  } catch (e) {
    log(`getFinancialFlowAnalysis: Erro ao obter análise de fluxo de caixa: ${e.message}`);
    return [];
  }
}

/**
 * Calcula os gastos médios mensais por categoria.
 * @returns {Object} Um objeto onde as chaves são nomes de categoria e os valores são gastos médios mensais.
 */
function getAverageMonthlySpendings() {
  try {
    const transacoesSheet = getSheet(SHEETS.TRANSACOES);
    const transacoes = transacoesSheet.getDataRange().getValues();
    const headers = transacoes[0];

    const dataColIndex = headers.indexOf('Data');
    const tipoColIndex = headers.indexOf('Tipo');
    const valorColIndex = headers.indexOf('Valor (R$)');
    const categoriaColIndex = headers.indexOf('Categoria');

    if ([dataColIndex, tipoColIndex, valorColIndex, categoriaColIndex].some(idx => idx === -1)) {
      throw new Error('Colunas de transação não encontradas para cálculo de gastos médios.');
    }

    const categoryMonthlySpendings = {}; 
    const categoryTotalSpendings = {};    
    const categoryMonthsCount = {};       

    for (let i = 1; i < transacoes.length; i++) {
      const row = transacoes[i];
      const transactionDate = new Date(row[dataColIndex]);
      const transactionType = row[tipoColIndex];
      const transactionValue = parseFloat(row[valorColIndex] || 0);
      const category = row[categoriaColIndex];

      if (transactionType === 'Saída' && !isNaN(transactionValue) && transactionValue > 0 && category) {
        const monthYear = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'yyyy-MM');

        if (!categoryMonthlySpendings[category]) {
          categoryMonthlySpendings[category] = {};
        }
        if (!categoryMonthlySpendings[category][monthYear]) {
          categoryMonthlySpendings[category][monthYear] = 0;
        }
        categoryMonthlySpendings[category][monthYear] += transactionValue;
      }
    }

    const averageSpendings = {};
    for (const category in categoryMonthlySpendings) {
      let totalMonths = 0;
      let totalSpent = 0;
      for (const monthYear in categoryMonthlySpendings[category]) {
        totalSpent += categoryMonthlySpendings[category][monthYear];
        totalMonths++;
      }
      if (totalMonths > 0) {
        averageSpendings[category] = totalSpent / totalMonths;
      }
    }

    return averageSpendings;

  } catch (e) {
    log(`getAverageMonthlySpendings: Erro ao calcular gastos médios mensais: ${e.message}`);
    return {};
  }
}

/**
 * Sugere categorias de alto gasto e potenciais economias.
 * @returns {Array<Object>} Lista de categorias com altos gastos e sugestões.
 */
function getSpendingSuggestions() {
  try {
    const avgSpendings = getAverageMonthlySpendings();
    const suggestions = [];

    const sortedCategories = Object.keys(avgSpendings).sort((a, b) => avgSpendings[b] - avgSpendings[a]);

    if (sortedCategories.length === 0) {
      return [{ category: 'N/A', averageSpend: 0, suggestion: 'Não há dados de gastos suficientes para gerar sugestões.' }];
    }

    const topCategories = sortedCategories.slice(0, Math.min(sortedCategories.length, 3));

    topCategories.forEach(category => {
      const avg = avgSpendings[category];
      let suggestionText = '';

      if (avg > 500) { 
        suggestionText = `Este é um gasto significativo (R$ ${avg.toFixed(2)}/mês). Considere revisar hábitos como "comer fora", "transporte individual" ou "compras por impulso" para esta categoria.`;
      } else if (avg > 200) {
        suggestionText = `Um gasto moderado (R$ ${avg.toFixed(2)}/mês). Pequenos cortes ou alternativas mais baratas podem fazer diferença ao longo do tempo.`;
      } else {
        suggestionText = `Gasto razoável (R$ ${avg.toFixed(2)}/mês). Mantenha o acompanhamento, mas o impacto de cortes pode ser menor aqui.`;
      }
      suggestions.push({ category: category, averageSpend: avg, suggestion: suggestionText });
    });

    return suggestions;
  } catch (e) {
    log(`getSpendingSuggestions: Erro ao gerar sugestões de gastos: ${e.message}`);
    return [];
  }
}


// --- Funções para Notificações (disparadas por gatilhos de tempo) ---

/**
 * Verifica dívidas e contas atrasadas na aba 'Dividas' e envia um e-mail de alerta.
 * Esta função deve ser configurada para ser executada por um gatilho baseado em tempo
 * (ex: diariamente, toda manhã).
 */
function checkOverdueBillsAndNotify() {
  try {
    const dividasSheet = getSheet(SHEETS.DIVIDAS);
    const dividas = dividasSheet.getDataRange().getValues();
    const headers = dividas[0];
    const today = new Date();
    today.setHours(0, 0, 0, 0); 

    const nomeDividaColIndex = headers.indexOf('Nome da Dívida');
    const dataVencimentoColIndex = headers.indexOf('Data Vencimento');
    const statusColIndex = headers.indexOf('Status');

    if (nomeDividaColIndex === -1 || dataVencimentoColIndex === -1 || statusColIndex === -1) {
      throw new Error('Colunas de dívida (Nome da Dívida, Data Vencimento, Status) não encontradas para verificação de atraso. Verifique os cabeçalhos.');
    }

    const overdueBills = [];

    for (let i = 1; i < dividas.length; i++) {
      const row = dividas[i];
      const status = row[statusColIndex];
      const dataVencimento = new Date(row[dataVencimentoColIndex]);
      dataVencimento.setHours(0, 0, 0, 0); 

      if ((status && typeof status === 'string' && (status.trim() === 'Ativa' || status.trim() === 'Aguardando Início')) && dataVencimento < today) {
        overdueBills.push(row[nomeDividaColIndex]);
        dividasSheet.getRange(i + 1, statusColIndex + 1).setValue('Atrasada');
        log(`checkOverdueBillsAndNotify: Status da dívida '${row[nomeDividaColIndex]}' atualizado para 'Atrasada'.`);
      }
    }

    if (overdueBills.length > 0) {
      const recipientEmail = Session.getActiveUser().getEmail(); 
      const subject = 'Alerta: Contas e Dívidas Atrasadas!';
      const body = `Olá,\n\nVocê tem as seguintes contas/dívidas em atraso:\n\n- ${overdueBills.join('\n- ')}\n\nPor favor, verifique-as no seu controle financeiro familiar para evitar juros e multas.\n\nAtenciosamente,\nSeu Controle Financeiro Familiar`;
      MailApp.sendEmail(recipientEmail, subject, body);
      log(`checkOverdueBillsAndNotify: E-mail de contas atrasadas enviado para ${recipientEmail}. Dívidas: ${overdueBills.join(', ')}`);
    } else {
      log('checkOverdueBillsAndNotify: Nenhuma conta ou dívida atrasada encontrada.');
    }
  } catch (e) {
    log(`checkOverdueBillsAndNotify: Erro ao verificar e notificar contas atrasadas: ${e.message}`);
  }
}

/**
 * Envia um e-mail de parabéns quando uma meta financeira é alcançada.
 * @param {string} goalName O nome da meta que foi alcançada.
 */
function sendGoalReachedEmail(goalName) {
  try {
    const recipientEmail = Session.getActiveUser().getEmail();
    const subject = `🥳 Parabéns! Meta "${goalName}" Alcançada!`;
    const body = `Olá,\n\nQue notícia fantástica! 🎉\n\nA meta "${goalName}" foi atingida com sucesso!\n\nEste é o resultado do seu planejamento e disciplina. Continue assim para alcançar ainda mais objetivos financeiros!\n\nAtenciosamente,\nSeu Controle Financeiro Familiar`;
    MailApp.sendEmail(recipientEmail, subject, body);
    log(`sendGoalReachedEmail: E-mail de meta alcançada enviado para ${recipientEmail} para a meta "${goalName}".`);
  } catch (e) {
    log(`sendGoalReachedEmail: Erro ao enviar e-mail de meta alcançada: ${e.message}`);
  }
}


// --- Funções Comuns para CRUD de Entidades (Categorias, Contas, Pessoas) ---

/**
 * Salva ou atualiza um registro em uma aba genérica.
 * @param {string} sheetName Nome da aba (e.g., 'Categorias', 'Contas', 'Pessoas').
 * @param {Object} recordData Objeto com os dados do registro (deve incluir 'id' para atualização).
 * @param {string} idPrefix Prefixo para o ID (e.g., 'CAT', 'CNT', 'PES').
 * @param {string} nameColumnHeader Cabeçalho da coluna de nome principal (e.g., 'Nome da Categoria', 'Nome da Conta', 'Nome').
 * @returns {boolean} true se salvo/atualizado, false caso contrário.
 */
function saveRecord(sheetName, recordData, idPrefix, nameColumnHeader) {
  try {
    const sheet = getSheet(sheetName);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const nameColIndex = headers.indexOf(nameColumnHeader);
    const idColIndex = headers.indexOf('ID'); 

    if (idColIndex === -1 || nameColIndex === -1) {
      throw new Error(`saveRecord: Colunas 'ID' ou '${nameColumnHeader}' não encontradas na aba '${sheetName}'. Headers: ${headers}`);
    }

    if (!recordData.name) { 
      throw new Error(`saveRecord: Nome do registro (${nameColumnHeader}) é obrigatório.`);
    }

    if (recordData.id && recordData.id.startsWith(idPrefix)) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][idColIndex] === recordData.id) {
          const rowToUpdate = allData[i]; 

          for (const key in recordData) {
            if (Object.prototype.hasOwnProperty.call(recordData, key)) {
              let headerToMatch = key;
              if (key === 'nome' && sheetName === SHEETS.CATEGORIAS) headerToMatch = 'Nome da Categoria';
              if (key === 'nome' && sheetName === SHEETS.CONTAS) headerToMatch = 'Nome da Conta';
              if (key === 'nome' && sheetName === SHEETS.PESSOAS) headerToMatch = 'Nome';

              const headerIndex = headers.indexOf(headerToMatch);
              if (headerIndex !== -1) {
                if (headerToMatch === 'Saldo Inicial' || headerToMatch === 'Saldo Atual') {
                    rowToUpdate[headerIndex] = parseFloat(recordData[key]);
                } else {
                    rowToUpdate[headerIndex] = recordData[key];
                }
              }
            }
          }
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`saveRecord: Registro '${recordData.name}' (ID: ${recordData.id}) atualizado na aba '${sheetName}'.`);
          return true;
        }
      }
    }

    const newId = `${idPrefix}${sheet.getLastRow() + 1}`;
    const newRow = new Array(headers.length).fill(''); 
    newRow[idColIndex] = newId;
    
    for (const key in recordData) {
        if (Object.prototype.hasOwnProperty.call(recordData, key)) {
            let headerToMatch = key;
            if (key === 'nome' && sheetName === SHEETS.CATEGORIAS) headerToMatch = 'Nome da Categoria';
            if (key === 'nome' && sheetName === SHEETS.CONTAS) headerToMatch = 'Nome da Conta';
            if (key === 'nome' && sheetName === SHEETS.PESSOAS) headerToMatch = 'Nome';

            const headerIndex = headers.indexOf(headerToMatch);
            if (headerIndex !== -1) {
                if (headerToMatch === 'Saldo Inicial' || headerToMatch === 'Saldo Atual') {
                    newRow[headerIndex] = parseFloat(recordData[key]);
                } else {
                    newRow[headerIndex] = recordData[key];
                }
            }
        }
    }
    if (nameColIndex !== -1) {
        newRow[nameColIndex] = recordData.name;
    }

    sheet.appendRow(newRow);
    log(`saveRecord: Novo registro '${recordData.name}' (ID: ${newId}) salvo na aba '${sheetName}'.`);
    return true;
  } catch (e) {
    log(`saveRecord: Erro ao salvar registro na aba '${sheetName}': ${e.message}`);
    return false;
  }
}

/**
 * Exclui um registro de uma aba genérica.
 * @param {string} sheetName Nome da aba.
 * @param {string} recordId ID do registro a ser excluído.
 * @returns {boolean} true se excluído, false caso contrário.
 */
function deleteRecord(sheetName, recordId) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID');

    if (idColIndex === -1) {
      throw new Error(`deleteRecord: Coluna 'ID' não encontrada na aba '${sheetName}'.`);
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === recordId) {
        sheet.deleteRow(i + 1);
        log(`deleteRecord: Registro '${recordId}' excluído da aba '${sheetName}'.`);
        return true;
      }
    }
    log(`deleteRecord: Erro: Registro com ID '${recordId}' não encontrado para exclusão na aba '${sheetName}'.`);
    return false;
  } catch (e) {
    log(`deleteRecord: Erro ao excluir registro da aba '${sheetName}': ${e.message}`);
    return false;
  }
}

// --- Funções Específicas para CRUD de Categorias ---
function saveCategory(categoryData) {
    return saveRecord(SHEETS.CATEGORIAS, { id: categoryData.id, name: categoryData.nome, tipo: categoryData.tipo }, 'CAT', 'Nome da Categoria');
}
function deleteCategory(categoryId) {
    return deleteRecord(SHEETS.CATEGORIAS, categoryId);
}

// --- Funções Específicas para CRUD de Contas ---
function saveAccount(accountData) {
    return saveRecord(SHEETS.CONTAS, { id: accountData.id, name: accountData.nome, banco: accountData.banco, saldoInicial: accountData.saldoInicial, saldoAtual: accountData.saldoAtual, tipo: accountData.tipo }, 'CNT', 'Nome da Conta');
}
function deleteAccount(accountId) {
    return deleteRecord(SHEETS.CONTAS, accountId);
}

// --- Funções Específicas para CRUD de Pessoas ---
function savePerson(personData) {
    return saveRecord(SHEETS.PESSOAS, { id: personData.id, name: personData.nome }, 'PES', 'Nome');
}
function deletePerson(personId) {
    return deleteRecord(SHEETS.PESSOAS, personId);
}

/**
 * Adiciona a coluna "Tipo de Pagamento" à aba "Transacoes".
 * Esta função deve ser executada UMA ÚNICA VEZ diretamente no editor do Apps Script.
 */
function addPaymentTypeColumn() {
  try {
    const sheet = getSheet(SHEETS.TRANSACOES);
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    const newColumnHeader = 'Tipo de Pagamento';

    if (headers.includes(newColumnHeader)) {
      log(`addPaymentTypeColumn: A coluna '${newColumnHeader}' já existe na aba '${SHEETS.TRANSACOES}'. Nenhuma ação necessária.`);
      return;
    }

    sheet.insertColumnsAfter(lastColumn, 1);
    sheet.getRange(1, lastColumn + 1).setValue(newColumnHeader);
    
    log(`addPaymentTypeColumn: Coluna '${newColumnHeader}' adicionada com sucesso à aba '${SHEETS.TRANSACOES}'.`);
  } catch (e) {
    log(`addPaymentTypeColumn: Erro ao adicionar coluna 'Tipo de Pagamento': ${e.message}`);
  }
}

/**
 * Adiciona as colunas "Quantidade de Parcelas" e "Periodicidade" à aba "Dividas".
 * Esta função deve ser executada UMA ÚNICA VEZ diretamente no editor do Apps Script.
 */
function addInstallmentColumns() {
  try {
    const sheet = getSheet(SHEETS.DIVIDAS);
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    const newColumn1Header = 'Quantidade de Parcelas';
    const newColumn2Header = 'Periodicidade';

    let columnsAdded = false;

    if (!headers.includes(newColumn1Header)) {
      sheet.insertColumnsAfter(lastColumn, 1);
      sheet.getRange(1, lastColumn + 1).setValue(newColumn1Header);
      log(`addInstallmentColumns: Coluna '${newColumn1Header}' adicionada com sucesso à aba '${SHEETS.DIVIDAS}'.`);
      columnsAdded = true;
    } else {
      log(`addInstallmentColumns: A coluna '${newColumn1Header}' já existe na aba '${SHEETS.DIVIDAS}'. Nenhuma ação necessária.`);
    }

    const currentLastColumn = sheet.getLastColumn();

    if (!headers.includes(newColumn2Header)) {
      sheet.insertColumnsAfter(currentLastColumn, 1);
      sheet.getRange(1, currentLastColumn + 1).setValue(newColumn2Header);
      log(`addInstallmentColumns: Coluna '${newColumn2Header}' adicionada com sucesso à aba '${SHEETS.DIVIDAS}'.`);
      columnsAdded = true;
    } else {
      log(`addInstallmentColumns: A coluna '${newColumn2Header}' já existe na aba '${SHEETS.DIVIDAS}'. Nenhuma ação necessária.`);
    }

    if (!columnsAdded) {
      log('addInstallmentColumns: Nenhuma nova coluna de parcela foi adicionada. Ambas já existiam.');
    }

  } catch (e) {
    log(`addInstallmentColumns: Erro ao adicionar colunas de parcelamento: ${e.message}`);
  }
}

/**
 * Adiciona as colunas "Tipo de Aporte" e "Tipo de Movimentação" à aba "Investimentos".
 * Esta função deve ser executada UMA ÚNICA VEZ diretamente no editor do Apps Script.
 */
function addInvestmentPlanColumns() {
  try {
    const sheet = getSheet(SHEETS.INVESTIMENTOS);
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    const newColumn1Header = 'Tipo de Aporte';
    const newColumn2Header = 'Tipo de Movimentação';

    let columnsAdded = false;

    if (!headers.includes(newColumn1Header)) {
      sheet.insertColumnsAfter(lastColumn, 1);
      sheet.getRange(1, lastColumn + 1).setValue(newColumn1Header);
      log(`addInvestmentPlanColumns: Coluna '${newColumn1Header}' adicionada com sucesso à aba '${SHEETS.INVESTIMENTOS}'.`);
      columnsAdded = true;
    } else {
      log(`addInvestmentPlanColumns: A coluna '${newColumn1Header}' já existe na aba '${SHEETS.INVESTIMENTOS}'. Nenhuma ação necessária.`);
    }

    const currentLastColumn = sheet.getLastColumn();

    if (!headers.includes(newColumn2Header)) {
      sheet.insertColumnsAfter(currentLastColumn, 1);
      sheet.getRange(1, currentLastColumn + 1).setValue(newColumn2Header);
      log(`addInvestmentPlanColumns: Coluna '${newColumn2Header}' adicionada com sucesso à aba '${SHEETS.INVESTIMENTOS}'.`);
      columnsAdded = true;
    } else {
      log(`addInvestmentPlanColumns: A coluna '${newColumn2Header}' já existe na aba '${SHEETS.INVESTIMENTOS}'. Nenhuma ação necessária.`);
    }

    if (!columnsAdded) {
      log('addInvestmentPlanColumns: Nenhuma nova coluna de plano de investimento foi adicionada. Ambas já existiam.');
    }

  } catch (e) {
    log(`addInvestmentPlanColumns: Erro ao adicionar colunas de plano de investimento: ${e.message}`);
  }
}