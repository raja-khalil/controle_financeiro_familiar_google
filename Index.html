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
    return sheet.getDataRange().getValues();
  } catch (e) {
    log(`Erro ao obter dados da aba '${sheetName}': ${e.message}`);
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
  for (const key in sheetsToFetch) {
    if (Object.prototype.hasOwnProperty.call(sheetsToFetch, key)) {
      result[key] = getSheetData(sheetsToFetch[key]);
    }
  }
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
        log(`Conta '${accountName}' atualizada para R$ ${newBalance.toFixed(2)}.`);
        return true;
      }
    }
    log(`Erro: Conta '${accountName}' não encontrada para atualização de saldo.`);
    return false;
  } catch (e) {
    log(`Erro ao atualizar saldo da conta: ${e.message}`);
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

    if (!transaction.data || !transaction.tipo || !transaction.valor || !transaction.conta || !transaction.descricao || !transaction.categoria || !transaction.pessoa) {
      throw new Error('Dados da transação incompletos. Verifique Data, Tipo, Valor, Descrição, Categoria, Conta e Pessoa.');
    }
    const valorNumerico = parseFloat(transaction.valor);
    if (isNaN(valorNumerico) || valorNumerico <= 0) {
      throw new Error('Valor da transação inválido. Deve ser um número positivo.');
    }

    const nextId = `TR${transacoesSheet.getLastRow() + 1}`; 
    const valorParaContas = transaction.tipo === 'Saída' ? -valorNumerico : valorNumerico;

    transacoesSheet.appendRow([
      nextId,
      transaction.data,
      transaction.tipo,
      valorNumerico,
      transaction.descricao,
      transaction.categoria,
      transaction.conta,
      transaction.pessoa,
      transaction.observacoes || ''
    ]);
    log(`Transação '${transaction.descricao}' (${transaction.tipo}) salva.`);

    updateAccountBalance(transaction.conta, valorParaContas);
    log(`Saldo da conta '${transaction.conta}' atualizado.`);

    return true;
  } catch (e) {
    log(`Erro ao salvar transação: ${e.message}`);
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

    if (idColIndex === -1) throw new Error('Coluna ID não encontrada na aba Transacoes.');

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

        const rowToUpdate = new Array(headers.length);
        rowToUpdate[idColIndex] = transactionData.id;
        rowToUpdate[dataColIndex] = transactionData.data;
        rowToUpdate[tipoColIndex] = newTipo;
        rowToUpdate[valorColIndex] = newValor;
        rowToUpdate[descricaoColIndex] = transactionData.descricao;
        rowToUpdate[categoriaColIndex] = transactionData.categoria;
        rowToUpdate[contaColIndex] = newConta;
        rowToUpdate[pessoaColIndex] = transactionData.pessoa;
        rowToUpdate[observacoesColIndex] = transactionData.observacoes || '';

        sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);

        // Aplica o novo valor na nova conta (se mudou)
        const valorParaAplicar = newTipo === 'Saída' ? -newValor : newValor;
        updateAccountBalance(newConta, valorParaAplicar);

        log(`Transação '${transactionData.id}' atualizada.`);
        return true;
      }
    }
    log(`Erro: Transação com ID '${transactionData.id}' não encontrada para atualização.`);
    return false;
  } catch (e) {
    log(`Erro ao atualizar transação: ${e.message}`);
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

    if (idColIndex === -1) throw new Error('Coluna ID não encontrada na aba Transacoes.');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === transactionId) {
        const tipo = data[i][tipoColIndex];
        const valor = parseFloat(data[i][valorColIndex] || 0);
        const conta = data[i][contaColIndex];

        // Estorna o valor da conta (o oposto do que foi registrado)
        const valorParaEstornar = tipo === 'Saída' ? valor : -valor;
        updateAccountBalance(conta, valorParaEstornar);

        sheet.deleteRow(i + 1);
        log(`Transação '${transactionId}' excluída.`);
        return true;
      }
    }
    log(`Erro: Transação com ID '${transactionId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`Erro ao excluir transação: ${e.message}`);
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
    const tipoColIndex = headers.indexOf('Tipo'); // NOVO: Tipo de orçamento (Receita/Despesa)
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
          row[tipoColIndex] === budgetData.tipo) { // Inclui tipo na comparação
        
        sheet.getRange(i + 1, valorOrcadoColIndex + 1).setValue(parseFloat(budgetData.valorOrcado));
        found = true;
        log(`Orçamento para '${budgetData.produtoServico}' em '${budgetData.categoria}' (${budgetData.tipo}, ${budgetData.anoMes}) atualizado.`);
        break;
      }
    }
    if (!found) {
      sheet.appendRow([budgetData.anoMes, budgetData.categoria, budgetData.produtoServico, budgetData.tipo, parseFloat(budgetData.valorOrcado)]);
      log(`Novo orçamento para '${budgetData.produtoServico}' em '${budgetData.categoria}' (${budgetData.tipo}, ${budgetData.anoMes}) salvo.`);
    }
    return true;
  } catch (e) {
    log(`Erro ao salvar orçamento: ${e.message}`);
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
        if (!orcamentosPorCategoria.despesas.hasOwnProperty(categoria) && !receitaRealPorCategoria.hasOwnProperty(categoria)) {
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
        if (!orcamentosPorCategoria.receitas.hasOwnProperty(categoria) && !gastosPorCategoria.hasOwnProperty(categoria)) {
            receitasResults.push({
                categoria: categoria,
                estimado: 0,
                realizado: receitaRealPorCategoria[categoria]
            });
        }
    }

    return { receitas: receitasResults, despesas: despesasResults };
  } catch (e) {
    log(`Erro ao obter análise de orçamento: ${e.message}`);
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
          log(`Meta '${goalData.nome}' (ID: ${goalData.id}) atualizada.`);
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
    log(`Nova meta '${goalData.nome}' (ID: ${nextId}) salva.`);
    return true;
  } catch (e) {
    log(`Erro ao salvar meta: ${e.message}`);
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

    if (Object.values(colIndices).some(idx => idx === -1)) {
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
          log(`Meta '${allData[i][nomeMetaColIndex]}' (ID: ${goalId}) alcançada.`);
        }
        log(`Contribuição de R$ ${amount.toFixed(2)} adicionada à meta '${allData[i][nomeMetaColIndex]}'.`);
        return true;
      }
    }
    log(`Erro: Meta com ID '${goalId}' não encontrada para adicionar contribuição.`);
    return false;
  } catch (e) {
    log(`Erro ao contribuir para meta: ${e.message}`);
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
        log(`Meta '${goalId}' excluída.`);
        return true;
      }
    }
    log(`Erro: Meta com ID '${goalId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`Erro ao excluir meta: ${e.message}`);
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
      observacoes: headers.indexOf('Observacoes')
    };

    if (Object.values(colIndices).some(idx => idx === -1)) {
        throw new Error('Colunas da aba Dividas não encontradas. Verifique os cabeçalhos.');
    }

    if (!debtData.nomeDivida || !debtData.credor || isNaN(parseFloat(debtData.valorTotal)) || parseFloat(debtData.valorTotal) <= 0 || !debtData.dataInicio || !debtData.dataVencimento || !debtData.status) {
      throw new Error('Dados da dívida incompletos ou inválidos.');
    }

    if (debtData.id && debtData.id.startsWith('DIV')) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][colIndices.id] === debtData.id) {
          const rowToUpdate = new Array(headers.length);
          rowToUpdate[colIndices.id] = debtData.id;
          rowToUpdate[colIndices.nomeDivida] = debtData.nomeDivida;
          rowToUpdate[colIndices.credor] = debtData.credor;
          rowToUpdate[colIndices.valorTotal] = parseFloat(debtData.valorTotal);
          rowToUpdate[colIndices.valorPago] = parseFloat(debtData.valorPago || 0);
          rowToUpdate[colIndices.dataInicio] = debtData.dataInicio;
          rowToUpdate[colIndices.dataVencimento] = debtData.dataVencimento;
          rowToUpdate[colIndices.status] = debtData.status;
          rowToUpdate[colIndices.observacoes] = debtData.observacoes || '';

          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`Dívida '${debtData.nomeDivida}' (ID: ${debtData.id}) atualizada.`);
          return true;
        }
      }
    }

    const nextId = `DIV${sheet.getLastRow() + 1}`;
    sheet.appendRow([
      nextId,
      debtData.nomeDivida,
      debtData.credor,
      parseFloat(debtData.valorTotal),
      parseFloat(debtData.valorPago || 0),
      debtData.dataInicio,
      debtData.dataVencimento,
      debtData.status,
      debtData.observacoes || ''
    ]);
    log(`Nova dívida '${debtData.nomeDivida}' (ID: ${nextId}) salva.`);
    return true;
  } catch (e) {
    log(`Erro ao salvar dívida: ${e.message}`);
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

    if (Object.values(colIndices).some(idx => idx === -1)) {
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
        log(`Pagamento de R$ ${paymentAmount.toFixed(2)} registrado para dívida '${allData[i][nomeDividaColIndex]}'.`);

        if (newPaid >= totalDebt) {
          sheet.getRange(i + 1, statusColIndex + 1).setValue('Paga');
          log(`Dívida '${allData[i][nomeDividaColIndex]}' quitada!`);
        } else if (allData[i][statusColIndex] === 'Aguardando Início' && newPaid > 0) {
          sheet.getRange(i + 1, statusColIndex + 1).setValue('Ativa');
          log(`Dívida '${allData[i][nomeDividaColIndex]}' ativada pelo pagamento.`);
        }

        saveTransaction({
          data: paymentDate,
          tipo: 'Saída',
          valor: paymentAmount,
          descricao: `Pagamento de Dívida: ${allData[i][nomeDividaColIndex]}`,
          categoria: 'Dívidas',
          conta: paymentAccount,
          pessoa: paymentPerson,
          observacoes: `Pagamento para ${allData[i][nomeDividaColIndex]}`
        });
        log(`Transação de saída para pagamento de dívida registrada.`);
        return true;
      }
    }
    log(`Erro: Dívida com ID '${debtId}' não encontrada para registro de pagamento.`);
    return false;
  } catch (e) {
    log(`Erro ao registrar pagamento de dívida: ${e.message}`);
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
        log(`Dívida '${debtId}' excluída.`);
        return true;
      }
    }
    log(`Erro: Dívida com ID '${debtId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`Erro ao excluir dívida: ${e.message}`);
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
      observacoes: headers.indexOf('Observacoes')
    };

    if (Object.values(colIndices).some(idx => idx === -1)) {
        throw new Error('Colunas da aba Investimentos não encontradas. Verifique os cabeçalhos.');
    }

    if (!investData.nomeInvestimento || !investData.instituicao || isNaN(parseFloat(investData.valorInicial)) || parseFloat(investData.valorInicial) <= 0 || !investData.tipo || !investData.dataAporteInicial) {
      throw new Error('Dados do investimento incompletos ou inválidos. Nome, Instituição, Valor Inicial, Tipo e Data de Aporte são obrigatórios.');
    }

    if (investData.id && investData.id.startsWith('INV')) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][colIndices.id] === investData.id) {
          const rowToUpdate = new Array(headers.length);
          rowToUpdate[colIndices.id] = investData.id;
          rowToUpdate[colIndices.nomeInvestimento] = investData.nomeInvestimento;
          rowToUpdate[colIndices.instituicao] = investData.instituicao;
          rowToUpdate[colIndices.valorInicial] = parseFloat(investData.valorInicial);
          rowToUpdate[colIndices.valorAtual] = parseFloat(investData.valorAtual || investData.valorInicial);
          rowToUpdate[colIndices.tipo] = investData.tipo;
          rowToUpdate[colIndices.rentabilidade] = parseFloat(investData.rentabilidade || 0);
          rowToUpdate[colIndices.dataAporteInicial] = investData.dataAporteInicial;
          rowToUpdate[colIndices.observacoes] = investData.observacoes || '';

          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowToUpdate]);
          log(`Investimento '${investData.nomeInvestimento}' (ID: ${investData.id}) atualizado.`);
          return true;
        }
      }
    }

    const nextId = `INV${sheet.getLastRow() + 1}`;
    sheet.appendRow([
      nextId,
      investData.nomeInvestimento,
      investData.instituicao,
      parseFloat(investData.valorInicial),
      parseFloat(investData.valorInicial),
      investData.tipo,
      0, // Rentabilidade inicial 0%
      investData.dataAporteInicial,
      investData.observacoes || ''
    ]);
    log(`Novo investimento '${investData.nomeInvestimento}' (ID: ${nextId}) salvo.`);
    return true;
  } catch (e) {
    log(`Erro ao salvar investimento: ${e.message}`);
    return false;
  }
}

/**
 * Atualiza o valor atual e/ou rentabilidade de um investimento.
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

    if (Object.values(colIndices).some(idx => idx === -1)) {
        throw new Error('Colunas de investimento não encontradas para atualização de valor.');
    }
    if (isNaN(newCurrentValue) || newCurrentValue < 0) {
        throw new Error('Novo valor atual inválido. Deve ser um número não negativo.');
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idColIndex] === investId) {
        sheet.getRange(i + 1, valorAtualColIndex + 1).setValue(parseFloat(newCurrentValue));
        log(`Valor atual de '${allData[i][nomeInvestimentoColIndex]}' atualizado para R$ ${newCurrentValue.toFixed(2)}.`);

        const initialValue = parseFloat(allData[i][valorInicialColIndex] || 0);
        if (initialValue > 0) {
            const calculatedRentability = ((newCurrentValue - initialValue) / initialValue) * 100;
            sheet.getRange(i + 1, rentabilidadeColIndex + 1).setValue(calculatedRentability);
            log(`Rentabilidade de '${allData[i][nomeInvestimentoColIndex]}' recalculada para ${calculatedRentability.toFixed(2)}%.`);
        } else if (newRentability !== undefined && !isNaN(newRentability)) {
             sheet.getRange(i + 1, rentabilidadeColIndex + 1).setValue(parseFloat(newRentability));
             log(`Rentabilidade de '${allData[i][nomeInvestimentoColIndex]}' definida para ${newRentability.toFixed(2)}%.`);
        }
        return true;
      }
    }
    log(`Erro: Investimento com ID '${investId}' não encontrado para atualização de valor.`);
    return false;
  } catch (e) {
    log(`Erro ao atualizar valor do investimento: ${e.message}`);
    return false;
  }
}

/**
 * Registra um aporte/resgate para um investimento.
 * Atualiza o 'Valor Atual' do investimento e registra a transação na aba 'Transacoes'.
 * @param {Object} aporteData Dados do aporte/resgate: { investId, data, tipoTransacao, valor, conta, pessoa, observacoes }
 * @returns {boolean} true se o aporte for registrado, false caso contrário.
 */
function recordInvestmentAporte(aporteData) {
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

    if (Object.values(investHeaders).some(idx => idx === -1)) {
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
      throw new Error(`Investimento com ID '${aporteData.investId}' não encontrado.`);
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
    log(`Aporte/Resgate '${aporteData.tipoTransacao}' de R$ ${parseFloat(aporteData.valor).toFixed(2)} para investimento '${currentInvestmentName}' salvo.`);

    // Atualiza o valor atual do investimento
    let newInvestmentValue = currentInvestmentValue;
    if (aporteData.tipoTransacao === 'Aporte') {
      newInvestmentValue += parseFloat(aporteData.valor);
    } else if (aporteData.tipoTransacao === 'Resgate') {
      newInvestmentValue -= parseFloat(aporteData.valor);
    }
    investimentosSheet.getRange(currentInvestmentRow, investValorAtualCol + 1).setValue(newInvestmentValue);

    // Recalcula rentabilidade (se o valor inicial for maior que 0)
    if (initialInvestmentValue > 0) {
      const calculatedRentability = ((newInvestmentValue - initialInvestmentValue) / initialInvestmentValue) * 100;
      investimentosSheet.getRange(currentInvestmentRow, investRentabilidadeCol + 1).setValue(calculatedRentability);
    }

    // Registra a transação na aba 'Transacoes'
    saveTransaction({
      data: aporteData.data,
      tipo: aporteData.tipoTransacao === 'Aporte' ? 'Saída' : 'Entrada', // Aporte é saída da conta, Resgate é entrada na conta
      valor: parseFloat(aporteData.valor),
      descricao: `${aporteData.tipoTransacao} em ${currentInvestmentName}`,
      categoria: 'Investimentos', // Categoria genérica para investimentos
      conta: aporteData.conta,
      pessoa: aporteData.pessoa,
      observacoes: `${aporteData.tipoTransacao} em ${currentInvestmentName}`
    });
    
    return true;
  } catch (e) {
    log(`Erro ao registrar aporte/resgate: ${e.message}`);
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
        log(`Investimento '${investId}' excluído.`);
        return true;
      }
    }
    log(`Erro: Investimento com ID '${investId}' não encontrada para exclusão.`);
    return false;
  } catch (e) {
    log(`Erro ao excluir investimento: ${e.message}`);
    return false;
  }
}


// --- Funções para Análises (Expandidas) ---

/**
 * Retorna dados para gráficos de fluxo de caixa (Entradas vs Saídas) por mês/ano.
 * @returns {Array<Object>} Array de objetos com { anoMes, totalEntrada, totalSaida }.
 */
function getMonthlyCashFlow() {
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

    const cashFlow = {}; // { 'YYYY-MM': { entrada: X, saida: Y } }

    for (let i = 1; i < transacoes.length; i++) {
      const row = transacoes[i];
      const date = new Date(row[dataColIndex]);
      const type = row[tipoColIndex];
      const value = parseFloat(row[valorColIndex] || 0);

      if (!isNaN(date.getTime()) && !isNaN(value) && value > 0) {
        const monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
        if (!cashFlow[monthYear]) {
          cashFlow[monthYear] = { totalEntrada: 0, totalSaida: 0 };
        }
        if (type === 'Entrada') {
          cashFlow[monthYear].totalEntrada += value;
        } else if (type === 'Saída') {
          cashFlow[monthYear].totalSaida += value;
        }
      }
    }

    // Converte para array de objetos e ordena por data
    const results = Object.keys(cashFlow).map(monthYear => ({
      anoMes: monthYear,
      totalEntrada: cashFlow[monthYear].totalEntrada,
      totalSaida: cashFlow[monthYear].totalSaida,
      saldo: cashFlow[monthYear].totalEntrada - cashFlow[monthYear].totalSaida
    })).sort((a, b) => a.anoMes.localeCompare(b.anoMes));

    return results;
  } catch (e) {
    log(`Erro ao obter fluxo de caixa mensal: ${e.message}`);
    return [];
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
      return [{ categoria: 'N/A', sugestao: 'Não há dados de gastos suficientes para gerar sugestões.' }];
    }

    // Top 3 categorias de maior gasto
    const topCategories = sortedCategories.slice(0, 3);

    topCategories.forEach(category => {
      const avg = avgSpendings[category];
      let suggestionText = '';

      if (avg > 500) { // Exemplo de limite, ajuste conforme necessário
        suggestionText = `Este é um gasto significativo (R$ ${avg.toFixed(2)}/mês). Considere revisar hábitos como "comer fora", "transporte individual" ou "compras por impulso" para esta categoria.`;
      } else if (avg > 200) {
        suggestionText = `Um gasto moderado (R$ ${avg.toFixed(2)}/mês). Pequenos cortes ou alternativas mais baratas podem fazer diferença ao longo do tempo.`;
      } else {
        suggestionText = `Gasto razoável (R$ ${avg.toFixed(2)}/mês). Mantenha o acompanhamento, mas o impacto de cortes pode ser menor aqui.`;
      }
      suggestions.push({ categoria: category, gastoMedio: avg, sugestao: suggestionText });
    });

    return suggestions;
  } catch (e) {
    log(`Erro ao gerar sugestões de gastos: ${e.message}`);
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
    today.setHours(0, 0, 0, 0); // Normaliza a data de hoje para comparação (sem hora/min/seg)

    const nomeDividaColIndex = headers.indexOf('Nome da Dívida');
    const dataVencimentoColIndex = headers.indexOf('Data Vencimento');
    const statusColIndex = headers.indexOf('Status');

    if (Object.values(colIndices).some(idx => idx === -1)) {
      throw new Error('Colunas de dívida (Nome da Dívida, Data Vencimento, Status) não encontradas para verificação de atraso. Verifique os cabeçalhos.');
    }

    const overdueBills = [];

    for (let i = 1; i < dividas.length; i++) {
      const row = dividas[i];
      const status = row[statusColIndex];
      const dataVencimento = new Date(row[dataVencimentoColIndex]);
      dataVencimento.setHours(0, 0, 0, 0); // Normaliza a data de vencimento

      if ((status === 'Ativa' || status === 'Aguardando Início') && dataVencimento < today) {
        overdueBills.push(row[nomeDividaColIndex]);
        dividasSheet.getRange(i + 1, statusColIndex + 1).setValue('Atrasada');
        log(`Status da dívida '${row[nomeDividaColIndex]}' atualizado para 'Atrasada'.`);
      }
    }

    if (overdueBills.length > 0) {
      const recipientEmail = Session.getActiveUser().getEmail(); 
      const subject = 'Alerta: Contas e Dívidas Atrasadas!';
      const body = `Olá,\n\nVocê tem as seguintes contas/dívidas em atraso:\n\n- ${overdueBills.join('\n- ')}\n\nPor favor, verifique-as no seu controle financeiro familiar para evitar juros e multas.\n\nAtenciosamente,\nSeu Controle Financeiro Familiar`;
      MailApp.sendEmail(recipientEmail, subject, body);
      log(`E-mail de contas atrasadas enviado para ${recipientEmail}. Dívidas: ${overdueBills.join(', ')}`);
    } else {
      log('Nenhuma conta ou dívida atrasada encontrada.');
    }
  } catch (e) {
    log(`Erro ao verificar e notificar contas atrasadas: ${e.message}`);
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
    log(`E-mail de meta alcançada enviado para ${recipientEmail} para a meta "${goalName}".`);
  } catch (e) {
    log(`Erro ao enviar e-mail de meta alcançada: ${e.message}`);
  }
}