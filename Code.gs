/**
 * @OnlyCurrentDoc
 *
 * Servidor Backend (Code.gs) para o Dashboard de Gerenciamento da JRS-HNaRe.
 * Este arquivo lida com todo o acesso à planilha, lógica de negócios e
 * serve o frontend (Index.html).
 */

// --- CONSTANTES GLOBAIS ---
const SPREADSHEET_ID = "12_X8hKR4T_ok33Tv-M8rwpKSUeJNwIAjo9rWzfoA2Nw";
const SHEET_DATA = "ListaControle";
const SHEET_REF = "ListaRef";

// Cache da planilha para evitar múltiplas chamadas 'openById'
const SS = SpreadsheetApp.openById(SPREADSHEET_ID);

// Mapeamento das colunas da 'ListaControle' para índices (base 1)
// Isso torna o código mais legível e fácil de manter.
const COL_MAP = {
  IS: 1,
  DataEntrevista: 2,
  Finalidade: 3,
  OM: 4,
  PGQ: 5,
  NIP: 6,
  Inspecionado: 7,
  StatusIS: 8,
  DataLaudo: 9,
  Laudo: 10,
  Restricoes: 11,
  TIS: 12,
  DS1a: 13,
  MSG: 14
};

// Mapeamento das colunas da 'ListaRef' (base 1)
const COL_MAP_REF = {
  FINALIDADES: 1,
  OM: 2,
  PG: 3,
  QUADRO_OF: 4,
  QUADRO_PR: 5,
  STATUS: 6,
  RESTRICOES: 7
};

// --- FUNÇÃO PRINCIPAL (WEB APP) ---

/**
 * Serve o Web App.
 * Esta função é chamada quando qualquer usuário (GET) acessa a URL do app.
 * @param {object} e O objeto de evento do Apps Script.
 * @returns {HtmlOutput} O conteúdo HTML da página.
 */
function doGet(e) {
  let html = HtmlService.createHtmlOutputFromFile('Index.html')
    .setTitle("JRS-HNaRe - Dashboard de Gerenciamento")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  // Permite o embedding em outros sites (como Google Sites), se necessário
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return html;
}

// --- FUNÇÕES DE LEITURA (EXPOSTAS AO CLIENTE) ---

/**
 * Busca os dados de referência (para filtros) da aba 'ListaRef'.
 * @returns {object} Um objeto contendo arrays para os dropdowns de filtro.
 */
function getRefData() {
  try {
    const refSheet = SS.getSheetByName(SHEET_REF);
    if (!refSheet) throw new Error("Aba 'ListaRef' não encontrada.");

    const finalidades = _getUniqueColumnValues(refSheet, COL_MAP_REF.FINALIDADES);
    const om = _getUniqueColumnValues(refSheet, COL_MAP_REF.OM);
    const status = _getUniqueColumnValues(refSheet, COL_MAP_REF.STATUS);
    const restricoes = _getUniqueColumnValues(refSheet, COL_MAP_REF.RESTRICOES);

    return {
      finalidades: finalidades,
      om: om,
      status: status,
      restricoes: restricoes
    };
  } catch (e) {
    Logger.log(`Erro em getRefData: ${e.message}`);
    throw new Error(`Não foi possível carregar os dados de referência. Detalhe: ${e.message}`);
  }
}

/**
 * Busca, filtra e processa os dados da 'ListaControle' para o dashboard e tabela.
 * @param {object} filters Um objeto contendo os filtros do cliente (ex: { dataInicio, dataFim, finalidade, om, statusIS }).
 * @returns {object} Um objeto com { kpis, chartsData, tableData }.
 */
function getSheetData(filters) {
  try {
    const sheet = SS.getSheetByName(SHEET_DATA);
    if (!sheet) throw new Error("Aba 'ListaControle' não encontrada.");

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const allDataValues = dataRange.getValues();
    const headers = _getHeaders();

    // Converte todos os dados para objetos para facilitar a manipulação
    const allDataObjects = _convertDataToObjects(allDataValues, headers);

    // --- 1. Lógica de Filtro ---
    const filteredData = _applyFilters(allDataObjects, filters);

    // --- 2. Lógica de KPI (com comparativo) ---
    const kpis = _calculateKpis(allDataObjects, filters);

    // --- 3. Lógica de Gráficos ---
    const chartsData = _prepareChartData(filteredData);

    // --- 4. Lógica da Tabela ---
    // Prepara os dados para a tabela, stringificando datas.
    const tableData = filteredData.map(row => ({
      ...row,
      DataEntrevista: row.DataEntrevista ? row.DataEntrevista.toISOString() : null,
      DataLaudo: row.DataLaudo ? row.DataLaudo.toISOString() : null
    }));

    return {
      kpis: kpis,
      chartsData: chartsData,
      tableData: tableData
    };

  } catch (e) {
    Logger.log(`Erro em getSheetData: ${e.message} \n ${e.stack}`);
    throw new Error(`Erro ao processar dados da planilha. Detalhe: ${e.message}`);
  }
}

// --- FUNÇÕES DE ESCRITA (CRUD - ROTEADOR) ---

/**
 * Roteador central para todas as ações de escrita (CRUD) na planilha.
 * @param {string} action O nome da ação a ser executada (ex: 'INSERIR_LAUDO').
 * @param {object} payload Os dados necessários para a ação (ex: { rowIndex, laudo }).
 * @returns {object} Um objeto de status de sucesso ou lança um erro.
 */
function executeCrudAction(action, payload) {
  // Trava o serviço para evitar concorrência de escrita
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // Espera até 10 segundos

  try {
    Logger.log(`Executando Ação: ${action} com payload: ${JSON.stringify(payload)}`);
    
    // Roteador de ações
    switch (action) {
      case 'REMARCAR_IS':
        return _remarcarIS(payload);
      case 'INSERIR_LAUDO':
        return _inserirLaudo(payload);
      case 'INSERIR_RESTRICOES':
        return _inserirRestricoes(payload);
      case 'AUDITORIA':
        return _mudarStatus(payload.rowIndex, 'Concluída', 'Auditoria');
      case 'RESTITUIR_AUDITORIA':
        return _mudarStatus(payload.rowIndex, 'Auditoria', 'Restituída Auditoria');
      case 'REVISAO_JSD':
        return _mudarStatus(payload.rowIndex, 'Concluída', 'JSD');
      case 'RESTITUIR_JSD':
        return _mudarStatus(payload.rowIndex, 'JSD', 'Restituída JSD');
      case 'INSERIR_TIS':
        return _inserirTIS(payload);
      case 'INSERIR_DS1A':
        return _inserirDS1a(payload);
      case 'REGISTRAR_MSG':
        return _registrarMSG(payload);
      default:
        throw new Error(`Ação CRUD desconhecida: ${action}`);
    }
  } catch (e) {
    Logger.log(`Falha na Ação CRUD (${action}): ${e.message}`);
    throw e; // Repassa o erro para o frontend
  } finally {
    lock.releaseLock(); // Sempre libera a trava
  }
}


// --- FUNÇÕES DE ASSESSORIA (ETAPA 6) ---

/**
 * [ETAPA 6] Analisa o conjunto de dados e gera alertas normativos.
 * @param {object[]} dadosPlanilha Os dados filtrados ou completos.
 * @returns {string[]} Uma lista de alertas em string.
 */
function getAssessoriaNormativa(dadosPlanilha) {
  // Lógica a ser implementada na Etapa 6.
  Logger.log("getAssessoriaNormativa chamada. Implementação pendente.");
  
  // Simula um retorno para testes
  return [
    "Implementação do Assessor Normativo (Etapa 6) pendente.",
    "Exemplo: 3 inspeções com prazo de LTS a vencer.",
    "Exemplo: 1 inspeção com MSG pendente há mais de 5 dias."
  ];
}

/**
 * [ETAPA 6] Sugere um texto de laudo baseado na finalidade e dados.
 * @param {string} finalidade A finalidade da IS.
 * @param {object} dadosInspecionado Os dados da linha do inspecionado.
 * @returns {string} O texto de laudo sugerido.
 */
function getSugestaoDeLaudo(finalidade, dadosInspecionado) {
  // Lógica a ser implementada na Etapa 6.
  Logger.log(`getSugestaoDeLaudo chamada para: ${finalidade}. Implementação pendente.`);
  
  // Simula um retorno para testes
  return `[Implementação da Sugestão de Laudo (Etapa 6) pendente]\n\nBaseado na finalidade '${finalidade}', o laudo padrão seria:\n"Apto / Incapaz para o SAM..."`;
}


// --- FUNÇÕES AUXILIARES (LÓGICA INTERNA) ---

// --- Funções Auxiliares de CRUD (Lógica de Negócio) ---

/**
 * (Privada) Altera o status de 'AGENDADA' para 'Remarcada' e atualiza a data.
 */
function _remarcarIS(payload) {
  const { rowIndex, novaData } = payload;
  if (!rowIndex || !novaData) throw new Error("Dados insuficientes para remarcação.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  if (statusAtual !== 'Agendada') {
    throw new Error(`Ação 'Remarcar' não permitida. Status atual é '${statusAtual}', mas 'Agendada' era esperado.`);
  }

  // Atualiza os valores no array
  values[COL_MAP.DataEntrevista - 1] = new Date(novaData);
  values[COL_MAP.StatusIS - 1] = 'Remarcada';

  // Escreve de volta na planilha
  range.setValues([values]);
  return { success: true, message: "Inspeção remarcada com sucesso." };
}

/**
 * (Privada) Insere o Laudo e muda o status para 'Concluída'.
 */
function _inserirLaudo(payload) {
  const { rowIndex, laudo } = payload;
  if (!rowIndex || !laudo) throw new Error("Dados insuficientes para inserir laudo.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  const laudoAtual = values[COL_MAP.Laudo - 1];

  if (statusAtual !== 'Agendada' && statusAtual !== 'Remarcada') {
    throw new Error(`Ação 'Inserir Laudo' não permitida. Status atual é '${statusAtual}'.`);
  }
  if (laudoAtual) {
    throw new Error("Ação não permitida. Já existe um laudo para esta inspeção.");
  }

  values[COL_MAP.Laudo - 1] = laudo;
  values[COL_MAP.DataLaudo - 1] = new Date(); // Define a data do laudo como 'hoje'
  values[COL_MAP.StatusIS - 1] = 'Concluída';

  range.setValues([values]);
  return { success: true, message: "Laudo inserido e inspeção concluída." };
}

/**
 * (Privada) Insere Restrições (não altera o status).
 */
function _inserirRestricoes(payload) {
  const { rowIndex, restricoes } = payload;
  if (!rowIndex || !restricoes) throw new Error("Dados insuficientes para inserir restrições.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  const finalidade = values[COL_MAP.Finalidade - 1];
  const restricoesAtuais = values[COL_MAP.Restricoes - 1];

  const statusPermitidos = ['Agendada', 'Remarcada', 'Concluída'];
  const finalidadesPermitidas = ['TÉRMINO DE INCAPACIDADE', 'TÉRMINO DE RESTRIÇÕES', 'VERIFICAÇÃO DE DEFICIÊNCIA FUNCIONAL'];

  if (!statusPermitidos.includes(statusAtual)) {
    throw new Error(`Ação 'Inserir Restrições' não permitida para o status '${statusAtual}'.`);
  }
  if (!finalidadesPermitidas.includes(finalidade)) {
     throw new Error(`Ação 'Inserir Restrições' não permitida para a finalidade '${finalidade}'.`);
  }
  if (restricoesAtuais) {
    throw new Error("Ação não permitida. Já existem restrições cadastradas.");
  }

  values[COL_MAP.Restricoes - 1] = restricoes.join(', '); // Assume que 'restricoes' é um array

  range.setValues([values]);
  return { success: true, message: "Restrições inseridas com sucesso." };
}

/**
 * (Privada) Insere o número TIS e muda o status para 'Votada JRS'.
 */
function _inserirTIS(payload) {
  const { rowIndex, tis } = payload;
  if (!rowIndex || !tis) throw new Error("Dados insuficientes para inserir TIS.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  const tisAtual = values[COL_MAP.TIS - 1];
  const statusPermitidos = ['Concluída', 'Restituída Auditoria', 'Restituída JSD'];

  if (!statusPermitidos.includes(statusAtual)) {
    throw new Error(`Ação 'Inserir TIS' não permitida para o status '${statusAtual}'.`);
  }
  if (tisAtual) {
    throw new Error("Ação não permitida. Já existe um TIS cadastrado.");
  }

  values[COL_MAP.TIS - 1] = tis;
  values[COL_MAP.StatusIS - 1] = 'Votada JRS';

  range.setValues([values]);
  return { success: true, message: "TIS inserido. Status alterado para 'Votada JRS'." };
}

/**
 * (Privada) Insere o DS-1a, muda status para 'TIS Assinado' e 'MSG' para 'PENDENTE'.
 */
function _inserirDS1a(payload) {
  const { rowIndex, ds1a } = payload;
  if (!rowIndex || !ds1a) throw new Error("Dados insuficientes para inserir DS-1a.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  const ds1aAtual = values[COL_MAP.DS1a - 1];

  if (statusAtual !== 'Votada JRS') {
    throw new Error(`Ação 'Inserir DS-1a' não permitida. Status atual é '${statusAtual}'.`);
  }
  if (ds1aAtual) {
    throw new Error("Ação não permitida. Já existe um DS-1a cadastrado.");
  }

  values[COL_MAP.DS1a - 1] = ds1a;
  values[COL_MAP.StatusIS - 1] = 'TIS Assinado';
  values[COL_MAP.MSG - 1] = 'PENDENTE'; // Regra automática

  range.setValues([values]);
  return { success: true, message: "DS-1a inserido. Status 'TIS Assinado' e MSG 'PENDENTE'." };
}

/**
 * (Privada) Registra a MSG como 'ENVIADA'.
 */
function _registrarMSG(payload) {
  const { rowIndex } = payload;
  if (!rowIndex) throw new Error("RowIndex não fornecido.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  const statusAtual = values[COL_MAP.StatusIS - 1];
  if (statusAtual !== 'TIS Assinado') {
    throw new Error(`Ação 'Registrar MSG' não permitida. Status atual é '${statusAtual}'.`);
  }

  values[COL_MAP.MSG - 1] = 'ENVIADA';

  range.setValues([values]);
  return { success: true, message: "Mensagem registrada como 'ENVIADA'." };
}

/**
 * (Privada) Função genérica para mudar status (Auditoria, JSD, etc.)
 */
function _mudarStatus(rowIndex, statusEsperado, novoStatus) {
  if (!rowIndex) throw new Error("RowIndex não fornecido.");

  const sheet = SS.getSheetByName(SHEET_DATA);
  const range = sheet.getRange(rowIndex, COL_MAP.StatusIS, 1, 1);
  const statusAtual = range.getValue();

  // Chain-of-Thought: Verificando a regra de negócio (Etapa 5)
  if (statusAtual !== statusEsperado) {
    throw new Error(`Ação não permitida. Status atual é '${statusAtual}', mas '${statusEsperado}' era esperado.`);
  }

  range.setValue(novoStatus);
  return { success: true, message: `Status alterado para '${novoStatus}'.` };
}


// --- Funções Auxiliares de Leitura e Processamento ---

/**
 * (Privada) Busca os cabeçalhos da aba 'ListaControle'.
 */
function _getHeaders() {
  const sheet = SS.getSheetByName(SHEET_DATA);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * (Privada) Converte dados da planilha (array de arrays) em um array de objetos.
 * Adiciona o 'rowIndex' para referência futura no CRUD.
 */
function _convertDataToObjects(dataValues, headers) {
  return dataValues.map((row, index) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    // Adiciona 2: 1 pela base-1 das planilhas, 1 pelo cabeçalho pulado
    obj.rowIndex = index + 2; 
    
    // Converte datas importantes
    obj.DataEntrevista = _parseDate(obj.DataEntrevista);
    obj.DataLaudo = _parseDate(obj.DataLaudo);
    
    return obj;
  });
}

/**
 * (Privada) Aplica os filtros recebidos do cliente ao conjunto de dados.
 */
function _applyFilters(dataObjects, filters) {
  const { dataInicio, dataFim, finalidade, om, statusIS } = filters;
  
  const startDate = dataInicio ? new Date(dataInicio) : null;
  const endDate = dataFim ? new Date(dataFim) : null;

  // Ajusta as datas para cobrir o dia inteiro
  if (startDate) startDate.setHours(0, 0, 0, 0);
  if (endDate) endDate.setHours(23, 59, 59, 999);
  
  return dataObjects.filter(row => {
    // Filtro de Data
    if (startDate && (!row.DataEntrevista || row.DataEntrevista < startDate)) return false;
    if (endDate && (!row.DataEntrevista || row.DataEntrevista > endDate)) return false;
    
    // Filtros de String
    if (finalidade && row.Finalidade !== finalidade) return false;
    if (om && row.OM !== om) return false;
    if (statusIS && row.StatusIS !== statusIS) return false;

    return true;
  });
}

/**
 * (Privada) Calcula todos os KPIs, incluindo o comparativo MoM.
 */
function _calculateKpis(allDataObjects, filters) {
  // 1. Calcula para o período ATUAL (usando os dados já filtrados)
  const currentFilteredData = _applyFilters(allDataObjects, filters);
  const kpiCurrent = _calculateKpiMetrics(currentFilteredData);

  // 2. Calcula para o período ANTERIOR
  const { dataInicio, dataFim } = filters;
  let kpiPrevious = { laudoAssinado: 0, faltas: 0, canceladas: 0 }; // Default

  if (dataInicio && dataFim) {
    try {
      const startDate = new Date(dataInicio);
      const endDate = new Date(dataFim);
      
      // Calcula a duração do período
      const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 para incluir o dia
      
      // Calcula o período anterior
      const prevEndDate = new Date(startDate.getTime() - (1000 * 60 * 60 * 24)); // 1 dia antes do início
      const prevStartDate = new Date(prevEndDate.getTime() - (diffDays - 1) * (1000 * 60 * 60 * 24));

      // Monta os filtros para o período anterior (mantendo os outros filtros)
      const previousFilters = {
        ...filters,
        dataInicio: prevStartDate.toISOString(),
        dataFim: prevEndDate.toISOString()
      };
      
      const previousFilteredData = _applyFilters(allDataObjects, previousFilters);
      kpiPrevious = _calculateKpiMetrics(previousFilteredData);

    } catch (e) {
      Logger.log("Não foi possível calcular o período anterior (datas inválidas?): " + e.message);
      // Continua sem os dados anteriores
    }
  }

  // 3. Monta o objeto final de KPIs
  return {
    agendadas: kpiCurrent.agendadas,
    auditoria: kpiCurrent.auditoria,
    msgPendente: kpiCurrent.msgPendente,
    laudoAssinado: {
      total: kpiCurrent.laudoAssinado,
      variacao: _calculateVariation(kpiCurrent.laudoAssinado, kpiPrevious.laudoAssinado)
    },
    faltas: {
      total: kpiCurrent.faltas,
      variacao: _calculateVariation(kpiCurrent.faltas, kpiPrevious.faltas)
    },
    canceladas: {
      total: kpiCurrent.canceladas,
      variacao: _calculateVariation(kpiCurrent.canceladas, kpiPrevious.canceladas)
    }
  };
}

/**
 * (Privada) Núcleo de cálculo de métricas para um conjunto de dados.
 */
function _calculateKpiMetrics(data) {
  let agendadas = 0;
  let auditoria = 0;
  let msgPendente = 0;
  let laudoAssinado = 0;
  let faltas = 0;
  let canceladas = 0;

  for (const row of data) {
    // KPI Agendadas
    if (row.StatusIS === 'Agendada' || row.StatusIS === 'Remarcada') {
      agendadas++;
    }
    // KPI Auditoria
    if (row.StatusIS === 'Auditoria' || row.StatusIS === 'JSD') {
      auditoria++;
    }
    // KPI MSG Pendente
    if (row.StatusIS === 'TIS Assinado' && row.MSG === 'PENDENTE') {
      msgPendente++;
    }
    // KPI Laudo Assinado (Qualquer status que tenha laudo)
    if (row.DataLaudo) {
      laudoAssinado++;
    }
    // KPI Faltas
    if (row.StatusIS === 'Faltou') {
      faltas++;
    }
    // KPI Canceladas
    if (row.StatusIS === 'Cancelada') {
      canceladas++;
    }
  }

  return { agendadas, auditoria, msgPendente, laudoAssinado, faltas, canceladas };
}

/**
 * (Privada) Calcula a variação percentual entre dois números.
 */
function _calculateVariation(current, previous) {
  if (previous === 0) {
    // Se o anterior era 0, qualquer aumento é "infinito"
    // Retornamos 100% se o atual for > 0, ou 0 se for 0.
    return (current > 0 ? 100.0 : 0.0);
  }
  const variacao = ((current - previous) / previous) * 100;
  return parseFloat(variacao.toFixed(1)); // Arredonda para 1 casa decimal
}

/**
 * (Privada) Prepara os dados agregados para os gráficos.
 */
function _prepareChartData(filteredData) {
  // 1. Gráfico de Tendência Mensal
  const tendenciaMap = _groupAndCount(filteredData, 'DataEntrevista', (date) => {
    if (!date) return 'Sem Data';
    // Formata como 'AAAA-MM' para agrupar
    return date.toISOString().substring(0, 7); 
  });
  const tendenciaMensal = Object.entries(tendenciaMap).map(([mes, total]) => [mes, total]);
  tendenciaMensal.sort((a, b) => a[0].localeCompare(b[0])); // Ordena por mês

  // 2. Gráfico de Finalidades
  const finalidadeMap = _groupAndCount(filteredData, 'Finalidade');
  const porFinalidade = Object.entries(finalidadeMap).map(([nome, total]) => [nome, total]);

  // 3. Gráfico de Status
  const statusMap = _groupAndCount(filteredData, 'StatusIS');
  const porStatus = Object.entries(statusMap).map(([nome, total]) => [nome, total]);

  // 4. Gráfico por OM
  const omMap = _groupAndCount(filteredData, 'OM');
  const porOM = Object.entries(omMap).map(([nome, total]) => [nome, total]);

  return { tendenciaMensal, porFinalidade, porStatus, porOM };
}

/**
 * (Privada) Helper genérico para agrupar e contar dados.
 */
function _groupAndCount(data, key, transformFn = (val) => val) {
  return data.reduce((acc, row) => {
    const groupKey = transformFn(row[key]);
    if (groupKey) { // Ignora nulos ou vazios
      acc[groupKey] = (acc[groupKey] || 0) + 1;
    }
    return acc;
  }, {});
}

/**
 * (Privada) Converte string de data (dd/MM/AAAA) ou objeto Date para um objeto Date válido.
 * Retorna null se a data for inválida ou vazia.
 */
function _parseDate(dateInput) {
  if (!dateInput) return null;
  
  // Se já for um objeto Date (do .setValue())
  if (dateInput instanceof Date) {
    return dateInput;
  }
  
  // Se for string (do .getValues())
  if (typeof dateInput === 'string') {
    const parts = dateInput.split(' ')[0].split('/'); // Remove horas se houver e divide
    if (parts.length === 3) {
      // Formato dd/MM/AAAA
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Mês é base 0
      const year = parseInt(parts[2], 10);
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        return new Date(year, month, day);
      }
    }
  }
  
  // Se for número (timestamp da planilha)
  if (typeof dateInput === 'number') {
    return new Date(dateInput);
  }

  return null; // Inválido
}

/**
 * (Privada) Lê uma coluna inteira e retorna valores únicos, filtrados e ordenados.
 */
function _getUniqueColumnValues(sheet, columnIndex) {
  const range = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1);
  return range.getValues()
    .flat()
    .filter(String) // Remove valores vazios
    .filter((value, index, self) => self.indexOf(value) === index) // Pega únicos
    .sort(); // Ordena
}
