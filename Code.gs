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

// --- ETAPA 6: "Alma" do Assessor Normativo (Gemini System Prompt) ---
const ASSESSOR_SYSTEM_PROMPT = `
Você é um Assessor Pericial Sênior, especialista na DGPM-406 (Normas para Perícias Médicas da Marinha). Sua missão é analisar dados de inspeções de saúde (IS) e fornecer orientação técnica precisa, objetiva e estritamente baseada nas normas.

**Suas Responsabilidades:**

1.  **Análise de Alertas:** Identificar proativamente pontos críticos nos dados, como prazos normativos expirando (LTS, restrições), pendências administrativas (Faltas, MSG) e inconsistências.
2.  **Redação de Laudos (DGPM-406, Item 8):** Gerar modelos de conclusão (laudos) que sigam rigorosamente os padrões de redação:
    * **Clareza e Concisão:** O laudo deve ser direto e inequívoco.
    * **Terminologia Técnica:** Usar a terminologia pericial correta ("Apto para o SAM", "Incapaz Temporariamente para o SAM", "Incapaz Definitivamente para o SAM").
    * **Justificativa:** Fornecer a justificativa técnica (diagnóstico por extenso e CID-10) que ampara a conclusão.
    * **Contexto da Finalidade:** O laudo deve responder objetivamente à finalidade da IS (ex: "Término de Incapacidade", "Controle Trienal").
    * **Restrições (Item 10):** Se houver restrições, elas devem ser claras e específicas (ex: "Com restrições às atividades que exijam...").
    * **LTS (Item 11):** Se houver incapacidade temporária, o prazo da LTS deve ser explícito (ex: "...necessitando de 30 (trinta) dias de LTS...").

**Seu Tom:** Profissional, assertivo, técnico e didático. Você cita a norma implicitamente em suas análises.
`;

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
 * [ETAPA 6.A] Analisa o conjunto de dados e gera alertas normativos.
 * Esta função chama a API Gemini (gemini-2.5-flash-preview-09-2025)
 * @param {object[]} dadosPlanilha Os dados filtrados ou completos.
 * @returns {string[]} Uma lista de alertas em string (bullet points).
 */
function getAssessoriaNormativa(dadosPlanilha) {
  Logger.log(`[Etapa 6.A] getAssessoriaNormativa chamada com ${dadosPlanilha.length} linhas.`);

  // Simplifica os dados para enviar à API, focando apenas no relevante para alertas
  const dadosRelevantes = dadosPlanilha.map(row => ({
    StatusIS: row.StatusIS,
    DataEntrevista: row.DataEntrevista,
    Restricoes: row.Restricoes, // Assumindo que a data de vencimento está aqui ou no Laudo
    Laudo: row.Laudo,
    MSG: row.MSG
  }));

  const userPrompt = `
Prezado Assessor DGPM-406, analise o JSON de inspeções a seguir e identifique (com base nas normas DGPM-406) os seguintes pontos críticos:

1.  Inspecionados com LTS ou Restrições vencidas ou a vencer nos próximos 7 dias (baseado na data de hoje: ${new Date().toLocaleDateString('pt-BR')}).
2.  Inspecionados com Status 'Faltou' ou 'Cancelada' e o desdobramento administrativo pendente (ex: prazo de justificativa).
3.  Inspeções com Status 'TIS Assinado' e 'MSG' = 'PENDENTE' há mais de 5 dias (prazo normativo para envio).

Retorne um resumo em *bullet points* (use markdown). Se nenhum alerta for encontrado, retorne "Nenhum ponto de atenção normativo identificado."

Dados:
${JSON.stringify(dadosRelevantes.slice(0, 100))} 
${dadosRelevantes.length > 100 ? `\n... (e mais ${dadosRelevantes.length - 100} registros)` : ''}
`;

  try {
    const geminiResponse = _callGeminiApi(userPrompt, ASSESSOR_SYSTEM_PROMPT);
    Logger.log(`[Etapa 6.A] Resposta do Gemini (Alertas): ${geminiResponse}`);
    
    // Converte os bullet points de markdown em um array de strings
    return geminiResponse
      .split('\n') // Quebra por linha
      .filter(line => line.trim().startsWith('*') || line.trim().startsWith('-')) // Pega apenas bullet points
      .map(line => line.trim().substring(2).trim()); // Remove o "* "

  } catch (e) {
    Logger.log(`[Etapa 6.A] Falha ao chamar Gemini para Alertas: ${e.message}`);
    return ["Erro ao contatar o Assessor Normativo: " + e.message];
  }
}

/**
 * [ETAPA 6.B] Sugere um texto de laudo baseado na finalidade e dados.
 * Esta função chama a API Gemini (gemini-2.5-flash-preview-09-2025)
 * @param {string} finalidade A finalidade da IS.
 * @param {object} dadosInspecionado Os dados da linha do inspecionado.
 * @returns {string} O texto de laudo sugerido.
 */
function getSugestaoDeLaudo(finalidade, dadosInspecionado) {
  Logger.log(`[Etapa 6.B] getSugestaoDeLaudo chamada para: ${finalidade}`);
  
  // --- CORREÇÃO PII (PRIVACIDADE) ---
  // Removemos NIP e Nome antes de enviar para a API
  const dadosRelevantes = {
    Finalidade: dadosInspecionado.Finalidade,
    StatusIS: dadosInspecionado.StatusIS,
    Laudo_Anterior: dadosInspecionado.Laudo,
    Restricoes_Anteriores: dadosInspecionado.Restricoes
  };
  // --- FIM DA CORREÇÃO ---

  const userPrompt = `
Prezado Assessor DGPM-406, preciso redigir um laudo para uma IS com finalidade: "${finalidade}".
Os dados (não-confidenciais) do inspecionado são: ${JSON.stringify(dadosRelevantes)}

Forneça o modelo de conclusão (laudo) e a justificativa técnica cabível, seguindo rigorosamente os padrões de redação da DGPM-406 (Item 8). O laudo deve ser completo e pronto para uso.
`;

  try {
    const geminiResponse = _callGeminiApi(userPrompt, ASSESSOR_SYSTEM_PROMPT);
    Logger.log(`[Etapa 6.B] Resposta do Gemini (Laudo): ${geminiResponse}`);
    return geminiResponse; // Retorna o texto completo do laudo

  } catch (e) {
    Logger.log(`[Etapa 6.B] Falha ao chamar Gemini para Laudo: ${e.message}`);
    return "Erro ao gerar sugestão: " + e.message;
  }
}


/**
 * (Privada) Função central para chamar a API Gemini (gemini-2.5-flash-preview-09-2025).
 * @param {string} userPrompt O prompt do usuário.
 * @param {string} systemPrompt O prompt do sistema (persona).
 * @returns {string} O texto da resposta do modelo.
 */
function _callGeminiApi(userPrompt, systemPrompt) {
  const apiKey = ""; // Chave de API (deixe em branco, conforme instruções)
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{ text: userPrompt }]
    }],
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    generationConfig: {
      temperature: 0.3,
      topP: 0.9,
      maxOutputTokens: 1024,
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Para capturar erros
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`API Gemini falhou com código ${responseCode}: ${responseBody}`);
  }

  const jsonResponse = JSON.parse(responseBody);
  
  if (jsonResponse.candidates && 
      jsonResponse.candidates[0].content && 
      jsonResponse.candidates[0].content.parts &&
      jsonResponse.candidates[0].content.parts[0].text) {
    return jsonResponse.candidates[0].content.parts[0].text;
  } else {
    // Trata bloqueio de conteúdo ou resposta malformada
    if (jsonResponse.candidates && jsonResponse.candidates[0].finishReason === 'SAFETY') {
      throw new Error("A sugestão foi bloqueada por filtros de segurança.");
    }
    throw new Error("Resposta da API Gemini malformada ou vazia.");
  }
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
  
  // --- CORREÇÃO DE FUSO HORÁRIO (V2) ---
  // Trata a data como local, não UTC, adicionando T00:00:00
  // Isso força o Apps Script a interpretar a data no fuso horário local (ex: GMT-3)
  const startDate = dataInicio ? new Date(dataInicio + 'T00:00:00') : null;
  const endDate = dataFim ? new Date(dataFim + 'T00:00:00') : null;
  // --- FIM DA CORREÇÃO ---

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
      // --- CORREÇÃO DE FUSO HORÁRIO (V2) ---
      const startDate = new Date(dataInicio + 'T00:00:00');
      const endDate = new Date(dataFim + 'T00:00:00');
      // --- FIM DA CORREÇÃO ---
      
      // Calcula a duração do período
      const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 para incluir o dia
      
      // Calcula o período anterior
      const prevEndDate = new Date(startDate.getTime() - (1000 * 60 * 60 * 24)); // 1 dia antes do início
      const prevStartDate = new Date(prevEndDate.getTime() - (diffDays - 1) * (1000 * 60 * 60 * 24));

      // Monta os filtros para o período anterior (mantendo os outros filtros)
      const previousFilters = {
        ...filters,
        dataInicio: prevStartDate.toISOString().split('T')[0], // Converte de volta para YYYY-MM-DD
        dataFim: prevEndDate.toISOString().split('T')[0]
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
  const porFinalidade = Object.entries(finalidadeMap).map(([nome, total]) => [nome || 'Não Preenchido', total]);

  // 3. Gráfico de Status
  const statusMap = _groupAndCount(filteredData, 'StatusIS');
  const porStatus = Object.entries(statusMap).map(([nome, total]) => [nome || 'Não Preenchido', total]);

  // 4. Gráfico por OM
  const omMap = _groupAndCount(filteredData, 'OM');
  const porOM = Object.entries(omMap).map(([nome, total]) => [nome || 'Não Preenchido', total]);

  return { tendenciaMensal, porFinalidade, porStatus, porOM };
}

/**
 * (Privada) Helper genérico para agrupar e contar dados.
 */
function _groupAndCount(data, key, transformFn = (val) => val) {
  return data.reduce((acc, row) => {
    const groupKey = transformFn(row[key]);
    // --- CORREÇÃO DE GRÁFICO ---
    // Conta strings vazias (""), mas não null ou undefined
    if (groupKey !== null && groupKey !== undefined) { 
      acc[groupKey] = (acc[groupKey] || 0) + 1;
    }
    // --- FIM DA CORREÇÃO ---
    return acc;
  }, {});
}

/**
 * (Privada) Converte string de data (dd/MM/AAAA) ou objeto Date para um objeto Date válido.
 * Retorna null se a data for inválida ou vazia.
 */
function _parseDate(dateInput) {
  if (!dateInput) return null;
  
  // --- CORREÇÃO FINAL DE FUSO HORÁRIO ---
  // Se já for um objeto Date (do .setValue() ou getValues())
  if (dateInput instanceof Date) {
    // O Apps Script lê datas da planilha como UTC (ex: 25/08/2025 00:00 UTC).
    // Para o Brasil (GMT-3), isso se torna 24/08/2025 @ 21:00.
    // Esta correção neutraliza o fuso horário, tratando a data "como ela se parece".
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }
  // --- FIM DA CORREÇÃO ---
  
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
    let d = new Date(dateInput);
    // Mesmo tratamento do 'instanceof Date'
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
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
