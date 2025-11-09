// ================================================================= //
//                      CONFIGURAÇÕES GLOBAIS                        //
// ================================================================= //

// --- Configs do Formulário de Inspeção e Dashboard (ORIGINAL) ---
const ss = SpreadsheetApp.getActiveSpreadsheet();
const listasRefSheet = ss.getSheetByName('ListasRef');
const listaControleSheet = ss.getSheetByName('ListaControle');
const militaresHnreSheet = ss.getSheetByName('MilitaresHNRe');
const concursosSheet = ss.getSheetByName('ListaConcursos');
// --- Configs do Gerador de Pareceres (ORIGINAL) ---
const TEMPLATE_IDS = {
  'Psiquiatria': '1dBBPpAUigm9DFUWqyP0NkTEVGxeHWwKZF8dnoJ7xqP8',
  'Psicologia': '1eQiRCtmF-V1Etn7fp6AbhTGZbYkqJlLrQYGjLRfePqQ',
  'Hepatologia': '12-s0h077A6hwipVe4oARcor4cNMKeAGOl6elV-pk_HQ',
  'Cardiologia': '1woGqRVEpe7hQkqZiPrPn2bUgIgcN7MtImn0plq8fbLE',
  'Ortopedia': '1qmi93Y_kZpNZEeYcqhSA41KVsZ4gFE1MqkK6kbAsheI'
};
const SHEET_ID_PARECER_MILITARES = "1yqZYEh2apMezknd0_0sBBEbliV3jwoDEw43FipXwsUQ";
const PDF_FOLDER_ID = "176YQTOOWXuE78Xul49XYpJsVNyu-eZFi";

const PERITO_EMAILS = {
  'CT Mauriston': 'mauriston.martins@marinha.mil.br',
  'CT Júlio César': 'julio.xaiver@marinha.mil.br',
  '1T Salyne': 'salyne.martins@marinha.mil.br',
  '1T Luz': 'lucas.luz@marinha.mil.br',
  '2T Trindade': 'marcelo.trindade@marinha.mil.br'
};

// ================================================================= //
//          ROTEADOR PRINCIPAL E FUNÇÕES DE SERVIÇO WEB              //
//          (SEU CÓDIGO ORIGINAL - INTOCADO)                         //
// ================================================================= //

function doGet(e) {
  let page = e.parameter.page;
  if (page == 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard').evaluate().setTitle('Dashboard JRS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page == 'parecer') {
    return HtmlService.createTemplateFromFile('Parecer').evaluate().setTitle('Gerador de Pareceres').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Página padrão é 'formulario'
  return HtmlService.createTemplateFromFile('Formulario').evaluate().setTitle('Formulário JRS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ================================================================= //
//     FUNÇÕES DO FORMULÁRIO DE INSPEÇÃO E DASHBOARD (ORIGINAL)      //
// ================================================================= //

function getDropdownData() {
  try {
    const omOptions = listasRefSheet.getRange('B2:B' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    const pgOptions = listasRefSheet.getRange('C2:C' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    const finalidadeOptions = listasRefSheet.getRange('A2:A' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    const statusOptions = listasRefSheet.getRange('F2:F' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    const restricoesOptions = listasRefSheet.getRange('G2:G' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    return { omOptions, pgOptions, finalidadeOptions, statusOptions, restricoesOptions };
  } catch (e) {
    console.error("Erro em getDropdownData: " + e.message);
    throw new Error("Não foi possível buscar os dados para os menus.");
  }
}

// FUNÇÃO RENOMEADA
function getHnreMilitaryDataForForm() {
  try {
    const data = militaresHnreSheet.getRange('A2:C' + militaresHnreSheet.getLastRow()).getValues();
    return data.map(row => ({ inspecionado: row[2], pg: row[0], nip: row[1] })).filter(m => m.inspecionado);
  } catch (e) {
    console.error("Erro em getHnreMilitaryDataForForm: " + e.message);
    throw new Error("Não foi possível buscar os dados dos militares.");
  }
}

function addNewInspection(formData) {
  try {
    const headers = listaControleSheet.getRange(1, 1, 1, listaControleSheet.getLastColumn()).getValues()[0];
    let restricoesString = formData.restricoes.join(', ');
    if (formData.restricoes.includes('Outros') && formData.outrosRestricao) {
      restricoesString = restricoesString.replace('Outros', `Outros: ${formData.outrosRestricao}`);
    }

    const rowData = {
      'OM': formData.om,
      'Inspecionado': formData.inspecionado,
      'P/G/Q': formData.pg,
      'NIP': formData.nip,
      'IS': formData.is,
      'Finalidade': formData.finalidade,
      'DataEntrevista': formData.dataEntrevista,
      'StatusIS': formData.statusIS,
      'TIS': formData.tis,
      'DS-1a': formData.ds1a,
      'Laudo': formData.laudo,
      'DataLaudo': formData.dataLaudo,
      'Restrições': restricoesString
  

    };

    const newRow = headers.map(header => rowData[header] || null);
    listaControleSheet.appendRow(newRow);
    return { status: 'success', message: 'Inspeção adicionada com sucesso!' };
  } catch(e) {
    console.error("Erro em addNewInspection: " + e.message);
    return { status: 'error', message: `Falha ao adicionar: ${e.message}` };
  }
}

// --- FUNÇÃO ATUALIZADA ---
function getDashboardData() {
  try {
    const rotinaData = listaControleSheet.getRange('A2:N' + listaControleSheet.getLastRow()).getValues();
    const concursoData = concursosSheet.getRange('A2:I' + concursosSheet.getLastRow()).getValues();

    const dashboardData = rotinaData.map(r => ({
      EventDate: r[1] ? new Date(r[1]).toLocaleDateString('pt-BR') : '', StatusIS: r[7], Finalidade: r[2], MSG: r[13], type: 'Rotina'
    }));
    concursoData.forEach(r => {
      dashboardData.push({
        EventDate: r[0] ? new Date(r[0]).toLocaleDateString('pt-BR') : '', StatusIS: r[8], Finalidade: r[6], MSG: '', type: 'Concurso'
      });
    });
    const tableData = rotinaData.map(r => ({
      DataEntrevista: r[1] ? new Date(r[1]).toLocaleDateString('pt-BR') : '', IS: r[0], Finalidade: r[2], StatusIS: r[7], Inspecionado: r[6], OM: r[3], 'P/G/Q': r[4], NIP: r[5], Laudo: r[9], DataLaudo: r[8] ? new Date(r[8]).toLocaleDateString('pt-BR') : '', Restrições: r[10], TIS: r[11], 'DS-1a': r[12], MSG: r[13],
    }));
    // --- ADICIONADO ---
    // Busca a lista de restrições para o novo modal de edição
    const restricoesOptions = listasRefSheet.getRange('G2:G' + listasRefSheet.getLastRow()).getValues().flat().filter(String);
    return { dashboardData, tableData, restricoesOptions }; // Retorna os dados E a lista de restrições
  } catch(e) {
    console.error("Erro em getDashboardData: " + e.message);
    throw new Error("Falha ao carregar dados do Dashboard.");
  }
}

// --- FUNÇÃO DE UPDATE DE MSG ---
function updateMsgStatus(isNumber) {
  try {
    if (!listaControleSheet) { throw new Error("Planilha 'ListaControle' não encontrada.");
    }
    const headers = listaControleSheet.getRange(1, 1, 1, listaControleSheet.getLastColumn()).getValues()[0];
    const isColIndex = headers.indexOf('IS');
    const msgColIndex = headers.indexOf('MSG');
    if (isColIndex === -1 || msgColIndex === -1) { throw new Error("Colunas 'IS' ou 'MSG' não encontradas.");
    }
    const dataRange = listaControleSheet.getRange(2, 1, listaControleSheet.getLastRow() - 1, listaControleSheet.getLastColumn());
    const values = dataRange.getValues();
    let rowFound = false;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][isColIndex]) == String(isNumber)) {
        listaControleSheet.getRange(i + 2, msgColIndex + 1).setValue('ENVIADA');
        rowFound = true;
        Logger.log(`IS ${isNumber} encontrada na linha ${i+2}. MSG atualizada para ENVIADA.`);
        break;
      }
    }
    if (rowFound) { return { success: true, message: 'Status atualizado para ENVIADA.'
    }; } 
    else { Logger.log(`IS ${isNumber} não encontrada.`);
      return { success: false, message: 'Número da IS não encontrado.' };
    }
  } catch (e) {
    Logger.log('Erro em updateMsgStatus: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// --- FUNÇÃO DE UPDATE DE CONCLUSÃO (ATUALIZADA) ---
function updateInspectionConclusion(formData) {
  try {
    if (!listaControleSheet) { throw new Error("Planilha 'ListaControle' não encontrada.");
    }

    // 1. Encontrar os índices das colunas
    const headers = listaControleSheet.getRange(1, 1, 1, listaControleSheet.getLastColumn()).getValues()[0];
    const colIndices = {
      is: headers.indexOf('IS'),
      laudo: headers.indexOf('Laudo'),
      dataLaudo: headers.indexOf('DataLaudo'),
      tis: headers.indexOf('TIS'),
      ds1a: headers.indexOf('DS-1a'),
      restricoes: headers.indexOf('Restrições'),
      statusIS: headers.indexOf('StatusIS') // <-- ADICIONADO
    };
    // <-- CÉLULA DE ERRO ATUALIZADA -->
    if (colIndices.is === -1 || colIndices.laudo === -1 || colIndices.dataLaudo === -1 || colIndices.tis === -1 || colIndices.ds1a === -1 || colIndices.restricoes === -1 || colIndices.statusIS === -1) {
      throw new Error("Colunas essenciais (IS, Laudo, DataLaudo, TIS, DS-1a, Restrições, StatusIS) não encontradas. Verifique a 1ª linha da planilha.");
    }

    // 2. Encontrar a linha da inspeção
    const dataRange = listaControleSheet.getRange(2, colIndices.is + 1, listaControleSheet.getLastRow() - 1, 1);
    const isValues = dataRange.getValues();
    let targetRowIndex = -1; // Esta será a linha no array (0-based)
    
    for (let i = 0; i < isValues.length; i++) {
      if (String(isValues[i][0]) == String(formData.isNumber)) {
        targetRowIndex = i;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`IS ${formData.isNumber} não encontrada.`);
    }

    // 3. Formatar Restrições
    let restricoesString = formData.restricoes.join(', ');
    if (formData.restricoes.includes('Outros') && formData.outrosRestricao) {
      restricoesString = restricoesString.replace('Outros', `Outros: ${formData.outrosRestricao}`);
    }

    // 4. Preparar dados e atualizar a planilha
    const targetRowOnSheet = targetRowIndex + 2;
    // +2 porque o array é 0-based e a planilha é 1-based + cabeçalho
    
    // Atualiza os valores dos campos
    listaControleSheet.getRange(targetRowOnSheet, colIndices.laudo + 1).setValue(formData.laudo || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.dataLaudo + 1).setValue(formData.dataLaudo || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.tis + 1).setValue(formData.tis || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.ds1a + 1).setValue(formData.ds1a || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.restricoes + 1).setValue(restricoesString || null);

    // --- LÓGICA DE STATUS (NOVA) ---
    const { laudo, dataLaudo, tis, ds1a } = formData;
    let newStatus = '';

    if (laudo && dataLaudo && tis && ds1a) {
      newStatus = 'TIS assinado';
    } else if (laudo && dataLaudo && tis && !ds1a) {
      newStatus = 'Votada JRS';
    } else if (laudo && dataLaudo && !tis && !ds1a) {
      newStatus = 'Concluída';
    }
    // --- FIM DA LÓGICA DE STATUS ---

    // 5. Preparar dados de retorno
    const updatedData = { 
      Laudo: formData.laudo ||
      '',
      DataLaudo: formData.dataLaudo ?
        new Date(formData.dataLaudo).toLocaleDateString('pt-BR') : '', // Re-formata para pt-BR
      TIS: formData.tis ||
      '',
      'DS-1a': formData.ds1a || '',
      Restrições: restricoesString ||
      ''
    };

    // 6. Atualizar o status se a lógica definiu um novo
    if (newStatus) {
      listaControleSheet.getRange(targetRowOnSheet, colIndices.statusIS + 1).setValue(newStatus);
      updatedData.StatusIS = newStatus; // Adiciona o status ao objeto de retorno
      Logger.log(`IS ${formData.isNumber} (Linha ${targetRowOnSheet}) atualizada. Novo Status: ${newStatus}`);
    } else {
      Logger.log(`IS ${formData.isNumber} (Linha ${targetRowOnSheet}) atualizada. Status não alterado.`);
    }

    // 7. Retornar os dados formatados para o frontend
    return { 
      success: true, 
      message: 'Conclusão salva com sucesso!',
      updatedData: updatedData
    };
  } catch (e) {
    Logger.log('Erro em updateInspectionConclusion: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// --- NOVA FUNÇÃO DE REMARCAÇÃO ---
function remarcarInspecao(formData) {
  try {
    if (!listaControleSheet) { throw new Error("Planilha 'ListaControle' não encontrada.");
    }

    const isNumber = formData.isNumber;
    const novaData = formData.novaData;
    // Formato YYYY-MM-DD

    if (!isNumber || !novaData) {
      throw new Error("Número da IS ou a Nova Data não foram fornecidos.");
    }

    // 1. Encontrar os índices das colunas
    const headers = listaControleSheet.getRange(1, 1, 1, listaControleSheet.getLastColumn()).getValues()[0];
    const colIndices = {
      is: headers.indexOf('IS'),
      dataEntrevista: headers.indexOf('DataEntrevista'),
      statusIS: headers.indexOf('StatusIS')
    };
    if (colIndices.is === -1 || colIndices.dataEntrevista === -1 || colIndices.statusIS === -1) {
      throw new Error("Colunas essenciais (IS, DataEntrevista, StatusIS) não encontradas. Verifique a 1ª linha da planilha.");
    }

    // 2. Encontrar a linha da inspeção
    const dataRange = listaControleSheet.getRange(2, colIndices.is + 1, listaControleSheet.getLastRow() - 1, 1);
    const isValues = dataRange.getValues();
    let targetRowIndex = -1; // 0-based
    
    for (let i = 0; i < isValues.length; i++) {
      if (String(isValues[i][0]) == String(isNumber)) {
        targetRowIndex = i;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`IS ${isNumber} não encontrada.`);
    }

    // 3. Atualizar a planilha
    const targetRowOnSheet = targetRowIndex + 2;
    // 1-based + cabeçalho
    
    // Converte a data YYYY-MM-DD para um objeto Date do GAS (respeitando o fuso)
    const [year, month, day] = novaData.split('-');
    const dataObj = new Date(year, month - 1, day);

    listaControleSheet.getRange(targetRowOnSheet, colIndices.dataEntrevista + 1).setValue(dataObj);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.statusIS + 1).setValue('Remarcada');
    Logger.log(`IS ${isNumber} (Linha ${targetRowOnSheet}) remarcada para ${novaData}.`);
    
    // 4. Retornar os dados formatados para o frontend
    return { 
      success: true, 
      message: 'Inspeção remarcada com sucesso!',
      updatedData: {
        DataEntrevista: dataObj.toLocaleDateString('pt-BR'), // Formata para dd/mm/yyyy
        StatusIS: 'Remarcada'
      } 
    };
  } catch (e) {
    Logger.log('Erro em remarcarInspecao: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}


// ================================================================= //
//           FUNÇÕES DO GERADOR DE PARECERES (ORIGINAL)              //
// ================================================================= //

// FUNÇÃO RENOMEADA
function getHnreMilitaryDataForParecer() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID_PARECER_MILITARES).getSheetByName("MILITAR");
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const militaryList = data.filter(row => row[2] === "HNRE" && row[1]).map(row => ({ nip: row[0], name: row[1], posto: row[3] })).sort((a, b) => a.name.localeCompare(b.name));
    return militaryList;
  } catch (e) {
    Logger.log('Erro em getHnreMilitaryDataForParecer: ' + e.toString());
    return { error: e.toString() };
  }
}

function processForm(formData) {
  try {
    const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const templateId = TEMPLATE_IDS[formData.ESPECIALIDADE];
    const templateDoc = DriveApp.getFileById(templateId);
    const newDocName = `Parecer - ${formData.NOME} - ${new Date().toLocaleDateString('pt-BR')}`;
    const newDocFile = templateDoc.makeCopy(newDocName, pdfFolder);
    const doc = DocumentApp.openById(newDocFile.getId());
    const body = doc.getBody();
    let peritoFullName = getPeritoFullName(formData.PERITO_SELECAO);
    let posto = getPeritoPosto(formData.PERITO_SELECAO);
    let membroValue = '';
    switch (formData.PERITO_SELECAO) { case 'CT Mauriston': membroValue = 'Presidente'; break; case 'CT Júlio César': membroValue = 'Médico Perito Isolado'; break;
    case '1T Salyne': case '1T Luz': membroValue = 'Membro da Junta Regular de Saúde'; break;
    case '2T Trindade': membroValue = 'Médico Perito Isolado'; break; }
    const hoje = new Date();
    const dia = hoje.getDate(); const meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
    const nomeMes = meses[hoje.getMonth()]; const ano = hoje.getFullYear(); const dataAtual = `${dia} de ${nomeMes} de ${ano}`;
    body.replaceText(`{{NOME}}`, formData.NOME || '');
    body.replaceText(`{{GRAU_HIERARQUICO}}`, formData.GRAU_HIERARQUICO || ''); body.replaceText(`{{NIP_MAT}}`, formData.NIP_MAT || ''); body.replaceText(`{{FINALIDADE_INSP}}`, formData.FINALIDADE_INSP || ''); body.replaceText(`{{INFO_COMPLEMENTARES}}`, formData.INFO_COMPLEMENTARES || ''); body.replaceText(`{{PERITO}}`, peritoFullName.toUpperCase()); body.replaceText(`{{POSTO}}`, posto);
    body.replaceText(`{{DATA}}`, dataAtual); body.replaceText(`{{MEMBRO}}`, membroValue);
    doc.saveAndClose();
    const pdfFile = pdfFolder.createFile(newDocFile.getAs(MimeType.PDF)).setName(newDocName + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    newDocFile.setTrashed(true);
    return { success: true, pdfUrl: pdfFile.getUrl(), pdfId: pdfFile.getId() };
  } catch (error) {
    Logger.log('Erro em processForm: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function sendPdfByEmail(pdfId, formData, emailTarget) {
  try {
    const pdfFile = DriveApp.getFileById(pdfId);
    const peritoSelecionado = formData.PERITO_SELECAO;
    const peritoEmail = PERITO_EMAILS[peritoSelecionado] || '';
    const peritoFullName = getPeritoFullName(peritoSelecionado);
    const recipient = emailTarget === 'zimbra' ? peritoEmail : 'hnre.jrs@marinha.mil.br';
    const subject = `SOL PARECER PSIQ ${formData.GRAU_HIERARQUICO} ${formData.NOME}`;
    const body = `TRM em anexo SOL de parecer psiquiátrico para conclusão de IS ${formData.FINALIDADE_INSP} atinente a ${formData.NIP_MAT} ${formData.GRAU_HIERARQUICO} ${formData.NOME}.<br><br>Atenciosamente,<br><br><b>${peritoFullName.toUpperCase()}</b><br>JRS/HNRe`;
    MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: body, name: peritoFullName, replyTo: peritoEmail, attachments: [pdfFile.getAs(MimeType.PDF)] });
    return { success: true, message: 'E-mail enviado para ' + recipient };
  } catch (e) {
    Logger.log('Erro em sendPdfByEmail: ' + e.toString());
    return { success: false, message: 'Erro ao enviar e-mail: ' + e.toString() };
  }
}

function getPeritoFullName(peritoSelection) { const nomes = { 'CT Mauriston': 'Mauriston Renan Martins Silva', 'CT Júlio César': 'Júlio César Xavier Filho', '1T Salyne': 'Salyne Regina Martins Roberto', '1T Luz': 'Lucas Luz Nunes', '2T Trindade': 'Marcelo Fulco Trindade' };
  return nomes[peritoSelection] || ''; }
function getPeritoPosto(peritoSelection) { const postos = { 'CT Mauriston': 'Capitão-Tenente (Md)', '1T Salyne': 'Primeiro-Tenente (Md)', '1T Luz': 'Primeiro-Tenente (Md)', 'CT Júlio César': 'Capitão-Tenente (RM2-Md)', '2T Trindade': 'Segundo-Tenente (RM2-Md)' };
  return postos[peritoSelection] || ''; }

// ================================================================= //
//     --- CÓDIGO DO MENU PERSONALIZADO (v3 - ROBUSTO) ---         //
//     (Esta é a única seção que foi adicionada / corrigida)         //
// ================================================================= //

// --- CONSTANTES GLOBAIS DO MENU ---
const DEPARA_SHEET_NAME = 'DEPARA_ANEXO_M'; // Nome da aba de-para
const MINUTA_OUTPUT_COL_NAME = 'MINUTA';    // Nome da coluna onde a minuta será escrita

const NATO_ALPHABET = [
  'ALFA', 'BRAVO', 'CHARLIE', 'DELTA', 'ECHO', 'FOXTROT', 'GOLF', 'HOTEL',
  'INDIA', 'JULIETT', 'KILO', 'LIMA', 'MIKE', 'NOVEMBER', 'OSCAR', 'PAPA',
  'QUEBEC', 'ROMEO', 'SIERRA', 'TANGO', 'UNIFORM', 'VICTOR', 'WHISKEY',
  'X-RAY', 'YANKEE', 'ZULU'
];
const MESES_ABREV = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];


// -------------------------------------------------------------------
// FUNÇÕES DO MENU
// -------------------------------------------------------------------

/**
 * Cria o menu 'Minutas JRS' ao abrir a planilha.
 * Esta função é a única 'onOpen' neste arquivo,
 * garantindo que não haja declarações duplicadas.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Minutas JRS')
    .addItem('1. Gerar Minuta(s) para Linha(s) Selecionada(s)', 'gerarMinutas')
    .addToUi();
}


/**
 * Função principal (controladora) acionada pelo menu.
 */
function gerarMinutas() {
  const ui = SpreadsheetApp.getUi();
  const sheet = listaControleSheet; // Usa a constante global do seu código
  if (!sheet) {
    ui.alert("Erro: A aba 'ListaControle' não foi encontrada.");
    return;
  }
  
  const range = SpreadsheetApp.getActiveRange();
  
  if (range.getSheet().getName() !== sheet.getName()) {
    ui.alert("Erro: Por favor, selecione as linhas na aba 'ListaControle' para gerar a minuta.");
    return;
  }

  // --- LÓGICA DINÂMICA DE COLUNAS (v3) ---
  let colIndices;
  try {
    // Esta função agora lê seu cabeçalho para encontrar as colunas corretas
    colIndices = getDynamicColumnIndices(sheet);
  } catch (e) {
    ui.alert(`Erro ao mapear colunas: ${e.message}`);
    return;
  }
  // --- FIM DA LÓGICA ---

  const numRows = range.getNumRows();
  const firstRowIndex = range.getRowIndex();
  const selectedData = range.getValues();

  try {
    if (numRows > 1) {
      processarMultiplasLinhas(sheet, selectedData, firstRowIndex, colIndices);
    } else {
      processarLinhaUnica(sheet, selectedData[0], firstRowIndex, colIndices);
    }
  } catch (e) {
    Logger.log(e);
    ui.alert(`Erro inesperado: ${e.message}`);
  }
}

// -------------------------------------------------------------------
// PROCESSADORES (Controladores de Lógica)
// -------------------------------------------------------------------

/**
 * Processa uma única linha selecionada.
 */
function processarLinhaUnica(sheet, rowData, rowIndex, colIndices) {
  const ui = SpreadsheetApp.getUi();
  // Lê o status da coluna correta, encontrada dinamicamente
  const status = getString(rowData[colIndices.STATUSIS]); 
  let minuta = '';

  if (!status) {
    ui.alert('A célula de Status (Coluna "StatusIS") está vazia.');
    return;
  }

  const row = mapRowData(rowData, colIndices);

  switch (status) {
    case 'TIS Assinado':
      minuta = gerarMsgConclusao(row, colIndices);
      break;
    case 'Cancelada':
      minuta = gerarMsgCanceladaUnica(row);
      break;
    case 'Faltou':
      minuta = gerarMsgFaltaUnica(row);
      break;
    default:
      return; // Silencioso se não for um status válido
  }

  if (minuta) {
    // Escreve na coluna correta (índice + 1 para se tornar número da coluna)
    sheet.getRange(rowIndex, colIndices.MINUTA + 1).setValue(minuta); 
  }
}

/**
 * Processa múltiplas linhas selecionadas.
 */
function processarMultiplasLinhas(sheet, selectedData, firstRowIndex, colIndices) {
  const ui = SpreadsheetApp.getUi();
  const firstRow = mapRowData(selectedData[0], colIndices);
  const statusBase = firstRow.StatusIS;
  const omBase = firstRow.OM;
  let minuta = '';

  if (!statusBase) {
    ui.alert('Erro: A primeira linha selecionada não possui StatusIS.');
    return;
  }
  
  if (statusBase === 'TIS Assinado') {
    ui.alert('Erro: Não é permitido gerar minutas múltiplas para "TIS Assinado". Processe uma de cada vez.');
    return;
  }

  const rows = [firstRow];

  for (let i = 1; i < selectedData.length; i++) {
    const currentRow = mapRowData(selectedData[i], colIndices);

    if (getString(currentRow.StatusIS) !== statusBase) {
      ui.alert('Erro: Para minutas múltiplas, todas as linhas devem ter o MESMO StatusIS.');
      return;
    }
    if (getString(currentRow.OM) !== omBase) {
      ui.alert('Erro: Para minutas múltiplas, todos os militares devem ser da MESMA OM.');
      return;
    }
    rows.push(currentRow);
  }

  switch (statusBase) {
    case 'Cancelada':
      minuta = gerarMsgCanceladaMultipla(rows);
      break;
    case 'Faltou':
      minuta = gerarMsgFaltaMultipla(rows);
      break;
  }

  if (minuta) {
    sheet.getRange(firstRowIndex, colIndices.MINUTA + 1).setValue(minuta); // +1 para índice
  }
}


// -------------------------------------------------------------------
// GERADORES DE TEMPLATE (StatusIS)
// -------------------------------------------------------------------

/**
 * GERADOR: TIS Assinado (Diretriz 12)
 */
function gerarMsgConclusao(row, colIndices) { // colIndices passado como argumento
  if (!row.Finalidade || !row.PGQ || !row.Laudo || !row.DataLaudo || !row.TIS || !row.DS1a) {
    return 'ERRO: Faltam dados essenciais (Finalidade, Laudo, DataLaudo, TIS, etc.) para gerar a minuta.';
  }

  const destinatarios = buscarDestinatarios(row); // Não precisa de colIndices aqui
  if (!destinatarios.AÇÃO) {
    return `ERRO: Não foi possível encontrar regra de tramitação para ${row.Finalidade} / ${row.Laudo} / RM2: ${row.isRM2} / Restr: ${row.temRestricao}. Verifique a aba ${DEPARA_SHEET_NAME}.`;
  }

  const dataLaudoFmt = formatarData(row.DataLaudo);
  const restricoesTxt = row.Restricoes ? ` Deverá ser dispensado de ${row.Restricoes}.` : '';

  const template = `
PARA: ${destinatarios.AÇÃO}
INFO:  ${destinatarios.INFORMAÇÃO}

ASSUNTO: INSPEÇÃO DE SAÚDE FIM ${row.Finalidade} - ${row.PGQ}	${row.NIP} ${row.Inspecionado} (${row.OM})

ALFA - PTC que a JRS/HNRe concluiu em ${dataLaudoFmt} a IS FIM ${row.Finalidade} atinente ao ${row.PGQ}	${row.NIP} ${row.Inspecionado} (${row.OM}) e exarou o seguinte laudo: "${row.Laudo}".${restricoesTxt} ACD TIS nº ${row.TIS}; e

BRAVO - O REF TIS (modelo DS-1A) pode ser verificado digitalmente por meio do acesso ao sítio “https://sinais.dsm.mb/tisonline”, informando o código de validação ${row.DS1a} e selecionando a opção “Download do TIS” BT
  `;
  return template.trim();
}

/**
 * GERADOR: Cancelada - IS Única (Diretriz 13.2)
 */
function gerarMsgCanceladaUnica(row) {
  if (!row.IS || !row.Finalidade || !row.PGQ || !row.NIP || !row.Inspecionado || !row.OM) {
    return 'ERRO: Faltam dados essenciais (IS, Finalidade, OM, NIP, etc.) para esta minuta.';
  }

  const template = `
PARA: OM DE ORIGEM
INFO: ÓRGÃO DE PESSOAL + SDP

ASSUNTO: CAN IS FIM ${row.Finalidade} - ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM})

CFM DGPM-406 (9ª rev) incisos 2.1.2 alíneas “a” e “b” e 2.3.3 alínea “e”, PTC:

ALFA - JRS/HNRe CAN IS nº ${row.IS} FIM ${row.Finalidade} atinente ao ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM}) por não comparecer a esta JRS para proceder aos agendamento da IS dentro do prazo de 7 dias úteis a contar da data de APS via SEIS; 

BRAVO - o REF MIL deverá ser APS novamente via SEIS pelo encarregado da Divisão de Pessoal dessa OM para nova IS; e

CHARLIE - ACD antecitada norma, os inspecionados são responsáveis pelo agendamento no prazo estipulado e comparecimento às JS na data marcada para realização de sua IS, podendo a inobservância de tal orientação ser caracterizada como contravenção disciplinar com medidas cabíveis a critério do Titular dessa OM BT
  `;
  return template.trim();
}

/**
 * GERADOR: Cancelada - IS Múltiplas (Diretriz 13.3)
 */
function gerarMsgCanceladaMultipla(rows) {
  let itensDinamicos = [];
  
  rows.forEach((row, index) => {
    if (!row.IS || !row.Finalidade || !row.PGQ || !row.NIP || !row.Inspecionado || !row.OM) {
      throw new Error(`A linha ${index + 1} da seleção (IS ${row.IS}) tem dados faltantes.`);
    }
    const letra = NATO_ALPHABET[index];
    itensDinamicos.push(
      `${letra} - IS nº ${row.IS} FIM ${row.Finalidade} atinente ao ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM});`
    );
  });

  const letraFinal1 = NATO_ALPHABET[rows.length];
  const letraFinal2 = NATO_ALPHABET[rows.length + 1];

  const template = `
PARA: OM DE ORIGEM
INFO: ÓRGÃO DE PESSOAL + SDP

ASSUNTO: CANCELAMENTO DE INSPEÇÕES DE SAÚDE

CFM DGPM-406 (9ª rev) incisos 2.1.2 alíneas “a” e “b” e 2.3.3 alínea “e”, PTC JRS/HNRe CAN IS MIL abaixo relacionados por não comparecerem a esta JRS para proceder aos agendamento da IS dentro do prazo de 7 dias úteis a contar da data de APS via SEIS:

${itensDinamicos.join('\n\n')}

${letraFinal1} - os REF MIL deverão ser APS novamente via SEIS pelo encarregado da Divisão de Pessoal dessa OM para nova IS; e

${letraFinal2} - ACD antecitada norma, os inspecionados são responsáveis pelo agendamento no prazo estipulado e comparecimento às JS na data marcada para realização de sua IS, podendo a inobservância de tal orientação ser caracterizada como contravenção disciplinar com medidas cabíveis a critério do Titular dessa OM BT
  `;
  return template.trim();
}

/**
 * GERADOR: Faltou - IS Única (Diretriz 14.2)
 */
function gerarMsgFaltaUnica(row) {
  if (!row.IS || !row.Finalidade || !row.PGQ || !row.NIP || !row.Inspecionado || !row.OM || !row.DataEntrevista) {
    return 'ERRO: Faltam dados essenciais (IS, DataEntrevista, OM, NIP, etc.) para esta minuta.';
  }

  const dataEntrevistaFmt = formatarData(row.DataEntrevista);

  const template = `
PARA: OM DE ORIGEM
INFO: ÓRGÃO DE PESSOAL + SDP

ASSUNTO: CAN IS FIM ${row.Finalidade} - ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM})

CFM DGPM-406 (9ª rev) incisos 2.1.2 alíneas “a” e “b” e 2.3.3 alínea “e”, PTC:

ALFA - JRS/HNRe CAN IS nº ${row.IS} FIM ${row.Finalidade} atinente ao ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM}) agendada em ${dataEntrevistaFmt} por não comparecimento; 

BRAVO - o REF MIL deverá ser APS novamente via SEIS pelo encarregado da Divisão de Pessoal dessa OM para nova IS; e

CHARLIE - ACD antecitada norma, os inspecionados são responsáveis pelo agendamento no prazo estipulado e comparecimento às JS na data marcada para realização de sua IS, podendo a inobservância de tal orientação ser caracterizada como contravenção disciplinar com medidas cabíveis a critério do Titular dessa OM BT
  `;
  return template.trim();
}

/**
 * GERADOR: Faltou - IS Múltiplas (Diretriz 14.3)
 */
function gerarMsgFaltaMultipla(rows) {
  let itensDinamicos = [];
  
  rows.forEach((row, index) => {
    if (!row.IS || !row.Finalidade || !row.PGQ || !row.NIP || !row.Inspecionado || !row.OM || !row.DataEntrevista) {
      throw new Error(`A linha ${index + 1} da seleção (IS ${row.IS}) não possui DataEntrevista.`);
    }
    const letra = NATO_ALPHABET[index];
    const dataEntrevistaFmt = formatarData(row.DataEntrevista);
    itensDinamicos.push(
      `${letra} - IS nº ${row.IS} FIM ${row.Finalidade} atinente ao ${row.PGQ} ${row.NIP} ${row.Inspecionado} (${row.OM}) agendada em ${dataEntrevistaFmt};`
    );
  });

  const letraFinal1 = NATO_ALPHABET[rows.length];
  const letraFinal2 = NATO_ALPHABET[rows.length + 1];

  const template = `
PARA: OM DE ORIGEM
INFO: ÓRGÃO DE PESSOAL + SDP

ASSUNTO: CANCELAMENTO DE INSPEÇÕES DE SAÚDE

CFM DGPM-406 (9ª rev) incisos 2.1.2 alíneas “a” e “b” e 2.3.3 alínea “e”, PTC JRS/HNRe CAN IS MIL abaixo relacionados por não comparecerem nas datas agendadas para as IS:

${itensDinamicos.join('\n\n')}

${letraFinal1} - os REF MIL deverão ser APS novamente via SEIS pelo encarregado da Divisão de Pessoal dessa OM para nova IS; e

${letraFinal2} - ACD antecitada norma, os inspecionados são responsáveis pelo agendamento no prazo estipulado e comparecimento às JS na data marcada para realização de sua IS, podendo a inobservância de tal orientação ser caracterizada como contravenção disciplinar com medidas cabíveis a critério do Titular dessa OM BT
  `;
  return template.trim();
}


// -------------------------------------------------------------------
// FUNÇÕES AUXILIARES (Helpers) - v3 ROBUSTA
// -------------------------------------------------------------------

/**
 * Normaliza um cabeçalho para ser usado como chave de objeto.
 * Ex: "DS-1a" -> "DS1A"
 * Ex: "Restrições" -> "RESTRICOES"
 * Ex: "P/G/Q" -> "PGQ"
 */
function normalizeHeader(header) {
  return getString(header)
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/[^A-Z0-9]/g, ''); // Remove não-alfanuméricos
}

/**
 * Lê a linha 1 da 'ListaControle' e cria um mapa de índices (0-based)
 * dos nomes das colunas de que precisamos.
 * Ex: { IS: 0, DATAENTREVISTA: 1, ..., MINUTA: 14 }
 */
function getDynamicColumnIndices(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  
  // Nomes de cabeçalho NORMALIZADOS (sem acento, sem traço, tudo maiúsculo)
  const requiredHeaders = [
    'IS', 'DATAENTREVISTA', 'NIP', 'PGQ', 'INSPECIONADO', 'OM', 'FINALIDADE',
    'DATALAUDO', 'LAUDO', 'RESTRICOES', 'TIS', 'DS1A', 'STATUSIS', 'MINUTA'
  ];

  // Mapeia o nome do cabeçalho normalizado para o índice (0-based)
  headers.forEach((header, index) => {
    const normalized = normalizeHeader(header);
    if (normalized) {
      indices[normalized] = index;
    }
  });

  // Validação: Verifica se todas as colunas necessárias foram encontradas
  requiredHeaders.forEach(headerName => {
    if (indices[headerName] === undefined) {
      // Tenta encontrar o nome original para a mensagem de erro
      const originalHeaders = ['IS', 'DataEntrevista', 'NIP', 'P/G/Q', 'Inspecionado', 'OM', 'Finalidade', 'DataLaudo', 'Laudo', 'Restrições', 'TIS', 'DS-1a', 'StatusIS', 'MINUTA'];
      const originalName = originalHeaders[requiredHeaders.indexOf(headerName)];
      throw new Error(`Coluna obrigatória "${originalName}" não encontrada no cabeçalho da 'ListaControle'. Verifique se o nome está correto.`);
    }
  });
  
  Logger.log("Mapeamento de colunas dinâmico criado:", indices);
  return indices;
}


/**
 * Mapeia um array de dados de linha (rowData) para um objeto de fácil utilização,
 * com base nos índices dinâmicos normalizados.
 */
function mapRowData(rowData, colIndices) {
  const data = {
    IS: rowData[colIndices.IS],
    DataEntrevista: rowData[colIndices.DATAENTREVISTA],
    NIP: rowData[colIndices.NIP],
    PGQ: rowData[colIndices.PGQ], // Usa a chave normalizada 'PGQ'
    Inspecionado: rowData[colIndices.INSPECIONADO],
    OM: rowData[colIndices.OM],
    Finalidade: rowData[colIndices.FINALIDADE],
    DataLaudo: rowData[colIndices.DATALAUDO],
    Laudo: rowData[colIndices.LAUDO],
    Restricoes: rowData[colIndices.RESTRICOES], // Usa a chave normalizada 'RESTRICOES'
    TIS: rowData[colIndices.TIS],
    DS1a: rowData[colIndices.DS1A], // Usa a chave normalizada 'DS1A'
    StatusIS: rowData[colIndices.STATUSIS]
  };

  // Adiciona as chaves lógicas do Anexo M
  data.temRestricao = getString(data.Restricoes) ? 'SIM' : 'NÃO';
  data.isRM2 = getString(data.PGQ).includes('RM2') ? 'SIM' : 'NÃO';
  
  // Limpa e padroniza os dados para correspondência
  data.Finalidade = getString(data.Finalidade);
  data.Laudo = getString(data.Laudo);
  data.StatusIS = getString(data.StatusIS);
  data.OM = getString(data.OM);
  data.PGQ = getString(data.PGQ);

  return data;
}

/**
 * Busca os destinatários AÇÃO e INFORMAÇÃO na aba DEPARA_ANEXO_M.
 */
function buscarDestinatarios(row) {
  const deparaSheet = ss.getSheetByName(DEPARA_SHEET_NAME); // Usa 'ss' global
  if (!deparaSheet) {
    throw new Error(`Aba de mapeamento "${DEPARA_SHEET_NAME}" não foi encontrada.`);
  }

  const deparaData = deparaSheet.getRange(2, 1, deparaSheet.getLastRow() - 1, 6).getValues();

  for (const deparaRow of deparaData) {
    const chaveFin = getString(deparaRow[0]);
    const chaveLaudo = getString(deparaRow[1]);
    const chaveRest = getString(deparaRow[2]);
    const chaveRM2 = getString(deparaRow[3]);

    // Lógica de correspondência (match)
    // Normaliza os valores da planilha DEPARA para a comparação
    const finMatch = (chaveFin.toUpperCase() === row.Finalidade.toUpperCase()) || (chaveFin === '(qualquer)');
    
    // O laudo da planilha de-para deve ser (qualquer) OU corresponder exatamente ao laudo da ListaControle
    const laudoMatch = (chaveLaudo.toUpperCase() === row.Laudo.toUpperCase()) || (chaveLaudo === '(qualquer)');
    
    const restMatch = (chaveRest.toUpperCase() === row.temRestricao.toUpperCase()) || (chaveRest === '(qualquer)');
    const rm2Match = (chaveRM2.toUpperCase() === row.isRM2.toUpperCase()) || (chaveRM2 === '(qualquer)');

    if (finMatch && laudoMatch && restMatch && rm2Match) {
      Logger.log(`Regra encontrada: ${chaveFin}, ${chaveLaudo}, ${chaveRest}, ${chaveRM2}`);
      return {
        AÇÃO: getString(deparaRow[4]),
        INFORMAÇÃO: getString(deparaRow[5]) || '-' // Se INFORMAÇÃO estiver vazia, usa '-'
      };
    }
  }

  return { AÇÃO: null, INFORMAÇÃO: null };
}

/**
 * Formata uma data (objeto Date ou string) para o padrão DDMMMAAA (ex: 13OUT2025).
 */
function formatarData(data) {
  try {
    if (!data) return '';
    
    if (!(data instanceof Date)) {
      data = new Date(data);
    }
    
    if (isNaN(data.getTime())) {
      return "(Data Inválida)";
    }
    
    // Adiciona o fuso horário para corrigir a data lida da planilha
    // Isso é crucial pois o Google Sheets pode passar datas em UTC
    const offset = data.getTimezoneOffset() * 60000;
    const correctedDate = new Date(data.getTime() + offset);

    const dia = correctedDate.getDate().toString().padStart(2, '0');
    const mes = MESES_ABREV[correctedDate.getMonth()];
    const ano = correctedDate.getFullYear();

    return `${dia}${mes}${ano}`;
  } catch (e) {
    Logger.log(`Erro ao formatar data: ${e}`);
    return "(Erro Data)";
  }
}

/**
 * Garante que o valor seja uma string e remove espaços em branco.
 */
function getString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim();
}
