// ================================================================= //
//                      CONFIGURAÇÕES GLOBAIS                        //
// ================================================================= //

// --- Configs do Formulário de Inspeção e Dashboard (ORIGINAL) ---
const ss = SpreadsheetApp.getActiveSpreadsheet();
const listasRefSheet = ss.getSheetByName('ListasRef');
const listaControleSheet = ss.getSheetByName('ListaControle');
const militaresHnreSheet = ss.getSheetByName('MilitaresHNRe');
const concursosSheet = ss.getSheetByName('ListaConcursos');
// --- Configs do Gerador de Pareceres (NOVO) ---
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

// --- NOVA FUNÇÃO DE UPDATE DE CONCLUSÃO (ATUALIZADA) ---
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
    // Validação dos índices
    if (colIndices.is === -1 || colIndices.laudo === -1 || colIndices.dataLaudo === -1 || colIndices.tis === -1 || colIndices.ds1a === -1 || colIndices.restricoes === -1 || colIndices.statusIS === -1) {
      throw new Error("Colunas essenciais (IS, Laudo, DataLaudo, StatusIS, etc.) não encontradas. Verifique a 1ª linha da planilha.");
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

    // 3. Formatar Restrições (lógica copiada de addNewInspection)
    let restricoesString = formData.restricoes.join(', ');
    if (formData.restricoes.includes('Outros') && formData.outrosRestricao) {
      restricoesString = restricoesString.replace('Outros', `Outros: ${formData.outrosRestricao}`);
    }

    // 4. Preparar dados e atualizar a planilha
    const targetRowOnSheet = targetRowIndex + 2;
    // +2 porque o array é 0-based e a planilha é 1-based + cabeçalho
    
    // Atualiza os valores individualmente
    listaControleSheet.getRange(targetRowOnSheet, colIndices.laudo + 1).setValue(formData.laudo || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.dataLaudo + 1).setValue(formData.dataLaudo || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.tis + 1).setValue(formData.tis || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.ds1a + 1).setValue(formData.ds1a || null);
    listaControleSheet.getRange(targetRowOnSheet, colIndices.restricoes + 1).setValue(restricoesString || null);
    
    // ATUALIZA O STATUSIS
    listaControleSheet.getRange(targetRowOnSheet, colIndices.statusIS + 1).setValue(formData.novoStatusIS || 'Conclusão Pendente');

    Logger.log(`IS ${formData.isNumber} (Linha ${targetRowOnSheet}) atualizada com sucesso.`);
    // 5. Retornar os dados formatados para o frontend
    return { 
      success: true, 
      message: 'Conclusão salva com sucesso!',
      updatedData: {
        Laudo: formData.laudo ||
        '',
        DataLaudo: formData.dataLaudo ?
        new Date(formData.dataLaudo).toLocaleDateString('pt-BR') : '', // Re-formata para pt-BR
        TIS: formData.tis ||
        '',
        'DS-1a': formData.ds1a ||
        '',
        Restrições: restricoesString ||
        '',
        StatusIS: formData.novoStatusIS || 'Conclusão Pendente' // <-- ADICIONADO
      } 
    };
  } catch (e) {
    Logger.log('Erro em updateInspectionConclusion: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// --- NOVA FUNÇÃO DE REMARCAÇÃO ---
function remarcarInspecao(formData) {
  try {
    if (!listaControleSheet) { throw new Error("Planilha 'ListaControle' não encontrada."); }

    const isNumber = formData.isNumber;
    const novaData = formData.novaData; // Formato YYYY-MM-DD

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
    const targetRowOnSheet = targetRowIndex + 2; // 1-based + cabeçalho
    
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
//           FUNÇÕES DO GERADOR DE PARECERES (NOVO MÓDULO)           //
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
