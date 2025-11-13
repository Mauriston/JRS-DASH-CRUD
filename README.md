<h1><img src="https://i.imgur.com/KUbQz08.png" width="50"> Análise Detalhada do Aplicativo JRS</h1>

Estrutura do Banco de Dados, Páginas e Funcionalidades de um aplicativo Web (Google Apps Script) para gerenciamento de Inspeções de Saúde (IS) e geração de pareceres médicos.

💾 1. Estrutura das Planilhas (O Banco de Dados)

O aplicativo utiliza o Google Sheets como um banco de dados relacional.

A. Planilha Principal (Vinculada ao Code.gs)

📋 ListaControle (A "Tabela" Principal de Inspeções)

Esta é a tabela central que armazena cada registro de Inspeção de Saúde (IS). É onde os dados do Formulario.html são salvos e de onde o Dashboard.html lê as informações.

Estrutura de Colunas (Inferida):

A: IS (ID)

B: DataEntrevista

C: Finalidade

D: OM

E: P/G/Q

F: NIP

G: Inspecionado

H: StatusIS

I: DataLaudo

J: Laudo

K: Restrições

L: TIS

M: DS-1a

N: MSG

📊 ListasRef (A "Tabela" de Referência)

Armazena as listas de opções usadas nos menus suspensos (dropdowns) do aplicativo, permitindo fácil atualização sem alterar o código.

Estrutura de Colunas:

A: Finalidade

B: OM

C: P/G

F: StatusIS

G: Restrições

🏷️ MilitaresHNRe (A "Tabela" de Militares Locais)

Usada pelo Formulario.html para autocompletar dados de militares do HNRe.

Estrutura de Colunas:

A: P/G

B: NIP

C: Inspecionado

🎓 ListaConcursos (A "Tabela" de Concursos)

Armazena dados de inspeções de concursos, que são agregados no Dashboard.html (KPIs e gráficos).

Estrutura de Colunas (Parcial):

A: EventDate

G: Finalidade

I: StatusIS

B. Planilha Externa (Dados para Parecer.html)

O aplicativo usa uma segunda planilha, externa, para a página "Gerar Parecer".

👤 Aba MILITAR (na planilha externa)

Usada pelo Parecer.html para autocompletar dados de militares.

Estrutura de Colunas:

A: NIP

B: name

C: OM

D: posto

🌐 2. Páginas do App Web

O aplicativo é composto por 3 páginas principais, controladas pela função doGet no Code.gs usando um parâmetro de URL (?page=...).

📝 Formulario.html (Página Padrão)

Propósito: Criar e adicionar novas Inspeções de Saúde na planilha ListaControle.

Interface: Um formulário de entrada de dados.

Lógica Dinâmica:

Se a OM "HNRe" é selecionada, o campo "Inspecionado" se transforma num menu de busca (com Select2) que puxa dados da MilitaresHNRe. NIP e P/G são preenchidos automaticamente.

Se "Outras" OMs são selecionadas, os campos são de entrada manual.

A seção "Conclusão" (Laudo, Data, TIS, etc.) só aparece se o "Status" for "Concluída", "Votada JRS" ou "TIS assinado".

A seção "Restrições" só aparece se o "Laudo" estiver preenchido, a OM for "HNRe" e a "Finalidade" for de verificação/término.

📈 Dashboard.html (Página: ?page=dashboard)

Propósito: Visualizar, filtrar e gerenciar todas as inspeções.

Interface: Uma página de dashboard com KPIs, gráficos e uma tabela de dados detalhada e interativa.

Lógica Dinâmica:

Carrega todos os dados das planilhas ListaControle e ListaConcursos de uma vez.

Filtros Globais (KPIs e Gráficos): Os filtros de "Ano" e "Mês" afetam os KPIs (ex: "Total de Inspeções") e os gráficos.

Filtros da Tabela (Client-Side): A tabela principal tem seus próprios filtros (busca, finalidades, "chips" de status) que operam no navegador (JavaScript), tornando a filtragem instantânea.

✍️ Parecer.html (Página: ?page=parecer)

Propósito: Gerar documentos PDF de pareceres médicos com base em templates do Google Docs.

Interface: Um formulário para solicitar o parecer (Especialidade, Militar, Finalidade, Perito, etc.).

Lógica Dinâmica:

Possui um seletor de OM que busca dados da planilha externa de militares.

Ao enviar, o formulário é desabilitado e 3 novos botões aparecem: "Novo Parecer", "Abrir PDF" e "Enviar Email".

O modal "Enviar Email" permite enviar o PDF para o perito ("zimbra") ou para a secretaria.

⚙️ 3. Funcionalidades Detalhadas (CRUD)

➕ Criação de Dados (CREATE)

Gatilho: Envio do Formulario.html.

Ação: A função addNewInspection no Code.gs é chamada.

Lógica: Recebe os dados, formata as "Restrições" (combinando a seleção com o texto "Outros") e adiciona uma nova linha (appendRow) à planilha ListaControle.

🔍 Leitura de Dados (READ)

Gatilho: Carregamento do Dashboard.html, Formulario.html e Parecer.html.

Ação: Funções como getDashboardData, getDropdownData, etc., são chamadas.

Lógica: Leem os dados das planilhas ListaControle, ListaConcursos, ListasRef, MilitaresHNRe e da planilha Externa para popular a interface (tabelas, gráficos, dropdowns).

🔄 Atualização de Dados (UPDATE)

Esta é a funcionalidade mais complexa, centrada no Dashboard.html:

Atualizar Status da Mensagem (MSG):

Gatilho: Clique no ícone sync na tabela.

Lógica (updateMsgStatus): Localiza a linha pelo Nº da IS e altera a coluna "MSG" para "ENVIADA".

Atualizar Conclusão da Inspeção (Edição Completa):

Gatilho: Clique no ícone edit na tabela.

Lógica (updateInspectionConclusion): Abre um modal, permite a edição e, ao salvar, o script automaticamente define o novo StatusIS com base nos campos preenchidos (Laudo, TIS, DS-1a).

Remarcar Inspeção:

Gatilho: Clique no ícone event_repeat na tabela.

Lógica (remarcarInspecao): Abre um modal, atualiza a DataEntrevista e define o StatusIS como "Remarcada".

Visualizar Detalhes (com Lógica):

Gatilho: Clique no ícone more_vert.

Lógica Oculta: O modal detailsModal gera automaticamente uma MINUTA MSG se a inspeção tiver um código DS-1a, usando os dados daquela linha.

📄 Geração de Documentos (LÓGICA)

Gatilho: Envio do formulário no Parecer.html.

Ação: A função processForm é chamada.

Lógica:

Encontra o "Template" (Google Doc) correto baseado na "Especialidade".

Faz uma cópia temporária do template.

Substitui os placeholders (ex: {{NOME}}) pelos dados do formulário.

Exporta a cópia como PDF para uma pasta no Google Drive.

Apaga a cópia temporária e retorna o link do PDF.

Ação de Email (sendPdfByEmail): Pega o PDF gerado e o envia como anexo para o perito ou secretaria.

❌ Funcionalidades Não Encontradas (DELETE)

Não foi identificada nenhuma função no código que permita ao usuário excluir um registro de inspeção. As operações de exclusão, aparentemente, devem ser feitas manualmente direto na planilha.
