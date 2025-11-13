<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <base target="_top">
    <title>Análise Detalhada - App JRS</title>
    
    <!-- Importando as fontes usadas no seu app (Oswald e Carlito) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Carlito:wght@400;700&family=Oswald:wght@400;700&display=swap" rel="stylesheet">
    
    <!-- Importando os ícones do Material Symbols (usados no seu app) -->
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" />
    
    <style>
        :root {
            --primary-color: #050f41;
            --success-color: #146c43;
            --warning-color: #FAB932;
            --danger-color: #990000;
            --info-color: #6c757d;
            --font-oswald: 'Oswald', sans-serif;
            --font-carlito: 'Carlito', sans-serif;
            --card-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
        }

        body {
            font-family: var(--font-carlito);
            background-color: #f4f4f9; /* Fundo cinza claro (do seu app) */
            margin: 0;
            padding: 20px;
            color: #333;
        }

        .container {
            max-width: 1000px;
            margin: 20px auto;
            background-color: transparent;
        }

        /* Cabeçalho da página de análise */
        .header-box {
            display: flex;
            align-items: center;
            background-color: var(--primary-color);
            color: #fff;
            padding: 20px 30px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: var(--card-shadow);
        }

        .header-box .icon {
            font-size: 50px;
            margin-right: 20px;
            font-variation-settings: 'FILL' 1, 'wght' 400;
        }

        .header-box h1 {
            font-family: var(--font-oswald);
            font-size: 32px;
            margin: 0;
            font-weight: 700;
            line-height: 1.1;
        }

        .header-box p {
            font-family: var(--font-carlito);
            font-size: 18px;
            margin: 5px 0 0;
            opacity: .9;
        }

        /* Estilo para as seções principais */
        .section {
            margin-bottom: 30px;
        }

        .section-title {
            display: flex;
            align-items: center;
            font-family: var(--font-oswald);
            font-size: 28px;
            font-weight: 700;
            text-transform: uppercase;
            gap: 15px;
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid var(--primary-color);
        }
        
        .section-title .icon {
            font-size: 36px;
            font-variation-settings: 'FILL' 1, 'wght' 600;
        }

        /* Estilo para os cards de conteúdo */
        .card {
            background-color: #fff;
            border-radius: 8px;
            box-shadow: var(--card-shadow);
            border: 1px solid #e9e9e9;
            margin-bottom: 20px;
            overflow: hidden; /* Para o cabeçalho do card */
        }

        .card-header {
            background-color: #f9f9f9;
            padding: 15px 20px;
            border-bottom: 1px solid #e9e9e9;
        }
        
        .card-header h3 {
            font-family: var(--font-oswald);
            font-size: 20px;
            margin: 0;
            color: var(--primary-color);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .card-header h3 .icon {
            font-size: 24px;
            font-variation-settings: 'wght' 600;
        }

        .card-header h4 {
            font-family: var(--font-oswald);
            font-size: 18px;
            margin: 15px 0 5px 0;
            color: #444;
        }
        
        .card-body {
            padding: 20px;
        }

        .card-body p {
            margin-top: 0;
            line-height: 1.6;
        }

        /* Estilos para listas (usadas na estrutura das planilhas) */
        ul.columns {
            padding-left: 20px;
            column-count: 2; /* Divide a lista em colunas */
            column-gap: 20px;
            margin: 0;
        }

        ul.columns li {
            margin-bottom: 8px;
            font-family: 'Courier New', Courier, monospace; /* Fonte monoespaçada para colunas */
            font-size: 15px;
            break-inside: avoid-column;
        }
        
        ul.columns li strong {
            font-family: var(--font-carlito);
            font-weight: 700;
            color: #000;
        }

        /* Estilos para listas de funcionalidades */
        ul.features {
            padding-left: 20px;
        }
        ul.features li {
            margin-bottom: 10px;
            line-height: 1.5;
        }

        /* Destaque para nomes de arquivos e funções */
        code {
            font-family: 'Courier New', Courier, monospace;
            background-color: #eef;
            padding: 2px 5px;
            border-radius: 4px;
            font-size: 0.95em;
            color: var(--primary-color);
            font-weight: 700;
        }
        
        /* Cores especiais para tags de lógica */
        .tag {
            font-family: var(--font-oswald);
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 4px;
            color: #fff;
            font-size: 0.9em;
        }
        .tag-create { background-color: var(--success-color); }
        .tag-read { background-color: var(--info-color); }
        .tag-update { background-color: var(--warning-color); color: #333;}
        .tag-delete { background-color: var(--danger-color); }
        .tag-logic { background-color: var(--primary-color); }

    </style>
</head>
<body>
    <div class="container">
        
        <!-- CABEÇALHO -->
        <div class="header-box">
            <span class="material-symbols-outlined icon">analytics</span>
            <div>
                <h1>Análise Detalhada do Aplicativo JRS</h1>
                <p>Estrutura do Banco de Dados, Páginas e Funcionalidades</p>
            </div>
        </div>

        <!-- SEÇÃO 1: ESTRUTURA DAS PLANILHAS -->
        <div class="section">
            <h2 class="section-title">
                <span class="material-symbols-outlined icon">storage</span>
                1. Estrutura das Planilhas (O Banco de Dados)
            </h2>
            
            <div class="card">
                <div class="card-header">
                    <h4>A. Planilha Principal (Vinculada ao <code>Code.gs</code>)</h4>
                </div>
                <div class="card-body">
                    
                    <div class="card" style="box-shadow: none; border-color: #ccc; margin-bottom: 15px;">
                        <div class="card-header">
                            <h3><span class="material-symbols-outlined icon">fact_check</span><code>ListaControle</code> (A "Tabela" Principal de Inspeções)</h3>
                        </div>
                        <div class="card-body">
                            <p>Esta é a tabela central que armazena cada registro de Inspeção de Saúde (IS). É onde os dados do <code>Formulario.html</code> são salvos e de onde o <code>Dashboard.html</code> lê as informações.</p>
                            <p><strong>Estrutura de Colunas (Inferida):</strong></p>
                            <ul class="columns">
                                <li><code>A: IS</code> <strong>(ID)</strong></li>
                                <li><code>B: DataEntrevista</code></li>
                                <li><code>C: Finalidade</code></li>
                                <li><code>D: OM</code></li>
                                <li><code>E: P/G/Q</code></li>
                                <li><code>F: NIP</code></li>
                                <li><code>G: Inspecionado</code></li>
                                <li><code>H: StatusIS</code></li>
                                <li><code>I: DataLaudo</code></li>
                                <li><code>J: Laudo</code></li>
                                <li><code>K: Restrições</code></li>
                                <li><code>L: TIS</code></li>
                                <li><code>M: DS-1a</code></li>
                                <li><code>N: MSG</code></li>
                            </ul>
                        </div>
                    </div>

                    <div class="card" style="box-shadow: none; border-color: #ccc; margin-bottom: 15px;">
                        <div class="card-header">
                            <h3><span class="material-symbols-outlined icon">list_alt</span><code>ListasRef</code> (A "Tabela" de Referência)</h3>
                        </div>
                        <div class="card-body">
                            <p>Armazena as listas de opções usadas nos menus suspensos (dropdowns) do aplicativo, permitindo fácil atualização sem alterar o código.</p>
                            <p><strong>Estrutura de Colunas:</strong></p>
                            <ul class="columns">
                                <li><code>A: Finalidade</code></li>
                                <li><code>B: OM</code></li>
                                <li><code>C: P/G</code></li>
                                <li><code>F: StatusIS</code></li>
                                <li><code>G: Restrições</code></li>
                            </ul>
                        </div>
                    </div>

                    <div class="card" style="box-shadow: none; border-color: #ccc; margin-bottom: 15px;">
                        <div class="card-header">
                            <h3><span class="material-symbols-outlined icon">badge</span><code>MilitaresHNRe</code> (A "Tabela" de Militares Locais)</h3>
                        </div>
                        <div class="card-body">
                            <p>Usada pelo <code>Formulario.html</code> para autocompletar dados de militares do HNRe.</p>
                            <p><strong>Estrutura de Colunas:</strong></p>
                            <ul class="columns">
                                <li><code>A: P/G</code></li>
                                <li><code>B: NIP</code></li>
                                <li><code>C: Inspecionado</code></li>
                            </ul>
                        </div>
                    </div>

                    <div class="card" style="box-shadow: none; border-color: #ccc; margin-bottom: 0;">
                        <div class="card-header">
                            <h3><span class="material-symbols-outlined icon">school</span><code>ListaConcursos</code> (A "Tabela" de Concursos)</h3>
                        </div>
                        <div class="card-body">
                            <p>Armazena dados de inspeções de concursos, que são agregados no <code>Dashboard.html</code> (KPIs e gráficos).</p>
                            <p><strong>Estrutura de Colunas (Parcial):</strong></p>
                            <ul class="columns">
                                <li><code>A: EventDate</code></li>
                                <li><code>G: Finalidade</code></li>
                                <li><code>I: StatusIS</code></li>
                            </ul>
                        </div>
                    </div>

                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <h4>B. Planilha Externa (Dados para <code>Parecer.html</code>)</h4>
                </div>
                <div class="card-body">
                    <div class="card" style="box-shadow: none; border-color: #ccc; margin-bottom: 0;">
                        <div class="card-header">
                            <h3><span class="material-symbols-outlined icon">person_search</span>Aba <code>MILITAR</code> (na planilha externa)</h3>
                        </div>
                        <div class="card-body">
                            <p>Usada pelo <code>Parecer.html</code> para autocompletar dados de militares.</p>
                            <p><strong>Estrutura de Colunas:</strong></p>
                            <ul class="columns">
                                <li><code>A: NIP</code></li>
                                <li><code>B: name</code></li>
                                <li><code>C: OM</code></li>
                                <li><code>D: posto</code></li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- SEÇÃO 2: PÁGINAS DO APP WEB -->
        <div class="section">
            <h2 class="section-title">
                <span class="material-symbols-outlined icon">web</span>
                2. Páginas do App Web
            </h2>
            
            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon">description</span><code>Formulario.html</code> (Página Padrão)</h3>
                </div>
                <div class="card-body">
                    <p><strong>Propósito:</strong> Criar e adicionar novas Inspeções de Saúde na planilha <code>ListaControle</code>.</p>
                    <p><strong>Interface:</strong> Um formulário de entrada de dados.</p>
                    <p><strong>Lógica Dinâmica:</strong></p>
                    <ul class="features">
                        <li>Se a OM "HNRe" é selecionada, o campo "Inspecionado" se transforma num menu de busca (com Select2) que puxa dados da <code>MilitaresHNRe</code>. NIP e P/G são preenchidos automaticamente.</li>
                        <li>Se "Outras" OMs são selecionadas, os campos são de entrada manual.</li>
                        <li>A seção "Conclusão" (Laudo, Data, TIS, etc.) só aparece se o "Status" for "Concluída", "Votada JRS" ou "TIS assinado".</li>
                        <li>A seção "Restrições" só aparece se o "Laudo" estiver preenchido, a OM for "HNRe" e a "Finalidade" for de verificação/término.</li>
                    </ul>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon">dashboard</span><code>Dashboard.html</code> (Página: <code>?page=dashboard</code>)</h3>
                </div>
                <div class="card-body">
                    <p><strong>Propósito:</strong> Visualizar, filtrar e gerenciar todas as inspeções.</p>
                    <p><strong>Interface:</strong> Uma página de dashboard com KPIs, gráficos e uma tabela de dados detalhada e interativa.</p>
                    <p><strong>Lógica Dinâmica:</strong></p>
                    <ul class="features">
                        <li>Carrega <strong>todos</strong> os dados das planilhas <code>ListaControle</code> e <code>ListaConcursos</code> de uma vez.</li>
                        <li><strong>Filtros Globais (KPIs e Gráficos):</strong> Os filtros de "Ano" e "Mês" afetam os KPIs (ex: "Total de Inspeções") e os gráficos.</li>
                        <li><strong>Filtros da Tabela (Client-Side):</strong> A tabela principal tem seus próprios filtros (busca, finalidades, "chips" de status) que operam <strong>no navegador (JavaScript)</strong>, tornando a filtragem instantânea.</li>
                    </ul>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon">edit_note</span><code>Parecer.html</code> (Página: <code>?page=parecer</code>)</h3>
                </div>
                <div class="card-body">
                    <p><strong>Propósito:</strong> Gerar documentos PDF de pareceres médicos com base em templates do Google Docs.</p>
                    <p><strong>Interface:</strong> Um formulário para solicitar o parecer (Especialidade, Militar, Finalidade, Perito, etc.).</p>
                    <p><strong>Lógica Dinâmica:</strong></p>
                    <ul class="features">
                        <li>Possui um seletor de OM que busca dados da planilha <strong>externa</strong> de militares.</li>
                        <li>Ao enviar, o formulário é desabilitado e 3 novos botões aparecem: "Novo Parecer", "Abrir PDF" e "Enviar Email".</li>
                        <li>O modal "Enviar Email" permite enviar o PDF para o perito ("zimbra") ou para a secretaria.</li>
                    </ul>
                </div>
            </div>
        </div>

        <!-- SEÇÃO 3: FUNCIONALIDADES DETALHADAS -->
        <div class="section">
            <h2 class="section-title">
                <span class="material-symbols-outlined icon">functions</span>
                3. Funcionalidades Detalhadas (CRUD)
            </h2>

            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon" style="color: var(--success-color);">add_circle</span>Criação de Dados <span class="tag tag-create">CREATE</span></h3>
                </div>
                <div class="card-body">
                    <ul class="features">
                        <li><strong>Gatilho:</strong> Envio do <code>Formulario.html</code>.</li>
                        <li><strong>Ação:</strong> A função <code>addNewInspection</code> no <code>Code.gs</code> é chamada.</li>
                        <li><strong>Lógica:</strong> Recebe os dados, formata as "Restrições" (combinando a seleção com o texto "Outros") e adiciona uma nova linha (`appendRow`) à planilha <code>ListaControle</code>.</li>
                    </ul>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon" style="color: var(--info-color);">search</span>Leitura de Dados <span class="tag tag-read">READ</span></h3>
                </div>
                <div class="card-body">
                    <ul class="features">
                        <li><strong>Gatilho:</strong> Carregamento do <code>Dashboard.html</code>, <code>Formulario.html</code> e <code>Parecer.html</code>.</li>
                        <li><strong>Ação:</strong> Funções como <code>getDashboardData</code>, <code>getDropdownData</code>, etc., são chamadas.</li>
                        <li><strong>Lógica:</strong> Leem os dados das planilhas <code>ListaControle</code>, <code>ListaConcursos</code>, <code>ListasRef</code>, <code>MilitaresHNRe</code> e da planilha <code>Externa</code> para popular a interface (tabelas, gráficos, dropdowns).</li>
                    </ul>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon" style="color: var(--warning-color);">edit_square</span>Atualização de Dados <span class="tag tag-update">UPDATE</span></h3>
                </div>
                <div class="card-body">
                    <p>Esta é a funcionalidade mais complexa, centrada no <code>Dashboard.html</code>:</p>
                    <ul class="features">
                        <li><strong>Atualizar Status da Mensagem (MSG):</strong>
                            <ul>
                                <li><strong>Gatilho:</strong> Clique no ícone <code>sync</code> na tabela.</li>
                                <li><strong>Lógica (<code>updateMsgStatus</code>):</strong> Localiza a linha pelo Nº da IS e altera a coluna "MSG" para "ENVIADA".</li>
                            </ul>
                        </li>
                        <li><strong>Atualizar Conclusão da Inspeção (Edição Completa):</strong>
                            <ul>
                                <li><strong>Gatilho:</strong> Clique no ícone <code>edit</code> na tabela.</li>
                                <li><strong>Lógica (<code>updateInspectionConclusion</code>):</strong> Abre um modal, permite a edição e, ao salvar, o script <strong>automaticamente define o novo <code>StatusIS</code></strong> com base nos campos preenchidos (Laudo, TIS, DS-1a).</li>
                            </ul>
                        </li>
                        <li><strong>Remarcar Inspeção:</strong>
                            <ul>
                                <li><strong>Gatilho:</strong> Clique no ícone <code>event_repeat</code> na tabela.</li>
                                <li><strong>Lógica (<code>remarcarInspecao</code>):</strong> Abre um modal, atualiza a <code>DataEntrevista</code> e define o <code>StatusIS</code> como "Remarcada".</li>
                            </ul>
                        </li>
                        <li><strong>Visualizar Detalhes (com Lógica):</strong>
                            <ul>
                                <li><strong>Gatilho:</strong> Clique no ícone <code>more_vert</code>.</li>
                                <li><strong>Lógica Oculta:</strong> O modal <code>detailsModal</code> <strong>gera automaticamente uma MINUTA MSG</strong> se a inspeção tiver um código <code>DS-1a</code>, usando os dados daquela linha.</li>
                            </ul>
                        </li>
                    </ul>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon" style="color: #6a1b9a;">picture_as_pdf</span>Geração de Documentos <span class="tag tag-logic">LÓGICA</span></h3>
                </div>
                <div class="card-body">
                     <ul class="features">
                        <li><strong>Gatilho:</strong> Envio do formulário no <code>Parecer.html</code>.</li>
                        <li><strong>Ação:</strong> A função <code>processForm</code> é chamada.</li>
                        <li><strong>Lógica:</strong>
                            <ol>
                                <li>Encontra o "Template" (Google Doc) correto baseado na "Especialidade".</li>
                                <li>Faz uma cópia temporária do template.</li>
                                <li>Substitui os placeholders (ex: <code>{{NOME}}</code>) pelos dados do formulário.</li>
                                <li>Exporta a cópia como PDF para uma pasta no Google Drive.</li>
                                <li>Apaga a cópia temporária e retorna o link do PDF.</li>
                            </ol>
                        </li>
                        <li><strong>Ação de Email (<code>sendPdfByEmail</code>):</strong> Pega o PDF gerado e o envia como anexo para o perito ou secretaria.</li>
                    </ul>
                </div>
            </div>

            <div class="card">
                <div class="card-header">
                    <h3><span class="material-symbols-outlined icon" style="color: var(--danger-color);">delete_forever</span>Funcionalidades Não Encontradas <span class="tag tag-delete">DELETE</span></h3>
                </div>
                <div class="card-body">
                    <p>Não foi identificada nenhuma função no código que permita ao usuário <strong>excluir</strong> um registro de inspeção. As operações de exclusão, aparentemente, devem ser feitas manualmente direto na planilha.</p>
                </div>
            </div>

        </div>
    </div>
</body>
</html>
