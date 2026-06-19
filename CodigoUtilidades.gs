// =====================================================================
// 🛠️ UTILIDADES DO SISTEMA HCPA
// ARQUIVO: FERRAMENTAS AVANÇADAS & CONTROLE DE ACESSO (DINÂMICO)
// =====================================================================

// --- CONFIGURAÇÃO DE SUPER ADMINS (LISTA FIXA) ---
// Estes usuários têm acesso total às configurações, independente da planilha.
var SUPER_ADMINS_LISTA = [
  "gsilveira@hcpa.edu.br", 
  "lklrocha@hcpa.edu.br", 
  "vregina@hcpa.edu.br"
];

// Nomes das Colunas na aba de configuração (Interno)
var COL_LEITORES = 0;      // A: Apenas Leitores
var COL_GUICHE = 1;        // B: Guichê (Parciais)
var COL_AUTORIZACAO = 2;   // C: Autorização (Intermediários)
var COL_ADMIN = 3;         // D: Agendamento/Admin (Completos)

// =====================================================================
// 1. VERIFICAÇÕES DE SUPER ADMIN (CORRIGIDA E ROBUSTA)
// =====================================================================

function ehSuperAdmin() {
  var emailsParaVerificar = [];
  
  // 1. Tenta pegar o usuário que está clicando na planilha (ActiveUser)
  try {
    var activeUser = Session.getActiveUser().getEmail().toLowerCase();
    if (activeUser) emailsParaVerificar.push(activeUser);
  } catch(e) {}

  // 2. Tenta pegar o usuário efetivo do script (EffectiveUser)
  try {
    var effectiveUser = Session.getEffectiveUser().getEmail().toLowerCase();
    if (effectiveUser && emailsParaVerificar.indexOf(effectiveUser) === -1) {
      emailsParaVerificar.push(effectiveUser);
    }
  } catch(e) {}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 3. Verifica se algum dos emails é o Dono do Arquivo
  try {
    var owner = DriveApp.getFileById(ss.getId()).getOwner().getEmail().toLowerCase();
    if (emailsParaVerificar.indexOf(owner) > -1) return true;
  } catch(e) {}

  // 4. Verifica se algum dos emails está na lista fixa de Super Admins
  for (var i = 0; i < emailsParaVerificar.length; i++) {
    if (SUPER_ADMINS_LISTA.indexOf(emailsParaVerificar[i]) > -1) return true;
  }

  return false;
}

// =====================================================================
// 2. NÚCLEO DE PERMISSÕES (LÊ DA PLANILHA)
// =====================================================================

// Função que carrega todas as listas da aba 'ConfigPermissoes'
function obterListasDePermissao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ConfigPermissoes');
  
  // Se a aba não existe, cria e popula com os dados antigos (Migração)
  if (!sheet) {
    return inicializarAbaPermissoes(ss);
  }

  var lastRow = sheet.getLastRow();
  // Se só tem cabeçalho ou está vazia, retorna arrays vazios
  if (lastRow < 2) return { leitores: [], guiche: [], autorizacao: [], admin: [] };

  // Lê do intervalo A2:D(ultimaLinha)
  var dados = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  var listas = {
    leitores: [],
    guiche: [],
    autorizacao: [],
    admin: []
  };

  // Itera linha a linha e preenche as listas
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][COL_LEITORES]) listas.leitores.push(String(dados[i][COL_LEITORES]).trim().toLowerCase());
    if (dados[i][COL_GUICHE]) listas.guiche.push(String(dados[i][COL_GUICHE]).trim().toLowerCase());
    if (dados[i][COL_AUTORIZACAO]) listas.autorizacao.push(String(dados[i][COL_AUTORIZACAO]).trim().toLowerCase());
    if (dados[i][COL_ADMIN]) listas.admin.push(String(dados[i][COL_ADMIN]).trim().toLowerCase());
  }

  return listas;
}

// Verifica se o usuário atual é ADMIN ou DONO (Para abrir o menu de gestão)
function ehAdminOuDono() {
  // Super Admin pode tudo
  if (ehSuperAdmin()) return true;

  var email = Session.getEffectiveUser().getEmail().toLowerCase();
  var listas = obterListasDePermissao();
  return listas.admin.indexOf(email) > -1;
}

// =====================================================================
// 2.1 VERIFICAÇÕES DE ACESSO (PERFIS) - AGORA DINÂMICAS
// =====================================================================

// PERFIL 1: ACESSO TOTAL (Agenda, Cancela, Move)
// Quem pode: Apenas ADMIN, SUPER ADMIN e DONO
function podeAlterarAgenda() {
  if (ehSuperAdmin()) return true;

  var email = Session.getEffectiveUser().getEmail().toLowerCase();
  var listas = obterListasDePermissao();
  
  // Apenas Admin pode agendar/reagendar/cancelar
  if (listas.admin.indexOf(email) > -1) return true;
  
  return false; 
}

// PERFIL 2: ACESSO RECEPÇÃO (AA:AB - Checks e Guichê)
// Quem pode: ADMIN, SUPER ADMIN e GUICHÊ (Parciais)
function podeEditarChecks() {
  if (ehSuperAdmin()) return true;

  var email = Session.getEffectiveUser().getEmail().toLowerCase();
  var listas = obterListasDePermissao();
  
  if (listas.admin.indexOf(email) > -1) return true;
  if (listas.guiche.indexOf(email) > -1) return true;
  
  return false; 
}

// PERFIL 3: ACESSO TÉCNICO (Info e AUT)
// Quem pode: ADMIN, SUPER ADMIN e AUTORIZAÇÃO (Intermediários)
function podeEditarInfoEAut() {
  if (ehSuperAdmin()) return true;

  var email = Session.getEffectiveUser().getEmail().toLowerCase();
  var listas = obterListasDePermissao();
  
  if (listas.admin.indexOf(email) > -1) return true;
  if (listas.autorizacao.indexOf(email) > -1) return true;
  
  return false; 
}

// =====================================================================
// 3. FERRAMENTAS DE CONFIGURAÇÃO (BACKEND DO MODAL)
// =====================================================================

// Função legada mantida para compatibilidade, mas redirecionando verificação
function abrirGerenciadorPermissoes() {
  if (!ehAdminOuDono()) {
    SpreadsheetApp.getUi().alert("⛔ ACESSO NEGADO\n\nApenas Administradores podem gerenciar permissões.");
    return;
  }
  
  var html = HtmlService.createTemplateFromFile('GerenciadorPermissoesHTML'); // Atenção: Verifique se este arquivo ainda existe ou se foi substituído pelo novo Painel
  var listas = obterListasDePermissao();
  html.dadosIniciais = JSON.stringify(listas);
  
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(700).setHeight(600), '👥 Gerenciar Permissões e Grupos');
}

// Agora protegida: Só Super Admin salva permissões
function salvarNovasPermissoes(novasListas) {
  if (!ehSuperAdmin()) throw new Error("⛔ Acesso Negado: Apenas Super Admins podem alterar permissões.");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ConfigPermissoes');
  if (!sheet) sheet = ss.insertSheet('ConfigPermissoes');
  
  // Limpa dados antigos (Mantém cabeçalho na linha 1)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  }
  
  // Descobre qual a maior lista para saber quantas linhas criar
  var maxLen = Math.max(
    novasListas.leitores.length, 
    novasListas.guiche.length, 
    novasListas.autorizacao.length, 
    novasListas.admin.length
  );
  
  if (maxLen === 0) return "Salvo (Listas vazias)";
  
  var dadosSaida = [];
  for (var i = 0; i < maxLen; i++) {
    dadosSaida.push([
      novasListas.leitores[i] || "",
      novasListas.guiche[i] || "",
      novasListas.autorizacao[i] || "",
      novasListas.admin[i] || ""
    ]);
  }
  
  sheet.getRange(2, 1, dadosSaida.length, 4).setValues(dadosSaida);
  
  // Aplica as permissões no Drive e nas Abas imediatamente
  configurarPermissoesPlanilhaAtual();
  
  return "✅ Permissões salvas e aplicadas com sucesso!";
}

// =====================================================================
// 4. MIGRAÇÃO INICIAL (CRIA A ABA SE NÃO EXISTIR COM DADOS ANTIGOS)
// =====================================================================
function inicializarAbaPermissoes(ss) {
  var sheet = ss.insertSheet('ConfigPermissoes');
  
  // --- AQUI ESTÃO OS DADOS ORIGINAIS PARA NÃO PERDER NADA ---
  var admins = [
    "gsilveira@hcpa.edu.br", "lklrocha@hcpa.edu.br", "vregina@hcpa.edu.br", 
    "acteixeira@hcpa.edu.br", "carolcdcosta@hcpa.edu.br", "asjorge@hcpa.edu.br", 
    "lzigue@hcpa.edu.br", "cmdasilva@hcpa.edu.br"
  ];
  
  var autorizacao = [
    "tamartins@hcpa.edu.br", "gclima@hcpa.edu.br", "gschild@hcpa.edu.br"
  ];
  
  var guiche = [
    "vcsilva@hcpa.edu.br", "rcmartins@hcpa.edu.br", "dgsilva@hcpa.edu.br", 
    "dslongo@hcpa.edu.br", "mleite@hcpa.edu.br", "marrdsilva@hcpa.edu.br", 
    "lmylius@hcpa.edu.br", "ljleal@hcpa.edu.br", "mromero@hcpa.edu.br"
  ];
  
  var leitores = [
    "mrrsilva@hcpa.edu.br", "jritt@hcpa.edu.br", "l-admcardiologia@hcpa.edu.br", "l-secret-bloco@hcpa.edu.br", "l-secret-cca@hcpa.edu.br", "l-genf-ubc-enfermeiros@hcpa.edu.br",
    "l-convenios-faturamento@hcpa.edu.br", "l-convenios-internacao@hcpa.edu.br", "fpetry@hcpa.edu.br", "mjunior@hcpa.edu.br", "gduarte@hcpa.edu.br", "fbadalotti@hcpa.edu.br",
    "rdblopes@hcpa.edu.br", "flima@hcpa.edu.br", "bbernardi@hcpa.edu.br", "lrangel@hcpa.edu.br", "smolin@hcpa.edu.br", "nsavaris@hcpa.edu.br", "ncpereira@hcpa.edu.br",
    "elcouto@hcpa.edu.br", "asvargas@hcpa.edu.br", "acarcuchinski@hcpa.edu.br", "ccaldana@hcpa.edu.br", "rminuzzi@hcpa.edu.br", "caires@hcpa.edu.br", "psdsantos@hcpa.edu.br",
    "atmiranda@hcpa.edu.br", "boliveira@hcpa.edu.br", "tmelo@hcpa.edu.br", "alexguimaraes@hcpa.edu.br", "cbrito@hcpa.edu.br", "trnunes@hcpa.edu.br", "llunardi@hcpa.edu.br",
    "rvidal@hcpa.edu.br", "ktabarkiewicz@hcpa.edu.br", "dsoliveira@hcpa.edu.br", "mebmachado@hcpa.edu.br", "bwallauer@hcpa.edu.br", "amalaquias@hcpa.edu.br", "dmello@hcpa.edu.br",
    "edaimone@hcpa.edu.br", "lmenezes@hcpa.edu.br", "gia.anestesiagia.anestesia@gmail.com", "echagas@hcpa.edu.br", "lguterres@hcpa.edu.br", "cschiavo@hcpa.edu.br",
    "mvirginio@hcpa.edu.br", "rcosta@hcpa.edu.br", "mfunari@hcpa.edu.br", "smanica@hcpa.edu.br", "ppackeiser@hcpa.edu.br", "mmferreira@hcpa.edu.br", "emallmann@hcpa.edu.br",
    "gsomm@hcpa.edu.br", "escala@gia.med.br", "valves@hcpa.edu.br", "bnovelo@hcpa.edu.br", "gpaniz@hcpa.edu.br", "fpelisoli@hcpa.edu.br", "jviesi@hcpa.edu.br",
    "gbraulio@hcpa.edu.br", "pbasso@hcpa.edu.br", "dsanto@hcpa.edu.br", "klferreira@hcpa.edu.br", "masturizaga@hcpa.edu.br"
  ];

  // Configura cabeçalho e cor
  sheet.getRange("A1:D1").setValues([["Apenas Leitores", "Guichê", "Autorização", "Agendamento/Admin"]])
        .setFontWeight("bold").setBackground("#e6e6e6");
  
  var max = Math.max(admins.length, autorizacao.length, guiche.length, leitores.length);
  var matrix = [];
  
  for(var i=0; i<max; i++) {
    matrix.push([
      leitores[i] || "",
      guiche[i] || "",
      autorizacao[i] || "",
      admins[i] || ""
    ]);
  }
  
  sheet.getRange(2, 1, matrix.length, 4).setValues(matrix);
  sheet.hideSheet(); // Oculta a aba para segurança visual
  
  return { leitores: leitores, guiche: guiche, autorizacao: autorizacao, admin: admins };
}

// =====================================================================
// 5. APLICAÇÃO DE PERMISSÕES (PROTEGE DRIVE E ABAS)
// =====================================================================

function configurarPermissoesPlanilhaAtual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  var file = DriveApp.getFileById(fileId);
  var me = Session.getEffectiveUser();

  // Carrega as listas atuais da aba de configuração
  var listas = obterListasDePermissao();

  Logger.log("--- INICIANDO CONFIGURAÇÃO ---");

  // ETAPA 1: Configurar acesso ao Arquivo (Drive)
  // Juntamos todos os tipos de editores (Admin + Guiche + Autorizacao)
  var todosEditores = listas.admin.concat(listas.guiche).concat(listas.autorizacao);
  
  Logger.log("Atualizando Editores no Drive...");
  todosEditores.forEach(function(email) {
    if(email) try { file.addEditor(email); } catch(e) {}
  });

  Logger.log("Atualizando Leitores no Drive...");
  listas.leitores.forEach(function(email) {
    if(email) try { file.addViewer(email); } catch(e) {}
  });

  // ETAPA 2: Configurar Proteções Internas (Abas)
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    var nomeAba = sheet.getName();
    
    // Ignora abas de sistema se necessário, ou protege tudo (recomendado)
    
    // Remove proteções antigas
    var protecoesAntigas = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < protecoesAntigas.length; i++) {
      protecoesAntigas[i].remove();
    }

    // Cria proteção nova
    var protecao = sheet.protect().setDescription("Acesso Restrito - Liberado AA:AB");
    
    // Libera AA:AB para TODOS os editores (inclusive Guichê)
    var rangeDesbloqueado = sheet.getRange("AA:AB");
    protecao.setUnprotectedRanges([rangeDesbloqueado]);

    // Bloqueia o resto da aba
    protecao.removeEditors(protecao.getEditors());
    protecao.addEditor(me);
    
    // Quem pode editar a parte BLOQUEADA (Dados principais, Info, Aut)?
    // Apenas ADMIN e AUTORIZAÇÃO (Intermediários)
    // O pessoal do Guichê (Parciais) NÃO entra aqui, pois eles só podem mexer no desbloqueado (AA:AB)
    var editoresProtegidos = listas.admin.concat(listas.autorizacao);
    
    // Adiciona Super Admins explicitamente se não estiverem na lista
    editoresProtegidos = editoresProtegidos.concat(SUPER_ADMINS_LISTA);
    
    if (editoresProtegidos.length > 0) {
      // Filtra vazios e duplicados
      editoresProtegidos = editoresProtegidos.filter(function(item, pos) {
          return item && editoresProtegidos.indexOf(item) == pos;
      });
      protecao.addEditors(editoresProtegidos);
    }
  });
}

// =====================================================================
// 6. FERRAMENTAS LEGADAS (MANTIDAS INTEGRALMENTE)
// =====================================================================

function relatorioELimpezaBlindada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  var ui = SpreadsheetApp.getUi();

  var resposta = ui.alert(
    'BACKUP E FAXINA (Modo Compatível)',
    'Este script irá:\n1. Salvar os comentários na aba "RELATÓRIO_FINAL".\n2. Tentar EXCLUIR os comentários.\n3. Se não conseguir excluir, vai OCULTÁ-LOS (Resolver).\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (resposta != ui.Button.YES) return;

  // --- PARTE 1: CRIAR RELATÓRIO (Backup) ---
  var mapaAbas = {};
  ss.getSheets().forEach(function(s) {
    mapaAbas[s.getSheetId()] = s.getName();
  });

  var nomeAba = "RELATÓRIO_FINAL_" + new Date().getTime();
  var sheetRep = ss.insertSheet(nomeAba);
  sheetRep.appendRow(["DATA", "STATUS ATUAL", "AUTOR", "ABA (ORIGEM)", "CONTEÚDO", "RESPOSTAS"]);
  sheetRep.getRange("A1:F1").setFontWeight("bold").setBackground("#fff2cc");
  sheetRep.setFrozenRows(1);

  var pageToken = null;
  var deletados = 0;
  var resolvidos = 0;
  var erros = 0;
  var dadosBackup = [];

  // Loop de Busca
  do {
    try {
      var resp = Drive.Comments.list(fileId, {
        pageToken: pageToken,
        maxResults: 100,
        includeDeleted: true, // Pega tudo
        fields: "nextPageToken, items(commentId, createdDate, status, anchor, content, author(displayName), replies(content, author(displayName)))"
      });

      if (resp.items && resp.items.length > 0) {
        resp.items.forEach(function(item) {
          
          // --- A. BACKUP ---
          var nomeDaAba = "Desconhecida";
          if (item.anchor) {
            try {
              var anc = JSON.parse(item.anchor);
              var uid = (anc.uid !== undefined) ? anc.uid : anc.s;
              if (uid !== undefined && mapaAbas[uid]) nomeDaAba = mapaAbas[uid];
            } catch (e) {}
          }

          var respostasTexto = "";
          if (item.replies) {
            item.replies.forEach(function(r) {
              var rAut = (r.author && r.author.displayName) ? r.author.displayName : "...";
              respostasTexto += "[" + rAut + "]: " + r.content + "\n";
            });
          }

          dadosBackup.push([
            item.createdDate,
            item.status,
            (item.author && item.author.displayName) ? item.author.displayName : "Anônimo",
            nomeDaAba,
            item.content,
            respostasTexto
          ]);

          // --- B. FAXINA ---
          var id = item.commentId || item.id;
          
          if (item.status === 'deleted') return;

          try {
            Drive.Comments.remove(fileId, id);
            deletados++;
          } catch (eDel) {
            if (item.status === 'open') {
              try {
                Drive.Comments.patch({status: 'resolved'}, fileId, id);
                resolvidos++;
              } catch (eResolve) {
                erros++;
              }
            } else {
              erros++;
            }
          }
        });
      }
      pageToken = resp.nextPageToken;
    } catch (e) {
      ui.alert("Erro Geral: " + e.message);
      return;
    }
  } while (pageToken);

  if (dadosBackup.length > 0) {
    sheetRep.getRange(2, 1, dadosBackup.length, 6).setValues(dadosBackup);
    sheetRep.autoResizeColumns(1, 6);
  }

  var msg = "FIM DO PROCESSO:\n" +
            "✅ Deletados permanentemente: " + deletados + "\n" +
            "👀 Ocultados (Resolvidos): " + resolvidos + "\n" +
            "❌ Falhas (Sem permissão): " + erros;
  ui.alert(msg);
}

function gerarRelatorioPermissoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  var file = DriveApp.getFileById(fileId);
  
  var sheetName = "Relatório de Permissões";
  var reportSheet = ss.getSheetByName(sheetName);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(sheetName);
  } else {
    reportSheet.clear();
  }

  var dadosGerais = [["--- ACESSO GERAL AO ARQUIVO ---", "", ""]];
  dadosGerais.push(["Nome", "Email", "Nível de Acesso"]);

  try {
    var owner = file.getOwner();
    dadosGerais.push([owner.getName(), owner.getEmail(), "PROPRIETÁRIO"]);
  } catch (e) {
    dadosGerais.push(["N/A", "N/A", "Erro ao ler proprietário"]);
  }

  var editors = file.getEditors();
  editors.forEach(function(u) {
    dadosGerais.push([u.getName(), u.getEmail(), "EDITOR"]);
  });

  var viewers = file.getViewers();
  viewers.forEach(function(u) {
    dadosGerais.push([u.getName(), u.getEmail(), "LEITOR"]);
  });

  reportSheet.getRange(1, 1, dadosGerais.length, 3).setValues(dadosGerais);
  reportSheet.getRange("A1:C1").setBackground("#f3f3f3").setFontWeight("bold").merge();
  reportSheet.getRange("A2:C2").setFontWeight("bold");

  var startRow = dadosGerais.length + 3;
  var dadosProtecao = [["--- INTERVALOS E ABAS PROTEGIDAS ---", "", "", ""]];
  dadosProtecao.push(["Descrição / Intervalo", "Tipo", "Aba", "Quem PODE Editar (além do dono)"]);

  var protecoes = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var protecoesSheet = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var todasProtecoes = protecoes.concat(protecoesSheet);

  if (todasProtecoes.length > 0) {
    todasProtecoes.forEach(function(p) {
      var rangeName = p.getRangeName();
      var rangeA1, sheetName, description, type, allowedEditors;

      try {
        rangeA1 = p.getRange().getA1Notation();
        sheetName = p.getRange().getSheet().getName();
        description = p.getDescription();
        type = (p.getRangeName() || description) ? "Intervalo" : "Aba Inteira";
      } catch (e) {
        rangeA1 = "Erro"; sheetName = "Erro"; description = "Erro"; type = "Erro";
      }

      try {
          allowedEditors = p.getEditors().map(function(u) { return u.getEmail(); }).join(", ");
          if (allowedEditors === "") allowedEditors = "Apenas Proprietário";
      } catch (e) {
          allowedEditors = "⚠️ SEM PERMISSÃO"; 
      }

      dadosProtecao.push([description || rangeA1, type, sheetName, allowedEditors]);
    });
  } else {
    dadosProtecao.push(["Nenhuma proteção específica encontrada", "-", "-", "-"]);
  }

  reportSheet.getRange(startRow, 1, dadosProtecao.length, 4).setValues(dadosProtecao);
  reportSheet.getRange(startRow, 1, 1, 4).setBackground("#f3f3f3").setFontWeight("bold").merge();
  reportSheet.getRange(startRow + 1, 1, 1, 4).setFontWeight("bold");
  reportSheet.autoResizeColumns(1, 4);
}

function fixarFormatoMatriculaGeral() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  
  var abasIgnoradas = [
    'DASHBOARD', 'ANALISE_DINAMICA', 'COLETA FINAL', 'COLETA', 'UTIL',
    'BASE', 'TABELAS', 'CENTRO OBSTETRICO', 'FIDELIZAÇÃO SALAS',
    'RESERVAS', '📅 CONSOLIDADO', 'CONSOLIDADO', 'CANCELADAS', 'RELATÓRIO_FINAL'
  ];

  var contagem = 0;

  sheets.forEach(function(sheet) {
    var nome = sheet.getName();
    var deveIgnorar = abasIgnoradas.some(function(ignorada) {
      return ignorada.toUpperCase() === nome.toUpperCase() || nome.indexOf("RELATÓRIO") > -1;
    });

    if (!deveIgnorar) {
      var lastRow = sheet.getMaxRows();
      var range = sheet.getRange(2, 14, lastRow - 1, 1);
      range.setNumberFormat("@");
      contagem++;
    }
  });

  ui.alert('✅ Processo Concluído!\n\nA formatação de Texto foi aplicada na coluna N de ' + contagem + ' abas de datas.');
}

// =====================================================================
// 🧹 FERRAMENTA DE PADRONIZAÇÃO DE NOMES DE SALAS (AÇÃO ÚNICA)
// =====================================================================

function padronizarNomesDeSalas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var contador = 0;

  // Mapa de Substituição (Termo Antigo -> Termo Novo)
  // Baseado na sua lista. Nota: "SALA 8" apareceu duas vezes no seu pedido 
  // (uma para CCA e outra para HMD). Mantive a lógica de:
  // - Números soltos ou "SALA 0X" genéricos -> HMD
  // - "SALA X" ou "SALA 2X" -> CCA
  // Ajuste conforme a realidade se necessário.
  
  var mapaDePara = {
    // --- HMD (Hemodinâmica) ---
    "SALA 01 CARDIOLOGIA": "SALA 01 HMD CARDIOLOGIA",
    "SALA 02 CARDIOLOGIA": "SALA 02 HMD CARDIOLOGIA",
    "SALA 02 ELETROFISIOLOGIA": "SALA 02 HMD ELETROFISIOLOGIA",
    "SALA 03 VASCULAR": "SALA 03 HMD VASCULAR",
    "SALA 03 RADIOLOGIA": "SALA 03 HMD RADIOLOGIA",
    "SALA 03 NEURO": "SALA 03 NEURO HMD",
    
    // Genéricos HMD (Assumidos)
    "SALA 01": "SALA 01 HMD",
    "SALA 1": "SALA 1 HMD",
    "1": "SALA 01 HMD",
    "8": "SALA 08 HMD",
    "SALA 8": "SALA 08 HMD", // Priorizei HMD pois você pediu "8 -> SALA 08 HMD" também
    
    // --- CCA (Ambulatório) ---
    "SALA 22": "SALA 22 CCA",
    "SALA 24": "SALA 24 CCA",
    "24": "SALA 24 CCA",
    "SALA 25": "SALA 25 CCA",
    "25": "SALA 25 CCA",
    "SALA 21": "SALA 21 CCA",
    "21": "SALA 21 CCA",
    "SALA 7": "SALA 7 CCA",
    "7": "SALA 7 CCA",
    "SALA 08": "SALA 08 CCA", // "SALA 08" explícito vai para CCA
    "SALA 8 CCA": "SALA 8 CCA" // Caso já exista, mantém
  };

  // Loop em todas as abas
  sheets.forEach(function(sheet) {
    var nomeAba = sheet.getName();
    // Ignora abas de sistema/dashboard para evitar acidentes
    if (nomeAba === "BASE" || nomeAba === "DASHBOARD" || nomeAba === "ConfigPermissoes") return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    // Pega apenas a Coluna A (onde ficam os nomes das salas)
    // Ler tudo de uma vez é muito mais rápido que ler linha a linha
    var range = sheet.getRange(1, 1, lastRow, 1);
    var valores = range.getValues();
    var alterouAba = false;

    for (var i = 0; i < valores.length; i++) {
      var valorAtual = String(valores[i][0]).trim(); // Remove espaços extras
      
      // Verifica se o valor exato existe no mapa
      // .toUpperCase() garante que pegue "sala 01" e "SALA 01"
      if (valorAtual !== "" && mapaDePara.hasOwnProperty(valorAtual.toUpperCase())) {
        valores[i][0] = mapaDePara[valorAtual.toUpperCase()];
        contador++;
        alterouAba = true;
      }
    }

    // Se houve mudança, grava de volta na planilha
    if (alterouAba) {
      range.setValues(valores);
      Logger.log("Atualizado na aba: " + nomeAba);
    }
  });

  var ui = SpreadsheetApp.getUi();
  ui.alert('✅ Padronização Concluída!', 
    'Foram atualizados ' + contador + ' nomes de salas em toda a planilha.\n\n' +
    'Verifique se as salas "SALA 8" (que foram para HMD) e "SALA 08" (que foram para CCA) estão corretas.', 
    ui.ButtonSet.OK);
}
