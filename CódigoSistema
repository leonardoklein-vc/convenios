// =====================================================================
// 🩺 SISTEMA DE AGENDAMENTO HCPA – CORE (Núcleo)
// ARQUIVO: CodigoSistema.gs
// Responsável por: Menus, Painel Admin, Visão do Dia, Edições Rápidas e Geração de Dias
// =====================================================================

// --- LISTA GLOBAL DE ABAS IGNORADAS ---
var ABAS_IGNORADAS = [
  'DASHBOARD', 'ANALISE_DINAMICA', 'COLETA FINAL', 'COLETA', 'UTIL',
  'BASE', 'TABELAS', 'CENTRO OBSTETRICO', 'FIDELIZAÇÃO SALAS',
  'RESERVAS', '📅 CONSOLIDADO', 'CONSOLIDADO', 'CANCELADAS', 'RELATÓRIO_FINAL',
  'ConfigPermissoes'
];

// ID DA PLANILHA DE MODELOS (PADRÃO GESTÃO CONVÊNIOS)
var ID_PLANILHA_MODELOS = "146FaarSV_tvOmlnf9Jrhydo_wqFgEf48MYIijtw-ozo";
// ID DA PASTA RAIZ DE ANEXOS (DRIVE)
var ID_PASTA_ANEXOS_RAIZ = "1fFxJP6ETIwq3Bhg0h3FkiAPOkUHxVwxB";

// =====================================================================
// --- 1. MENUS E INICIALIZAÇÃO ---
// =====================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🏥 Gestão');
  
  // 1. Gestão principal
  menu.addItem('📋 Gestão Convênios', 'abrirVisaoDia'); 
  menu.addSeparator();
  
  // 2. Agendamento (Chama funções do AgendamentoScript.gs ou locais wrappers)
  menu.addItem('➕ Novo Agendamento (Modal)', 'wrapperMostrarViewNovo'); 
  menu.addItem('➕ Novo Agendamento (Lateral)', 'wrapperAbrirSidebar');
  menu.addItem('🔄 Reagendar (Seleção)', 'wrapperAbrirReagendar'); 
  menu.addItem('❌ Cancelar (Seleção)', 'wrapperAbrirCancelar');   
  menu.addSeparator();
  
  // 3. Ferramentas
  menu.addItem('📝 Inserir Nota (Cor)', 'ferramentaInserirNota');
  menu.addItem('📅 Ir para Data', 'ferramentaIrParaData');
  menu.addSeparator();
  menu.addItem('📱 Ferramenta Checklist (WhatsApp)', 'abrirFerramentaChecklist'); // <--- NOVO ITEM

  // 4. Configurações
  menu.addSeparator();
  menu.addItem('⚙️ Configurações', 'abrirConfiguracoesAdmin');
  menu.addItem('⚙️ Configurar Permissões', 'abrirConfiguradorPermissoes'); // <--- NOVO
  menu.addItem('🏗️ Gerenciar Salas/Áreas', 'abrirFerramentaEstrutura'); // <--- ADICIONE ISTO
  menu.addSeparator();
  menu.addItem('🖌️ Gerar Coluna AC (Lote)', 'abrirFerramentaFormatacaoAC');
  
  menu.addToUi();
}

// --- WRAPPERS PARA FUNÇÕES DO AGENDAMENTOSCRIPT.GS ---
function wrapperMostrarViewNovo() {
  if (typeof abrirFormulario === 'function') abrirFormulario();
  else SpreadsheetApp.getUi().alert("Função abrirFormulario não encontrada.");
}
function wrapperAbrirSidebar() {
  if (typeof abrirFormularioLateral === 'function') abrirFormularioLateral();
  else SpreadsheetApp.getUi().alert("Função abrirFormularioLateral não encontrada.");
}
function wrapperAbrirReagendar() {
  if (typeof abrirReagendamento === 'function') abrirReagendamento();
  else SpreadsheetApp.getUi().alert("Função abrirReagendamento não encontrada.");
}
function wrapperAbrirCancelar() {
  Logger.log("INÍCIO: wrapperAbrirCancelar acionado pelo menu.");
  try {
    if (typeof cancelarAgendamento === 'function') {
      Logger.log("-> Função cancelarAgendamento encontrada. Transferindo comando...");
      cancelarAgendamento();
      Logger.log("FIM: wrapperAbrirCancelar (Sucesso)");
    } else {
      Logger.log("-> ERRO: Função cancelarAgendamento NÃO encontrada no escopo global.");
      SpreadsheetApp.getUi().alert("Função cancelarAgendamento não encontrada.");
    }
  } catch (e) {
    Logger.log("-> ERRO CRÍTICO no wrapperAbrirCancelar: " + e.message + "\nStack: " + e.stack);
    SpreadsheetApp.getUi().alert("Erro Crítico no Wrapper: " + e.message);
  }
}

// =====================================================================
// --- 2. CONFIGURAÇÕES ADMIN (PAINEL GERAL) ---
// =====================================================================

function abrirConfiguracoesAdmin() {
  var userEmail = Session.getActiveUser().getEmail();
  
  if (typeof ehSuperAdmin === 'function' && !ehSuperAdmin()) {
    SpreadsheetApp.getUi().alert('⛔ Acesso Restrito\n\nEsta área é exclusiva para Super Administradores.');
    return;
  }

  var permissoesAtuais = (typeof obterListasDePermissao === 'function') ? obterListasDePermissao() : { leitores:[], guiche:[], autorizacao:[], admin:[] };
  var coresAtuais = obterPreferenciasCores();     
  
  var template = HtmlService.createTemplateFromFile('ConfiguracoesAdminHTML');
  template.dadosIniciais = JSON.stringify(permissoesAtuais);
  template.cores = JSON.stringify(coresAtuais);
  
  var html = template.evaluate()
      .setWidth(1000)
      .setHeight(700)
      .setTitle('⚙️ Painel Administrativo');
      
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ Painel Administrativo');
}

function salvarItemNaBaseExterna(dados) {
  if (typeof ehSuperAdmin === 'function' && !ehSuperAdmin()) return "⛔ Sem permissão.";
  
  var idPlanilha = "10u3xcFlOtESIHGkR74lPpJ6fkj6zh3MS-mDVvJdIjp0";
  var nomeGuia = "Base";
  
  var mapaColunas = {
    'procedimento': 1, 'equipamento': 2, 'material': 3, 'exame': 4,
    'convenio': 5, 'equipe': 6, 'anestesista': 6, 'origem': 7,
    'lateralidade': 8, 'ia': 9, 'marcacao': 10, 'status': 11
  };
  
  var colIndex = mapaColunas[dados.tipo];
  if (!colIndex) return "Erro: Tipo de dado não mapeado.";
  
  try {
    var ss = SpreadsheetApp.openById(idPlanilha);
    var sheet = ss.getSheetByName(nomeGuia);
    if (!sheet) return "Erro: Guia 'Base' não encontrada.";
    
    var maxRows = sheet.getMaxRows();
    var valores = sheet.getRange(1, colIndex, maxRows, 1).getValues();
    var proximaLinha = 0;
    
    for (var i = 0; i < valores.length; i++) {
      if (valores[i][0] === "") {
        proximaLinha = i + 1;
        break;
      }
    }
    if (proximaLinha === 0) proximaLinha = maxRows + 1;
    
    // 🔥 ALTERADO AQUI: Removido .toUpperCase() para respeitar a formatação do HTML
    sheet.getRange(proximaLinha, colIndex).setValue(dados.valor);
    
    return "Sucesso";
  } catch(e) {
    return "Erro ao salvar: " + e.message;
  }
}

// =====================================================================
// --- 3. VISÃO GERAL DO DIA (FRONTEND PRINCIPAL) ---
// =====================================================================

function abrirVisaoDia() {
  DriveApp.getRootFolder();
  var html = HtmlService.createTemplateFromFile('VisaoDiaHTML');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var nomeAba = sheet.getName();
  
  if (!ehAbaDeData(nomeAba)) nomeAba = "";
  
  html.abaInicial = nomeAba;
  
  // Carrega listas de dropdown
  if (typeof obterListasValidacao === 'function') {
      html.listas = JSON.stringify(obterListasValidacao(ss));
  } else {
      html.listas = "{}";
  }
  
  // --- NOVA LÓGICA DE IDENTIDADE E GRUPOS ---
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var userGroup = "Visitante (Sem Permissão)"; // Padrão
  var isAdm = false;

  // 1. Tenta identificar Super Admin
  if (typeof ehSuperAdmin === 'function' && ehSuperAdmin()) {
      userGroup = "💎 Super Admin";
      isAdm = true;
  } 
  // 2. Se não for Super, verifica nas listas da planilha
  else if (typeof obterListasDePermissao === 'function') {
      var perms = obterListasDePermissao();
      
      // Verifica cada lista (ordem de prioridade)
      if (perms.admin && perms.admin.indexOf(userEmail) > -1) {
          userGroup = "🛠️ Admin / Agendamento";
          isAdm = true;
      } else if (perms.autorizacao && perms.autorizacao.indexOf(userEmail) > -1) {
          userGroup = "📋 Autorização";
      } else if (perms.guiche && perms.guiche.indexOf(userEmail) > -1) {
          userGroup = "🛎️ Guichê / Recepção";
      } else if (perms.leitores && perms.leitores.indexOf(userEmail) > -1) {
          userGroup = "👀 Leitor";
      }
  }

  // Se o email vier vazio (erro comum no GAS), avisa
  if (!userEmail) userEmail = "Desconhecido (Logue no Google)";

  // Passa as variáveis para o HTML
  html.isUserAdmin = isAdm;
  html.displayUser = userEmail;
  html.displayGroup = userGroup;
  // ------------------------------------------------
  
  // Carrega disponibilidade se houver aba selecionada
  if (nomeAba) {
      if (typeof lerDisponibilidade === 'function') {
          html.dispInicial = JSON.stringify(lerDisponibilidade(sheet));
      } else {
          html.dispInicial = "{}";
      }
  } else {
      html.dispInicial = "{}";
  }

  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1200).setHeight(850), '👁️ Visão Geral do Dia');
}

// --- LISTAGEM DE DADOS BLINDADA ---
function listarCirurgiasDoDia(nomeAba) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "cir_dia_" + nomeAba;
  var cachedData = cache.get(cacheKey);

  // 1. TENTA SERVIR PELO CACHE (Sem tocar na API do Sheets)
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  // 2. SE NÃO TEM CACHE, ORGANIZA A FILA
  var lock = LockService.getScriptLock();
  try {
    // Aguarda até 15 segundos na fila para ler a planilha
    lock.waitLock(15000); 
    
    // Checa o cache novamente (alguém pode ter populado enquanto estávamos na fila)
    cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) return { erro: "Aba não encontrada." };

    var lastRow = sheet.getLastRow();
    var startRow = 5; 
    if (lastRow < startRow) return { cirurgias: [] };

    var maxColumns = sheet.getMaxColumns();
    var limitCol = (maxColumns < 30) ? maxColumns : 30; 
    var numRows = lastRow - startRow + 1;
    var range = sheet.getRange(startRow, 1, numRows, limitCol); 
    
    var valores = range.getDisplayValues();
    var cores = sheet.getRange(startRow, 3, numRows, 1).getBackgrounds();
    var notas = range.getNotes(); 
    var pesosFonteOpme = sheet.getRange(startRow, 15, numRows, 1).getFontWeights();

    var coresGuiche = [];
    if (maxColumns >= 27) {
        var colsToRead = (maxColumns >= 28) ? 2 : 1;
        coresGuiche = sheet.getRange(startRow, 27, numRows, colsToRead).getBackgrounds();
    }

    var cirurgias = [];
    var areaAtual = "BLOCO"; 

    var getSalaMescladaLocal = (typeof getSalaMesclada === 'function') ? getSalaMesclada : function(s, r){ return ""; };

    for (var i = 0; i < valores.length; i++) {
      var linhaReal = i + startRow;
      var row = valores[i];
      
      var paciente = row[2]; 
      var sala = row[0];     
      var hora = row[1];     
      var cor = (cores[i][0] || "#ffffff").toLowerCase();
      var isOpmeBold = (pesosFonteOpme[i][0] === 'bold');

      var bgAA = (coresGuiche[i] && coresGuiche[i][0]) ? coresGuiche[i][0].toLowerCase() : '#ffffff';
      var bgAB = (coresGuiche[i] && coresGuiche[i][1]) ? coresGuiche[i][1].toLowerCase() : '#ffffff';
      
      var statusGuiche = 'vazio';
      var tonsVermelho = ['#ff0000', '#ea4335', '#e06666', '#f4c7c3', '#ff9900'];
      var tonsAmarelo = ['#ffff00', '#fff2cc', '#f1c232', '#bf9000'];
      var tonsCinza = ['#d9d9d9', '#cccccc', '#efefef', '#999999', '#b7b7b7', '#d0e0e3', '#eeeeee'];

      if (tonsVermelho.includes(bgAA) || tonsVermelho.includes(bgAB)) statusGuiche = 'erro';
      else if (tonsAmarelo.includes(bgAA) || tonsAmarelo.includes(bgAB)) statusGuiche = 'atencao'; 
      else if (tonsCinza.includes(bgAA) || tonsCinza.includes(bgAB)) statusGuiche = 'ok';       

      var agendaClean = String(row[18] || "").replace(/\D/g, ''); 
      var codigoPaciente = row[28] || ""; // 🔥 LÊ A COLUNA AC (Índice 28)

      // Localize dentro de listarCirurgiasDoDia o bloco de detecção de área:
if ((!paciente || paciente === "") && sala && sala !== "") {
    var sUp = sala.toUpperCase();
    // Prioridade total para CCA para não confundir com CCV
    if (/\bCCA\b/.test(sUp)) areaAtual = "CCA"; 
    else if (/\bCCV\b/.test(sUp)) areaAtual = "CCV";
    else if (/\bHEMO\b/.test(sUp)) areaAtual = "HEMODINÂMICA";
    else if (/\bBC\b|\bBLOCO\b/.test(sUp)) areaAtual = "BLOCO";
}

      if (paciente && paciente !== "") {
        var blocos = 1;
        for (var j = i + 1; j < valores.length; j++) {
          var corProx = (cores[j][0] || "#ffffff").toLowerCase();
          var pacProx = valores[j][2];
          if (corProx === cor && (!pacProx || pacProx === "")) blocos++;
          else break;
        }
        
        var rowNotes = notas[i];
        var notasObj = {};
        var mapCols = {
           2: 'paciente', 3: 'nascimento', 4: 'procedimento', 5: 'cid', 6: 'lado',
           7: 'equipe', 8: 'origem', 9: 'marcacao', 10: 'anestesia', 11: 'ia',
           12: 'convenio', 13: 'matricula', 14: 'opme', 15: 'materiais',
           16: 'equipamentos', 17: 'exames', 19: 'status', 22: 'info',
           23: 'telefone', 24: 'cti'
        };
        for (var cIdx in mapCols) {
           if (rowNotes[cIdx] && rowNotes[cIdx].trim() !== "") {
              notasObj[mapCols[cIdx]] = rowNotes[cIdx];
           }
        }

        cirurgias.push({
          id: linhaReal,
          area: areaAtual,
          sala: sala || getSalaMescladaLocal(sheet, linhaReal),
          hora: hora,
          paciente: paciente,
          nasc: row[3],
          procedimento: row[4],
          cid: row[5],
          lado: row[6],
          equipe: row[7],
          origem: row[8],
          marcacao: row[9],
          anestesia: row[10],
          ia: row[11],
          convenio: row[12],
          matricula: row[13],
          opme: row[14],
          opmeBold: isOpmeBold,
          codigoPaciente: codigoPaciente, // 🔥 DADO ENVIADO PARA O FRONTEND 
          statusGuiche: statusGuiche, 
          instrumentais: row[15],
          equipamentos: row[16],
          exames: row[17],
          agenda: agendaClean, 
          status: row[19],
          motivo: row[20],
          revisao: row[21],
          info: row[22],
          telefone: row[23],
          cti: row[24],
          aut: row[25],
          email: (row.length > 26) ? row[26] : "", 
          obs: (row.length > 27) ? row[27] : "",
          anexos: (row.length > 29) ? row[29] : "",   
          cor: cor,
          duracao: blocos * 30,
          notas: notasObj 
        });
      }
    }

    var resultadoFinal = { cirurgias: cirurgias, total: cirurgias.length };
    
    // 3. GRAVA NO CACHE POR 20 SEGUNDOS
    // Isso garante que picos de acesso sejam absorvidos sem travar o Google Sheets
    cache.put(cacheKey, JSON.stringify(resultadoFinal), 20);
    
    return resultadoFinal;

  } catch (e) {
    // Fallback amigável se a fila estourar
    return { erro: "O sistema está processando muitos acessos simultâneos. A página recarregará em instantes.", cirurgias: [] };
  } finally {
    // 4. SEMPRE LIBERA A FILA
    lock.releaseLock();
  }
}

// =====================================================================
// --- 4. FUNÇÕES DE EDIÇÃO (BACKEND) BLINDADAS ---
// =====================================================================

// Helper interno para gerenciar a fila e limpar o cache
function executarComTrava(guia, acaoParaExecutar) {
  var lock = LockService.getScriptLock();
  try {
    // Aguarda na fila por até 10 segundos
    lock.waitLock(10000); 
    
    // Executa a ação específica que foi passada
    var resultado = acaoParaExecutar();
    
    // Invalida o cache APENAS se a ação foi bem-sucedida
    if (resultado && resultado.success) {
      CacheService.getScriptCache().remove("cir_dia_" + guia);
    }
    return resultado;
  } catch (e) {
    return { success: false, message: "Aguarde um momento, sistema ocupado. Detalhe: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// Edição rápida (Email/Obs/Checks)
function salvarEdicaoRapida(form) {
  if (typeof podeEditarChecks === 'function' && !podeEditarChecks()) {
      return { success: false, message: "⛔ Não autorizado." };
  }
  return executarComTrava(form.guia, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(form.guia);
    var linha = parseInt(form.id);
    if (!sheet || isNaN(linha)) throw new Error("Referência perdida.");
    
    if (sheet.getMaxColumns() < 28) sheet.insertColumnsAfter(sheet.getMaxColumns(), 28 - sheet.getMaxColumns());
    sheet.getRange(linha, 27, 1, 2).setValues([[form.email, form.obs]]);
    return { success: true, message: "✅ Check 1 e 2 atualizados!" };
  });
}

// Edição de Autorização (AUT)
function salvarAut(guia, id, valor) {
  if (typeof podeEditarInfoEAut === 'function' && !podeEditarInfoEAut()) {
      return { success: false, message: "⛔ Não autorizado." };
  }
  return executarComTrava(guia, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    var linha = parseInt(id);
    if (!sheet || isNaN(linha)) throw new Error("Erro na referência da aba ou linha.");
    
    sheet.getRange(linha, 26).setValue(valor);
    return { success: true, message: "Status AUT alterado." };
  });
}

// Edição de Status Guichê (Badge)
function salvarStatusGuiche(guia, id, novoStatus) {
  if (typeof podeEditarChecks === 'function' && !podeEditarChecks()) {
      return { success: false, message: "⛔ Não autorizado." };
  }
  return executarComTrava(guia, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    var linha = parseInt(id);
    if (!sheet || isNaN(linha)) throw new Error("Referência perdida.");
    
    if (sheet.getMaxColumns() < 28) {
       sheet.insertColumnsAfter(sheet.getMaxColumns(), 28 - sheet.getMaxColumns());
    }
    
    var cor = '#ffffff';
    if (novoStatus === 'ok') cor = '#d9d9d9'; 
    else if (novoStatus === 'atencao') cor = '#ffff00';
    else if (novoStatus === 'erro') cor = '#ea4335';
    
    sheet.getRange(linha, 27, 1, 2).setBackground(cor);
    return { success: true };
  });
}

// Edição de Info Técnica
function salvarInfoTecnica(guia, id, texto) {
  if (typeof podeEditarInfoEAut === 'function' && !podeEditarInfoEAut()) {
      return { success: false, message: "⛔ Não autorizado." };
  }
  return executarComTrava(guia, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    var linha = parseInt(id);
    if (!sheet || isNaN(linha)) throw new Error("Erro de referência.");
    
    sheet.getRange(linha, 23).setValue(texto); 
    return { success: true, message: "Info salva!" };
  });
}

// Toggle OPME Bold
function alternarOpmeBold(guia, id, deveFicarNegrito) {
  // 🔥 CORREÇÃO: Liberado para grupo Autorização também (podeEditarInfoEAut)
  if (typeof podeEditarInfoEAut === 'function' && !podeEditarInfoEAut()) {
      return { success: false, message: "⛔ Não autorizado. Apenas Autorização ou Admin." };
  }
  return executarComTrava(guia, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    var linha = parseInt(id);
    if (!sheet || isNaN(linha)) throw new Error("Erro de referência.");
    
    var cell = sheet.getRange(linha, 15); // Coluna O (15) = OPME
    
    if (deveFicarNegrito) cell.setFontWeight("bold");
    else cell.setFontWeight("normal");
    
    return { success: true };
  });
}

// Carrega dados para edição/reagendamento (Backend)
function carregarDadosParaReagendamento(guia, id) {
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      return { error: "⛔ Não autorizado." };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    if(!sheet) throw new Error("Aba '" + guia + "' não encontrada.");
    
    // lerDadosDaLinha está no AgendamentoScript, mas é acessível globalmente
    // Se não estiver, precisaremos copiá-la para cá. 
    // Como estão no mesmo projeto, deve funcionar. 
    // Por segurança, verificamos:
    if (typeof lerDadosDaLinha === 'function') {
        return lerDadosDaLinha(sheet, parseInt(id));
    } else {
        throw new Error("Função lerDadosDaLinha não encontrada. Verifique se AgendamentoScript.gs está salvo.");
    }
  } catch(e) {
    return { error: e.message };
  }
}

// =====================================================================
// 🔥 NOVA FUNÇÃO: ATUALIZAÇÃO COMPLETA DE DADOS (EDICAO TUDO)
// =====================================================================
function atualizarAgendamentoCompleto(form) {
  // --- 🔒 TRAVA DE SEGURANÇA OBRIGATÓRIA ---
  // Se a função de verificação não existir ou retornar falso, bloqueia tudo.
  if (typeof podeAlterarAgenda !== 'function' || !podeAlterarAgenda()) {
      return { success: false, message: "⛔ ACESSO NEGADO: Você não tem permissão de Administrador para editar estes dados." };
  }
  // -------------------------------------------
  
  var guiaAlvo = form.guiaEdicao || form.guia;
  
  return executarComTrava(guiaAlvo, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guiaAlvo);
    if (!sheet) throw new Error("Aba não encontrada.");

    var linha = parseInt(form.idEdicao);
    if (isNaN(linha)) throw new Error("ID inválido.");

    // Mapeamento dos campos do formulário para as colunas
    // Coluna 1 (A) = Sala, 2 (B) = Hora -> NÃO MEXEMOS AQUI NO EDITAR SIMPLES
    
    // Helper para juntar arrays (select multiplo)
    var joinArr = function(v) { return Array.isArray(v) ? v.join(", ") : v; };

    // Atualiza Colunas C até Y (3 a 25)
    // C(3)=Paciente, D(4)=Nasc, E(5)=Proc ...
    
    var dadosSalvar = [[
        form.paciente,             // 3 (C)
        form.nascimento,           // 4 (D)
        joinArr(form.procedimentos),// 5 (E)
        form.cid,                  // 6 (F)
        form.lado,                 // 7 (G)
        form.equipe,               // 8 (H)
        form.origem,               // 9 (I)
        form.marcadoPor,           // 10 (J)
        form.anestesista,          // 11 (K)
        form.ia,                   // 12 (L)
        form.convenio,             // 13 (M)
        form.matricula,            // 14 (N)
        form.opme,                 // 15 (O)
        joinArr(form.materiais),   // 16 (P)
        joinArr(form.equipamentos),// 17 (Q)
        joinArr(form.exames),      // 18 (R)
        "",                        // 19 (S) - ID/Agenda (Não mexer ou manter) -> ideal ler antes, mas vamos pular por segurança visual
        form.status,               // 20 (T)
        "",                        // 21 (U) - Motivo Canc (Só usado em cancelar)
        "",                        // 22 (V) - Revisão
        form.outrasInfo,           // 23 (W)
        form.telefone,             // 24 (X)
        form.cti                   // 25 (Y)
    ]];

    // Nota: A coluna S (19) geralmente é gerada auto. Se quiser manter o valor antigo,
    // o ideal seria ler, mas para simplificar, se ela for fórmula, não sobrescreva.
    // O código abaixo grava do C ao R (3 ao 18) e depois do T ao Y (20 ao 25), pulando a S.
    
    // Bloco 1: C até R (3 a 18) -> 16 colunas
    var parte1 = dadosSalvar[0].slice(0, 16);
    sheet.getRange(linha, 3, 1, 16).setValues([parte1]);

    // Bloco 2: T até Y (20 a 25) -> 6 colunas
    var parte2 = dadosSalvar[0].slice(17, 23); // Pega do form.status até form.cti
    sheet.getRange(linha, 20, 1, 6).setValues([parte2]);
    
    // 🔥 LÓGICA DO CÓDIGO (COLUNA AC = 29) E NOTA INTELIGENTE NA C (3)
    var codPac = form.codigoPaciente || "";
    var celCodigo = sheet.getRange(linha, 29); // Coluna AC
    celCodigo.setValue(codPac);
    
    var celNome = sheet.getRange(linha, 3); // Coluna C
    var notaAtual = celNome.getNote() || "";
    var regexCodigo = /Código do Paciente:\s*[^\n]*/g; 

    if (codPac !== "") {
        if (regexCodigo.test(notaAtual)) {
            celNome.setNote(notaAtual.replace(regexCodigo, "Código do Paciente: " + codPac));
        } else {
            var quebra = (notaAtual !== "") ? "\n" : "";
            celNome.setNote(notaAtual + quebra + "Código do Paciente: " + codPac);
        }
    } else {
        if (regexCodigo.test(notaAtual)) {
            var notaLimpa = notaAtual.replace(regexCodigo, "").trim();
            if (notaLimpa === "") celNome.clearNote();
            else celNome.setNote(notaLimpa);
        }
    }

    return { success: true };
  });
}

// =====================================================================
// --- 5. FUNÇÕES DE GERAÇÃO DE DIAS ---
// =====================================================================

function obterUltimaDataFormatada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var maiorData = null;
  var anoAtual = new Date().getFullYear();

  for (var i = 0; i < sheets.length; i++) {
    var nome = sheets[i].getName().toUpperCase();
    var match = nome.match(/(\d{1,2})\/(\d{1,2})$/);
    
    if (match && ehAbaDeData(nome)) {
      var dia = parseInt(match[1]);
      var mes = parseInt(match[2]) - 1; 
      var dataObj = new Date(anoAtual, mes, dia, 12, 0, 0); 
      if (!maiorData || dataObj > maiorData) maiorData = dataObj;
    }
  }
  
  if (!maiorData) return "Nenhuma data padrão encontrada.";
  return Utilities.formatDate(maiorData, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function adicionarNovosDias(qtdDias) {
  if (typeof ehSuperAdmin === 'function' && !ehSuperAdmin()) return { success: false, msg: "Sem permissão." };
  
  try {
    var ssAtual = SpreadsheetApp.getActiveSpreadsheet();
    
    var ultimaData = null;
    var anoAtual = new Date().getFullYear();
    var sheets = ssAtual.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var nome = sheets[i].getName().toUpperCase();
      var match = nome.match(/(\d{1,2})\/(\d{1,2})$/);
      if (match && ehAbaDeData(nome)) {
        var d = new Date(anoAtual, parseInt(match[2])-1, parseInt(match[1]), 12,0,0);
        if (!ultimaData || d > ultimaData) ultimaData = d;
      }
    }
    
    if (!ultimaData) return { success: false, msg: "Não consegui identificar a última data." };

    var ssModelos = SpreadsheetApp.openById(ID_PLANILHA_MODELOS);
    var novasAbasCriadas = [];
    
    for (var i = 1; i <= qtdDias; i++) {
      var proximaData = new Date(ultimaData);
      proximaData.setDate(ultimaData.getDate() + i);
      
      var diasSemana = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"];
      var nomeDia = diasSemana[proximaData.getDay()];
      
      var diaStr = ("0" + proximaData.getDate()).slice(-2);
      var mesStr = ("0" + (proximaData.getMonth() + 1)).slice(-2);
      var anoStr = proximaData.getFullYear();
      var nomeAbaFinal = nomeDia + " " + diaStr + "/" + mesStr; 
      
      if (ssAtual.getSheetByName(nomeAbaFinal)) continue; 
      
      var templateSheet = ssModelos.getSheetByName(nomeDia);
      if (!templateSheet) {
        if (nomeDia === "SÁBADO") templateSheet = ssModelos.getSheetByName("SABADO");
        if (nomeDia === "TERÇA") templateSheet = ssModelos.getSheetByName("TERCA");
      }
      
      if (!templateSheet) return { success: false, msg: "Modelo não encontrado para: " + nomeDia };
      
      var novaSheet = templateSheet.copyTo(ssAtual);
      novaSheet.setName(nomeAbaFinal);
      novaSheet.getRange("A1").setValue("AGENDAMENTO DE CIRURGIAS - " + diaStr + "/" + mesStr + "/" + anoStr);
      
      ssAtual.setActiveSheet(novaSheet);
      ssAtual.moveActiveSheet(ssAtual.getNumSheets());
      novasAbasCriadas.push(nomeAbaFinal);
    }
    
    return { success: true, msg: "Criadas: " + novasAbasCriadas.join(", ") };
    
  } catch (e) {
    return { success: false, msg: "Erro: " + e.message };
  }
}

// =====================================================================
// --- 6. UTILITÁRIOS GLOBAIS (RECUPERADOS!) ---
// =====================================================================

function ehAbaDeData(nomeAba) {
  if (!nomeAba) return false;
  var nome = nomeAba.toUpperCase();
  for (var i = 0; i < ABAS_IGNORADAS.length; i++) {
    if (ABAS_IGNORADAS[i].toUpperCase() === nome) return false;
  }
  if (nome.indexOf('BACKUP') > -1) return false;
  if (nome.indexOf('RELATÓRIO') > -1) return false;
  if (nome.indexOf('RELATORIO') > -1) return false;
  if (nome.indexOf('CANCELADA') > -1) return false;
  if (nome.indexOf('CONFIG') > -1) return false;
  return true;
}

// 🔥🔥 FUNÇÃO QUE ESTAVA FALTANDO E CAUSOU O ERRO 🔥🔥
function obterListaGuias() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var lista = [];
    for (var i = 0; i < sheets.length; i++) {
      var nome = sheets[i].getName();
      if (ehAbaDeData(nome)) {
        lista.push(nome);
      }
    }
    return lista;
  } catch (e) {
    Logger.log("Erro em obterListaGuias: " + e.message);
    throw new Error("Falha ao acessar planilha atual: " + e.message);
  }
}

// 🔥🔥 OUTRA FUNÇÃO NECESSÁRIA 🔥🔥
function obterDisponibilidadeDaGuia(nomeGuia) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomeGuia);
  if (!sheet) return null;
  
  // Chama a função lerDisponibilidade que já está aqui embaixo
  return lerDisponibilidade(sheet);
}

function irParaGuia(nome) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (s) { 
    s.activate(); 
    return true; 
  }
  return false;
}

function obterPreferenciasCores() {
  var props = PropertiesService.getUserProperties();
  return {
    principal: props.getProperty('corPrincipal') || '#d9d2e9',
    backup: props.getProperty('corBackup') || '#b4a7d6'
  };
}

function salvarPreferenciasCores(form) {
  if (typeof ehSuperAdmin === 'function' && !ehSuperAdmin()) throw new Error("Acesso Negado.");
  PropertiesService.getUserProperties().setProperties(form);
  return "Cores salvas!";
}

function ferramentaIrParaData() {
  var html = HtmlService.createHtmlOutputFromFile('NavegadorHTML').setTitle('Navegar').setWidth(300).setHeight(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function ferramentaInserirNota() {
  var html = HtmlService.createHtmlOutputFromFile('NotaHTML').setWidth(400).setHeight(320); 
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Inserir Nota e Cor');
}

function salvarNotaECor(texto, cor) {
  // CORREÇÃO: Trocamos .getActiveCell() por .getActiveRange()
  var cell = SpreadsheetApp.getActiveRange(); 
  
  if (!cell) {
    throw new Error("Nenhuma célula selecionada.");
  }

  // Define a nota (o Google aplica a nota na primeira célula do intervalo selecionado)
  cell.setNote(texto);

  // Lógica da cor
  if (texto && texto.trim() !== "") {
    // Se tiver texto, aplica a cor selecionada em TUDO o que estiver selecionado
    cell.setBackground(cor);
  } else {
    // Se o texto estiver vazio, limpa a nota
    cell.clearNote();
    // Opcional: Se quiser que ele tire a cor ao apagar a nota, descomente abaixo:
    // cell.setBackground('#ffffff'); 
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obterConteudoAjuda() {
  return HtmlService.createHtmlOutputFromFile('AjudaHTML').getContent();
}

// =====================================================================
// HELPER: LER DISPONIBILIDADE (ATUALIZADO PARA CIANO)
// =====================================================================
function lerDisponibilidade(sheet) {
  var lastRow = sheet.getLastRow();
  // Se não tiver dados suficientes (só cabeçalho), retorna vazio
  if (lastRow < 5) return {}; 

  // Pega da linha 5 até o fim
  // Coluna 1=Sala (A), 2=Hora (B), 3=Paciente (C)
  var rangeData = sheet.getRange(5, 1, lastRow - 4, 3).getDisplayValues();
  // Pega as cores da coluna C (Paciente)
  var rangeCores = sheet.getRange(5, 3, lastRow - 4, 1).getBackgrounds();

  var disponibilidade = {};
  var ultimaSala = "";

  for (var i = 0; i < rangeData.length; i++) {
    var sala = String(rangeData[i][0]).trim();
    var hora = String(rangeData[i][1]).trim();
    var paciente = rangeData[i][2]; 
    var corAtual = (rangeCores[i][0] || '#ffffff').toLowerCase();

    // Lógica de mesclagem visual (Sala e Hora)
    if (sala !== "") {
      ultimaSala = sala;
    } else if (ultimaSala !== "" && hora !== "") {
      sala = ultimaSala;
    }

    var isSemNome = (!paciente || String(paciente).trim() === "");

    // --- 🔥 ALTERAÇÃO AQUI: BRANCO OU CIANO SÃO LIVRES ---
    // #00ffff é o Ciano padrão do seletor de cores
    var isCorLivre = (corAtual === '#ffffff' || corAtual === '#00ffff');

    if (sala && hora && isSemNome && isCorLivre) {
      if (!disponibilidade[sala]) disponibilidade[sala] = [];
      // Ajusta o índice da linha (i + 5, pois começamos na linha 5)
      disponibilidade[sala].push({ hora: hora, linha: i + 5 });
    }
  }
  return disponibilidade;
}

// =====================================================================
// 7. SISTEMA DE ANEXOS (GOOGLE DRIVE) -> COLUNA AD (30)
// =====================================================================

function processarUploadDrive(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(dados.guia);
    var linha = parseInt(dados.id);
    
    // 1. Acessa/Cria Estrutura de Pastas
    var pastaRaiz = DriveApp.getFolderById(ID_PASTA_ANEXOS_RAIZ);
    
    // Nível 1: PRONTUARIO - NOME (Ex: 12345 - JOAO SILVA)
    var nomePastaN1 = (dados.matricula || "S_Matricula") + " - " + dados.paciente;
    var pastaN1 = getOrCreateSubfolder(pastaRaiz, nomePastaN1);
    
    // Nível 2: DATA - 5 PALAVRAS PROC - EQUIPE
    // Ex: 25/10 - COLECISTECTOMIA VIDEOLAPAROSCOPICA - EQUIPE A
    var procCurto = (dados.procedimento || "").split(" ").slice(0, 5).join(" ");
    // Limpa caracteres proibidos em nomes de pasta (/ ou \)
    var dataLimpa = dados.data.replace(/\//g, "-"); 
    var nomePastaN2 = dataLimpa + " - " + procCurto + " - " + (dados.equipe || "S_Equipe");
    
    var pastaN2 = getOrCreateSubfolder(pastaN1, nomePastaN2);
    
    // 2. Cria o Arquivo
    var contentType = dados.fileData.substring(5, dados.fileData.indexOf(';'));
    var bytes = Utilities.base64Decode(dados.fileData.substr(dados.fileData.indexOf('base64,')+7));
    var blob = Utilities.newBlob(bytes, contentType, dados.fileName);
    var arquivo = pastaN2.createFile(blob);
    
    // 3. Salva Link na Planilha (Coluna AD = 30)
    // Se a planilha não tiver 30 colunas, cria
    if (sheet.getMaxColumns() < 30) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 30 - sheet.getMaxColumns());
    }

    var cell = sheet.getRange(linha, 30); // 🔥 ESCREVE NA AD
    var valorAtual = cell.getValue();
    var novoRegistro = dados.fileName + "|" + arquivo.getUrl();
    
    // Adiciona ao existente (separado por quebra de linha)
    var valorFinal = valorAtual ? (valorAtual + "\n" + novoRegistro) : novoRegistro;
    cell.setValue(valorFinal);
    
    return { success: true, links: valorFinal };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Helper para não duplicar pastas
function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

// =====================================================================
// FUNÇÃO DE DELEÇÃO (ARQUIVO + REFERÊNCIA)
// =====================================================================

function excluirArquivoAnexo(guia, idRow, urlParaDeletar) {
  // 🔥 Segurança: Apenas quem pode alterar agenda pode deletar arquivos
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      return { success: false, error: "⛔ Sem permissão para deletar arquivos." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(guia);
    var cell = sheet.getRange(parseInt(idRow), 30); // Coluna AD (30)
    var valorAtual = cell.getValue();

    if (!valorAtual) return { success: false, error: "Célula já está vazia." };

    // 1. Tentar Deletar do Drive
    try {
        // Extrai o ID do arquivo a partir da URL padrão do Drive
        // Formato comum: https://drive.google.com/file/d/ID_DO_ARQUIVO/view...
        var match = urlParaDeletar.match(/\/d\/(.+?)\//);
        if (match && match[1]) {
            var fileId = match[1];
            var arquivo = DriveApp.getFileById(fileId);
            arquivo.setTrashed(true); // Move para lixeira (segurança)
        }
    } catch (e) {
        // Se der erro no Drive (ex: arquivo já não existe), seguimos para limpar a planilha
        Logger.log("Aviso: Arquivo não encontrado no Drive ou já deletado. Limpando apenas referência.");
    }

    // 2. Atualizar a Célula (Remove a linha que contém a URL)
    var linhas = valorAtual.split('\n');
    var novasLinhas = [];
    
    for (var i = 0; i < linhas.length; i++) {
        // Se a linha NÃO contém a URL que queremos deletar, mantemos ela
        if (linhas[i].indexOf(urlParaDeletar) === -1) {
            novasLinhas.push(linhas[i]);
        }
    }

    var textoFinal = novasLinhas.join('\n');
    cell.setValue(textoFinal);

    return { success: true, links: textoFinal };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ARQUIVO: CodigoSistema.gs

function abrirFerramentaChecklist() {
  // Cria o objeto Template
  var template = HtmlService.createTemplateFromFile('ChecklistHTML');
  
  // 1. Obtém a lista de guias de data de forma segura
  var guias = [];
  try {
    if (typeof obterListaGuias === 'function') {
        guias = obterListaGuias();
    }
  } catch (e) {
    Logger.log("Erro ao obter guias: " + e.message);
  }
  
  // Passa para o HTML (Garante que nunca seja null/undefined)
  template.listaGuias = JSON.stringify(guias || []);
  
  // 2. Tenta definir a guia inicial baseada na aba ativa
  var nomeAbaAtiva = "";
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    nomeAbaAtiva = sheet.getName();
  } catch (e) {}

  if (typeof ehAbaDeData === 'function' && ehAbaDeData(nomeAbaAtiva)) {
      template.guiaInicial = nomeAbaAtiva;
  } else {
      template.guiaInicial = (guias.length > 0) ? guias[0] : "";
  }

  // Cria a interface
  var html = template.evaluate()
      .setTitle('📱 Ferramenta WhatsApp')
      .setWidth(400)
      .setHeight(750);

  SpreadsheetApp.getUi().showModelessDialog(html, '📱 Ferramenta WhatsApp');
}

function CONSOLIDAR_DIAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var consolidado = [];
  
  // Lista de meses que você quer incluir (ex: 01, 02, 03... até 07)
  var mesesValidos = ["01", "02", "03", "04", "05", "06", "07"];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var mesAba = name.split('/')[1]; // Pega o MM do nome "DIA DD/MM"

    if (mesesValidos.indexOf(mesAba) !== -1) {
      var data = sheets[i].getRange("A3:AD" + sheets[i].getLastRow()).getValues();
      var dataAba = name.split(' ')[1]; // Pega o DD/MM

      for (var j = 0; j < data.length; j++) {
        if (data[j][0] !== "") { // Verifica se a coluna A não está vazia
          data[j].push(dataAba); // Adiciona a data no final
          consolidado.push(data[j]);
        }
      }
    }
  }
  return consolidado;
}

// =====================================================================
// VIGIA AUTOMÁTICO DE PLANILHA (GATILHO ONEDIT)
// =====================================================================
function onEdit(e) {
  if (!e || !e.range) return;
  
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();
  
  if (typeof ehAbaDeData === 'function' && !ehAbaDeData(sheet.getName())) return;
  
  // Se a edição foi na Coluna AC (29) abaixo do cabeçalho
  if (col === 29 && row >= 5) {
    var codPac = String(range.getValue()).trim();
    var celNome = sheet.getRange(row, 3); // Coluna C
    var nomePac = celNome.getValue();
    
    if (nomePac !== "") {
      var notaAtual = celNome.getNote() || "";
      var regexCodigo = /Código do Paciente:\s*[^\n]*/g;
      
      if (codPac !== "") {
          if (regexCodigo.test(notaAtual)) {
              celNome.setNote(notaAtual.replace(regexCodigo, "Código do Paciente: " + codPac));
          } else {
              var quebra = (notaAtual !== "") ? "\n" : "";
              celNome.setNote(notaAtual + quebra + "Código do Paciente: " + codPac);
          }
      } else {
          // Usuário deletou o código na Coluna AC, limpa só ele
          if (regexCodigo.test(notaAtual)) {
              var notaLimpa = notaAtual.replace(regexCodigo, "").replace(/^\s+|\s+$/g, '');
              if (notaLimpa === "") celNome.clearNote();
              else celNome.setNote(notaLimpa);
          }
      }
    }
  }
}

// =====================================================================
// ATUALIZADOR RETROATIVO DE NOTAS DE CÓDIGO
// =====================================================================
function forcarAtualizacaoDeNotasNaAbaAtual() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var maxRow = sheet.getLastRow();
  
  if (maxRow < 5) {
    SpreadsheetApp.getUi().alert("Aba sem dados suficientes.");
    return;
  }
  
  // Lê todos os nomes e códigos de uma vez para não travar o Google
  var nomes = sheet.getRange(5, 3, maxRow - 4, 1).getValues();
  var codigos = sheet.getRange(5, 29, maxRow - 4, 1).getDisplayValues();
  
  var atualizados = 0;
  
  for (var i = 0; i < nomes.length; i++) {
    var nome = String(nomes[i][0]).trim();
    var cod = String(codigos[i][0]).trim();
    var linhaReal = i + 5;
    
    var celNome = sheet.getRange(linhaReal, 3);
    
    if (nome !== "" && cod !== "") {
      celNome.setNote("Código do Paciente: " + cod);
      atualizados++;
    }
  }
  
  SpreadsheetApp.getUi().alert("✅ Varredura concluída! " + atualizados + " notas inseridas/atualizadas.");
}

// =====================================================================
// 🖌️ FERRAMENTA EM LOTE: FORMATAR COLUNA AC (CÓD. DO PACIENTE)
// =====================================================================

function abrirFerramentaFormatacaoAC() {
  if (typeof ehSuperAdmin === 'function' && !ehSuperAdmin()) {
    SpreadsheetApp.getUi().alert('⛔ Acesso Restrito\n\nApenas Administradores podem rodar ações em lote.');
    return;
  }
  
  var html = HtmlService.createTemplateFromFile('FormatacaoACHTML')
      .evaluate()
      .setTitle('🖌️ Criador da Coluna AC (Lote)')
      .setWidth(750)
      .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, '🖌️ Criador da Coluna AC');
}

// 1. Escaneia a planilha e classifica as abas
function listarAbasParaFormatacaoAC() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var lista = [];
    
    // Regex que busca os dias da semana seguidos de DD/MM
    var regexData = /(SEGUNDA|TERÇA|TERCA|QUARTA|QUINTA|SEXTA|SÁBADO|SABADO|DOMINGO)\s\d{2}\/\d{2}/i;

    sheets.forEach(function(s) {
      var nome = s.getName();
      if (regexData.test(nome) && ABAS_IGNORADAS.indexOf(nome) === -1) {
        
        var statusLocal = 'pendente';
        try {
           // Verifica se a célula AC2 já tem o cabeçalho
           var ac2 = String(s.getRange("AC2").getValue()).trim().toUpperCase();
           if (ac2 === "CÓD. DO PACIENTE" || ac2 === "COD. DO PACIENTE") {
               statusLocal = 'feito';
           }
        } catch(e) {
           // Se a coluna não existir, vai dar erro e cair no pendente
        }
        
        lista.push({ nome: nome, status: statusLocal });
      }
    });
    
    return lista;
  } catch (e) {
    throw new Error("Falha ao escanear abas: " + e.message);
  }
}

// 2. Aplica a formatação na aba específica
function aplicarFormatacaoAC(nomeAba) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) throw new Error("Aba não encontrada.");

    // Verifica se tem colunas suficientes. Se não tiver até a 29 (AC), cria.
    if (sheet.getMaxColumns() < 29) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 29 - sheet.getMaxColumns());
    }

    // Achar a última linha preenchida na Coluna B (Hora)
    var valoresB = sheet.getRange(1, 2, sheet.getMaxRows(), 1).getDisplayValues();
    var ultimaLinhaB = 0;
    
    for (var i = valoresB.length - 1; i >= 0; i--) {
      if (valoresB[i][0] && valoresB[i][0].trim() !== "") {
        ultimaLinhaB = i + 1; // i é índice 0, linha é 1
        break;
      }
    }
    
    // Segurança: se a aba estiver muito vazia, vai até a linha 5
    if (ultimaLinhaB < 5) ultimaLinhaB = 5;

    // Regra do Gestor: Até a célula POSTERIOR à última com hora (+1)
    var linhaLimite = ultimaLinhaB + 1;

    // 1. Inserir o Texto na AC2
    var celulaAlvo = sheet.getRange(2, 29); // Linha 2, Coluna 29 (AC)
    celulaAlvo.setValue("Cód. do Paciente");
    
    // 2. Copiar Formatação da AB (28) para AC (29) a partir da linha 2 até a linha limite
    var numLinhasParaCopiar = (linhaLimite - 2) + 1;
    
    var rangeOrigem = sheet.getRange(2, 28, numLinhasParaCopiar, 1);
    var rangeDestino = sheet.getRange(2, 29, numLinhasParaCopiar, 1);
    
    // Copia APENAS a formatação (cores, fontes, bordas), preservando os dados (vazios)
    rangeOrigem.copyTo(rangeDestino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    return { success: true };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}
