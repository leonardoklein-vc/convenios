// =====================================================================
// 🔐 CONFIGURADOR DE PERMISSÕES OTIMIZADO (LOTE V2 + INTERFACE VISUAL)
// =====================================================================

// --- LISTAS DE EMAILS (Padrão de Segurança) ---
// Estes valores serão carregados na tela para edição antes de aplicar.
var LISTA_LEITORES = [
  "mrrsilva@hcpa.edu.br", "l-admcardiologia@hcpa.edu.br", "l-secret-bloco@hcpa.edu.br",
  "l-secret-cca@hcpa.edu.br", "l-genf-ubc-enfermeiros@hcpa.edu.br", "l-convenios-faturamento@hcpa.edu.br",
  "l-convenios-internacao@hcpa.edu.br", "fpetry@hcpa.edu.br", "mjunior@hcpa.edu.br", "gduarte@hcpa.edu.br",
  "fbadalotti@hcpa.edu.br", "rdblopes@hcpa.edu.br", "flima@hcpa.edu.br", "bbernardi@hcpa.edu.br",
  "lrangel@hcpa.edu.br", "smolin@hcpa.edu.br", "nsavaris@hcpa.edu.br", "ncpereira@hcpa.edu.br",
  "elcouto@hcpa.edu.br", "asvargas@hcpa.edu.br", "acarcuchinski@hcpa.edu.br", "ccaldana@hcpa.edu.br",
  "rminuzzi@hcpa.edu.br", "caires@hcpa.edu.br", "psdsantos@hcpa.edu.br", "atmiranda@hcpa.edu.br",
  "boliveira@hcpa.edu.br", "tmelo@hcpa.edu.br", "alexguimaraes@hcpa.edu.br", "cbrito@hcpa.edu.br",
  "trnunes@hcpa.edu.br", "llunardi@hcpa.edu.br", "rvidal@hcpa.edu.br", "ktabarkiewicz@hcpa.edu.br",
  "dsoliveira@hcpa.edu.br", "mebmachado@hcpa.edu.br", "bwallauer@hcpa.edu.br", "amalaquias@hcpa.edu.br",
  "dmello@hcpa.edu.br", "edaimone@hcpa.edu.br", "lmenezes@hcpa.edu.br", "gia.anestesiagia.anestesia@gmail.com",
  "echagas@hcpa.edu.br", "lguterres@hcpa.edu.br", "cschiavo@hcpa.edu.br", "mvirginio@hcpa.edu.br",
  "rcosta@hcpa.edu.br", "mfunari@hcpa.edu.br", "smanica@hcpa.edu.br", "ppackeiser@hcpa.edu.br",
  "mmferreira@hcpa.edu.br", "emallmann@hcpa.edu.br", "gsomm@hcpa.edu.br", "escala@gia.med.br",
  "valves@hcpa.edu.br", "bnovelo@hcpa.edu.br", "gpaniz@hcpa.edu.br", "fpelisoli@hcpa.edu.br",
  "jviesi@hcpa.edu.br", "gbraulio@hcpa.edu.br", "pbasso@hcpa.edu.br", "dsanto@hcpa.edu.br",
  "klferreira@hcpa.edu.br", "masturizaga@hcpa.edu.br"
];

var LISTA_GUICHE = [
  "dgsilva@hcpa.edu.br", "dslongo@hcpa.edu.br", "ljleal@hcpa.edu.br", "lmylius@hcpa.edu.br",
  "marrdsilva@hcpa.edu.br", "mleite@hcpa.edu.br", "mromero@hcpa.edu.br", "rcmartins@hcpa.edu.br",
  "vcsilva@hcpa.edu.br"
];

var LISTA_AUTORIZACAO = [
  "tamartins@hcpa.edu.br", "gclima@hcpa.edu.br", "gschild@hcpa.edu.br","dsoliveira@hcpa.edu.br"
];

var LISTA_ADMIN = [
  "acteixeira@hcpa.edu.br", "asjorge@hcpa.edu.br", "carolcdcosta@hcpa.edu.br", "cmdasilva@hcpa.edu.br",
  "gsilveira@hcpa.edu.br", "lklrocha@hcpa.edu.br", "lzigue@hcpa.edu.br", "mluiz@hcpa.edu.br",
  "vregina@hcpa.edu.br", "jritt@hcpa.edu.br"
];

var SUPER_ADMINS_LISTA = [
  "gsilveira@hcpa.edu.br", "lklrocha@hcpa.edu.br", "vregina@hcpa.edu.br"
];

// Abre o Modal Visual (Aumentado para caber as listas de email)
function abrirConfiguradorPermissoes() {
  var html = HtmlService.createTemplateFromFile('ConfiguradorPermissoesHTML')
      .evaluate()
      .setTitle('⚙️ Configurador de Permissões e Acessos')
      .setWidth(750)  // <-- Mais largo para os grids
      .setHeight(650); // <-- Mais alto para caber os emails
  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Configurador de Permissões');
}

// Envia as listas padrão para o Frontend carregar nas Textareas
function obterListasIniciais() {
  return {
    admin: LISTA_ADMIN,
    autorizacao: LISTA_AUTORIZACAO,
    guiche: LISTA_GUICHE,
    leitores: LISTA_LEITORES
  };
}

// 1. Lista apenas as abas de Data (Backend)
function listarAbasDeDataParaConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var lista = [];
  
  // Regex para identificar abas de data: "DIA DD/MM"
  var regexData = /\d{2}\/\d{2}/; 

  sheets.forEach(function(s) {
    var nome = s.getName();
    if (ABAS_IGNORADAS.indexOf(nome) === -1 && regexData.test(nome)) {
      lista.push(nome);
    }
  });
  
  return lista;
}

// 2. Processa Permissões do DRIVE (Arquivo Geral) 
// AGORA RECEBE AS LISTAS DIRETAMENTE DO FRONTEND
function processarPermissoesDrive(listasAtualizadas) {
  try {
    var fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var file = DriveApp.getFileById(fileId);
    
    // 1. Leitores
    if (listasAtualizadas.leitores && listasAtualizadas.leitores.length > 0) {
      file.addViewers(listasAtualizadas.leitores);
    }

    // 2. Editores (Junta Admin + Guichê + Autorização)
    var todosEditores = [];
    if (listasAtualizadas.admin) todosEditores = todosEditores.concat(listasAtualizadas.admin);
    if (listasAtualizadas.guiche) todosEditores = todosEditores.concat(listasAtualizadas.guiche);
    if (listasAtualizadas.autorizacao) todosEditores = todosEditores.concat(listasAtualizadas.autorizacao);
    
    // Remove duplicados
    todosEditores = todosEditores.filter(function(item, pos) {
      return item && todosEditores.indexOf(item) == pos;
    });
    
    if (todosEditores.length > 0) {
      file.addEditors(todosEditores);
    }
    
    return { success: true, msg: "Permissões do ARQUIVO (Drive) atualizadas com sucesso!" };
  } catch (e) {
    return { success: false, msg: "Erro no Drive: " + e.message };
  }
}

// 3. Aplica permissão em UMA aba (Backend)
// AGORA RECEBE AS LISTAS DIRETAMENTE DO FRONTEND
function aplicarPermissaoEmUmaAba(nomeAba, listasAtualizadas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) return { success: false, msg: "Aba não encontrada" };

  try {
    // Remove proteções antigas
    var protecoes = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < protecoes.length; i++) {
      protecoes[i].remove();
    }

    // Cria nova proteção
    var protecao = sheet.protect().setDescription("Acesso Restrito - Liberado AA:AB");
    
    // Configura exceção (Coluna AA:AB liberada para os parciais poderem editar)
    var maxCols = sheet.getMaxColumns();
    if (maxCols >= 28) {
       var rangeDesbloqueado = sheet.getRange(1, 27, sheet.getMaxRows(), 2); // AA:AB
       protecao.setUnprotectedRanges([rangeDesbloqueado]);
    }

    // Define Editores do Bloqueado
    var me = Session.getEffectiveUser();
    protecao.addEditor(me);
    protecao.removeEditors(protecao.getEditors()); // Limpa todos
    protecao.addEditor(me); // Garante eu

    // Junta as listas editadas na tela + Super Admins Fixos
    var editoresPermitidos = [];
    if (listasAtualizadas.admin) editoresPermitidos = editoresPermitidos.concat(listasAtualizadas.admin);
    if (listasAtualizadas.autorizacao) editoresPermitidos = editoresPermitidos.concat(listasAtualizadas.autorizacao);
    
    editoresPermitidos = editoresPermitidos.concat(SUPER_ADMINS_LISTA);
    
    // Remove duplicados e vazios
    editoresPermitidos = editoresPermitidos.filter(function(item, pos) {
      return item && editoresPermitidos.indexOf(item) == pos;
    });

    protecao.addEditors(editoresPermitidos);

    return { success: true, msg: "Protegida com sucesso." };

  } catch (e) {
    return { success: false, msg: "Erro: " + e.message };
  }
}

// =====================================================================
// 8. HISTÓRICO DE PERMISSÕES
// =====================================================================

function salvarHistoricoPermissoes(abasProcessadas) {
  if (!abasProcessadas || abasProcessadas.length === 0) return;
  
  var props = PropertiesService.getDocumentProperties();
  var histStr = props.getProperty('HistoricoPermissoes_V1');
  var historico = histStr ? JSON.parse(histStr) : [];
  
  // Cria o novo registro
  var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var email = Session.getEffectiveUser().getEmail();
  
  historico.unshift({
    data: dataAtual,
    usuario: email,
    abas: abasProcessadas
  });
  
  // Limita a 30 registros para não pesar a memória
  if (historico.length > 30) historico = historico.slice(0, 30);
  
  props.setProperty('HistoricoPermissoes_V1', JSON.stringify(historico));
}

function obterHistoricoPermissoes() {
  var props = PropertiesService.getDocumentProperties();
  var histStr = props.getProperty('HistoricoPermissoes_V1');
  return histStr ? JSON.parse(histStr) : [];
}
