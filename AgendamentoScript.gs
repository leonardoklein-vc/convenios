// =====================================================================
// 🩺 SISTEMA DE AGENDAMENTO HCPA – MÓDULO DE AGENDAMENTO & DADOS
// ARQUIVO: AgendamentoScript.gs
// Responsável por: Formulários de criação, Reagendamento, Cancelamento e Helpers de Leitura
// =====================================================================

// =====================================================================
// 1. ABERTURA DE FORMULÁRIOS (Novo Agendamento)
// =====================================================================

function abrirFormulario() {
  // 🔥 SEGURANÇA: Bloqueia leitores, parciais e intermediários
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      SpreadsheetApp.getUi().alert("⛔ NÃO AUTORIZADO\n\nSeu perfil não permite criar agendamentos.");
      return;
  }

  var html = montarHtmlFormulario('modal');
  if (html) SpreadsheetApp.getUi().showModalDialog(html, 'Agendar Cirurgia');
}

function abrirFormularioLateral() {
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      SpreadsheetApp.getUi().alert("⛔ NÃO AUTORIZADO\n\nSeu perfil não permite criar agendamentos.");
      return;
  }

  var html = montarHtmlFormulario('lateral');
  if (html) {
    html.setTitle('Novo Agendamento');
    SpreadsheetApp.getUi().showSidebar(html);
  }
}

function montarHtmlFormulario(modo) {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();
    var utilSheet = ss.getSheetByName('Util');

    // Validação de aba de data (função global no CodigoSistema ou local aqui se necessário)
    if (typeof ehAbaDeData === 'function' && !ehAbaDeData(sheetName)) {
        ui.alert("⚠️ Selecione uma guia de Data (ex: '25/02') para começar.");
        return null;
    }

    if (!utilSheet) {
      ui.alert("❌ A aba 'Util' é obrigatória e não foi encontrada.");
      return null;
    }

    var disponibilidade = lerDisponibilidade(sheet);
    var dadosValidacao = obterListasValidacao(ss);
    dadosValidacao.opme = ["SIM", "NÃO"];

    var dadosParaFront = {
      guia: sheetName,
      disponibilidade: disponibilidade,
      listas: dadosValidacao,
      modo: modo || 'modal'
    };

    var template = HtmlService.createTemplateFromFile('FormularioHTML');
    template.dadosIniciais = JSON.stringify(dadosParaFront);

    return template.evaluate().setWidth(950).setHeight(750);

  } catch (e) {
    ui.alert("❌ Erro ao preparar formulário: " + e.message);
    Logger.log(e);
    return null;
  }
}

// =====================================================================
// 2. REAGENDAMENTO (Abrir e Executar)
// =====================================================================

function abrirReagendamento() {
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      SpreadsheetApp.getUi().alert("⛔ NÃO AUTORIZADO\n\nSeu perfil não permite mover pacientes.");
      return;
  }
  
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert("⚠️ Selecione uma linha de agendamento (não o cabeçalho).");
    return;
  }

  var paciente = sheet.getRange(row, 3).getValue();
  if (!paciente || String(paciente).trim() === "") {
    ui.alert("⚠️ A linha selecionada não parece ter um paciente agendado (Coluna C vazia).");
    return;
  }

  try {
    var dadosOrigem = lerDadosDaLinha(sheet, row);
    var html = HtmlService.createTemplateFromFile('ReagendarHTML');

    html.dadosOrigem = JSON.stringify(dadosOrigem);
    html.listas = JSON.stringify(obterListasValidacao(ss));
    
    // obterListaGuias deve estar disponível globalmente (CodigoSistema.gs)
    var listaGuias = (typeof obterListaGuias === 'function') ? obterListaGuias() : [];
    html.guias = JSON.stringify(listaGuias);

    var output = html.evaluate().setTitle('🔄 Reagendar Cirurgia').setWidth(350);
    ui.showSidebar(output);

  } catch (e) {
    ui.alert("Erro ao ler dados: " + e.message);
  }
}

function executarReagendamento(form) {
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      return { success: false, message: "⛔ Não autorizado." };
  }

  Logger.log("Reagendando: " + JSON.stringify(form));

  // 1. Tenta salvar no novo local
  var resultadoSalvar = processFormData(form);

  // 2. Se salvou, apaga o antigo
      if (resultadoSalvar.success) {
        try {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          var sheetAntiga = ss.getSheetByName(form.origemSheet);
          var rowAntiga = parseInt(form.origemRow);
          var blocosParaLimpar = parseInt(form.origemDuracaoBlocos);

          // 🔥 CORREÇÃO: Limpa de C até AD (28 colunas)
          sheetAntiga.getRange(rowAntiga, 3, blocosParaLimpar, 28).clearContent();
          sheetAntiga.getRange(rowAntiga, 3, blocosParaLimpar, 28).setBackground('#ffffff');
          sheetAntiga.getRange(rowAntiga, 3, blocosParaLimpar, 28).clearNote(); 

          return { success: true, message: "✅ Movido com sucesso!" };

    } catch (e) {
      return {
        success: false,
        message: "⚠️ Salvo na nova data, mas erro ao limpar antiga: " + e.message
      };
    }

  } else {
    return resultadoSalvar;
  }
}

// =====================================================================
// 3. CANCELAMENTO (Lê e limpa até a coluna AD / Blindagem & M. Detetive)
// =====================================================================

function cancelarAgendamento() {
  Logger.log("--- INÍCIO CANCELAMENTO (MODO DETETIVE) ---");
  
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      SpreadsheetApp.getUi().alert("⛔ NÃO AUTORIZADO\n\nSeu perfil não permite cancelar.");
      return;
  }

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = ss.getActiveRange().getRow();

  if (row < 2) {
    ui.alert("⚠️ Selecione uma linha de agendamento válido na planilha (abaixo do cabeçalho) antes de clicar em Cancelar.");
    return;
  }

  var paciente = sheet.getRange(row, 3).getValue();
  if (!paciente || String(paciente).trim() === "") {
    ui.alert("⚠️ A linha selecionada não parece ter um paciente agendado.");
    return;
  }

  // Captura do Alerta para evitar que o fechamento da janela pareça um erro
  try {
      var resp = ui.alert("Confirmar cancelamento?", "Deseja mover para a aba 'CANCELADAS'?", ui.ButtonSet.YES_NO);
      if (resp !== ui.Button.YES) return;
  } catch(e) { 
      return; 
  }

  try {
    var info = lerDadosDaLinha(sheet, row);
    var blocos = info.origemDuracaoBlocos;
    var sheetName = sheet.getName();
    var salaOrig = getSalaMesclada(sheet, row);
    var lastRowSheet = sheet.getLastRow();

    if (row + blocos - 1 > lastRowSheet) {
      blocos = lastRowSheet - row + 1;
    }

    // 🔥 GARANTE QUE A ABA DO DIA TEM LARGURA SUFICIENTE PARA LER ATÉ A AD (30)
    if (sheet.getMaxColumns() < 30) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 30 - sheet.getMaxColumns());
    }

    // 🔥 LÊ DE B (2) ATÉ AD (30) = 29 COLUNAS TOTAIS
    var valoresBZ = sheet.getRange(row, 2, blocos, 29).getDisplayValues(); 
    var coresBZ = sheet.getRange(row, 2, blocos, 29).getBackgrounds();
    var notasBZ = sheet.getRange(row, 2, blocos, 29).getNotes();

    var cancelSheet = ss.getSheetByName('CANCELADAS');
    if (!cancelSheet) {
      cancelSheet = ss.insertSheet('CANCELADAS');
      var header = ['Dia', 'Sala'];
      // Gera colunas dinâmicas até a 30 (que cobrirá os anexos e códigos)
      for (var c = 2; c <= 30; c++) { header.push('Col ' + c); }
      cancelSheet.getRange(1, 1, 1, 31).setValues([header]);
    }

    // 🔥 GARANTE QUE A ABA CANCELADAS TEM LARGURA SUFICIENTE (31)
    if (cancelSheet.getMaxColumns() < 31) {
        cancelSheet.insertColumnsAfter(cancelSheet.getMaxColumns(), 31 - cancelSheet.getMaxColumns());
    }

    var lastRowCancel = cancelSheet.getLastRow();
    var newRow = lastRowCancel < 1 ? 2 : lastRowCancel + 1;
    var needRows = newRow + blocos - 1;
    
    if (needRows > cancelSheet.getMaxRows()) {
      cancelSheet.insertRowsAfter(cancelSheet.getMaxRows(), needRows - cancelSheet.getMaxRows());
    }

    var saidaValores = [];
    var saidaCores = [];
    var saidaNotas = [];

    for (var i = 0; i < blocos; i++) {
      var linhaDados = [sheetName, salaOrig];    
      var linhaCores = ['#ffffff', '#ffffff']; 
      var linhaNotas = ['', '']; 
      
      for (var j = 0; j < 29; j++) { 
          linhaDados.push(valoresBZ[i][j]); 
          linhaCores.push(coresBZ[i][j]);
          linhaNotas.push(notasBZ[i][j]);
      }
      saidaValores.push(linhaDados);
      saidaCores.push(linhaCores);
      saidaNotas.push(linhaNotas);
    }

    // =========================================================================
    // 🔥 MODO DETETIVE: SEPARANDO AS AÇÕES PARA VER QUAL CÉLULA ESTÁ BLOQUEADA
    // =========================================================================

    // AÇÃO 1: COLAR NAS CANCELADAS (31 colunas totais: 2 de metadados + 29 de dados capturados)
    var rangeDestino = cancelSheet.getRange(newRow, 1, blocos, 31);
    try {
        rangeDestino.setValues(saidaValores);
        rangeDestino.setBackgrounds(saidaCores);
        rangeDestino.setNotes(saidaNotas);
    } catch (eDestino) {
        var msgErro = "⛔ BLOQUEIO NA ABA 'CANCELADAS'!\n\n" +
                      "Você não tem permissão para editar as células: " + rangeDestino.getA1Notation() + "\n\n" +
                      "Verifique se esta aba está protegida no Google Planilhas.";
        Logger.log(msgErro);
        ui.alert(msgErro);
        return; 
    }

    // AÇÃO 2: APAGAR DADOS DA ABA ORIGINAL (Aba do Dia) - Da Coluna C até AD (28 colunas de largura)
    var rangeOrigemDados = sheet.getRange(row, 3, blocos, 28);
    try {
        rangeOrigemDados.clearContent();
        rangeOrigemDados.clearNote();
        rangeOrigemDados.setBackground('#ffffff'); 
    } catch (eOrigem) {
        var msgErro = "⛔ BLOQUEIO NA ABA '" + sheetName + "'!\n\n" +
                      "O paciente foi copiado para as Canceladas, mas você não tem permissão para APAGAR os dados dele nas células: " + rangeOrigemDados.getA1Notation() + "\n\n" +
                      "Motivo: Alguém pode ter protegido especificamente esta linha/célula.";
        Logger.log(msgErro);
        ui.alert(msgErro);
        return; 
    }

    ui.alert("✅ Agendamento cancelado com sucesso!");

  } catch (e) {
    Logger.log("ERRO FATAL: " + e.message + "\nStack: " + e.stack);
    ui.alert("❌ Erro inesperado: " + e.message);
  }
}

// =====================================================================
// 4. PROCESSAMENTO DE DADOS (Salvar Novo / Editar / Reagendar)
// =====================================================================

function processFormData(form) {
  // 🔥 SEGURANÇA
  if (typeof podeAlterarAgenda === 'function' && !podeAlterarAgenda()) {
      return { success: false, message: "⛔ Não autorizado." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(form.guia);
    if (!sheet) throw new Error("Aba '" + form.guia + "' não existe.");

    var startRow = parseInt(form.linhaInicial);
    if (isNaN(startRow) || startRow < 2) throw new Error("Linha inválida.");

    // Expande planilha se necessário
    var duracao = parseInt(form.duracao) || 60;
    var linhasNecessarias = Math.ceil((duracao + 30) / 30);
    if (startRow + linhasNecessarias > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 20);
    if (sheet.getMaxColumns() < 28) sheet.insertColumnsAfter(sheet.getMaxColumns(), 28 - sheet.getMaxColumns());
    
    // Helper para juntar arrays (Select2)
    var joinArr = function(v) { return (v && Array.isArray(v)) ? v.join(' + ') : (v || ""); };

    // --- RECUPERAÇÃO DE DADOS (Blindagem) ---
    var valoresBackup = form.valoresOriginais;
    
    // Se o backup veio vazio, força leitura da origem
    if ((!valoresBackup || valoresBackup.length === 0) && form.origemSheet && form.origemRow) {
         var sheetOrigem = ss.getSheetByName(form.origemSheet);
         if (sheetOrigem) {
             var dadosRecuperados = lerDadosDaLinha(sheetOrigem, parseInt(form.origemRow));
             valoresBackup = dadosRecuperados.valoresOriginais;
         }
    }

    var isEdit = (valoresBackup && valoresBackup.length > 0);
    
    // Helper inteligente
    function getValor(campoForm, indexOriginal, isArray) {
        if (form[campoForm] !== undefined) {
            return isArray ? joinArr(form[campoForm]) : form[campoForm];
        }
        if (isEdit && valoresBackup[indexOriginal] !== undefined) {
            return valoresBackup[indexOriginal];
        }
        return "";
    }

    var dadosLinha = [];

    if (isEdit) {
      // === MODO REAGENDAMENTO / EDIÇÃO ===
      dadosLinha = [
        getValor('paciente', 0),            // Col C
        getValor('nascimento', 1),          // Col D
        getValor('procedimentos', 2, true), // Col E
        getValor('cid', 3),                 // Col F
        getValor('lado', 4),                // Col G
        getValor('equipe', 5),              // Col H
        getValor('origem', 6),              // Col I
        getValor('marcadoPor', 7),          // Col J
        getValor('anestesista', 8),         // Col K
        getValor('ia', 9),                  // Col L
        getValor('convenio', 10),           // Col M
        "'" + getValor('matricula', 11).replace("'",""), // Col N
        getValor('opme', 12),               // Col O
        getValor('materiais', 13, true),    // Col P
        getValor('equipamentos', 14, true), // Col Q
        getValor('exames', 15, true),       // Col R
        
        getValor('agenda', 16),             // Col S (ID)
        getValor('status', 17),             // Col T (Status) - Pega o original por enquanto
        
        getValor('motivo', 18),             // Col U
        getValor('revisao', 19),            // Col V
        getValor('outrasInfo', 20),         // Col W
        getValor('telefone', 21),           // Col X
        getValor('cti', 22),                // Col Y
        getValor('aut', 23)                 // Col Z
      ];

      // --- LÓGICA DE MOVIMENTO (AQUI ESTÁ A MUDANÇA) ---
      var mudouDeLugar = (form.origemSheet && (form.origemSheet !== form.guia || parseInt(form.origemRow) !== startRow));
      
      if (mudouDeLugar) {
          dadosLinha[16] = ""; // Limpa ID
          dadosLinha[23] = ""; // Limpa AUT
          
          // 🔥 ALTERAÇÃO: Força "Remarcado" se o formulário não trouxe um status novo
          if (!form.status) {
              dadosLinha[17] = "Remarcado"; 
          }
      }

    } else {
      // === MODO NOVO AGENDAMENTO ===
      dadosLinha = [
        (form.paciente || ''), form.nascimento, joinArr(form.procedimentos),
        form.cid, form.lado, form.equipe, form.origem, form.marcadoPor,
        form.anestesista, form.ia, form.convenio, "'" + form.matricula, 
        form.opme, joinArr(form.materiais), joinArr(form.equipamentos), joinArr(form.exames),
        form.agenda, form.status, form.motivo, form.revisao, form.outrasInfo,
        form.telefone, form.cti, form.aut
      ];
    }

    // --- ESCRITA NO BANCO ---
    sheet.getRange(startRow, 3, 1, dadosLinha.length).setValues([dadosLinha]);
    sheet.getRange(startRow, 14).setNumberFormat("@");

    // Cores
    var props = PropertiesService.getUserProperties();
    var corPrincipal = (props.getProperty('corPrincipal') || '#d9d2e9').toLowerCase();
    var corBackup = (props.getProperty('corBackup') || '#b4a7d6').toLowerCase();
    var bgCor = corPrincipal;
    
    if (startRow > 5) {
       var corAcima = sheet.getRange(startRow-1, 3).getBackground().toLowerCase();
       if (corAcima === corPrincipal) bgCor = corBackup;
    }
    sheet.getRange(startRow, 3, linhasNecessarias, 24).setBackground(bgCor);

    // Checks
    if (isEdit && mudouDeLugar) {
       sheet.getRange(startRow, 27, 1, 2).setBackground('#ffff00').setValue("");
    } else if (!isEdit) {
       sheet.getRange(startRow, 27, 1, 2).setBackground('#ffffff').setValue("");
    }

    // 🔥 LÓGICA DO CÓDIGO (COLUNA AC = 29) E NOTA INTELIGENTE NA C (3)
    var codPac = form.codigoPaciente || "";
    var celCodigo = sheet.getRange(startRow, 29); // Coluna AC
    celCodigo.setValue(codPac);
    
    var celNome = sheet.getRange(startRow, 3); // Coluna C
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

    return { success: true, message: isEdit ? "✅ Reagendado com sucesso!" : "✅ Agendado com sucesso!" };

  } catch (error) {
    return { success: false, message: "❌ Erro: " + error.message };
  }
}

// =====================================================================
// 5. HELPER FUNCTIONS (Usadas pelo CodigoSistema.gs também)
// =====================================================================

function getSalaMesclada(sheet, row) {
  while (row >= 2) {
    var val = sheet.getRange(row, 1).getDisplayValue().trim();
    if (val !== "") return val;
    row--;
  }
  return "";
}

function lerDadosDaLinha(sheet, row) {
  var maxCols = sheet.getMaxColumns();
  
  // 🔥 CORREÇÃO: Lê até a coluna 30 (AD) para pegar Anexos
  // Se a planilha for menor, lê o máximo possível.
  var numColsToRead = (maxCols >= 30) ? 28 : (maxCols - 2); 
  
  // Pega valores e notas começando da coluna 3 (C)
  var values = sheet.getRange(row, 3, 1, numColsToRead).getDisplayValues()[0];
  var notas = sheet.getRange(row, 3, 1, numColsToRead).getNotes()[0];
  
  while (values.length < 28) { values.push(""); } // Garante array cheio

  // Calcula duração pelos blocos de cor
  var corBase = sheet.getRange(row, 3).getBackground();
  var duracaoBlocos = 1;
  var maxRows = sheet.getLastRow();

  for (var i = row + 1; i <= maxRows; i++) {
    var corProxima = sheet.getRange(i, 3).getBackground(); 
    var pacienteProximo = sheet.getRange(i, 3).getValue();
    if (corProxima === corBase && (!pacienteProximo || pacienteProximo === "")) {
      duracaoBlocos++;
    } else {
      break;
    }
  }

  // Mapeia notas ocultas
  var notasObj = {};
  var colMap = ['paciente','nascimento','procedimento','cid','lado','equipe','origem','marcacao','anestesia','ia','convenio','matricula','opme','materiais','equipamentos','exames','agenda','status','motivo','revisao','info','telefone','cti','aut','email','obs','guiche','anexos'];
  
  for(var n=0; n < colMap.length; n++) {
     if(notas[n] && notas[n] !== "") {
        notasObj[colMap[n]] = notas[n];
     }
  }

  var duracaoMinutos = duracaoBlocos * 30;

  return {
    id: row,
    origemSheet: sheet.getName(),
    origemRow: row,
    origemDuracaoBlocos: duracaoBlocos,
    valoresOriginais: values, 
    // Mapeamento
    paciente: values[0], nascimento: values[1],
    procedimentos: values[2] ? values[2].split(' + ') : [],
    cid: values[3], lado: values[4], equipe: values[5], origem: values[6],
    marcadoPor: values[7], anestesista: values[8], ia: values[9],
    convenio: values[10], matricula: values[11], opme: values[12],
    materiais: values[13] ? values[13].split(' + ') : [],
    equipamentos: values[14] ? values[14].split(' + ') : [],
    exames: values[15] ? values[15].split(' + ') : [],
    agenda: values[16], status: values[17], motivo: values[18],
    revisao: values[19], outrasInfo: values[20], telefone: values[21],
    cti: values[22], aut: values[23], 
    
    // --- NOVOS CAMPOS ---
    email: values[24],       // AA
    obs: values[25],         // AB
    codigoPaciente: values[26], // 🔥 NÓS ASSUMIMOS A COLUNA AC (29) AQUI
    anexos: values[27],      // AD (Link dos anexos)

    notas: notasObj,
    duracaoEstimada: duracaoMinutos
  };
}

function obterListasValidacao(ss) {
  var utilSheet = ss.getSheetByName('Util');
  var uLastRow = utilSheet.getLastRow();
  var uData = uLastRow > 1 ? utilSheet.getRange(2, 1, uLastRow - 1, 11).getValues() : [];

  function getCol(colIndex) {
    var result = [];
    var seen = {};
    for (var r = 0; r < uData.length; r++) {
      var val = uData[r][colIndex];
      if (val !== "" && val != null) {
        var txt = String(val).trim();
        if (!seen[txt]) {
          seen[txt] = true;
          result.push(txt);
        }
      }
    }
    return result.sort();
  }

  return {
    procedimentos: getCol(0), equipamentos: getCol(1), materiais: getCol(2),
    exames: getCol(3), convenios: getCol(4), equipes: getCol(5),
    origens: getCol(6), lados: getCol(7), ias: getCol(8),
    marcacoes: getCol(9), status: getCol(10),
    opme: ["SIM", "NÃO"] // <--- ADICIONE ESTA LINHA AQUI (com a vírgula na linha anterior)
  };
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
