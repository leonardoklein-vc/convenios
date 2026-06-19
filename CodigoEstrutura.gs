// =====================================================================
// 🏗️ MÓDULO DE ESTRUTURA (ÁREAS + SALAS OBRIGATÓRIAS) - V5 (VALIDAÇÃO EXPLÍCITA)
// =====================================================================

function abrirFerramentaEstrutura() {
  var html = HtmlService.createTemplateFromFile('EstruturaHTML')
      .evaluate()
      .setTitle('🏗️ Gerenciador de Estrutura')
      .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function obterEstruturaAtual() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];

  var dadosA = sheet.getRange(5, 1, lastRow - 4, 1).getValues();
  var estrutura = [];
  
  for (var i = 0; i < dadosA.length; i++) {
    var val = String(dadosA[i][0]);
    if (val && val.trim() !== "") {
      estrutura.push({ nome: val, linha: i + 5 });
    }
  }
  return estrutura;
}

function processarInsercaoEstrutura(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetUtil = ss.getSheetByName('Util'); // Referência para validações
  
  if (!sheetUtil) return { success: false, msg: "Erro: Aba 'Util' não encontrada." };

  // 1. Calcula ponto de inserção
  var linhaInsercao = parseInt(form.referencia);
  var posicao = form.posicao; 
  
  if (posicao === 'depois') {
    var maxRows = sheet.getLastRow();
    if (maxRows < 5) {
      linhaInsercao = 5;
    } else {
      for (var i = linhaInsercao; i <= maxRows; i++) {
         var proximaCel = sheet.getRange(i + 1, 1).getValue();
         if (i == maxRows || (i > linhaInsercao && proximaCel && proximaCel !== "")) {
           linhaInsercao = i + 1;
           break;
         }
      }
    }
  }

  // --- FLUXO A: NOVA ÁREA (COM SALA OBRIGATÓRIA) ---
  if (form.tipo === 'area') {
    sheet.insertRows(linhaInsercao, 2);
    
    // Configuração Visual da Área
    // Preto até a coluna AB (28)
    var rangeAreaTotal = sheet.getRange(linhaInsercao, 1, 2, 28); 
    rangeAreaTotal.setBackground("black");
    rangeAreaTotal.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    rangeAreaTotal.clearDataValidations(); // 🔥 Garante que não tenha validação
    
    // Mesclagem e Texto
    var rangeNomeArea = sheet.getRange(linhaInsercao, 1, 2, 14); // Mescla A:N para o texto
    rangeNomeArea.merge();
    rangeNomeArea.setValue(form.nomeArea.toUpperCase());
    rangeNomeArea.setFontColor("white");
    rangeNomeArea.setFontWeight("bold");
    rangeNomeArea.setFontSize(25); // 🔥 Tamanho 25
    rangeNomeArea.setHorizontalAlignment("center").setVerticalAlignment("middle"); // 🔥 Centralizado Total

    // Limpa colunas após AB (se houver) para ficar branco
    var colFinal = sheet.getLastColumn();
    if (colFinal > 28) {
       sheet.getRange(linhaInsercao, 29, 2, colFinal - 28).setBackground("white").setBorder(null, null, null, null, null, null);
    }

    // Insere a Sala logo abaixo (linhaInsercao + 2)
    inserirSalaLogica(sheet, sheetUtil, linhaInsercao + 2, form);
    
    return { success: true, msg: "Área '" + form.nomeArea + "' criada!" };
  }

  // --- FLUXO B: APENAS SALA ---
  if (form.tipo === 'sala') {
    inserirSalaLogica(sheet, sheetUtil, linhaInsercao, form);
    return { success: true, msg: "Sala '" + form.nomeSala + "' adicionada!" };
  }
}

// === FUNÇÃO CENTRAL DE CRIAÇÃO DE SALA ===
function inserirSalaLogica(sheet, sheetUtil, linhaInicio, form) {
  
  var qtdBlocos = calcularQtdBlocos(form.horaInicio, form.horaFim);
  if (qtdBlocos <= 0) throw new Error("Hora Fim deve ser maior que Início.");

  // Insere linhas: Blocos + 1 (separador)
  sheet.insertRows(linhaInicio, qtdBlocos + 1);
  
  // 1. Nome da Sala (Coluna A)
  var rangeSala = sheet.getRange(linhaInicio, 1, qtdBlocos, 1);
  rangeSala.merge();
  rangeSala.setValue(form.nomeSala.toUpperCase());
  
  // Estilo Sala
  rangeSala.setBackground("#38761d"); // Verde Escuro 2
  rangeSala.setFontColor("black");    // 🔥 Fonte Preta
  rangeSala.setFontWeight("bold");
  rangeSala.setFontSize(15);          // 🔥 Tamanho 15
  
  rangeSala.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  rangeSala.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);

  // 2. Horários (Coluna B)
  var horarios = gerarListaHorarios(form.horaInicio, qtdBlocos);
  var rangeHoras = sheet.getRange(linhaInicio, 2, qtdBlocos, 1);
  var matrizHoras = horarios.map(function(h){ return [h]; });
  rangeHoras.setValues(matrizHoras);
  rangeHoras.setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("white");
  rangeHoras.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  // 3. Dados (C até o fim) - Cor e Bordas
  var colFinal = sheet.getLastColumn();
  var rangeDados = sheet.getRange(linhaInicio, 3, qtdBlocos, colFinal - 2);
  var corFundo = (form.corTipo === 'ciano') ? '#00FFFF' : '#FFFFFF';
  
  rangeDados.setBackground(corFundo);
  rangeDados.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // =================================================================
  // 🔥 4. APLICAÇÃO DE VALIDAÇÕES EXPLÍCITAS (HARDCODED) 🔥
  // =================================================================
  
  // Helper para aplicar regra em uma coluna inteira do bloco
  var aplicarValidacao = function(colunaLetra, regra) {
    var colIndex = letraParaNumero(colunaLetra);
    var rangeDestino = sheet.getRange(linhaInicio, colIndex, qtdBlocos, 1);
    rangeDestino.setDataValidation(regra);
  };

  // G - Lado/Origem? (Util!H2:H23)
  var regraG = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$H$2:$H$23"), true).setAllowInvalid(false).build();
  aplicarValidacao("G", regraG);

  // H - Equipe? (Util!F2:F418)
  var regraH = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$F$2:$F$418"), true).setAllowInvalid(true).build(); // AllowInvalid true pq médicos mudam
  aplicarValidacao("H", regraH);

  // I - Origem? (Util!G2:G25)
  var regraI = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$G$2:$G$25"), true).setAllowInvalid(false).build();
  aplicarValidacao("I", regraI);

  // J - Marcação? (Util!J2:J15)
  var regraJ = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$J$2:$J$15"), true).setAllowInvalid(false).build();
  aplicarValidacao("J", regraJ);

  // L - I/A? (Util!I2:I24)
  var regraL = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$I$2:$I$24"), true).setAllowInvalid(false).build();
  aplicarValidacao("L", regraL);

  // M - Convênio (Util!E2:E100)
  var regraM = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$E$2:$E$100"), true).setAllowInvalid(true).build();
  aplicarValidacao("M", regraM);

  // O - OPME (SIM/NÃO)
  var regraO = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).setAllowInvalid(false).build();
  aplicarValidacao("O", regraO);

  // T - Status (Util!K2:K25)
  var regraT = SpreadsheetApp.newDataValidation().requireValueInRange(sheetUtil.getRange("$K$2:$K$25"), true).setAllowInvalid(false).build();
  aplicarValidacao("T", regraT);

  // =================================================================

  // 5. Linha Separadora (A última inserida) - LIMPEZA TOTAL
  var linhaSep = linhaInicio + qtdBlocos;
  var rangeSep = sheet.getRange(linhaSep, 1, 1, 28); // A até AB
  rangeSep.setBackground("black");
  rangeSep.clearContent();
  rangeSep.clearDataValidations(); // 🔥 Garante que não tenha validação na linha preta
  rangeSep.setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID); 
}

// Helpers
function letraParaNumero(letra) {
  // A=1, B=2 ...
  var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  // Suporte simples para A..Z e AA..AZ
  if (letra.length === 1) return base.indexOf(letra) + 1;
  if (letra.length === 2) {
     var d1 = base.indexOf(letra[0]) + 1;
     var d2 = base.indexOf(letra[1]) + 1;
     return (d1 * 26) + d2;
  }
  return 1;
}

function calcularQtdBlocos(inicio, fim) {
  var minIni = horaParaMinutos(inicio);
  var minFim = horaParaMinutos(fim);
  var diff = minFim - minIni;
  return Math.floor(diff / 30);
}

function horaParaMinutos(hStr) {
  var p = hStr.split(':');
  return parseInt(p[0])*60 + parseInt(p[1]);
}

function gerarListaHorarios(inicio, qtd) {
  var lista = [];
  var atual = horaParaMinutos(inicio);
  for (var i = 0; i < qtd; i++) {
    var h = Math.floor(atual / 60);
    var m = atual % 60;
    lista.push((h<10?'0':'')+h + ':' + (m<10?'0':'')+m);
    atual += 30;
  }
  return lista;
}
