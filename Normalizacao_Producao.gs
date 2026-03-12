/****************************************************************
 * 🏭 MOTOR DE NORMALIZAÇÃO - DASHBOARD (Baseado na v4)
 * * Lógica de leitura ORIGINAL RESTAURADA (Funciona 100%)
 * * Saída em formato "Lista Contínua" para Dashboards
 ****************************************************************/

function normalizarPlanejamentoOperacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // NOME DA ABA ORIGEM
  const nomeAbaOrigem = "Base- Planejamento Operações"; 
  let sheetOrigem = ss.getSheetByName(nomeAbaOrigem);

  if (!sheetOrigem) {
    SpreadsheetApp.getUi().alert(`❌ Aba "${nomeAbaOrigem}" não encontrada! Verifique o nome.`);
    return;
  }

  const data = sheetOrigem.getDataRange().getValues();
  
  // Matriz única para o Dashboard
  let baseDashboard = [];
  
  // Listas de Controle (Idênticas ao V4)
  const produtosAlvo = ["AT1", "AT2", "AT5", "CT1", "CT2", "CT5"];
  const materiaisAlvo = ["Sobra", "Consumo", "Chegada", "Saldo"];
  
  let datasSemana = [];

  // Função interna com a sua regra exata
  function processarDataInteligente(valorBruto) {
    if (valorBruto instanceof Date) {
      return valorBruto;
    }
    
    let s = String(valorBruto).trim();
    let match = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
    
    if (match) {
      let d = parseInt(match[1], 10);
      let m = parseInt(match[2], 10);
      let y;
      
      if (match[3]) {
        y = parseInt(match[3], 10);
        if (y < 100) y += 2000; 
      } else {
        y = 2025; // Sem ano, fixa em 2025
      }
      
      return new Date(y, m - 1, d); 
    }
    return null;
  }

  function formatarDataFinal(dateObj) {
    if (!dateObj) return "";
    let d = String(dateObj.getDate()).padStart(2, '0');
    let m = String(dateObj.getMonth() + 1).padStart(2, '0');
    let y = dateObj.getFullYear();
    return `${d}/${m}/${y}`; 
  }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // 1. Detectar Cabeçalho de Semana (Lógica exata da v4)
    const strCol22 = String(row[22]).trim().toLowerCase();
    const strCol2 = String(row[2]).trim();
    
    if (strCol22 === "sábado" || strCol2.match(/\d{1,2}\/\d{1,2}/)) {
      datasSemana = [];
      const indicesDias = [2, 6, 10, 14, 18];
      let dateObjSexta = null;

      // Varrer Segunda a Sexta
      for (let j = 0; j < indicesDias.length; j++) {
        let valRaw = row[indicesDias[j]];
        let dateObj = processarDataInteligente(valRaw);
        
        if (dateObj) {
          datasSemana[j] = formatarDataFinal(dateObj);
          if (j === 4) dateObjSexta = dateObj;
        } else {
          datasSemana[j] = String(valRaw).trim();
        }
      }
      
      // Calcular Sábado e Domingo a partir da Sexta-feira
      if (dateObjSexta) {
        let dateSabado = new Date(dateObjSexta.getTime());
        dateSabado.setDate(dateSabado.getDate() + 1); 
        datasSemana[5] = formatarDataFinal(dateSabado);
        
        let dateDomingo = new Date(dateObjSexta.getTime());
        dateDomingo.setDate(dateDomingo.getDate() + 2); 
        datasSemana[6] = formatarDataFinal(dateDomingo);
      } else {
        datasSemana[5] = "Sábado";
        datasSemana[6] = "Domingo";
      }
      continue;
    }
    
    if (String(row[2]).trim().toLowerCase() === "saída") continue;
    
    // 2. Extração dos Dados (Lógica exata da v4)
    const label = String(row[1]).trim();
    if (!label) continue; 
    
    if (produtosAlvo.includes(label.toUpperCase())) {
      const indicesDias = [2, 6, 10, 14, 18];
      
      for (let d = 0; d < 5; d++) {
        let baseCol = indicesDias[d];
        let dataDia = datasSemana[d];
        if (!dataDia) continue;
        
        let saida = limparValor(row[baseCol]);
        let saldo = limparValor(row[baseCol + 1]);
        let prod = limparValor(row[baseCol + 2]);
        
        // MUDANÇA AQUI: Escrevendo em formato de Lista para Dashboard
        if (saida !== "") baseDashboard.push([dataDia, "Produto", label, "Saída", saida]);
        if (saldo !== "") baseDashboard.push([dataDia, "Produto", label, "Saldo", saldo]);
        if (prod !== "")  baseDashboard.push([dataDia, "Produto", label, "Produção", prod]);
      }
      
      let prodSab = limparValor(row[22]);
      if (prodSab !== "" && datasSemana[5] && datasSemana[5] !== "Sábado") {
        baseDashboard.push([datasSemana[5], "Produto", label, "Produção", prodSab]);
      }
      
      let prodDom = limparValor(row[26]);
      if (prodDom !== "" && datasSemana[6] && datasSemana[6] !== "Domingo") {
        baseDashboard.push([datasSemana[6], "Produto", label, "Produção", prodDom]);
      }
      
    } else if (materiaisAlvo.includes(label)) {
      const indicesDias = [2, 6, 10, 14, 18];
      for (let d = 0; d < 5; d++) {
        let baseCol = indicesDias[d];
        let dataDia = datasSemana[d];
        if (!dataDia) continue;
        
        let valor = limparValor(row[baseCol]);
        if (valor !== "") {
          baseDashboard.push([dataDia, "Insumo", label, "Valor", valor]);
        }
      }
    }
  }

  // ==========================================
  // LÓGICA DE ORDENAÇÃO CRONOLÓGICA 
  // ==========================================
  
  function ordernarPorData(a, b) {
    let partesA = String(a[0]).split('/');
    let partesB = String(b[0]).split('/');
    
    if (partesA.length === 3 && partesB.length === 3) {
      let dataA = new Date(partesA[2], partesA[1] - 1, partesA[0]).getTime();
      let dataB = new Date(partesB[2], partesB[1] - 1, partesB[0]).getTime();
      return dataA - dataB; // Crescente (Mais antigo -> Mais novo)
    }
    return 0; // Mantém igual se não for data válida
  }

  baseDashboard.sort(ordernarPorData);

  // Adicionar cabeçalho após a ordenação
  baseDashboard.unshift(["Data", "Categoria", "Item", "Indicador", "Quantidade"]);
  
  // 3. Escrever na Nova Aba
  escreverAbaDashboard(ss, "BD_Dashboard_Producao", baseDashboard);
  
  try {
    SpreadsheetApp.getUi().alert("✅ Dados normalizados em LISTA e ORDENADOS POR DATA com sucesso!\n\nAba 'BD_Dashboard_Producao' pronta.");
  } catch(e) {}
}

// ================= FUNÇÕES AUXILIARES =================

function limparValor(val) {
  if (val == null) return "";
  const s = String(val).trim();
  if (s === "#N/A" || s === "-" || s === "") return "";
  
  // Força o valor a ser número se possível (melhor para o Dashboard somar)
  let num = Number(s);
  if (!isNaN(num) && s !== "") return num;
  
  return val;
}

function escreverAbaDashboard(ss, nomeAba, dados) {
  let aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  } else {
    aba.clear();
  }
  
  if (dados.length > 0) {
    aba.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
    
    const headerRange = aba.getRange(1, 1, 1, dados[0].length);
    headerRange.setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, dados[0].length);
  }
}

// ================= AUTOMAÇÃO (GATILHOS) =================

function ativarAutomacao() {
  desativarAutomacaoSemAviso(); 
  ScriptApp.newTrigger('normalizarPlanejamentoOperacoes').timeBased().everyHours(1).create();
  try { SpreadsheetApp.getUi().alert('⏰ Automação Ativada!\n\nOs dados serão extraídos a cada 1 hora automaticamente.'); } catch(e) {}
}

function desativarAutomacao() {
  desativarAutomacaoSemAviso();
  try { SpreadsheetApp.getUi().alert('🛑 Automação Desativada!'); } catch(e) {}
}

function desativarAutomacaoSemAviso() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'normalizarPlanejamentoOperacoes') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Gestão de Produção')
    .addItem('▶️ Gerar Lista para Dashboard', 'normalizarPlanejamentoOperacoes')
    .addSeparator()
    .addItem('⏰ Ativar Atualização Automática (1h)', 'ativarAutomacao')
    .addItem('🛑 Desativar Atualização Automática', 'desativarAutomacao')
    .addToUi();
}
