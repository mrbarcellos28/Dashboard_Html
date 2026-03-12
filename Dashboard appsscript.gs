// ═══════════════════════════════════════════════════════════════════════════════
// 📊 DASHBOARD - CÓDIGO APPS SCRIPT (com doGet para Web App / TV)
// ═══════════════════════════════════════════════════════════════════════════════
//
// INSTRUÇÕES DE INSTALAÇÃO:
//
// 1. No Apps Script da planilha, crie um NOVO arquivo Script:
//    → "+" > "Script" > Nomeie "Dashboard" > Cole ESTE código
//
// 2. Crie um arquivo HTML:
//    → "+" > "HTML" > Nomeie "DashboardPage" > Cole o HTML do DashboardPage.html
//
// 3. No seu código EXISTENTE, adicione ao onOpen() apenas:
//    .addSeparator()
//    .addItem('📊 Abrir Dashboard', 'abrirDashboard')
//
// 4. Para exibir em TV:
//    → Menu "Implantar" > "Nova implantação"
//    → Tipo: "App da Web"
//    → Executar como: "Eu"
//    → Quem tem acesso: "Qualquer pessoa" (ou da sua organização)
//    → Copie a URL gerada e abra no navegador da TV
//
// 5. Para atualizar após mudanças:
//    → "Implantar" > "Gerenciar implantações" > "Editar" > Nova versão
//
// NENHUMA função existente é modificada ou conflitada.
// ═══════════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════════
// doGet — OBRIGATÓRIA para funcionar como Web App (URL para TV)
// ═══════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  var htmlContent = gerarHTMLComDados_();
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Dashboard pronew')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ═══════════════════════════════════════════════════════════════════════════════
// abrirDashboard — Para abrir via menu dentro da planilha
// ═══════════════════════════════════════════════════════════════════════════════

function abrirDashboard() {
  var htmlContent = gerarHTMLComDados_();
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(1300)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Dashboard Operacional — PRO NEW');
}


// ═══════════════════════════════════════════════════════════════════════════════
// MOTOR DE DADOS — Lê TUDO direto das abas da planilha
// ═══════════════════════════════════════════════════════════════════════════════

function gerarHTMLComDados_() {
  var template = HtmlService.createHtmlOutputFromFile('Dashboardpage').getContent();
  var dados = coletarTodosDados_();
  var json = JSON.stringify(dados);
  return template.replace('__DATA_PLACEHOLDER__', json);
}

function coletarTodosDados_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    resumoMensal:        lerResumoMensal_(ss),
    margemCliente:       lerMargemCliente_(ss),
    margemProduto:       lerMargemProduto_(ss),
    pedidosMes:          lerPedidosMes_(ss),
    clientes30d:         lerClientes30d_(ss),
    meta30d:             lerMeta30d_(ss),
    producao:            lerProducao_(ss),
    positivacaoMensal:   calcPositivacao_(ss),
    fretePorMes:         calcFretePorMes_(ss),
    ultimaCompraValor:   calcUltimaCompraValor_(ss),
    ticketMedioPedidos:  calcTicketMedioPedidos_(ss),
    freteSavingMensal:   lerFreteSaving_(ss),
    vendasDiario:        calcVendasDiario_(ss),
    logisticaDiario:     calcLogisticaDiario_(ss),
    pagamentosCarreg:    lerPagamentosCarregamento_(ss),
    otifLogistica:       calcOtifLogistica_(ss),
    clientesEmRisco:     lerClientesEmRisco_(ss),
    metaAnual:           40000000
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// LEITORES — Cada função lê diretamente de uma aba específica
// ═══════════════════════════════════════════════════════════════════════════════

// Aba: "Resumo Mensal"
function lerResumoMensal_(ss) {
  var aba = ss.getSheetByName('Resumo Mensal');
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0] || String(row[0]).toUpperCase().indexOf('TOTAL') >= 0) continue;
    var mes = fmtMes_(row[0]);
    if (!mes) continue;
    result.push({
      mes: mes,
      faturamento: toN_(row[1]),
      custos: toN_(row[2]),
      margem: toN_(row[3]),
      margemPct: toN_(row[4]),
      qtd: toN_(row[5])
    });
  }
  return result;
}

// Aba: "Margem por Cliente"
function lerMargemCliente_(ss) {
  var aba = ss.getSheetByName('Margem por Cliente');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0] || !row[1]) continue;
    var mes = fmtMes_(row[0]);
    if (!mes) continue;
    var cli = String(row[1]).trim();
    if (!r[mes]) r[mes] = {};
    if (!r[mes][cli]) r[mes][cli] = { faturamento: 0, custos: 0 };
    r[mes][cli].faturamento += toN_(row[2]);
    r[mes][cli].custos += toN_(row[3]);
  }
  return r;
}

// Aba: "Margem por Produto"
function lerMargemProduto_(ss) {
  var aba = ss.getSheetByName('Margem por Produto');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0] || !row[1]) continue;
    var mes = fmtMes_(row[0]);
    if (!mes) continue;
    var prod = String(row[1]).trim();
    if (!r[mes]) r[mes] = {};
    if (!r[mes][prod]) r[mes][prod] = { faturamento: 0, custos: 0, quantidade: 0 };
    r[mes][prod].faturamento += toN_(row[2]);
    r[mes][prod].custos += toN_(row[3]);
    r[mes][prod].quantidade += toN_(row[6]);
  }
  return r;
}

// Aba: "Pedidos por Mês"
function lerPedidosMes_(ss) {
  var aba = ss.getSheetByName('Pedidos por Mês');
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var r = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0] || String(row[0]).toUpperCase().indexOf('TOTAL') >= 0) continue;
    var mes = fmtMes_(row[0]);
    if (!mes) continue;
    r.push({
      mes: mes,
      qtdProdutos: toN_(row[1]),
      faturamento: toN_(row[2]),
      ticketMedio: toN_(row[3])
    });
  }
  return r;
}

// Aba: "N° de vendas 30 dias" — lista de clientes com última compra
function lerClientes30d_(ss) {
  var aba = ss.getSheetByName('N° de vendas 30 dias');
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var r = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0]) continue;
    var entry = { nome: String(row[0]).trim() };
    entry.ultimaCompra = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : null;
    r.push(entry);
  }
  return r;
}

// Aba: "N° de vendas 30 dias" — célula C2, D2, E2 (resumo)
function lerMeta30d_(ss) {
  var aba = ss.getSheetByName('N° de vendas 30 dias');
  if (!aba) return { clientesSemCompra: 0, totalClientes: 0, percSemCompra: 0 };
  var dados = aba.getDataRange().getValues();
  if (dados.length < 2) return { clientesSemCompra: 0, totalClientes: 0, percSemCompra: 0 };
  return {
    clientesSemCompra: parseInt(dados[1][2]) || 0,
    totalClientes: parseInt(dados[1][3]) || 0,
    percSemCompra: parseFloat(dados[1][4]) || 0
  };
}

// Aba: "Dados_Producao_Suprimentos"
function lerProducao_(ss) {
  var aba = ss.getSheetByName('Dados_Producao_Suprimentos');
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var r = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0]) continue;
    var mes = fmtMes_(row[0]);
    if (!mes) continue;
    r.push({
      mes: mes,
      metaCaixas: toN_(row[2]),
      caixasProd: toN_(row[3]),
      percMeta: toN_(row[4]),
      custoFolha: toN_(row[5]),
      custoUnit: toN_(row[6]),
      fatBruto: toN_(row[7]),
      devolucoes: toN_(row[8]),
      percDevol: toN_(row[9]),
      consumoTeorico: toN_(row[10]),
      compras: toN_(row[11]),
      percEfic: toN_(row[12]),
      custoMP: toN_(row[13]),
      percMP: toN_(row[14]),
      estoqueQtd: toN_(row[15]),
      estoqueRS: toN_(row[16])
    });
  }
  return r;
}

// Aba: "Base faturamento" — conta clientes únicos por mês (positivação)
function calcPositivacao_(ss) {
  var aba = ss.getSheetByName('Base faturamento');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var porMes = {};
  for (var i = 3; i < dados.length; i++) {
    var data = dados[i][2];
    var cliente = dados[i][4];
    if (!data || !(data instanceof Date) || !cliente) continue;
    var mes = fmtMes_(data);
    if (!mes) continue;
    if (!porMes[mes]) porMes[mes] = {};
    porMes[mes][String(cliente).trim()] = true;
  }
  var r = {};
  for (var m in porMes) r[m] = Object.keys(porMes[m]).length;
  return r;
}

// Aba: "Base faturamento" — soma coluna S (Frete) por mês
function calcFretePorMes_(ss) {
  var aba = ss.getSheetByName('Base faturamento');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 3; i < dados.length; i++) {
    var data = dados[i][2];
    if (!data || !(data instanceof Date)) continue;
    var mes = fmtMes_(data);
    if (!mes) continue;
    var frete = parseBRL_(dados[i][18]);
    if (!r[mes]) r[mes] = 0;
    r[mes] += frete;
  }
  return r;
}


// ═══════════════════════════════════════════════════════════════════════════════
// NOVOS LEITORES — Pagamentos Carregamento, OTIF, Clientes em Risco
// ═══════════════════════════════════════════════════════════════════════════════

// Aba: "Pagamentos carregamento"
// A:A = ID, B:B = Tipo, C:C = Valor, D:D = Status, E:E = Data
// Considera TUDO (Pago + Pendente) para custo logístico correto
function lerPagamentosCarregamento_(ss) {
  var aba = ss.getSheetByName('Pagamentos carregamento');
  if (!aba) return { porMes: {}, porTipo: {}, porId: {} };
  var dados = aba.getDataRange().getValues();
  var porMes = {};   // mes -> {frete, descarga, pernoite, canhoto, total, count}
  var porTipo = {};   // tipo -> total
  var porId = {};     // id -> {total, tipos: {tipo: valor}}

  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    var id = String(row[0] || '').trim();
    var tipo = String(row[1] || '').trim();
    var valor = toN_(row[2]);
    var status = String(row[3] || '').trim();
    var dataP = row[4];

    if (!id || !tipo || valor === 0) continue;

    // Agrupa por mês
    if (dataP && (dataP instanceof Date)) {
      var mes = fmtMes_(dataP);
      if (mes) {
        if (!porMes[mes]) porMes[mes] = { frete: 0, descarga: 0, pernoite: 0, canhoto: 0, total: 0, count: 0, pago: 0, pendente: 0 };
        porMes[mes].total += valor;
        porMes[mes].count += 1;
        if (tipo === 'Custo de Frete') porMes[mes].frete += valor;
        else if (tipo === 'Custo de Descarga') porMes[mes].descarga += valor;
        else if (tipo === 'Pernoite') porMes[mes].pernoite += valor;
        else if (tipo === 'Canhoto') porMes[mes].canhoto += valor;
        if (status === 'Pago') porMes[mes].pago += valor;
        else porMes[mes].pendente += valor;
      }
    }

    // Agrupa por tipo
    if (!porTipo[tipo]) porTipo[tipo] = 0;
    porTipo[tipo] += valor;

    // Agrupa por ID
    if (!porId[id]) porId[id] = { total: 0, tipos: {} };
    porId[id].total += valor;
    if (!porId[id].tipos[tipo]) porId[id].tipos[tipo] = 0;
    porId[id].tipos[tipo] += valor;
  }

  // Round values
  for (var m in porMes) {
    porMes[m].frete = Math.round(porMes[m].frete * 100) / 100;
    porMes[m].descarga = Math.round(porMes[m].descarga * 100) / 100;
    porMes[m].pernoite = Math.round(porMes[m].pernoite * 100) / 100;
    porMes[m].total = Math.round(porMes[m].total * 100) / 100;
    porMes[m].pago = Math.round(porMes[m].pago * 100) / 100;
    porMes[m].pendente = Math.round(porMes[m].pendente * 100) / 100;
  }

  return { porMes: porMes, porTipo: porTipo, porId: porId };
}


// Aba: "Logs Logística" — Calcula OTIF e Lead Time
// Col B = Operação ("Carga Criada" para início)
// Col C = Novo estado ("Entregue" para fim)
// Col D = Horário de alteração
// Col H = ID do pedido (deve bater nos 2 casos)
function calcOtifLogistica_(ss) {
  var aba = ss.getSheetByName('Logs Logística');
  if (!aba) return { porMes: {}, leadTimes: [], resumo: {} };
  var dados = aba.getDataRange().getValues();

  // Mapeia: ID -> primeira "Carga Criada" (data) e "Entregue" (data)
  var criadas = {};    // id -> {data, cliente}
  var entregues = {};  // id -> {data}
  var canceladas = {}; // id -> true
  var totalCriadas = 0;
  var totalEntregues = 0;

  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    var operacao = String(row[1] || '').trim();
    var estado = String(row[2] || '').trim();
    var dataLog = row[3];
    var cliente = String(row[4] || '').trim();
    var idPedido = String(row[7] || '').trim();

    if (!idPedido || !dataLog || !(dataLog instanceof Date)) continue;

    // Carga Criada = início da contagem
    if (operacao === 'Carga Criada') {
      if (!criadas[idPedido]) {
        criadas[idPedido] = { data: dataLog, cliente: cliente };
        totalCriadas++;
      }
    }

    // Entregue = fim da contagem (col C)
    if (estado === 'Entregue') {
      entregues[idPedido] = { data: dataLog };
      totalEntregues++;
    }

    // Cancelada
    if (operacao === 'Carga Cancelada' || estado === 'Cancelada') {
      canceladas[idPedido] = true;
    }
  }

  // Calcula lead times e agrupa por mês
  var porMes = {};
  var leadTimes = [];
  var allIds = Object.keys(criadas);

  for (var j = 0; j < allIds.length; j++) {
    var id = allIds[j];
    var criadaData = criadas[id].data;
    var mes = fmtMes_(criadaData);
    if (!mes) continue;

    if (!porMes[mes]) porMes[mes] = {
      totalCargas: 0,
      entregues: 0,
      canceladas: 0,
      pendentes: 0,
      leadTimesH: [],
      leadTimeMedio: 0,
      leadTimeMin: 0,
      leadTimeMax: 0,
      otifPct: 0
    };

    porMes[mes].totalCargas++;

    if (canceladas[id]) {
      porMes[mes].canceladas++;
    } else if (entregues[id]) {
      porMes[mes].entregues++;
      var diffMs = entregues[id].data.getTime() - criadaData.getTime();
      var diffH = Math.round((diffMs / 3600000) * 10) / 10;
      porMes[mes].leadTimesH.push(diffH);
      leadTimes.push({ id: id, horas: diffH, mes: mes, cliente: criadas[id].cliente });
    } else {
      porMes[mes].pendentes++;
    }
  }

  // Calcula médias por mês
  for (var m in porMes) {
    var lt = porMes[m].leadTimesH;
    if (lt.length > 0) {
      var soma = 0;
      var mn = Infinity;
      var mx = -Infinity;
      for (var k = 0; k < lt.length; k++) {
        soma += lt[k];
        if (lt[k] < mn) mn = lt[k];
        if (lt[k] > mx) mx = lt[k];
      }
      porMes[m].leadTimeMedio = Math.round((soma / lt.length) * 10) / 10;
      porMes[m].leadTimeMin = Math.round(mn * 10) / 10;
      porMes[m].leadTimeMax = Math.round(mx * 10) / 10;
    }
    // OTIF = entregues / (total - canceladas)
    var base = porMes[m].totalCargas - porMes[m].canceladas;
    porMes[m].otifPct = base > 0 ? Math.round((porMes[m].entregues / base) * 10000) / 10000 : 0;
    // Remove array de lead times individuais para não poluir o JSON
    delete porMes[m].leadTimesH;
  }

  // Resumo geral
  var totalLT = 0;
  var countLT = 0;
  for (var li = 0; li < leadTimes.length; li++) {
    totalLT += leadTimes[li].horas;
    countLT++;
  }

  var resumo = {
    totalCriadas: totalCriadas,
    totalEntregues: totalEntregues,
    totalCanceladas: Object.keys(canceladas).length,
    leadTimeMedioGeral: countLT > 0 ? Math.round((totalLT / countLT) * 10) / 10 : 0,
    otifGeral: (totalCriadas - Object.keys(canceladas).length) > 0
      ? Math.round((totalEntregues / (totalCriadas - Object.keys(canceladas).length)) * 10000) / 10000
      : 0
  };

  // Top 5 lead times mais longos (para alertas)
  leadTimes.sort(function(a, b) { return b.horas - a.horas; });
  var top5Lentos = leadTimes.slice(0, 5).map(function(lt) {
    return { id: lt.id, horas: lt.horas, cliente: lt.cliente };
  });

  return {
    porMes: porMes,
    resumo: resumo,
    top5Lentos: top5Lentos
  };
}


// Aba: "Clientes em Risco"
// A = Cliente, B = Última Compra, C = Dias Inativo, D = Status
function lerClientesEmRisco_(ss) {
  var aba = ss.getSheetByName('Clientes em Risco');
  if (!aba) return { lista: [], resumo: {} };
  var dados = aba.getDataRange().getValues();
  var lista = [];
  var emRisco = 0;
  var ativo = 0;

  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[0]) continue;
    var nome = String(row[0]).trim();
    var ultimaCompra = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[1] || '');
    var diasInativo = toN_(row[2]);
    var status = String(row[3] || '').trim();

    lista.push({
      nome: nome,
      ultimaCompra: ultimaCompra,
      diasInativo: diasInativo,
      status: status
    });

    if (status.toUpperCase().indexOf('RISCO') >= 0) emRisco++;
    else ativo++;
  }

  // Ordena por dias inativo (maior primeiro)
  lista.sort(function(a, b) { return b.diasInativo - a.diasInativo; });

  return {
    lista: lista.slice(0, 50), // Top 50 para não poluir o JSON
    resumo: {
      totalClientes: lista.length,
      emRisco: emRisco,
      ativos: ativo,
      percRisco: lista.length > 0 ? Math.round((emRisco / lista.length) * 10000) / 10000 : 0
    }
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// DADOS ADICIONAIS (cruzamento de abas) — MANTIDOS
// ═══════════════════════════════════════════════════════════════════════════════

// Aba: "Base frete saving" — custo logístico e saving por mês
function lerFreteSaving_(ss) {
  var aba = ss.getSheetByName('Base frete saving');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 1; i < dados.length; i++) {
    var data = dados[i][0];
    var preco = dados[i][1];
    var precoAntigo = dados[i][7];
    var saving = dados[i][26];
    if (!data || !(data instanceof Date)) continue;
    var mes = fmtMes_(data);
    if (!mes) continue;
    if (!r[mes]) r[mes] = { precoRealizado: 0, precoAntigo: 0, saving: 0, entregas: 0 };
    r[mes].precoRealizado += (typeof preco === 'number') ? preco : 0;
    r[mes].precoAntigo += (typeof precoAntigo === 'number') ? precoAntigo : 0;
    r[mes].saving += (typeof saving === 'number') ? saving : 0;
    r[mes].entregas += 1;
  }
  for (var m in r) {
    r[m].precoRealizado = Math.round(r[m].precoRealizado * 100) / 100;
    r[m].precoAntigo = Math.round(r[m].precoAntigo * 100) / 100;
    r[m].saving = Math.round(r[m].saving * 100) / 100;
  }
  return r;
}

// Vendas diário
function calcVendasDiario_(ss) {
  var aba = ss.getSheetByName('Base faturamento');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 3; i < dados.length; i++) {
    var data = dados[i][2]; var nf = dados[i][6]; var fat = dados[i][12]; var qtd = dados[i][13]; var cli = dados[i][4];
    if (!data || !(data instanceof Date)) continue;
    var key = Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!r[key]) r[key] = { fat: 0, qtd: 0, numPedidos: 0, numClientes: 0, _nfs: {}, _clis: {} };
    r[key].fat += (typeof fat === 'number') ? fat : 0;
    r[key].qtd += (typeof qtd === 'number') ? qtd : 0;
    if (nf != null) r[key]._nfs[String(nf)] = true;
    if (cli) r[key]._clis[String(cli).trim()] = true;
  }
  for (var k in r) {
    r[k].numPedidos = Object.keys(r[k]._nfs).length;
    r[k].numClientes = Object.keys(r[k]._clis).length;
    r[k].fat = Math.round(r[k].fat * 100) / 100;
    delete r[k]._nfs; delete r[k]._clis;
  }
  return r;
}

// Logística diário
function calcLogisticaDiario_(ss) {
  var aba = ss.getSheetByName('Base frete saving');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var r = {};
  for (var i = 1; i < dados.length; i++) {
    var data = dados[i][0]; var preco = dados[i][1]; var saving = dados[i][26];
    var antigo = dados[i][7]; var pallets = dados[i][6];
    if (!data || !(data instanceof Date)) continue;
    var key = Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!r[key]) r[key] = { custo: 0, saving: 0, antigo: 0, entregas: 0, pallets: 0 };
    r[key].custo += (typeof preco === 'number') ? preco : 0;
    r[key].saving += (typeof saving === 'number') ? saving : 0;
    r[key].antigo += (typeof antigo === 'number') ? antigo : 0;
    r[key].entregas += 1;
    r[key].pallets += (typeof pallets === 'number') ? pallets : 0;
  }
  for (var k in r) {
    r[k].custo = Math.round(r[k].custo * 100) / 100;
    r[k].saving = Math.round(r[k].saving * 100) / 100;
    r[k].antigo = Math.round(r[k].antigo * 100) / 100;
    r[k].custoPorPallet = r[k].pallets > 0 ? Math.round((r[k].custo / r[k].pallets) * 100) / 100 : 0;
  }
  return r;
}

// Valor da última compra
function calcUltimaCompraValor_(ss) {
  var aba = ss.getSheetByName('Base faturamento');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var clientLast = {};
  for (var i = 3; i < dados.length; i++) {
    var data = dados[i][2];
    var cliente = dados[i][4];
    var fat = dados[i][12];
    if (!data || !(data instanceof Date) || !cliente) continue;
    var cli = String(cliente).trim();
    var fatVal = (typeof fat === 'number') ? fat : 0;
    if (!clientLast[cli] || data > clientLast[cli].date) {
      clientLast[cli] = { date: data, value: fatVal };
    } else if (data.getTime() === clientLast[cli].date.getTime()) {
      clientLast[cli].value += fatVal;
    }
  }
  var r = {};
  for (var c in clientLast) r[c] = Math.round(clientLast[c].value * 100) / 100;
  return r;
}

// Ticket médio por pedido
function calcTicketMedioPedidos_(ss) {
  var aba = ss.getSheetByName('Base faturamento');
  if (!aba) return {};
  var dados = aba.getDataRange().getValues();
  var porMes = {};
  for (var i = 3; i < dados.length; i++) {
    var data = dados[i][2];
    var nf = dados[i][6];
    var fat = dados[i][12];
    if (!data || !(data instanceof Date)) continue;
    var mes = fmtMes_(data); if (!mes) continue;
    var fatVal = (typeof fat === 'number') ? fat : 0;
    if (!porMes[mes]) porMes[mes] = { nfs: {}, fat: 0 };
    if (nf != null) porMes[mes].nfs[String(nf)] = true;
    porMes[mes].fat += fatVal;
  }
  var r = {};
  for (var m in porMes) {
    var nPed = Object.keys(porMes[m].nfs).length;
    r[m] = {
      numPedidos: nPed,
      faturamento: Math.round(porMes[m].fat * 100) / 100,
      ticketMedio: nPed > 0 ? Math.round((porMes[m].fat / nPed) * 100) / 100 : 0
    };
  }
  return r;
}


// ═══════════════════════════════════════════════════════════════════════════════
// UTILITÁRIOS
// ═══════════════════════════════════════════════════════════════════════════════

function fmtMes_(v) {
  if (!v) return null;
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  var nomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return nomes[d.getMonth()] + '/' + d.getFullYear();
}

function toN_(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  var n = parseFloat(String(v).replace(/[^\d,.-]/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function parseBRL_(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
