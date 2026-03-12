// Dashboard.gs — Arquivo Principal

function doGet(e) {
  var htmlContent = gerarHTMLComDados_();
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Dashboard PRO NEW')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function abrirDashboard() {
  var htmlContent = gerarHTMLComDados_();
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(1300).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Dashboard Operacional — PRO NEW');
}

function gerarHTMLComDados_() {
  var template = HtmlService.createHtmlOutputFromFile('DashboardPage').getContent();
  var dados = coletarTodosDados_();
  var json = JSON.stringify(dados);
  return template.replace('__DATA_PLACEHOLDER__', json);
}

// ── Cache de leitura: cada aba é lida UMA única vez, independente de quantas funções a usam ──
function makeCachedSS_(ss) {
  var cache = {};
  return {
    getSheetByName: function(nome) {
      if (!(nome in cache)) {
        var aba = ss.getSheetByName(nome);
        if (!aba) { cache[nome] = null; return null; }
        var dados = aba.getDataRange().getValues();
        cache[nome] = {
          getDataRange: function() {
            return { getValues: function() { return dados; } };
          }
        };
      }
      return cache[nome];
    }
  };
}

function coletarTodosDados_() {
  var ss = makeCachedSS_(SpreadsheetApp.getActiveSpreadsheet());
  return {
    resumoMensal:           lerResumoMensal_(ss),
    margemCliente:          lerMargemCliente_(ss),
    margemProduto:          lerMargemProduto_(ss),
    pedidosMes:             lerPedidosMes_(ss),
    clientes30d:            lerClientes30d_(ss),
    meta30d:                lerMeta30d_(ss),
    producao:               lerProducao_(ss),
    positivacaoMensal:      calcPositivacao_(ss),
    fretePorMes:            calcFretePorMes_(ss),
    ultimaCompraValor:      calcUltimaCompraValor_(ss),
    ticketMedioPedidos:     calcTicketMedioPedidos_(ss),
    freteSavingMensal:      lerFreteSaving_(ss),
    vendasDiario:           calcVendasDiario_(ss),
    logisticaDiario:        calcLogisticaDiario_(ss),
    pagamentosCarreg:       lerPagamentosCarregamento_(ss),
    otifLogistica:          calcOtifLogistica_(ss),
    otifMensal:             calcOtifMensal_(ss),
    clientesEmRisco:        lerClientesEmRisco_(ss),
    churnMensal:            calcChurnMensal_(ss),
    metaAnual:              40000000,
    saldoPallets:           lerSaldoPallets_(ss),
    palletsMRP:             lerControlePalletMRP_(ss),
    palletsPorMes:          calcPalletsPorMes_(ss),
    // Novos dados
    pedidosCadastradosOMIE: lerPedidosCadastradosOMIE_(ss),
    margemComercialReal:    calcMargemComercialReal_(ss),
    pedidosCadastrados:     lerPedidosCadastrados_(ss),
    margemComercialMeta:    calcMargemComercialMeta_(ss),
    clientesAtivos:         calcClientesAtivos_(ss),
    clientesSemCompra:      calcClientesSemCompra_(ss),
    pedidosRepresados:      lerPedidosRepresados_(ss),
    entradaCaixa:           (function(){ try { return calcEntradaCaixa_(ss); } catch(e) { return {diario:{},mensal:{}}; } })(),
    vendedores:             (function(){ try { return calcVendedores_(ss); } catch(e) { return {porVendedor:{},meses:[]}; } })(),
    mrp:                    (function(){ try { return calcMRP_(ss); } catch(e) { return {porSku:{},porMes:{},rupturas:[]}; } })()
  };
}
