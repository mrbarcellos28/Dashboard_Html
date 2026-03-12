// ═══════════════════════════════════════════════════════════════════════════════
// DashboardLeitores.gs — Leitores de Dados (abas Base-*)
// ═══════════════════════════════════════════════════════════════════════════════
// NOMES DAS ABAS: "Base-faturamento", "Base-pedidos- Aplicativo",
//   "Base-PEDIDOS - OMIE", "Base-frete saving", "Base-Controle de Pallet",
//   "Base-Saldo de Pallets", "Base-Pagamentos carregamento",
//   "Base-Logs Logística", "Base-Dados_Producao_Suprimentos"
// ═══════════════════════════════════════════════════════════════════════════════

// ── RESUMO MENSAL (Base-faturamento) ──
// "custos" aqui = IMPOSTOS. Margem comercial REAL vem de calcMargemComercialReal_
function lerResumoMensal_(ss) {
  var aba = ss.getSheetByName('Base-faturamento');
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var porMes = {};
  for (var i = 3; i < dados.length; i++) {
    var sit = String(dados[i][11] || '').trim();
    if (sit === 'Cancelado' || sit === 'Devolvido') continue;
    var data = toDate_(dados[i][2]); if (!data) continue;
    var mes = fmtMes_(data); if (!mes) continue;
    var fat = toN_(dados[i][12]);
    var qtd = toN_(dados[i][13]);
    var imp = parseBRL_(dados[i][17]) + parseBRL_(dados[i][21]) + parseBRL_(dados[i][23]) + parseBRL_(dados[i][24]) + parseBRL_(dados[i][25]);
    if (!porMes[mes]) porMes[mes] = { faturamento: 0, impostos: 0, qtd: 0 };
    porMes[mes].faturamento += fat; porMes[mes].qtd += qtd; porMes[mes].impostos += imp;
  }
  var MORD = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var r = [];
  for (var m in porMes) {
    var d = porMes[m]; var mg = d.faturamento - d.impostos;
    r.push({ mes: m, faturamento: round2_(d.faturamento), custos: round2_(d.impostos), margem: round2_(mg),
      margemPct: d.faturamento > 0 ? round2_(mg / d.faturamento * 10000) / 10000 : 0, qtd: Math.round(d.qtd) });
  }
  r.sort(function(a, b) { var pa=a.mes.split('/'),pb=b.mes.split('/'); return (parseInt(pa[1])*100+MORD.indexOf(pa[0]))-(parseInt(pb[1])*100+MORD.indexOf(pb[0])); });
  return r;
}

// ── MARGEM POR CLIENTE (Base-faturamento) ──
function lerMargemCliente_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 3; i < dados.length; i++) {
    var sit = String(dados[i][11] || '').trim();
    if (sit === 'Cancelado' || sit === 'Devolvido') continue;
    var data = toDate_(dados[i][2]); if (!data) continue;
    var mes = fmtMes_(data); if (!mes) continue;
    var cli = String(dados[i][4] || '').trim(); if (!cli) continue;
    var fat = toN_(dados[i][12]);
    var imp = parseBRL_(dados[i][17])+parseBRL_(dados[i][21])+parseBRL_(dados[i][23])+parseBRL_(dados[i][24])+parseBRL_(dados[i][25]);
    if (!r[mes]) r[mes] = {}; if (!r[mes][cli]) r[mes][cli] = { faturamento: 0, custos: 0 };
    r[mes][cli].faturamento += fat; r[mes][cli].custos += imp;
  }
  return r;
}

// ── MARGEM POR PRODUTO (Base-faturamento) ──
function lerMargemProduto_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 3; i < dados.length; i++) {
    var sit = String(dados[i][11] || '').trim();
    if (sit === 'Cancelado' || sit === 'Devolvido') continue;
    var data = toDate_(dados[i][2]); if (!data) continue;
    var mes = fmtMes_(data); if (!mes) continue;
    var prod = String(dados[i][7] || '').trim(); if (!prod) continue;
    var fat = toN_(dados[i][12]); var qtd = toN_(dados[i][13]);
    var imp = parseBRL_(dados[i][17])+parseBRL_(dados[i][21])+parseBRL_(dados[i][23])+parseBRL_(dados[i][24])+parseBRL_(dados[i][25]);
    if (!r[mes]) r[mes] = {}; if (!r[mes][prod]) r[mes][prod] = { faturamento: 0, custos: 0, qtd: 0 };
    r[mes][prod].faturamento += fat; r[mes][prod].custos += imp; r[mes][prod].qtd += qtd;
  }
  return r;
}

// ── PEDIDOS POR MÊS (Base-faturamento) ──
function lerPedidosMes_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return [];
  var dados = aba.getDataRange().getValues(); var pm = {};
  for (var i = 3; i < dados.length; i++) {
    var sit = String(dados[i][11]||'').trim(); if (sit==='Cancelado'||sit==='Devolvido') continue;
    var data = toDate_(dados[i][2]); if (!data) continue;
    var mes = fmtMes_(data); if (!mes) continue;
    if (!pm[mes]) pm[mes]={f:0,q:0}; pm[mes].f+=toN_(dados[i][12]); pm[mes].q+=toN_(dados[i][13]);
  }
  var r=[]; for (var m in pm) r.push({mes:m,qtdProdutos:Math.round(pm[m].q),faturamento:round2_(pm[m].f),ticketMedio:0});
  return r;
}

// ── CLIENTES 30D (Base-faturamento) ──
function lerClientes30d_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return [];
  var dados = aba.getDataRange().getValues(); var cl = {};
  for (var i = 3; i < dados.length; i++) {
    var sit = String(dados[i][11]||'').trim(); if (sit==='Cancelado'||sit==='Devolvido') continue;
    var data = toDate_(dados[i][2]); var cli = String(dados[i][4]||'').trim();
    if (!data||!cli) continue; if (!cl[cli]||data>cl[cli]) cl[cli]=data;
  }
  var r=[]; for (var c in cl) r.push({nome:c,ultimaCompra:fmtDateISO_(cl[c])}); return r;
}

// ── META 30D ──
function lerMeta30d_(ss) {
  var cl = lerClientes30d_(ss); var hoje=new Date(); var tot=cl.length; var sem=0;
  for (var i=0;i<cl.length;i++){var d=cl[i].ultimaCompra?new Date(cl[i].ultimaCompra):null;if(!d||(hoje.getTime()-d.getTime())/86400000>30)sem++;}
  return {clientesSemCompra:sem,totalClientes:tot,percSemCompra:tot>0?round2_(sem/tot*1000)/1000:0};
}

// ── PRODUÇÃO (Base-Dados_Producao_Suprimentos) ──
function lerProducao_(ss) {
  var aba = ss.getSheetByName('Base-Dados_Producao_Suprimentos'); if (!aba) return [];
  var dados = aba.getDataRange().getValues(); var r = [];
  for (var i = 1; i < dados.length; i++) {
    var row=dados[i]; if(!row[0]) continue; var mes=fmtMes_(row[0]); if(!mes) continue;
    r.push({mes:mes,metaCaixas:toN_(row[2]),caixasProd:toN_(row[3]),percMeta:toN_(row[4]),
      custoFolha:toN_(row[5]),custoUnit:toN_(row[6]),fatBruto:toN_(row[7]),devolucoes:toN_(row[8]),
      percDevol:toN_(row[9]),consumoTeorico:toN_(row[10]),compras:toN_(row[11]),percEfic:toN_(row[12]),
      custoMP:toN_(row[13]),percMP:toN_(row[14]),estoqueQtd:toN_(row[15]),estoqueRS:toN_(row[16])});
  }
  return r;
}

// ── POSITIVAÇÃO (Base-faturamento) ──
function calcPositivacao_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var pm = {};
  for (var i = 3; i < dados.length; i++) {
    var sit=String(dados[i][11]||'').trim(); if(sit==='Cancelado'||sit==='Devolvido') continue;
    var data=toDate_(dados[i][2]); var cli=dados[i][4]; if(!data||!cli) continue;
    var mes=fmtMes_(data); if(!mes) continue;
    if(!pm[mes]) pm[mes]={}; pm[mes][String(cli).trim()]=true;
  }
  var r={}; for(var m in pm) r[m]=Object.keys(pm[m]).length; return r;
}

// ── VENDAS DIÁRIO (Base-faturamento) ──
function calcVendasDiario_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 3; i < dados.length; i++) {
    var sit=String(dados[i][11]||'').trim(); if(sit==='Cancelado'||sit==='Devolvido') continue;
    var data=toDate_(dados[i][2]); if(!data) continue;
    var key=fmtDateISO_(data); if(!key) continue;
    if(!r[key]) r[key]={fat:0,qtd:0,numPedidos:0,numClientes:0,_n:{},_c:{}};
    r[key].fat+=toN_(dados[i][12]); r[key].qtd+=toN_(dados[i][13]);
    if(dados[i][6]!=null) r[key]._n[String(dados[i][6])]=true;
    if(dados[i][4]) r[key]._c[String(dados[i][4]).trim()]=true;
  }
  for(var k in r){r[k].numPedidos=Object.keys(r[k]._n).length;r[k].numClientes=Object.keys(r[k]._c).length;r[k].fat=round2_(r[k].fat);delete r[k]._n;delete r[k]._c;}
  return r;
}

// ── LOGÍSTICA DIÁRIO (Base-frete saving) ──
function calcLogisticaDiario_(ss) {
  var aba = ss.getSheetByName('Base-frete saving'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 1; i < dados.length; i++) {
    var data=toDate_(dados[i][0]); if(!data) continue;
    var key=fmtDateISO_(data); if(!key) continue;
    if(!r[key]) r[key]={custo:0,saving:0,antigo:0,entregas:0,pallets:0};
    r[key].custo+=toN_(dados[i][1]); r[key].pallets+=toN_(dados[i][6]);
    r[key].antigo+=toN_(dados[i][7]); r[key].saving+=toN_(dados[i][26]); r[key].entregas+=1;
  }
  for(var k in r){r[k].custo=round2_(r[k].custo);r[k].saving=round2_(r[k].saving);r[k].antigo=round2_(r[k].antigo);}
  return r;
}

// ── FRETE SAVING MENSAL (Base-frete saving) ──
function lerFreteSaving_(ss) {
  var aba = ss.getSheetByName('Base-frete saving'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 1; i < dados.length; i++) {
    var data=toDate_(dados[i][0]); if(!data) continue; var mes=fmtMes_(data); if(!mes) continue;
    if(!r[mes]) r[mes]={precoRealizado:0,precoAntigo:0,saving:0,entregas:0,pallets:0};
    r[mes].precoRealizado+=toN_(dados[i][1]); r[mes].precoAntigo+=toN_(dados[i][7]);
    r[mes].saving+=toN_(dados[i][26]); r[mes].pallets+=toN_(dados[i][6]); r[mes].entregas+=1;
  }
  for(var m in r){r[m].precoRealizado=round2_(r[m].precoRealizado);r[m].precoAntigo=round2_(r[m].precoAntigo);r[m].saving=round2_(r[m].saving);r[m].pallets=round2_(r[m].pallets);}
  return r;
}

// ── ÚLTIMA COMPRA VALOR (Base-faturamento) ──
function calcUltimaCompraValor_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var cl = {};
  for (var i = 3; i < dados.length; i++) {
    var sit=String(dados[i][11]||'').trim(); if(sit==='Cancelado'||sit==='Devolvido') continue;
    var data=toDate_(dados[i][2]); var cli=String(dados[i][4]||'').trim(); var fat=toN_(dados[i][12]);
    if(!data||!cli) continue;
    if(!cl[cli]||data>cl[cli].d) cl[cli]={d:data,v:fat};
    else if(data.getTime()===cl[cli].d.getTime()) cl[cli].v+=fat;
  }
  var r={}; for(var c in cl) r[c]=round2_(cl[c].v); return r;
}

// ── TICKET MÉDIO (Base-faturamento) ──
function calcTicketMedioPedidos_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var pm = {};
  for (var i = 3; i < dados.length; i++) {
    var sit=String(dados[i][11]||'').trim(); if(sit==='Cancelado'||sit==='Devolvido') continue;
    var data=toDate_(dados[i][2]); if(!data) continue; var mes=fmtMes_(data); if(!mes) continue;
    if(!pm[mes]) pm[mes]={n:{},f:0}; if(dados[i][6]!=null) pm[mes].n[String(dados[i][6])]=true; pm[mes].f+=toN_(dados[i][12]);
  }
  var r={}; for(var m in pm){var n=Object.keys(pm[m].n).length;r[m]={numPedidos:n,faturamento:round2_(pm[m].f),ticketMedio:n>0?round2_(pm[m].f/n):0};}
  return r;
}

// ── FRETE POR MÊS (Base-faturamento col S) ──
function calcFretePorMes_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var r = {};
  for (var i = 3; i < dados.length; i++) {
    var data=toDate_(dados[i][2]); if(!data) continue; var mes=fmtMes_(data); if(!mes) continue;
    if(!r[mes]) r[mes]=0; r[mes]+=parseBRL_(dados[i][18]);
  }
  return r;
}

// ── CHURN MENSAL (Base-faturamento) ──
function calcChurnMensal_(ss) {
  var aba = ss.getSheetByName('Base-faturamento'); if (!aba) return {};
  var dados = aba.getDataRange().getValues(); var pm = {};
  for (var i = 3; i < dados.length; i++) {
    var sit=String(dados[i][11]||'').trim(); if(sit==='Cancelado'||sit==='Devolvido') continue;
    var data=toDate_(dados[i][2]); var cli=String(dados[i][4]||'').trim(); var fat=toN_(dados[i][12]);
    if(!data||!cli) continue; var mes=fmtMes_(data); if(!mes) continue;
    if(!pm[mes]) pm[mes]={}; if(!pm[mes][cli]) pm[mes][cli]=0; pm[mes][cli]+=fat;
  }
  var MORD=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var ms=Object.keys(pm).sort(function(a,b){var pa=a.split('/'),pb=b.split('/');return(parseInt(pa[1])*100+MORD.indexOf(pa[0]))-(parseInt(pb[1])*100+MORD.indexOf(pb[0]));});
  var r={};
  for(var mi=1;mi<ms.length;mi++){var ant=pm[ms[mi-1]],atu=pm[ms[mi]];var p=0,rp=0;for(var c in ant){if(!atu[c]){p++;rp+=ant[c];}}r[ms[mi]]={clientesPerdidos:p,receitaPerdida:round2_(rp),clientesAnteriores:Object.keys(ant).length};}
  return r;
}

// ── OTIF (Base-Logs Logística) ──
function calcOtifLogistica_(ss) {
  var aba = ss.getSheetByName('Base-Logs Logística'); if (!aba) return {porMes:{},resumo:{},top5Lentos:[]};
  var dados = aba.getDataRange().getValues();
  var cr={},en={},ca={};
  for(var i=1;i<dados.length;i++){var op=String(dados[i][1]||'').trim();var est=String(dados[i][2]||'').trim();var dl=toDate_(dados[i][3]);var cli=String(dados[i][4]||'').trim();var id=String(dados[i][7]||'').trim();if(!id||!dl)continue;if(op==='Carga Criada'&&!cr[id])cr[id]={data:dl,cliente:cli};if(est==='Entregue')en[id]={data:dl};if(op==='Carga Cancelada'||est==='Cancelada')ca[id]=true;}
  var pm={},lt=[];
  for(var id in cr){var mes=fmtMes_(cr[id].data);if(!mes)continue;if(!pm[mes])pm[mes]={totalCargas:0,entregues:0,canceladas:0,lts:[],leadTimeMedio:0,otifPct:0};pm[mes].totalCargas++;if(ca[id])pm[mes].canceladas++;else if(en[id]){pm[mes].entregues++;var h=round2_((en[id].data.getTime()-cr[id].data.getTime())/3600000);pm[mes].lts.push(h);lt.push({horas:h});}}
  for(var m in pm){var l=pm[m].lts;if(l.length>0){var s=0;for(var k=0;k<l.length;k++)s+=l[k];pm[m].leadTimeMedio=round2_(s/l.length);}var b=pm[m].totalCargas-pm[m].canceladas;pm[m].otifPct=b>0?round2_(pm[m].entregues/b*10000)/10000:0;delete pm[m].lts;}
  var tC=Object.keys(cr).length,tE=Object.keys(en).length,tX=Object.keys(ca).length;var tLT=0;for(var li=0;li<lt.length;li++)tLT+=lt[li].horas;
  return{porMes:pm,resumo:{totalCriadas:tC,totalEntregues:tE,totalCanceladas:tX,leadTimeMedioGeral:lt.length>0?round2_(tLT/lt.length):0,otifGeral:(tC-tX)>0?round2_(tE/(tC-tX)*10000)/10000:0},top5Lentos:[]};
}
function calcOtifMensal_(ss){var o=calcOtifLogistica_(ss);if(!o||!o.porMes)return{};var r={};for(var m in o.porMes){var d=o.porMes[m];r[m]={total:d.totalCargas||0,perfeitas:d.entregues||0,avarias:0,falhas:d.canceladas||0,otifPct:d.otifPct||0,tempoMedioH:d.leadTimeMedio||0};}return r;}

// ── CLIENTES EM RISCO ──
function lerClientesEmRisco_(ss){var cl=lerClientes30d_(ss);var h=new Date();var li=[],er=0;for(var i=0;i<cl.length;i++){var d=cl[i].ultimaCompra?new Date(cl[i].ultimaCompra):null;var dias=d?Math.floor((h.getTime()-d.getTime())/86400000):999;var st=dias>60?'Em Risco':'Ativo';if(st==='Em Risco')er++;li.push({nome:cl[i].nome,ultimaCompra:cl[i].ultimaCompra||'',diasInativo:dias,status:st});}li.sort(function(a,b){return b.diasInativo-a.diasInativo;});return{lista:li.slice(0,50),resumo:{totalClientes:li.length,emRisco:er,ativos:li.length-er,percRisco:li.length>0?round2_(er/li.length*10000)/10000:0}};}

// ── PAGAMENTOS CARREGAMENTO (Base-Pagamentos carregamento) ──
function lerPagamentosCarregamento_(ss){var aba=ss.getSheetByName('Base-Pagamentos carregamento');if(!aba)return{porMes:{},porTipo:{},porId:{}};var dados=aba.getDataRange().getValues();var pm={},pt={};for(var i=1;i<dados.length;i++){var id=String(dados[i][0]||'').trim();var tipo=String(dados[i][1]||'').trim();var val=toN_(dados[i][2]);var stat=String(dados[i][3]||'').trim();var dp=toDate_(dados[i][4]);if(!id||!tipo||val===0)continue;if(dp){var mes=fmtMes_(dp);if(mes){if(!pm[mes])pm[mes]={total:0,count:0,pago:0,pendente:0};pm[mes].total+=val;pm[mes].count++;if(stat==='Pago')pm[mes].pago+=val;else pm[mes].pendente+=val;}}if(!pt[tipo])pt[tipo]=0;pt[tipo]+=val;}return{porMes:pm,porTipo:pt,porId:{}};}

// ── PALLETS ──
function lerSaldoPallets_(ss){var aba=ss.getSheetByName('Base-Saldo de Pallets');if(!aba)return{porCliente:[],porMotorista:[],resumo:{totalRetirados:0,totalDevolvidos:0,saldoPendente:0,numClientes:0}};var dados=aba.getDataRange().getValues();var cm={},mm={},tR=0,tD=0;for(var i=1;i<dados.length;i++){var cli=String(dados[i][1]||'').trim();var mot=String(dados[i][2]||'').trim();var plt=toN_(dados[i][3]);var op=String(dados[i][4]||'').trim();if(!cli||!op)continue;var cn=cli.split(' - ')[0].trim();if(!cm[cn])cm[cn]={r:0,d:0};if(!mm[mot])mm[mot]={r:0,d:0};if(op==='Retirados'){cm[cn].r+=plt;mm[mot].r+=plt;tR+=plt;}else if(op==='Devolvidos'){cm[cn].d+=plt;mm[mot].d+=plt;tD+=plt;}}var pc=[];for(var c in cm)pc.push({nome:c,retirados:cm[c].r,devolvidos:cm[c].d,saldo:cm[c].r-cm[c].d});pc.sort(function(a,b){return b.saldo-a.saldo;});var pmo=[];for(var m in mm)if(m)pmo.push({nome:m,retirados:mm[m].r,devolvidos:mm[m].d,saldo:mm[m].r-mm[m].d});pmo.sort(function(a,b){return b.saldo-a.saldo;});return{porCliente:pc,porMotorista:pmo,resumo:{totalRetirados:tR,totalDevolvidos:tD,saldoPendente:tR-tD,numClientes:pc.filter(function(c){return c.saldo>0}).length}};}
function lerControlePalletMRP_(ss){return{dias:[],resumo:{}};}
function calcPalletsPorMes_(ss){var aba=ss.getSheetByName('Base-frete saving');if(!aba)return{};var dados=aba.getDataRange().getValues();var r={};for(var i=1;i<dados.length;i++){var data=toDate_(dados[i][0]);if(!data)continue;var mes=fmtMes_(data);if(!mes)continue;if(!r[mes])r[mes]={pallets:0,entregas:0,ct:0};r[mes].pallets+=toN_(dados[i][6]);r[mes].entregas+=1;r[mes].ct+=toN_(dados[i][1]);}for(var m in r){r[m].pallets=round2_(r[m].pallets);r[m].custoPorPallet=r[m].pallets>0?round2_(r[m].ct/r[m].pallets):0;r[m].palletsPorEntrega=r[m].entregas>0?round2_(r[m].pallets/r[m].entregas):0;delete r[m].ct;}return r;}


// ═══════════════════════════════════════════════════════════════════════════════
// PEDIDOS CADASTRADOS OMIE + REPRESADOS
// Base-PEDIDOS - OMIE (ATENÇÃO: colunas desalinhadas dos headers!)
// Col real 4=Nº Pedido, 6=Cliente, 11=Data Inclusão, 14=Operação,
//   15=Situação, 18=Qtd, 20=Data Faturamento (null=REPRESADO), 24=Valor R$
// ═══════════════════════════════════════════════════════════════════════════════
function lerPedidosCadastradosOMIE_(ss){
  var aba=ss.getSheetByName('Base-PEDIDOS - OMIE');if(!aba)return{porMes:{},resumo:{totalRepresados:0,valorRepresado:0}};
  var dados=aba.getDataRange().getValues();var pm={};var tR=0,vR=0;
  for(var i=2;i<dados.length;i++){
    var op=String(dados[i][13]||'').trim();var sit=String(dados[i][14]||'').trim();
    if(op!=='Pedidos Cadastrados')continue;if(sit==='Cancelado'||sit==='Inutilizado'||sit==='Rejeitado')continue;
    var dInc=toDate_(dados[i][10]);var dFat=toDate_(dados[i][19]);var val=toN_(dados[i][23]);var qtd=toN_(dados[i][17]);var ped=dados[i][3];
    if(!dInc)continue;var mes=fmtMes_(dInc);if(!mes)continue;
    if(!pm[mes])pm[mes]={totalValor:0,faturadoValor:0,represadoValor:0,qtd:0,peds:{}};
    pm[mes].totalValor+=val;pm[mes].qtd+=qtd;if(ped&&ped!=='N/D')pm[mes].peds[String(ped)]=true;
    if(dFat instanceof Date){pm[mes].faturadoValor+=val;}else{pm[mes].represadoValor+=val;tR++;vR+=val;}
  }
  for(var m in pm){pm[m].totalValor=round2_(pm[m].totalValor);pm[m].faturadoValor=round2_(pm[m].faturadoValor);pm[m].represadoValor=round2_(pm[m].represadoValor);pm[m].represadoPct=pm[m].totalValor>0?round2_(pm[m].represadoValor/pm[m].totalValor*100):0;pm[m].numPedidos=Object.keys(pm[m].peds).length;delete pm[m].peds;}
  return{porMes:pm,resumo:{totalRepresados:tR,valorRepresado:round2_(vR)}};
}

// ═══════════════════════════════════════════════════════════════════════════════
// MARGEM COMERCIAL REAL — Base-pedidos- Aplicativo
// Margem = Total Vendido (col J) - Custo (col I)  ← CUSTO REAL do produto
// MargemPct = (Total Vendido - Custo) / Total Vendido
// ═══════════════════════════════════════════════════════════════════════════════
function calcMargemComercialReal_(ss){
  var aba=ss.getSheetByName('Base-pedidos- Aplicativo');if(!aba)return{porMes:{},meta12Pct:{}};
  var dados=aba.getDataRange().getValues();var pm={};var pd={};
  for(var i=1;i<dados.length;i++){
    var data=toDate_(dados[i][0]);var ped=dados[i][1];var tv=toN_(dados[i][9]);var cst=toN_(dados[i][8]);var mrg=toN_(dados[i][10]);var qtd=toN_(dados[i][6]);
    if(!data)continue;
    // col AA (idx26): só conta se Cadastrado no OMIE = "Sim"
    if(String(dados[i][26]||'').trim()!=='Sim')continue;
    var mes=fmtMes_(data);if(!mes)continue;
    if(!pm[mes])pm[mes]={tv:0,c:0,m:0,q:0,p:{}};pm[mes].tv+=tv;pm[mes].c+=cst;if(mrg>0)pm[mes].m+=mrg;pm[mes].q+=qtd;if(ped)pm[mes].p[String(ped)]=true;
    // Agrupamento diário para o gráfico de progresso
    var dKey=fmtDateISO_(data);if(!dKey)continue;
    if(!pd[dKey])pd[dKey]=0;pd[dKey]+=tv;
  }
  var r={},mt={};
  for(var m in pm){var d=pm[m];r[m]={totalVendido:round2_(d.tv),custo:round2_(d.c),margem:round2_(d.m),margemPct:d.tv>0?round2_(d.m/d.tv*10000)/10000:0,numPedidos:Object.keys(d.p).length,qtd:Math.round(d.q)};var meta=round2_(d.tv*0.12);mt[m]={totalVendido:round2_(d.tv),metaMargem:meta,margemReal:round2_(d.m),atingimento:meta>0?round2_(d.m/meta*100):0};}
  var rdaily={};for(var k in pd)rdaily[k]=round2_(pd[k]);
  return{porMes:r,meta12Pct:mt,diario:rdaily};
}

// ── CLIENTES ATIVOS (Base-pedidos- Aplicativo) ──
function calcClientesAtivos_(ss){
  var aba=ss.getSheetByName('Base-pedidos- Aplicativo');if(!aba)return{resumo:{ativos:0,alerta:0,perdidos:0,inativos:0,total12m:0},clientes:[]};
  var dados=aba.getDataRange().getValues();var hoje=new Date();var cl={};
  for(var i=1;i<dados.length;i++){
    var data=toDate_(dados[i][0]);var cli=dados[i][4]?String(dados[i][4]):'';var sku=dados[i][5]?String(dados[i][5]):'';var qtd=toN_(dados[i][6]);var total=toN_(dados[i][9]);
    if(!data||!cli)continue;var cn=cli.split(' - ')[0].trim();
    if(!cl[cn]){cl[cn]={ud:data,tc:total,pr:[{sku:sku,qtd:qtd,total:round2_(total)}]};}
    else if(data>cl[cn].ud){cl[cn]={ud:data,tc:total,pr:[{sku:sku,qtd:qtd,total:round2_(total)}]};}
    else if(data.getTime()===cl[cn].ud.getTime()){cl[cn].tc+=total;if(cl[cn].pr.length<5)cl[cn].pr.push({sku:sku,qtd:qtd,total:round2_(total)});}
  }
  var res={ativos:0,alerta:0,perdidos:0,inativos:0,total12m:0};var li=[];
  for(var c in cl){var info=cl[c];var dias=Math.floor((hoje.getTime()-info.ud.getTime())/86400000);var st;if(dias<=30){st='Ativo';res.ativos++;}else if(dias<=60){st='Alerta';res.alerta++;}else if(dias<=90){st='Lead Perdido';res.perdidos++;}else{st='Inativo';res.inativos++;}if(dias<=365)res.total12m++;li.push({nome:c,ultimaCompra:fmtDateISO_(info.ud),diasSemCompra:dias,status:st,totalUltCompra:round2_(info.tc),produtos:info.pr});}
  li.sort(function(a,b){return b.diasSemCompra-a.diasSemCompra;});
  return{resumo:res,clientes:li};
}

// ── CLIENTES SEM COMPRA ──
function calcClientesSemCompra_(ss){
  var aba=ss.getSheetByName('Base-pedidos- Aplicativo');if(!aba)return{sem30d:0,sem60d:0,sem90d:0,totalClientes:0};
  var dados=aba.getDataRange().getValues();var hoje=new Date();var cl={};
  for(var i=1;i<dados.length;i++){var data=toDate_(dados[i][0]);var c=dados[i][4]?String(dados[i][4]).split(' - ')[0].trim():'';if(!data||!c)continue;if(!cl[c]||data>cl[c])cl[c]=data;}
  var t=Object.keys(cl).length;var s3=0,s6=0,s9=0;for(var c in cl){var d=Math.floor((hoje.getTime()-cl[c].getTime())/86400000);if(d>30)s3++;if(d>60)s6++;if(d>90)s9++;}
  return{sem30d:s3,sem60d:s6,sem90d:s9,totalClientes:t};
}

// ── Legacy compat ──
function lerPedidosRepresados_(ss) {
  // Situações que indicam pedido represado (não faturado)
  var SITS = {'Aguardando faturamento':1,'Faturamento atrasado':1,'Pendente':1,'Previsto para hoje':1};
  var MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var aba = ss.getSheetByName('Base-PEDIDOS - OMIE');
  if (!aba) return {porMes:{}, total:0, totalPedidos:0};
  var dados = aba.getDataRange().getValues();
  // idx14=Situação, idx18=Data Inclusão (col S), idx22=Total Mercadoria (col W), idx3=Nº Pedido
  var porMes = {}; var totalVal = 0; var pedsUnicos = {};
  for (var i = 2; i < dados.length; i++) {
    var r = dados[i];
    if (!r[0] && !r[3]) continue;
    var sit = String(r[14]||'').trim();
    if (!SITS[sit]) continue;
    var dInc = toDate_(r[18]);
    if (!dInc) continue;
    var val = typeof r[22]==='number' ? r[22] : 0;
    var ped = String(r[3]||'').trim();
    var mes = MESES[dInc.getMonth()]+'/'+dInc.getFullYear();
    if (!porMes[mes]) porMes[mes] = {valor:0, pedidos:{}, sit:{}};
    porMes[mes].valor += val;
    if (ped) porMes[mes].pedidos[ped] = true;
    if (ped) pedsUnicos[ped] = true;
    if (!porMes[mes].sit[sit]) porMes[mes].sit[sit] = 0;
    porMes[mes].sit[sit] += val;
    totalVal += val;
  }
  // Formatar resultado
  var result = {};
  for (var m in porMes) {
    result[m] = {
      valor: round2_(porMes[m].valor),
      numPedidos: Object.keys(porMes[m].pedidos).length,
      sit: porMes[m].sit
    };
  }
  return {porMes: result, total: round2_(totalVal), totalPedidos: Object.keys(pedsUnicos).length};
}
function calcMargemComercialMeta_(ss){return calcMargemComercialReal_(ss).meta12Pct||{};}
function lerPedidosCadastrados_(ss){var o=lerPedidosCadastradosOMIE_(ss);return{porMes:o.porMes,pedidosDetalhe:[]};}

function calcVendedores_(ss) {
  // ── Fonte: Base-pedidos- Aplicativo, filtro AA=Sim (idx26) ──
  // Mesma lógica do calcMargemComercialReal_ — só pedidos cadastrados no OMIE
  var aba = ss.getSheetByName('Base-pedidos- Aplicativo');
  if (!aba) return {porVendedor:{}, meses:[]};
  var dados = aba.getDataRange().getValues();
  if (!dados || !dados.length) return {porVendedor:{}, meses:[]};

  var MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var vend  = {};
  var mesesSet = {};

  var getV = function(nome) {
    if (!vend[nome]) vend[nome] = {tv:0, cst:0, mrg:0, peds:{}, clis:{}, porMes:{}, porSku:{}};
    return vend[nome];
  };

  // Header índice 0, dados a partir do índice 1
  for (var i = 1; i < dados.length; i++) {
    var r = dados[i];
    if (!r[0]) continue;
    // ── Filtro AA=Sim (idx26): só pedidos cadastrados no OMIE ──
    if (String(r[26]||'').trim() !== 'Sim') continue;

    var data = toDate_(r[0]); if (!data) continue;
    var mes  = MESES[data.getMonth()] + '/' + data.getFullYear();

    var nome = String(r[3]||'N/D').trim();          // idx3 = Vendedor
    var ped  = String(r[1]||'').trim();             // idx1 = N° Pedido
    var cli  = String(r[4]||'').trim();             // idx4 = Cliente
    var sku  = String(r[5]||'').trim();             // idx5 = SKU
    var tv   = toN_(r[9]);                          // idx9 = Total Vendido
    var cst  = toN_(r[8]);                          // idx8 = Custo
    var mrg  = toN_(r[10]);                         // idx10= Margem

    var v = getV(nome);
    v.tv  += tv;
    v.cst += cst;
    v.mrg += (mrg > 0 ? mrg : 0);                  // só margem positiva, igual calcMargemComercialReal_
    if (ped) v.peds[ped] = true;
    if (cli) v.clis[cli] = true;
    if (sku) { if (!v.porSku[sku]) v.porSku[sku]=0; v.porSku[sku]+=tv; }
    if (!v.porMes[mes]) v.porMes[mes]=0;
    v.porMes[mes] += tv;
    mesesSet[mes] = true;
  }

  // Montar resultado final
  var result = {};
  for (var nome in vend) {
    var v = vend[nome];
    var skuArr = Object.keys(v.porSku).map(function(s){ return {sku:s, val:v.porSku[s]}; });
    skuArr.sort(function(a,b){ return b.val-a.val; });
    result[nome] = {
      totalRs:   round2_(v.tv),
      nPedidos:  Object.keys(v.peds).length,
      nClientes: Object.keys(v.clis).length,
      margem:    round2_(v.mrg),
      margemPct: v.tv>0 ? round2_(v.mrg/v.tv) : 0,
      devPct:    0,                                 // não disponível no Aplicativo — mantido para compatibilidade HTML
      porMes:    v.porMes,
      topSkus:   skuArr.slice(0,3).map(function(x){ return {sku:x.sku, val:round2_(x.val)}; })
    };
  }

  var meses = Object.keys(mesesSet).sort(function(a,b){
    var pa=a.split('/'),pb=b.split('/');
    var MA=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    return (parseInt(pa[1])*100+MA.indexOf(pa[0]))-(parseInt(pb[1])*100+MA.indexOf(pb[0]));
  });

  return {porVendedor: result, meses: meses};
}

function calcMRP_(ss) {
  // Estrutura real da aba Base-Controle de Pallet:
  // L3 (idx2): col3='MRP', col4+ = datas "d/mm/aaaa"
  // L4 (idx3): col3='Demanda Produção', col4+ = valores diários
  // L5 (idx4): col3='Retorno de Pallets', col4+ = valores
  // L6 (idx5): col3='Projeção de Compras', col4+ = valores
  // L7 (idx6): col3='Estoque Final Projetado', col4+ = valores (pode ser negativo = ruptura)
  var aba = ss.getSheetByName('Base-Controle de Pallet');
  if (!aba) return {porSku:{},porMes:{},rupturas:[]};
  var dados = aba.getDataRange().getValues();
  if (!dados || dados.length < 7) return {porSku:{},porMes:{},rupturas:[]};

  var MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var SKU = 'PALLETS';

  var dateRow   = dados[2]; // L3
  var demRow    = dados[3]; // L4 Demanda
  var retRow    = dados[4]; // L5 Retorno
  var compRow   = dados[5]; // L6 Compras
  var estoqRow  = dados[6]; // L7 Estoque Final

  var porMes = {};
  var porSku = {};
  porSku[SKU] = [];
  var rupturas = [];

  for (var ci = 4; ci < dateRow.length; ci++) {
    var dv = dateRow[ci];
    if (!dv) continue;
    var ds = String(dv).trim();
    // Formato "d/mm/aaaa"
    var dm = ds.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!dm) continue;
    var day = parseInt(dm[1]), mon = parseInt(dm[2]), year = parseInt(dm[3]);
    try {
      var dt = new Date(year, mon-1, day);
      if (isNaN(dt.getTime())) continue;
    } catch(e) { continue; }

    var mes     = MESES[mon-1] + '/' + year;
    var demanda = typeof demRow[ci]  === 'number' ? demRow[ci]  : 0;
    var retorno = typeof retRow[ci]  === 'number' ? retRow[ci]  : 0;
    var compras = typeof compRow[ci] === 'number' ? compRow[ci] : 0;
    var estoque = typeof estoqRow[ci]=== 'number' ? estoqRow[ci]: null;
    var dKey    = fmtDateISO_(dt);

    porSku[SKU].push({data:dKey, mes:mes, saida:demanda, saldo:estoque, prod:compras+retorno});

    if (!porMes[mes]) porMes[mes] = {};
    if (!porMes[mes][SKU]) porMes[mes][SKU] = {prodPlan:0, saidaPlan:0, saldoMin:null};
    porMes[mes][SKU].prodPlan  += demanda;      // Demanda planejada como "produção planificada"
    porMes[mes][SKU].saidaPlan += retorno;       // Retorno como "saída planejada"
    if (estoque !== null) {
      if (porMes[mes][SKU].saldoMin === null || estoque < porMes[mes][SKU].saldoMin)
        porMes[mes][SKU].saldoMin = estoque;
      if (estoque < 0) rupturas.push({data:dKey, sku:SKU, saldo:estoque});
    }
  }

  rupturas.sort(function(a,b){ return a.data < b.data ? -1 : 1; });
  return {porSku:porSku, porMes:porMes, rupturas:rupturas};
}

// ── ENTRADA DE CAIXA (Base-pedidos- Aplicativo: col A=Data, col J=Total Vendido, col L=Parcelas) ──
function calcEntradaCaixa_(ss){
  var aba=ss.getSheetByName('Base-pedidos- Aplicativo');if(!aba)return{diario:{},mensal:{}};
  var dados=aba.getDataRange().getValues();
  var MESES=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var diario={},mensal={};
  for(var i=1;i<dados.length;i++){
    var data=toDate_(dados[i][0]);if(!data)continue;
    var val=toN_(dados[i][9]);if(!val)continue;
    var prazoStr=String(dados[i][11]||'').trim();
    // Extrai primeiro número do prazo: "Para 28 dias"->28, "14/21/28"->14, "A Vista"->0
    var diasPrazo=0;
    if(prazoStr&&prazoStr.toLowerCase().indexOf('vista')===-1){
      var m=prazoStr.match(/(\d+)/);if(m)diasPrazo=parseInt(m[1]);
    }
    // Data de recebimento = data do pedido + prazo
    var dataReceb=new Date(data.getTime()+diasPrazo*86400000);
    var dKey=fmtDateISO_(dataReceb);if(!dKey)continue;
    var mes=MESES[dataReceb.getMonth()]+'/'+dataReceb.getFullYear();
    if(!diario[dKey])diario[dKey]=0;diario[dKey]+=val;
    if(!mensal[mes])mensal[mes]=0;mensal[mes]+=val;
  }
  // Arredondar
  for(var k in diario)diario[k]=round2_(diario[k]);
  for(var m in mensal)mensal[m]=round2_(mensal[m]);
  return{diario:diario,mensal:mensal};
}
