// =========================================================================
// FUNÇÕES AUXILIARES PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================

function atualizarCentroCirurgicoIndependente() {
  executarAtualizacaoIsolada(atualizarCentroCirurgico, 'Centro.Cir.', 'Dash.C.Cir');
}

function atualizarFiosIndependente() {
  executarAtualizacaoIsolada(atualizarFiosCentroCirurgico, 'Fios CentroCirurgico', 'Dash.Fios');
}

function atualizarEndoscopiaIndependente() {
  executarAtualizacaoIsolada(atualizarEndoscopia, 'Endoscopia', 'Dash.Endoscopia');
}

function executarAtualizacaoIsolada(funcaoAtualizacao, abaDados, abaDash) {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  ssLocal.toast(`Buscando dados no Estoque Externo para ${abaDados}...`, 'Atualização', 3);
  
  const idExterna = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
  try {
    const ssExterna = SpreadsheetApp.openById(idExterna);
    const guiaEstoque = ssExterna.getSheetByName('DadosEstoque');
    const dadosBrutos = guiaEstoque.getDataRange().getDisplayValues();
    
    funcaoAtualizacao(dadosBrutos);
    
    SpreadsheetApp.getUi().alert('Sucesso', `Valores atualizados na aba ${abaDados} e Gráfico gerado na aba ${abaDash}!`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', 'Sem permissão para acessar a base externa.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =========================================================================
// FUNÇÕES ESPECÍFICAS DE CADA SETOR (Chamadas pela Função Mestre ou Menu)
// =========================================================================

function atualizarCentroCirurgico(dadosEstoqueBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, "Centro.Cir.", "Dash.C.Cir", "Evolução de Status - Centro Cirúrgico (Últimos 7 Dias)");
}

function atualizarFiosCentroCirurgico(dadosEstoqueBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, "Fios CentroCirurgico", "Dash.Fios", "Evolução de Status - Fios (Últimos 7 Dias)");
}

function atualizarEndoscopia(dadosEstoqueBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, "Endoscopia", "Dash.Endoscopia", "Evolução de Status - Endoscopia (Últimos 7 Dias)");
}

// =========================================================================
// FUNÇÃO MOTOR GENÉRICA (Processa os dados para qualquer guia nos mesmos moldes)
// =========================================================================

function atualizarSetorGenerico(dadosEstoqueBrutos, nomeAbaDados, nomeAbaDash, tituloGrafico) {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ssLocal.getSheetByName(nomeAbaDados);
  const abaDash = ssLocal.getSheetByName(nomeAbaDash);
  
  if (!abaDados) {
    SpreadsheetApp.getUi().alert('Erro', `Aba "${nomeAbaDados}" não encontrada!`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!abaDash) {
    SpreadsheetApp.getUi().alert('Erro', `Aba "${nomeAbaDash}" não encontrada! Crie a guia para gerar o Dashboard.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // =========================================================================
  // 1. PROCESSAMENTO DOS DADOS
  // =========================================================================
  const mapaEstoque = new Map();
  
  for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
    const codItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase(); 
    
    if (codItem) {
      const procArr = [];
      const procAta = String(dadosEstoqueBrutos[i][19] || '').trim(); 
      const procAnd = String(dadosEstoqueBrutos[i][27] || '').trim(); 
      
      if (procAta) {
        procAta.split('\n').forEach(p => {
          if (p.trim()) procArr.push(p.trim() + " - Ata");
        });
      }
      if (procAnd) {
        procAnd.split('\n').forEach(p => {
          if (p.trim()) procArr.push(p.trim() + " - Em andamento");
        });
      }

      mapaEstoque.set(codItem, {
        desc: dadosEstoqueBrutos[i][2],       
        cmm: dadosEstoqueBrutos[i][8],        
        saldo: dadosEstoqueBrutos[i][7],      
        obs: dadosEstoqueBrutos[i][12],       
        valAta: dadosEstoqueBrutos[i][25],    
        ae: dadosEstoqueBrutos[i][30],        
        empenho: dadosEstoqueBrutos[i][13],   
        processo: procArr.join('\n')          
      });
    }
  }

  // =========================================================================
  // 2. ATUALIZAÇÃO DOS DADOS NA GUIA DO SETOR
  // =========================================================================
  const ultimaLinha = abaDados.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert('Aviso', `Nenhum código encontrado na Coluna A da aba ${nomeAbaDados}.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const codigos = abaDados.getRange(2, 1, ultimaLinha - 1, 1).getDisplayValues();
  const matrizAtualizacao = [];
  const contagemObs = {}; 

  for (let i = 0; i < codigos.length; i++) {
    const cod = String(codigos[i][0]).trim().toUpperCase();
    if (!cod) {
      matrizAtualizacao.push(["", "", "", "", "", "", "", ""]);
      continue;
    }

    const info = mapaEstoque.get(cod);
    if (info) {
      matrizAtualizacao.push([
        info.desc, 
        info.cmm, 
        info.saldo, 
        info.obs, 
        info.valAta,
        info.ae,
        info.empenho,
        info.processo
      ]);

      const obsStr = String(info.obs).trim() || "Sem Dados";
      contagemObs[obsStr] = (contagemObs[obsStr] || 0) + 1;
    } else {
      matrizAtualizacao.push(["Item não encontrado", "", "", "", "", "", "", ""]);
    }
  }

  abaDados.getRange(2, 2, matrizAtualizacao.length, 8).setValues(matrizAtualizacao);

  // =========================================================================
  // 3. CONSTRUÇÃO DA TABELA DE HISTÓRICO NA GUIA DASH
  // =========================================================================
  const pesosCategorias = {
    "Abaixo 30 dias": 1,
    "Entre 30 e 59 Dias": 2,
    "Entre 60 e 89 Dias": 3,
    "Igual/Maior 90 Dias": 4,
    "Igual ou Maior a 90 Dias": 4, 
    "Sem Dados": 99 
  };

  const topObs = Object.keys(contagemObs).sort((a, b) => {
    const pesoA = pesosCategorias[a] || 50; 
    const pesoB = pesosCategorias[b] || 50;
    return pesoA - pesoB;
  }).slice(0, 5); 

  const cabecalhoHist = ["Data", ...topObs];
  
  let histData = [];
  const checkHeader = abaDash.getRange(1, 1).getDisplayValue(); 
  
  if (checkHeader === "Data") {
    let lastRowHist = 1;
    const valsI = abaDash.getRange(1, 1, abaDash.getMaxRows(), 1).getDisplayValues();
    for (let k = 0; k < valsI.length; k++) {
      if (valsI[k][0] !== "") lastRowHist = k + 1;
    }
    if (lastRowHist > 1) {
      histData = abaDash.getRange(2, 1, lastRowHist - 1, cabecalhoHist.length).getValues();
    }
  }

  let hojeStr = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
  let novaLinhaHist = [hojeStr];
  
  for (let i = 0; i < 5; i++) {
    let obsKey = topObs[i];
    novaLinhaHist.push(obsKey ? (contagemObs[obsKey] || 0) : 0);
  }

  histData.push(novaLinhaHist);

  if (histData.length > 7) {
    histData = histData.slice(histData.length - 7);
  }

  abaDash.getRange(1, 1, abaDash.getMaxRows(), 10).clearContent();

  abaDash.getRange(1, 1, 1, cabecalhoHist.length)
         .setValues([cabecalhoHist])
         .setBackground("#444444")
         .setFontColor("white")
         .setFontWeight("bold");
           
  abaDash.getRange(2, 1, histData.length, 1).setNumberFormat("@"); 
  abaDash.getRange(2, 1, histData.length, novaLinhaHist.length).setValues(histData);

  // =========================================================================
  // 4. CRIAÇÃO DO GRÁFICO NO DASH
  // =========================================================================
  const charts = abaDash.getCharts();
  charts.forEach(c => abaDash.removeChart(c));

  const rangeGrafico = abaDash.getRange(1, 1, histData.length + 1, cabecalhoHist.length);
  const coresGrafico = ["#D32F2F", "#F57C00", "#FBC02D", "#388E3C", "#1976D2"];

  const chartBuilder = abaDash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeGrafico)
    .setPosition(10, 1, 0, 0) 
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('hAxis.textStyle', {fontSize: 11})
    .setOption('title', tituloGrafico)
    .setOption('colors', coresGrafico)
    .setOption('legend', {position: 'top', textStyle: {fontSize: 12}})
    .setOption('pointSize', 6)
    .setOption('lineWidth', 3)
    .setOption('vAxis.gridlines.count', 7) 
    .setOption('series', {
      0: {dataLabel: 'value'},
      1: {dataLabel: 'value'},
      2: {dataLabel: 'value'},
      3: {dataLabel: 'value'},
      4: {dataLabel: 'value'}
    })
    .setOption('width', 900)
    .setOption('height', 400);

  abaDash.insertChart(chartBuilder.build());
}