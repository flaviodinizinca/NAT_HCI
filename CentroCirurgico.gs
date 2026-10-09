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
    const dadosEstoqueBrutos = guiaEstoque.getDataRange().getDisplayValues();
    
    const guiaEntradas = ssExterna.getSheetByName('EntradaEmpenhos');
    let dadosEntradasBrutos = [];
    if (guiaEntradas) {
      dadosEntradasBrutos = guiaEntradas.getDataRange().getValues();
    }
    
    funcaoAtualizacao(dadosEstoqueBrutos, dadosEntradasBrutos);
    
    SpreadsheetApp.getUi().alert('Sucesso', `Valores atualizados na aba ${abaDados} e Gráfico gerado na aba ${abaDash}!`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', 'Sem permissão para acessar a base externa.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =========================================================================
// FUNÇÕES ESPECÍFICAS DE CADA SETOR
// =========================================================================

function atualizarCentroCirurgico(dadosEstoqueBrutos, dadosEntradasBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, dadosEntradasBrutos, "Centro.Cir.", "Dash.C.Cir", "Evolução de Status - Centro Cirúrgico (Últimos 7 Dias)", "CENTRO");
}

function atualizarFiosCentroCirurgico(dadosEstoqueBrutos, dadosEntradasBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, dadosEntradasBrutos, "Fios CentroCirurgico", "Dash.Fios", "Evolução de Status - Fios (Últimos 7 Dias)", "FIOS_ENDO");
}

function atualizarEndoscopia(dadosEstoqueBrutos, dadosEntradasBrutos) {
  atualizarSetorGenerico(dadosEstoqueBrutos, dadosEntradasBrutos, "Endoscopia", "Dash.Endoscopia", "Evolução de Status - Endoscopia (Últimos 7 Dias)", "FIOS_ENDO");
}


// =========================================================================
// FUNÇÃO MOTOR GENÉRICA
// =========================================================================

function atualizarSetorGenerico(dadosEstoqueBrutos, dadosEntradasBrutos, nomeAbaDados, nomeAbaDash, tituloGrafico, tipoLayout) {
  
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
  // 1. PROCESSAMENTO DE ENTRADAS (Coluna X)
  // =========================================================================
  const mapaEntradas = new Map();
  if (dadosEntradasBrutos && dadosEntradasBrutos.length > 0) {
    for (let i = 1; i < dadosEntradasBrutos.length; i++) {
      const codItem = String(dadosEntradasBrutos[i][2]).trim().toUpperCase(); 
      const dataEntradaRaw = dadosEntradasBrutos[i][23]; 

      if (codItem && dataEntradaRaw) {
        let dataEntrada = null;
        
        if (dataEntradaRaw instanceof Date) {
          dataEntrada = dataEntradaRaw;
        } else {
          const stringData = String(dataEntradaRaw).trim();
          if (stringData !== "" && stringData !== "00:00:00" && stringData !== "0") {
            const parts = stringData.split('/');
            if (parts.length === 3) {
              dataEntrada = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
              dataEntrada = new Date(stringData);
            }
          }
        }
        
        if (dataEntrada instanceof Date && !isNaN(dataEntrada) && dataEntrada.getFullYear() > 1900) {
          if (!mapaEntradas.has(codItem)) {
            mapaEntradas.set(codItem, dataEntrada);
          } else {
            const dataAtual = mapaEntradas.get(codItem);
            if (dataEntrada > dataAtual) {
              mapaEntradas.set(codItem, dataEntrada);
            }
          }
        }
      }
    }
  }

  // =========================================================================
  // 2. PROCESSAMENTO DOS DADOS (Estoque)
  // =========================================================================
  const INDICE_SALDO_ATA = 27; // Atualizado para Coluna AB
  const mapaEstoque = new Map();

  for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
    const codItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase();       
    
    if (codItem) {
      const procArr = [];
      const procAta = String(dadosEstoqueBrutos[i][22] || '').trim(); // Atualizado para 22
      const procAnd = String(dadosEstoqueBrutos[i][30] || '').trim(); // Atualizado para 30
      
      if (procAta) {
        procAta.split('\n').forEach(p => { if (p.trim()) procArr.push(p.trim() + " - Ata"); });
      }
      if (procAnd) {
        procAnd.split('\n').forEach(p => { if (p.trim()) procArr.push(p.trim() + " - Em andamento"); });
      }

      mapaEstoque.set(codItem, {
        desc: dadosEstoqueBrutos[i][2],
        cmm: dadosEstoqueBrutos[i][8],
        saldo: dadosEstoqueBrutos[i][7],
        obs: dadosEstoqueBrutos[i][15], // Atualizado para 15
        saldoAta: dadosEstoqueBrutos[i][INDICE_SALDO_ATA], 
        valAta: dadosEstoqueBrutos[i][28], // Atualizado para 28
        ae: dadosEstoqueBrutos[i][33], // Atualizado para 33
        empenho: dadosEstoqueBrutos[i][16], // Atualizado para 16
        processo: procArr.join('\n')
      });
    }
  }

  // =========================================================================
  // 3. ATUALIZAÇÃO DOS DADOS NA GUIA DO SETOR (Sem formatar a planilha)
  // =========================================================================
  const ultimaLinha = abaDados.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert('Aviso', `Nenhum código encontrado na aba ${nomeAbaDados}.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const codigos = abaDados.getRange(2, 1, ultimaLinha - 1, 1).getDisplayValues();

  // Matrizes de atualização dependendo do layout
  const matrizBlocoPrincipal = []; // Usado por ambos, mas com larguras diferentes
  const matrizUltimaEntradaIsolada = []; // Usado apenas pelo Centro Cirúrgico
  const contagemObs = {}; 

  for (let i = 0; i < codigos.length; i++) {
    const cod = String(codigos[i][0]).trim().toUpperCase();
    
    let dataUltimaEntradaStr = "Ainda não houve entrada";
    if (mapaEntradas.has(cod)) {
      const d = mapaEntradas.get(cod);
      const dia = String(d.getDate()).padStart(2, '0');
      const mes = String(d.getMonth() + 1).padStart(2, '0');
      const ano = d.getFullYear();
      dataUltimaEntradaStr = `${dia}/${mes}/${ano}`;
    }

    if (!cod) {
      if (tipoLayout === "CENTRO") {
        matrizBlocoPrincipal.push(["", "", "", "", "", "", "", ""]);
        matrizUltimaEntradaIsolada.push([""]);
      } else {
        matrizBlocoPrincipal.push(["", "", "", "", "", "", "", "", "", ""]);
      }
      continue;
    }

    const info = mapaEstoque.get(cod);
    if (info) {
      if (tipoLayout === "CENTRO") {
        // Layout: B(Desc), C(CMM), D(Saldo), E(Obs), F(ValAta), G(AE), H(Empenho), I(Processo) ... L(UltimaEntrada)
        matrizBlocoPrincipal.push([info.desc, info.cmm, info.saldo, info.obs, info.valAta, info.ae, info.empenho, info.processo]);
        matrizUltimaEntradaIsolada.push([dataUltimaEntradaStr]);
      } else {
        // Layout FIOS/ENDO: B(Desc), C(CMM), D(Saldo), E(Obs), F(SaldoAta), G(ValAta), H(AE), I(Empenho), J(Processo), K(UltimaEntrada)
        matrizBlocoPrincipal.push([info.desc, info.cmm, info.saldo, info.obs, info.saldoAta, info.valAta, info.ae, info.empenho, info.processo, dataUltimaEntradaStr]);
      }
      
      const obsStr = String(info.obs).trim() || "Sem Dados";
      contagemObs[obsStr] = (contagemObs[obsStr] || 0) + 1;

    } else {
      if (tipoLayout === "CENTRO") {
        matrizBlocoPrincipal.push(["Item não encontrado", "", "", "", "", "", "", ""]);
        matrizUltimaEntradaIsolada.push([dataUltimaEntradaStr]);
      } else {
        matrizBlocoPrincipal.push(["Item não encontrado", "", "", "", "", "", "", "", "", dataUltimaEntradaStr]);
      }
    }
  }

  // Despeja os dados preservando formatação
  if (tipoLayout === "CENTRO") {
    abaDados.getRange(2, 2, matrizBlocoPrincipal.length, 8).setValues(matrizBlocoPrincipal); // Colunas B até I
    abaDados.getRange(2, 12, matrizUltimaEntradaIsolada.length, 1).setValues(matrizUltimaEntradaIsolada); // Coluna L isolada
  } else {
    abaDados.getRange(2, 2, matrizBlocoPrincipal.length, 10).setValues(matrizBlocoPrincipal); // Colunas B até K
  }

  // =========================================================================
  // 4. CONSTRUÇÃO DA TABELA DE HISTÓRICO NA GUIA DASH
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

  while (topObs.length < 5) {
    topObs.push("");
  }

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

  let existeHoje = false;
  if (histData.length > 0) {
    let dataUltimaLinha = histData[histData.length - 1][0];
    let strUltima = (dataUltimaLinha instanceof Date) 
          ? Utilities.formatDate(dataUltimaLinha, "GMT-3", "dd/MM/yyyy") 
          : String(dataUltimaLinha);
          
    if (strUltima === hojeStr) {
      existeHoje = true;
    }
  }

  if (existeHoje) {
    histData[histData.length - 1] = novaLinhaHist; 
  } else {
    histData.push(novaLinhaHist); 
    if (histData.length > 7) {
      histData = histData.slice(histData.length - 7);
    }
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
  // 5. CRIAÇÃO DO GRÁFICO NO DASH
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