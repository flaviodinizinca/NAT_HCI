function atualizarDashboardComRelatorio(guia, contagem) {
  guia.clear();
  
  // Limpar gráficos antigos para não sobrepor a cada atualização
  const charts = guia.getCharts();
  for (let i = 0; i < charts.length; i++) {
    guia.removeChart(charts[i]);
  }

  // =========================================================================
  // 1. PREPARAÇÃO DOS DADOS DO RESUMO
  // =========================================================================
  let totalGeral = 0;
  for (let cat in contagem) {
    totalGeral += contagem[cat].total;
  }
  
  const cabecalhoResumo = [["ESTADO DO ESTOQUE", "QTD TOTAL", "% ESTOQUE", "EMPENHADO", "C/ AE EMITIDA", "EM PROCESSO SEI", "SEM PROCESSO"]];
  const resumoValores = [];
  
  const coresTabela = ["#E67C73", "#FFD666", "#FFF2CC", "#D9EAD3", "#CFE2F3"];
  const coresGrafico = ["#D32F2F", "#F57C00", "#FBC02D", "#388E3C", "#1976D2"];
  
  let somaTotal = 0, somaEmp = 0, somaAe = 0, somaProc = 0, somaSemProc = 0;

  for (let cat in contagem) {
    let qtd = contagem[cat].total;
    let percentual = totalGeral > 0 ? (qtd / totalGeral) : 0;
    
    resumoValores.push([
      cat, 
      qtd, 
      percentual, 
      contagem[cat].empenho, 
      contagem[cat].ae, 
      contagem[cat].apenasProcesso, 
      contagem[cat].semProcesso
    ]);

    somaTotal   += qtd;
    somaEmp     += contagem[cat].empenho;
    somaAe      += contagem[cat].ae;
    somaProc    += contagem[cat].apenasProcesso;
    somaSemProc += contagem[cat].semProcesso;
  }

  resumoValores.push([
    "TOTAL GERAL", 
    somaTotal, 
    1.0, 
    somaEmp, 
    somaAe, 
    somaProc, 
    somaSemProc
  ]);

  // =========================================================================
  // 2. INSERÇÃO E FORMATAÇÃO DA TABELA PRINCIPAL
  // =========================================================================
  
  guia.getRange(1, 1, 1, 7).merge()
      .setValue("RESUMO GERENCIAL DE ESTOQUE - SEABA")
      .setFontSize(14)
      .setFontWeight("bold")
      .setFontColor("#202124")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

  guia.getRange(2, 1, 1, 7).setValues(cabecalhoResumo)
      .setBackground("#444444").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true);
      
  guia.getRange(3, 1, resumoValores.length, 7).setValues(resumoValores)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true);
      
  guia.getRange(3, 3, resumoValores.length, 1).setNumberFormat("0.0%");
  
  for (let r = 0; r < 5; r++) {
    if (resumoValores[r]) {
      guia.getRange(r + 3, 1, 1, 7).setBackground(coresTabela[r]);
    }
  }
  
  let linhaTotal = resumoValores.length + 2;
  guia.getRange(linhaTotal, 1, 1, 7)
      .setBackground("#D9E1F2")
      .setFontWeight("bold");

  guia.setColumnWidth(1, 200); 
  for (let c = 2; c <= 7; c++) {
    guia.setColumnWidth(c, 130);
  }

  guia.setHiddenGridlines(true);

  // =========================================================================
  // 3. CRIAÇÃO DO GRÁFICO DE LINHAS (Evolução dos Últimos 7 Dias)
  // =========================================================================
  const ss = guia.getParent();
  const guiaHistorico = ss.getSheetByName('Historico_Resumo');
  
  if (guiaHistorico && guiaHistorico.getLastRow() > 1) {
    const ultimaLinhaHist = guiaHistorico.getLastRow();
    
    let qtdDias = 7;
    let linhaInicio = 2;
    let numLinhas = ultimaLinhaHist - 1;
    
    if (numLinhas > qtdDias) {
      linhaInicio = ultimaLinhaHist - qtdDias + 1;
      numLinhas = qtdDias;
    }
    
    // --- ÁREA DE ESPELHO ---
    guiaHistorico.getRange("H:M").clearContent();
    
    const cabecalhoHist = guiaHistorico.getRange(1, 1, 1, 6).getValues();
    guiaHistorico.getRange(1, 8, 1, 6).setValues(cabecalhoHist);
    
    const dadosHist = guiaHistorico.getRange(linhaInicio, 1, numLinhas, 6).getValues();
    
    // TRUQUE: Transforma a data em texto para o gráfico não criar linha do tempo contínua
    for (let i = 0; i < dadosHist.length; i++) {
      let dataOriginal = dadosHist[i][0];
      if (dataOriginal instanceof Date) {
        let dia = String(dataOriginal.getDate()).padStart(2, '0');
        let mes = String(dataOriginal.getMonth() + 1).padStart(2, '0');
        let ano = dataOriginal.getFullYear();
        dadosHist[i][0] = dia + "/" + mes + "/" + ano; // Fica como texto fixo
      }
    }
    
    // Formata a coluna inteira do espelho como TEXTO PURO ("@") antes de colar os dados
    guiaHistorico.getRange(2, 8, dadosHist.length, 1).setNumberFormat("@");
    guiaHistorico.getRange(2, 8, dadosHist.length, 6).setValues(dadosHist);
    
    const rangeGraficoContinua = guiaHistorico.getRange(1, 8, dadosHist.length + 1, 6);

    const chartBuilder = guia.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(rangeGraficoContinua) 
        .setPosition(linhaTotal + 2, 1, 0, 0) 
        .setNumHeaders(1)
        .setOption('useFirstColumnAsDomain', true)
        // Força a tratar as labels como texto explícito no eixo horizontal
        .setOption('hAxis.textStyle', {fontSize: 11})
        .setOption('title', 'Evolução Histórica do Estoque (Últimos 7 Dias)')
        .setOption('colors', coresGrafico)
        .setOption('legend', {position: 'top', textStyle: {fontSize: 12}})
        .setOption('pointSize', 6)
        .setOption('lineWidth', 3)
        .setOption('series', {
          0: {dataLabel: 'value'},
          1: {dataLabel: 'value'},
          2: {dataLabel: 'value'},
          3: {dataLabel: 'value'},
          4: {dataLabel: 'value'}
        })
        .setOption('width', 980) 
        .setOption('height', 450);
        
    guia.insertChart(chartBuilder.build());
  }
}