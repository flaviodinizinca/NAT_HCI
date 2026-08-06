// =========================================================================
// FUNÇÃO AUXILIAR PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================
function atualizarCentroCirurgicoIndependente() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  ssLocal.toast('Buscando dados no Estoque Externo...', 'Atualização', 3);
  
  const idExterna = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
  try {
    const ssExterna = SpreadsheetApp.openById(idExterna);
    const guiaEstoque = ssExterna.getSheetByName('DadosEstoque');
    const dadosBrutos = guiaEstoque.getDataRange().getDisplayValues();
    
    atualizarCentroCirurgico(dadosBrutos);
    
    SpreadsheetApp.getUi().alert('Sucesso', 'Valores atualizados na aba Centro.Cir. e Gráfico de Histórico gerado na aba Dash.C.Cir!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', 'Sem permissão para acessar a base externa.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =========================================================================
// FUNÇÃO PRINCIPAL REFATORADA (Recebe os dados como parâmetro)
// =========================================================================
function atualizarCentroCirurgico(dadosEstoqueBrutos) {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const abaCentro = ssLocal.getSheetByName("Centro.Cir.");
  const abaDash = ssLocal.getSheetByName("Dash.C.Cir");
  
  if (!abaCentro) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba "Centro.Cir." não encontrada!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!abaDash) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba "Dash.C.Cir" não encontrada! Crie a guia para gerar o Dashboard.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // =========================================================================
  // 1. PROCESSAMENTO DOS DADOS (Recebidos da Função Mestre)
  // =========================================================================
  
  // Mapear as colunas baseadas na última atualização
  const mapaEstoque = new Map();
  // Assume-se que as linhas 0 e 1 são cabeçalhos (inicia no índice 2)
  for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
    const codItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase(); // Coluna B (Item)
    
    if (codItem) {
      const procArr = [];
      const procAta = String(dadosEstoqueBrutos[i][19] || '').trim(); // Coluna T (Ata)
      const procAnd = String(dadosEstoqueBrutos[i][27] || '').trim(); // Coluna AB (Em Andamento)
      
      // Quebra por linha caso exista mais de um processo na mesma célula originalmente
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
        desc: dadosEstoqueBrutos[i][2],       // Coluna C (Descrição)
        cmm: dadosEstoqueBrutos[i][8],        // Coluna I (CMM)
        saldo: dadosEstoqueBrutos[i][7],      // Coluna H (Saldo)
        obs: dadosEstoqueBrutos[i][12],       // Coluna M (Obs)
        valAta: dadosEstoqueBrutos[i][25],    // Coluna Z (Validade Ata)
        ae: dadosEstoqueBrutos[i][30],        // Coluna AE (Quantidade AE)
        empenho: dadosEstoqueBrutos[i][13],   // Coluna N (Empenho)
        processo: procArr.join('\n')          // Combinação dos processos
      });
    }
  }

  // =========================================================================
  // 2. ATUALIZAÇÃO DOS DADOS NA GUIA CENTRO.CIR. (Colunas B até I)
  // =========================================================================
  const ultimaLinha = abaCentro.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert('Aviso', 'Nenhum código encontrado na Coluna A da aba Centro.Cir.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const codigos = abaCentro.getRange(2, 1, ultimaLinha - 1, 1).getDisplayValues();
  const matrizAtualizacao = [];
  const contagemObs = {}; // Objeto para contar as ocorrências na Coluna E

  for (let i = 0; i < codigos.length; i++) {
    const cod = String(codigos[i][0]).trim().toUpperCase();
    if (!cod) {
      // São 8 colunas agora (B, C, D, E, F, G, H, I)
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

      // Conta a frequência dos agrupamentos da Coluna E para criar o gráfico. Trata células vazias.
      const obsStr = String(info.obs).trim() || "Sem Dados";
      contagemObs[obsStr] = (contagemObs[obsStr] || 0) + 1;
    } else {
      matrizAtualizacao.push(["Item não encontrado", "", "", "", "", "", "", ""]);
    }
  }

  // Grava as informações da coluna B (2) até I (9) que cobrem um range de 8 colunas
  abaCentro.getRange(2, 2, matrizAtualizacao.length, 8).setValues(matrizAtualizacao);

  // =========================================================================
  // 3. CONSTRUÇÃO DA TABELA DE HISTÓRICO NA GUIA DASH.C.CIR
  // =========================================================================
  
  // Define pesos personalizados para forçar a ordem crescente das observações
  const pesosCategorias = {
    "Abaixo 30 dias": 1,
    "Entre 30 e 59 Dias": 2,
    "Entre 60 e 89 Dias": 3,
    "Igual/Maior 90 Dias": 4,
    "Igual ou Maior a 90 Dias": 4, // Tratamento para caso o texto venha com pequena variação
    "Sem Dados": 99 // Peso alto para forçar a ficar sempre por último
  };

  // Ordena os cabeçalhos baseando-se nos pesos estipulados acima
  const topObs = Object.keys(contagemObs).sort((a, b) => {
    const pesoA = pesosCategorias[a] || 50; // Se a categoria for desconhecida, fica no meio
    const pesoB = pesosCategorias[b] || 50;
    return pesoA - pesoB;
  }).slice(0, 5); // Mantém o limite visual de 5 linhas para o gráfico

  const cabecalhoHist = ["Data", ...topObs];
  
  // Resgata o histórico antigo na coluna A da aba Dash.C.Cir
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

  // Gera o dia de hoje formatado como texto puro para não gerar falsas datas no Sheets
  let hojeStr = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
  let novaLinhaHist = [hojeStr];
  
  // Preenche a linha de hoje cruzando com a ordem definida em topObs
  for (let i = 0; i < 5; i++) {
    let obsKey = topObs[i];
    novaLinhaHist.push(obsKey ? (contagemObs[obsKey] || 0) : 0);
  }

  histData.push(novaLinhaHist);

  // Trava em exatamente 7 atualizações (remove as mais antigas se passar)
  if (histData.length > 7) {
    histData = histData.slice(histData.length - 7);
  }

  // Limpa a região a partir da Coluna A até J para colar os valores ajustados sem deixar resíduos
  abaDash.getRange(1, 1, abaDash.getMaxRows(), 10).clearContent();

  // Escreve cabeçalho e dados formatados
  abaDash.getRange(1, 1, 1, cabecalhoHist.length)
         .setValues([cabecalhoHist])
         .setBackground("#444444")
         .setFontColor("white")
         .setFontWeight("bold");
           
  abaDash.getRange(2, 1, histData.length, 1).setNumberFormat("@"); // Trava a Data como texto
  abaDash.getRange(2, 1, histData.length, novaLinhaHist.length).setValues(histData);

  // =========================================================================
  // 4. CRIAÇÃO DO GRÁFICO NO DASH.C.CIR
  // =========================================================================
  // Remove gráficos anteriores da aba Dash
  const charts = abaDash.getCharts();
  charts.forEach(c => abaDash.removeChart(c));

  // O gráfico vai pegar exatamente a tabela nova gerada na aba Dash
  const rangeGrafico = abaDash.getRange(1, 1, histData.length + 1, cabecalhoHist.length);

  const coresGrafico = ["#D32F2F", "#F57C00", "#FBC02D", "#388E3C", "#1976D2"];

  const chartBuilder = abaDash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeGrafico)
    .setPosition(10, 1, 0, 0) // O gráfico ancorado na Linha 10, Coluna 1 (A)
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('hAxis.textStyle', {fontSize: 11})
    .setOption('title', 'Evolução de Status - Observações (Últimos 7 Dias)')
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