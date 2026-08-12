function parseBRNumber(valor) {
  if (typeof valor === 'number') return valor;
  if (!valor) return 0;
  let str = String(valor).trim();
  if (str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  let num = Number(str);
  return isNaN(num) ? 0 : num;
}

// =========================================================================
// FUNÇÃO AUXILIAR PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================
function importarDadosIndependente() {
  const ssOrigemId = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
  try {
    const ssOrigem = SpreadsheetApp.openById(ssOrigemId);
    const guiaOrigem = ssOrigem.getSheetByName('DadosEstoque');
    const dadosBrutos = guiaOrigem.getDataRange().getDisplayValues();
    
    // Chama a função principal passando os dados capturados
    importarDados(dadosBrutos);
    
    SpreadsheetApp.getUi().alert('Processamento concluído com agrupamento de informações e registro de histórico atualizado!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro ao acessar a base externa.');
  }
}

// =========================================================================
// FUNÇÃO PRINCIPAL REFATORADA (Recebe os dados como parâmetro)
// =========================================================================
function importarDados(dadosEstoqueBrutos) { 
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  const ssStatusId = '1ZLebBqhR1bMZgrnr_dfXikyIY22oi0B2pqXDz1UdRZM';
  
  const ssStatus = SpreadsheetApp.openById(ssStatusId);
  
  const guiaCompilados = ssStatus.getSheetByName('Compilados');
  const guiaDelReport = ssDestino.getSheetByName('Del_Report');
  const guiaNATHCI = ssDestino.getSheetByName('NATHCI') || ssDestino.insertSheet('NATHCI');
  const guiaDash = ssDestino.getSheetByName('Dashboard') || ssDestino.insertSheet('Dashboard');
  const guiaDadosExtras = ssDestino.getSheetByName('DadosExtras');

  if (!guiaCompilados) return;

  const dadosCompilados = guiaCompilados.getDataRange().getValues();
  const statusValidos = ["Pendente com Resíduo", "Pendente", "Rec. Prov. / Com Residuo", "Recebimento Provisório"];
  const mapaEmpenhosValidos = new Set();
  const mapaRecebimento = {}; 

  for (let i = 1; i < dadosCompilados.length; i++) {
    const numEmpenho = String(dadosCompilados[i][0]).trim();
    const codItemComp = String(dadosCompilados[i][5]).trim().toUpperCase(); 
    const qtdRecebida = parseBRNumber(dadosCompilados[i][15]);
    const statusEmpenho = String(dadosCompilados[i][18]).trim();
    
    if (statusValidos.includes(statusEmpenho)) {
      mapaEmpenhosValidos.add(numEmpenho);
    }

    if (statusEmpenho === "Recebimento Provisório" && codItemComp !== "") {
      mapaRecebimento[codItemComp] = (mapaRecebimento[codItemComp] || 0) + qtdRecebida;
    }
  }

  const itensParaExcluir = new Set();
  if (guiaDelReport && guiaDelReport.getLastRow() > 1) {
    const listaExclusao = guiaDelReport.getRange(2, 1, guiaDelReport.getLastRow() - 1, 1).getValues();
    listaExclusao.forEach(row => { if (row[0]) itensParaExcluir.add(String(row[0]).trim()); });
  }

  const itensExtras = new Set();
  if (guiaDadosExtras && guiaDadosExtras.getLastRow() > 1) {
    const listaExtras = guiaDadosExtras.getRange(2, 1, guiaDadosExtras.getLastRow() - 1, 1).getValues();
    listaExtras.forEach(row => { 
      if (row[0]) {
        itensExtras.add(String(row[0]).trim().toUpperCase()); 
      } 
    });
  }

  // Utiliza a matriz recebida por parâmetro (eliminando as duas primeiras linhas de cabeçalho)
  const dadosBrutos = dadosEstoqueBrutos.slice(2);
  const dicionarioItens = {};

  dadosBrutos.forEach(linha => {
    const colA = String(linha[0]).trim();        
    const codItem = String(linha[1]).trim().toUpperCase();     
    const grupoEstoque = String(linha[3]).trim();
    
    const isCondicaoNormal = colA === "ALM" && (codItem.startsWith("A0") || codItem.startsWith("A1") || codItem.startsWith("A2")) && grupoEstoque === "136";
    const isItemExtra = itensExtras.has(codItem);
    
    if ((isCondicaoNormal || isItemExtra) && !itensParaExcluir.has(codItem)) {
      
      const cmmFormatado = parseBRNumber(linha[8]); 
      if (cmmFormatado === 0) return; 
      
      const saldoEstoque = parseBRNumber(linha[7]); 
      const qtdProv = mapaRecebimento[codItem] || 0;
      
      const saldoFinalReal = saldoEstoque + qtdProv;
      
      const obsLower = String(linha[12]).toLowerCase();
      const obsRaw = linha[12];
      
      const procAta = String(linha[19]).trim();
      const procAnalise = String(linha[27]).trim();
      const proc = [procAta, procAnalise].filter(Boolean).join(" / ");
      
      const validadeAta = linha[25]; 
      
      const aeOriginal = String(linha[30]).trim();
      const empOriginal = String(linha[13]).trim();
      
      const aeFiltrada = aeOriginal.startsWith("1") ? aeOriginal : "";
      const empFiltrado = mapaEmpenhosValidos.has(empOriginal) ? empOriginal : "";

      let cat = "";
      if (obsLower.includes("primeira")) cat = "Primeira Compra";
      else if (saldoFinalReal <= 0) cat = "Itens Zerados";
      else if (obsLower.includes("30") && obsLower.includes("59")) cat = "Entre 30 e 59 Dias";
      else if (obsLower.includes("60") && obsLower.includes("89")) cat = "Entre 60 e 89 Dias";
      else if (obsLower.includes("maior") && obsLower.includes("90")) cat = "Igual ou Maior a 90 Dias";
      
      if (cat === "" && isItemExtra) {
        let sd = Number(linha[11]); 
        if (sd <= 0 || isNaN(sd)) cat = "Itens Zerados";
        else if (sd > 0 && sd <= 59) cat = "Entre 30 e 59 Dias";
        else if (sd >= 60 && sd <= 89) cat = "Entre 60 e 89 Dias";
        else cat = "Igual ou Maior a 90 Dias";
      }

      if (cat) {
        if (!dicionarioItens[codItem]) {
          dicionarioItens[codItem] = {
            codItemBase: codItem,
            desc: linha[2],
            grupo: linha[3],
            familia: linha[4], 
            saldo: saldoFinalReal, 
            cmm: linha[8], 
            saldoDias: linha[11], 
            obs: obsRaw, 
            validadeAta: validadeAta,
            processos: new Set(),
            aes: new Set(),
            empenhos: new Set(),
            categoria: cat
          };
        }
        
        if (proc !== "") dicionarioItens[codItem].processos.add(proc);
        if (aeFiltrada !== "") dicionarioItens[codItem].aes.add(aeFiltrada);
        if (empFiltrado !== "") dicionarioItens[codItem].empenhos.add(empFiltrado);
      }
    }
  });

  const contagem = {
    "Itens Zerados": { total: 0, empenho: 0, ae: 0, apenasProcesso: 0, semProcesso: 0, itens: [] },
    "Entre 30 e 59 Dias": { total: 0, empenho: 0, ae: 0, apenasProcesso: 0, semProcesso: 0, itens: [] },
    "Entre 60 e 89 Dias": { total: 0, empenho: 0, ae: 0, apenasProcesso: 0, semProcesso: 0, itens: [] },
    "Igual ou Maior a 90 Dias": { total: 0, empenho: 0, ae: 0, apenasProcesso: 0, semProcesso: 0, itens: [] },
    "Primeira Compra": { total: 0, empenho: 0, ae: 0, apenasProcesso: 0, semProcesso: 0, itens: [] }
  };
  const listaFinal = [];

  Object.values(dicionarioItens).forEach(obj => {
    const strProcessos = Array.from(obj.processos).join('\n');
    const strAes = Array.from(obj.aes).join('\n');
    const strEmpenhos = Array.from(obj.empenhos).join('\n');
    const cat = obj.categoria;

    if (!contagem[cat]) return;

    contagem[cat].total++;
    
    const temEmpenho = obj.empenhos.size > 0;
    const temAe = obj.aes.size > 0;
    const temProcesso = obj.processos.size > 0;

    if (temEmpenho) {
      contagem[cat].empenho++;
    } else if (temAe) {
      contagem[cat].ae++;
    } else if (temProcesso) {
      contagem[cat].apenasProcesso++;
    } else {
      contagem[cat].semProcesso++;
    }

    listaFinal.push([
      "'" + obj.codItemBase, obj.desc, obj.grupo, "'" + obj.familia, obj.saldo, 
      obj.cmm, obj.saldoDias, obj.obs, strProcessos, obj.validadeAta, strAes, strEmpenhos
    ]);

    contagem[cat].itens.push([obj.codItemBase, obj.desc, cat, strEmpenhos, strAes, strProcessos]);
  });

  guiaNATHCI.clear().getRange(1, 1, 1, 12).setValues([["Item", "Descrição", "Grupo", "Família", "Saldo", "CMM", "Saldo Dias", "Obs", "Processo SEI", "Validade Ata", "AE (Filtro 1)", "Empenho (Status)"]])
    .setBackground("#444444").setFontColor("white");
  if (listaFinal.length > 0) {
    const rangeDestino = guiaNATHCI.getRange(2, 1, listaFinal.length, 12);
    rangeDestino.setValues(listaFinal);
    rangeDestino.setWrap(true);
  }

  // --- LÓGICA DE GRAVAÇÃO DO HISTÓRICO ATUALIZADA ---
  const guiaHistorico = ssDestino.getSheetByName('Historico_Resumo') || ssDestino.insertSheet('Historico_Resumo');
  let lastRowHist = guiaHistorico.getLastRow();
  
  if (lastRowHist === 0) {
    guiaHistorico.appendRow(["Data", "Itens Zerados", "Entre 30 e 59 Dias", "Entre 60 e 89 Dias", "Igual ou Maior a 90 Dias", "Primeira Compra"]);
    guiaHistorico.getRange(1, 1, 1, 6).setBackground("#444444").setFontColor("white").setFontWeight("bold");
    lastRowHist = 1;
  }
  
  let dataAtual = new Date();
  dataAtual.setHours(0, 0, 0, 0);
  let strHoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
  
  let linhaParaGravar = lastRowHist + 1;
  
  if (lastRowHist > 1) {
    let valorUltimaData = guiaHistorico.getRange(lastRowHist, 1).getValue();
    let strUltimaData = (valorUltimaData instanceof Date) 
        ? Utilities.formatDate(valorUltimaData, "GMT-3", "dd/MM/yyyy") 
        : String(valorUltimaData);
        
    // Verifica se a data da última linha é igual a hoje
    if (strUltimaData === strHoje) {
      linhaParaGravar = lastRowHist; // Se for igual, sobrescreve a mesma linha
    }
  }
  
  const novaLinha = [
    dataAtual,
    contagem["Itens Zerados"].total,
    contagem["Entre 30 e 59 Dias"].total,
    contagem["Entre 60 e 89 Dias"].total,
    contagem["Igual ou Maior a 90 Dias"].total,
    contagem["Primeira Compra"].total
  ];
  
  guiaHistorico.getRange(linhaParaGravar, 1, 1, 6).setValues([novaLinha]);
  guiaHistorico.getRange(2, 1, Math.max(2, guiaHistorico.getLastRow()), 1).setNumberFormat("dd/MM/yyyy");
  
  if (typeof atualizarDashboardComRelatorio === "function") {
    atualizarDashboardComRelatorio(guiaDash, contagem);
  }
}