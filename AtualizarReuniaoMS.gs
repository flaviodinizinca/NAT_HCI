// =========================================================================
// FUNÇÃO AUXILIAR PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================
function atualizarReuniaoMSIndependente() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  ssLocal.toast('Buscando dados no Estoque Externo...', 'Atualização', 3);

  const idExterna = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';

  try {
    const ssExterna = SpreadsheetApp.openById(idExterna);
    
    // Puxa os dados do Estoque
    const guiaEstoque = ssExterna.getSheetByName('DadosEstoque');
    const dadosEstoqueBrutos = guiaEstoque.getDataRange().getDisplayValues();
    
    // Puxa os dados de Entradas (usando getValues para manipular as datas corretamente)
    const guiaEntradas = ssExterna.getSheetByName('EntradaEmpenhos');
    let dadosEntradasBrutos = [];
    if (guiaEntradas) {
      dadosEntradasBrutos = guiaEntradas.getDataRange().getValues();
    }
    
    atualizarReuniaoMS(dadosEstoqueBrutos, dadosEntradasBrutos);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', 'Sem permissão para acessar a base externa.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// =========================================================================
// FUNÇÃO PRINCIPAL DE ATUALIZAÇÃO DA GUIA
// =========================================================================
function atualizarReuniaoMS(dadosEstoqueBrutos, dadosEntradasBrutos) {
  // Ajuste do nome exato da guia conforme você informou
  const nomeGuia = 'itens avaliados na Reuniao com MS - 12/08/2026';
  
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const guiaDestino = ssLocal.getSheetByName(nomeGuia);

  if (!guiaDestino) {
    SpreadsheetApp.getUi().alert('Erro', `A guia "${nomeGuia}" não foi encontrada! Verifique o nome exato da aba.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // =========================================================================
  // 1. PROCESSAMENTO DE ENTRADAS (Buscando a data mais recente)
  // =========================================================================
  const mapaEntradas = new Map();
  if (dadosEntradasBrutos && dadosEntradasBrutos.length > 0) {
    for (let i = 1; i < dadosEntradasBrutos.length; i++) {
      const codItem = String(dadosEntradasBrutos[i][2]).trim().toUpperCase(); // Coluna C (Índice 2)
      const dataEntradaRaw = dadosEntradasBrutos[i][23]; // Coluna X (Índice 23)

      if (codItem && dataEntradaRaw) {
        let dataEntrada = null;
        
        if (dataEntradaRaw instanceof Date) {
          dataEntrada = dataEntradaRaw;
        } else {
          const stringData = String(dataEntradaRaw).trim();
          
          // Ignora strings de tempo zerado ou valores vazios
          if (stringData !== "" && stringData !== "00:00:00" && stringData !== "0") {
            const parts = stringData.split('/');
            if (parts.length === 3) {
              dataEntrada = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
              dataEntrada = new Date(stringData);
            }
          }
        }
        
        // Validação rigorosa: Apenas datas válidas com ano superior a 1900 (evita 1899)
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
  // 2. PROCESSAMENTO DE ESTOQUE
  // =========================================================================
  const mapaEstoque = new Map();
  for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
    const codItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase(); // Coluna B (Índice 1)
    
    if (codItem) {
      mapaEstoque.set(codItem, {
        estoque: dadosEstoqueBrutos[i][7],  // Coluna H (Índice 7)
        cmm: dadosEstoqueBrutos[i][8],      // Coluna I (Índice 8)
        empenho: dadosEstoqueBrutos[i][16], // Coluna Q (Índice 16) - Atualizado
        ae: dadosEstoqueBrutos[i][33]       // Coluna AH (Índice 33) - Atualizado
      });
    }
  }

  // =========================================================================
  // 3. INJEÇÃO DOS DADOS NAS COLUNAS ALVO (Sem formatar a planilha)
  // =========================================================================
  const ultimaLinha = guiaDestino.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert('Aviso', `Nenhum dado encontrado na aba "${nomeGuia}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Lê apenas a Coluna A, assumindo que os códigos dos itens estejam lá
  const codigos = guiaDestino.getRange(2, 1, ultimaLinha - 1, 1).getDisplayValues();

  // Matrizes isoladas para não sobrescrever colunas adjacentes
  const valoresCMM = [];      // Coluna D
  const valoresAE = [];       // Coluna G
  const valoresEmpenho = [];  // Coluna H
  const valoresEstoque = [];  // Coluna I
  const valoresData = [];     // Coluna M

  for (let i = 0; i < codigos.length; i++) {
    const cod = String(codigos[i][0]).trim().toUpperCase();
    
    if (!cod) {
      valoresCMM.push([""]);
      valoresAE.push([""]);
      valoresEmpenho.push([""]);
      valoresEstoque.push([""]);
      valoresData.push([""]);
      continue;
    }

    // Processamento da Data Mais Recente
    let dataUltimaEntradaStr = "Ainda não houve entrada";
    if (mapaEntradas.has(cod)) {
      const d = mapaEntradas.get(cod);
      const dia = String(d.getDate()).padStart(2, '0');
      const mes = String(d.getMonth() + 1).padStart(2, '0');
      const ano = d.getFullYear();
      dataUltimaEntradaStr = `${dia}/${mes}/${ano}`;
    }
    valoresData.push([dataUltimaEntradaStr]);

    // Processamento das Informações de Estoque
    const info = mapaEstoque.get(cod);
    if (info) {
      valoresCMM.push([info.cmm]);
      valoresAE.push([info.ae]);
      valoresEmpenho.push([info.empenho]);
      valoresEstoque.push([info.estoque]);
    } else {
      // Deixa em branco caso o item não conste na base, mantendo a limpeza da planilha
      valoresCMM.push([""]);
      valoresAE.push([""]);
      valoresEmpenho.push([""]);
      valoresEstoque.push([""]);
    }
  }

  // Escrita em lote isolada para cada coluna específica
  guiaDestino.getRange(2, 4, valoresCMM.length, 1).setValues(valoresCMM);          // Atualiza Coluna D (CMM)
  guiaDestino.getRange(2, 7, valoresAE.length, 1).setValues(valoresAE);            // Atualiza Coluna G (AE / Notes)
  guiaDestino.getRange(2, 8, valoresEmpenho.length, 1).setValues(valoresEmpenho);  // Atualiza Coluna H (Empenho)
  guiaDestino.getRange(2, 9, valoresEstoque.length, 1).setValues(valoresEstoque);  // Atualiza Coluna I (Estoque)
  guiaDestino.getRange(2, 13, valoresData.length, 1).setValues(valoresData);       // Atualiza Coluna M (Última Entrada)

  SpreadsheetApp.getUi().alert('Sucesso', `Valores atualizados na guia "${nomeGuia}" com sucesso!`, SpreadsheetApp.getUi().ButtonSet.OK);
}