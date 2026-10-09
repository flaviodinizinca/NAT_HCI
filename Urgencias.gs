// =========================================================================
// FUNÇÃO AUXILIAR PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================
function criarPlanilhaUrgenciasIndependente() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  ssLocal.toast('Buscando dados no Estoque Externo...', 'Atualização', 3);

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
    
    criarPlanilhaUrgencias(dadosEstoqueBrutos, dadosEntradasBrutos);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', 'Sem permissão para acessar a base externa.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =========================================================================
// FUNÇÃO PRINCIPAL DE CRIAÇÃO DA GUIA URGÊNCIAS
// =========================================================================
function criarPlanilhaUrgencias(dadosEstoqueBrutos, dadosEntradasBrutos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const guiaNATHCI = ss.getSheetByName('NATHCI');

  if (!guiaNATHCI) return;

  // =========================================================================
  // 1. LEITURA DA GUIA DE REGRAS (PARÂMETROS DINÂMICOS)
  // =========================================================================
  const guiaRegras = ss.getSheetByName('Regras');
  let usarDadosExtras = true;
  let aceitarCmmZero = false;
  let limiteDiasUrgencia = 12;
  let termosBloqueados = ["primeira compra", "debito direto"];

  if (guiaRegras) {
    usarDadosExtras = String(guiaRegras.getRange("B2").getValue()).trim().toUpperCase() === "SIM";
    aceitarCmmZero = String(guiaRegras.getRange("B3").getValue()).trim().toUpperCase() === "SIM";
    
    let diasLidos = parseInt(guiaRegras.getRange("B4").getValue());
    if (!isNaN(diasLidos)) limiteDiasUrgencia = diasLidos;

    let textoBloqueados = String(guiaRegras.getRange("B5").getValue()).toLowerCase();
    if (textoBloqueados !== "") {
      termosBloqueados = textoBloqueados.split(",").map(termo => termo.trim()).filter(termo => termo !== "");
    }
  }

  const limiteCmm = aceitarCmmZero ? 0 : 1; 

  // =========================================================================
  // 2. LER A GUIA DADOS EXTRAS
  // =========================================================================
  const itensExtras = new Set();
  if (usarDadosExtras) {
    const guiaDadosExtras = ss.getSheetByName('DadosExtras');
    if (guiaDadosExtras) {
      const maxRow = guiaDadosExtras.getLastRow();
      if (maxRow > 1) {
        const dadosExtras = guiaDadosExtras.getRange(2, 1, maxRow - 1, 1).getValues();
        dadosExtras.forEach(row => {
          if (row[0]) itensExtras.add(String(row[0]).trim().toUpperCase());
        });
      }
    }
  }

  // =========================================================================
  // 3. MAPEAR ESTOQUE E ENTRADAS (Para Saldo Ata e Última Entrada)
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

  const INDICE_SALDO_ATA = 27; // Atualizado para 27 (Coluna AB)
  const mapaEstoque = new Map();

  if (dadosEstoqueBrutos && dadosEstoqueBrutos.length > 0) {
    for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
      const codItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase();
      if (codItem) {
        mapaEstoque.set(codItem, {
          saldoAta: dadosEstoqueBrutos[i][INDICE_SALDO_ATA],
          validadeAta: dadosEstoqueBrutos[i][28] // Atualizado para 28
        });
      }
    }
  }

  // =========================================================================
  // 4. LIMPAR A GUIA PRESERVANDO A FORMATAÇÃO DA LINHA 1
  // =========================================================================
  let guiaUrgencias = ss.getSheetByName('Urgências');
  if (!guiaUrgencias) {
    guiaUrgencias = ss.insertSheet('Urgências');
  }

  const maxRows = guiaUrgencias.getMaxRows();
  if (maxRows > 1) {
    guiaUrgencias.getRange(2, 1, maxRows - 1, guiaUrgencias.getMaxColumns()).clearContent();
    guiaUrgencias.getRange(2, 1, maxRows - 1, guiaUrgencias.getMaxColumns()).clearDataValidations();
  }

  const dados = guiaNATHCI.getDataRange().getValues();

  // Novo cabeçalho alinhado com as suas alterações
  const cabecalho = ["Item", "Descrição", "CMM", "Saldo", "Obs", "Saldo Ata", "Validade Ata", "Última Entrada"];
  const itensFiltrados = [];

  const parseNumero = (val) => {
    if (val === "" || val == null) return 0;
    if (typeof val === 'number') return val;
    const num = Number(String(val).replace(/\./g, '').replace(',', '.'));
    return isNaN(num) ? 0 : num;
  };

  // =========================================================================
  // 5. FILTRAGEM DINÂMICA E CRUZAMENTO DOS DADOS
  // =========================================================================
  for (let i = 1; i < dados.length; i++) {
    const itemCodeOriginal = String(dados[i][0]).trim();
    const itemCode = itemCodeOriginal.toUpperCase();
    
    const saldo = parseNumero(dados[i][4]);          
    const cmm = parseNumero(dados[i][5]);
    const saldoDiasOriginal = dados[i][6];
    const obs = String(dados[i][7]).toLowerCase();
    const categoriaColunaH = String(dados[i][7]).trim().toLowerCase();

    let possuiTermoBloqueado = false;
    for (let t = 0; t < termosBloqueados.length; t++) {
      if (obs.includes(termosBloqueados[t])) {
        possuiTermoBloqueado = true;
        break;
      }
    }

    const isExtra = usarDadosExtras && itensExtras.has(itemCode);
    const isEntre30e59Dias = categoriaColunaH.includes('entre 30 e 59 dias');
    
    let atendeUrgencia = false;
    let saldoDiasCalculado = saldoDiasOriginal !== "" ? parseNumero(saldoDiasOriginal) : null;
    
    if (cmm > 0) {
      saldoDiasCalculado = (saldo / cmm) * 30;
    }

    if (isEntre30e59Dias) {
      atendeUrgencia = true;
    } else if (isExtra) {
      atendeUrgencia = true;
    } else if (!possuiTermoBloqueado) {
      if (cmm >= limiteCmm && saldo <= 0) {
        atendeUrgencia = true;
      } else if (saldoDiasCalculado !== null && saldoDiasCalculado <= limiteDiasUrgencia && cmm >= limiteCmm) {
        atendeUrgencia = true;
      } else if (saldoDiasOriginal !== "" && parseNumero(saldoDiasOriginal) <= limiteDiasUrgencia) {
        atendeUrgencia = true;
      }
    }

    if (atendeUrgencia) {
      let saldoAta = "";
      let valAta = "";
      let dataUltimaEntradaStr = "Ainda não houve entrada";
      
      if (mapaEstoque.has(itemCode)) {
        saldoAta = mapaEstoque.get(itemCode).saldoAta;
        valAta = mapaEstoque.get(itemCode).validadeAta;
      }
      
      if (mapaEntradas.has(itemCode)) {
        const d = mapaEntradas.get(itemCode);
        const dia = String(d.getDate()).padStart(2, '0');
        const mes = String(d.getMonth() + 1).padStart(2, '0');
        const ano = d.getFullYear();
        dataUltimaEntradaStr = `${dia}/${mes}/${ano}`;
      }

      itensFiltrados.push([
        itemCodeOriginal,
        dados[i][1], // Descrição
        dados[i][5], // CMM
        dados[i][4], // Saldo
        dados[i][7], // Obs
        saldoAta,    // Saldo Ata (F)
        valAta,      // Validade Ata (G)
        dataUltimaEntradaStr // Última Entrada (H)
      ]);
    }
  }

  // =========================================================================
  // 6. GRAVAÇÃO DOS DADOS E FORMATAÇÃO
  // =========================================================================
  guiaUrgencias.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);

  if (itensFiltrados.length > 0) {
    guiaUrgencias.getRange(2, 1, itensFiltrados.length, cabecalho.length).setValues(itensFiltrados);
    
    // Força a formatação de data na coluna G (Validade Ata - índice 7)
    // A coluna H (Última Entrada) já está indo como texto formatado dd/mm/yyyy
    guiaUrgencias.getRange(2, 7, itensFiltrados.length, 1).setNumberFormat("dd/MM/yyyy");
  }

  guiaUrgencias.autoResizeColumns(1, cabecalho.length);
  guiaUrgencias.setFrozenRows(1);

  if (itensFiltrados.length > 0) {
    SpreadsheetApp.getUi().alert('Guia de Urgências atualizada com ' + itensFiltrados.length + ' itens.');
  } else {
    SpreadsheetApp.getUi().alert('Nenhum item atende aos critérios de urgência configurados na aba Regras.');
  }
}