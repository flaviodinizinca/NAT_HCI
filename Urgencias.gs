function criarPlanilhaUrgencias() {
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

  let guiaUrgencias = ss.getSheetByName('Urgências');
  
  // =========================================================================
  // 3. SALVAR NA MEMÓRIA AS INFORMAÇÕES EDITÁVEIS (HCI a HCIV)
  // =========================================================================
  const historicoUrgencias = {};

  if (guiaUrgencias) {
    const ultimaLinha = guiaUrgencias.getLastRow();
    if (ultimaLinha > 1) {
      // Como a estrutura nova vai até a coluna J (índice 10)
      const ultimaColunaParaPreservar = Math.max(10, guiaUrgencias.getLastColumn());
      const dadosExistentes = guiaUrgencias.getRange(2, 1, ultimaLinha - 1, ultimaColunaParaPreservar).getValues();
      for (let i = 0; i < dadosExistentes.length; i++) {
        const itemCode = String(dadosExistentes[i][0]).trim().toUpperCase();
        if (itemCode !== "") {
          historicoUrgencias[itemCode] = {
            hci: dadosExistentes[i][6],   // Coluna G
            hcii: dadosExistentes[i][7],  // Coluna H
            hciii: dadosExistentes[i][8], // Coluna I
            hciv: dadosExistentes[i][9]   // Coluna J
          };
        }
      }
    }
  } else {
    guiaUrgencias = ss.insertSheet('Urgências');
  }

  // =========================================================================
  // 4. LIMPAR A GUIA PRESERVANDO A FORMATAÇÃO DA LINHA 1
  // =========================================================================
  const maxRows = guiaUrgencias.getMaxRows();
  if (maxRows > 1) {
    // Limpa conteúdo e validações antigas apenas da linha 2 para baixo
    guiaUrgencias.getRange(2, 1, maxRows - 1, guiaUrgencias.getMaxColumns()).clearContent();
    guiaUrgencias.getRange(2, 1, maxRows - 1, guiaUrgencias.getMaxColumns()).clearDataValidations();
  }

  const dados = guiaNATHCI.getDataRange().getValues();
  // Novo cabeçalho adaptado de A até J
  const cabecalho = ["Item", "Descrição", "CMM", "Saldo", "Obs", "Validade Ata", "HCI", "HCII", "HCIII", "HCIV"];
  const itensFiltrados = [];

  const parseNumero = (val) => {
    if (val === "" || val == null) return 0;
    if (typeof val === 'number') return val;
    const num = Number(String(val).replace(/\./g, '').replace(',', '.'));
    return isNaN(num) ? 0 : num;
  };

  // =========================================================================
  // 5. FILTRAGEM DINÂMICA
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
      const historicoItem = historicoUrgencias[itemCode] || {};
      
      itensFiltrados.push([
        itemCodeOriginal,
        dados[i][1],
        dados[i][5],
        dados[i][4],
        dados[i][7],
        dados[i][9], // Informação vinda agora da Coluna Y mapeada em Importacao.gs
        historicoItem.hci || "",
        historicoItem.hcii || "",
        historicoItem.hciii || "",
        historicoItem.hciv || ""
      ]);
    }
  }

  // =========================================================================
  // 6. GRAVAÇÃO DOS DADOS E FORMATAÇÕES
  // =========================================================================
  
  // Usamos apenas setValues na linha 1. Assim ele não afeta as cores e formatações manuais que você fez!
  guiaUrgencias.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);

  if (itensFiltrados.length > 0) {
    guiaUrgencias.getRange(2, 1, itensFiltrados.length, cabecalho.length).setValues(itensFiltrados);
    
    // Força a formatação de data na coluna F (índice 6)
    guiaUrgencias.getRange(2, 6, itensFiltrados.length, 1).setNumberFormat("dd/MM/yyyy");
  }

  guiaUrgencias.autoResizeColumns(1, cabecalho.length);
  guiaUrgencias.setFrozenRows(1);
  
  if (itensFiltrados.length > 0) {
    SpreadsheetApp.getUi().alert('Guia de Urgências criada com ' + itensFiltrados.length + ' itens baseados nas suas regras parametrizadas.');
  } else {
    SpreadsheetApp.getUi().alert('Nenhum item atende aos critérios de urgência configurados na aba Regras.');
  }
}