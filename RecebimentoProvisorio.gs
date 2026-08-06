function verificarRecebimentoProvisorio() {
  const ssAtiva = SpreadsheetApp.getActiveSpreadsheet();
  const guiaUrgencias = ssAtiva.getSheetByName('Urgências');
  
  // Interrompe se a guia Urgências não existir
  if (!guiaUrgencias) {
    SpreadsheetApp.getUi().alert('A guia "Urgências" não foi encontrada.');
    return;
  }

  // 1. Acessar a planilha externa e a guia Material
  const idExterna = '1jXd4uEnyGZvLv4ozDfFi5ZMlumw0TvleGdMKlgGcPtU';
  let ssExterna, guiaMaterial;
  
  try {
    ssExterna = SpreadsheetApp.openById(idExterna);
    guiaMaterial = ssExterna.getSheetByName('Material');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro ao acessar a planilha externa. Verifique se você tem permissão.');
    return;
  }

  if (!guiaMaterial) {
    SpreadsheetApp.getUi().alert('A guia "Material" não foi encontrada na planilha externa.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Verificando status de empenhos na planilha externa...', 'Aguarde');

  // 2. Mapear os itens que têm "Recebimento Provisório"
  const dadosMaterial = guiaMaterial.getDataRange().getValues();
  const itensComRecebimento = new Set();

  for (let i = 1; i < dadosMaterial.length; i++) {
    const codItem = String(dadosMaterial[i][5]).trim(); // Coluna F (índice 5) - Item
    const status = String(dadosMaterial[i][18]).trim(); // Coluna S (índice 18) - Status do Empenho
    
    if (status === "Recebimento Provisório" && codItem !== "") {
      itensComRecebimento.add(codItem);
    }
  }

  // 3. Cruzar com a guia Urgências e preparar a Coluna J
  const dadosUrgencias = guiaUrgencias.getDataRange().getValues();
  const valoresColunaJ = [];
  
  // Define o cabeçalho para a Coluna J (J1)
  valoresColunaJ.push(["Status Externo"]);

  // Analisa a partir da linha 2
  for (let i = 1; i < dadosUrgencias.length; i++) {
    const itemUrgencia = String(dadosUrgencias[i][0]).trim(); // Coluna A (índice 0) - Item
    
    // Verifica se o item atual está no nosso Set de Recebimento Provisório
    if (itensComRecebimento.has(itemUrgencia)) {
      valoresColunaJ.push(["Recebimento Provisório"]);
    } else {
      valoresColunaJ.push([""]); // Deixa em branco se não tiver esse status
    }
  }

  // 4. Escrever os resultados na Coluna J da guia Urgências de uma só vez (mais rápido)
  // Limpa a coluna J primeiro para evitar dados residuais
  guiaUrgencias.getRange(1, 10, guiaUrgencias.getMaxRows(), 1).clearContent();
  
  // Aplica os novos valores
  guiaUrgencias.getRange(1, 10, valoresColunaJ.length, 1).setValues(valoresColunaJ)
               .setVerticalAlignment("middle");
               
  // Formata o cabeçalho (J1) para manter o padrão visual
  guiaUrgencias.getRange(1, 10).setBackground("#CC0000").setFontColor("white").setFontWeight("bold");

  SpreadsheetApp.getActiveSpreadsheet().toast('Verificação concluída com sucesso!', 'Finalizado');
}