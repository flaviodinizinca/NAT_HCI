function enviarParaStatusReport() {
  const ssUrgencias = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Urgências');
  if (!ssUrgencias || ssUrgencias.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Gere a guia de Urgências primeiro!');
    return;
  }

  const idStatusReport = '1ZLebBqhR1bMZgrnr_dfXikyIY22oi0B2pqXDz1UdRZM';
  try {
    const ssExterna = SpreadsheetApp.openById(idStatusReport);
    const guiaStatus = ssExterna.getSheetByName('Status Report');

    // Pega apenas a coluna A (Itens) da guia Urgências
    const codigos = ssUrgencias.getRange(2, 1, ssUrgencias.getLastRow() - 1, 1).getValues();
    
    // Limpa e insere na externa a partir da linha 2
    guiaStatus.getRange(2, 1, guiaStatus.getLastRow() > 1 ? guiaStatus.getLastRow() : 1, 1).clearContent();
    guiaStatus.getRange(2, 1, codigos.length, 1).setValues(codigos);

    SpreadsheetApp.getUi().alert('Códigos enviados com sucesso para o Status Report!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro ao acessar a planilha externa. Verifique as permissões.');
  }
}