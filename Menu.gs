/**
 * Cria o menu personalizado na planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Gestão de Estoque')
    .addItem('▶️ Executar Rotina Completa', 'processarRotinaCompleta')
    .addSeparator()
    .addItem('1. Atualizar Base de Dados', 'importarDadosIndependente')
    .addItem('2. Gerar Planilha de Urgências', 'criarPlanilhaUrgencias')
    .addItem('3. Atualizar Guia HCI', 'atualizarGuiaHC1Independente')
    .addItem('4. Atualizar Centro Cirúrgico', 'atualizarCentroCirurgicoIndependente')
    .addItem('5. Atualizar Fios', 'atualizarFiosIndependente')
    .addItem('6. Atualizar Endoscopia', 'atualizarEndoscopiaIndependente')
    .addSeparator()
    .addItem('7. Enviar para Status Report', 'enviarParaStatusReport')
    .addToUi();
}

/**
 * Função Mestre que orquestra toda a atualização buscando os dados externos APENAS UMA VEZ.
 */
function processarRotinaCompleta() {
  const ui = SpreadsheetApp.getUi();
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  
  ssLocal.toast('Conectando à base externa...', 'Passo 1 de 7', 5);
  
  // =========================================================================
  // 1. CONEXÃO ÚNICA COM A BASE EXTERNA
  // =========================================================================
  const ssOrigemId = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
  let ssOrigem;
  try {
    ssOrigem = SpreadsheetApp.openById(ssOrigemId);
  } catch (e) {
    ui.alert('Erro', 'Sem permissão para acessar a base externa.', ui.ButtonSet.OK);
    return;
  }
  
  const guiaOrigem = ssOrigem.getSheetByName('DadosEstoque');
  if (!guiaOrigem) {
    ui.alert('Erro', 'Aba "DadosEstoque" não encontrada na base externa!', ui.ButtonSet.OK);
    return;
  }
  
  // Puxa toda a matriz de dados de uma vez só (como texto para evitar bugs de data)
  const dadosEstoqueBrutos = guiaOrigem.getDataRange().getDisplayValues();
  
  // =========================================================================
  // 2. DISPARO DAS FUNÇÕES EM SEQUÊNCIA
  // =========================================================================
  
  ssLocal.toast('Atualizando Base de Dados (NATHCI)...', 'Passo 2 de 7', 5);
  importarDados(dadosEstoqueBrutos); 
  
  ssLocal.toast('Gerando Planilha de Urgências...', 'Passo 3 de 7', 5);
  criarPlanilhaUrgencias(); // Esta não precisa receber dados externos, ela lê o NATHCI interno
  
  ssLocal.toast('Atualizando Centro Cirúrgico...', 'Passo 4 de 7', 5);
  atualizarCentroCirurgico(dadosEstoqueBrutos);

  ssLocal.toast('Atualizando Fios Centro Cirúrgico...', 'Passo 5 de 7', 5);
  atualizarFiosCentroCirurgico(dadosEstoqueBrutos);

  ssLocal.toast('Atualizando Endoscopia...', 'Passo 6 de 7', 5);
  atualizarEndoscopia(dadosEstoqueBrutos);
  
  ssLocal.toast('Atualizando Guia HCI...', 'Passo 7 de 7', 5);
  // Esta fica por último pois dispara a tela de confirmação de e-mail ao final
  atualizarGuiaHC1(dadosEstoqueBrutos); 
}