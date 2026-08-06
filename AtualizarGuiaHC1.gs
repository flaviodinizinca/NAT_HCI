// =========================================================================
// FUNÇÃO AUXILIAR PARA RODAR ISOLADAMENTE PELO MENU
// =========================================================================
function atualizarGuiaHC1Independente() {
  const ssOrigemId = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ssOrigem = SpreadsheetApp.openById(ssOrigemId);
    const guiaOrigem = ssOrigem.getSheetByName('DadosEstoque');
    const dadosBrutos = guiaOrigem.getDataRange().getDisplayValues();
    
    atualizarGuiaHC1(dadosBrutos);
  } catch (e) {
    ui.alert('Erro', 'Sem permissão para acessar a base externa.', ui.ButtonSet.OK);
  }
}

// =========================================================================
// FUNÇÃO PRINCIPAL REFATORADA (Recebe os dados como parâmetro)
// =========================================================================
function atualizarGuiaHC1(dadosEstoqueBrutos) {
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  const guiaDestino = ssDestino.getSheetByName('Itens_Prioritarios_HC1');

  if (!guiaDestino) {
    SpreadsheetApp.getUi().alert('A guia "Itens_Prioritarios_HC1" não foi encontrada na planilha atual.');
    return;
  }

  const ultimaLinhaDestino = guiaDestino.getLastRow();

  if (dadosEstoqueBrutos.length < 2 || ultimaLinhaDestino < 2) {
    SpreadsheetApp.getUi().alert('Não há dados suficientes para atualizar a guia Itens_Prioritarios_HC1.');
    return;
  }

  const parseNumero = (valor) => {
    if (valor === '' || valor == null) return 0;
    if (typeof valor === 'number') return valor;
    const texto = String(valor).trim();
    if (!texto) return 0;
    const normalizado = texto.includes(',')
      ? texto.replace(/\./g, '').replace(',', '.')
      : texto;
    const numero = Number(normalizado);
    return isNaN(numero) ? 0 : numero;
  };

  // Mapeia os dados recebidos da base externa
  const mapaEstoque = {};
  for (let i = 2; i < dadosEstoqueBrutos.length; i++) {
    const codigoItem = String(dadosEstoqueBrutos[i][1]).trim().toUpperCase(); // Coluna B (índice 1)
    const valorColunaH = dadosEstoqueBrutos[i][7]; // Coluna H (índice 7)

    if (codigoItem) {
      mapaEstoque[codigoItem] = valorColunaH;
    }
  }

  const dadosDestino = guiaDestino.getRange(2, 1, ultimaLinhaDestino - 1, 6).getValues();
  const valoresColunaF = [];
  const itensRegularizados = [];

  dadosDestino.forEach(linha => {
    const codigo = linha[0]; // Coluna A
    const descricao = linha[1]; // Coluna B
    const cmm = linha[2]; // Coluna C
    const codigoItem = String(codigo).trim().toUpperCase();
    const estoqueAnterior = parseNumero(linha[5]); // Coluna F

    let valorAtualizado = '';

    if (Object.prototype.hasOwnProperty.call(mapaEstoque, codigoItem)) {
      valorAtualizado = mapaEstoque[codigoItem];
    }

    const estoqueAtual = parseNumero(valorAtualizado);

    if (estoqueAtual > estoqueAnterior) {
      itensRegularizados.push({
        codigo: codigo,
        descricao: descricao,
        cmm: cmm,
        estoque: valorAtualizado
      });
    }

    valoresColunaF.push([valorAtualizado]);
  });

  guiaDestino.getRange(2, 6, valoresColunaF.length, 1).setValues(valoresColunaF);

  const mensagemAtualizacao = 'Coluna F da guia Itens_Prioritarios_HC1 atualizada com os dados da base.';

  if (itensRegularizados.length === 0) {
    SpreadsheetApp.getUi().alert(mensagemAtualizacao + ' Nenhum item com aumento de estoque para envio por email.');
    return;
  }

  const props = PropertiesService.getUserProperties();
  props.setProperty('HCI_ITENS_REGULARIZADOS', JSON.stringify(itensRegularizados));

  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            color: #222;
          }
          h3 {
            margin-top: 0;
            font-size: 18px;
          }
          p {
            font-size: 13px;
            line-height: 1.5;
          }
          .actions {
            margin-top: 20px;
            display: flex;
            gap: 10px;
            justify-content: flex-end;
          }
          button {
            border: none;
            border-radius: 6px;
            padding: 10px 14px;
            cursor: pointer;
            font-size: 13px;
          }
          .primary {
            background: #1a73e8;
            color: white;
          }
          .secondary {
            background: #f1f3f4;
            color: #222;
          }
          .danger {
            background: #d93025;
            color: white;
          }
        </style>
      </head>
      <body>
        <h3>Confirmação de envio de email</h3>
        <p>Foram encontrados <strong>${itensRegularizados.length}</strong> itens com aumento de estoque.</p>
        <p>Escolha como deseja enviar o relatório:</p>
        <div class="actions">
          <button class="secondary" onclick="enviar('me')">Para mim mesmo</button>
          <button class="primary" onclick="enviar('todos')">Para Todos</button>
          <button class="danger" onclick="cancelar()">Cancelar</button>
        </div>

        <script>
          function enviar(destino) {
            google.script.run
              .withSuccessHandler(function(mensagem) {
                google.script.host.close();
                alert(mensagem);
              })
              .enviarEmailAtualizacaoHC1(destino);
          }

          function cancelar() {
            google.script.host.close();
          }
        </script>
      </body>
    </html>
  `).setWidth(480).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar relatório HCI');
}

function enviarEmailAtualizacaoHC1(destino) {
  const props = PropertiesService.getUserProperties();
  const dados = props.getProperty('HCI_ITENS_REGULARIZADOS');

  if (!dados) {
    return 'Nenhum dado encontrado para envio do email.';
  }

  const itensRegularizados = JSON.parse(dados);
  if (!itensRegularizados.length) {
    return 'Nenhum item com aumento de estoque para envio por email.';
  }

  const assunto = 'Itens Prioritários Regularizados HCI';
  const paraTodos = 'asilva@inca.gov.br, lsalomao@inca.gov.br';
  const ccTodos = 'jssampaio@inca.gov.br, elmoraes@inca.gov.br';
  const meuEmail = Session.getActiveUser().getEmail();

  let para = '';
  let cc = '';

  if (destino === 'todos') {
    para = paraTodos;
    cc = ccTodos;
  } else if (destino === 'me') {
    para = meuEmail;
    cc = '';
  } else {
    return 'Envio cancelado.';
  }

  const escapeHtml = (valor) => String(valor == null ? '' : valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  const linhasTabela = itensRegularizados.map(item => `
    <tr>
      <td style="padding:8px;border:1px solid #d9d9d9;">${escapeHtml(item.codigo)}</td>
      <td style="padding:8px;border:1px solid #d9d9d9;">${escapeHtml(item.descricao)}</td>
      <td style="padding:8px;border:1px solid #d9d9d9;text-align:center;">${escapeHtml(item.cmm)}</td>
      <td style="padding:8px;border:1px solid #d9d9d9;text-align:center;">${escapeHtml(item.estoque)}</td>
    </tr>
  `).join('');

  const corpoHtml = `
    <div style="font-family:Arial,sans-serif;font-size:13px;color:#222;">
      <p>Prezados,</p>
      <p>Segue abaixo a relação dos itens prioritários do HCI que apresentaram aumento de estoque após a atualização da base.</p>
      <table style="border-collapse:collapse;width:100%;max-width:900px;">
        <thead>
          <tr style="background-color:#f2f2f2;">
            <th style="padding:8px;border:1px solid #d9d9d9;text-align:left;">Código</th>
            <th style="padding:8px;border:1px solid #d9d9d9;text-align:left;">Descrição</th>
            <th style="padding:8px;border:1px solid #d9d9d9;text-align:center;">CMM</th>
            <th style="padding:8px;border:1px solid #d9d9d9;text-align:center;">Estoque</th>
          </tr>
        </thead>
        <tbody>
          ${linhasTabela}
        </tbody>
      </table>
      <p style="margin-top:16px;">Atenciosamente,<br>Controle SA</p>
    </div>
  `;

  MailApp.sendEmail({
    to: para,
    cc: cc,
    subject: assunto,
    htmlBody: corpoHtml
  });

  props.deleteProperty('HCI_ITENS_REGULARIZADOS');
  return 'Email enviado com sucesso.';
}