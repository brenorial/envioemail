function enviarEmailSeCondicaoAtendida() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var planilhaAtiva = planilha.getSheetByName('Contratos'); 

  if (!planilhaAtiva) {
    Logger.log('A aba especificada não foi encontrada na planilha.');
    return;
  }

  var dados = planilhaAtiva.getRange('K2:K').getValues(); 
  var dadosColunaC = planilhaAtiva.getRange('C2:C').getValues(); 
  var dadosColunaA = planilhaAtiva.getRange('A2:A').getValues(); 
  var dadosColunaE = planilhaAtiva.getRange('E2:E').getValues(); 
  var dadosColunaB = planilhaAtiva.getRange('B2:B').getValues(); 

  var destinatarios = [
	  'robertorangelalvessilva@gmail.com',
    'douglassouto.riosaude@gmail.com',
    'alessandrasantos@prefeitura.rio',
    'leonardo.calixto@prefeitura.rio',
    'dgovi.inteligencia@gmail.com',
  ];

  var mensagens = []; 

  dados.forEach(function(fila, index) {
    var valorCelulaK = fila[0]; 
    var valorCelulaC = dadosColunaC[index][0]; 
    var valorCelulaA = dadosColunaA[index][0];
    var valorCelulaE = dadosColunaE[index][0];
    var valorCelulaB = dadosColunaB[index][0];
        
    if (valorCelulaK > 0 && valorCelulaK < 180) {
      var mensagem = 
                     'Contrato: ' + valorCelulaA + '<br>' +
                     'Processo: ' + valorCelulaE + '<br>' +
                     'Objeto: ' + valorCelulaB + '<br>' +
                     'Fornecedor: ' + valorCelulaC + '<br>' +
                     'Prazo de Vencimento: ' + valorCelulaK + ' dias' + '<br><br>';
      mensagens.push(mensagem); 
    }
  });

  mensagens.sort(function(a, b) {
    var prazoA = extrairPrazo(a);
    var prazoB = extrairPrazo(b);
    return prazoA - prazoB;
  });

  if (mensagens.length > 0) {
    var assunto = 'Alerta de vencimento de contratos';

    var mensagemHTML = '<html><body>';
    mensagemHTML += '<p>Prezados,</p>'
    mensagemHTML += '<p>Segue abaixo os contratos próximos do prazo de vencimento:</p>';
    mensagemHTML += '<ul>';
    mensagens.forEach(function(mensagem) {
      mensagemHTML += '<li>' + mensagem + '</li>';
    });
    mensagemHTML += '</ul>';
    mensagemHTML += '<p>Atenciosamente,<br>Inteligência - DGOVI</p>';
    mensagemHTML += '<div style="background-color: #ffff00; padding: 5px; border-radius: 5px; width: fit-content;">';
    mensagemHTML += '<p><b>Este é um e-mail automático. Por favor, não responda.</b></p>';
    mensagemHTML += '<p><b>Para assistência, entre em contato diretamente com o responsável pelo setor.</b></p>';
    mensagemHTML += '</div>';
    mensagemHTML += '</body></html>';

    destinatarios.forEach(function(destinatario) {
      GmailApp.sendEmail(destinatario, assunto, '', { htmlBody: mensagemHTML });
    });
  }
}

function extrairPrazo(mensagem) {
  var inicio = mensagem.indexOf('Prazo de Vencimento: ') + 'Prazo de Vencimento: '.length;
  var fim = mensagem.indexOf(' dias', inicio);
  var prazo = mensagem.substring(inicio, fim);
  return parseInt(prazo);
}
