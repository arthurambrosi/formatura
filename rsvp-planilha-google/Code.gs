const RSVP_CONFIG = {
  spreadsheetName: 'RSVP - Formatura de Arthur',
  timezone: 'America/Sao_Paulo',
  sheetConfig: 'Config',
  sheetResumo: 'Resumo',
  sheetSolenidade: 'Solenidade',
  sheetBaile: 'Baile',
  defaultEventName: 'Formatura de Arthur',
  defaultDeadline: '2026-06-01',
  simpleStatusValues: 'confirm | decline | pending',
  combinedStatusValues: 'yes | no | maybe',
  eventSheetHeaders: [
    'Grupo ID',
    'Recebido em',
    'Nome completo',
    'Status',
    'Origem',
    'Pagina',
    'Canal',
    'Payload bruto'
  ]
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RSVP')
    .addItem('Configurar sistema', 'configurarSistemaRSVP')
    .addItem('Liberar visualizacao por link', 'liberarVisualizacaoPorLink')
    .addItem('Liberar edicao publica por link', 'liberarEdicaoPublicaDaPlanilha')
    .addSeparator()
    .addItem('Registrar URL da Web App', 'registrarUrlWebAppViaPrompt')
    .addItem('Mostrar info para HTML', 'mostrarInformacoesParaHtml')
    .addSeparator()
    .addItem('Inserir envio de teste', 'inserirEnvioDeTeste')
    .addToUi();
}

function configurarSistemaRSVP() {
  const contexto = obterOuCriarPlanilha_();
  const ss = contexto.ss;

  if (contexto.criadaNova) {
    ss.rename(RSVP_CONFIG.spreadsheetName);
  }

  ss.setSpreadsheetTimeZone(RSVP_CONFIG.timezone);
  salvarScriptProperties_(ss);
  criarOuAtualizarAbaEvento_(ss, RSVP_CONFIG.sheetSolenidade, '#b91c1c');
  criarOuAtualizarAbaEvento_(ss, RSVP_CONFIG.sheetBaile, '#a16207');
  criarOuAtualizarAbaResumo_(ss);
  liberarVisualizacaoPorLink_();
  criarOuAtualizarAbaConfig_(ss);

  return obterInformacoesParaHtml();
}

function liberarVisualizacaoPorLink() {
  configurarSistemaRSVP();
  return liberarVisualizacaoPorLink_();
}

function liberarEdicaoPublicaDaPlanilha() {
  const ss = obterOuCriarPlanilha_().ss;
  DriveApp.getFileById(ss.getId()).setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  );
  criarOuAtualizarAbaConfig_(ss);
  return ss.getUrl();
}

function registrarUrlWebAppViaPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    'Registrar URL da Web App',
    'Cole a URL final que termina com /exec:',
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  registrarUrlWebApp(resposta.getResponseText());
}

function registrarUrlWebApp(url) {
  const valor = String(url || '').trim();

  if (!valor || valor.indexOf('/exec') === -1) {
    throw new Error('Cole a URL final da Web App, terminando com /exec.');
  }

  const ss = obterOuCriarPlanilha_().ss;
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', valor);
  criarOuAtualizarAbaConfig_(ss);
  return obterInformacoesParaHtml();
}

function mostrarInformacoesParaHtml() {
  const ui = SpreadsheetApp.getUi();
  const info = obterInformacoesParaHtml();
  ui.alert(
    'Info para HTML',
    [
      'URL da Web App: ' + (info.webAppUrl || 'PREENCHER APOS O DEPLOY'),
      '',
      'Pagina simples:',
      '- nome_completo',
      '- evento',
      '- status',
      '',
      'Pagina combinada:',
      '- nome_completo',
      '- solenidade',
      '- baile'
    ].join('\n'),
    ui.ButtonSet.OK
  );
  return info;
}

function obterInformacoesParaHtml() {
  const ss = obterOuCriarPlanilha_().ss;
  salvarScriptProperties_(ss);

  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
  const info = {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    publicSpreadsheetUrl: ss.getUrl(),
    webAppUrl: webAppUrl,
    formMethod: 'POST',
    paginaEventoSimples: {
      action: webAppUrl || 'COLE_AQUI_A_URL_EXEC',
      campos: {
        nome: 'nome_completo',
        evento: 'evento',
        status: 'status',
        origem: 'origem',
        pagina: 'pagina',
        canal: 'canal'
      },
      valores: {
        evento: 'solenidade | baile',
        status: RSVP_CONFIG.simpleStatusValues
      }
    },
    paginaCombinada: {
      action: webAppUrl || 'COLE_AQUI_A_URL_EXEC',
      campos: {
        nome: 'nome_completo',
        solenidade: 'solenidade',
        baile: 'baile',
        origem: 'origem',
        pagina: 'pagina',
        canal: 'canal'
      },
      valores: {
        status: RSVP_CONFIG.combinedStatusValues + ' ou ' + RSVP_CONFIG.simpleStatusValues
      }
    }
  };

  criarOuAtualizarAbaConfig_(ss, info);
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function inserirEnvioDeTeste() {
  const payload = {
    nome_completo: 'Teste Automacao',
    origem: 'teste',
    pagina: 'teste-manual',
    canal: 'apps-script',
    solenidade: 'confirm',
    baile: 'pending'
  };

  registrarSubmissao_(payload);
}

function doGet(e) {
  const formato = (e && e.parameter && e.parameter.format) || '';
  const ping = e && e.parameter && e.parameter.ping === '1';
  const info = obterInformacoesParaHtml();

  if (ping || formato === 'json') {
    return jsonOutput_({
      ok: true,
      status: 'online',
      webAppUrl: info.webAppUrl || '',
      spreadsheetUrl: info.spreadsheetUrl
    });
  }

  return HtmlService.createHtmlOutput(gerarPaginaInfo_(info))
    .setTitle('RSVP - Formatura de Arthur');
}

function doPost(e) {
  try {
    const payload = extrairPayload_(e);
    const resultado = registrarSubmissao_(payload);

    if (querJson_(payload)) {
      return jsonOutput_(resultado);
    }

    return HtmlService.createHtmlOutput(gerarPaginaSucesso_(resultado))
      .setTitle('Confirmacao registrada');
  } catch (error) {
    if (e && e.parameter && e.parameter.format === 'json') {
      return jsonOutput_({ ok: false, erro: error.message });
    }

    return HtmlService.createHtmlOutput(gerarPaginaErro_(error.message))
      .setTitle('Erro ao registrar confirmacao');
  }
}

function registrarSubmissao_(payload) {
  const ss = obterOuCriarPlanilha_().ss;
  prepararEstruturaMinima_(ss);

  const submissao = normalizarPayload_(payload);
  validarSubmissao_(submissao);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const recebidosEm = new Date();
    const eventosRegistrados = [];

    submissao.eventos.forEach(function (evento) {
      const sheet = ss.getSheetByName(nomeDaAbaDoEvento_(evento.evento));
      sheet.appendRow([
        submissao.grupoId,
        recebidosEm,
        submissao.nomeCompleto,
        evento.status,
        submissao.origem,
        submissao.pagina,
        submissao.canal,
        JSON.stringify(submissao.payloadOriginal)
      ]);

      eventosRegistrados.push({
        evento: tituloDoEvento_(evento.evento),
        status: evento.status
      });
    });

    criarOuAtualizarAbaResumo_(ss);
    criarOuAtualizarAbaConfig_(ss);

    return {
      ok: true,
      grupoId: submissao.grupoId,
      nomeCompleto: submissao.nomeCompleto,
      eventos: eventosRegistrados,
      spreadsheetUrl: ss.getUrl()
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizarPayload_(payload) {
  const bruto = payload || {};
  const nomeCompleto = limparTexto_(
    primeiroValor_(
      bruto.nome_completo,
      bruto.nomeCompleto,
      bruto.full_name,
      bruto['full-name'],
      bruto.nome
    )
  );

  const origem = limparTexto_(primeiroValor_(bruto.origem, bruto.site, bruto.projeto, 'site'));
  const pagina = limparTexto_(primeiroValor_(bruto.pagina, bruto.page, ''));
  const canal = limparTexto_(primeiroValor_(bruto.canal, 'site'));
  const grupoId = Utilities.getUuid();
  const eventos = [];

  const eventoUnico = normalizarEvento_(primeiroValor_(bruto.evento, bruto.tipo_evento));
  const statusUnicoBruto = primeiroValor_(bruto.status, bruto.attendance, bruto.resposta);
  const statusUnico = normalizarStatus_(statusUnicoBruto);

  if (eventoUnico && statusUnico) {
    eventos.push({ evento: eventoUnico, status: statusUnico });
  }

  const statusSolenidade = normalizarStatus_(bruto.solenidade);
  if (statusSolenidade) {
    eventos.push({ evento: 'solenidade', status: statusSolenidade });
  }

  const statusBaile = normalizarStatus_(bruto.baile);
  if (statusBaile) {
    eventos.push({ evento: 'baile', status: statusBaile });
  }

  return {
    grupoId: grupoId,
    nomeCompleto: nomeCompleto,
    origem: origem,
    pagina: pagina,
    canal: canal,
    eventos: deduplicarEventos_(eventos),
    payloadOriginal: bruto
  };
}

function validarSubmissao_(submissao) {
  if (!submissao.nomeCompleto) {
    throw new Error('O campo nome_completo e obrigatorio.');
  }

  if (!submissao.eventos.length) {
    throw new Error('Nenhuma resposta valida foi recebida.');
  }
}

function prepararEstruturaMinima_(ss) {
  salvarScriptProperties_(ss);
  criarOuAtualizarAbaEvento_(ss, RSVP_CONFIG.sheetSolenidade, '#b91c1c');
  criarOuAtualizarAbaEvento_(ss, RSVP_CONFIG.sheetBaile, '#a16207');
  criarOuAtualizarAbaResumo_(ss);
  criarOuAtualizarAbaConfig_(ss);
}

function criarOuAtualizarAbaEvento_(ss, nomeAba, tabColor) {
  const sheet = obterOuCriarAba_(ss, nomeAba);
  sheet.setTabColor(tabColor);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, RSVP_CONFIG.eventSheetHeaders.length)
    .setValues([RSVP_CONFIG.eventSheetHeaders])
    .setFontWeight('bold')
    .setBackground('#203934')
    .setFontColor('#f5f1e8');

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 360);
  sheet.getRange('B2:B').setNumberFormat('dd/mm/yyyy hh:mm:ss');

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), RSVP_CONFIG.eventSheetHeaders.length).createFilter();
  }

  aplicarRegrasVisuaisDeStatus_(sheet);
}

function criarOuAtualizarAbaResumo_(ss) {
  const sheet = obterOuCriarAba_(ss, RSVP_CONFIG.sheetResumo);
  sheet.clear();
  sheet.setFrozenRows(2);
  sheet.setTabColor('#065f46');

  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue('Resumo RSVP - Formatura de Arthur');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#04201b').setFontColor('#f5f1e8');

  const headers = [['Evento', 'Confirmados', 'Recusados', 'Pendentes', 'Total respostas', 'Ultima resposta']];
  sheet.getRange(2, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#203934')
    .setFontColor('#f5f1e8');

  sheet.getRange(3, 1, 2, 6).setValues([
    [
      'Solenidade',
      '=COUNTIF(Solenidade!D2:D,"CONFIRMADO")',
      '=COUNTIF(Solenidade!D2:D,"RECUSADO")',
      '=COUNTIF(Solenidade!D2:D,"PENDENTE")',
      '=COUNTA(Solenidade!A2:A)',
      '=IFERROR(MAX(Solenidade!B2:B),"")'
    ],
    [
      'Baile',
      '=COUNTIF(Baile!D2:D,"CONFIRMADO")',
      '=COUNTIF(Baile!D2:D,"RECUSADO")',
      '=COUNTIF(Baile!D2:D,"PENDENTE")',
      '=COUNTA(Baile!A2:A)',
      '=IFERROR(MAX(Baile!B2:B),"")'
    ]
  ]);

  sheet.getRange('F3:F4').setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 180);
}

function criarOuAtualizarAbaConfig_(ss, infoOpcional) {
  const sheet = obterOuCriarAba_(ss, RSVP_CONFIG.sheetConfig);
  sheet.clear();
  sheet.setTabColor('#7c2d12');

  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || 'PREENCHER_APOS_O_DEPLOY';
  const compartilhamento = DriveApp.getFileById(ss.getId()).getSharingAccess();
  const permissao = DriveApp.getFileById(ss.getId()).getSharingPermission();
  const info = infoOpcional || obterInformacoesBasicas_(ss, webAppUrl);

  sheet.getRange('A1:B1').merge();
  sheet.getRange('A1').setValue('Configuracao');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14).setBackground('#04201b').setFontColor('#f5f1e8');

  const configRows = [
    ['Nome do evento', RSVP_CONFIG.defaultEventName],
    ['Spreadsheet ID', ss.getId()],
    ['Spreadsheet URL', ss.getUrl()],
    ['URL publica da planilha', ss.getUrl()],
    ['Acesso publico atual', compartilhamento],
    ['Permissao publica atual', permissao],
    ['URL da Web App', webAppUrl],
    ['Fuso horario', RSVP_CONFIG.timezone],
    ['Aba Solenidade', RSVP_CONFIG.sheetSolenidade],
    ['Aba Baile', RSVP_CONFIG.sheetBaile],
    ['Data limite RSVP', RSVP_CONFIG.defaultDeadline],
    ['Metodo recomendado', 'POST']
  ];

  sheet.getRange(3, 1, configRows.length, 2).setValues(configRows);

  sheet.getRange('D1:E1').merge();
  sheet.getRange('D1').setValue('Para os HTMLs');
  sheet.getRange('D1').setFontWeight('bold').setFontSize(14).setBackground('#04201b').setFontColor('#f5f1e8');

  const htmlRows = [
    ['Action do form', info.webAppUrl || 'COLE_AQUI_A_URL_EXEC'],
    ['Campo nome', 'nome_completo'],
    ['Campo evento simples', 'evento'],
    ['Campo status simples', 'status'],
    ['Valores status simples', RSVP_CONFIG.simpleStatusValues],
    ['Campo origem', 'origem'],
    ['Campo pagina', 'pagina'],
    ['Campo canal', 'canal'],
    ['Campo Solenidade', 'solenidade'],
    ['Campo Baile', 'baile'],
    ['Valores pagina combinada', RSVP_CONFIG.combinedStatusValues + ' ou ' + RSVP_CONFIG.simpleStatusValues],
    ['Evento simples valido', 'solenidade | baile'],
    ['Deploy manual', 'Deploy > New deployment > Web app'],
    ['Quem acessa a Web App', 'Anyone'],
    ['Executar como', 'Voce']
  ];

  sheet.getRange(3, 4, htmlRows.length, 2).setValues(htmlRows);

  sheet.getRange(3, 1, configRows.length, 1).setFontWeight('bold');
  sheet.getRange(3, 4, htmlRows.length, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 340);
}

function aplicarRegrasVisuaisDeStatus_(sheet) {
  const range = sheet.getRange('D2:D');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('CONFIRMADO')
      .setBackground('#dcfce7')
      .setFontColor('#166534')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('RECUSADO')
      .setBackground('#fee2e2')
      .setFontColor('#991b1b')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PENDENTE')
      .setBackground('#fef3c7')
      .setFontColor('#92400e')
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

function obterInformacoesBasicas_(ss, webAppUrl) {
  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    publicSpreadsheetUrl: ss.getUrl(),
    webAppUrl: webAppUrl
  };
}

function obterOuCriarPlanilha_() {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();

  if (ativa) {
    return { ss: ativa, criadaNova: false };
  }

  return {
    ss: SpreadsheetApp.create(RSVP_CONFIG.spreadsheetName),
    criadaNova: true
  };
}

function obterOuCriarAba_(ss, nomeAba) {
  return ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
}

function salvarScriptProperties_(ss) {
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEB_APP_URL') || '';

  props.setProperties({
    SPREADSHEET_ID: ss.getId(),
    SPREADSHEET_URL: ss.getUrl(),
    SHEET_SOLENIDADE: RSVP_CONFIG.sheetSolenidade,
    SHEET_BAILE: RSVP_CONFIG.sheetBaile,
    TIMEZONE: RSVP_CONFIG.timezone,
    WEB_APP_URL: webAppUrl
  }, true);
}

function liberarVisualizacaoPorLink_() {
  const ss = obterOuCriarPlanilha_().ss;
  DriveApp.getFileById(ss.getId()).setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );
  return ss.getUrl();
}

function extrairPayload_(e) {
  if (!e) {
    return {};
  }

  const base = Object.assign({}, e.parameter || {});
  const postData = e.postData && e.postData.contents ? e.postData.contents : '';
  const tipo = e.postData && e.postData.type ? e.postData.type : '';

  if (tipo.indexOf('application/json') !== -1 && postData) {
    return Object.assign(base, JSON.parse(postData));
  }

  return base;
}

function querJson_(payload) {
  const valor = String(
    primeiroValor_(payload.format, payload.response, payload.resposta, '')
  ).toLowerCase();
  return valor === 'json';
}

function deduplicarEventos_(eventos) {
  const mapa = {};

  eventos.forEach(function (item) {
    mapa[item.evento] = item;
  });

  return Object.keys(mapa).map(function (chave) {
    return mapa[chave];
  });
}

function normalizarEvento_(valor) {
  const texto = String(valor || '').trim().toLowerCase();

  if (!texto) return '';
  if (texto === 'solenidade' || texto === 'colacao' || texto === 'colacao_de_grau') return 'solenidade';
  if (texto === 'baile' || texto === 'baile_de_gala') return 'baile';
  return '';
}

function normalizarStatus_(valor) {
  const texto = String(valor || '').trim().toLowerCase();

  if (!texto) return '';

  if ([
    'confirm',
    'confirmed',
    'yes',
    'sim',
    'irei',
    'presente',
    'comparecer'
  ].indexOf(texto) !== -1) {
    return 'CONFIRMADO';
  }

  if ([
    'decline',
    'declined',
    'no',
    'nao',
    'não',
    'ausente',
    'recusar'
  ].indexOf(texto) !== -1) {
    return 'RECUSADO';
  }

  if ([
    'pending',
    'maybe',
    'talvez',
    'depois',
    'pendente'
  ].indexOf(texto) !== -1) {
    return 'PENDENTE';
  }

  return '';
}

function tituloDoEvento_(evento) {
  return evento === 'baile' ? 'Baile de Gala' : 'Solenidade de Colacao de Grau';
}

function nomeDaAbaDoEvento_(evento) {
  return evento === 'baile' ? RSVP_CONFIG.sheetBaile : RSVP_CONFIG.sheetSolenidade;
}

function primeiroValor_() {
  for (var i = 0; i < arguments.length; i += 1) {
    var valor = arguments[i];
    if (valor !== undefined && valor !== null && String(valor).trim() !== '') {
      return valor;
    }
  }

  return '';
}

function limparTexto_(valor) {
  return String(valor || '').trim();
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function gerarPaginaInfo_(info) {
  return [
    '<!DOCTYPE html>',
    '<html lang="pt-BR">',
    '<head>',
    '<meta charset="utf-8"/>',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0"/>',
    '<title>RSVP - Web App</title>',
    '<style>',
    'body{font-family:Arial,sans-serif;background:#001713;color:#cbe9e0;padding:32px;line-height:1.5;}',
    '.card{max-width:780px;margin:0 auto;background:#09241f;border:1px solid #203934;border-radius:16px;padding:24px;}',
    'code{background:#04201b;padding:2px 6px;border-radius:6px;}',
    'a{color:#e9c176;}',
    '</style>',
    '</head>',
    '<body>',
    '<div class="card">',
    '<h1>Web App RSVP online</h1>',
    '<p>Esta Web App recebe os envios dos sites de confirmacao e grava nas abas <strong>Solenidade</strong> e <strong>Baile</strong>.</p>',
    '<p><strong>Planilha:</strong> <a href="' + info.spreadsheetUrl + '" target="_blank">abrir planilha</a></p>',
    '<p><strong>URL atual da Web App:</strong> <code>' + (info.webAppUrl || 'ainda nao registrada na aba Config') + '</code></p>',
    '<h2>Campos esperados</h2>',
    '<p><strong>Pagina simples:</strong> <code>nome_completo</code>, <code>evento</code>, <code>status</code></p>',
    '<p><strong>Pagina combinada:</strong> <code>nome_completo</code>, <code>solenidade</code>, <code>baile</code></p>',
    '</div>',
    '</body>',
    '</html>'
  ].join('');
}

function gerarPaginaSucesso_(resultado) {
  const lista = resultado.eventos.map(function (evento) {
    return '<li><strong>' + evento.evento + ':</strong> ' + evento.status + '</li>';
  }).join('');

  return [
    '<!DOCTYPE html>',
    '<html lang="pt-BR">',
    '<head>',
    '<meta charset="utf-8"/>',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0"/>',
    '<title>Confirmacao registrada</title>',
    '<style>',
    'body{font-family:Arial,sans-serif;background:#001713;color:#cbe9e0;padding:32px;line-height:1.5;}',
    '.card{max-width:720px;margin:0 auto;background:#09241f;border:1px solid #203934;border-radius:18px;padding:28px;}',
    '.btn{display:inline-block;margin-top:20px;background:#e9c176;color:#412d00;text-decoration:none;padding:12px 18px;border-radius:12px;font-weight:bold;}',
    'ul{padding-left:18px;}',
    '</style>',
    '</head>',
    '<body>',
    '<div class="card">',
    '<h1>Confirmacao registrada com sucesso</h1>',
    '<p>Obrigado, <strong>' + escaparHtml_(resultado.nomeCompleto) + '</strong>.</p>',
    '<p>Registramos as seguintes respostas:</p>',
    '<ul>' + lista + '</ul>',
    '<p>Voce ja pode fechar esta pagina ou voltar ao convite.</p>',
    '<a class="btn" href="javascript:history.back()">Voltar</a>',
    '</div>',
    '</body>',
    '</html>'
  ].join('');
}

function gerarPaginaErro_(mensagem) {
  return [
    '<!DOCTYPE html>',
    '<html lang="pt-BR">',
    '<head>',
    '<meta charset="utf-8"/>',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0"/>',
    '<title>Erro ao registrar confirmacao</title>',
    '<style>',
    'body{font-family:Arial,sans-serif;background:#001713;color:#cbe9e0;padding:32px;line-height:1.5;}',
    '.card{max-width:720px;margin:0 auto;background:#09241f;border:1px solid #7f1d1d;border-radius:18px;padding:28px;}',
    '.btn{display:inline-block;margin-top:20px;background:#e9c176;color:#412d00;text-decoration:none;padding:12px 18px;border-radius:12px;font-weight:bold;}',
    '</style>',
    '</head>',
    '<body>',
    '<div class="card">',
    '<h1>Nao foi possivel registrar a confirmacao</h1>',
    '<p>' + escaparHtml_(mensagem) + '</p>',
    '<a class="btn" href="javascript:history.back()">Voltar</a>',
    '</div>',
    '</body>',
    '</html>'
  ].join('');
}

function escaparHtml_(texto) {
  return String(texto || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
