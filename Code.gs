// ═══════════════════════════════════════════════════════
// NEXUS CLIN — Automação Clínica v18
// Google Apps Script v19
// ═══════════════════════════════════════════════════════

// ── CONSTANTES ──────────────────────────────────────────
const props        = PropertiesService.getScriptProperties();
const FEEGOW_TOKEN = props.getProperty('FEEGOW_TOKEN') || '';
const ANTHROPIC_KEY = props.getProperty('ANTHROPIC_KEY') || '';
const EMAIL_CLINICA      = 'nexusclinpb@gmail.com';
const EMAIL_PROFISSIONAL = 'sosthenes53@gmail.com';
const EMAIL_JULIA        = 'nutri.juliaquiteria@gmail.com';
const CLINICA            = 'NEXUS CLIN';
const PROFISSIONAL       = 'Sosthenes dos Santos Alves';
const COREN              = 'COREN-PB 568176';
const CARGO              = 'Enfermeiro de Prática Avançada';
const CIDADE             = 'Livramento-PB';
const TELEFONE           = '(83) 9 9858-5691';

// ═══════════════════════════════════════════════════════
// FUNÇÃO 0 — salvarNaPlanilha(dados)
// ═══════════════════════════════════════════════════════
function salvarNaPlanilha(dados) {
  try {
    var ss;
    var spreadsheetId = PropertiesService
      .getScriptProperties()
      .getProperty('SPREADSHEET_ID');

    if (!spreadsheetId) {
      // Cria a planilha pela primeira vez
      ss = SpreadsheetApp.create('NEXUS CLIN — Anamneses');
      PropertiesService.getScriptProperties()
        .setProperty('SPREADSHEET_ID', ss.getId());
      Logger.log('📊 Planilha criada: ' + ss.getId());

      // Compartilha com Sosthenes e Julia (só na criação)
      try {
        var arquivo = DriveApp.getFileById(ss.getId());
        arquivo.addEditor('sosthenes53@gmail.com');
        arquivo.addEditor('nutri.juliaquiteria@gmail.com');
        Logger.log('✓ Planilha compartilhada com Sosthenes e Julia');
      } catch (shareErr) {
        Logger.log('⚠️ Erro ao compartilhar: ' + shareErr.message);
      }

    } else {
      ss = SpreadsheetApp.openById(spreadsheetId);
    }

    // Aba do mês atual ex: "Mai/2026"
    var mesAno = Utilities.formatDate(
      new Date(), 'America/Recife', 'MMM/yyyy'
    );
    var aba = ss.getSheetByName(mesAno);

    if (!aba) {
      aba = ss.insertSheet(mesAno);

      // Cabeçalho
      var cabecalho = [
        'Data/Hora', 'Nome', 'CPF', 'Nascimento',
        'Celular', 'Email', 'Sexo', 'Cidade', 'Estado',
        'Objetivo', 'Disposição', 'Qualidade Vida', 'Estresse',
        'Sint. Gerais', 'Sint. Emocionais',
        'Medicamentos', 'Suplementos',
        'Pratica Exercício', 'Tipo Exercício',
        'Horas Sono', 'Agua', 'Peso Atual', 'Altura',
        'Comportamento Alimentar', 'Intolerância',
        'Observações Finais', 'Feegow ID', 'JSON Completo'
      ];
      aba.appendRow(cabecalho);
      aba.getRange(1, 1, 1, cabecalho.length)
        .setBackground('#0B1F3A')
        .setFontColor('#C9A84C')
        .setFontWeight('bold');
      aba.setFrozenRows(1);
    }

    // Linha de dados
    var linha = [
      Utilities.formatDate(new Date(), 'America/Recife', 'dd/MM/yyyy HH:mm:ss'),
      dados.nome                    || '',
      dados.cpf                     || '',
      dados.data_nascimento         || '',
      dados.whatsapp               || '',
      dados.email                   || '',
      dados.sexo                    || '',
      dados.cidade                  || '',
      dados.estado                  || '',
      dados.objetivo                || '',
      dados.disposicao              || '',
      dados.qualidade_vida          || '',
      dados.estresse                || '',
      dados.sint_gerais             || '',
      dados.sint_emocionais         || '',
      dados.medicamentos            || '',
      dados.suplementos             || '',
      dados.pratica_exercicio       || '',
      dados.tipo_exercicio          || '',
      dados.horas_sono              || '',
      dados.agua                    || '',
      dados.peso_atual              || '',
      dados.altura                  || '',
      dados.comportamento_alimentar || '',
      dados.intolerancia            || '',
      dados.observacoes_finais      || '',
      dados.feegow_id            || '',
      JSON.stringify(dados)
    ];

    aba.appendRow(linha);
    Logger.log('✓ Dados salvos na planilha — ' + dados.nome);

  } catch (err) {
    // Nunca interrompe o fluxo principal
    Logger.log('⚠️ Erro ao salvar planilha: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 1 — doPost(e) — WEBHOOK PRINCIPAL
// ═══════════════════════════════════════════════════════
function doPost(e) {
  var resp = {};
  try {
    var dados;

    // Tenta JSON direto primeiro
    try {
      dados = JSON.parse(e.postData.contents);
    } catch(jsonErr) {
      // Fallback: form-data com campo "payload"
      var payload = e.parameter.payload || e.parameters.payload;
      if (payload) {
        dados = JSON.parse(Array.isArray(payload) ? payload[0] : payload);
      }
    }

    // Validação básica
    if (!dados || !dados.nome) {
      Logger.log('✗ Payload vazio ou inválido: ' + JSON.stringify(e.postData));
      throw new Error('Payload inválido: nome obrigatório');
    }

    Logger.log('📥 Dados recebidos: ' + JSON.stringify(dados));

    // ── ETAPA 0: Salvar na planilha (sempre primeiro) ──
    Logger.log('⏳ Etapa 0: salvarNaPlanilha');
    salvarNaPlanilha(dados);
    Logger.log('✓ ETAPA 0: dados salvos na planilha');

    // ── ETAPA 0b: Cadastrar no Feegow ──
    var pacienteId = 'NAO_CADASTRADO';
    try {
      pacienteId = cadastrarNoFeegow(dados);
      Logger.log('✓ ETAPA 0b: paciente_id Feegow: ' + pacienteId);
    } catch (feegowErr) {
      Logger.log('⚠️ ETAPA 0b: Feegow falhou (continua): ' + feegowErr.message);
    }
    dados.feegow_id = pacienteId;

    // ── ETAPA 1: Gerar resumo com Claude ──
    Logger.log('⏳ Etapa 1: gerarResumoClaude');
    var resumo;
    try {
      resumo = gerarResumoClaude(dados);
    } catch (err) {
      Logger.log('⚠️ Claude falhou — usando fallback: ' + err.message);
      resumo = 'Olá! Estamos muito felizes em receber você na NEXUS CLIN. ' +
               'Seu protocolo personalizado será construído com cuidado e dedicação para você.';
    }
    Logger.log('✓ Etapa 1 concluída');

    // ── ETAPA 2: Gerar PDF de requisição ──
    Logger.log('⏳ Etapa 2: gerarPDFRequisicao');
    var pdfBlob = null;
    try {
      pdfBlob = gerarPDFRequisicao(dados);
      Logger.log('✓ Etapa 2 concluída');
    } catch (err) {
      Logger.log('⚠️ Erro ao gerar PDF (não crítico): ' + err.message);
    }

    // ── ETAPA 3: Enviar e-mails ──
    Logger.log('⏳ Etapa 3: enviarEmails');
    enviarEmails(dados, resumo, pdfBlob);
    Logger.log('✓ Etapa 3 concluída');

    resp = {
      success: true,
      emails_enviados: true,
      paciente: dados.nome
    };

  } catch (err) {
    var msg = err.message || String(err);
    Logger.log('✗ ERRO: ' + msg);
    resp = { success: false, error: msg };
  }

  return ContentService
    .createTextOutput(JSON.stringify(resp))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 2 — gerarResumoClaude(dados)
// ═══════════════════════════════════════════════════════
function gerarResumoClaude(dados) {
  var FALLBACK =
    'Olá! Estamos muito felizes em receber você na NEXUS CLIN. ' +
    'Seu protocolo personalizado será construído com cuidado e dedicação para você.';

  try {
    var primeiroNome = dados.nome.split(' ')[0];

    var promptUsuario =
      'Paciente: ' + primeiroNome + '\n' +
      'Objetivo(s): ' + (dados.objetivo || 'Não informado') + '\n' +
      'Disposição/energia (1-5): ' + (dados.disposicao || '?') + '\n' +
      'Nível de estresse (1-5): ' + (dados.estresse || '?') + '\n' +
      'Qualidade de vida (1-5): ' + (dados.qualidade_vida || '?') + '\n' +
      'Saúde mental (1-5): ' + (dados.saude_mental || '?') + '\n' +
      'Horas de sono/noite: ' + (dados.horas_sono || '?') + '\n' +
      'Sintomas gerais: ' + (dados.sint_gerais || 'Nenhum relatado') + '\n' +
      'Sintomas emocionais: ' + (dados.sint_emocionais || 'Nenhum relatado') + '\n' +
      'Diagnósticos metabólicos: ' + (dados.diag_metab || 'Nenhum informado') + '\n' +
      'Medicamentos: ' + (dados.medicamentos || 'Nenhum') + '\n' +
      'Principais desafios: ' + (dados.desafios || 'Não informado') + '\n\n' +
      'Escreva um parágrafo acolhedor e empático (máximo 4 frases), iniciando pelo ' +
      'primeiro nome do paciente (' + primeiroNome + '), que apresente brevemente o ' +
      'panorama clínico real (use os dados acima) e termine com uma frase de encorajamento ' +
      'sobre o início do protocolo personalizado na NEXUS CLIN.';

    var apiBody = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 500,
      system:
        'Você é um enfermeiro de prática avançada especializado em medicina metabólica ' +
        'e longevidade. Escreva em português brasileiro. Tom acolhedor, empático e ' +
        'profissional. Sem jargão sem explicação. Máximo 4 frases. Inicie pelo primeiro nome.',
      messages: [{ role: 'user', content: promptUsuario }]
    };

    var options = {
      method: 'POST',
      headers: {
        'x-api-key': ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(apiBody),
      muteHttpExceptions: true,
      timeout: 30000
    };

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var statusCode = response.getResponseCode();
    var bodyText = response.getContentText();
    Logger.log('📥 Claude [' + statusCode + ']: ' + bodyText.substring(0, 200));

    if (statusCode === 200) {
      var content = JSON.parse(bodyText);
      if (content.content && content.content[0] && content.content[0].text) {
        return content.content[0].text;
      }
    }

    Logger.log('⚠️ Claude falhou — usando fallback');
    return FALLBACK;

  } catch (err) {
    Logger.log('⚠️ Erro em gerarResumoClaude: ' + err.message);
    return FALLBACK;
  }
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 3 — gerarPDFRequisicao(dados)
// ═══════════════════════════════════════════════════════
function gerarPDFRequisicao(dados) {
  var primeiroNome = dados.nome.split(' ')[0];
  var docId = null;
  var pdfBlob = null;

  var exames;
  if (dados.sexo === 'Masculino') {
    exames = [
      'Glicemia de Jejum',
      'Insulina de Jejum',
      'HbA1c',
      'TSH + T4 Livre',
      'Hemograma Completo',
      'Ferritina',
      'Vitamina D (25-OH)',
      'Cortisol Matinal',
      'Perfil Lipídico Completo',
      'TGO + TGP',
      'Testosterona Total + SHBG'
    ];
  } else {
    exames = [
      'Glicemia de Jejum',
      'Insulina de Jejum',
      'HbA1c',
      'TSH + T4 Livre',
      'Hemograma Completo',
      'Ferritina',
      'Vitamina D (25-OH)',
      'Cortisol Matinal',
      'Perfil Lipídico Completo',
      'TGO + TGP',
      'Testosterona Total + SHBG + DHEA-S + LH + FSH'
    ];
  }

  var hoje = Utilities.formatDate(new Date(), 'America/Recife', 'dd/MM/yyyy');

  var doc = DocumentApp.create('nexus_req_' + Date.now());
  docId = doc.getId();
  var body = doc.getBody();
  body.setMarginTop(56);
  body.setMarginBottom(56);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  var t1 = body.appendParagraph(CLINICA);
  t1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t1.getChild(0).asText().setFontSize(18).setBold(true).setForegroundColor('#0B1F3A');

  var t2 = body.appendParagraph('Centro de Performance Metabólica & Longevidade');
  t2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t2.getChild(0).asText().setFontSize(11).setBold(false);

  var t3 = body.appendParagraph(PROFISSIONAL + ' | ' + COREN);
  t3.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t3.getChild(0).asText().setFontSize(10);

  var t4 = body.appendParagraph(CARGO + ' | ' + CIDADE);
  t4.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t4.getChild(0).asText().setFontSize(10);

  var t5 = body.appendParagraph(TELEFONE + ' | ' + EMAIL_CLINICA);
  t5.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t5.getChild(0).asText().setFontSize(10);

  body.appendHorizontalRule();

  var pn = body.appendParagraph('Paciente: ' + dados.nome);
  pn.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pn.getChild(0).asText().setFontSize(11);

  var pd = body.appendParagraph('Data de Emissão: ' + hoje);
  pd.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pd.getChild(0).asText().setFontSize(11);

  body.appendHorizontalRule();

  var tr = body.appendParagraph('Painel Metabólico NEXUS — Bloqueios ao Emagrecimento');
  tr.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  tr.getChild(0).asText().setFontSize(13).setBold(true).setForegroundColor('#0B1F3A');

  body.appendParagraph('');

  for (var i = 0; i < exames.length; i++) {
    var item = body.appendParagraph((i + 1) + '. ' + exames[i]);
    item.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    item.getChild(0).asText().setFontSize(11);
  }

  body.appendParagraph('');

  var just =
    'Solicitação de mapeamento metabólico direcionado à identificação de resistência ' +
    'insulínica, disfunção tireoidiana, hipovitaminoses, sobrecarga hepática e ' +
    'desequilíbrio hormonal — condições de alta prevalência nessa população e com ' +
    'impacto direto na resposta ao emagrecimento.';

  var pj = body.appendParagraph(just);
  pj.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pj.getChild(0).asText().setFontSize(10).setItalic(true);

  var pc = body.appendParagraph('CID: E66.9');
  pc.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pc.getChild(0).asText().setFontSize(10).setItalic(true);

  body.appendHorizontalRule();

  var la = body.appendParagraph('_____________________________________________');
  la.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  la.getChild(0).asText().setFontSize(11);

  var pp = body.appendParagraph(PROFISSIONAL);
  pp.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pp.getChild(0).asText().setFontSize(11).setBold(true);

  var pk = body.appendParagraph(COREN + ' | ' + CARGO);
  pk.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pk.getChild(0).asText().setFontSize(11);

  var pl = body.appendParagraph(CIDADE + ' | ' + hoje);
  pl.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pl.getChild(0).asText().setFontSize(11);

  doc.saveAndClose();

  try {
    var exportUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
    var pdfResp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    pdfBlob = pdfResp.getBlob().setName('Painel_Metabolico_NEXUS_' + primeiroNome + '.pdf');
  } finally {
    if (docId) {
      try {
        DriveApp.getFileById(docId).setTrashed(true);
        Logger.log('🗑️ Doc temporário removido');
      } catch (delErr) {
        Logger.log('⚠️ Não foi possível remover doc: ' + delErr.message);
      }
    }
  }

  return pdfBlob;
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 4 — enviarEmails(dados, resumo, pdfBlob)
// ═══════════════════════════════════════════════════════
function enviarEmails(dados, resumo, pdfBlob) {
  var primeiroNome = dados.nome.split(' ')[0];
  var dataHora = Utilities.formatDate(new Date(), 'America/Recife', "dd/MM/yyyy 'às' HH:mm");

  var htmlPaciente =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#f5f5f5;margin:0}' +
    '.wrap{max-width:600px;margin:0 auto;background:#fff}' +
    '.hdr{background:#0B1F3A;padding:32px 24px;text-align:center}' +
    '.hdr h1{color:#fff;font-size:26px;margin:0;letter-spacing:3px}' +
    '.hdr p{color:#C9A84C;margin:6px 0 0;font-size:12px}' +
    '.bod{padding:32px 24px}' +
    '.ola{font-size:20px;color:#0B1F3A;font-weight:bold;margin-bottom:16px}' +
    '.resumo{border-left:4px solid #C9A84C;background:#FAF7F2;padding:16px 20px;margin:24px 0;border-radius:0 8px 8px 0}' +
    '.resumo h3{color:#0B1F3A;margin:0 0 10px;font-size:13px;text-transform:uppercase;letter-spacing:1px}' +
    '.resumo p{color:#333;line-height:1.7;margin:0;font-size:15px}' +
    '.anexo{background:#e8f4f0;border-radius:8px;padding:16px 20px;margin:20px 0;color:#0F6E56;font-size:14px}' +
    '.ass{margin-top:32px;border-top:1px solid #eee;padding-top:20px;color:#555;font-size:13px;line-height:1.8}' +
    '.ftr{background:#0B1F3A;padding:16px 24px;text-align:center;color:#888;font-size:11px}' +
    '.ftr span{color:#C9A84C}' +
    '</style></head>' +
    '<body><div class="wrap">' +
    '<div class="hdr"><h1>NEXUS CLIN</h1>' +
    '<p>Centro de Performance Metab&oacute;lica &amp; Longevidade</p></div>' +
    '<div class="bod">' +
    '<p class="ola">Ol&aacute;, ' + primeiroNome + '! &#128075;</p>' +
    '<p style="color:#444;font-size:15px;line-height:1.7">É com muito prazer que recebemos você na ' +
    '<strong>NEXUS CLIN</strong>. Nossa missão é oferecer um cuidado personalizado, baseado em evidências, ' +
    'focado no seu equilíbrio metabólico e qualidade de vida a longo prazo.</p>' +
    '<div class="resumo"><h3>Seu panorama inicial</h3><p>' + resumo + '</p></div>';

  var semExames = (dados.tem_exames === 'Não tenho exames recentes');

  if (pdfBlob && semExames) {
    htmlPaciente += '<div class="anexo">📄 <strong>Sua solicitação de exames está em anexo.</strong><br>' +
      'Realize a coleta o quanto antes para que possamos iniciar seu protocolo personalizado com precisão.</div>';
  }

  htmlPaciente +=
    '<p style="color:#555;font-size:14px;line-height:1.7">Em caso de dúvidas, entre em contato pelo WhatsApp ' +
    'ou aguarde nosso retorno após análise dos seus exames. Estamos com você em cada etapa dessa jornada. &#128154;</p>' +
    '<div class="ass"><strong>Sosthenes dos Santos Alves</strong><br>' +
    'COREN-PB 568176 | Enfermeiro de Prática Avançada<br>' +
    'NEXUS CLIN — Livramento-PB<br>&#128241; (83) 9 9858-5691<br>&#9993; nexusclinpb@gmail.com</div>' +
    '</div>' +
    '<div class="ftr"><span>NEXUS CLIN</span> — Livramento-PB<br>' +
    'Este e-mail contém informações confidenciais de saúde.</div>' +
    '</div></body></html>';

  try {
    var anexosPaciente = [];
    if (pdfBlob && semExames) anexosPaciente.push(pdfBlob);
    GmailApp.sendEmail(
      dados.email,
      'Bem-vindo(a) à NEXUS CLIN',
      'Visualize em HTML para melhor experiência.',
      {
        name: 'NEXUS CLIN',
        htmlBody: htmlPaciente,
        attachments: anexosPaciente
      }
    );
    Logger.log('✓ E-mail paciente enviado para ' + dados.email);
  } catch (e) {
    Logger.log('✗ ERRO e-mail paciente: ' + e.message);
    throw e;
  }

  var nascFormatado = dados.data_nascimento || '';
  if (nascFormatado.indexOf('-') !== -1) {
    var partes = nascFormatado.split('-');
    if (partes.length === 3) nascFormatado = partes[2] + '/' + partes[1] + '/' + partes[0];
  }

  var camposIgnorar = ['nome', 'cpf', 'email', 'whatsapp',
    'data_nascimento', 'cep', 'endereco', 'numero_endereco', 'complemento', 'bairro',
    'cidade', 'estado', 'sexo', 'timestamp', 'feegow_id'];

  var labelCampo = {
    objetivo:              'Objetivos',
    prazo:                 'Prazo desejado',
    desafios:              'Principais desafios',
    percepcao_corpo:       'Percepção corporal',
    variacao_peso:         'Variação de peso',
    rotina_trabalho:       'Rotina de trabalho',
    disposicao:            'Disposição/energia (1-5)',
    estresse:              'Estresse (1-5)',
    saude_mental:          'Saúde mental (1-5)',
    saude_fisica:          'Saúde física (1-5)',
    qualidade_vida:        'Qualidade de vida (1-5)',
    horas_sono:            'Horas de sono/noite',
    queixas_sono:          'Queixas de sono',
    sonolencia:            'Sonolência diurna',
    pratica_exercicio:     'Pratica exercício?',
    tipo_exercicio:        'Tipo de exercício',
    freq_exercicio:        'Frequência (dias/sem)',
    duracao_treino:        'Duração treino (min)',
    agua:                  'Consumo de água/dia',
    alcool:                'Consumo de álcool',
    tabagismo:             'Tabagismo',
    padrao_alimentar:      'Padrão alimentar',
    intolerancia:          'Intolerâncias alimentares',
    comportamento_alimentar: 'Comportamento alimentar',
    sint_digestivos:       'Sintomas digestivos',
    sint_gerais:           'Sintomas gerais',
    sint_emocionais:       'Sintomas emocionais',
    sint_neuro:            'Sintomas neurológicos',
    sint_cardio:           'Sintomas cardiovasculares',
    sint_derm:             'Sintomas dermatológicos',
    diag_cardio:           'Diagnósticos cardiovasculares',
    diag_metab:            'Diagnósticos metabólicos',
    diag_digest:           'Diagnósticos digestivos',
    diag_musculo:          'Diagnósticos músculo-esqueléticos',
    diag_autoimune:        'Diagnósticos autoimunes',
    diag_psiq:             'Diagnósticos psiquiátricos',
    medicamentos:          'Medicamentos em uso',
    suplementos:           'Suplementos em uso',
    medicamentos_emagrec:  'Medicamentos p/ emagrecimento',
    libido:                'Libido',
    tem_exames:            'Exames de sangue recentes',
    observacoes_finais:    'Observações finais'
  };

  var linhasAnamnese = '';
  Object.keys(dados).forEach(function(campo) {
    if (camposIgnorar.indexOf(campo) === -1) {
      var valor = dados[campo];
      if (valor !== undefined && valor !== null && String(valor).trim() !== '') {
        linhasAnamnese +=
          '<tr>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:bold;color:#0B1F3A;width:38%">' +
          (labelCampo[campo] || campo.replace(/_/g,' ')) +
          '</td>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #eee;color:#333">' +
          String(valor) +
          '</td>' +
          '</tr>';
      }
    }
  });

  // --- Triagem rapida: PHQ (humor) / GAD (ansiedade) ---
  var tagHumor   = dados.phq_1 ? dados.phq_1.split(' - ')[0] : '-';
  var tagAnsied  = dados.gad_1 ? dados.gad_1.split(' - ')[0] : '-';
  var tagTriagem = '[PHQ:' + tagHumor + ' | GAD:' + tagAnsied + ']';

  var htmlInterno =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#f0f0f0;margin:0}' +
    '.wrap{max-width:700px;margin:0 auto;background:#fff}' +
    '.hdr{background:#0B1F3A;padding:20px 24px}' +
    '.hdr h2{color:#C9A84C;margin:0;font-size:16px}' +
    '.hdr p{color:#aaa;margin:4px 0 0;font-size:12px}' +
    '.bod{padding:24px}' +
    'table{width:100%;border-collapse:collapse;font-size:14px;margin-bottom:24px}' +
    'tr:nth-child(even){background:#fafafa}' +
    'td{padding:10px 12px;border-bottom:1px solid #eee;color:#333}' +
    'td:first-child{font-weight:bold;color:#0B1F3A;width:38%}' +
    'h3{color:#0B1F3A;margin:24px 0 12px;font-size:14px;border-bottom:2px solid #C9A84C;padding-bottom:6px}' +
    '.ftr{background:#f5f5f5;padding:12px 24px;text-align:center;color:#999;font-size:11px}' +
    '</style></head>' +
    '<body><div class="wrap">' +
    '<div class="hdr"><h2>&#128276; Novo paciente aguardando atendimento ' + tagTriagem + '</h2>' +
    '<p>Recebido em ' + dataHora + '</p></div>' +
    '<div class="bod">' +
    '<h3>Dados Pessoais</h3>' +
    '<table>' +
    '<tr><td>Nome completo</td><td>' + (dados.nome || '') + '</td></tr>' +
    '<tr><td>Data de nascimento</td><td>' + nascFormatado + '</td></tr>' +
    '<tr><td>Celular / WhatsApp</td><td>' + (dados.whatsapp || '') + '</td></tr>' +
    '<tr><td>E-mail</td><td>' + (dados.email || '') + '</td></tr>' +
    '<tr><td>Sexo</td><td>' + (dados.sexo || '') + '</td></tr>' +
    '<tr><td>Cidade / UF</td><td>' + (dados.cidade || '') + ' / ' + (dados.estado || '') + '</td></tr>' +
    '<tr><td>CEP</td><td>' + (dados.cep || '') + '</td></tr>' +
    '</table>' +
      '<div style="background:#f0f7ff;border-left:4px solid #0B1F3A;padding:12px 16px;margin:0 0 20px;border-radius:0 4px 4px 0">' +
      '<strong style="color:#0B1F3A">Triagem rápida</strong><br>' +
      '<span style="margin-right:16px">&#128203; PHQ: <strong>' + tagHumor + '</strong></span>' +
      '<span>&#128165; GAD: <strong>' + tagAnsied + '</strong></span>' +
      '</div>' +
    '<h3>Respostas da Anamnese</h3>' +
    '<table>' + linhasAnamnese + '</table>' +
    '</div>' +
    '<div class="ftr">Sistema de Anamnese NEXUS CLIN — Não responda este e-mail.</div>' +
    '</div></body></html>';

  try {
    var anexosInternos = (pdfBlob && semExames) ? [pdfBlob] : [];
    GmailApp.sendEmail(
      EMAIL_PROFISSIONAL,
      '[NOVO PACIENTE] ' + dados.nome + ' ' + tagTriagem,  // triagem
      'Novo paciente cadastrado.',
      { name: 'NEXUS CLIN Sistema', htmlBody: htmlInterno, attachments: anexosInternos }
    );
    Logger.log('✓ E-mail Sosthenes enviado');
  } catch (e) {
    Logger.log('⚠️ Falha ao enviar e-mail Sosthenes (silenciosa): ' + e.message);
  }

  try {
    GmailApp.sendEmail(
      EMAIL_JULIA,
      '[NOVO PACIENTE] ' + dados.nome + ' ' + tagTriagem,  // triagem
      'Novo paciente cadastrado.',
      { name: 'NEXUS CLIN Sistema', htmlBody: htmlInterno, attachments: (pdfBlob && semExames) ? [pdfBlob] : [] }
    );
    Logger.log('✓ E-mail Julia enviado');
  } catch (e) {
    Logger.log('⚠️ Falha ao enviar e-mail Julia (silenciosa): ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 5 — doGet(e) — TESTE DE STATUS
// ═══════════════════════════════════════════════════════
function doGet(e) {
  var agora = Utilities.formatDate(new Date(), 'America/Recife', 'dd/MM/yyyy HH:mm');
  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#0B1F3A;display:flex;' +
    'align-items:center;justify-content:center;min-height:100vh;margin:0}' +
    '.card{background:#fff;border-radius:12px;padding:40px 48px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.3)}' +
    'h1{color:#0B1F3A;margin:0 0 8px;font-size:28px}' +
    '.gold{color:#C9A84C;font-weight:bold}' +
    'p{color:#555;margin:8px 0;font-size:14px}' +
    '.ok{font-size:48px;margin-bottom:16px}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="ok">✅</div>' +
    '<h1><span class="gold">NEXUS CLIN</span> — Webhook ativo v19.2</h1>' +
    '<p>Livramento-PB | ' + agora + '</p>' +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html);
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO TESTE — testarWebhook()
// ═══════════════════════════════════════════════════════
function testarWebhook() {
  var dadosTeste = {
    nome:                   'João Teste Silva',
    cpf:                    '529.982.247-25',
    data_nascimento:        '15/05/1990',
    whatsapp:               '83999999999',
    email:                  'nexusclinpb@gmail.com',
    sexo:                   'Masculino',
    cep:                    '58190000',
    endereco:               'Rua das Flores',
    numero_endereco:        '123',
    complemento:            '',
    bairro:                 'Centro',
    cidade:                 'Livramento',
    estado:                 'PB',
    profissao:              'Empresário',
    carga_horaria:          '50',
    estado_civil:           'Casado(a)',
    origem:                 'Instagram',
    objetivo:               'Emagrecer e reduzir gordura corporal | Ter mais energia e disposição no dia a dia',
    prazo:                  'Entre 3 e 6 meses',
    desafios:               'Falta de tempo, ansiedade, cansaço crônico',
    peso_atual:             '92',
    altura:                 '175',
    percepcao_corpo:        'Gordura concentrada na barriga | Retenção de líquido ou inchaço',
    variacao_peso:          'Aumentou',
    rotina_trabalho:        'Sedentária (fico sentado(a) a maior parte do tempo)',
    disposicao:             '2',
    estresse:               '4',
    saude_mental:           '2',
    saude_fisica:           '3',
    qualidade_vida:         '3',
    horas_sono:             '6',
    queixas_sono:           'Acordo cansado(a), mesmo tendo dormido | Tenho dificuldade para pegar no sono',
    sonolencia:             'Com frequência',
    pratica_exercicio:      'Não pratico',
    tipo_exercicio:         '',
    agua:                   'Entre 1 e 2 litros',
    alcool:                 'Raramente (menos de 1x por semana)',
    tabagismo:              'Nunca fumei',
    padrao_alimentar:       'Como de tudo, sem restrição',
    intolerancia:           'Nenhum desconforto',
    comportamento_alimentar:'Come mais quando estou estressado(a) ou ansioso(a) | Como grandes volumes de uma vez',
    sint_digestivos:        'Gases excessivos | Barriga muito inchada',
    sint_gerais:            'Cansaço, dor de cabeça frequente, pressão alta',
    sint_emocionais:        'Ansiedade, irritabilidade, baixa autoestima',
    diag_metab:             'Pré-diabetes (glicemia alterada)',
    diag_cardio:            '',
    diag_psiq:              'Ansiedade',
    medicamentos:           'Nenhum',
    suplementos:            'Vitamina D ocasional',
    phq_1:                  '2 - Mais da metade dos dias',
    gad_1:                  '3 - Quase todos os dias',
    tem_exames:             'Não tenho exames recentes',
    observacoes_finais:     'Quero perder 15kg e ter mais disposição para trabalhar',
  };

  Logger.log('🚀 === INICIANDO TESTE COMPLETO ===');
  Logger.log('👤 Paciente: ' + dadosTeste.nome);
  Logger.log('📧 Email destino: ' + dadosTeste.email);

  // ETAPA 0: Planilha
  Logger.log('--- ETAPA 0: Planilha ---');
  salvarNaPlanilha(dadosTeste);

  // ETAPA 1: Claude
  Logger.log('--- ETAPA 1: Claude ---');
  var resumo;
  try {
    resumo = gerarResumoClaude(dadosTeste);
    Logger.log('✓ Resumo gerado: ' + resumo.substring(0, 100) + '...');
  } catch(err) {
    resumo = 'Texto fallback de teste.';
    Logger.log('⚠️ Claude falhou: ' + err.message);
  }

  // ETAPA 2: PDF
  Logger.log('--- ETAPA 2: PDF ---');
  var pdfBlob = null;
  try {
    pdfBlob = gerarPDFRequisicao(dadosTeste);
    Logger.log('✓ PDF gerado: ' + (pdfBlob ? pdfBlob.getName() : 'null'));
  } catch(err) {
    Logger.log('⚠️ PDF falhou: ' + err.message);
  }

  // ETAPA 3: E-mails
  Logger.log('--- ETAPA 3: E-mails ---');
  try {
    enviarEmails(dadosTeste, resumo, pdfBlob);
    Logger.log('✓ E-mails enviados');
  } catch(err) {
    Logger.log('✗ ERRO e-mails: ' + err.message);
  }

  // ETAPA 0b: Feegow
  Logger.log('--- ETAPA 0b: Feegow ---');
  try {
    var feegowResult = cadastrarNoFeegow(dadosTeste);
    Logger.log('✓ ETAPA 0b resultado: ' + JSON.stringify(feegowResult));
  } catch(err) {
    Logger.log('⚠️ ETAPA 0b falhou: ' + err.message);
  }


  Logger.log('🏁 === TESTE CONCLUÍDO ===');
}

// ════════════════════════════════════════════════════════
// FUNÇÃO — cadastrarNoFeegow(dados)
// Endpoint oficial: POST /v1/api/pacientes/novo-paciente
// Header: x-access-token (não Bearer)
// Data: converte DD/MM/YYYY -> YYYY-MM-DD antes de enviar
// ════════════════════════════════════════════════════════
function cadastrarNoFeegow(dados) {
  var token = FEEGOW_TOKEN;
  var unidadeId = PropertiesService.getScriptProperties().getProperty('FEEGOW_UNIDADE_ID') || '';
  if (!token) {
    Logger.log('⚠️ FEEGOW_TOKEN não configurado — pulando cadastro');
    return 'TOKEN_NAO_CONFIGURADO';
  }

  var payload = {
    nome_completo: dados.nome,
    cpf:           (dados.cpf || '').replace(/\D/g, ''),
    data_nascimento: formatarDataFeegow(dados.data_nascimento),
    genero:        dados.sexo === 'Masculino' ? 'M' : 'F',
    celular1:      (dados.whatsapp || '').replace(/\D/g, ''),
    email1:        dados.email || '',
    cep:           (dados.cep || '').replace(/\D/g, ''),
    endereco:      dados.endereco || '',
    cidade:        dados.cidade   || '',
    estado:        dados.estado   || '',
    unidade_id:   unidadeId
  };

  var options = {
    method: 'POST',
    headers: {
      'x-access-token': token,
      'Content-Type':   'application/json'
    },
    payload:           JSON.stringify(payload),
    muteHttpExceptions: true
  };

  Logger.log('📤 Feegow payload: ' + JSON.stringify(payload));
  var response   = UrlFetchApp.fetch('https://api.feegow.com/v1/api/patient/create', options);
  var statusCode = response.getResponseCode();
  var body       = response.getContentText();

  Logger.log('🔁 Feegow [' + statusCode + ']: ' + body);

  try {
    var json = JSON.parse(body);
    if (json.success) {
      Logger.log('✓ ETAPA 0b: paciente_id Feegow: ' + json.content.paciente_id);
      return json.content.paciente_id;
    } else {
      Logger.log('⚠️ Feegow falhou: ' + JSON.stringify(json));
      return null;
    }
  } catch(e) {
    Logger.log('✗ Feegow parse error: ' + e.message + ' | body: ' + body);
    return null;
  }
}
// ════════════════════════════════════════════════════════
// HELPER — formatarDataFeegow(dataStr)
// Converte DD/MM/YYYY -> YYYY-MM-DD (formato ISO esperado pelo Feegow)
// ════════════════════════════════════════════════════════
function formatarDataFeegow(dataStr) {
  if (!dataStr) return '';
  var p = dataStr.split('/');
  if (p.length === 3) return p[2] + '-' + p[1] + '-' + p[0];
  return dataStr; // já está em YYYY-MM-DD ou outro formato
}
