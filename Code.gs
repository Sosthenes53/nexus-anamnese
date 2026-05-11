// ═══════════════════════════════════════════════════════
// NEXUS CLIN — Automação Clínica
// Google Apps Script v1.1
// Conta de execução: nexusclinpb@gmail.com
// ═══════════════════════════════════════════════════════

// ── CONSTANTES ──────────────────────────────────────────
const FEEGOW_TOKEN = 'SEU_TOKEN_FEEGOW';
const ANTHROPIC_KEY = 'SUA_CHAVE_ANTHROPIC';
const EMAIL_CLINICA = 'nexusclinpb@gmail.com';
const EMAIL_PROFISSIONAL = 'sosthenes53@gmail.com';
const CLINICA = 'NEXUS CLIN';
const PROFISSIONAL = 'Sosthenes dos Santos Alves';
const COREN = 'COREN-PB 568176';
const CARGO = 'Enfermeiro de Prática Avançada';
const CIDADE = 'Livramento-PB';
const TELEFONE = '(83) 9 9858-5691';

// ═══════════════════════════════════════════════════════
// FUNÇÃO 1 — doPost(e) — WEBHOOK PRINCIPAL
// FIX PROBLEMA 1 (CORS): aceita JSON puro OU form-data c/ campo payload
// ═══════════════════════════════════════════════════════

function doPost(e) {
  var resp = {};

  try {
    var dados;

    // Tenta JSON direto primeiro
    try {
      dados = JSON.parse(e.postData.contents);
    } catch(jsonErr) {
      // Fallback: form data com campo "payload"
      var payload = e.parameter.payload ||
                    e.parameters.payload;
      if (payload) {
        dados = JSON.parse(
          Array.isArray(payload) ? payload[0] : payload
        );
      }
    }

    if (!dados || !dados.nome) {
      throw new Error('Payload vazio ou inválido: ' +
        JSON.stringify(e.postData) + ' | params: ' +
        JSON.stringify(e.parameter));
    }

    Logger.log('📥 Dados recebidos: ' + JSON.stringify(dados));

    Logger.log('⏳ Iniciando ETAPA 1: cadastrarNoFeegow');
    var pacienteId = cadastrarNoFeegow(dados);
    Logger.log('✓ ETAPA 1 concluída — paciente_id: ' + pacienteId);

    Logger.log('⏳ Iniciando ETAPA 2: gerarResumoClaude');
    var resumo = gerarResumoClaude(dados);
    Logger.log('✓ ETAPA 2 concluída — resumo gerado');

    Logger.log('⏳ Iniciando ETAPA 3: gerarPDFRequisicao');
    var pdfBlob = gerarPDFRequisicao(dados);
    Logger.log('✓ ETAPA 3 concluída — PDF gerado');

    Logger.log('⏳ Iniciando ETAPA 4: enviarEmails');
    enviarEmails(dados, resumo, pdfBlob);
    Logger.log('✓ ETAPA 4 concluída — e-mails enviados');

    resp = { success: true, paciente_id: pacienteId, emails_enviados: true };

  } catch (err) {
    var etapa = 'desconhecida';
    var msg = err.message || String(err);
    if (msg.indexOf('Feegow') !== -1 || msg.indexOf('paciente') !== -1) { etapa = 'cadastrarNoFeegow'; }
    else if (msg.indexOf('Claude') !== -1 || msg.indexOf('Anthropic') !== -1) { etapa = 'gerarResumoClaude'; }
    else if (msg.indexOf('PDF') !== -1 || msg.indexOf('Document') !== -1) { etapa = 'gerarPDFRequisicao'; }
    else if (msg.indexOf('mail') !== -1 || msg.indexOf('Email') !== -1) { etapa = 'enviarEmails'; }
    Logger.log('✗ ERRO na etapa "' + etapa + '": ' + msg);
    resp = { success: false, etapa: etapa, error: msg };
  }

  return ContentService
    .createTextOutput(JSON.stringify(resp))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 2 — cadastrarNoFeegow(dados)
// FIX PROBLEMA 3: log detalhado + validação de token
// ═══════════════════════════════════════════════════════

function cadastrarNoFeegow(dados) {
  // FIX: validação do token antes de enviar
  if (!FEEGOW_TOKEN || FEEGOW_TOKEN === 'SEU_TOKEN_FEEGOW') {
    Logger.log('⚠️ FEEGOW_TOKEN não configurado — pulando cadastro');
    return 'TOKEN_NAO_CONFIGURADO';
  }

  var url = 'https://api.feegow.com/v1/api/pacientes/novo-paciente';

  var payload = {
    nome: dados.nome,
    cpf: dados.cpf.replace(/\D/g, ''),
    nascimento: dados.data_nascimento,
    celular: dados.celular.replace(/\D/g, ''),
    email: dados.email,
    sexo: dados.sexo === 'Masculino' ? 'M' : 'F',
    cep: dados.cep.replace(/\D/g, ''),
    endereco: dados.endereco,
    cidade: dados.cidade,
    estado: dados.estado
  };

  var options = {
    method: 'POST',
    headers: {
      'x-access-token': FEEGOW_TOKEN,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 15000
  };

  // FIX: logs detalhados para diagnóstico
  Logger.log('📤 URL: ' + url);
  Logger.log('📤 Token (primeiros 8): ' + FEEGOW_TOKEN.substring(0, 8) + '...');
  Logger.log('📤 Payload: ' + JSON.stringify(payload));

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();
  var bodyText = response.getContentText();
  var content = {};

  Logger.log('📥 Status: ' + statusCode);
  Logger.log('📥 Response: ' + bodyText);

  try { content = JSON.parse(bodyText); } catch (e) { content = {}; }

  if (statusCode === 200 || statusCode === 201) {
    if (content.paciente_id) return content.paciente_id;
    if (content.data && content.data.paciente_id) return content.data.paciente_id;
    if (content.id) return content.id;
  }

  var isDuplicate = (
    statusCode === 409 ||
    bodyText.toLowerCase().indexOf('duplicado') !== -1 ||
    bodyText.toLowerCase().indexOf('já cadastrado') !== -1 ||
    bodyText.toLowerCase().indexOf('ja cadastrado') !== -1 ||
    bodyText.toLowerCase().indexOf('cpf') !== -1
  );

  if (isDuplicate) {
    Logger.log('⚠️ CPF já cadastrado no Feegow — extraindo paciente_id existente');
    var existingId =
      (content.paciente_id) ||
      (content.data && content.data.paciente_id) ||
      (content.id) ||
      (content.data && content.data.id) ||
      'DUPLICADO';
    Logger.log('↩️ paciente_id existente: ' + existingId);
    return existingId;
  }

  throw new Error(
    'Feegow API retornou status ' + statusCode + ': ' +
    (content.message || content.error || bodyText.substring(0, 200))
  );
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 3 — gerarResumoClaude(dados)
// ═══════════════════════════════════════════════════════

function gerarResumoClaude(dados) {
  var FALLBACK =
    'Olá! Estamos muito felizes em receber você na NEXUS CLIN. ' +
    'Seu protocolo personalizado será construído com cuidado e dedicação para você.';

  try {
    var primeiroNome = dados.nome.split(' ')[0];

    var promptUsuario =
      'Paciente: ' + primeiroNome + '\n' +
      'Objetivo principal: ' + (dados.objetivo || 'Não informado') + '\n' +
      'Nível de energia: ' + (dados.nivel_energia || '?') + '/10\n' +
      'Qualidade do sono: ' + (dados.qualidade_sono || '?') + '/10\n' +
      'Nível de estresse: ' + (dados.nivel_estresse || 'Não informado') + '\n' +
      'Sintomas gerais: ' + (dados.sint_gerais || 'Nenhum relatado') + '\n' +
      'Sintomas emocionais: ' + (dados.sint_emocionais || 'Nenhum relatado') + '\n\n' +
      'Escreva um parágrafo acolhedor e empático (máximo 4 frases), iniciando pelo ' +
      'primeiro nome do paciente (' + primeiroNome + '), que apresente brevemente o ' +
      'panorama clínico e termine com uma frase de encorajamento sobre o início do ' +
      'protocolo na NEXUS CLIN.';

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
    Logger.log('📥 Claude RESPONSE [' + statusCode + ']: ' + bodyText.substring(0, 300));

    if (statusCode === 200) {
      var content = JSON.parse(bodyText);
      if (content.content && content.content[0] && content.content[0].text) {
        return content.content[0].text;
      }
    }

    Logger.log('⚠️ Claude retornou status ' + statusCode + ' — usando fallback');
    return FALLBACK;

  } catch (err) {
    Logger.log('⚠️ Erro em gerarResumoClaude: ' + err.message + ' — usando fallback');
    return FALLBACK;
  }
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 4 — gerarPDFRequisicao(dados)
// ═══════════════════════════════════════════════════════

function gerarPDFRequisicao(dados) {
  var primeiroNome = dados.nome.split(' ')[0];
  var docId = null;
  var pdfBlob = null;

  var exames = [];
  if (dados.sexo === 'Masculino') {
    exames = [
      'Glicemia de Jejum', 'Insulina de Jejum', 'HbA1c', 'TSH + T4 Livre',
      'Hemograma Completo', 'Ferritina', 'Vitamina D (25-OH)', 'Cortisol Matinal',
      'Perfil Lipídico Completo', 'TGO + TGP', 'Testosterona Total + SHBG'
    ];
  } else {
    exames = [
      'Glicemia de Jejum', 'Insulina de Jejum', 'HbA1c', 'TSH + T4 Livre',
      'Hemograma Completo', 'Ferritina', 'Vitamina D (25-OH)', 'Cortisol Matinal',
      'Perfil Lipídico Completo', 'TGO + TGP',
      'Testosterona Total + SHBG + DHEA-S + LH + FSH'
    ];
  }

  var hoje = Utilities.formatDate(new Date(), 'America/Recife', 'dd/MM/yyyy');
  var doc = DocumentApp.create('nexus_req_temp_' + Date.now());
  docId = doc.getId();
  var body = doc.getBody();

  body.setMarginTop(56);
  body.setMarginBottom(56);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // CABEÇALHO
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

  // DADOS DO PACIENTE
  var pn = body.appendParagraph('Paciente: ' + dados.nome);
  pn.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pn.getChild(0).asText().setFontSize(11);

  var pd = body.appendParagraph('Data de Emissão: ' + hoje);
  pd.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pd.getChild(0).asText().setFontSize(11);

  body.appendHorizontalRule();

  // TÍTULO DA REQUISIÇÃO
  var tr = body.appendParagraph('Painel Metabólico NEXUS — Bloqueios ao Emagrecimento');
  tr.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  tr.getChild(0).asText().setFontSize(13).setBold(true).setForegroundColor('#0B1F3A');

  body.appendParagraph('');

  // LISTA DE EXAMES
  for (var i = 0; i < exames.length; i++) {
    var item = body.appendParagraph((i + 1) + '. ' + exames[i]);
    item.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    item.getChild(0).asText().setFontSize(11);
  }

  body.appendParagraph('');

  // JUSTIFICATIVA CLÍNICA
  var just =
    'Solicitação de mapeamento metabólico direcionado à identificação de resistência ' +
    'insulínica, disfunção tireoidiana, hipovitaminoses, sobrecarga hepática e ' +
    'desequilíbrio hormonal — condições de alta prevalência nessa população e com ' +
    'impacto direto na resposta ao emagrecimento. Os achados nortearão a construção ' +
    'de um protocolo individualizado, podendo envolver condutas hormonais, nutricionais, ' +
    'farmacológicas ou ortomoleculares, com objetivo de personalizar a intervenção, ' +
    'antecipar riscos e garantir segurança no uso de análogos de GLP-1 e demais ' +
    'estratégias metabólicas.';

  var pj = body.appendParagraph(just);
  pj.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pj.getChild(0).asText().setFontSize(10).setItalic(true);

  var pc = body.appendParagraph('CID: E66.9');
  pc.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  pc.getChild(0).asText().setFontSize(10).setItalic(true);

  body.appendHorizontalRule();

  // ASSINATURA
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
        Logger.log('🗑️ Documento temporário removido: ' + docId);
      } catch (delErr) {
        Logger.log('⚠️ Não foi possível remover doc temporário: ' + delErr.message);
      }
    }
  }

  return pdfBlob;
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 5 — enviarEmails(dados, resumo, pdfBlob)
// FIX PROBLEMA 2: e-mail interno inclui tabela com TODAS as respostas
// ═══════════════════════════════════════════════════════

function enviarEmails(dados, resumo, pdfBlob) {
  var primeiroNome = dados.nome.split(' ')[0];

  // ── E-MAIL 1: PACIENTE ───────────────────────────────
  var htmlPaciente =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
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
    '<p class="ola">Ol&aacute;, ' + primeiroNome + '! 👋</p>' +
    '<p style="color:#444;font-size:15px;line-height:1.7">É com muito prazer que recebemos você na ' +
    '<strong>NEXUS CLIN</strong>. Nossa missão é oferecer um cuidado personalizado, baseado em evidências, ' +
    'focado no seu equilíbrio metabólico e qualidade de vida a longo prazo.</p>' +
    '<div class="resumo"><h3>Seu panorama inicial</h3><p>' + resumo + '</p></div>' +
    '<div class="anexo">📎 <strong>Sua solicitação de exames está em anexo.</strong><br>' +
    'Realize a coleta o quanto antes para que possamos iniciar seu protocolo personalizado com precisão.</div>' +
    '<p style="color:#555;font-size:14px;line-height:1.7">Em caso de dúvidas, entre em contato pelo WhatsApp ' +
    'ou aguarde nosso retorno após análise dos seus exames. Estamos com você em cada etapa dessa jornada. 💚</p>' +
    '<div class="ass"><strong>Sosthenes dos Santos Alves</strong><br>' +
    'COREN-PB 568176 | Enfermeiro de Prática Avançada<br>' +
    'NEXUS CLIN — Livramento-PB<br>📱 (83) 9 9858-5691<br>✉️ nexusclinpb@gmail.com</div>' +
    '</div>' +
    '<div class="ftr"><span>NEXUS CLIN</span> — Livramento-PB<br>' +
    'Este e-mail contém informações confidenciais de saúde. Caso não seja o destinatário, por favor desconsidere.</div>' +
    '</div></body></html>';

  try {
    GmailApp.sendEmail(
      dados.email,
      'Bem-vindo(a) à NEXUS CLIN — sua jornada começa agora 🌿',
      'Olá ' + primeiroNome + '! Sua anamnese foi recebida. Visualize em HTML para melhor experiência.',
      { name: 'NEXUS CLIN', htmlBody: htmlPaciente, attachments: [pdfBlob] }
    );
    Logger.log('✓ E-mail paciente enviado para ' + dados.email);
  } catch (e) {
    Logger.log('✗ ERRO e-mail paciente: ' + e.message);
    throw e;
  }

  // ── E-MAIL 2: PROFISSIONAL (alerta interno) ──────────
  var dataHora = Utilities.formatDate(new Date(), 'America/Recife', "dd/MM/yyyy 'às' HH:mm");

  var nascFormatado = dados.data_nascimento || '';
  if (nascFormatado.indexOf('-') !== -1) {
    var partes = nascFormatado.split('-');
    if (partes.length === 3) nascFormatado = partes[2] + '/' + partes[1] + '/' + partes[0];
  }

  // FIX PROBLEMA 2: gera linhas com TODAS as respostas (exceto campos da tabela principal)
  var linhasRespostas = '';
  var camposIgnorar = ['timestamp', 'sexo', 'nome',
    'cpf', 'data_nascimento', 'celular', 'email',
    'cep', 'endereco', 'cidade', 'estado'];
  Object.keys(dados).forEach(function(campo) {
    if (camposIgnorar.indexOf(campo) === -1) {
      var valor = dados[campo];
      if (valor && String(valor).trim() !== '') {
        linhasRespostas += '<tr>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:bold;color:#555;width:40%">' +
          campo.replace(/_/g, ' ') + '</td>' +
          '<td style="padding:8px 12px;border-bottom:1px solid #eee;color:#333">' + valor + '</td>' +
          '</tr>';
      }
    }
  });

  var htmlInterno =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#f0f0f0;margin:0}' +
    '.wrap{max-width:600px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden}' +
    '.hdr{background:#0B1F3A;padding:20px 24px}' +
    '.hdr h2{color:#C9A84C;margin:0;font-size:16px}' +
    '.hdr p{color:#aaa;margin:4px 0 0;font-size:12px}' +
    '.bod{padding:24px}' +
    'table{width:100%;border-collapse:collapse;font-size:14px}' +
    'tr:nth-child(even){background:#fafafa}' +
    'td{padding:10px 12px;border-bottom:1px solid #eee;color:#333}' +
    'td:first-child{font-weight:bold;color:#0B1F3A;width:38%}' +
    '.btn{display:block;width:fit-content;margin:24px auto 0;background:#0B1F3A;color:#fff;' +
    'padding:12px 28px;border-radius:6px;text-decoration:none;font-size:14px;font-weight:bold}' +
    '.ftr{background:#f5f5f5;padding:12px 24px;text-align:center;color:#999;font-size:11px;margin-top:24px}' +
    '</style></head>' +
    '<body><div class="wrap">' +
    '<div class="hdr"><h2>🔔 Novo paciente aguardando atendimento</h2>' +
    '<p>Recebido em ' + dataHora + '</p></div>' +
    '<div class="bod"><table>' +
    '<tr><td>Nome completo</td><td>' + dados.nome + '</td></tr>' +
    '<tr><td>Data de nascimento</td><td>' + nascFormatado + '</td></tr>' +
    '<tr><td>Celular</td><td>' + dados.celular + '</td></tr>' +
    '<tr><td>E-mail</td><td>' + dados.email + '</td></tr>' +
    '<tr><td>Sexo</td><td>' + dados.sexo + '</td></tr>' +
    '<tr><td>Cidade / UF</td><td>' + dados.cidade + ' / ' + dados.estado + '</td></tr>' +
    '<tr><td>CEP</td><td>' + dados.cep + '</td></tr>' +
    '<tr><td>Objetivo principal</td><td>' + (dados.objetivo || 'Não informado') + '</td></tr>' +
    '<tr><td>Sintomas relatados</td><td>' + (dados.sint_gerais || 'Não informado') + '</td></tr>' +
    '</table>' +
    '<h3 style="color:#0B1F3A;margin:24px 0 12px;font-size:14px">Respostas completas da anamnese</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
    linhasRespostas +
    '</table>' +
    '<a class="btn" href="https://app.feegow.com/main/?P=pacientes">Abrir paciente no Feegow →</a>' +
    '</div>' +
    '<div class="ftr">Mensagem automática — Sistema de Anamnese NEXUS CLIN<br>Não responda este e-mail.</div>' +
    '</div></body></html>';

  try {
    GmailApp.sendEmail(
      EMAIL_PROFISSIONAL,
      '[NOVO PACIENTE] ' + dados.nome,
      'Novo paciente cadastrado. Veja o e-mail HTML para detalhes completos.',
      { name: 'NEXUS CLIN Sistema', htmlBody: htmlInterno }
    );
    Logger.log('✓ Alerta interno enviado para ' + EMAIL_PROFISSIONAL);
  } catch (e) {
    Logger.log('✗ ERRO alerta interno: ' + e.message);
    // não propaga — falha silenciosa
  }
}

// ═══════════════════════════════════════════════════════
// FUNÇÃO 6 — doGet(e) — TESTE DE STATUS
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
    '<h1><span class="gold">NEXUS CLIN</span> — Webhook ativo</h1>' +
    '<p>Versão 1.1 | Livramento-PB</p>' +
    '<p>Data: ' + agora + '</p>' +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html);
}
