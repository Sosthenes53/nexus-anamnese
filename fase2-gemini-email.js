// ═══════════════════════════════════════════════
//  NEXUS — Apps Script FASE 2
//  Responsabilidade: ler última linha → chamar
//  Gemini → gravar sugestão de exames → e-mail
//
//  Este script roda SEPARADO da Fase 1.
//  Configure um trigger onEdit ou time-based.
// ═══════════════════════════════════════════════

const SHEET_RESPOSTAS_F2 = 'Respostas';
const SHEET_EXAMES       = 'Sugestão de Exames';
const GPT_MODEL          = 'gpt-4o-mini';
const WPP_NEXUS          = '5583999086787';

// ── Lê a API key do PropertiesService (seguro) ──
// Para configurar: vá em Projeto → Propriedades do script → Adicione
// Chave: OPENAI_KEY  |  Valor: sua chave da OpenAI
function getOpenAIKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_KEY');
  if (!key) throw new Error('OPENAI_KEY não configurada em Propriedades do script');
  return key;
}

// ── Trigger principal ──
// Configure para rodar a cada 5 minutos OU onEdit na planilha
function processarNovosLeads() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(SHEET_RESPOSTAS_F2);
  if (!sheet) return;

  const dados  = sheet.getDataRange().getValues();
  const header = dados[0];
  const colStatus = header.indexOf('Status');
  const colNome   = header.indexOf('Nome');
  const colEmail  = header.indexOf('E-mail');

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (linha[colStatus] !== 'Novo') continue;

    try {
      const d = linhaParaObjeto(header, linha);
      const sugestao = chamarGemini(d);

      salvarExames(ss, d, sugestao);

      if (d['E-mail']) {
        enviarEmail(d, sugestao);
      }

      sheet.getRange(i + 1, colStatus + 1).setValue('Processado');
    } catch (err) {
      sheet.getRange(i + 1, colStatus + 1).setValue('Erro: ' + err.message.substring(0, 50));
    }

    Utilities.sleep(2000); // Respeita rate limit da API
  }
}

// ── Converte linha da planilha em objeto ──
function linhaParaObjeto(header, linha) {
  const obj = {};
  header.forEach((col, i) => { obj[col] = linha[i] || ''; });
  return obj;
}

// ── Chama OpenAI GPT ──
function chamarGemini(d) {
  const key  = getOpenAIKey();
  const url  = 'https://api.openai.com/v1/chat/completions';
  const sexo = (d['Sexo'] || '').toLowerCase();

  const prompt = `Você é um enfermeiro especialista em medicina integrativa e performance metabólica.

Com base na anamnese abaixo, sugira exames laboratoriais organizados por eixo clínico.
Para cada exame, dê uma justificativa objetiva baseada nos dados da anamnese.
Seja direto e técnico. Máximo 2 linhas de justificativa por exame.

FORMATO OBRIGATÓRIO:
EIXO [NOME DO EIXO]
- [Nome do exame]: [Justificativa clínica]

Eixos obrigatórios: METABÓLICO E GLICÊMICO, HORMONAL, LIPÍDICO E CARDIOVASCULAR, TIREOIDIANO, INFLAMATÓRIO, NUTRICIONAL E MICRONUTRIENTES, HEPÁTICO E RENAL, HEMATOLÓGICO${sexo.includes('femin') ? ', GINECOLÓGICO E REPRODUTIVO' : ''}.

Finalize com:
ALERTAS CLÍNICOS
[Máximo 3 pontos de atenção antes da consulta]

---
DADOS DO PACIENTE:
Nome: ${d['Nome']}
Sexo: ${d['Sexo']} | Idade: ${d['Idade']} anos
Peso: ${d['Peso (kg)']} kg | Altura: ${d['Altura (cm)']} cm | IMC: ${d['IMC']}
Objetivo: ${d['Objetivo(s)']}
Sono: ${d['Horas sono']}h — ${d['Queixas sono']}
Estresse: ${d['Estresse (0-10)']}/10
Exercício: ${d['Exercício']} — ${d['Tipo exercício']} ${d['Freq. treino (x/sem)']}x/sem
Padrão alimentar: ${d['Padrão alimentar']}
Intolerâncias: ${d['Intolerâncias']}
Sint. digestivos: ${d['Sint. digestivos']}
${sexo.includes('femin')
  ? `Ciclo: ${d['Ciclo']} | Libido: ${d['Libido F']} | TPM: ${d['TPM']}
Anticoncepcional: ${d['Anticoncepcional']} | TRH: ${d['TRH']}
Diag. ginecológico: ${d['Diag. ginecológico']}`
  : `Libido: ${d['Libido M']} | Ereções: ${d['Ereções']}
Sint. andropausa: ${d['Sint. andropausa']}
Testosterona: ${d['Testosterona/anabolizantes']}`}
Saúde mental: ${d['Saúde mental (0-10)']}/10 | PHQ-2: ${d['PHQ-2']} | GAD-2: ${d['GAD-2']}
Diag. metabólico: ${d['Diag. metabólico']}
Diag. cardiovascular: ${d['Diag. cardiovascular']}
Diag. autoimune: ${d['Diag. autoimune']}
Medicamentos: ${d['Medicamentos']}
Suplementos: ${d['Suplementos']}
Sint. gerais: ${d['Sint. gerais']}
Sint. neurológicos: ${d['Sint. neurológicos']}
Sint. dermatológicos: ${d['Sint. dermatológicos']}
Observações do paciente: ${d['Observações finais']}`;

  const res = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + key },
    payload: JSON.stringify({
      model: GPT_MODEL,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.2,
      max_tokens: 1500
    })
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('OpenAI HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 200));
  }

  const json = JSON.parse(res.getContentText());
  return json.choices[0].message.content;
}

// ── Salva na aba de exames ──
function salvarExames(ss, d, texto) {
  const sheet = ss.getSheetByName(SHEET_EXAMES) || ss.insertSheet(SHEET_EXAMES);

  if (sheet.getLastRow() === 0) {
    const h = sheet.getRange(1, 1, 1, 5);
    h.setValues([['Data/Hora','Nome','Sexo / Idade','IMC','Sugestão de Exames']]);
    h.setBackground('#0D1F3C').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 600);
  }

  sheet.appendRow([
    d['Data/Hora'],
    d['Nome'],
    `${d['Sexo']} · ${d['Idade']} anos`,
    (function(){ const p=parseFloat(d['Peso (kg)']),h=parseFloat(d['Altura (cm)']); return (p>0&&h>0)?(p/Math.pow(h/100,2)).toFixed(1)+' kg/m²':''; })(),
    texto
  ]);

  const r = sheet.getRange(sheet.getLastRow(), 1, 1, 5);
  r.setFontSize(10).setVerticalAlignment('top').setWrap(true);
  r.setBackground(sheet.getLastRow() % 2 === 0 ? '#F0F4F8' : '#FFFFFF');
}

// ── Converte texto Gemini em HTML de exames por eixo ──
function examesParaHtml(sugestao) {
  const linhas = sugestao.split('\n');
  let html = '';
  let dentroDeEixo = false;
  const coresEixo = ['#0B1E3D','#1a3358','#0d2b4e','#162d4a','#0f2640','#1c3560','#0a1e38','#19304f','#112338'];
  let eixoIdx = 0;

  for (const linha of linhas) {
    const trim = linha.trim();
    if (!trim) continue;

    if (trim.startsWith('EIXO ') || trim.startsWith('ALERTAS CLÍNICOS')) {
      if (dentroDeEixo) html += '</table></div>';
      const titulo = trim.startsWith('EIXO ') ? trim.replace('EIXO ', '') : '⚠ ALERTAS CLÍNICOS';
      const cor = trim.startsWith('ALERTAS') ? '#7B3F00' : (coresEixo[eixoIdx % coresEixo.length]);
      eixoIdx++;
      html += `<div style="margin:20px 0 0;border-radius:6px;overflow:hidden">
        <div style="background:${cor};padding:10px 16px;font-family:Georgia,serif;font-size:12px;font-weight:bold;color:#C9A860;letter-spacing:1.5px;text-transform:uppercase">${titulo}</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">`;
      dentroDeEixo = true;

    } else if (trim.startsWith('- ') && dentroDeEixo) {
      const partes = trim.replace(/^- /, '').split(':');
      const exame = partes[0].replace(/\*\*/g, '').trim();
      const justif = partes.slice(1).join(':').replace(/\*\*/g, '').trim();
      const bgRow = eixoIdx % 2 === 0 ? '#F8F6F0' : '#FFFFFF';
      html += `<tr style="border-bottom:1px solid #e8e4dc">
        <td style="width:35%;padding:10px 14px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#0B1E3D;vertical-align:top;background:${bgRow}">${exame}</td>
        <td style="padding:10px 14px;font-family:Arial,sans-serif;font-size:12px;color:#444;line-height:1.5;background:${bgRow}">${justif}</td>
      </tr>`;

    } else if (dentroDeEixo && !trim.startsWith('EIXO ')) {
      const bgRow = eixoIdx % 2 === 0 ? '#F8F6F0' : '#FFFFFF';
      html += `<tr><td colspan="2" style="padding:8px 14px;font-family:Arial,sans-serif;font-size:12px;color:#555;background:${bgRow};font-style:italic">${trim.replace(/\*\*/g, '')}</td></tr>`;
    }
  }

  if (dentroDeEixo) html += '</table></div>';
  return html;
}

// ── Envia e-mail HTML premium para o paciente ──
function enviarEmail(d, sugestao) {
  const nome     = d['Nome'] || 'Paciente';
  const primeiro = nome.split(' ')[0];
  const assunto  = `NEXUS CLIN — Sua análise personalizada está pronta`;
  const objetivo = (d['Objetivo(s)'] || '—').split('|')[0].trim();
  const wppLink  = `https://wa.me/${WPP_NEXUS}`;

  // Calcula IMC a partir de peso e altura (evita bug do campo IMC retornar Date do Sheets)
  const peso  = parseFloat(d['Peso (kg)']);
  const altCm = parseFloat(d['Altura (cm)']);
  const imcVal = (peso > 0 && altCm > 0) ? (peso / Math.pow(altCm / 100, 2)).toFixed(1) : '—';

  const examesHtml = examesParaHtml(sugestao);
  const logoBlob = DriveApp.getFileById('1EVrecK9y-1J23j3G6XGjVi9ay2bP2ib9').getBlob();

  const htmlBody = `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#EDEAE3;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#EDEAE3;padding:32px 16px">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%">

  <!-- HEADER -->
  <tr><td style="background:#0B1E3D;border-radius:8px 8px 0 0;padding:28px 40px;text-align:center">
    <img src="cid:logo" width="60" alt="NEXUS CLIN" style="display:block;margin:0 auto 12px">
    <div style="font-family:Georgia,serif;font-size:28px;font-weight:bold;color:#C9A860;letter-spacing:6px">NEXUS CLIN</div>
    <div style="width:48px;height:2px;background:#C9A860;margin:10px auto 0"></div>
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#8899BB;letter-spacing:2px;margin-top:10px;text-transform:uppercase">Centro de Performance Metabólica &amp; Longevidade</div>
  </td></tr>

  <!-- BODY -->
  <tr><td style="background:#FFFFFF;padding:40px 40px 32px">

    <!-- SAUDAÇÃO -->
    <p style="font-family:Georgia,serif;font-size:22px;color:#0B1E3D;margin:0 0 8px">Olá, ${primeiro}.</p>
    <p style="font-family:Arial,sans-serif;font-size:14px;color:#555;line-height:1.7;margin:0 0 28px">
      Suas informações chegaram com segurança ao Sistema NEXUS. Com base na sua anamnese,
      preparamos uma <strong style="color:#0B1E3D">sugestão de exames personalizada</strong> para potencializar
      sua consulta e antecipar os pontos que vamos trabalhar juntos.
    </p>

    <!-- DIVISOR -->
    <div style="border-top:1px solid #E0DAD0;margin:0 0 28px"></div>

    <!-- PERFIL RESUMIDO -->
    <div style="font-family:Georgia,serif;font-size:11px;color:#C9A860;letter-spacing:2px;text-transform:uppercase;margin-bottom:14px">Seu perfil resumido</div>
    <table width="100%" cellpadding="0" cellspacing="8" style="margin-bottom:32px">
      <tr>
        <td width="50%" style="padding:4px">
          <div style="background:#F4F1EB;border-left:3px solid #C9A860;border-radius:4px;padding:14px 16px">
            <div style="font-size:10px;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">IMC</div>
            <div style="font-size:18px;font-weight:bold;color:#0B1E3D">${imcVal} <span style="font-size:11px;font-weight:normal;color:#666">kg/m²</span></div>
          </div>
        </td>
        <td width="50%" style="padding:4px">
          <div style="background:#F4F1EB;border-left:3px solid #C9A860;border-radius:4px;padding:14px 16px">
            <div style="font-size:10px;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">Objetivo</div>
            <div style="font-size:13px;font-weight:bold;color:#0B1E3D;line-height:1.3">${objetivo}</div>
          </div>
        </td>
      </tr>
      <tr>
        <td width="50%" style="padding:4px">
          <div style="background:#F4F1EB;border-left:3px solid #0B1E3D;border-radius:4px;padding:14px 16px">
            <div style="font-size:10px;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">Sono</div>
            <div style="font-size:18px;font-weight:bold;color:#0B1E3D">${d['Horas sono'] || '—'}<span style="font-size:11px;font-weight:normal;color:#666">h / noite</span></div>
          </div>
        </td>
        <td width="50%" style="padding:4px">
          <div style="background:#F4F1EB;border-left:3px solid #0B1E3D;border-radius:4px;padding:14px 16px">
            <div style="font-size:10px;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">Estresse</div>
            <div style="font-size:18px;font-weight:bold;color:#0B1E3D">${d['Estresse (0-10)'] || '—'}<span style="font-size:11px;font-weight:normal;color:#666"> / 10</span></div>
          </div>
        </td>
      </tr>
    </table>

    <!-- EXAMES -->
    <div style="font-family:Georgia,serif;font-size:11px;color:#C9A860;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px">Exames sugeridos para sua consulta</div>
    <p style="font-family:Arial,sans-serif;font-size:12px;color:#888;margin:0 0 12px;line-height:1.5">
      Esta lista foi gerada com base no seu perfil clínico. Você pode realizá-los em qualquer laboratório de sua preferência antes da consulta.
    </p>
    ${examesHtml}

    <!-- PRÓXIMOS PASSOS -->
    <div style="background:#F4F1EB;border-radius:6px;padding:24px 28px;margin-top:32px">
      <div style="font-family:Georgia,serif;font-size:11px;color:#C9A860;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px">Próximos passos</div>
      <table cellpadding="0" cellspacing="0" width="100%">
        <tr>
          <td style="vertical-align:top;padding-bottom:12px">
            <span style="background:#0B1E3D;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">1</span>
            <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Realize os exames em qualquer laboratório de sua preferência.</span>
          </td>
        </tr>
        <tr>
          <td style="vertical-align:top;padding-bottom:12px">
            <span style="background:#0B1E3D;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">2</span>
            <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Traga os resultados na consulta <em>ou envie pelo WhatsApp antes</em> — chegamos com sua rota já traçada.</span>
          </td>
        </tr>
        <tr>
          <td style="vertical-align:top">
            <span style="background:#0B1E3D;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">3</span>
            <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Na consulta, definiremos juntos o seu <strong>Protocolo NEXUS</strong> personalizado.</span>
          </td>
        </tr>
      </table>
    </div>

    <!-- CTA -->
    <div style="text-align:center;margin-top:32px">
      <a href="${wppLink}" style="display:inline-block;background:#C9A860;color:#0B1E3D;font-family:Georgia,serif;font-size:14px;font-weight:bold;letter-spacing:1px;text-decoration:none;padding:14px 36px;border-radius:4px">
        Enviar exames ou Agendar consulta
      </a>
    </div>

  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#0B1E3D;border-radius:0 0 8px 8px;padding:24px 40px;text-align:center">
    <div style="font-family:Georgia,serif;font-size:14px;color:#C9A860;letter-spacing:2px;margin-bottom:8px">NEXUS CLIN</div>
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#7A8FAA;line-height:1.8">
      Enf. Sosthenes Santos · COREN-PB 568176<br>
      <span style="color:#4A5F7A;font-size:10px;margin-top:8px;display:block">
        Esta análise tem finalidade de triagem inicial e não substitui avaliação clínica presencial.
      </span>
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;

  GmailApp.sendEmail(d['E-mail'], assunto, '', { htmlBody: htmlBody, inlineImages: { logo: logoBlob } });
}

// ── Configura o trigger ──
// Rode esta função UMA VEZ para criar o trigger automático
function configurarTrigger() {
  // Remove triggers anteriores para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processarNovosLeads') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria trigger que roda a cada 5 minutos
  ScriptApp.newTrigger('processarNovosLeads')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger configurado: processarNovosLeads roda a cada 5 minutos.');
}

// ── Teste manual ──
function testeManual() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESPOSTAS_F2);
  if (!sheet) {
    Logger.log('Planilha "Respostas" não encontrada. Execute a Fase 1 primeiro.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Nenhuma resposta na planilha ainda. Envie uma anamnese primeiro.');
    return;
  }

  // Força status "Novo" na última linha para testar
  const header    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colStatus = header.indexOf('Status');
  sheet.getRange(lastRow, colStatus + 1).setValue('Novo');

  processarNovosLeads();
  Logger.log('Teste concluído. Verifique aba "Sugestão de Exames" e seu e-mail.');
}
