// ═══════════════════════════════════════════════
// NEXUS — Apps Script FASE 2
// Responsabilidade: ler última linha → chamar
// Claude API → gerar resumo clínico → PDF → e-mail
//
// Este script roda SEPARADO da Fase 1.
// Configure um trigger onEdit ou time-based.
// ═══════════════════════════════════════════════

const SHEET_RESPOSTAS_F2 = 'Respostas';
const SHEET_EXAMES       = 'Sugestão de Exames';
const WPP_NEXUS          = '5583999086787';

// ────────────────────────────────────────────────────────────────
// FUNÇÃO 5 — enviarEmail(dados, resumo, pdfBlob)
// Envia dois e-mails independentes:
//   1. Paciente → boas-vindas com resumo clínico + PDF anexo
//   2. Profissional → alerta interno com dados do paciente
// ────────────────────────────────────────────────────────────────
function enviarEmail(dados, resumo, pdfBlob) {

  const primeiroNome = (dados.nome || 'Paciente').split(' ')[0];

  // ── FORMATAÇÕES AUXILIARES ─────────────────────────────────────
  // Data/hora atual para o alerta interno
  const agora     = new Date();
    const dataHora  = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                          ' às ' +
                          Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm');

  // Data de nascimento formatada
  let dataNascFormatada = dados.data_nascimento || '';
    if (dataNascFormatada instanceof Date) {
          dataNascFormatada = Utilities.formatDate(dataNascFormatada, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } else if (typeof dataNascFormatada === 'string' && dataNascFormatada.match(/^\d{4}-\d{2}-\d{2}/)) {
          const parts = dataNascFormatada.substring(0, 10).split('-');
          dataNascFormatada = parts[2] + '/' + parts[1] + '/' + parts[0];
    }

  // ══════════════════════════════════════════════════════════════
  // E-MAIL 1 — PACIENTE (boas-vindas + resumo clínico + PDF)
  // ══════════════════════════════════════════════════════════════
  try {

      const htmlPaciente = `<!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          </head>
          <body style="margin:0;padding:0;background:#EDEAE3;font-family:Arial,sans-serif">
          <table width="100%" cellpadding="0" cellspacing="0" style="background:#EDEAE3;padding:32px 16px">
          <tr><td align="center">
          <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%">

            <!-- HEADER -->
              <tr><td style="background:#0B1F3A;border-radius:8px 8px 0 0;padding:32px 40px;text-align:center">
                  <div style="font-family:Georgia,serif;font-size:30px;font-weight:bold;color:#C9A860;letter-spacing:6px;margin-bottom:6px">
                        NEXUS CLIN
                            </div>
                                <div style="width:48px;height:2px;background:#C9A860;margin:0 auto 12px"></div>
                                    <div style="font-family:Arial,sans-serif;font-size:11px;color:#8899BB;letter-spacing:2px;text-transform:uppercase">
                                          Centro de Performance Metabólica &amp; Longevidade
                                              </div>
                                                </td></tr>

                                                  <!-- CORPO -->
                                                    <tr><td style="background:#FFFFFF;padding:40px 40px 32px">

                                                        <!-- SAUDAÇÃO -->
                                                            <p style="font-family:Georgia,serif;font-size:22px;color:#0B1F3A;margin:0 0 10px">
                                                                  Bem-vindo(a), ${primeiroNome}. 🌿
                                                                      </p>
                                                                          <p style="font-family:Arial,sans-serif;font-size:14px;color:#555;line-height:1.7;margin:0 0 28px">
                                                                                Sua anamnese foi recebida com sucesso pelo Sistema NEXUS.<br>
                                                                                      Preparamos abaixo um <strong style="color:#0B1F3A">resumo clínico personalizado</strong>
                                                                                            com base nas informações que você nos enviou.
                                                                                                </p>

                                                                                                    <!-- DIVISOR -->
                                                                                                        <div style="border-top:1px solid #E0DAD0;margin:0 0 28px"></div>
                                                                                                        
                                                                                                            <!-- CAIXA DOURADA — RESUMO CLÍNICO -->
                                                                                                                <div style="background:#FFFDF5;border:1px solid #C9A860;border-left:4px solid #C9A860;border-radius:6px;padding:24px 28px;margin-bottom:28px">
                                                                                                                      <div style="font-family:Georgia,serif;font-size:11px;color:#C9A860;letter-spacing:2px;text-transform:uppercase;margin-bottom:14px">
                                                                                                                              ✦ Resumo Clínico Personalizado
                                                                                                                                    </div>
                                                                                                                                          <div style="font-family:Arial,sans-serif;font-size:13px;color:#333;line-height:1.8;white-space:pre-line">
                                                                                                                                          ${resumo || 'Resumo em processamento. Em breve você receberá mais detalhes.'}
                                                                                                                                                </div>
                                                                                                                                                    </div>
                                                                                                                                                    
                                                                                                                                                        <!-- AVISO ANEXO (verde) -->
                                                                                                                                                            <div style="background:#F0FFF4;border:1px solid #68D391;border-left:4px solid #38A169;border-radius:6px;padding:16px 20px;margin-bottom:32px;display:flex;align-items:center">
                                                                                                                                                                  <span style="font-size:20px;margin-right:12px">📎</span>
                                                                                                                                                                        <div>
                                                                                                                                                                                <div style="font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#276749">
                                                                                                                                                                                          Sua solicitação de exames está em anexo
                                                                                                                                                                                                  </div>
                                                                                                                                                                                                          <div style="font-family:Arial,sans-serif;font-size:12px;color:#4A7C59;margin-top:2px">
                                                                                                                                                                                                                    Arquivo: <em>Solicitacao_Exames_${primeiroNome}.pdf</em> — leve na consulta ou realize antes dela.
                                                                                                                                                                                                                            </div>
                                                                                                                                                                                                                                  </div>
                                                                                                                                                                                                                                      </div>
                                                                                                                                                                                                                                      
                                                                                                                                                                                                                                          <!-- PRÓXIMOS PASSOS -->
                                                                                                                                                                                                                                              <div style="background:#F4F1EB;border-radius:6px;padding:24px 28px;margin-bottom:32px">
                                                                                                                                                                                                                                                    <div style="font-family:Georgia,serif;font-size:11px;color:#C9A860;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px">
                                                                                                                                                                                                                                                            Próximos passos
                                                                                                                                                                                                                                                                  </div>
                                                                                                                                                                                                                                                                        <table cellpadding="0" cellspacing="0" width="100%">
                                                                                                                                                                                                                                                                                <tr><td style="vertical-align:top;padding-bottom:12px">
                                                                                                                                                                                                                                                                                          <span style="background:#0B1F3A;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">1</span>
                                                                                                                                                                                                                                                                                                    <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Realize os exames do PDF anexo em qualquer laboratório de sua preferência.</span>
                                                                                                                                                                                                                                                                                                            </td></tr>
                                                                                                                                                                                                                                                                                                                    <tr><td style="vertical-align:top;padding-bottom:12px">
                                                                                                                                                                                                                                                                                                                              <span style="background:#0B1F3A;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">2</span>
                                                                                                                                                                                                                                                                                                                                        <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Traga os resultados na consulta ou envie pelo WhatsApp — chegamos com sua rota já traçada.</span>
                                                                                                                                                                                                                                                                                                                                                </td></tr>
                                                                                                                                                                                                                                                                                                                                                        <tr><td style="vertical-align:top">
                                                                                                                                                                                                                                                                                                                                                                  <span style="background:#0B1F3A;color:#C9A860;font-size:11px;font-weight:bold;border-radius:50%;width:22px;height:22px;display:inline-block;text-align:center;line-height:22px">3</span>
                                                                                                                                                                                                                                                                                                                                                                            <span style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-left:10px">Na consulta, definiremos juntos o seu <strong>Protocolo NEXUS</strong> personalizado.</span>
                                                                                                                                                                                                                                                                                                                                                                                    </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                          </table>
                                                                                                                                                                                                                                                                                                                                                                                              </div>
                                                                                                                                                                                                                                                                                                                                                                                              
                                                                                                                                                                                                                                                                                                                                                                                                  <!-- ASSINATURA -->
                                                                                                                                                                                                                                                                                                                                                                                                      <div style="border-top:1px solid #E0DAD0;padding-top:24px;text-align:center">
                                                                                                                                                                                                                                                                                                                                                                                                            <div style="font-family:Georgia,serif;font-size:14px;color:#0B1F3A;font-weight:bold;margin-bottom:4px">
                                                                                                                                                                                                                                                                                                                                                                                                                    Sosthenes dos Santos Alves
                                                                                                                                                                                                                                                                                                                                                                                                                          </div>
                                                                                                                                                                                                                                                                                                                                                                                                                                <div style="font-family:Arial,sans-serif;font-size:12px;color:#777;margin-bottom:2px">
                                                                                                                                                                                                                                                                                                                                                                                                                                        Enfermeiro · COREN-PB 568176
                                                                                                                                                                                                                                                                                                                                                                                                                                              </div>
                                                                                                                                                                                                                                                                                                                                                                                                                                                    <div style="font-family:Arial,sans-serif;font-size:12px;color:#777">
                                                                                                                                                                                                                                                                                                                                                                                                                                                            NEXUS CLIN — Performance Metabólica &amp; Longevidade
                                                                                                                                                                                                                                                                                                                                                                                                                                                                  </div>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                      </div>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                                                                                                                                                                                                                                          <!-- FOOTER -->
                                                                                                                                                                                                                                                                                                                                                                                                                                                                            <tr><td style="background:#0B1F3A;border-radius:0 0 8px 8px;padding:20px 40px;text-align:center">
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                <div style="font-family:Arial,sans-serif;font-size:10px;color:#4A5F7A;line-height:1.7">
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      Esta mensagem contém informações clínicas confidenciais destinadas exclusivamente ao seu destinatário.<br>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            Caso tenha recebido por engano, por favor descarte esta mensagem e notifique-nos.<br>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  <span style="color:#6B7FA0">© 2026 NEXUS CLIN — Todos os direitos reservados</span>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      </div>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </table>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </table>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </body>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        </html>`;

      GmailApp.sendEmail(
              dados.email,
              'Bem-vindo(a) à NEXUS CLIN — sua jornada começa agora 🌿',
              'Olá! Sua anamnese foi recebida. Veja o e-mail em formato HTML.',
        {
                  from:        'nexusclinpb@gmail.com',
                  name:        'NEXUS CLIN',
                  htmlBody:    htmlPaciente,
                  attachments: [pdfBlob.setName('Solicitacao_Exames_' + primeiroNome + '.pdf')]
        }
            );
        Logger.log('E-mail paciente enviado: ' + dados.email);

  } catch (errPaciente) {
        Logger.log('ERRO email paciente: ' + errPaciente);
        throw errPaciente; // propaga para doPost registrar
  }

  // ══════════════════════════════════════════════════════════════
  // E-MAIL 2 — PROFISSIONAL (alerta interno, falha silenciosa)
  // ══════════════════════════════════════════════════════════════
  try {

      const htmlInterno = `<!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          </head>
          <body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif">
          <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:24px 16px">
          <tr><td align="center">
          <table width="580" cellpadding="0" cellspacing="0" style="max-width:580px;width:100%;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08)">

            <!-- CABEÇALHO -->
              <tr><td style="padding:28px 32px 20px">
                  <div style="font-family:Arial,sans-serif;font-size:20px;font-weight:bold;color:#222;margin-bottom:6px">
                        🔔 Novo paciente aguardando atendimento
                            </div>
                                <div style="font-family:Arial,sans-serif;font-size:13px;color:#888">
                                      Recebido em ${dataHora}
                                          </div>
                                            </td></tr>

                                              <!-- DIVISOR -->
                                                <tr><td style="padding:0 32px">
                                                    <div style="border-top:1px solid #E8E8E8"></div>
                                                      </td></tr>

                                                        <!-- TABELA DE DADOS -->
                                                          <tr><td style="padding:20px 32px 24px">
                                                              <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #E8E8E8;border-radius:6px;overflow:hidden;font-family:Arial,sans-serif;font-size:13px">

                                                                    <tr style="background:#fafafa">
                                                                            <td style="padding:10px 14px;width:38%;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Nome completo</td>
                                                                                    <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.nome || '—'}</td>
                                                                                          </tr>
                                                                                                <tr style="background:#ffffff">
                                                                                                        <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Data de nascimento</td>
                                                                                                                <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dataNascFormatada || '—'}</td>
                                                                                                                      </tr>
                                                                                                                            <tr style="background:#fafafa">
                                                                                                                                    <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Celular</td>
                                                                                                                                            <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.celular || '—'}</td>
                                                                                                                                                  </tr>
                                                                                                                                                        <tr style="background:#ffffff">
                                                                                                                                                                <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">E-mail</td>
                                                                                                                                                                        <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.email || '—'}</td>
                                                                                                                                                                              </tr>
                                                                                                                                                                                    <tr style="background:#fafafa">
                                                                                                                                                                                            <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Sexo</td>
                                                                                                                                                                                                    <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.sexo || '—'}</td>
                                                                                                                                                                                                          </tr>
                                                                                                                                                                                                                <tr style="background:#ffffff">
                                                                                                                                                                                                                        <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Cidade / UF</td>
                                                                                                                                                                                                                                <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${(dados.cidade || '—') + ' / ' + (dados.estado || '—')}</td>
                                                                                                                                                                                                                                      </tr>
                                                                                                                                                                                                                                            <tr style="background:#fafafa">
                                                                                                                                                                                                                                                    <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">CEP</td>
                                                                                                                                                                                                                                                            <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.cep || '—'}</td>
                                                                                                                                                                                                                                                                  </tr>
                                                                                                                                                                                                                                                                        <tr style="background:#ffffff">
                                                                                                                                                                                                                                                                                <td style="padding:10px 14px;color:#666;font-weight:bold;border-bottom:1px solid #EFEFEF;border-right:1px solid #EFEFEF">Objetivo principal</td>
                                                                                                                                                                                                                                                                                        <td style="padding:10px 14px;color:#222;border-bottom:1px solid #EFEFEF">${dados.objetivo || '—'}</td>
                                                                                                                                                                                                                                                                                              </tr>
                                                                                                                                                                                                                                                                                                    <tr style="background:#fafafa">
                                                                                                                                                                                                                                                                                                            <td style="padding:10px 14px;color:#666;font-weight:bold;border-right:1px solid #EFEFEF">Sintomas relatados</td>
                                                                                                                                                                                                                                                                                                                    <td style="padding:10px 14px;color:#222">${dados.sint_gerais || '—'}</td>
                                                                                                                                                                                                                                                                                                                          </tr>
                                                                                                                                                                                                                                                                                                                          
                                                                                                                                                                                                                                                                                                                              </table>
                                                                                                                                                                                                                                                                                                                                </td></tr>
                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                  <!-- BOTÃO FEEGOW -->
                                                                                                                                                                                                                                                                                                                                    <tr><td style="padding:0 32px 32px;text-align:center">
                                                                                                                                                                                                                                                                                                                                        <a href="https://app.feegow.com/main/?P=pacientes"
                                                                                                                                                                                                                                                                                                                                               style="display:inline-block;background:#0B1F3A;color:#ffffff;font-family:Arial,sans-serif;font-size:14px;font-weight:bold;text-decoration:none;padding:14px 32px;border-radius:5px;letter-spacing:0.5px">
                                                                                                                                                                                                                                                                                                                                                     Abrir paciente no Feegow →
                                                                                                                                                                                                                                                                                                                                                         </a>
                                                                                                                                                                                                                                                                                                                                                           </td></tr>
                                                                                                                                                                                                                                                                                                                                                           
                                                                                                                                                                                                                                                                                                                                                             <!-- RODAPÉ -->
                                                                                                                                                                                                                                                                                                                                                               <tr><td style="background:#f9f9f9;border-top:1px solid #EFEFEF;padding:14px 32px;text-align:center">
                                                                                                                                                                                                                                                                                                                                                                   <div style="font-family:Arial,sans-serif;font-size:11px;color:#AAAAAA;line-height:1.6">
                                                                                                                                                                                                                                                                                                                                                                         Mensagem automática gerada pelo sistema de anamnese NEXUS CLIN<br>
                                                                                                                                                                                                                                                                                                                                                                               Não responda este e-mail.
                                                                                                                                                                                                                                                                                                                                                                                   </div>
                                                                                                                                                                                                                                                                                                                                                                                     </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                     
                                                                                                                                                                                                                                                                                                                                                                                     </table>
                                                                                                                                                                                                                                                                                                                                                                                     </td></tr>
                                                                                                                                                                                                                                                                                                                                                                                     </table>
                                                                                                                                                                                                                                                                                                                                                                                     </body>
                                                                                                                                                                                                                                                                                                                                                                                     </html>`;

      GmailApp.sendEmail(
              'sosthenes53@gmail.com',
              '[NOVO PACIENTE] ' + (dados.nome || 'Sem nome'),
              'Novo paciente cadastrado. Veja o e-mail em formato HTML.',
        {
                  from:     'nexusclinpb@gmail.com',
                  name:     'NEXUS CLIN Sistema',
                  htmlBody: htmlInterno
        }
            );
        Logger.log('E-mail interno enviado para sosthenes53@gmail.com');

  } catch (errInterno) {
        // Falha silenciosa — o fluxo do paciente nunca é afetado
      Logger.log('ERRO email interno: ' + errInterno);
  }

}

// ────────────────────────────────────────────────────────────────
// Lê a API key do PropertiesService (seguro)
// Para configurar: Projeto → Propriedades do script → Adicione:
//   Chave: CLAUDE_KEY | Valor: sua chave da Anthropic (sk-ant-...)
// ────────────────────────────────────────────────────────────────
function getClaudeKey() {
    const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_KEY');
    if (!key) throw new Error('CLAUDE_KEY não configurada em Propriedades do script');
    return key;
}

// ── Trigger principal ──
// Configure para rodar a cada 5 minutos OU onEdit na planilha
function processarNovosLeads() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RESPOSTAS_F2);
    if (!sheet) return;

  const dados     = sheet.getDataRange().getValues();
    const header    = dados[0];
    const colStatus = header.indexOf('Status');

  for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (linha[colStatus] !== 'Novo') continue;

      try {
              const d       = linhaParaObjeto(header, linha);
              const resumo  = chamarClaude(d);
              const pdfBlob = gerarPdfSolicitacao(d, resumo);

          salvarExames(ss, d, resumo);
              enviarEmail(d, resumo, pdfBlob);

          sheet.getRange(i + 1, colStatus + 1).setValue('Processado');
      } catch (err) {
              sheet.getRange(i + 1, colStatus + 1).setValue('Erro: ' + err.message.substring(0, 50));
      }

      Utilities.sleep(2000);
  }
}

// ── Converte linha da planilha em objeto ──
function linhaParaObjeto(header, linha) {
    const obj = {};
    header.forEach((col, i) => { obj[col] = linha[i] || ''; });
    return obj;
}

// ── Chama Claude API (Anthropic) ──
function chamarClaude(d) {
    const key  = getClaudeKey();
    const url  = 'https://api.anthropic.com/v1/messages';
    const sexo = (d['sexo'] || d['Sexo'] || '').toLowerCase();

  const prompt = `Você é um enfermeiro especialista em medicina integrativa e performance metabólica.

  Com base na anamnese abaixo, gere:
  1. Um RESUMO CLÍNICO objetivo (máximo 5 linhas) destacando os pontos mais relevantes para a consulta.
  2. Uma SUGESTÃO DE EXAMES por eixo clínico, com justificativa objetiva (máximo 2 linhas por exame).

  EIXOS OBRIGATÓRIOS: METABÓLICO E GLICÊMICO, HORMONAL, LIPÍDICO E CARDIOVASCULAR, TIREOIDIANO, INFLAMATÓRIO, NUTRICIONAL E MICRONUTRIENTES, HEPÁTICO E RENAL, HEMATOLÓGICO${sexo.includes('femin') ? ', GINECOLÓGICO E REPRODUTIVO' : ''}.

  Finalize com:
  ALERTAS CLÍNICOS
  [Máximo 3 pontos de atenção antes da consulta]

  DADOS DO PACIENTE:
  Nome: ${d['nome'] || d['Nome']}
  Sexo: ${d['sexo'] || d['Sexo']} | Data Nasc: ${d['data_nascimento'] || d['Data de Nascimento']}
  Objetivo: ${d['objetivo'] || d['Objetivo(s)']}
  Sint. gerais: ${d['sint_gerais'] || d['Sint. gerais']}
  Celular: ${d['celular'] || d['Celular']}`;

  const res = UrlFetchApp.fetch(url, {
        method:           'POST',
        contentType:      'application/json',
        muteHttpExceptions: true,
        headers: {
                'x-api-key':         key,
                'anthropic-version': '2023-06-01',
                'content-type':      'application/json'
        },
        payload: JSON.stringify({
                model:      'claude-opus-4-5',
                max_tokens: 1500,
                messages:   [{ role: 'user', content: prompt }]
        })
  });

  if (res.getResponseCode() !== 200) {
        throw new Error('Claude API HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 200));
  }

  const json = JSON.parse(res.getContentText());
    return json.content[0].text;
}

// ── Gera PDF de solicitação de exames (placeholder) ──
// Implemente aqui a geração real do PDF via HtmlService + DriveApp
function gerarPdfSolicitacao(d, resumo) {
    const html = `<html><body style="font-family:Arial,sans-serif;padding:32px">
        <h1 style="color:#0B1F3A">NEXUS CLIN — Solicitação de Exames</h1>
            <p><strong>Paciente:</strong> ${d['nome'] || d['Nome']}</p>
                <p><strong>Data:</strong> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')}</p>
                    <hr>
                        <pre style="white-space:pre-wrap;font-size:12px">${resumo}</pre>
                            <hr>
                                <p style="font-size:11px;color:#777">
                                      Enf. Sosthenes dos Santos Alves · COREN-PB 568176<br>
                                            Documento gerado automaticamente pelo Sistema NEXUS CLIN
                                                </p>
                                                  </body></html>`;

  const blob = Utilities.newBlob(html, MimeType.HTML, 'temp.html');
    const file = DriveApp.createFile(blob);
    const pdf  = file.getAs(MimeType.PDF);
    file.setTrashed(true); // remove o arquivo HTML temporário
  return pdf;
}

// ── Salva na aba de exames ──
function salvarExames(ss, d, texto) {
    const sheet = ss.getSheetByName(SHEET_EXAMES) || ss.insertSheet(SHEET_EXAMES);

  if (sheet.getLastRow() === 0) {
        const h = sheet.getRange(1, 1, 1, 5);
        h.setValues([['Data/Hora', 'Nome', 'Sexo / Nasc.', 'Objetivo', 'Resumo Clínico']]);
        h.setBackground('#0B1F3A').setFontColor('#fff').setFontWeight('bold');
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(5, 600);
  }

  sheet.appendRow([
        new Date(),
        d['nome']        || d['Nome'],
        (d['sexo']       || d['Sexo']),
        (d['objetivo']   || d['Objetivo(s)']),
        texto
      ]);

  const r = sheet.getRange(sheet.getLastRow(), 1, 1, 5);
    r.setFontSize(10).setVerticalAlignment('top').setWrap(true);
    r.setBackground(sheet.getLastRow() % 2 === 0 ? '#F0F4F8' : '#FFFFFF');
}

// ── Configura o trigger automático ──
// Rode esta função UMA VEZ para criar o trigger
function configurarTrigger() {
    ScriptApp.getProjectTriggers().forEach(t => {
          if (t.getHandlerFunction() === 'processarNovosLeads') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('processarNovosLeads').timeBased().everyMinutes(5).create();
    Logger.log('Trigger configurado: processarNovosLeads roda a cada 5 minutos.');
}

// ── Teste manual ──
function testeManual() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RESPOSTAS_F2);
    if (!sheet) { Logger.log('Planilha "Respostas" não encontrada.'); return; }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('Nenhuma resposta ainda.'); return; }

  const header    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colStatus = header.indexOf('Status');
    sheet.getRange(lastRow, colStatus + 1).setValue('Novo');

  processarNovosLeads();
    Logger.log('Teste concluído. Verifique aba "Sugestão de Exames" e os e-mails.');
}
