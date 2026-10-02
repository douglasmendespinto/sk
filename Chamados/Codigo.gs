/**
 * CHAMADOS DE COMPRA — SK Infraestrutura
 * Recebe os pedidos do formulário, grava na planilha, guarda os anexos
 * no Google Drive e envia o e-mail para os responsáveis.
 *
 * Cole este arquivo em: Planilha Google > Extensões > Apps Script.
 */

// ====== PREENCHA AQUI ======================================================
const CONFIG = {
  // Senha que os colaboradores digitam para abrir o formulário
  SENHA: 'TROQUE-ESTA-SENHA',

  // Quem recebe cada chamado (você, sua assistente e seu líder)
  EMAILS: [
    'seu.email@exemplo.com',
    'assistente@exemplo.com',
    'lider@exemplo.com'
  ],

  OBRA: 'Projeto Sucuriú (ARAUCO)',
  ABA: 'Chamados',
  PASTA_ANEXOS: 'Chamados de compra - Anexos',
  FUSO: 'America/Campo_Grande'
};
// ===========================================================================

const LIMITE_ANEXOS = 6;
const LIMITE_BYTES = 15 * 1024 * 1024; // soma dos anexos por chamado
const LIMITE_ITENS = 40;
const TIPOS_ACEITOS = ['image/jpeg', 'image/png', 'application/pdf'];
const URGENCIAS = ['Normal', 'Urgente', 'Emergencial'];

const CABECALHO = [
  'Nº', 'Data/hora', 'Solicitante', 'Setor / frente', 'Contato', 'Urgência',
  'Necessário até', 'Item', 'Descrição', 'Qtd', 'Un', 'Aplicação / obs. do item',
  'Justificativa', 'Anexos', 'Status'
];

/** Rode esta função UMA vez pelo editor para liberar as permissões. */
function autorizar() {
  obterAba_();
  obterPasta_();
  Logger.log('Envios de e-mail disponíveis hoje: ' + MailApp.getRemainingDailyQuota());
}

function doGet() {
  return ContentService.createTextOutput('Chamados de compra: serviço ativo.');
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (String(d.senha || '') !== CONFIG.SENHA) return json_({ ok: false, erro: 'senha' });
    if (d.acao === 'entrar') return json_({ ok: true });
    if (d.acao === 'enviar') return json_(registrar_(d));
    return json_({ ok: false, erro: 'acao' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, erro: 'falha', detalhe: String(err) });
  }
}

function registrar_(d) {
  const solicitante = texto_(d.solicitante, 80);
  const setor = texto_(d.setor, 80);
  const contato = texto_(d.contato, 60);
  const urgencia = URGENCIAS.indexOf(d.urgencia) >= 0 ? d.urgencia : 'Normal';
  const prazo = /^\d{4}-\d{2}-\d{2}$/.test(d.prazo || '') ? d.prazo.split('-').reverse().join('/') : '';
  const justificativa = texto_(d.justificativa, 1000);

  const itens = (Array.isArray(d.itens) ? d.itens : []).slice(0, LIMITE_ITENS).map(function (i) {
    return {
      descricao: texto_(i.descricao, 300),
      qtd: Number(String(i.qtd).replace(',', '.')),
      un: texto_(i.un, 12),
      obs: texto_(i.obs, 300)
    };
  }).filter(function (i) { return i.descricao && i.qtd > 0; });

  if (!solicitante || !setor || itens.length === 0) return { ok: false, erro: 'dados' };

  // Anexos
  const brutos = (Array.isArray(d.anexos) ? d.anexos : []).slice(0, LIMITE_ANEXOS);
  let total = 0;
  const blobs = [];
  brutos.forEach(function (a) {
    if (TIPOS_ACEITOS.indexOf(a.tipo) < 0 || !a.dados) return;
    const bytes = Utilities.base64Decode(a.dados);
    total += bytes.length;
    if (total > LIMITE_BYTES) throw new Error('Anexos acima do limite.');
    const nome = String(a.nome || 'anexo').replace(/[^\w.\- ()À-ÿ]/g, '_').slice(0, 80);
    blobs.push(Utilities.newBlob(bytes, a.tipo, nome));
  });

  // Número sequencial + gravação (com trava para não repetir número)
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  let numero, links = [];
  const agora = new Date();
  try {
    const props = PropertiesService.getScriptProperties();
    const seq = Number(props.getProperty('SEQ') || 0) + 1;
    props.setProperty('SEQ', String(seq));
    numero = ('0000' + seq).slice(-4);

    if (blobs.length) {
      const pasta = obterPasta_();
      blobs.forEach(function (b) {
        b.setName(numero + ' - ' + b.getName());
        links.push(pasta.createFile(b).getUrl());
      });
    }

    const dataHora = Utilities.formatDate(agora, CONFIG.FUSO, 'dd/MM/yyyy HH:mm');
    const linhas = itens.map(function (i, n) {
      return [
        numero, dataHora, solicitante, setor, contato, urgencia, prazo,
        n + 1, i.descricao, i.qtd, i.un, i.obs, justificativa,
        links.join('\n'), 'Aberto'
      ].map(seguro_);
    });
    const aba = obterAba_();
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, CABECALHO.length).setValues(linhas);
  } finally {
    trava.releaseLock();
  }

  enviarEmail_({
    numero: numero, agora: agora, solicitante: solicitante, setor: setor, contato: contato,
    urgencia: urgencia, prazo: prazo, justificativa: justificativa, itens: itens,
    blobs: blobs, links: links
  });

  return { ok: true, numero: numero };
}

function enviarEmail_(c) {
  const cor = { Normal: '#1E7005', Urgente: '#B7791F', Emergencial: '#B42318' }[c.urgencia];
  const td = 'padding:8px 10px;border-bottom:1px solid #E3E7EC;font-size:14px;vertical-align:top;';
  const th = td + 'text-align:left;background:#283556;color:#fff;font-weight:600;';

  const linhasItens = c.itens.map(function (i, n) {
    return '<tr><td style="' + td + '">' + (n + 1) + '</td>' +
      '<td style="' + td + '">' + esc_(i.descricao) +
      (i.obs ? '<br><span style="color:#5B6676;font-size:13px;">' + esc_(i.obs) + '</span>' : '') + '</td>' +
      '<td style="' + td + 'white-space:nowrap;text-align:right;">' + esc_(String(i.qtd).replace('.', ',')) + ' ' + esc_(i.un) + '</td></tr>';
  }).join('');

  function linha(rotulo, valor) {
    return valor ? '<tr><td style="padding:3px 14px 3px 0;color:#5B6676;font-size:14px;">' + rotulo +
      '</td><td style="padding:3px 0;font-size:14px;">' + esc_(valor) + '</td></tr>' : '';
  }

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#1B2433;max-width:640px;">' +
    '<div style="height:4px;background:#1E7005;"></div><div style="height:4px;background:#004AAD;"></div>' +
    '<div style="height:4px;background:#FFDE59;"></div>' +
    '<h2 style="margin:18px 0 4px;font-size:20px;">Pedido de compra nº ' + c.numero + '</h2>' +
    '<p style="margin:0 0 14px;color:#5B6676;font-size:14px;">' + esc_(CONFIG.OBRA) + '</p>' +
    '<p style="margin:0 0 14px;"><span style="background:' + cor + ';color:#fff;padding:4px 10px;' +
    'border-radius:4px;font-size:13px;font-weight:700;">' + c.urgencia + '</span></p>' +
    '<table style="border-collapse:collapse;margin-bottom:16px;">' +
    linha('Solicitante', c.solicitante) + linha('Setor / frente', c.setor) + linha('Contato', c.contato) +
    linha('Necessário até', c.prazo) +
    linha('Enviado em', Utilities.formatDate(c.agora, CONFIG.FUSO, 'dd/MM/yyyy HH:mm')) +
    '</table>' +
    '<table style="border-collapse:collapse;width:100%;margin-bottom:16px;">' +
    '<tr><th style="' + th + '">#</th><th style="' + th + '">Item</th><th style="' + th + 'text-align:right;">Qtd</th></tr>' +
    linhasItens + '</table>' +
    (c.justificativa ? '<p style="font-size:14px;margin:0 0 14px;"><b>Justificativa:</b><br>' +
      esc_(c.justificativa).replace(/\n/g, '<br>') + '</p>' : '') +
    (c.links.length ? '<p style="font-size:14px;margin:0 0 14px;">' + c.links.length +
      ' anexo(s) neste e-mail e na pasta do Drive.</p>' : '') +
    '<p style="font-size:13px;color:#5B6676;">Histórico completo: <a href="' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">abrir planilha</a></p></div>';

  const prefixo = c.urgencia === 'Normal' ? '' : '[' + c.urgencia.toUpperCase() + '] ';
  MailApp.sendEmail({
    to: CONFIG.EMAILS.join(','),
    subject: prefixo + 'Pedido de compra nº ' + c.numero + ' - ' + c.solicitante +
      ' (' + c.itens.length + (c.itens.length === 1 ? ' item)' : ' itens)'),
    htmlBody: html,
    attachments: c.blobs,
    name: 'Chamados de compra SK'
  });
}

// ---------- utilitários ----------
function obterAba_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(CONFIG.ABA);
  if (!aba) aba = planilha.insertSheet(CONFIG.ABA);
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO])
      .setFontWeight('bold').setBackground('#283556').setFontColor('#FFFFFF');
    aba.setFrozenRows(1);
  }
  return aba;
}

function obterPasta_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* pasta apagada: cria outra */ }
  }
  const pasta = DriveApp.createFolder(CONFIG.PASTA_ANEXOS);
  props.setProperty('PASTA_ID', pasta.getId());
  return pasta;
}

function texto_(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
}

/** Impede que um texto digitado vire fórmula na planilha. */
function seguro_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
