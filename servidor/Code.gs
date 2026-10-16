/*
  Servidor das pesquisas Maxipark — Google Apps Script ligado a uma Planilha Google.
  Recebe as respostas dos tablets (pesquisa com clientes, pesquisa com a equipe e avaliação do conselho) e devolve todas elas para
  qualquer aparelho que tenha a mesma chave. Uma linha por resposta, em três abas: "Clientes", "Equipe" e "Conselho".
  Instalação passo a passo: LEIA-ME.md, seção 11.

  Como funciona:
    POST  { op:'add', chave, itens:[ ...respostas... ] }   -> grava as respostas novas (id repetido é ignorado)
    GET   ?op=list&app=pesquisa-clientes-maxipark&chave=... -> devolve todas as respostas daquele app
    GET   ?op=ping&chave=...                                -> teste de ligação
*/
const CHAVE = 'TROQUE-ESTA-CHAVE';   // a MESMA chave de sincronização digitada nas Configurações dos apps (letras e números, sem espaços)
const ABAS = { 'pesquisa-clientes-maxipark': 'Clientes', 'pesquisa-equipe-maxipark': 'Equipe', 'pesquisa-conselho-maxipark': 'Conselho' };
const FIXAS = ['recebido_em', 'app', 'id', 'unidade', 'data', 'hora_ou_rodada', 'json'];   // colunas fixas; depois vem uma coluna por campo da resposta
const CHAVE_EXEMPLO = 'TROQUE-ESTA-CHAVE';

function verificaChave(chave) {   // devolve null se está tudo certo, ou a resposta de erro
  if (CHAVE === CHAVE_EXEMPLO || !CHAVE) return json({ ok: false, erro: 'configurar-chave' });   // ninguém lê nem grava enquanto a chave de exemplo não for trocada
  if (String(chave || '') !== CHAVE) return json({ ok: false, erro: 'chave' });
  return null;
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const erro = verificaChave(body.chave); if (erro) return erro;
    if (body.op === 'relatorio-equipe') return json(enviaRelatorioEquipe(body.relatorio || {}));   // relatorio da visita (pesquisa com a equipe) por e-mail
    if (body.op === 'responder-relatorio') return json(responderRelatorio(body));   // resposta escrita no app: sai na mesma conversa do e-mail, com o histórico
    if (body.op === 'historico-add') {            // eventos novos vindos dos apps
      const ws = abaHistorico();
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      let n = 0;
      try {
        const ids = idsHistorico(ws);
        const linhas = [];
        (body.itens || []).forEach(function (it) {
          const id = String((it && it.id) || ''); if (!id || ids[id]) return; ids[id] = true;
          linhas.push(linhaHistorico(it)); n++;
        });
        if (linhas.length) ws.getRange(ws.getLastRow() + 1, 1, linhas.length, COLS_HIST.length).setValues(linhas);
      } finally { lock.releaseLock(); }
      return json({ ok: true, n: n });
    }
    if (body.op === 'historico-apagar') {         // so o administrador chama isto
      const ws = abaHistorico();
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      let n = 0;
      try {
        const alvo = {}; (body.ids || []).forEach(function (x) { alvo[String(x)] = true; });
        const last = ws.getLastRow();
        if (last >= 2) {
          const vals = ws.getRange(2, 1, last - 1, COLS_HIST.length).getValues();
          for (var i = vals.length - 1; i >= 0; i--) {
            if (alvo[String(vals[i][0])]) { ws.deleteRow(i + 2); n++; }
          }
        }
      } finally { lock.releaseLock(); }
      return json({ ok: true, n: n });
    }
    if (body.op === 'historico-editar') {         // so o administrador chama isto
      const ws = abaHistorico();
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      let n = 0;
      try {
        const last = ws.getLastRow();
        if (last >= 2) {
          const vals = ws.getRange(2, 1, last - 1, COLS_HIST.length).getValues();
          const porId = {}; vals.forEach(function (r, i) { porId[String(r[0])] = i; });
          (body.itens || []).forEach(function (it) {
            const i = porId[String((it && it.id) || '')];
            if (i === undefined) return;
            ws.getRange(i + 2, 1, 1, COLS_HIST.length).setValues([linhaHistorico(it)]); n++;
          });
        }
      } finally { lock.releaseLock(); }
      return json({ ok: true, n: n });
    }
    if (body.op === 'usuarios-salvar') {          // cadastro de usuarios do app: guarda a lista inteira
      const ws = abaUsuarios();
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      try { ws.getRange(1, 1).setValue(JSON.stringify(body.dados || {})); } finally { lock.releaseLock(); }
      return json({ ok: true, salvo: true });
    }
    const itens = Array.isArray(body.itens) ? body.itens : [body];
    let n = 0, dup = 0, ignorados = 0;
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      const porApp = {};
      itens.forEach(function (it) { if (it && ABAS[it.app]) (porApp[it.app] = porApp[it.app] || []).push(it); else ignorados++; });
      Object.keys(porApp).forEach(function (app) {
        const ws = aba(app); const ids = idsExistentes(ws);
        const chaves = camposDe(ws, porApp[app]);   // garante uma coluna por campo
        const linhas = [];
        porApp[app].forEach(function (it) {
          const id = String(it.id || ''); if (!id || ids.has(id)) { dup++; return; } ids.add(id);
          const fixas = [new Date(), app, id, semFormula(String(it.unidade || '')), String(it.data || ''), String(it.hora || it.rodada || ''), semFormula(JSON.stringify(it))];
          linhas.push(fixas.concat(chaves.map(function (k) { return celula(it[k]); }))); n++;
        });
        if (linhas.length) ws.getRange(ws.getLastRow() + 1, 1, linhas.length, FIXAS.length + chaves.length).setValues(linhas);
      });
    } finally { lock.releaseLock(); }
    return json({ ok: true, n: n, dup: dup, ignorados: ignorados });
  } catch (err) { return json({ ok: false, erro: String(err) }); }
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const erro = verificaChave(p.chave); if (erro) return erro;
    if (p.op === 'respostas-relatorio') {            // relatórios enviados + respostas recebidas por e-mail (filtra por unidade e rodada)
      let novas = 0, aviso = '';
      try { novas = coletaRespostasRelatorios(); } catch (e) { aviso = String(e); }
      const filtro = o => (!p.unidade || o.unidade === p.unidade) && (!p.rodada || o.rodada === p.rodada);
      return json({ ok: true, novas: novas, aviso: aviso, relatorios: lerAba(ABA_RELATORIOS, COLS_REL).filter(filtro), respostas: lerAba(ABA_RESP_REL, COLS_RESP).filter(filtro) });
    }
    if (p.op === 'ping') return json({ ok: true, servidor: 'apps-script', versao: 5, abas: Object.keys(ABAS), usuarios: true, historico: true, relatorio: true });
    if (p.op === 'historico') {
      const ws = abaHistorico(); const last = ws.getLastRow();
      if (last < 2) return json({ ok: true, itens: [] });
      const vals = ws.getRange(2, 1, last - 1, COLS_HIST.length).getValues();
      const itens = vals.map(function (r) {
        const o = {}; COLS_HIST.forEach(function (c, i) { o[c] = (r[i] instanceof Date) ? r[i].toISOString() : String(r[i] === null || r[i] === undefined ? '' : r[i]); });
        return o;
      });
      return json({ ok: true, itens: itens });
    }
    if (p.op === 'usuarios') {                    // cadastro de usuarios do app
      const t = String(abaUsuarios().getRange(1, 1).getValue() || '').trim();
      let d = null; try { d = t ? JSON.parse(t) : null; } catch (x) { d = null; }
      return json({ ok: true, dados: d });
    }
    if (!ABAS[p.app]) return json({ ok: false, erro: 'app' });
    const ws = aba(p.app); const last = ws.getLastRow();
    if (last < 2) return json({ ok: true, itens: [] });
    const desde = p.desde ? new Date(p.desde) : null;                 // incremental: só o que chegou depois de "desde" (os apps mandam a última atualização menos 15 min)
    const semContatos = p.contatos !== '1';                           // contatos de clientes só saem para quem pediu de propósito (computador da área de CX)
    const vals = ws.getRange(2, 1, last - 1, FIXAS.length).getValues();
    const iJson = FIXAS.indexOf('json'); const itens = [];
    vals.forEach(function (r) {
      if (desde && !isNaN(desde.getTime()) && r[0] instanceof Date && r[0] < desde) return;
      if (!r[iJson]) return;
      try { const it = JSON.parse(r[iJson]); if (semContatos && it && it.contato !== undefined) it.contato = ''; itens.push(it); } catch (x) { /* linha editada à mão: ignora */ }
    });
    return json({ ok: true, itens: itens });
  } catch (err) { return json({ ok: false, erro: String(err) }); }
}

function aba(app) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(ABAS[app]);
  if (!ws) { ws = ss.insertSheet(ABAS[app]); ws.appendRow(FIXAS); ws.setFrozenRows(1); ws.getRange(1, 1, 1, FIXAS.length).setFontWeight('bold'); }
  return ws;
}
const COLS_HIST = ['id', 'quando', 'quem_nome', 'quem_email', 'evento', 'pesquisa', 'unidade', 'detalhe'];
function abaHistorico() {   // uma linha por evento: quem fez o que, quando
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('Historico');
  if (!ws) {
    ws = ss.insertSheet('Historico');
    ws.appendRow(COLS_HIST); ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, COLS_HIST.length).setFontWeight('bold');
    ws.setColumnWidth(2, 160); ws.setColumnWidth(8, 320);
  }
  /* A data fica como TEXTO puro: se a planilha converter em data, ela mexe no fuso
     e o horario mostrado deixa de ser o que aconteceu de verdade. */
  try { ws.getRange('A:H').setNumberFormat('@'); } catch (e) {}
  return ws;
}
function idsHistorico(ws) {
  const last = ws.getLastRow(); const m = {};
  if (last >= 2) ws.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { if (r[0] !== '') m[String(r[0])] = true; });
  return m;
}
function linhaHistorico(it) {
  return COLS_HIST.map(function (c) { return semFormula(String((it && it[c]) === undefined || (it && it[c]) === null ? '' : it[c])); });
}
function abaUsuarios() {   // uma aba so para o cadastro de quem entra no app; o JSON inteiro fica na celula A1
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('Usuarios');
  if (!ws) {
    ws = ss.insertSheet('Usuarios');
    ws.getRange(1, 2).setValue('<- Cadastro de usuarios do app, na celula A1. Nao editar a mao: use a tela AVALIADORES/USUARIOS da pagina inicial.');
    ws.setColumnWidth(1, 420);
  }
  return ws;
}
function cabecalho(ws) { const lc = ws.getLastColumn(); return lc ? ws.getRange(1, 1, 1, lc).getValues()[0].map(String) : []; }
function camposDe(ws, itens) {   // colunas depois das fixas: os campos das respostas, na ordem em que aparecem; acrescenta as que faltam
  const head = cabecalho(ws); const atuais = head.slice(FIXAS.length); const set = {}; atuais.forEach(function (k) { set[k] = true; });
  const novas = [];
  itens.forEach(function (it) { Object.keys(it).forEach(function (k) { if (!set[k]) { set[k] = true; novas.push(k); } }); });
  if (novas.length) { ws.getRange(1, head.length + 1, 1, novas.length).setValues([novas]); ws.getRange(1, head.length + 1, 1, novas.length).setFontWeight('bold'); }
  return atuais.concat(novas);
}
function idsExistentes(ws) { const last = ws.getLastRow(); const s = new Set(); if (last >= 2) ws.getRange(2, FIXAS.indexOf('id') + 1, last - 1, 1).getValues().forEach(function (r) { if (r[0] !== '') s.add(String(r[0])); }); return s; }
function celula(v) {   // texto digitado pelo público nunca vira fórmula na planilha (=, +, -, @ no início ganham um apóstrofo)
  if (v === null || v === undefined) return '';
  if (Array.isArray(v)) return semFormula(v.map(function (x) { return x === null || x === undefined ? '' : String(x); }).join(' | '));
  if (typeof v === 'object') return semFormula(JSON.stringify(v));
  if (typeof v === 'string') return semFormula(v);
  return v;
}
function semFormula(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Relatório da visita (pesquisa com a equipe) ----------
   O tablet monta o relatório (HTML) ao encerrar a visita numa unidade e diz para quem vai (o e-mail digitado na hora,
   até 5). Aqui ele vai por e-mail, com um PDF anexo. Se o tablet não mandar destinatário, vai para o administrador. */
const RELATORIO_PADRAO = ['raphael@maxipark.com.br'];
const RELATORIO_MAX_PARA = 5;
function enviaRelatorioEquipe(r) {
  const html = String(r.html || '');
  if (!html) return { ok: false, erro: 'relatorio-vazio' };
  if (html.length > 400000) return { ok: false, erro: 'relatorio-grande' };
  const ok = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  let para = (Array.isArray(r.para) ? r.para : []).map(function (x) { return String(x || '').trim().toLowerCase(); }).filter(Boolean);
  if (para.some(function (x) { return !ok.test(x); })) return { ok: false, erro: 'email-invalido' };
  if (para.length > RELATORIO_MAX_PARA) return { ok: false, erro: 'emails-demais' };
  if (!para.length) para = RELATORIO_PADRAO;
  const assunto = String(r.assunto || 'Relatório da visita - Pesquisa equipe').slice(0, 200);
  const nome = String(r.arquivo || 'relatorio-equipe').replace(/[^\w.-]+/g, '_').slice(0, 120) + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'relatorio.html').getAs('application/pdf').setName(nome);
  const rid = idRelatorio();
  const assuntoFinal = assunto.slice(0, 180) + ' [' + rid + ']';
  const aviso = '<div style="font-family:Arial,Helvetica,sans-serif;background:#efe9fb;border:1px solid #d9cdf3;border-radius:10px;padding:12px 14px;margin:0 auto 14px;max-width:720px;color:#1f1a33;font-size:14px;text-align:left">'
    + '<b>Responda a este e-mail</b> com o que vai ser feito em cada ponto de melhoria e o <b>prazo</b> de cada ação. '
    + 'Pode responder mais de uma vez: todas as respostas ficam guardadas no histórico desta pesquisa (código ' + rid + ').</div>';
  MailApp.sendEmail({ to: para.join(','), subject: assuntoFinal, htmlBody: aviso + html, attachments: [pdf], name: 'Pesquisas Maxipark' });
  try { abaSimples(ABA_RELATORIOS, COLS_REL).appendRow([rid, new Date(), semFormula(String(r.unidade || '')), String(r.rodada || ''), "'" + String(r.data || ''), para.join(', '), semFormula(assuntoFinal)]); } catch (e) {}
  return { ok: true, enviado: para.length, id: rid };
}

/* ---------- Respostas aos relatórios da visita (servidor v13, 15/10/2026) ----------
   Cada relatório sai com um código no assunto ([REL-aammdd-XXXXX]) e fica registrado na aba "Relatorios".
   Quem recebe responde o e-mail com o plano de melhoria e o prazo. Sempre que o app pede as respostas
   (op=respostas-relatorio), o servidor procura no Gmail desta conta as respostas com esse código e guarda
   cada uma na aba "RespostasRelatorios" (uma linha por e-mail; pode haver várias por relatório).
   Precisa da permissão do Gmail (rodar autorizarRespostas pelo editor uma vez). */
const ABA_RELATORIOS = 'Relatorios', ABA_RESP_REL = 'RespostasRelatorios';
const REMETENTE_RELATORIO = 'raphaelfrascinofl@gmail.com';   // conta do Gmail de onde saem os relatórios (e onde chegam as respostas)
const COLS_REL = ['id', 'enviado_em', 'unidade', 'rodada', 'data', 'para', 'assunto'];
const COLS_RESP = ['msg_id', 'relatorio_id', 'unidade', 'rodada', 'data', 'de', 'recebido_em', 'texto'];
function abaSimples(nome, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(nome);
  if (!ws) { ws = ss.insertSheet(nome); ws.appendRow(cols); ws.setFrozenRows(1); ws.getRange(1, 1, 1, cols.length).setFontWeight('bold'); }
  return ws;
}
function lerAba(nome, cols) {
  const ws = abaSimples(nome, cols), l = ws.getLastRow();
  if (l < 2) return [];
  return ws.getRange(2, 1, l - 1, cols.length).getValues().map(function (r) {
    const o = {}; cols.forEach(function (c, i) { o[c] = r[i] instanceof Date ? r[i].toISOString() : String(r[i] === null || r[i] === undefined ? '' : r[i]); }); return o;
  });
}
function idRelatorio() { return 'REL-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 5).toUpperCase(); }
function limpaResposta(t) {   // guarda a resposta inteira com a assinatura e o cabeçalho (quem respondeu a quem); só tira o relatório citado que vem embaixo
  const linhas = String(t || '').replace(/\r/g, '').split('\n'), out = [];
  for (let i = 0; i < linhas.length; i++) {
    const l = linhas[i];
    if (/^\s*>/.test(l)) break;                                                     // começo do e-mail citado (estilo Gmail)
    if (/(Em .{0,250}escreveu:|On .{0,250}wrote:)\s*$/.test(l)) { out.push(l.trim()); break; }   // mantém a linha "Fulano escreveu:"
    if (/^\s*\*?(De|From):\s/.test(l)) {                                          // cabeçalho do Outlook: mantém De/Enviado/Para/Cc/Assunto
      for (let k = i; k < linhas.length && k - i < 8 && linhas[k].trim() !== ''; k++) out.push(linhas[k]);
      break;
    }
    out.push(l);
  }
  return out.join('\n').trim().slice(0, 5000);
}
function coletaRespostasRelatorios() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return 0;
  try {
    const rels = {};
    lerAba(ABA_RELATORIOS, COLS_REL).forEach(function (r) { rels[r.id] = r; });
    if (!Object.keys(rels).length) return 0;
    const wsA = abaSimples(ABA_RESP_REL, COLS_RESP), vistos = {};
    lerAba(ABA_RESP_REL, COLS_RESP).forEach(function (r) { vistos[r.msg_id] = true; });
    const eu = REMETENTE_RELATORIO;   // a conta dona do servidor: o próprio envio não é resposta
    const novas = [];
    GmailApp.search('subject:"Relatório da visita" newer_than:365d', 0, 200).forEach(function (th) {
      const m = /\[(REL-\d{6}-[A-Z0-9]{5})\]/.exec(th.getFirstMessageSubject() || ''); if (!m || !rels[m[1]]) return;
      const rel = rels[m[1]];
      th.getMessages().forEach(function (msg) {
        const id = msg.getId(); if (vistos[id]) return;
        const de = String(msg.getFrom() || ''); if (eu && de.toLowerCase().indexOf(eu) >= 0) return;
        vistos[id] = true;
        novas.push([id, m[1], rel.unidade, rel.rodada, rel.data, semFormula(soEmail(de)), msg.getDate(), semFormula(limpaResposta(msg.getPlainBody()))]);
      });
    });
    if (novas.length) wsA.getRange(wsA.getLastRow() + 1, 1, novas.length, COLS_RESP.length).setValues(novas);
    return novas.length;
  } finally { lock.releaseLock(); }
}
/* Rode UMA vez pelo editor para o Google pedir a permissão de ler as respostas no Gmail. */
function autorizarRespostas() { const n = GmailApp.search('subject:"Relatório da visita"', 0, 1).length; Logger.log('Acesso ao Gmail OK (' + n + ')'); return coletaRespostasRelatorios(); }
/* Rode UMA vez pelo editor para o Google pedir a permissão de enviar e-mail. Não envia nada. */
function autorizarEmail() { const resta = MailApp.getRemainingDailyQuota(); Logger.log('Permissão de e-mail OK. E-mails que ainda podem sair hoje: ' + resta); return resta; }

/* ---------- Responder pelo app (servidor v17, 15/10/2026) ----------
   O app manda { id do relatório, texto, autor, email }. O servidor acha a conversa no Gmail pelo código do relatório,
   responde para todos da conversa com o texto + o histórico (relatório e todas as respostas) e guarda a resposta
   na aba "RespostasRelatorios" (msg_id começa com "app-"), para aparecer no histórico do app. */
/* No histórico fica só o e-mail de quem respondeu (pedido do Raphael em 15/10/2026). */
function soEmail(de) { const m = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/.exec(String(de || '')); return m ? m[0].toLowerCase() : String(de || '').trim(); }
function escHtml(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function quandoBR(d) { const x = d instanceof Date ? d : new Date(d); return isNaN(x) ? String(d || '') : Utilities.formatDate(x, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'); }
function responderRelatorio(b) {
  const id = String(b.id || '').trim(), texto = String(b.texto || '').trim();
  if (!/^REL-\d{6}-[A-Z0-9]{5}$/.test(id)) return { ok: false, erro: 'relatorio-invalido' };
  if (!texto) return { ok: false, erro: 'texto-vazio' };
  if (texto.length > 8000) return { ok: false, erro: 'texto-grande' };
  const rel = lerAba(ABA_RELATORIOS, COLS_REL).filter(function (r) { return r.id === id; })[0];
  if (!rel) return { ok: false, erro: 'relatorio-nao-encontrado' };
  try { coletaRespostasRelatorios(); } catch (e) {}   // pega as respostas mais novas antes de montar o histórico
  const ths = GmailApp.search('subject:"' + id + '"', 0, 10).filter(function (t) { return (t.getFirstMessageSubject() || '').indexOf(id) >= 0; });
  if (!ths.length) return { ok: false, erro: 'conversa-nao-encontrada' };
  const th = ths[0], msgs = th.getMessages(), ultima = msgs[msgs.length - 1];
  const autor = String(b.autor || '').trim().slice(0, 80) || 'Experiência do Cliente', email = String(b.email || '').trim().slice(0, 120);
  const hist = lerAba(ABA_RESP_REL, COLS_RESP).filter(function (r) { return r.relatorio_id === id; })
    .sort(function (a, c) { return String(a.recebido_em).localeCompare(String(c.recebido_em)); });
  const bloco = function (de, quando, txt) { return '<div style="border-left:3px solid #d9cdf3;padding:6px 0 6px 12px;margin:10px 0"><div style="font-size:12px;color:#6e6a7c"><b>' + escHtml(de) + '</b> · ' + escHtml(quandoBR(quando)) + '</div><div style="white-space:pre-wrap;font-size:14px;margin-top:4px">' + escHtml(txt) + '</div></div>'; };
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#1f1a33;max-width:720px;font-size:14px;line-height:1.45">'
    + '<div style="white-space:pre-wrap;font-size:15px">' + escHtml(texto) + '</div>'
    + '<p style="margin:10px 0 0">' + escHtml(autor) + (email ? ' · ' + escHtml(email) : '') + '<br><span style="color:#6e6a7c;font-size:12px">Enviado pelo app de pesquisas Maxipark</span></p>'
    + '<div style="margin-top:22px;padding-top:12px;border-top:1px solid #e3dff0"><div style="font-size:13px;font-weight:700;color:#4e3572">Histórico desta pesquisa · ' + escHtml(rel.unidade) + ' · ' + escHtml(rel.rodada) + ' · código ' + escHtml(id) + '</div>'
    + bloco('Relatório enviado para ' + rel.para, rel.enviado_em, rel.assunto)
    + hist.map(function (r) { return bloco(soEmail(r.de), r.recebido_em, r.texto); }).join('')
    + '</div></div>';
  ultima.replyAll(texto + '\n\n' + autor, { htmlBody: html, name: 'Pesquisas Maxipark' });
  abaSimples(ABA_RESP_REL, COLS_RESP).appendRow(['app-' + Utilities.getUuid(), id, rel.unidade, rel.rodada, "'" + String(rel.data || '').slice(0, 10), semFormula(email || autor), new Date(), semFormula(texto)]);
  return { ok: true };
}
