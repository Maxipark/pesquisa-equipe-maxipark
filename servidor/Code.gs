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
    if (body.op === 'recuperar-senha') return json(recuperarSenha(body));   // públicos: esqueci a senha, criar senha pelo link e entrar
    if (body.op === 'definir-senha') return json(definirSenha(body));
    if (body.op === 'entrar') return json(entrarPeloServidor(body));
    if (body.op === 'definir-pin') return json(definirPin(body));   // público, com a prova da senha: cada usuário escolhe o próprio PIN
    const erro = verificaChave(body.chave); if (erro) return erro;
    if (body.op === 'revisao-salvar') return json(revisaoSalvar(body.item || {}));   // revisão das perguntas: uma linha por pergunta e por pessoa (21/10/2026)
    if (body.op === 'convidar') return json(convidarUsuario(body));   // administrador: manda o e-mail de primeiro acesso
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
      try { ws.getRange(1, 1).setValue(JSON.stringify(protegeCredenciais(body.dados || {}))); } finally { lock.releaseLock(); }
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
    if (p.state && (p.code || p.error)) return outlookRetorno(e);   // volta do login da Microsoft (conectar a caixa pesquisa@)
    if (p.op === 'token-info') return json(infoToken(p));   // públicos: dados do link de senha e sal para entrar
    if (p.op === 'sal') return json(salDoUsuario(p));
    const erro = verificaChave(p.chave); if (erro) return erro;
    if (p.op === 'respostas-relatorio') {            // relatórios enviados + respostas recebidas por e-mail (filtra por unidade e rodada)
      let novas = 0, aviso = '';
      try { novas = coletaRespostasRelatorios(); } catch (e) { aviso = String(e); }
      const filtro = o => (!p.unidade || o.unidade === p.unidade) && (!p.rodada || o.rodada === p.rodada);
      return json({ ok: true, novas: novas, aviso: aviso, relatorios: lerAba(ABA_RELATORIOS, COLS_REL).filter(filtro), respostas: lerAba(ABA_RESP_REL, COLS_RESP).filter(filtro) });
    }
    if (p.op === 'revisao') return json({ ok: true, itens: lerAba(ABA_REVISAO, COLS_REVISAO) });   // revisão das perguntas: tudo, para todo mundo ver
    if (p.op === 'ping') return json({ ok: true, servidor: 'apps-script', versao: 6, revisao: true, abas: Object.keys(ABAS), usuarios: true, historico: true, relatorio: true, outlook: outlookConectado() });
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
  let canal = 'gmail';
  if (outlookConectado()) { outlookEnviar(para, assuntoFinal, aviso + html, pdf); canal = 'outlook'; }   // sai da caixa pesquisa@maxipark.com.br
  else MailApp.sendEmail({ to: para.join(','), subject: assuntoFinal, htmlBody: aviso + html, attachments: [pdf], name: 'Pesquisas Maxipark' });
  try { abaSimples(ABA_RELATORIOS, COLS_REL).appendRow([rid, new Date(), semFormula(String(r.unidade || '')), String(r.rodada || ''), "'" + String(r.data || ''), para.join(', '), semFormula(assuntoFinal), canal]); } catch (e) {}
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
const COLS_REL = ['id', 'enviado_em', 'unidade', 'rodada', 'data', 'para', 'assunto', 'canal'];   // canal: outlook (caixa pesquisa@) ou gmail (antigo)
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
function coletaRespostasRelatorios() {   // respostas que chegam na caixa pesquisa@ (Outlook). O Gmail pessoal não é mais usado (16/10/2026)
  let n = 0;
  if (outlookConectado()) { try { n += coletaRespostasOutlook(); } catch (e) { Logger.log(e); } }
  return n;
}
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
  if (rel.canal === 'outlook') {                       // conversa da caixa pesquisa@ (Outlook)
    if (!outlookConectado()) return { ok: false, erro: 'outlook-desconectado' };
    if (!outlookResponder(id, html)) return { ok: false, erro: 'conversa-nao-encontrada' };
  } else return { ok: false, erro: 'relatorio-antigo-gmail' };   // só os testes antigos saíram pelo Gmail pessoal, que não é mais usado
  abaSimples(ABA_RESP_REL, COLS_RESP).appendRow(['app-' + Utilities.getUuid(), id, rel.unidade, rel.rodada, "'" + String(rel.data || '').slice(0, 10), semFormula(email || autor), new Date(), semFormula(texto)]);
  return { ok: true };
}

/* ---------- Caixa pesquisa@maxipark.com.br (Outlook / Microsoft 365, servidor v22, 16/10/2026: dono pesquisa@, sem Gmail) ----------
   O servidor envia os relatórios, lê as respostas e responde pela caixa pesquisa@ usando o Microsoft Graph,
   com o aplicativo "Pesquisas Maxipark" registrado no Entra (delegado: só age na caixa que entrou nele; sem segredo).
   Para conectar (ou reconectar): rode conectarOutlook() pelo editor, abra o endereço do registro e entre com
   pesquisa@maxipark.com.br. A Microsoft volta para o próprio servidor (doGet com code/state), que guarda o token
   nas Propriedades do script (MS_REFRESH). Nenhuma senha fica guardada. */
const MS_TENANT = 'e6f85d01-20ad-497d-bcb5-bf47461eba9e';
const MS_CLIENTE = '9612dd11-1818-45d5-a803-9435c30036b9';
const MS_CAIXA = 'pesquisa@maxipark.com.br';
const MS_ESCOPOS = 'offline_access User.Read Mail.ReadWrite Mail.Send';
function msRedirect() { return 'https://script.google.com/macros/s/AKfycbyDp_t-Nbpa2Ikjg2_MB--sEyGJzbR_eV3dMmfdtlk4IA0IBvEKF6hkQVQHQT5I_Gtz/exec'; }   // o próprio servidor (roda sempre como o dono)
function outlookConectado() { return !!PropertiesService.getScriptProperties().getProperty('MS_REFRESH'); }
function conectarOutlook() {
  const verif = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('MS_VERIF', verif);
  const desafio = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, verif, Utilities.Charset.UTF_8)).replace(/=+$/, '');
  const estado = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('MS_ESTADO', estado);
  UrlFetchApp.fetch('https://login.microsoftonline.com/' + MS_TENANT + '/v2.0/.well-known/openid-configuration', { muteHttpExceptions: true });   // garante a permissão de acessar a Microsoft
  const url = 'https://login.microsoftonline.com/' + MS_TENANT + '/oauth2/v2.0/authorize?' + [
    'client_id=' + MS_CLIENTE, 'response_type=code', 'response_mode=query', 'redirect_uri=' + encodeURIComponent(msRedirect()),
    'scope=' + encodeURIComponent(MS_ESCOPOS), 'state=' + estado, 'code_challenge=' + desafio, 'code_challenge_method=S256',
    'login_hint=' + encodeURIComponent(MS_CAIXA), 'prompt=select_account'].join('&');
  Logger.log('Abra este endereço e entre com ' + MS_CAIXA + ':\n' + url);
  return url;
}
function outlookRetorno(req) {
  const p = (req && req.parameter) || {};
  const pagina = function (t, ok) { return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:30px;font-size:18px">' + (ok ? '✅ ' : '❌ ') + t + '</div>'); };
  if (p.error) return pagina('A Microsoft não autorizou: ' + escHtml(p.error_description || p.error), false);
  const props = PropertiesService.getScriptProperties();
  if (!p.state || p.state !== props.getProperty('MS_ESTADO')) return pagina('Pedido de conexão inválido ou vencido. Rode conectarOutlook de novo.', false);
  const r = UrlFetchApp.fetch('https://login.microsoftonline.com/' + MS_TENANT + '/oauth2/v2.0/token', { method: 'post', muteHttpExceptions: true,
    payload: { client_id: MS_CLIENTE, grant_type: 'authorization_code', code: p.code, redirect_uri: msRedirect(), code_verifier: props.getProperty('MS_VERIF') || '', scope: MS_ESCOPOS } });
  const j = JSON.parse(r.getContentText() || '{}');
  if (!j.refresh_token) return pagina('Não consegui conectar: ' + escHtml(j.error_description || j.error || r.getResponseCode()), false);
  const eu = JSON.parse(UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName', { headers: { Authorization: 'Bearer ' + j.access_token }, muteHttpExceptions: true }).getContentText() || '{}');
  const conta = String(eu.mail || eu.userPrincipalName || '').toLowerCase();
  if (conta !== MS_CAIXA) return pagina('Você entrou com ' + escHtml(conta) + '. É preciso entrar com ' + MS_CAIXA + '. Rode conectarOutlook de novo.', false);
  props.setProperty('MS_REFRESH', j.refresh_token); props.setProperty('MS_CONTA', conta); props.deleteProperty('MS_VERIF'); props.deleteProperty('MS_ESTADO');
  CacheService.getScriptCache().put('MS_ACESSO', j.access_token, Math.max(60, (j.expires_in || 3600) - 300));
  return pagina('Caixa ' + MS_CAIXA + ' conectada ao servidor das pesquisas. Pode fechar esta aba.', true);
}
function msAcesso() {
  const c = CacheService.getScriptCache(), a = c.get('MS_ACESSO'); if (a) return a;
  const props = PropertiesService.getScriptProperties(), rt = props.getProperty('MS_REFRESH');
  if (!rt) throw new Error('outlook-desconectado');
  const r = UrlFetchApp.fetch('https://login.microsoftonline.com/' + MS_TENANT + '/oauth2/v2.0/token', { method: 'post', muteHttpExceptions: true,
    payload: { client_id: MS_CLIENTE, grant_type: 'refresh_token', refresh_token: rt, scope: MS_ESCOPOS } });
  const j = JSON.parse(r.getContentText() || '{}');
  if (!j.access_token) throw new Error('outlook-token: ' + (j.error || r.getResponseCode()));
  if (j.refresh_token) props.setProperty('MS_REFRESH', j.refresh_token);
  c.put('MS_ACESSO', j.access_token, Math.max(60, (j.expires_in || 3600) - 300));
  return j.access_token;
}
function graph(metodo, caminho, corpo, extras) {
  const op = { method: metodo, muteHttpExceptions: true, headers: Object.assign({ Authorization: 'Bearer ' + msAcesso() }, extras || {}) };
  if (corpo !== undefined) { op.contentType = 'application/json'; op.payload = JSON.stringify(corpo); }
  const r = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0' + caminho, op), cod = r.getResponseCode(), t = r.getContentText();
  if (cod >= 300) throw new Error('graph ' + cod + ': ' + t.slice(0, 300));
  return t ? JSON.parse(t) : {};
}
function outlookEnviar(para, assunto, html, pdf) {
  graph('post', '/me/sendMail', { saveToSentItems: true, message: {
    subject: assunto, body: { contentType: 'HTML', content: html },
    toRecipients: para.map(function (e) { return { emailAddress: { address: e } }; }),
    attachments: pdf ? [{ '@odata.type': '#microsoft.graph.fileAttachment', name: pdf.getName(), contentType: 'application/pdf', contentBytes: Utilities.base64Encode(pdf.getBytes()) }] : [] } });
}
function outlookMensagens() {   // mensagens da caixa pesquisa@ (todas as pastas) com "Relatório da visita" no assunto
  let url = '/me/messages?$search=' + encodeURIComponent('"Relatório da visita"') + '&$top=100&$select=id,internetMessageId,subject,from,receivedDateTime,body';
  const out = []; let voltas = 0;
  while (url && voltas++ < 10) {
    const j = graph('get', url, undefined, { Prefer: 'outlook.body-content-type="text"' });
    (j.value || []).forEach(function (m) { out.push(m); });
    url = j['@odata.nextLink'] ? j['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : '';
  }
  return out;
}
function coletaRespostasOutlook() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return 0;
  try {
    const rels = {}; lerAba(ABA_RELATORIOS, COLS_REL).forEach(function (r) { rels[r.id] = r; });
    if (!Object.keys(rels).length) return 0;
    const wsA = abaSimples(ABA_RESP_REL, COLS_RESP), vistos = {};
    lerAba(ABA_RESP_REL, COLS_RESP).forEach(function (r) { vistos[r.msg_id] = true; });
    const novas = [];
    outlookMensagens().forEach(function (msg) {
      const m = /\[(REL-\d{6}-[A-Z0-9]{5})\]/.exec(msg.subject || ''); if (!m || !rels[m[1]]) return;
      const de = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '').toLowerCase();
      if (!de || de === MS_CAIXA) return;               // o que a própria caixa enviou não é resposta
      const id = msg.internetMessageId || msg.id; if (vistos[id]) return; vistos[id] = true;
      const rel = rels[m[1]];
      novas.push([id, m[1], rel.unidade, rel.rodada, rel.data, semFormula(de), new Date(msg.receivedDateTime), semFormula(limpaResposta((msg.body && msg.body.content) || ''))]);
    });
    if (novas.length) wsA.getRange(wsA.getLastRow() + 1, 1, novas.length, COLS_RESP.length).setValues(novas);
    return novas.length;
  } finally { lock.releaseLock(); }
}
function outlookResponder(id, html) {   // responde a todos na conversa do relatório, a partir da mensagem mais recente
  const msgs = outlookMensagens().filter(function (m) { return (m.subject || '').indexOf(id) >= 0; })
    .sort(function (a, b) { return String(b.receivedDateTime).localeCompare(String(a.receivedDateTime)); });
  if (!msgs.length) return false;
  const rascunho = graph('post', '/me/messages/' + msgs[0].id + '/createReplyAll', {});
  graph('patch', '/me/messages/' + rascunho.id, { body: { contentType: 'HTML', content: html } });
  graph('post', '/me/messages/' + rascunho.id + '/send', {});
  return true;
}

/* ---------- Acesso dos usuários (17/10/2026): convite por e-mail, primeiro acesso, "esqueci minha senha"
   e entrada pelo servidor. O administrador cadastra só nome, e-mail e nível; a pessoa recebe um link
   (válido por 3 dias) para criar a própria senha. "Esqueci minha senha" manda um link válido por 2 horas.
   O servidor nunca recebe a senha: o aparelho calcula a prova (PBKDF2) e só ela é guardada e conferida. ---------- */
const APP_URL = 'https://maxipark.github.io/pesquisa-equipe-maxipark/';
function lerUsuarios() { const v = abaUsuarios().getRange(1, 1).getValue(); try { return JSON.parse(v || '{}') || {}; } catch (e) { return {}; } }
function salvarUsuarios(d) { abaUsuarios().getRange(1, 1).setValue(JSON.stringify(d)); }
function hexSha(t) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(t), Utilities.Charset.UTF_8).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join(''); }
function contaPorEmail(d, email) { email = String(email || '').trim().toLowerCase(); return ((d && d.contas) || []).filter(function (c) { return String(c.email || '').toLowerCase() === email; })[0]; }
function criaToken(email, tipo, horas) {
  const props = PropertiesService.getScriptProperties(), agora = Date.now(), todas = props.getProperties();
  Object.keys(todas).forEach(function (k) { if (k.indexOf('tok_') === 0) { try { if (JSON.parse(todas[k]).exp < agora) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } } });
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  props.setProperty('tok_' + hexSha(t), JSON.stringify({ email: email, tipo: tipo, exp: agora + horas * 3600000 }));
  return t;
}
function lerToken(t) {
  if (!/^[0-9a-f]{64}$/.test(String(t || ''))) return null;
  const v = PropertiesService.getScriptProperties().getProperty('tok_' + hexSha(t)); if (!v) return null;
  const o = JSON.parse(v); return o.exp >= Date.now() ? o : null;
}
function enviaEmailSimples(para, assunto, html) {
  if (outlookConectado()) outlookEnviar([para], assunto, html, null);
  else MailApp.sendEmail({ to: para, subject: assunto, htmlBody: html, name: 'Pesquisas Maxipark' });
}
function emailLink(nome, texto, botao, url, validade) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1f1a33;max-width:560px;font-size:15px;line-height:1.5">'
    + '<div style="background:#4E3572;color:#fff;border-radius:12px;padding:16px 20px;font-size:18px;font-weight:700">Maxipark · App de pesquisas</div>'
    + '<p>Olá, ' + escHtml(String(nome || '').split(' ')[0]) + '!</p><p>' + texto + '</p>'
    + '<p style="margin:24px 0"><a href="' + url + '" style="background:#6B4A98;color:#fff;text-decoration:none;padding:12px 22px;border-radius:10px;font-weight:700;display:inline-block">' + botao + '</a></p>'
    + '<p style="font-size:13px;color:#6e6a7c">O link vale por ' + validade + ' e só funciona uma vez. Se o botão não abrir, copie este endereço no navegador:<br>' + url + '</p>'
    + '<p style="font-size:13px;color:#6e6a7c">Se você não esperava este e-mail, pode ignorá-lo.</p></div>';
}
function convidarUsuario(b) {   // administrador (com a chave) pede o convite depois de cadastrar
  const c = contaPorEmail(lerUsuarios(), b.email);
  if (!c) return { ok: false, erro: 'usuario-nao-encontrado' };
  if (c.bloqueado) return { ok: false, erro: 'usuario-bloqueado' };
  const t = criaToken(c.email, 'convite', 72);
  enviaEmailSimples(c.email, 'Seu acesso ao app de pesquisas Maxipark — crie sua senha',
    emailLink(c.nome, 'Você foi cadastrado(a) no app de pesquisas da Maxipark. Para concluir o cadastro, toque no botão abaixo e crie a sua senha.', 'Criar minha senha', APP_URL + '#senha=' + t, '3 dias'));
  return { ok: true };
}
function recuperarSenha(b) {   // público: responde sempre "ok", para não revelar quem está cadastrado
  const email = String(b.email || '').trim().toLowerCase();
  if (!email) return { ok: true };
  const cache = CacheService.getScriptCache(), k = 'rec_' + hexSha(email);
  if (cache.get(k)) return { ok: true };             // no máximo um e-mail a cada 5 minutos por pessoa
  cache.put(k, '1', 300);
  const c = contaPorEmail(lerUsuarios(), email);
  if (!c || c.bloqueado) return { ok: true };
  const t = criaToken(c.email, 'senha', 2);
  enviaEmailSimples(c.email, 'App de pesquisas Maxipark — criar uma nova senha',
    emailLink(c.nome, 'Recebemos um pedido para criar uma nova senha no app de pesquisas da Maxipark. Toque no botão abaixo para escolher a nova senha.', 'Criar nova senha', APP_URL + '#senha=' + t, '2 horas'));
  return { ok: true };
}
function infoToken(p) {
  const o = lerToken(p.token); if (!o) return { ok: false, erro: 'link-invalido' };
  const c = contaPorEmail(lerUsuarios(), o.email); if (!c || c.bloqueado) return { ok: false, erro: 'link-invalido' };
  return { ok: true, nome: c.nome, email: c.email, tipo: o.tipo };
}
function definirSenha(b) {   // público, mas só com um link válido
  const o = lerToken(b.token); if (!o) return { ok: false, erro: 'link-invalido' };
  if (!/^[A-Za-z0-9+/=]{16,64}$/.test(String(b.salt || '')) || !/^[A-Za-z0-9+/=]{40,64}$/.test(String(b.hash || '')) || !(Number(b.iter) >= 100000 && Number(b.iter) <= 1000000)) return { ok: false, erro: 'dados-invalidos' };
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const d = lerUsuarios(), c = contaPorEmail(d, o.email);
    if (!c || c.bloqueado) return { ok: false, erro: 'link-invalido' };
    c.salt = String(b.salt); c.iter = Number(b.iter); c.hash = String(b.hash); delete c.pendente;
    d.versao = (d.versao || 0) + 1; salvarUsuarios(d);
    PropertiesService.getScriptProperties().deleteProperty('tok_' + hexSha(b.token));
  } finally { lock.releaseLock(); }
  return { ok: true, email: o.email };
}
function salDoUsuario(p) {   // público: o sal para o aparelho calcular a prova da senha (sal falso para quem não existe)
  const c = contaPorEmail(lerUsuarios(), p.email);
  if (c && c.salt && c.hash) return { ok: true, salt: c.salt, iter: c.iter };
  return { ok: true, salt: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'maxipark-sal-' + String(p.email || '').toLowerCase())).slice(0, 24), iter: 150000 };
}
function entrarPeloServidor(b) {   // público: confere a prova e devolve o perfil, para entrar em qualquer aparelho
  const email = String(b.email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), k = 'ent_' + hexSha(email), n = Number(cache.get(k) || 0);
  if (n >= 10) return { ok: false, erro: 'muitas-tentativas' };   // 10 erros seguidos: espera 15 minutos
  const d = lerUsuarios(), c = contaPorEmail(d, email);
  if (!c || !c.hash || String(b.hash || '') !== c.hash) { cache.put(k, String(n + 1), 900); return { ok: false, erro: 'nao-confere' }; }
  if (c.bloqueado) return { ok: false, erro: 'bloqueado' };
  cache.remove(k);
  return { ok: true, conta: c, sairTodos: d.sairTodos || 0, chave: CHAVE };   // a chave liga o aparelho ao servidor sozinho (17/10/2026)
}

/* ---------- PIN de cada usuário que aplica pesquisa (17/10/2026) ----------
   Quem aplica pesquisa escolhe o próprio PIN (4 a 6 números) na página inicial. O servidor guarda só o sal e a
   impressão (SHA-256) do PIN, e só aceita a troca de quem prova que sabe a senha. O PIN vai no crachá (#a=...)
   para o tablet conferir sem internet. Senha e PIN nunca são trocados pelo cadastro do administrador. */
function definirPin(b) {
  if (!/^[0-9a-f]{32}$/.test(String(b.pinSalt || '')) || !/^[0-9a-f]{64}$/.test(String(b.pinHash || ''))) return { ok: false, erro: 'dados-invalidos' };
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const d = lerUsuarios(), c = contaPorEmail(d, b.email);
    if (!c || !c.hash || String(b.prova || '') !== c.hash) return { ok: false, erro: 'nao-confere' };
    if (c.bloqueado) return { ok: false, erro: 'bloqueado' };
    c.pinSalt = String(b.pinSalt); c.pinHash = String(b.pinHash);
    d.versao = (d.versao || 0) + 1; salvarUsuarios(d);
    return { ok: true, versao: d.versao };
  } finally { lock.releaseLock(); }
}
function protegeCredenciais(novo) {   // no "salvar usuários" do administrador, senha e PIN sempre vêm do servidor
  const atual = lerUsuarios();
  ((novo && novo.contas) || []).forEach(function (c) {
    const a = contaPorEmail(atual, c.email) || (c.id && ((atual && atual.contas) || []).filter(function (x) { return x.id === c.id; })[0]);   // casa pelo e-mail ou, se o e-mail foi corrigido, pelo id (21/10/2026)
    if (!a) return;
    ['salt', 'iter', 'hash', 'pinSalt', 'pinHash'].forEach(function (k) { if (a[k] !== undefined) c[k] = a[k]; else delete c[k]; });
    if (a.hash) delete c.pendente;
  });
  return novo;
}

/* Conselho desativado (15/10/2026): rode UMA vez pelo editor para esconder a aba da planilha (nada é apagado). */
function esconderAbaConselho() { const a = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conselho'); if (a) a.hideSheet(); Logger.log(a ? 'Aba Conselho escondida' : 'Aba Conselho não existe'); }

/* ---------------- revisão das perguntas (pedido do Raphael em 21/10/2026) ----------------
   Cada pessoa com acesso marca "manter", "mudar o texto" ou "tirar" para cada pergunta das pesquisas e deixa
   uma observação. Uma linha por pergunta e por pessoa, na aba "RevisaoPerguntas"; marcar de novo substitui a linha. */
const ABA_REVISAO = 'RevisaoPerguntas';
const COLS_REVISAO = ['chave_item', 'pesquisa', 'pergunta', 'email', 'nome', 'decisao', 'observacao', 'atualizado_em'];
function revisaoSalvar(it) {
  const ch = String(it.chave_item || '').trim(), email = String(it.email || '').trim().toLowerCase();
  if (!ch || !email) return { ok: false, erro: 'dados' };
  const dec = ['manter', 'mudar', 'tirar', ''].indexOf(String(it.decisao || '')) >= 0 ? String(it.decisao || '') : '';
  const obs = String(it.observacao || '').slice(0, 4000);
  const linha = [ch, String(it.pesquisa || ''), semFormula(String(it.pergunta || '').slice(0, 500)), email, semFormula(String(it.nome || '').slice(0, 120)), dec, semFormula(obs), new Date()];
  const ws = abaSimples(ABA_REVISAO, COLS_REVISAO);
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const last = ws.getLastRow(); let achou = 0;
    if (last >= 2) {
      const vals = ws.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === ch && String(vals[i][3]).toLowerCase() === email) { achou = i + 2; break; }
    }
    if (achou) ws.getRange(achou, 1, 1, COLS_REVISAO.length).setValues([linha]); else ws.appendRow(linha);
  } finally { lock.releaseLock(); }
  return { ok: true };
}
