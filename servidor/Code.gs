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
    if (p.op === 'ping') return json({ ok: true, servidor: 'apps-script', versao: 4, abas: Object.keys(ABAS), usuarios: true, historico: true });
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
