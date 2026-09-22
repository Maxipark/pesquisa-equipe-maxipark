/* Servidor local só para os testes: serve a pasta do site numa porta qualquer.
   Não é o servidor das respostas — aquele é o Apps Script, e o container não
   alcança o script.google.com. */
const http = require('http'), fs = require('fs'), path = require('path');

const TIPOS = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml',
  '.webmanifest': 'application/manifest+json', '.json': 'application/json', '.txt': 'text/plain; charset=utf-8'
};

/* raiz do site = a pasta acima desta */
const RAIZ = path.resolve(__dirname, '..');

function liga(porta) {
  const s = http.createServer((req, res) => {
    let p = decodeURIComponent(req.url.split('?')[0]);
    if (p.endsWith('/')) p += 'index.html';
    const abs = path.join(RAIZ, p);
    if (!abs.startsWith(RAIZ) || !fs.existsSync(abs) || fs.statSync(abs).isDirectory()) {
      res.writeHead(404); return res.end('nao achei');
    }
    res.writeHead(200, { 'content-type': TIPOS[path.extname(abs)] || 'application/octet-stream' });
    res.end(fs.readFileSync(abs));
  });
  return new Promise(r => s.listen(porta, () => r(s)));
}

/* onde está o Chromium do Playwright neste computador */
function chromium() {
  const cands = [];
  const base = process.env.PLAYWRIGHT_BROWSERS_PATH || '/opt/pw-browsers';
  try {
    for (const d of fs.readdirSync(base))
      if (/^chromium-/.test(d)) cands.push(path.join(base, d, 'chrome-linux', 'chrome'));
  } catch (e) {}
  for (const c of cands) if (fs.existsSync(c)) return c;
  return undefined;   /* deixa o Playwright achar sozinho */
}

const TELAS = [
  ['Pagina inicial', '/'],
  ['Clientes',       '/cliente/'],
  ['Equipe',         '/equipe/'],
  ['Conselho',       '/conselho/'],
  ['Painel',         '/painel/'],
];

module.exports = { liga, chromium, RAIZ, TELAS };
