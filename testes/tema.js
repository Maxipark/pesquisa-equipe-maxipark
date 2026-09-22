/* Bateria do botao claro/escuro (a lua e o sol), nas cinco telas.
   Roda num Chromium de verdade, contra a copia local do site.
   Uso:  node testes/tema.js
   Sai com codigo 1 se qualquer conferencia falhar, para travar a publicacao. */
const { chromium } = require('playwright');
const S = require('./servidor');

const PORTA = 8099;
let passou = 0, falhou = 0;
const ok  = m => { passou++; console.log('   ✓ ' + m); };
const mal = m => { falhou++; console.log('   ✗ ' + m); };

/* contraste WCAG entre duas cores "rgb(...)" */
function lum(c) {
  const [r, g, b] = c.match(/\d+(\.\d+)?/g).slice(0, 3).map(Number).map(v => {
    v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
const contraste = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
const cruza = (a, b) => a && b && a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;

(async () => {
  const servidor = await S.liga(PORTA);
  const nav = await chromium.launch({ executablePath: S.chromium() });
  const url = r => `http://localhost:${PORTA}${r}`;

  /* ═══ 1. o botao, tela por tela ═══ */
  for (const [nome, rota] of S.TELAS) {
    console.log('\n══ ' + nome + '  ' + rota);
    const ctx = await nav.newContext({ viewport: { width: 900, height: 780 } });
    const pag = await ctx.newPage();
    const erros = [];
    pag.on('pageerror', e => erros.push('JS: ' + e.message));
    pag.on('console', m => { if (m.type() === 'error') erros.push('console: ' + m.text().slice(0, 160)); });
    pag.on('requestfailed', r => erros.push('rede: ' + r.url().slice(-60)));

    await pag.goto(url(rota), { waitUntil: 'networkidle' });
    await pag.waitForTimeout(350);

    const b = pag.locator('#btn-tema');
    if (await b.count() !== 1) { mal('o botao nao foi criado'); await ctx.close(); continue; }
    ok('o botao existe');
    if (await b.isVisible()) ok('o botao aparece na tela'); else mal('o botao esta escondido');

    const cx = await b.boundingBox(), vp = pag.viewportSize();
    if (cx && cx.x + cx.width > vp.width - 90 && cx.y + cx.height > vp.height - 130)
      ok(`canto de baixo a direita (x=${Math.round(cx.x)}, y=${Math.round(cx.y)})`);
    else mal('fora do canto: ' + JSON.stringify(cx));
    if (cx && cx.width >= 44 && cx.height >= 44) ok(`alvo de toque ${Math.round(cx.width)}x${Math.round(cx.height)}px`);
    else mal('botao pequeno demais para o dedo');

    const tema0 = await pag.evaluate(() => document.documentElement.getAttribute('data-theme'));
    const icone0 = (await b.textContent()).trim();
    if (tema0 === 'light') ok('comeca no claro'); else mal('comecou em ' + tema0);
    if (icone0 === '🌙') ok('mostra a lua (toque para escurecer)'); else mal('icone inicial: ' + icone0);

    const fundoClaro = await pag.evaluate(() => getComputedStyle(document.body).backgroundColor);
    const textoClaro = await pag.evaluate(() => getComputedStyle(document.body).color);

    await b.click(); await pag.waitForTimeout(220);
    const tema1 = await pag.evaluate(() => document.documentElement.getAttribute('data-theme'));
    const icone1 = (await b.textContent()).trim();
    if (tema1 === 'dark') ok('o toque escurece'); else mal('depois do toque: ' + tema1);
    if (icone1 === '☀') ok('mostra o sol (toque para clarear)'); else mal('icone: ' + icone1);

    const fundoEscuro = await pag.evaluate(() => getComputedStyle(document.body).backgroundColor);
    const textoEscuro = await pag.evaluate(() => getComputedStyle(document.body).color);
    if (lum(fundoEscuro) < lum(fundoClaro)) ok(`o fundo escureceu de verdade (${fundoClaro} -> ${fundoEscuro})`);
    else mal('o fundo nao mudou: ' + fundoEscuro);

    const cC = contraste(textoClaro, fundoClaro), cE = contraste(textoEscuro, fundoEscuro);
    if (cC >= 4.5) ok(`contraste do texto no claro ${cC.toFixed(1)}:1`); else mal(`contraste no claro fraco: ${cC.toFixed(1)}:1`);
    if (cE >= 4.5) ok(`contraste do texto no escuro ${cE.toFixed(1)}:1`); else mal(`contraste no escuro fraco: ${cE.toFixed(1)}:1`);

    /* o icone do botao nao pode herdar a cor dos outros botoes (ver LEIA-ME) */
    const cb = await b.evaluate(el => { const s = getComputedStyle(el); return [s.color, s.backgroundColor]; });
    const cBot = contraste(cb[0], cb[1]);
    if (cBot >= 4.5) ok(`o icone do botao se le no escuro (${cBot.toFixed(1)}:1)`);
    else mal(`icone quase invisivel: ${cBot.toFixed(1)}:1 (${cb[0]} sobre ${cb[1]})`);

    const meta = await pag.evaluate(() => { const m = document.querySelector('meta[name="theme-color"]'); return m && m.content; });
    if (meta === '#17111E') ok('a barra do navegador escureceu'); else mal('theme-color: ' + meta);

    const guardado = await pag.evaluate(() => localStorage.getItem('maxipark_tema'));
    if (guardado === 'escuro') ok('a escolha ficou guardada no aparelho'); else mal('guardado: ' + guardado);

    await pag.reload({ waitUntil: 'networkidle' }); await pag.waitForTimeout(300);
    const tema2 = await pag.evaluate(() => document.documentElement.getAttribute('data-theme'));
    if (tema2 === 'dark') ok('continua escuro depois de recarregar'); else mal('ao recarregar virou ' + tema2);

    await pag.locator('#btn-tema').click(); await pag.waitForTimeout(200);
    const tema3 = await pag.evaluate(() => document.documentElement.getAttribute('data-theme'));
    if (tema3 === 'light') ok('o segundo toque clareia'); else mal('nao voltou ao claro: ' + tema3);

    if (!erros.length) ok('nenhum erro de JS, de console ou de rede');
    else erros.slice(0, 6).forEach(mal);

    await ctx.close();
  }

  /* ═══ 2. a escolha vale para as cinco telas ═══ */
  console.log('\n══ a escolha vale para as cinco telas');
  {
    const ctx = await nav.newContext({ viewport: { width: 900, height: 780 } });
    const p1 = await ctx.newPage();
    await p1.goto(url('/'), { waitUntil: 'networkidle' });
    await p1.locator('#btn-tema').click(); await p1.waitForTimeout(200);
    for (const [nome, rota] of S.TELAS.slice(1)) {
      const p = await ctx.newPage();
      await p.goto(url(rota), { waitUntil: 'networkidle' }); await p.waitForTimeout(250);
      const t = await p.evaluate(() => document.documentElement.getAttribute('data-theme'));
      if (t === 'dark') ok(`${nome} abriu ja no escuro`); else mal(`${nome} abriu em ${t}`);
      await p.close();
    }
    await ctx.close();
  }

  /* ═══ 3. aparelho no escuro e ninguem escolheu ═══ */
  console.log('\n══ aparelho no modo escuro, sem escolha do usuario');
  {
    const ctx = await nav.newContext({ colorScheme: 'dark', viewport: { width: 900, height: 780 } });
    for (const [nome, rota] of S.TELAS) {
      const p = await ctx.newPage();
      await p.goto(url(rota), { waitUntil: 'networkidle' }); await p.waitForTimeout(250);
      const t = await p.evaluate(() => document.documentElement.getAttribute('data-theme'));
      if (t === 'dark') ok(`${nome} segue o aparelho`); else mal(`${nome} ficou em ${t}`);
      await p.close();
    }
    await ctx.close();
  }

  /* ═══ 4. o botao nao cobre o que ja existia no canto ═══ */
  console.log('\n══ o botao nao cobre o rodape de quiosque nem o engrenagem Usuario');
  const VIZINHOS = [
    ['Clientes', '/cliente/',  '<div class="rodape-kiosk"><span>Maxipark</span><button>Area do usuario</button></div>', '.rodape-kiosk button'],
    ['Conselho', '/conselho/', '<div class="rodape-kiosk"><span>Maxipark</span><button>Area do usuario</button></div>', '.rodape-kiosk button'],
    ['Equipe',   '/equipe/',   '<button class="corner">Usuario</button>', '.corner'],
  ];
  for (const [nome, rota, html, sel] of VIZINHOS) {
    for (const w of [360, 420, 900]) {
      const ctx = await nav.newContext({ viewport: { width: w, height: 740 } });
      const p = await ctx.newPage();
      await p.goto(url(rota), { waitUntil: 'networkidle' }); await p.waitForTimeout(300);
      await p.evaluate(h => document.body.insertAdjacentHTML('beforeend', h), html);
      await p.waitForTimeout(150);
      const a = await p.locator('#btn-tema').boundingBox();
      const b2 = await p.locator(sel).first().boundingBox();
      if (cruza(a, b2)) mal(`${nome} @${w}px: o botao cobre ${sel}`);
      else ok(`${nome} @${w}px: ${sel} livre`);
      await ctx.close();
    }
  }

  await nav.close(); servidor.close();
  console.log(`\n${'═'.repeat(58)}\n${passou} conferencias passaram, ${falhou} falharam.`);
  process.exit(falhou ? 1 : 0);
})();
