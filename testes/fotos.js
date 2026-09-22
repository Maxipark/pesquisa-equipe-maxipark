/* Fotografa as cinco telas no claro e no escuro, para olhar com os olhos.
   Uso:  node testes/fotos.js         (grava em testes/fotos/, que nao vai para o repositorio) */
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const S = require('./servidor');

const PORTA = 8098;
const SAI = path.join(__dirname, 'fotos');

(async () => {
  fs.mkdirSync(SAI, { recursive: true });
  const servidor = await S.liga(PORTA);
  const nav = await chromium.launch({ executablePath: S.chromium() });

  for (const tema of ['claro', 'escuro']) {
    const ctx = await nav.newContext({ viewport: { width: 1000, height: 820 } });
    await ctx.addInitScript(t => { try { localStorage.setItem('maxipark_tema', t); } catch (e) {} }, tema);
    for (const [nome, rota] of S.TELAS) {
      const p = await ctx.newPage();
      await p.goto(`http://localhost:${PORTA}${rota}`, { waitUntil: 'networkidle' });
      await p.waitForTimeout(500);
      const arq = path.join(SAI, nome.toLowerCase().replace(/[^a-z]+/g, '-') + '-' + tema + '.png');
      await p.screenshot({ path: arq });
      console.log('  ' + path.basename(arq));
      await p.close();
    }
    await ctx.close();
  }
  await nav.close(); servidor.close();
  console.log('\nFotos em ' + SAI);
})();
