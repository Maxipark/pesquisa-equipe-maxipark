# Testes do app de pesquisas

Esta pasta **não vai para o ar**. Ela só serve para conferir o app antes de
publicar. Os tablets nunca baixam nada daqui.

Ela existe porque a regra 2 do projeto diz que *toda mudança é refeita por
inteiro: republicar os arquivos alterados, rodar a bateria completa e conferir
byte a byte contra o GitHub*. Sem os testes guardados aqui, essa regra vira uma
promessa que ninguém consegue cumprir.

---

## Como rodar

Precisa de Node 18 ou mais novo.

```bash
cd testes
npm install        # só na primeira vez
npm run teste      # roda o auditor e a bateria do tema
```

Se qualquer conferência falhar, o comando sai com erro — é o sinal para **não
publicar**.

Comandos separados, quando quiser só uma parte:

| Comando | O que faz |
|---|---|
| `npm run audita` | procura erro de sintaxe e nome indefinido no JavaScript das seis páginas |
| `npm run tema`   | a bateria do botão claro/escuro nas cinco telas (103 conferências) |
| `npm run fotos`  | grava as cinco telas no claro e no escuro em `testes/fotos/`, para olhar |

---

## O que cada arquivo é

- **`servidor.js`** — sobe um servidor local servindo a pasta do site, e acha o
  Chromium do Playwright. Os outros testes usam este.
- **`audita.js`** — lê cada `<script>` inline das páginas com o acorn e procura
  erro de sintaxe e nome indefinido. Trata **todos os `<script>` de uma página
  como um escopo só**, porque é assim que o navegador trata: o que o bloco 3
  declara, o bloco 4 enxerga. Um auditor que olha bloco a bloco acusa dezenas
  de erros que não existem.
- **`tema.js`** — a bateria do botão claro/escuro. Confere, nas cinco telas:
  o botão existe, aparece, fica no canto de baixo à direita e tem alvo de toque
  de pelo menos 44px; começa no claro com a lua; o toque escurece e mostra o
  sol; o fundo escurece de verdade; o contraste do texto e do ícone do botão
  passa de 4,5:1 nos dois temas; a barra do navegador acompanha; a escolha fica
  guardada no aparelho e sobrevive a recarregar; o segundo toque clareia; a
  escolha atravessa as cinco telas; o aparelho no escuro é respeitado enquanto
  ninguém escolher; e o botão não cobre o rodapé de quiosque nem o botão
  ⚙ Usuário, em 360, 420 e 900px de largura.
- **`fotos.js`** — tira as fotos das telas nos dois temas. Não confere nada
  sozinho: serve para você olhar.

---

## Duas coisas que economizam tempo

**O Chromium.** O `servidor.js` procura o Chromium já instalado em
`/opt/pw-browsers` e passa o caminho para o Playwright. Sem isso, o Playwright
procura uma versão que pode não estar instalada e reclama pedindo
`npx playwright install`.

**O contraste do ícone do botão.** Na página inicial existe a regra
`:root[data-theme="dark"] button{color:#1B1226}`, que é **mais específica** do
que `.botao-tema{color:…}`. Da primeira vez, o ícone do botão ficou quase
invisível no escuro por causa disso. A correção foi dar ao botão uma regra de
peso igual e posterior:

```css
:root[data-theme="dark"] .botao-tema,
:root[data-theme="light"] .botao-tema{background:var(--card);color:var(--texto)}
```

O `tema.js` mede esse contraste justamente para que ninguém repita o erro ao
mexer nas cores.

---

## O que estes testes NÃO cobrem

Vale saber, para não confiar demais:

- **O servidor de verdade.** O Apps Script não é alcançável de dentro do
  container, então nada aqui testa envio de resposta, cadastro de usuário ou
  histórico. Isso continua sendo conferido à mão, pelo Diagnóstico do servidor
  dentro de Configurações (o PIN está no bilhete de manutenção).
- **O site publicado.** Os testes rodam contra a cópia local. Depois de
  publicar, a conferência final é abrir o endereço de verdade no navegador —
  de preferência num tablet, que é onde o app vive.
- **As telas de dentro das pesquisas.** A bateria entra na primeira tela de
  cada app. Percorrer uma pesquisa inteira, exportar Excel e PDF: ainda é à mão.
