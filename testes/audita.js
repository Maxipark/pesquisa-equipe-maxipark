/* Auditor de JavaScript das paginas do app.
   Procura erro de sintaxe e nome indefinido em cada <script> inline.
   Trata TODOS os <script> de uma mesma pagina como um escopo so, porque e
   assim que o navegador trata: o que o bloco 3 declara, o bloco 4 enxerga.
   Uso:  node testes/audita.js            (audita as seis paginas)
         node testes/audita.js a.html     (audita so o que voce passar)      */
const fs=require('fs'), acorn=require('acorn'), walk=require('acorn-walk');
const G=new Set(['window','document','navigator','location','localStorage','sessionStorage','console','JSON','Math','Date','Number','String','Boolean','Array','Object','Promise','Set','Map','WeakMap','WeakSet','RegExp','Error','TypeError','RangeError','fetch','Request','Response','Headers','setTimeout','clearTimeout','setInterval','clearInterval','requestAnimationFrame','cancelAnimationFrame','alert','confirm','prompt','encodeURIComponent','decodeURIComponent','encodeURI','decodeURI','escape','unescape','isNaN','isFinite','parseInt','parseFloat','URL','URLSearchParams','Intl','Blob','File','FileReader','FormData','TextEncoder','TextDecoder','crypto','history','screen','matchMedia','self','globalThis','undefined','NaN','Infinity','structuredClone','AbortController','Uint8Array','Uint16Array','Int32Array','Float64Array','ArrayBuffer','DataView','Symbol','Function','atob','btoa','XMLHttpRequest','CustomEvent','Event','MouseEvent','Image','caches','performance','getComputedStyle','scrollTo','scrollBy','open','close','print','addEventListener','removeEventListener','dispatchEvent','innerWidth','innerHeight','devicePixelRatio','Node','Element','HTMLElement','DocumentFragment','MutationObserver','IntersectionObserver','ResizeObserver','queueMicrotask','reportError','Proxy','Reflect','BigInt','Notification','indexedDB','DOMParser','XMLSerializer','getSelection','module','require','exports','process','eval','arguments','XLSX','jspdf','jsPDF','html2canvas','Chart','QRCode','saveAs','Papa']);
function blocos(html){const out=[];const re=/<script\b([^>]*)>([\s\S]*?)<\/script>/gi;let m;while((m=re.exec(html))){const a=m[1]||'';if(/\bsrc\s*=/.test(a))continue;if(/type\s*=\s*["'](?!text\/javascript|application\/javascript|module)/i.test(a))continue;out.push({code:m[2],linha:html.slice(0,m.index).split('\n').length});}return out;}
function declara(n,add){const rec=q=>{if(!q)return;if(q.type==='Identifier')add(q.name);else if(q.type==='ObjectPattern')q.properties.forEach(p=>rec(p.value||p.argument));else if(q.type==='ArrayPattern')q.elements.forEach(rec);else if(q.type==='AssignmentPattern')rec(q.left);else if(q.type==='RestElement')rec(q.argument);};rec(n);}
let total=0;
const path=require('path');
const RAIZ=path.resolve(__dirname,'..');
const PADRAO=['index.html','cliente/index.html','equipe/index.html','conselho/index.html','painel/index.html','cliente/qr.html']
  .map(f=>path.join(RAIZ,f));
const ALVOS = process.argv.length>2 ? process.argv.slice(2) : PADRAO;
for(const f of ALVOS){
  const html=fs.readFileSync(f,'utf8'); const bs=blocos(html); const asts=[]; const decl=new Set(); let falhou=false;
  bs.forEach((b,i)=>{ let ast;
    try{ast=acorn.parse(b.code,{ecmaVersion:2023,sourceType:'script',allowReturnOutsideFunction:true});}
    catch(e){console.log(`SINTAXE  ${f}  bloco#${i+1} (linha ~${b.linha}): ${e.message}`);total++;falhou=true;return;}
    asts.push({ast,i});
    walk.full(ast,n=>{
      if(n.type==='VariableDeclarator')declara(n.id,x=>decl.add(x));
      if((n.type==='FunctionDeclaration'||n.type==='ClassDeclaration')&&n.id)decl.add(n.id.name);
      if(n.type==='FunctionExpression'||n.type==='ArrowFunctionExpression'||n.type==='FunctionDeclaration')n.params.forEach(p=>declara(p,x=>decl.add(x)));
      if(n.type==='CatchClause'&&n.param)declara(n.param,x=>decl.add(x));
      if(n.type==='ObjectPattern')n.properties.forEach(p=>declara(p.value||p.argument,x=>decl.add(x)));
      if(n.type==='ArrayPattern')n.elements.forEach(e=>declara(e,x=>decl.add(x)));
      if(n.type==='AssignmentExpression'&&n.left.type==='Identifier')decl.add(n.left.name);
    });
  });
  if(falhou)continue;
  const faltando=new Map();
  for(const{ast,i}of asts) walk.ancestor(ast,{Identifier(n,st,anc){
    const p=anc[anc.length-2]; if(!p)return;
    if(p.type==='MemberExpression'&&p.property===n&&!p.computed)return;
    if((p.type==='Property'||p.type==='MethodDefinition'||p.type==='PropertyDefinition')&&p.key===n&&!p.computed)return;
    if(p.type==='LabeledStatement'||p.type==='BreakStatement'||p.type==='ContinueStatement')return;
    if(decl.has(n.name)||G.has(n.name))return;
    const k=`${n.name}|bloco#${i+1}`; faltando.set(k,(faltando.get(k)||0)+1);
  }});
  if(faltando.size){for(const[k,q]of faltando){const[nome,bl]=k.split('|');console.log(`INDEFINIDO  ${f}  ${bl}: ${nome} (${q}x)`);total++;}}
  else console.log(`ok  ${f}  (${bs.length} blocos inline)`);
}
console.log(total?`\n>>> ${total} PROBLEMA(S)`:'\n>>> OK — nenhum erro de sintaxe, nenhum nome indefinido');
