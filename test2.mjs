/* Стресс-проверка: имитируем ЖИВЫЕ клики пользователя, а не вызовы render().
   Прошлые сборки падали именно на клике по фильтру, поэтому проверяем
   через реальные click(), включая случайные последовательности. */
import { JSDOM } from 'jsdom';
import fs from 'fs';
let pass=0,fail=0;
const ok=(c,m)=>{c?(pass++,console.log('  ok  '+m)):(fail++,console.log('  FAIL '+m));};
const drawn=[];
const html=fs.readFileSync('index.html','utf8');
const dom=new JSDOM(html,{runScripts:'outside-only',pretendToBeVisual:true});
const w=dom.window;
w.HTMLCanvasElement.prototype.getContext=function(){return{canvas:this,save(){},restore(){},clearRect(){},
 fillRect(){},strokeRect(){},beginPath(){},closePath(){},moveTo(){},lineTo(){},arc(){},fill(){},stroke(){},
 clip(){},rect(){},translate(){},rotate(){},scale(){},setTransform(){},setLineDash(){},quadraticCurveTo(){},
 bezierCurveTo(){},createLinearGradient(){return{addColorStop(){}}},createRadialGradient(){return{addColorStop(){}}},
 measureText(t){return{width:String(t).length*6,actualBoundingBoxAscent:6,actualBoundingBoxDescent:2}},
 fillText(){},strokeText(){},drawImage(){},putImageData(){},getImageData(){return{data:new Uint8ClampedArray(4)}},
 isPointInPath(){return false},arcTo(){},ellipse(){},resetTransform(){},transform(){},createPattern(){return null},
 roundRect(){}};};
w.ResizeObserver=class{observe(){}unobserve(){}disconnect(){}};
w.requestAnimationFrame=()=>1; w.cancelAnimationFrame=()=>{};
w.matchMedia=()=>({matches:false,addEventListener(){},removeEventListener(){}});
w.scrollTo=()=>{}; w.URL.createObjectURL=()=>'blob:x'; w.URL.revokeObjectURL=()=>{};
const scripts=[...w.document.querySelectorAll('script:not([type="application/json"])')];
w.eval(scripts[0].textContent);
const R=w.Chart;
function Spy(cv,cfg){ drawn.push({id:cv.id,
  sets:(cfg.data.datasets||[]).map(d=>({nonNull:(d.data||[]).filter(v=>v!=null).length}))});
  return new R(cv,cfg); }
Spy.prototype=R.prototype; Object.setPrototypeOf(Spy,R); w.Chart=Spy;
const errs=[];
w.addEventListener('error',e=>errs.push(e.message));
for(const s of scripts.slice(1)) w.eval(s.textContent);
w.dispatchEvent(new w.Event('DOMContentLoaded'));
const $=(id)=>w.document.getElementById(id);
const els=(s)=>[...w.document.querySelectorAll(s)];
const D=w.__D,S=w.__S,API=w.__API,C=API.C;

/** Экран считается живым, если ЛИБО есть график с данными,
    ЛИБО честно написано «нет данных». Молча пустой экран = провал. */
function screenAlive(view){
  const root=$('v-'+view);
  if(!root) return {alive:false,why:'нет секции'};
  // живые графики этой вкладки = экземпляры Chart, чьё полотно внутри неё
  const live=Object.keys(API.CH).filter(id=>{const c=$(id);return c&&root.contains(c);}).length;
  const painted=live||drawn.filter(d=>d.sets.some(s=>s.nonNull>0)).length;
  const said=/нет данных|нужно минимум|плана в таблице нет/.test(root.textContent);
  const rows=C.filterRows(D,S).length;
  // Методика не зависит от среза — она описывает формулы, а не данные.
  // Методика и «Гугл-таблица» — статичные экраны, от среза не зависят
  if(view==='method'||view==='gs') return {alive:root.textContent.trim().length>200,why:'статичный экран — текст'};
  if(rows===0) return {alive:said,why:'пустой срез, сообщение='+said};
  if(['heat','table','pf','share'].includes(view)){
    const has=root.querySelectorAll('.hc,tbody tr,.card').length>0;
    return {alive:has||said,why:'элементов '+root.querySelectorAll('.hc,tbody tr').length};
  }
  return {alive:painted>0||said,why:'графиков '+painted+' строк '+rows};
}

console.log('=== A. Живые клики: вкладки ===');
let bad=0;
for(const t of els('.tab')){
  drawn.length=0; t.click();
  const v=t.dataset.v, r=screenAlive(v);
  if(!r.alive){ console.log('    FAIL вкладка '+v+': '+r.why); bad++; }
}
ok(bad===0,'все 9 вкладок живы после клика');

console.log('\n=== B. Живые клики: КАЖДЫЙ чип направления на КАЖДОЙ вкладке ===');
bad=0; let runs=0;
for(const t of els('.tab')){
  t.click();
  const v=t.dataset.v;
  for(const ch of els('#fDir .chip')){
    drawn.length=0; ch.click();               // включить
    runs++; let r=screenAlive(v);
    if(!r.alive){ console.log('    FAIL '+v+' + '+ch.dataset.v+': '+r.why); bad++; }
    drawn.length=0; ch.click();               // выключить
    runs++; r=screenAlive(v);
    if(!r.alive){ console.log('    FAIL '+v+' - '+ch.dataset.v+': '+r.why); bad++; }
  }
}
ok(bad===0,`${runs} кликов по фильтрам на всех вкладках — ни одного мёртвого экрана`);

console.log('\n=== C. Живые клики: разрез недели/месяцы на каждой вкладке ===');
bad=0;
for(const t of els('.tab')){
  t.click(); const v=t.dataset.v;
  for(const l of ['weekly','monthly']){
    drawn.length=0;
    w.document.querySelector(`#segLevel button[data-l="${l}"]`).click();
    const r=screenAlive(v);
    if(!r.alive){ console.log('    FAIL '+v+' / '+l+': '+r.why); bad++; }
  }
}
ok(bad===0,'смена разреза на всех вкладках безопасна');

console.log('\n=== D. Случайные последовательности (200 шагов) ===');
let seed=42; const rnd=()=>(seed=(seed*1103515245+12345)&0x7fffffff)/0x7fffffff;
const acts=[];
els('.tab').forEach(t=>acts.push(['tab',t]));
els('#fDir .chip').forEach(c=>acts.push(['dir',c]));
els('#fSrc .chip').slice(0,12).forEach(c=>acts.push(['src',c]));
els('#segLevel button').forEach(b=>acts.push(['lvl',b]));
els('#heatMetric .chip').forEach(c=>acts.push(['hm',c]));
els('#heatDim .chip').forEach(c=>acts.push(['hd',c]));
els('#fcMethod .chip').forEach(c=>acts.push(['fc',c]));
bad=0;
for(let i=0;i<200;i++){
  const [,el]=acts[Math.floor(rnd()*acts.length)];
  drawn.length=0;
  try{ el.click(); }catch(e){ console.log('    FAIL исключение на шаге '+i+': '+e.message); bad++; continue; }
  const r=screenAlive(S.view);
  if(!r.alive){ console.log('    FAIL шаг '+i+' вкладка '+S.view+': '+r.why); bad++; if(bad>4) break; }
}
ok(bad===0,'200 случайных кликов: ни одного мёртвого экрана и ни одного исключения');
ok(errs.length===0,'необработанных ошибок в окне: '+errs.length+(errs.length?' → '+errs[0]:''));

console.log('\n=== E. Календарь при живых кликах ===');
w.document.querySelector('#segLevel button[data-l="weekly"]').click();
els('.tab')[1].click();
const cal=$('cal');
for(let i=0;i<5;i++){
  $('calBtn').click();
  if(cal.classList.contains('hide')){ ok(false,'цикл '+i+': не открылся'); break; }
  const ws=els('#calBody .cw').filter(b=>!b.disabled);
  ws[Math.floor(rnd()*ws.length)].click();
  ws[Math.floor(rnd()*ws.length)].click();
  const r=screenAlive('dyn');
  if(!r.alive){ ok(false,'после выбора недель экран мёртв: '+r.why); break; }
  if(!cal.classList.contains('hide')) $('calClose').click();
}
ok(cal.classList.contains('hide'),'календарь закрыт после 5 циклов выбора');
$('calBtn').click(); els('#calQuick .chip').forEach(c=>{ c.click(); });
ok(true,'быстрые интервалы отработали, S='+S.from+'→'+S.to);
if(!cal.classList.contains('hide')) $('calClose').click();
ok(cal.classList.contains('hide'),'календарь закрывается в конце');

console.log('\n=== F. Сброс фильтров восстанавливает экран ===');
els('#fDir .chip')[0].click(); els('#fSrc .chip')[0].click();
drawn.length=0; $('btnReset').click();
ok(S.dirs.length===0&&S.srcs.length===0,'фильтры сброшены');
ok(screenAlive(S.view).alive,'экран жив после сброса');

console.log('\n=== G. CSV выгрузка ===');
let dl=null;
const realCreate=w.document.createElement.bind(w.document);
w.document.createElement=(t)=>{ const e=realCreate(t); if(t==='a'){ e.click=()=>{dl=e.download;}; } return e; };
$('btnCsv').click();
ok(!!dl,'CSV сформирован: '+dl);

console.log(`\n${'='.repeat(46)}\nПРОЙДЕНО ${pass}   ПРОВАЛЕНО ${fail}`);
process.exit(fail?1:0);
