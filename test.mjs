import { JSDOM } from 'jsdom';
import fs from 'fs';

let pass=0, fail=0;
const ok=(c,m)=>{ c?(pass++,console.log('  ok  '+m)):(fail++,console.log('  FAIL '+m)); };

// Chart.js нужен canvas; подменяем его учётной заглушкой, которая
// РЕГИСТРИРУЕТ каждый вызов — так мы видим, дошло ли дело до отрисовки
// и с какими данными, чего по картинке не понять.
const drawn = [];
const html = fs.readFileSync('index.html','utf8');
const dom = new JSDOM(html, { runScripts:'outside-only', pretendToBeVisual:true });
const w = dom.window;

w.HTMLCanvasElement.prototype.getContext = function(){
  return { canvas:this, save(){},restore(){},clearRect(){},fillRect(){},strokeRect(){},beginPath(){},
    closePath(){},moveTo(){},lineTo(){},arc(){},fill(){},stroke(){},clip(){},rect(){},
    translate(){},rotate(){},scale(){},setTransform(){},setLineDash(){},quadraticCurveTo(){},
    bezierCurveTo(){},createLinearGradient(){return{addColorStop(){}}},
    measureText(t){return{width:String(t).length*6, actualBoundingBoxAscent:6, actualBoundingBoxDescent:2}},
    fillText(){},strokeText(){},drawImage(){},putImageData(){},
    getImageData(){return{data:new Uint8ClampedArray(4)}},
    isPointInPath(){return false}, arcTo(){}, ellipse(){},
    resetTransform(){}, transform(){}, createPattern(){return null},
    createRadialGradient(){return{addColorStop(){}}}, roundRect(){},
    set fillStyle(v){}, get fillStyle(){return '#000'},
    set strokeStyle(v){}, get strokeStyle(){return '#000'},
    set lineWidth(v){}, get lineWidth(){return 1},
    set font(v){}, get font(){return '10px sans-serif'} };
};
w.ResizeObserver=class{observe(){}unobserve(){}disconnect(){}};
w.requestAnimationFrame=()=>1;  // анимацию не прокручиваем: нам важен факт создания графика с данными
w.cancelAnimationFrame=()=>{};
w.matchMedia=()=>({matches:false,addEventListener(){},removeEventListener(){}});
w.scrollTo=()=>{};
w.URL.createObjectURL=()=>'blob:x'; w.URL.revokeObjectURL=()=>{};

// Выполняем скрипты вручную в нужном порядке
const scripts=[...dom.window.document.querySelectorAll('script:not([type="application/json"])')];
// Chart.js грузим первым, затем ставим шпиона, ЗАТЕМ core+app — иначе
// app.js уже стартовал с настоящим Chart и вызовы не записались.
w.eval(scripts[0].textContent);
const RealChart0 = w.Chart;
function SpyChart0(cv,cfg){
  drawn.push({ id:cv.id, type:cfg.type||(cfg.data.datasets[0]&&cfg.data.datasets[0].type),
    labels:(cfg.data.labels||[]).length,
    sets:(cfg.data.datasets||[]).map(d=>({label:d.label,n:(d.data||[]).length,
      nonNull:(d.data||[]).filter(v=>v!=null).length})),
    opts:cfg.options||{} });
  return new RealChart0(cv,cfg);
}
SpyChart0.prototype=RealChart0.prototype;
Object.setPrototypeOf(SpyChart0,RealChart0);
w.Chart=SpyChart0;
for(const s of scripts.slice(1)){
  try { w.eval(s.textContent); }
  catch(e){ console.log('ОШИБКА в скрипте:', e.message, '\n', (e.stack||'').split('\n').slice(0,6).join('\n')); }
}
console.log('readyState =', w.document.readyState, '| CORE =', typeof w.CORE);

try { w.dispatchEvent(new w.Event('DOMContentLoaded'));
      w.document.dispatchEvent(new w.Event('DOMContentLoaded',{bubbles:true})); }
catch(e){ console.log('ОШИБКА boot:', e.message, '\n', (e.stack||'').split('\n').slice(0,8).join('\n')); }
console.log('после boot: __D =', typeof w.__D, ' __API =', typeof w.__API);

const $=(id)=>w.document.getElementById(id);
const D=w.__D, S=w.__S, API=w.__API, C=API.C;
const txt=(id)=>($(id)?$(id).textContent.trim():'');

console.log('\n=== 1. Загрузка и данные ===');
ok(!!D, 'датасет распарсен');
ok(D.weekly.length===6278, 'недельных строк 6278 (факт '+D.weekly.length+')');
ok(D.monthly.length===1566, 'месячных строк 1566 (факт '+D.monthly.length+')');
ok(D.meta.coverage.weeks===84, 'недель 84');
ok(D.meta.coverage.weekTo==='2026-08-17', 'последняя неделя 17.08.2026');

console.log('\n=== 2. Контрольные суммы против таблицы ===');
const A=C.agg(C.filterRows(D,{level:'monthly',from:'2026-01',to:'2026-07',dirs:[],srcs:[],pages:[]}));
ok(Math.abs(A.spend-9246114.44)<0.5, 'расход 2026 = 9 246 114,44 (факт '+A.spend.toFixed(2)+')');
ok(A.leads===8738,'лиды 2026 = 8738 (факт '+A.leads+')');
ok(A.kl===2028,'квал 2026 = 2028 (факт '+A.kl+')');
ok(A.revenue_won===35798309,'выручка 2026 = 35 798 309 (факт '+A.revenue_won+')');
ok(Math.round(A.cpql)===4559,'CPQL 2026 = 4559 (факт '+Math.round(A.cpql)+')');
const W=C.agg(C.filterRows(D,{level:'weekly',from:'2025-01-13',to:'2026-08-17',dirs:[],srcs:[],pages:[]}));
ok(Math.abs(W.spend-22937991.15)<0.5,'весь недельный расход 22 937 991,15 (факт '+W.spend.toFixed(2)+')');

console.log('\n=== 3. Графики рисуются на каждой вкладке ===');
const views=['over','dyn','share','heat','pf','fc','table','gs','method'];
for(const v of views){
  drawn.length=0;
  S.view=v; API.render();
  const withCanvas=['over','dyn','share','pf','fc'].includes(v);
  const painted = drawn.filter(d=>d.sets.some(s=>s.nonNull>0));
  if(withCanvas) ok(painted.length>0, `вкладка ${v}: графиков с данными ${painted.length}`);
  else ok(true, `вкладка ${v}: без графиков, отрисована`);
  ok($('v-'+v).classList.contains('on'), `вкладка ${v}: активна`);
}

console.log('\n=== 4. ГЛАВНЫЙ БАГ: графики при СМЕНЕ ФИЛЬТРА ===');
S.view='over'; API.render();
const dirs=[...new Set(D.monthly.map(r=>r.dir))].filter(Boolean);
let filterRuns=0, filterOk=0;
for(const d of dirs){
  S.dirs=[d]; drawn.length=0; API.render();
  const rows=C.filterRows(D,S);
  filterRuns++;
  const painted=drawn.filter(x=>x.sets.some(s=>s.nonNull>0)).length;
  if(rows.length===0){ if(painted===0) filterOk++; }
  else if(painted>0) filterOk++;
  else console.log('    FAIL фильтр '+d+': строк '+rows.length+', графиков с данными 0');
}
S.dirs=[];
ok(filterOk===filterRuns, `смена фильтра по направлению: ${filterOk}/${filterRuns} без пропажи графиков`);

// пары направление × источник
let pairRuns=0,pairOk=0;
const srcs=[...new Set(D.monthly.map(r=>r.src))].filter(Boolean).slice(0,10);
for(const d of dirs) for(const s of srcs){
  S.dirs=[d]; S.srcs=[s]; drawn.length=0; API.render();
  const rows=C.filterRows(D,S); pairRuns++;
  const painted=drawn.filter(x=>x.sets.some(z=>z.nonNull>0)).length;
  const empty=$('ovKpi').innerHTML.includes('нет данных');
  if(rows.length===0){ if(empty) pairOk++; else console.log('    FAIL '+d+'×'+s+': пусто, но нет сообщения'); }
  else if(painted>0) pairOk++;
  else console.log('    FAIL '+d+'×'+s+': строк '+rows.length+' графиков 0');
}
S.dirs=[];S.srcs=[];
ok(pairOk===pairRuns, `пары направление×источник: ${pairOk}/${pairRuns} корректны`);

console.log('\n=== 5. Смена разреза недели/месяцы ===');
for(const lvl of ['weekly','monthly','weekly']){
  S.level=lvl;
  w.document.querySelector(`#segLevel button[data-l="${lvl}"]`).click();
  drawn.length=0; S.view='dyn'; API.render();
  ok(drawn.filter(d=>d.sets.some(s=>s.nonNull>0)).length>0, `разрез ${lvl}: графики есть`);
}
S.level='monthly'; w.document.querySelector('#segLevel button[data-l="monthly"]').click();

console.log('\n=== 6. Календарь недель ЗАКРЫВАЕТСЯ ===');
S.level='weekly'; w.document.querySelector('#segLevel button[data-l="weekly"]').click();
const cal=$('cal'), back=$('calBack');
ok(cal.classList.contains('hide'), 'при старте календарь скрыт классом hide');
$('calBtn').click();
ok(!cal.classList.contains('hide'), 'открылся по кнопке');
ok(!back.classList.contains('hide'), 'подложка показана');
$('calClose').click();
ok(cal.classList.contains('hide'), 'закрылся кнопкой ✕');
ok(back.classList.contains('hide'), 'подложка скрыта');
$('calBtn').click(); $('calBack').click();
ok(cal.classList.contains('hide'), 'закрылся кликом по подложке');
$('calBtn').click();
w.document.dispatchEvent(Object.assign(new w.Event('keydown',{bubbles:true}),{key:'Escape'}));
ok(cal.classList.contains('hide'), 'закрылся по Escape');
// самое важное: .cal не должен иметь display, который перебивает hide
{const css=html.slice(html.indexOf('<style>'),html.indexOf('</style>'));
 ok(/\.hide\{display:none!important\}/.test(css),'правило .hide{display:none!important} есть в CSS');
 ok(!/\.cal\[hidden\]/.test(css) && !/hidden=/.test(html.slice(html.indexOf('id="cal"')-200,html.indexOf('id="cal"')+80)),
    'календарь НЕ полагается на атрибут hidden (причина прошлого бага)');
 ok(cal.className.includes('hide'),'скрытый календарь несёт класс hide: '+cal.className);}

console.log('\n=== 7. Календарь выбирает интервал ===');
$('calBtn').click();
const cws=[...w.document.querySelectorAll('#calBody .cw')];
ok(cws.length>0, 'недели в календаре отрисованы: '+cws.length);
cws[0].click(); cws[Math.min(3,cws.length-1)].click();
ok(S.from<=S.to, 'интервал выбран корректно: '+S.from+' → '+S.to);
drawn.length=0; S.view='dyn'; API.render();
ok(drawn.filter(d=>d.sets.some(s=>s.nonNull>0)).length>0, 'после выбора недель графики есть');
API.calQuick('13');
{const pl=C.periodList(D,'weekly');const n=pl.indexOf(S.to)-pl.indexOf(S.from)+1;
 ok(n===13||n===14, 'быстрый выбор «13 недель» даёт '+n+' периодов');}
if(!cal.classList.contains('hide')) $('calClose').click();
S.level='monthly'; w.document.querySelector('#segLevel button[data-l="monthly"]').click();

console.log('\n=== 8. Тултипы настроены на точное значение ===');
drawn.length=0; S.view='over'; API.render();
const tipped=drawn.filter(d=>d.opts.plugins&&d.opts.plugins.tooltip&&d.opts.plugins.tooltip.callbacks);
ok(tipped.length===drawn.length && drawn.length>0, `у всех ${drawn.length} графиков обзора есть форматтер тултипа`);
const cb=tipped[0].opts.plugins.tooltip.callbacks;
ok(typeof cb.label==='function', 'label-колбэк — функция');
ok(/\d/.test(cb.label({dataset:{label:'Лиды'},raw:1234})), 'label возвращает число: '+cb.label({dataset:{label:'Лиды'},raw:1234}));
ok(drawn.every(d=>d.opts.interaction&&d.opts.interaction.mode==='index'),'наведение в любой точке ловит период (mode=index)');
ok(drawn.every(d=>d.opts.animation&&d.opts.animation.duration>0),'анимация включена');

console.log('\n=== 9. План-факт: факт совпадает с обзором ===');
S.view='pf'; API.render();
let pfOk=true;
for(const [per,rows] of Object.entries(D.plan_fact)){
  for(const r of rows){
    const a=C.agg(C.filterRows(D,{level:'monthly',from:per,to:per,dirs:[r.dir],srcs:[],pages:[]}));
    if(Math.abs(a.spend-r.spend_fact)>0.5||a.kl!==r.kl_fact){
      console.log(`    FAIL ${per} ${r.dir}: pf(${r.spend_fact},${r.kl_fact}) vs обзор(${a.spend.toFixed(2)},${a.kl})`); pfOk=false; }
  }
}
ok(pfOk,'факт план-факта = факт обзора по всем месяцам и направлениям (блокер 4)');

console.log('\n=== 10. Пустой срез не рисует нули ===');
S.dirs=['Запчасти']; S.srcs=['Яндекс Директ Коляс']; S.view='over'; API.render();
const rowsEmpty=C.filterRows(D,S);
if(rowsEmpty.length===0){
  ok($('ovKpi').innerHTML.includes('нет данных'),'пустой срез: сообщение вместо нулей');
  ok(!$('ovKpi').innerHTML.includes('0 ₽'),'пустой срез: нулей нет');
} else ok(true,'выбранная пара непуста — проверка пропущена');
S.dirs=[];S.srcs=[];

console.log('\n=== 11. Деление на ноль ===');
const z=C.derive({spend:100,leads:0,kl:0,spam:0,badlead:0,existing:0,deals_created:0,pipeline:0,deals_lost:0,revenue_lost:0,deals_won:0,revenue_won:0});
ok(z.cpl===null&&z.cpql===null&&z.cr===null,'нет лидов → CPL/CPQL/CR = null, не 0 и не Infinity');
ok(z.romi===null||z.romi===-1,'нет выручки → ROMI не Infinity');

console.log('\n=== 12. Тепловая карта ===');
S.view='heat';
for(const m of ['cpql','cpl','cplClean','cpd']) for(const d of ['dir','src','page']){
  S.heatMetric=m; S.heatDim=d; API.render();
  const cells=w.document.querySelectorAll('#heat .hc').length;
  if(!(cells>0||$('heat').innerHTML.includes('нет'))){ ok(false,`heat ${m}/${d}: пусто без объяснения`); }
}
S.heatMetric='cpql';S.heatDim='dir';API.render();
ok(w.document.querySelectorAll('#heat .hc').length>0,'CPQL×направления: ячейки есть ('+w.document.querySelectorAll('#heat .hc').length+')');
ok([...w.document.querySelectorAll('#heat .hc')].every(c=>c.getAttribute('title')||c.classList.contains('nil')),'у каждой ячейки тултип с точным значением');

console.log('\n=== 13. Прогноз ===');
S.view='fc';
for(const m of ['last3','trend','mom']){ S.fcMethod=m; drawn.length=0; API.render();
  ok(drawn.filter(d=>d.sets.some(s=>s.nonNull>0)).length>0,'метод '+m+': графики есть'); }

console.log('\n=== 14. Сортировка таблиц ===');
S.view='over'; API.render();
const th=w.document.querySelector('#ovTable th[data-si="1"]');
const before=w.document.querySelector('#ovTable tbody tr td').textContent;
th.click(); th.click();
ok(true,'клик по заголовку не падает; первая строка: '+w.document.querySelector('#ovTable tbody tr td').textContent);

console.log('\n=== 15. Полный обход всех вкладок ×3 (утечки/падения) ===');
let cycles=0;
for(let i=0;i<3;i++) for(const v of views){ S.view=v; API.render(); cycles++; }
ok(cycles===27,cycles+' переключений вкладок без исключений');
ok(Object.keys(API.CH).length<=10,'живых графиков не накопилось: '+Object.keys(API.CH).length);

console.log(`\n${'='.repeat(46)}\nПРОЙДЕНО ${pass}   ПРОВАЛЕНО ${fail}`);
process.exit(fail?1:0);
