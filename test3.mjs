/* Проверки новых экранов: доля источников, редактируемый план-факт,
   интеграция с Гугл-таблицей. Здесь важна АРИФМЕТИКА долей и то, что
   правка плана реально доезжает до процентов и полос. */
import { JSDOM } from 'jsdom';
import fs from 'fs';
let pass=0,fail=0;
const ok=(c,m)=>{c?(pass++,console.log('  ok  '+m)):(fail++,console.log('  FAIL '+m));};
const drawn=[];
const html=fs.readFileSync('index.html','utf8');
const store={};
const dom=new JSDOM(html,{runScripts:'outside-only',pretendToBeVisual:true,url:'https://local.test/'});
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
w.requestAnimationFrame=()=>1; w.matchMedia=()=>({matches:false,addEventListener(){},removeEventListener(){}});
w.scrollTo=()=>{}; w.URL.createObjectURL=()=>'blob:x'; w.URL.revokeObjectURL=()=>{};
Object.defineProperty(w,'localStorage',{value:{getItem:(k)=>store[k]??null,
  setItem:(k,v)=>{store[k]=String(v);},removeItem:(k)=>{delete store[k];}},configurable:true});
// сеть подменяем: проверяем обработку ответов, а не живой Google
let fetchMode='ok';
w.fetch=async(url)=>{
  if(fetchMode==='neterr') throw new Error('Failed to fetch');
  if(fetchMode==='html') return {ok:true,status:200,text:async()=>'<!doctype html><html>login</html>'};
  if(fetchMode==='500') return {ok:false,status:500,text:async()=>'boom'};
  if(fetchMode==='badjson') return {ok:true,status:200,text:async()=>'{oops'};
  if(fetchMode==='noarray') return {ok:true,status:200,text:async()=>JSON.stringify({hello:1})};
  if(fetchMode==='missing') return {ok:true,status:200,text:async()=>JSON.stringify({weekly:[{period:'2026-08-17',dir:'X'}]})};
  const D=JSON.parse(w.document.getElementById('data').textContent);
  const rows=D.weekly.map(r=>({period:r.period,end:r.end,dir:r.dir,src:r.src,page:r.page,
    spend:r.spend,leads:r.leads,kl:r.kl,badlead:r.badlead,spam:r.spam,
    deals_created:r.deals_created,pipeline:r.pipeline,deals_won:r.deals_won,revenue_won:r.revenue_won}));
  if(fetchMode==='mismatch') rows.forEach(r=>r.spend=r.spend*2);
  if(fetchMode==='fresh') rows.push({period:'2026-08-24',dir:'Запчасти',src:'SEO',page:null,
    spend:1000,leads:10,kl:3,badlead:0,spam:0,deals_created:0,pipeline:0,deals_won:0,revenue_won:0});
  return {ok:true,status:200,text:async()=>JSON.stringify({weekly:rows,count:rows.length})};
};
const scripts=[...w.document.querySelectorAll('script:not([type="application/json"])')];
w.eval(scripts[0].textContent);
const R=w.Chart;
function Spy(cv,cfg){drawn.push({id:cv.id,type:cfg.type,
  sets:(cfg.data.datasets||[]).map(d=>({label:d.label,nonNull:(d.data||[]).filter(v=>v!=null).length})),
  opts:cfg.options||{}});return new R(cv,cfg);}
Spy.prototype=R.prototype;Object.setPrototypeOf(Spy,R);w.Chart=Spy;
for(const s of scripts.slice(1)) w.eval(s.textContent);
w.dispatchEvent(new w.Event('DOMContentLoaded'));
const $=(id)=>w.document.getElementById(id);
const els=(s)=>[...w.document.querySelectorAll(s)];
const D=w.__D,S=w.__S,API=w.__API,C=API.C;
const sleep=(ms)=>new Promise(r=>setTimeout(r,ms));

console.log('=== 1. Доля по источникам: арифметика ===');
S.view='share';S.level='monthly';S.from='2026-01';S.to='2026-07';S.dirs=[];S.srcs=[];API.render();
let sh=C.shares(D,S,'src');
ok(Math.abs(sh.rows.reduce((a,r)=>a+(r.shareKl||0),0)-1)<1e-9,'доли квала суммируются в 100%');
ok(Math.abs(sh.rows.reduce((a,r)=>a+(r.shareLeads||0),0)-1)<1e-9,'доли лидов суммируются в 100%');
ok(sh.raw.kl===2028,'квал в срезе 2028 (факт '+sh.raw.kl+')');
const top=sh.rows[0];
ok(top.key==='Яндекс Директ'&&top.kl===462,'лидер — Яндекс Директ, 462 квала');
ok(Math.abs(top.shareKl-462/2028)<1e-12,'доля лидера = 462/2028 = '+(top.shareKl*100).toFixed(1)+'%');
ok(sh.rows.every((r,i,a)=>i===0||a[i-1].cumKl<=r.cumKl+1e-12),'накопленная доля не убывает');
ok(Math.abs(sh.rows[sh.rows.length-1].cumKl-1)<1e-9,'накопленная доля доходит до 100%');

console.log('\n=== 2. Доля пересчитывается ОТ СРЕЗА при фильтре ===');
let bad=0;
for(const d of [...new Set(D.monthly.map(r=>r.dir))].filter(Boolean)){
  S.dirs=[d]; API.render();
  const s2=C.shares(D,S,'src');
  const rows=C.filterRows(D,S);
  const klSum=rows.reduce((a,r)=>a+r.kl,0);
  if(s2.raw.kl!==klSum){console.log('    FAIL '+d+': итог квала '+s2.raw.kl+' vs '+klSum);bad++;}
  const sum=s2.rows.reduce((a,r)=>a+(r.shareKl||0),0);
  if(klSum>0&&Math.abs(sum-1)>1e-9){console.log('    FAIL '+d+': доли не 100% ('+sum+')');bad++;}
}
S.dirs=[];
ok(bad===0,'по каждому направлению доли считаются от его собственного итога');
S.dirs=['Запчасти'];API.render();
const shZ=C.shares(D,S,'src');
ok(shZ.raw.kl===640,'Запчасти: квал 640 (факт '+shZ.raw.kl+')');
ok(shZ.rows[0].key==='SEO','внутри Запчастей лидер SEO с '+(shZ.rows[0].shareKl*100).toFixed(1)+'%');
S.dirs=[];

console.log('\n=== 3. Экран доли: графики и таблица ===');
for(const dim of ['src','dir','page']) for(const m of ['kl','leads']){
  S.shDim=dim;S.shMetric=m;drawn.length=0;API.render();
  const painted=drawn.filter(d=>d.sets.some(s=>s.nonNull>0)).length;
  if(painted===0){ok(false,`доля ${dim}/${m}: графиков 0`);}
}
S.shDim='src';S.shMetric='kl';drawn.length=0;API.render();
ok(drawn.length===4,'на экране доли 4 графика (факт '+drawn.length+')');
ok(drawn.some(d=>d.type==='doughnut'),'есть кольцевая диаграмма структуры');
ok(drawn.every(d=>d.opts.plugins&&d.opts.plugins.tooltip&&d.opts.plugins.tooltip.callbacks),'у всех графиков доли есть тултип');
ok(els('#shTable tbody tr').length>0,'таблица раскладки заполнена: строк '+els('#shTable tbody tr').length);
ok($('shTable').textContent.includes('%'),'в таблице есть проценты');
ok(els('#shKpi .kpi').length===4,'4 KPI на экране доли');

console.log('\n=== 4. План-факт: правка плана доезжает до процентов ===');
S.view='pf';API.render();
S.pfPeriod='2026-07';API.render();
const pins=els('#pfT .pin');
ok(pins.length>0,'поля плана редактируемы: '+pins.length+' полей');
const pin=els('#pfT .pin').find(i=>i.dataset.k==='kl_plan');
const dirEd=pin.dataset.dir;
const factKl=API.factFor('2026-07',dirEd).kl;
pin.value=String(factKl*2); pin.onchange();
ok(API.planFor('2026-07',dirEd).kl_plan===factKl*2,'план сохранён в состоянии: '+API.planFor('2026-07',dirEd).kl_plan);
ok(store['alfakpd.plans.v1'] && store['alfakpd.plans.v1'].includes(dirEd),'план записан в localStorage');
{const row=els('#pfT tbody tr').find(tr=>tr.textContent.includes(dirEd));
 ok(/\b50%/.test(row.textContent),'выполнение показано как 50% (план вдвое больше факта): '+/\d+%/.exec(row.textContent));}
pin.value=String(factKl); pin.onchange();
{const row=els('#pfT tbody tr').find(tr=>tr.textContent.includes(dirEd));
 ok(/100%/.test(row.textContent),'план=факт → 100%');}

console.log('\n=== 5. План-факт: новое направление и перенос плана ===');
API.pfAddDir('Тест-Направление');
ok(els('#pfT tbody tr').some(t=>t.textContent.includes('Тест-Направление')),'новое направление добавлено в таблицу');
const before=Object.keys(API.PLANS()['2026-07']||{}).length;
API.pfCopyFrom('2026-02');
ok(Object.keys(API.PLANS()['2026-07']||{}).length>=before,'перенос плана из другого месяца отработал');
API.pfClearManual();
ok(!API.PLANS()['2026-07'],'ручные правки месяца снимаются одной кнопкой');
ok(!els('#pfT tbody tr').some(t=>t.textContent.includes('Тест-Направление')),'тестовое направление ушло вместе с правками');

console.log('\n=== 6. План-факт: факт по-прежнему = факт обзора ===');
bad=0;
for(const per of Object.keys(D.plan_fact)){
  for(const r of (D.plan_fact[per]||[])){
    const f=API.factFor(per,r.dir);
    const a=C.agg(C.filterRows(D,{level:'monthly',from:per,to:per,dirs:[r.dir],srcs:[],pages:[]}));
    if(Math.abs(f.spend-a.spend)>0.5||f.kl!==a.kl){console.log('    FAIL '+per+' '+r.dir);bad++;}
  }
}
ok(bad===0,'факт план-факта считается из тех же строк, что и обзор');
ok(els('#pfKpi .kpi').length===4,'инфографика выполнения: 4 сводных показателя');
ok($('pfKpi').innerHTML.includes('pf-bar'),'в сводке есть полосы выполнения');

console.log('\n=== 7. Гугл-таблица: успешная проверка ===');
S.view='gs';API.render();
ok($('gsBody').textContent.includes('Apps Script'),'инструкция отрисована');
ok($('gsCode').textContent.includes('doGet'),'код скрипта на странице');
ok($('gsCode').textContent.includes('Недельный отчёт по трафику'),'в коде правильное имя листа');
ok(els('.chk tbody tr').length>=7,'чек-лист: '+els('.chk tbody tr').length+' пунктов');
$('gsUrl').value='https://script.google.com/macros/s/AAA/exec';
fetchMode='ok'; await API.gsCheck(); await sleep(10);
ok($('gsLog').textContent.includes('Интеграция работает'),'итог: интеграция работает');
ok($('gsLog').textContent.includes('Расход сходится'),'проверка сверяет расход');
ok(!$('gsApply').disabled,'кнопка обновления разблокирована');
ok(store['alfakpd.gsurl.v1']==='https://script.google.com/macros/s/AAA/exec','ссылка запомнена');

console.log('\n=== 8. Гугл-таблица: диагностика поломок ===');
const cases=[
  ['neterr','Соединение не установлено'],
  ['html','HTML-страница вместо данных'],
  ['500','HTTP 500'],
  ['badjson','не разобран как JSON'],
  ['noarray','нет массива строк'],
  ['missing','не хватает'],
  ['mismatch','цифры расходятся']
];
for(const [mode,expect] of cases){
  fetchMode=mode; await API.gsCheck(); await sleep(5);
  const got=$('gsLog').textContent;
  ok(got.includes(expect),`режим «${mode}» → сообщение содержит «${expect}»`);
}

console.log('\n=== 9. Гугл-таблица: обновление данных ===');
fetchMode='fresh'; await API.gsCheck(); await sleep(5);
ok($('gsLog').textContent.includes('более свежие'),'свежие недели замечены');
const wasWeeks=D.meta.coverage.weeks, wasTo=D.meta.coverage.weekTo;
API.gsApply();
// gsApply заменяет D.weekly внутри модуля — перечитываем актуальный объект
const D2=w.__D;
ok(D2.meta.coverage.weekTo==='2026-08-24','покрытие обновилось: '+wasTo+' → '+D2.meta.coverage.weekTo);
ok(D2.meta.coverage.weeks===wasWeeks+1,'недель стало на одну больше: '+wasWeeks+' → '+D2.meta.coverage.weeks);
ok(D2.weekly.some(r=>r.period==='2026-08-24'),'новая неделя доехала до данных');
S.view='over';S.level='weekly';API.render();
ok(true,'после обновления дэшборд перерисовался без падения');
S.level='monthly';API.render();

console.log('\n=== 10. Живые клики по новым экранам ===');
bad=0;
for(const t of els('.tab')){
  t.click();
  for(const ch of els('#fDir .chip').slice(0,4)){
    ch.click();
    const root=$('v-'+S.view);
    const rows=C.filterRows(D,S).length;
    const live=Object.keys(API.CH).filter(id=>{const c=$(id);return c&&root.contains(c);}).length;
    const said=/нет данных|нужно минимум|плана|Apps Script|формула/i.test(root.textContent);
    // heat / table / gs / method живут содержимым, а не графиками
    const content=root.querySelectorAll('.hc,tbody tr,.card').length;
    const alive = live>0 || said || content>0;
    if(rows>0&&!alive){console.log('    FAIL '+S.view+' + '+ch.dataset.v);bad++;}
    ch.click();
  }
}
ok(bad===0,'клики по фильтрам на всех 9 вкладках безопасны');

console.log(`\n${'='.repeat(46)}\nПРОЙДЕНО ${pass}   ПРОВАЛЕНО ${fail}`);
process.exit(fail?1:0);
