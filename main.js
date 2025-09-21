
/* Utility helpers */
const html = String.raw;
const $ = sel => document.querySelector(sel);
function systemId(s){return (s||"").toString().trim().toLowerCase().replace(/\s+/g,' ').replace(/[（）()]/g,'').replace(/[^\w\u4e00-\u9fa5\s\-\/\+]/g,'')}
function toNumber(x){
  if(x==null) return null;
  let s = String(x).trim();
  if(!s) return null;
  // normalize dashes
  s = s.replace(/[～~≈∼]/g,'~').replace(/[—–−]/g,'-');
  // n/a
  if(/\b(n\/?a|na)\b/i.test(s)) return null;
  // BR tokens turn single number to negative
  const hasBR = /\bbr\b/i.test(s) || /\bbr[\s\-]?\d+/i.test(s);
  // grab all integers with sign if present
  const nums = (s.match(/-?\d+/g)||[]).map(v=>parseInt(v,10));
  if(nums.length===0) return null;
  // if 2+ numbers present, take midpoint of min/max (range/approx handled)
  if(nums.length>=2){
    const lo = Math.min(...nums), hi = Math.max(...nums);
    return (lo + hi) / 2;
  }
  let val = nums[0];
  if(hasBR) val = -Math.abs(val);
  return val;
}
function quantile(arr, q){
  if(!arr || arr.length===0) return null;
  const a = arr.slice().sort((x,y)=>x-y);
  const pos = (a.length-1)*q;
  const base = Math.floor(pos);
  const rest = pos-base;
  if(a[base+1]!==undefined) return a[base] + rest*(a[base+1]-a[base]);
  return a[base];
}
function median(arr){return quantile(arr, 0.5)}
function uniq(arr){return Array.from(new Set(arr))}
function clamp(x,a,b){return Math.max(a, Math.min(b,x))}

/* Child-facing and parent-facing narratives */
function bandTag(L, bands){
  if(L==null || !bands) return '未知';
  const {p20,p40,p60,p80} = bands;
  if([p20,p40,p60,p80].some(v=>v==null)) return '参考';
  if(L < p20) return '偏易';
  if(L < p40) return '舒适';
  if(L <= p60) return '舒适';
  if(L <= p80) return '最佳挑战';
  return '偏难';
}

function childEvaluationText(lexiles, bands){
  const tags = lexiles.map(L => bandTag(L, bands));
  const count = t => tags.filter(x=>x===t).length;
  const easy=count('偏易'), comfy=count('舒适'), best=count('最佳挑战'), hard=count('偏难');
  const parts=[];
  if(hard>=1 && best>=1){
    parts.push('今天有一部分内容会有点“挑战”，但也搭配了较合适的书，先稳稳读，再去挑战就很好。');
  }else if(best>=1 && easy>=1){
    parts.push('先读轻松的热身书，再读一点点挑战的内容，你会感觉既顺利又有提升。');
  }else if(comfy>=2){
    parts.push('今天的书大多在你的舒适区内，你会读得比较流畅，适合练速度与表达。');
  }else if(hard>=2){
    parts.push('今天的书整体偏难，别担心，我们可以配合听读或一起读。');
  }else{
    parts.push('今天这几本书难度搭配比较均衡，可以安心开读。');
  }
  if(easy>=1) parts.push('有一本比较容易，适合做“热身”或“复读”，把语音语调练稳。');
  if(best>=1) parts.push('也有一本稍微有挑战，遇到生词可以先猜测、再查证，或者和家长一起读。');
  if(hard>=1) parts.push('如果感觉有点吃力，先听一遍音频再跟读，会轻松很多。');
  return parts.join(' ');
}

function parentSequenceAdvice(lexiles, bands){
  const idxs = lexiles.map((L,i)=>({i, L, tag: bandTag(L,bands)}));
  // recommended order: easy -> comfy -> challenge -> hard
  const order = idxs.slice().sort((a,b)=>{
    const rank = t => ({'偏易':0,'舒适':1,'最佳挑战':2,'偏难':3,'参考':1,'未知':1}[t] ?? 1);
    return rank(a.tag) - rank(b.tag);
  }).map(o=>o.i);
  const desc = [];
  desc.push('建议阅读顺序：从较容易的开始热身 → 进入舒适区的主读本 → 最后处理略有挑战的文本。');
  if(idxs.some(x=>x.tag==='偏难')) desc.push('如出现“偏难”文本，请采用“先听后读/分段共读/关键词扫读后精读”等方式。');
  if(!idxs.some(x=>x.tag==='偏易')) desc.push('当前组合缺少易文，建议补充一篇用于流利度或重复朗读练习。');
  return {order, text: desc.join(' ')};
}

function theoryNotes(){
  return [
    '难度匹配：组合的中位难度落在舒适—轻挑战区间（约P40–P80）通常阅读体验最佳。',
    '流利度支持：加入一篇易文做热身/复读，有助于速度、准确与理解提升。',
    '听读结合：当包含偏难文本时，先听后读或边听边读能降低解码负荷，提升理解。',
    '体裁多样：若今天全是故事类，可偶尔加入信息类/说明文，帮助知识建构与结构意识。',
    '动机与选择：允许孩子对书目有选择权，更能促进投入与理解。'
  ];
}

/* State */
const AppState = {
  loaded:false,
  basicRows:[],            // from 基本信息
  diffRows:[],             // from 语言难度数值表
  systems:[],              // [{id, name}]
  levelsBySystem:{},       // id -> [{level, lexileRaw, lexile, band, features, source}]
  basicBySystem:{},        // id -> basic row
  ageBuckets:{},           // age(int) -> [lexiles]
  ageBands:{},             // age(int) -> {p20,p40,p60,p80}
  globalLexiles:[],
  calibRows:[],
};

/* Load Excel fully */
async function loadExcel(){
  const resp = await fetch('./data/用户体验数据库.xlsx');
  const respCal = await fetch('./data/leveled_readers_detailed_levels.xlsx');
  if(!resp.ok) throw new Error('无法加载 Excel：'+resp.status);
  const buf = await resp.arrayBuffer();
  const bufCal = await respCal.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'});
  const wbCal = XLSX.read(bufCal, {type:'array'});
  // Find sheets by name fallback by index
  const basicName = wb.SheetNames.find(n=>/基本信息/.test(n)) || wb.SheetNames[0];
  const diffName = wb.SheetNames.find(n=>/语言难度数值表/.test(n)) || wb.SheetNames[1] || wb.SheetNames[0];
  const basic = XLSX.utils.sheet_to_json(wb.Sheets[basicName], {defval:null, raw:false});
  const diff  = XLSX.utils.sheet_to_json(wb.Sheets[diffName],  {defval:null, raw:false});
  const calName = wbCal.SheetNames.find(n=>/Sheet1|校准|calib/i.test(n)) || wbCal.SheetNames[0];
  const calib = XLSX.utils.sheet_to_json(wbCal.Sheets[calName], {defval:null, raw:false});
  AppState.basicRows = basic;
  AppState.diffRows = diff;
  AppState.calibRows = calib;
  normalize();
  AppState.loaded = true;
}

/* Normalize and build indices */
function parseAges(s){
  if(s==null) return [];
  s = String(s).trim();
  if(!s) return [];
  // Extract numbers and ranges like 6–7 / 3-5 / 单值 6
  const m = s.match(/(\d+)\s*[–-]\s*(\d+)/);
  if(m){
    const a = parseInt(m[1],10), b = parseInt(m[2],10);
    const lo = Math.min(a,b), hi = Math.max(a,b);
    const out=[];
    for(let k=lo;k<=hi;k++) out.push(k);
    return out;
  }
  const one = s.match(/^\d+$/);
  if(one) return [parseInt(one[0],10)];
  // fallback: capture all digits and clamp 3..15
  const nums = (s.match(/\d+/g)||[]).map(x=>parseInt(x,10)).filter(x=>x>=3 && x<=15);
  return uniq(nums);
}

function normalize(){
  // Build basicBySystem
  AppState.basicBySystem = {};
  for(const row of AppState.basicRows){
    const name = (row['系列'] ?? row['系统'] ?? '').toString().trim();
    if(!name) continue;
    const id = systemId(name);
    AppState.basicBySystem[id] = {
      id, name,
      pub: row['出版者/来源'] ?? '',
      position: row['特点/课程定位'] ?? '',
      levelRange: row['分级范围'] ?? '',
      difficultyNote: row['难度/单词量/句子复杂度'] ?? '',
      textFeatures: row['文本形式/可读性特征'] ?? ''
    };
  }

  // Prepare containers
  AppState.levelsBySystem = {};
  AppState.ageBuckets = {};
  AppState.globalLexiles = [];

  // 1) Use calibration rows as authoritative level→lexile mapping
  const addLevel = (id, level, lexRaw, band, features, source) => {
    if(!AppState.levelsBySystem[id]) AppState.levelsBySystem[id] = [];
    const lex = toNumber(lexRaw);
    const item = {level, band, lexileRaw: String(lexRaw||''), lexile: lex, features: (features||'').toString().trim(), source:(source||'')+''};
    // avoid duplicates by level name (keep calibration first)
    const existsIdx = AppState.levelsBySystem[id].findIndex(x=> (x.level||'')===(level||''));
    if(existsIdx>=0){
      // Prefer non-null numeric lexile
      const prev = AppState.levelsBySystem[id][existsIdx];
      if(prev.lexile==null && item.lexile!=null) AppState.levelsBySystem[id][existsIdx] = item;
      else if(prev.lexile!=null && item.lexile==null) {/* keep prev */}
      else AppState.levelsBySystem[id][existsIdx] = item; // overwrite with calib
    }else{
      AppState.levelsBySystem[id].push(item);
    }
    return item;
  };

  // Parse ages from a variety of column names
  const ageVal = r => r['国内小读者年龄 (近似)'] ?? r['中国年纪/年龄 (近似)'] ?? r['年龄'] ?? '';

  // From calibration table
  for(const r of (AppState.calibRows||[])){
    const id = systemId((r['系统'] ?? r['系列'] ?? '').toString().trim());
    const level = (r['分级'] ?? r['级别'] ?? '').toString().trim();
    const band = (r['对应 Book Band / 注'] ?? '').toString().trim();
    const features = (r['可比难度特征'] ?? '').toString().trim();
    const source = (r['来源'] ?? '').toString().trim();
    const added = addLevel(id, level, (r['蓝思值 (Lexile)'] ?? r['Lexile'] ?? ''), band, features, source);

    const ages = parseAges(ageVal(r));
    if(added.lexile!=null){
      AppState.globalLexiles.push(added.lexile);
      for(const a of ages){
        if(!AppState.ageBuckets[a]) AppState.ageBuckets[a] = [];
        AppState.ageBuckets[a].push(added.lexile);
      }
    }
  }

  // 2) Merge difficulty table as fallback (fills missing levels or missing lexile)
  for(const r of (AppState.diffRows||[])){
    const id = systemId((r['系统'] ?? r['系列'] ?? '').toString().trim());
    const level = (r['分级'] ?? r['级别'] ?? '').toString().trim();
    const band = (r['对应 Book Band / 注'] ?? '').toString().trim();
    const features = (r['可比难度特征'] ?? '').toString().trim();
    const source = (r['来源'] ?? '').toString().trim();
    const added = addLevel(id, level, (r['蓝思值 (Lexile)'] ?? r['Lexile'] ?? ''), band, features, source);

    const ages = parseAges(ageVal(r));
    if(added.lexile!=null){
      AppState.globalLexiles.push(added.lexile);
      for(const a of ages){
        if(!AppState.ageBuckets[a]) AppState.ageBuckets[a] = [];
        AppState.ageBuckets[a].push(added.lexile);
      }
    }
  }

  // Build systems list
  const ids = uniq(Object.keys(AppState.levelsBySystem));
  AppState.systems = ids.map(id => ({id, name: (AppState.basicBySystem[id]?.name) || id}))
                        .sort((a,b)=>a.name.localeCompare(b.name,'zh'));

  // Age bands
  const g = AppState.globalLexiles.slice().sort((x,y)=>x-y);
  const gBands = {p20:quantile(g,0.2), p40:quantile(g,0.4), p60:quantile(g,0.6), p80:quantile(g,0.8)};
  for(let age=3; age<=15; age++){
    const arr = (AppState.ageBuckets[age]||[]).slice().sort((x,y)=>x-y);
    const bands = arr.length>=8 ? {
      p20: quantile(arr,0.2), p40:quantile(arr,0.4), p60:quantile(arr,0.6), p80:quantile(arr,0.8)
    } : gBands;
    AppState.ageBands[age] = bands;
  }

  // Sort levels within system
  for(const id of Object.keys(AppState.levelsBySystem)){
    AppState.levelsBySystem[id].sort((a,b)=>{
      if(a.lexile!=null && b.lexile!=null) return a.lexile - b.lexile;
      if(a.lexile!=null) return -1;
      if(b.lexile!=null) return 1;
      return (a.level||'').localeCompare(b.level||'', 'zh');
    });
  }
}


  // Build levelsBySystem + age buckets
  AppState.levelsBySystem = {};
  AppState.ageBuckets = {};
  AppState.globalLexiles = [];

  for(const row of AppState.diffRows){
    const sysName = (row['系统'] ?? row['系列'] ?? '').toString().trim();
    const id = systemId(sysName);
    const level = (row['分级'] ?? row['级别'] ?? '').toString().trim();
    const band = (row['对应 Book Band / 注'] ?? '').toString().trim();
    const lexRaw = row['蓝思值 (Lexile)'] ?? row['Lexile'] ?? '';
    const lex = toNumber(lexRaw);
    const ages = parseAges(row['国内小读者年龄 (近似)'] ?? '');
    const features = (row['可比难度特征'] ?? '').toString().trim();
    const source = (row['来源'] ?? '').toString().trim();

    if(!AppState.levelsBySystem[id]) AppState.levelsBySystem[id] = [];
    AppState.levelsBySystem[id].push({
      level, band, lexileRaw: String(lexRaw||''), lexile: lex,
      features, source
    });

    if(lex!=null){
      AppState.globalLexiles.push(lex);
      for(const a of ages){
        if(!AppState.ageBuckets[a]) AppState.ageBuckets[a] = [];
        AppState.ageBuckets[a].push(lex);
      }
    }
  }

  // Build systems list: intersection that actually has levels
  const ids = uniq(Object.keys(AppState.levelsBySystem));
  AppState.systems = ids.map(id => ({id, name: (AppState.basicBySystem[id]?.name) || id}))
                        .sort((a,b)=>a.name.localeCompare(b.name,'zh'));

  // Age bands with fallback to global
  const g = AppState.globalLexiles.slice().sort((x,y)=>x-y);
  const gBands = {p20:quantile(g,0.2), p40:quantile(g,0.4), p60:quantile(g,0.6), p80:quantile(g,0.8)};
  for(let age=3; age<=15; age++){
    const arr = (AppState.ageBuckets[age]||[]).slice().sort((x,y)=>x-y);
    const bands = arr.length>=8 ? {
      p20: quantile(arr,0.2), p40:quantile(arr,0.4), p60:quantile(arr,0.6), p80:quantile(arr,0.8)
    } : gBands; // fallback
    AppState.ageBands[age] = bands;
  }

  // Sort levels within system by numeric Lexile if possible, otherwise by natural order
  for(const id of Object.keys(AppState.levelsBySystem)){
    AppState.levelsBySystem[id].sort((a,b)=>{
      if(a.lexile!=null && b.lexile!=null) return a.lexile - b.lexile;
      if(a.lexile!=null) return -1;
      if(b.lexile!=null) return 1;
      return (a.level||'').localeCompare(b.level||'', 'zh');
    });
  }
}

/* UI Components (no build) */
const e = React.createElement;

function SelectorRow({index, systems, levelsBySystem, value, onChange}){
  const sysId = value.systemId || '';
  const level = value.level || '';
  const levels = sysId ? (levelsBySystem[sysId]||[]) : [];
  return e('div', {className:'row'},
    e('div', null,
      e('label', null, `选项 ${index+1} · 选择读物`),
      e('select', {
        value: sysId,
        onChange: ev => onChange({systemId:ev.target.value, level:''})
      },
        e('option',{value:''}, '（不选）'),
        systems.map(s => e('option', {key:s.id, value:s.id}, s.name))
      )
    ),
    e('div', null,
      e('label', null, '选择该读物的级别'),
      e('select', {
        value: level,
        onChange: ev => onChange({systemId:sysId, level:ev.target.value})
      },
        e('option',{value:''}, sysId ? '（请选择级别）' : '（先选择读物）'),
        levels.map((lv,i)=> e('option',{key:i, value:lv.level}, lv.level || '(无名级别)'))
      ),
      e('div', {style:{marginTop:'4px'}}, e('small',{className:'help'}, sysId && level ?
        (()=>{
          const lv = (levels.find(x=>x.level===level)||{});
          const lex = lv.lexile!=null ? `${lv.lexile}L` : (lv.lexileRaw||'N/A');
          return `Lexile: ${lex}${lv.band?` · ${lv.band}`:''}`
        })() : '')
      )
    )
  );
}

function Report({profile, selections, ageBands, basicBySystem, levelsBySystem}){
  const age = profile.age, gender = profile.gender;
  const bands = ageBands[age] || {};
  const chosen = selections
    .map(s => {
      if(!s.systemId || !s.level) return null;
      const sys = basicBySystem[s.systemId] || {name:s.systemId};
      const lv = (levelsBySystem[s.systemId]||[]).find(x=>x.level===s.level) || {};
      return { systemId:s.systemId, systemName: sys.name, level:s.level, lexile: lv.lexile, lexileRaw: lv.lexileRaw, band:lv.band, features: lv.features, textFeatures: sys.textFeatures };
    }).filter(Boolean);

  const warnings = [];
  const lexiles = chosen.map(x=>x.lexile).filter(x=>x!=null);
  if(chosen.length===0) return e('div', null, e('div',{className:'warntext'}, '未选择有效读物，请至少选择 1 本并指定级别。'));
  if(lexiles.length < chosen.length) warnings.push('部分级别为区间或文本型数值，已取中位估计或暂缺。');

  const med = lexiles.length? median(lexiles): null;
  const avg = lexiles.length? (lexiles.reduce((a,b)=>a+b,0)/lexiles.length): null;
  const spread = lexiles.length? (Math.max(...lexiles) - Math.min(...lexiles)) : null;

  function classify(L){
    if(L==null) return {tag:'未知', cls:'', msg:'区间或文本型数值，已取中位估计'};
    const p20=bands.p20, p40=bands.p40, p60=bands.p60, p80=bands.p80;
    if([p20,p40,p60,p80].some(v=>v==null)) return {tag:'参考', cls:'', msg:'年龄基线不足，按全局估计'};
    if(L < p20) return {tag:'偏易', cls:'warn', msg:'建议提高级别或增加任务复杂度'};
    if(L < p40) return {tag:'舒适', cls:'ok', msg:'流畅阅读，适合巩固'};
    if(L <= p60) return {tag:'舒适', cls:'ok', msg:'流畅阅读，适合巩固'};
    if(L <= p80) return {tag:'最佳挑战', cls:'ok', msg:'略高于舒适，有助于提升'};
    return {tag:'偏难', cls:'bad', msg:'建议拆分任务或加入听读支持'};
  }
  const overall = med!=null ? classify(med) : {tag:'未知',cls:'',msg:'无可计算的 Lexile'};

  return e('div', {id:'report', className:'card'},
    e('h2', null, '阅读体验报告'),
    e('div', {className:'kv'},
      e('div', null, '年龄'), e('div', null, age),
      e('div', null, '性别'), e('div', null, gender==='male'?'男':'女'),
      e('div', null, '总体评估'),
      e('div', null, e('span',{className:`badge ${overall.cls}`}, overall.tag), '　', overall.msg)
    ),
    // --- 新增：孩子评价语 ---
    (function(){
      const narrative = childEvaluationText(lexiles, bands);
      return e('div', null,
        e('h3', null, '给孩子的话'),
        e('div', {className:'sub'}, narrative)
      );
    })(),
    // --- 新增：家长视角（顺序与建议） ---
    (function(){
      const adv = parentSequenceAdvice(lexiles, bands);
      const seq = adv.order.map(i=>`第${i+1}本`).join(' → ');
      return e('div', null,
        e('h3', null, '家长视角'),
        e('div', {className:'sub'}, adv.text),
        e('div', {className:'sub'}, '建议顺序：', seq || '（按孩子兴趣调整）')
      );
    })(),
    // --- 新增：理论依据（好的当日阅读） ---
    (function(){
      const notes = theoryNotes();
      return e('div', null,
        e('h3', null, '理论依据（好的当日阅读）'),
        e('ul', null, notes.map((t,i)=> e('li', {key:i}, t)))
      );
    })(),
    e('table', {className:'table'},
      e('thead', null, e('tr', null,
        e('th', null, '读物'),
        e('th', null, '级别'),
        e('th', null, 'Lexile'),
        e('th', null, '匹配'),
        e('th', null, '提示')
      )),
      e('tbody', null,
        chosen.map((c,i)=>{
          const cls = classify(c.lexile);
          return e('tr', {key:i},
            e('td', null, c.systemName),
            e('td', null, c.level),
            e('td', null, (c.lexile!=null? `${c.lexile}L` : null) || (c.lexileRaw || 'N/A')),
            e('td', null, e('span',{className:`badge ${cls.cls}`}, cls.tag)),
            e('td', null, (c.lexile!=null? cls.msg : (c.lexileRaw? '区间或文本型数值，已取中位估计' : cls.msg)))
          )
        })
      )
    ),
    e('div', {className:'sub', style:{marginTop:'10px'}},
      `统计口径：按所选年龄段在数据集中观测到的 Lexile 分布计算分位（P20/P40/P60/P80）；当样本不足时使用全局分布作为回退。`
    ),
    e('hr'),
    e('h3', null, '阅读建议'),
    e('ul', null,
      chosen.map((c,i)=> e('li', {key:i},
        e('strong', null, `${c.systemName} · ${c.level}：`),
        '来自数据的可读性提示：', (c.features||c.textFeatures||'暂无；可根据孩子表现加入复述、找关键词、跟读等任务。')
      ))
    ),
    warnings.length ? e('div', {className:'warntext', style:{marginTop:'8px'}}, '注意：', warnings.join('；')) : null,
    e('div', {className:'footer'}, '本结果为基于分位的快速估计，建议结合孩子当日状态做微调。')
  );
}

function App(){
  const [ready, setReady] = React.useState(false);
  const [loading, setLoading] = React.useState(true);
  const [systems, setSystems] = React.useState([]);
  const [profile, setProfile] = React.useState({age:8, gender:'female'});
  const [rows, setRows] = React.useState([{},{},{}]);
  const [showReport, setShowReport] = React.useState(false);

  React.useEffect(()=>{
    (async()=>{
      try{
        await loadExcel();
        setSystems(AppState.systems);
        setReady(true);
      }catch(e){
        const c=document.createElement('div');c.className='card';c.style.maxWidth='900px';c.style.margin='16px auto';c.innerHTML='<h2>数据加载失败</h2><div class="sub">'+(e&&e.message?e.message:'未知错误')+'。请确认仓库中存在 <code>data/用户体验数据库.xlsx</code> 与 <code>data/leveled_readers_detailed_levels.xlsx</code> ，并稍后刷新重试。</div>';document.getElementById('app').appendChild(c);
      }finally{
        setLoading(false);
      }
    })();
  },[]);

  function updateRow(idx, val){
    const copy = rows.slice();
    copy[idx] = {...copy[idx], ...val};
    setRows(copy);
  }

  function calculate(){
    // Validate: if picked system must pick level
    for(const r of rows){
      if(r.systemId && !r.level){
        alert('已选择读物但未选择级别，请为该行选择级别。');
        return;
      }
    }
    setShowReport(true);
    // scroll
    setTimeout(()=>document.getElementById('report')?.scrollIntoView({behavior:'smooth'}), 50);
  }

  async function exportPNG(){
    const el = document.getElementById('report');
    if(!el){ alert('请先生成报告'); return; }
    const canvas = await html2canvas(el, {scale:2, backgroundColor:'#0b1020'});
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url; a.download = `阅读体验报告_${profile.age}_${profile.gender}.png`;
    a.click();
  }

  return e('div', {className:'container'},
    e('div', {className:'card '+(loading?'loading':'')},
      e('h1', null, '少儿英语分级读物 · 阅读体验评估'),
      e('div', {className:'sub'}, '选择孩子画像与 1–3 本读物（每行先选读物，再选该读物的真实级别），系统将基于 Excel 的完整数据计算当日阅读体验。'),
      e('div', {className:'grid'},
        e('div', null,
          e('label', null, '孩子年龄（3–15）'),
          e('select', {
            value: profile.age,
            onChange: ev => setProfile({...profile, age: parseInt(ev.target.value,10)})
          },
            Array.from({length:13}, (_,i)=>3+i).map(a => e('option',{key:a, value:a}, a))
          )
        ),
        e('div', null,
          e('label', null, '性别'),
          e('select', {
            value: profile.gender,
            onChange: ev => setProfile({...profile, gender: ev.target.value})
          },
            e('option',{value:'female'}, '女'),
            e('option',{value:'male'}, '男')
          )
        )
      ),
      e('hr', null),
      e(SelectorRow, {index:0, systems, levelsBySystem:AppState.levelsBySystem, value:rows[0], onChange:v=>updateRow(0,v)}),
      e(SelectorRow, {index:1, systems, levelsBySystem:AppState.levelsBySystem, value:rows[1], onChange:v=>updateRow(1,v)}),
      e(SelectorRow, {index:2, systems, levelsBySystem:AppState.levelsBySystem, value:rows[2], onChange:v=>updateRow(2,v)}),
      e('div', {className:'btns'},
        e('button', {onClick:calculate, disabled:loading || !ready}, '开始计算'),
        e('button', {className:'secondary', onClick:()=>{setRows([{},{},{}]); setShowReport(false)}}, '重置选择'),
        e('button', {className:'secondary', onClick:exportPNG}, '导出报告 PNG')
      ),
      !ready ? e('div', {className:'sub'}, '正在加载 Excel 全量数据……') : null,
      showReport ? e(Report, {profile, selections:rows, ageBands:AppState.ageBands, basicBySystem:AppState.basicBySystem, levelsBySystem:AppState.levelsBySystem}) : null
    )
  );
}

/* Render */
const root = ReactDOM.createRoot(document.getElementById('app'));
root.render(React.createElement(App));
