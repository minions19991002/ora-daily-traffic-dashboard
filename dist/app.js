const TYPES = {
  mtStore: '美团门店数据', eleStore: '饿了么门店数据',
  mtPromo: '美团推广', elePromo: '饿了么推广', stores: '门店信息表'
};
const state = { files: {}, data: {}, mode: 'single', platform: 'mt', results: null };
const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];
const fmt = new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 0 });
const money = new Intl.NumberFormat('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pct = v => v == null || !Number.isFinite(v) ? '—' : `${(v * 100).toFixed(1)}%`;
const num = (v, d = 0) => v == null || !Number.isFinite(v) ? '—' : Number(v).toLocaleString('zh-CN', {minimumFractionDigits:d, maximumFractionDigits:d});
const clean = v => String(v ?? '').replace(/[\s　]+/g, '').replace(/[（）]/g, m => m === '（' ? '(' : ')').toLowerCase();
const keyDate = v => {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  if (typeof v === 'number') {
    if (v > 30000 && v < 70000) { const d = XLSX.SSF.parse_date_code(v); return d ? `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}` : null; }
    v = String(Math.trunc(v));
  }
  const s = String(v).trim();
  const m = s.match(/(20\d{2})[\/-]?(\d{1,2})[\/-]?(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}` : null;
};
const n = v => { const x = Number(String(v ?? '').replace(/,/g,'').replace(/%/g,'')); return Number.isFinite(x) ? x : 0; };
const id = v => String(v ?? '').replace(/\.0$/,'').trim();
const between = (d,a,b) => d && d >= a && d <= b;

function headerIndex(headers, aliases) {
  const hs = headers.map(clean);
  for (const a of aliases) { const i = hs.indexOf(clean(a)); if (i >= 0) return i; }
  return -1;
}
function val(row, headers, aliases) { const i = headerIndex(headers, aliases); return i < 0 ? null : row[i]; }

function detect(headers) {
  const h = headers.map(clean);
  const has = x => h.includes(clean(x));
  if (has('美团门店ID') && has('饿了么门店ID')) return 'stores';
  if (has('推广消费实付(元)')) return 'mtPromo';
  if (has('推广现金消费(元)')) return 'elePromo';
  if (has('门店id') && has('入店人数') && has('顾客实付')) return 'mtStore';
  if (has('门店编号') && has('进店人数') && has('顾客实付总额')) return 'eleStore';
  return null;
}

async function parseFile(file) {
  const book = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true });
  for (const name of book.SheetNames) {
    if (clean(name) === 'meta') continue;
    const rows = XLSX.utils.sheet_to_json(book.Sheets[name], { header:1, raw:true, defval:null });
    const first = rows.findIndex(r => r.some(v => v != null && String(v).trim() !== ''));
    if (first < 0) continue;
    const headers = rows[first].map(v => String(v ?? '').trim());
    const type = detect(headers);
    if (type) return { type, headers, rows: rows.slice(first+1).filter(r => r.some(v => v != null && String(v).trim() !== '')), sheet:name };
  }
  throw new Error('无法按表头识别文件类型');
}

async function ingest(files) {
  setMessage('正在读取文件…');
  const errors = [];
  for (const file of files) {
    try {
      const parsed = await parseFile(file);
      state.files[parsed.type] = file;
      state.data[parsed.type] = parsed;
    } catch (e) { errors.push(`${file.name}：${e.message}`); }
  }
  renderSlots(); updateDates();
  const missing = Object.keys(TYPES).filter(k => !state.data[k]);
  if (errors.length) setMessage(errors.join('；'), true);
  else if (missing.length) setMessage(`已识别 ${5-missing.length}/5 个文件，还缺少：${missing.map(k=>TYPES[k]).join('、')}`);
  else setMessage('5 个文件已全部识别，可以选择日期生成报表。', false, true);
}

function renderSlots() {
  $('#fileGrid').innerHTML = Object.entries(TYPES).map(([k,label]) => `<div class="file-slot ${state.files[k]?'ok':''}"><span class="dot"></span><div class="slot-title">${label}</div><div class="file-name">${state.files[k]?.name || '等待上传'}</div></div>`).join('');
  $('#generateBtn').disabled = Object.keys(TYPES).some(k => !state.data[k]);
}

function allDates() {
  const out = new Set();
  ['mtStore','eleStore'].forEach(t => {
    const d = state.data[t]; if (!d) return;
    d.rows.forEach(r => { const x = keyDate(val(r,d.headers,['日期'])); if (x) out.add(x); });
  });
  return [...out].sort();
}
function updateDates() {
  const dates = allDates(); if (!dates.length) return;
  const min = dates[0], max = dates.at(-1);
  ['startDate','endDate'].forEach(x => { $(`#${x}`).min=min; $(`#${x}`).max=max; $(`#${x}`).disabled=false; });
  $('#startDate').value=max; $('#endDate').value=max;
}
function setMessage(text='', error=false, success=false) { const el=$('#message'); el.textContent=text; el.className=`message${error?' error':''}${success?' success':''}`; }

function storeMaster() {
  const d=state.data.stores;
  return d.rows.map(r => ({
    name:String(val(r,d.headers,['店名','门店名称'])??'').trim(), code:id(val(r,d.headers,['店号'])),
    mt:id(val(r,d.headers,['美团门店ID'])), ele:id(val(r,d.headers,['饿了么门店ID']))
  })).filter(x => x.code && x.code.toUpperCase() !== 'OSZ01');
}
function latestRate(rows, headers, aliases, end) {
  const candidates=rows.map(r=>({d:keyDate(val(r,headers,['日期'])),v:val(r,headers,aliases)})).filter(x=>x.d&&x.d<=end&&x.v!==null&&x.v!=='').sort((a,b)=>a.d.localeCompare(b.d));
  if (!candidates.length) return null;
  let v=n(candidates.at(-1).v); if (v>1) v/=100; return v;
}
function promoAllowed(type,row,headers) {
  const text=['推广产品','计划名称','营销场景'].map(a=>String(val(row,headers,[a])??'')).join('|');
  const blocked=type==='mtPromo' ? ['津贴联盟','赏金联盟','流量助手','金字招牌','袋鼠店长','品牌装修','应用市场','短信通','拼好饭'] : ['增量助手'];
  return !blocked.some(x=>text.includes(x));
}
function promoFor(type, platformId, start, end) {
  const d=state.data[type], idAliases=['门店ID'];
  const rows=d.rows.filter(r=>id(val(r,d.headers,idAliases))===platformId && between(keyDate(val(r,d.headers,['日期'])),start,end) && promoAllowed(type,r,d.headers));
  const spendAlias=type==='mtPromo'?['推广消费实付(元)']:['推广现金消费(元)'];
  const spend=rows.reduce((s,r)=>s+n(val(r,d.headers,spendAlias)),0);
  const gmv=rows.reduce((s,r)=>s+n(val(r,d.headers,['订单交易额(元)'])),0);
  return {spend,gmv,roi:spend?gmv/spend:null};
}
function storeFor(platform, master, start, end) {
  const type=platform==='mt'?'mtStore':'eleStore', d=state.data[type], pid=master[platform];
  const idAliases=platform==='mt'?['门店id']:['门店编号'];
  const rows=d.rows.filter(r=>id(val(r,d.headers,idAliases))===pid && between(keyDate(val(r,d.headers,['日期'])),start,end));
  const exposure=rows.reduce((s,r)=>s+n(val(r,d.headers,['曝光人数'])),0);
  const visits=rows.reduce((s,r)=>s+n(val(r,d.headers,platform==='mt'?['入店人数']:['进店人数'])),0);
  const buyers=rows.reduce((s,r)=>s+n(val(r,d.headers,['下单人数'])),0);
  const orders=rows.reduce((s,r)=>s+n(val(r,d.headers,['有效订单'])),0);
  const pay=rows.reduce((s,r)=>s+n(val(r,d.headers,platform==='mt'?['顾客实付']:['顾客实付总额'])),0);
  const repeat=latestRate(rows,d.headers,platform==='mt'?['复购率']:['近30日复购率'],end);
  const promo=promoFor(platform==='mt'?'mtPromo':'elePromo',pid,start,end);
  return {code:master.code,name:master.name,pid,has:rows.length>0,exposure,visits,buyers,orders,pay,repeat,...promo};
}
function totalRows(rows) {
  const has=rows.filter(r=>r.has), sum=k=>has.reduce((s,r)=>s+r[k],0);
  const exposure=sum('exposure'), visits=sum('visits'), buyers=sum('buyers'), orders=sum('orders'), pay=sum('pay'), spend=rows.reduce((s,r)=>s+r.spend,0), gmv=rows.reduce((s,r)=>s+r.gmv,0);
  const repeatDen=has.reduce((s,r)=>s+(r.repeat==null?0:r.orders),0);
  const repeatNum=has.reduce((s,r)=>s+(r.repeat==null?0:r.repeat*r.orders),0);
  return {code:'全国',name:'',has:has.length>0,exposure,visits,buyers,orders,pay,repeat:repeatDen?repeatNum/repeatDen:null,spend,gmv,roi:spend?gmv/spend:null};
}
function buildResults(start,end) {
  const stores=storeMaster();
  const mt=stores.map(s=>storeFor('mt',s,start,end)); const ele=stores.map(s=>storeFor('ele',s,start,end));
  mt.push(totalRows(mt)); ele.push(totalRows(ele));
  return {start,end,mt,ele,storeCount:stores.length};
}
function viewRow(r, period) {
  return [period,r.code,r.name,r.has?fmt.format(Math.round(r.exposure)):'—',r.has&&r.exposure?pct(r.visits/r.exposure):'—',r.has&&r.visits?pct(r.buyers/r.visits):'—',r.has?fmt.format(Math.round(r.orders)):'—',r.has&&r.orders?money.format(r.pay/r.orders):'—',r.has?pct(r.repeat):'—',fmt.format(Math.round(r.spend)),r.roi==null?'—':r.roi.toFixed(1)];
}
function renderReport() {
  const r=state.results, rows=r[state.platform], total=rows.at(-1), label=r.start===r.end?r.start:`${r.start} 至 ${r.end}`;
  $('#reportTitle').textContent=`ORA 流量日报 · ${label}`; $('#reportMeta').textContent=`${r.storeCount} 家门店（已剔除 OSZ01） · ${state.platform==='mt'?'美团':'淘闪'}`;
  $('#summaryCards').innerHTML=[['曝光人数',fmt.format(Math.round(total.exposure))],['进店率',total.exposure?pct(total.visits/total.exposure):'—'],['有效订单',fmt.format(Math.round(total.orders))],['推广实付 ROI',total.roi==null?'—':total.roi.toFixed(1)]].map(x=>`<div class="metric"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join('');
  const headers=['日期','店号','店名','曝光人数','进店率','下单转化率','客单量','客单价','复购率','推广消耗','推广ROI'];
  $('#reportTable').innerHTML=`<thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map((x,i)=>`<tr class="${i===rows.length-1?'total':''}">${viewRow(x,label).map(v=>`<td class="${v==='—'?'na':''}">${v}</td>`).join('')}</tr>`).join('')}</tbody>`;
  $('#methodText').innerHTML='进店率 = 进店人数 ÷ 曝光人数；下单转化率 = 下单人数 ÷ 进店人数；客单量 = 有效订单；客单价 = 顾客实付 ÷ 有效订单；复购率取所选周期最后一天；推广 ROI = 订单交易额 ÷ 推广实付消耗。总计均由明细分子、分母重新汇总。多日指标不平均每日比例。';
  $('#report').classList.remove('hidden'); $('#report').scrollIntoView({behavior:'smooth',block:'start'});
}
const excelSerial = iso => {
  const [y,m,d]=iso.split('-').map(Number);
  return Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/86400000);
};
function worksheetCell(doc,address) {
  return [...doc.getElementsByTagName('c')].find(cell=>cell.getAttribute('r')===address);
}
function replaceCellValue(doc,address,value,type='number') {
  const cell=worksheetCell(doc,address);
  if (!cell) throw new Error(`母版缺少单元格 ${address}`);
  [...cell.childNodes].forEach(node=>cell.removeChild(node));
  cell.removeAttribute('t');
  if (value==null || value==='' || (typeof value==='number' && !Number.isFinite(value))) return;
  const ns=doc.documentElement.namespaceURI;
  if (type==='text') {
    cell.setAttribute('t','inlineStr');
    const is=doc.createElementNS(ns,'is'), text=doc.createElementNS(ns,'t');
    text.textContent=String(value); is.appendChild(text); cell.appendChild(is);
  } else {
    const node=doc.createElementNS(ns,'v'); node.textContent=String(value); cell.appendChild(node);
  }
}
function rowAt(doc,rowNumber) {
  return [...doc.getElementsByTagName('row')].find(row=>Number(row.getAttribute('r'))===rowNumber);
}
function shiftA1Rows(reference,startRow,delta) {
  return String(reference).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(match,column,row)=>`${column}${Number(row)>=startRow?Number(row)+delta:row}`);
}
function shiftWorksheetRows(doc,startRow,delta) {
  [...doc.getElementsByTagName('row')].filter(row=>Number(row.getAttribute('r'))>=startRow).forEach(row=>{
    const next=Number(row.getAttribute('r'))+delta;
    row.setAttribute('r',next);
    [...row.getElementsByTagName('c')].forEach(cell=>cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,next)));
  });
  const dimension=doc.getElementsByTagName('dimension')[0];
  if(dimension) dimension.setAttribute('ref',shiftA1Rows(dimension.getAttribute('ref'),startRow,delta));
  [['mergeCell','ref'],['conditionalFormatting','sqref'],['dataValidation','sqref'],['autoFilter','ref'],['hyperlink','ref']].forEach(([tag,attribute])=>{
    [...doc.getElementsByTagName(tag)].forEach(node=>{
      if(node.hasAttribute(attribute)) node.setAttribute(attribute,shiftA1Rows(node.getAttribute(attribute),startRow,delta));
    });
  });
}
function setRowNumber(row,rowNumber) {
  row.setAttribute('r',rowNumber);
  [...row.getElementsByTagName('c')].forEach(cell=>cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,rowNumber)));
}
function insertStoreRows(doc,totalRow,stores,sourceRow) {
  if(!stores.length) return totalRow;
  const source=rowAt(doc,sourceRow).cloneNode(true);
  shiftWorksheetRows(doc,totalRow,stores.length);
  const shiftedTotal=rowAt(doc,totalRow+stores.length);
  stores.forEach((store,index)=>{
    const row=source.cloneNode(true), rowNumber=totalRow+index;
    setRowNumber(row,rowNumber);
    shiftedTotal.parentNode.insertBefore(row,shiftedTotal);
    replaceCellValue(doc,`B${rowNumber}`,store.code,'text');
    replaceCellValue(doc,`C${rowNumber}`,store.name,'text');
  });
  return totalRow+stores.length;
}
function prepareExportStyles(stylesDoc) {
  const ns=stylesDoc.documentElement.namespaceURI;
  const numFmts=stylesDoc.getElementsByTagName('numFmts')[0];
  let oneDecimal=[...numFmts.getElementsByTagName('numFmt')].find(node=>node.getAttribute('formatCode')==='0.0');
  if(!oneDecimal) {
    oneDecimal=stylesDoc.createElementNS(ns,'numFmt'); oneDecimal.setAttribute('numFmtId','177'); oneDecimal.setAttribute('formatCode','0.0');
    numFmts.appendChild(oneDecimal); numFmts.setAttribute('count',numFmts.children.length);
  }
  const oneDecimalId=oneDecimal.getAttribute('numFmtId');
  const fonts=stylesDoc.getElementsByTagName('fonts')[0], boldFont=fonts.children[3].cloneNode(true);
  if(!boldFont.getElementsByTagName('b').length) boldFont.insertBefore(stylesDoc.createElementNS(ns,'b'),boldFont.firstChild);
  fonts.appendChild(boldFont); fonts.setAttribute('count',fonts.children.length);
  const boldFontId=fonts.children.length-1, cellXfs=stylesDoc.getElementsByTagName('cellXfs')[0];
  const addStyle=(baseIndex,{numFmtId=null,fontId=null}={})=>{
    const xf=cellXfs.children[baseIndex].cloneNode(true);
    if(numFmtId!=null) { xf.setAttribute('numFmtId',numFmtId); xf.setAttribute('applyNumberFormat','1'); }
    if(fontId!=null) { xf.setAttribute('fontId',fontId); xf.setAttribute('applyFont','1'); }
    cellXfs.appendChild(xf); cellXfs.setAttribute('count',cellXfs.children.length);
    return cellXfs.children.length-1;
  };
  const totalText=addStyle(17,{fontId:boldFontId}), totalGeneral=addStyle(5,{fontId:boldFontId});
  return {
    detail:{D:addStyle(41,{numFmtId:'3'}),G:addStyle(42,{numFmtId:'3'}),J:addStyle(43,{numFmtId:'3'}),K:addStyle(44,{numFmtId:oneDecimalId})},
    total:{A:addStyle(16,{fontId:boldFontId}),B:totalText,C:totalText,D:addStyle(4,{numFmtId:'3',fontId:boldFontId}),E:totalGeneral,F:totalGeneral,G:addStyle(5,{numFmtId:'3',fontId:boldFontId}),H:totalGeneral,I:totalGeneral,J:addStyle(5,{numFmtId:'3',fontId:boldFontId}),K:addStyle(6,{numFmtId:oneDecimalId,fontId:boldFontId})}
  };
}
function applyRowStyles(doc,rowNumber,styles,isTotal=false) {
  const mapping=isTotal?styles.total:styles.detail;
  Object.entries(mapping).forEach(([column,style])=>worksheetCell(doc,`${column}${rowNumber}`).setAttribute('s',style));
}
function centerWorksheetCells(doc,stylesDoc) {
  const ns=stylesDoc.documentElement.namespaceURI, cellXfs=stylesDoc.getElementsByTagName('cellXfs')[0], centered=new Map();
  [...doc.getElementsByTagName('c')].forEach(cell=>{
    const baseIndex=Number(cell.getAttribute('s')||0);
    if(!centered.has(baseIndex)) {
      const xf=cellXfs.children[baseIndex].cloneNode(true);
      let alignment=xf.getElementsByTagName('alignment')[0];
      if(!alignment) { alignment=stylesDoc.createElementNS(ns,'alignment'); xf.appendChild(alignment); }
      alignment.setAttribute('horizontal','center'); xf.setAttribute('applyAlignment','1');
      cellXfs.appendChild(xf); centered.set(baseIndex,cellXfs.children.length-1);
    }
    cell.setAttribute('s',centered.get(baseIndex));
  });
  cellXfs.setAttribute('count',cellXfs.children.length);
}
function templateMetricValues(row) {
  return [
    row.has?Math.round(row.exposure):null,
    row.has&&row.exposure?row.visits/row.exposure:null,
    row.has&&row.visits?row.buyers/row.visits:null,
    row.has?Math.round(row.orders):null,
    row.has&&row.orders?row.pay/row.orders:null,
    row.has?row.repeat:null,
    Math.round(row.spend),
    row.roi
  ];
}
function fillTemplateBlock(doc,platform,startRow,codes,totalRow,dateValue,dateType,styles) {
  const byCode=new Map(state.results[platform].map(row=>[row.code,row]));
  codes.forEach((code,index)=>{
    const rowNumber=startRow+index;
    const result=byCode.get(code);
    replaceCellValue(doc,`A${rowNumber}`,dateValue,dateType);
    const values=result?templateMetricValues(result):Array(8).fill(null);
    'DEFGHIJK'.split('').forEach((column,index)=>replaceCellValue(doc,`${column}${rowNumber}`,values[index]));
    applyRowStyles(doc,rowNumber,styles);
  });
  const total=byCode.get('全国');
  replaceCellValue(doc,`A${totalRow}`,dateValue,dateType);
  'DEFGHIJK'.split('').forEach((column,index)=>replaceCellValue(doc,`${column}${totalRow}`,templateMetricValues(total)[index]));
  applyRowStyles(doc,totalRow,styles,true);
}
async function downloadExcel() {
  const button=$('#downloadBtn'); button.disabled=true; button.textContent='正在生成…';
  try {
    const response=await fetch('Ora日报-流量数据 T+1.xlsx');
    if(!response.ok) throw new Error('无法读取报表母版');
    const templateBytes=await response.arrayBuffer();
    const templateBook=XLSX.read(templateBytes,{type:'array',cellDates:false});
    const templateSheet=templateBook.Sheets[templateBook.SheetNames[0]];
    const zip=await JSZip.loadAsync(templateBytes);
    const sheetPath='xl/worksheets/sheet1.xml', stylesPath='xl/styles.xml';
    const xml=await zip.file(sheetPath).async('string'), stylesXml=await zip.file(stylesPath).async('string');
    const doc=new DOMParser().parseFromString(xml,'application/xml');
    const stylesDoc=new DOMParser().parseFromString(stylesXml,'application/xml');
    if(doc.getElementsByTagName('parsererror').length||stylesDoc.getElementsByTagName('parsererror').length) throw new Error('母版结构读取失败');
    const r=state.results, single=r.start===r.end;
    const dateValue=single?excelSerial(r.start):`${r.start.replaceAll('-','/')}至${r.end.replaceAll('-','/')}`;
    const dateType=single?'number':'text';
    const styles=prepareExportStyles(stylesDoc);
    const originalCodes=(start,end)=>Array.from({length:end-start+1},(_,index)=>String(templateSheet[`B${start+index}`]?.v??'').trim()).filter(Boolean);
    const mtCodes=originalCodes(4,26), eleCodes=originalCodes(31,53);
    const newRows=(platform,codes)=>state.results[platform].filter(row=>row.code!=='全国'&&!codes.includes(row.code));
    const mtNew=newRows('mt',mtCodes), mtTotal=insertStoreRows(doc,27,mtNew,5), mtAll=[...mtCodes,...mtNew.map(row=>row.code)];
    const eleStart=31+mtNew.length, eleNew=newRows('ele',eleCodes), eleTotal=insertStoreRows(doc,54+mtNew.length,eleNew,eleStart+1), eleAll=[...eleCodes,...eleNew.map(row=>row.code)];
    fillTemplateBlock(doc,'mt',4,mtAll,mtTotal,dateValue,dateType,styles);
    fillTemplateBlock(doc,'ele',eleStart,eleAll,eleTotal,dateValue,dateType,styles);
    centerWorksheetCells(doc,stylesDoc);
    zip.file(sheetPath,new XMLSerializer().serializeToString(doc));
    zip.file(stylesPath,new XMLSerializer().serializeToString(stylesDoc));
    const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
    const link=document.createElement('a');
    const objectUrl=URL.createObjectURL(blob); link.href=objectUrl;
    link.download=`Ora日报-流量数据_${r.start}_${r.end}.xlsx`;
    document.body.appendChild(link); link.click(); link.remove();
    setTimeout(()=>URL.revokeObjectURL(objectUrl),1000);
    setMessage('已按 Ora日报-流量数据 T+1 母版格式导出。',false,true);
  } catch(e) { setMessage(`导出失败：${e.message}`,true); }
  finally { button.disabled=false; button.textContent='下载 Excel'; }
}

$('#fileInput').addEventListener('change',e=>ingest([...e.target.files]));
const dz=$('#dropzone'); ['dragenter','dragover'].forEach(x=>dz.addEventListener(x,e=>{e.preventDefault();dz.classList.add('drag')})); ['dragleave','drop'].forEach(x=>dz.addEventListener(x,e=>{e.preventDefault();dz.classList.remove('drag')})); dz.addEventListener('drop',e=>ingest([...e.dataTransfer.files].filter(f=>/\.xlsx?$/i.test(f.name))));
$('#clearBtn').addEventListener('click',()=>{state.files={};state.data={};state.results=null;$('#fileInput').value='';$('#report').classList.add('hidden');['startDate','endDate'].forEach(x=>{$(`#${x}`).value='';$(`#${x}`).disabled=true});renderSlots();setMessage('已清空。')});
$$('.segmented button').forEach(b=>b.addEventListener('click',()=>{$$('.segmented button').forEach(x=>x.classList.remove('active'));b.classList.add('active');state.mode=b.dataset.mode;$('#endWrap').classList.toggle('hidden',state.mode==='single')}));
$('#generateBtn').addEventListener('click',()=>{const start=$('#startDate').value,end=state.mode==='single'?start:$('#endDate').value;if(!start||!end)return setMessage('请选择报表日期。',true);if(start>end)return setMessage('开始日期不能晚于结束日期。',true);try{state.results=buildResults(start,end);renderReport();setMessage('报表已生成。',false,true)}catch(e){setMessage(`生成失败：${e.message}`,true)}});
$$('.tabs button').forEach(b=>b.addEventListener('click',()=>{$$('.tabs button').forEach(x=>x.classList.remove('active'));b.classList.add('active');state.platform=b.dataset.platform;renderReport()}));
$('#downloadBtn').addEventListener('click',downloadExcel);
$('#copyBtn').addEventListener('click',async()=>{try{await navigator.clipboard.writeText($('#reportTable').innerText);$('#copyBtn').textContent='已复制';setTimeout(()=>$('#copyBtn').textContent='复制表格',1400)}catch{setMessage('浏览器未允许复制，请手动选择表格。',true)}});
renderSlots();
