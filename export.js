const ExcelJS = require('exceljs');

const SERVICE_URL = 'https://roszdravnadzor.gov.ru/services/turnover';
const AJAX_URL = 'https://roszdravnadzor.gov.ru/ajax/services/turnover';
const PAGE_SIZE = 500;
const COLUMNS = [
  'Поисковый запрос','Тип ЛС','Дата внесения в АИС Росздравнадзора','№ РУ','Дата РУ',
  'Торговое наименование','Производитель (выпускающий контроль)','Страна',
  'Сведения о стадиях производства','Нормативная документация',
  'Организация, выпустившая в гражданский оборот','№ серии, партии',
  '№ разрешения, дата разрешения Росздравнадзора на ИЛП',
  'Письма, приказы об изъятии ЛС (номер, дата)'
];

function cleanHtml(value) {
  if (value == null) return '';
  return String(value)
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/\s*(div|p|li)\s*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;|&#160;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#039;|&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/[ \t]+/g, ' ').replace(/\n[ \t]+/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}
function extractCell(record, n) {
  const cell = record[`col${n}`] ?? '';
  if (!cell || typeof cell !== 'object') return cleanHtml(cell);
  const title = cleanHtml(cell.title ?? '');
  return title || cleanHtml(cell.label ?? '');
}
function payload(name, dateFrom, dateTo, start, draw, previousTotal) {
  const p = new URLSearchParams({
    draw:String(draw), start:String(start), length:String(PAGE_SIZE),
    'search[value]':'', 'search[regex]':'false', prev_total:String(previousTotal),
    q_label:'', q_series:'', q_type_ls:'2', q_mnn:'', q_dt_ru_from:'', q_dt_ru_to:'',
    q_no_ru:'', q_country:'', q_tn:name, dt_from:dateFrom, dt_to:dateTo,
    q_producer:'', q_inf_letter:'0', q_org:''
  });
  for (let i=0;i<13;i++) {
    p.set(`columns[${i}][data]`,`col${i+1}.label`); p.set(`columns[${i}][name]`, '');
    p.set(`columns[${i}][searchable]`,'true'); p.set(`columns[${i}][orderable]`,'true');
    p.set(`columns[${i}][search][value]`,''); p.set(`columns[${i}][search][regex]`,'false');
  }
  return p;
}
async function fetchJson(url, options, timeoutMs=55000) {
  const controller = new AbortController(); const timer = setTimeout(()=>controller.abort(), timeoutMs);
  try { const r = await fetch(url, {...options, signal:controller.signal});
    if (!r.ok) throw new Error(`HTTP ${r.status} от Росздравнадзора`);
    const text = await r.text();
    try { return {json:JSON.parse(text), response:r}; } catch { throw new Error('Росздравнадзор вернул не JSON'); }
  } finally { clearTimeout(timer); }
}
async function loadName(name, dateFrom, dateTo) {
  const home = await fetch(SERVICE_URL, {headers:{'User-Agent':'Mozilla/5.0','Accept-Language':'ru-RU,ru;q=0.9'}});
  if (!home.ok) throw new Error(`Не удалось открыть сервис: HTTP ${home.status}`);
  const cookie = home.headers.get('set-cookie') || '';
  let all=[], start=0, draw=1, total=null, previousTotal=0;
  while (true) {
    const {json} = await fetchJson(AJAX_URL, {method:'POST',headers:{
      'User-Agent':'Mozilla/5.0','Accept-Language':'ru-RU,ru;q=0.9','Accept':'application/json, text/javascript, */*; q=0.01',
      'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8','Origin':'https://roszdravnadzor.gov.ru',
      'Referer':SERVICE_URL,'X-Requested-With':'XMLHttpRequest','Cookie':cookie
    }, body:payload(name,dateFrom,dateTo,start,draw,previousTotal).toString()});
    if (json.message) throw new Error(json.message);
    if (total === null) { total = Number(json.recordsFiltered ?? json.recordsTotal ?? 0); previousTotal = Number(json.recordsTotal ?? total); }
    const batch = json.data || []; if (!batch.length) break;
    all.push(...batch); start += batch.length; draw += 1;
    if (all.length >= total) break;
  }
  if (all.length !== total) throw new Error(`${name}: загружено ${all.length} из ${total}`);
  return all.map(r => [name, ...Array.from({length:13},(_,i)=>extractCell(r,i+1))]);
}
function safeFilename(names) {
  const base = names.join('_').replace(/[\\/:*?"<>|]+/g,'_').replace(/\s+/g,' ').slice(0,120) || 'Результат';
  return `${base}_Росздравнадзор_ЛП.xlsx`;
}
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return {statusCode:405, body:JSON.stringify({error:'Method not allowed'})};
  try {
    const body = JSON.parse(event.body || '{}');
    const names = [...new Set((body.trade_names || []).map(x=>String(x).trim()).filter(Boolean))];
    if (!names.length) throw new Error('Укажите минимум одно торговое наименование');
    if (!body.date_from || !body.date_to) throw new Error('Укажите обе даты');
    let rows=[]; for (const name of names) rows.push(...await loadName(name,body.date_from,body.date_to));
    const seen=new Set(); rows=rows.filter(r=>{const k=JSON.stringify(r);if(seen.has(k))return false;seen.add(k);return true;});
    const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('Результаты',{views:[{state:'frozen',ySplit:1}]});
    ws.addRow(COLUMNS); rows.forEach(r=>ws.addRow(r)); ws.autoFilter={from:'A1',to:`N${Math.max(1,rows.length+1)}`};
    ws.getRow(1).font={bold:true}; ws.getRow(1).alignment={horizontal:'center',vertical:'middle',wrapText:true};
    ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9EAF7'}};
    [20,12,20,20,13,28,38,18,48,40,42,22,32,34].forEach((w,i)=>ws.getColumn(i+1).width=w);
    ws.eachRow((row,n)=>{if(n>1)row.alignment={vertical:'top',wrapText:true};});
    const buffer=await wb.xlsx.writeBuffer(); const filename=safeFilename(names);
    return {statusCode:200,isBase64Encoded:true,headers:{
      'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition':`attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
      'X-Filename':encodeURIComponent(filename),'Cache-Control':'no-store'
    },body:Buffer.from(buffer).toString('base64')};
  } catch(e) { return {statusCode:400,headers:{'Content-Type':'application/json; charset=utf-8'},body:JSON.stringify({error:e.message||String(e)})}; }
};
