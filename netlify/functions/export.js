const ExcelJS = require("exceljs");

const SERVICE_URL = "https://roszdravnadzor.gov.ru/services/turnover";
const AJAX_URL = "https://roszdravnadzor.gov.ru/ajax/services/turnover";
const PAGE_SIZE = 500;

const COLUMN_NAMES = [
  "Поисковый запрос", "Тип ЛС", "Дата внесения в АИС Росздравнадзора", "№ РУ", "Дата РУ",
  "Торговое наименование", "Производитель (выпускающий контроль)", "Страна",
  "Сведения о стадиях производства", "Нормативная документация",
  "Организация, выпустившая в гражданский оборот", "№ серии, партии",
  "№ разрешения, дата разрешения Росздравнадзора на ИЛП",
  "Письма, приказы об изъятии ЛС (номер, дата)"
];

let diagnosticLog = [];
function resetDiagnosticLog(){ diagnosticLog = []; }
function log(message, data){
  const stamp = new Date().toISOString();
  let line = `[${stamp}] ${message}`;
  if (data !== undefined) {
    let serialized;
    try { serialized = JSON.stringify(data, null, 2); } catch { serialized = String(data); }
    line += `\n${serialized}`;
  }
  diagnosticLog.push(line);
  console.log(line);
}
function getDiagnosticLog(){ return diagnosticLog.join("\n"); }
function truncate(value, maxLength=3000){
  value = String(value ?? "");
  return value.length <= maxLength ? value : value.slice(0,maxLength) + "... [TRUNCATED]";
}
function cleanHtml(value){
  if (value == null) return "";
  return String(value)
    .replace(/<\s*br\s*\/?\s*>/gi,"\n")
    .replace(/<\/\s*(div|p|li)\s*>/gi,"\n")
    .replace(/<[^>]+>/g,"")
    .replace(/&nbsp;|&#160;/g," ")
    .replace(/&amp;/g,"&").replace(/&quot;/g,'"').replace(/&#039;|&#39;/g,"'")
    .replace(/&lt;/g,"<").replace(/&gt;/g,">")
    .replace(/[ \t]+/g," ").replace(/\n[ \t]+/g,"\n").replace(/\n{3,}/g,"\n\n").trim();
}
function extractCell(record,n){
  const cell = record[`col${n}`] ?? "";
  if (!cell || typeof cell !== "object") return cleanHtml(cell);
  const title = cleanHtml(cell.title ?? "");
  return title || cleanHtml(cell.label ?? "");
}
function buildPayload({tradeName,dateFrom,dateTo,start,draw,previousTotal}){
  const p = new URLSearchParams({
    draw:String(draw), start:String(start), length:String(PAGE_SIZE),
    "search[value]":"", "search[regex]":"false", prev_total:String(previousTotal),
    q_label:"", q_series:"", q_type_ls:"2", q_mnn:"", q_dt_ru_from:"", q_dt_ru_to:"",
    q_no_ru:"", q_country:"", q_tn:tradeName, dt_from:dateFrom, dt_to:dateTo,
    q_producer:"", q_inf_letter:"0", q_org:""
  });
  for(let i=0;i<13;i++){
    p.set(`columns[${i}][data]`,`col${i+1}.label`); p.set(`columns[${i}][name]`,"");
    p.set(`columns[${i}][searchable]`,"true"); p.set(`columns[${i}][orderable]`,"true");
    p.set(`columns[${i}][search][value]`,""); p.set(`columns[${i}][search][regex]`,"false");
  }
  return p;
}
function extractCookies(response){
  let setCookies=[];
  if(typeof response.headers.getSetCookie === "function") setCookies=response.headers.getSetCookie();
  else { const raw=response.headers.get("set-cookie"); if(raw) setCookies=[raw]; }
  const pairs=setCookies.map(c=>c.split(";")[0].trim()).filter(Boolean);
  return {count:setCookies.length,names:pairs.map(x=>x.split("=")[0]),header:pairs.join("; ")};
}
async function openService(){
  log("Opening Roszdravnadzor service",{url:SERVICE_URL});
  const response=await fetch(SERVICE_URL,{method:"GET",redirect:"follow",headers:{
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language":"ru-RU,ru;q=0.9,en;q=0.8","Cache-Control":"no-cache","Pragma":"no-cache"
  }});
  const text=await response.text();
  log("Initial GET completed",{status:response.status,final_url:response.url,content_type:response.headers.get("content-type"),bytes:text.length});
  if(!response.ok){ log("Initial GET failed",{body:truncate(text)}); throw new Error(`Первичный GET Росздравнадзора: HTTP ${response.status}`); }
  const cookies=extractCookies(response); log("Cookies received",{count:cookies.count,names:cookies.names});
  log("Initial HTML received",{has_turnover_form:text.includes("id_q_tn")});
  return cookies.header;
}
async function requestPage(args){
  const payload=buildPayload(args);
  const headers={
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept":"application/json, text/javascript, */*; q=0.01","Accept-Language":"ru-RU,ru;q=0.9,en;q=0.8",
    "Content-Type":"application/x-www-form-urlencoded; charset=UTF-8","Origin":"https://roszdravnadzor.gov.ru",
    "Referer":SERVICE_URL,"X-Requested-With":"XMLHttpRequest","Cache-Control":"no-cache","Pragma":"no-cache"
  };
  if(args.cookieHeader) headers.Cookie=args.cookieHeader;
  log("Roszdravnadzor AJAX request",{trade_name:args.tradeName,date_from:args.dateFrom,date_to:args.dateTo,start:args.start,draw:args.draw,prev_total:args.previousTotal,cookie_attached:Boolean(args.cookieHeader)});
  const response=await fetch(AJAX_URL,{method:"POST",headers,body:payload.toString(),redirect:"follow"});
  const text=await response.text();
  log("Roszdravnadzor AJAX response",{status:response.status,content_type:response.headers.get("content-type"),content_length:text.length,start:args.start,draw:args.draw});
  if(!response.ok){ log("Roszdravnadzor AJAX HTTP ERROR",{status:response.status,response_body:truncate(text)}); throw new Error(`Росздравнадзор AJAX: HTTP ${response.status}`); }
  let json; try{json=JSON.parse(text);}catch{log("AJAX response is not JSON",{body:truncate(text)});throw new Error("Росздравнадзор вернул AJAX-ответ не в формате JSON");}
  if(json.message) throw new Error(`Росздравнадзор: ${json.message}`);
  return json;
}
async function loadTradeName(tradeName,dateFrom,dateTo){
  log("Starting trade name",{trade_name:tradeName});
  const cookieHeader=await openService();
  let all=[],start=0,draw=1,total=null,previousTotal=0;
  while(true){
    const json=await requestPage({tradeName,dateFrom,dateTo,start,draw,previousTotal,cookieHeader});
    if(total===null){
      total=Number(json.recordsFiltered ?? json.recordsTotal ?? 0);
      previousTotal=Number(json.recordsTotal ?? total);
      log("Search initialized",{trade_name:tradeName,records_filtered:json.recordsFiltered,records_total:json.recordsTotal,expected:total});
    }
    const batch=json.data || [];
    log("Batch received",{trade_name:tradeName,start,batch:batch.length,accumulated:all.length+batch.length,expected:total});
    if(!batch.length) break;
    all.push(...batch); start+=batch.length; draw+=1;
    if(all.length>=total) break;
  }
  if(all.length!==total) throw new Error(`${tradeName}: загружено ${all.length} из ${total}`);
  log("Trade name completed",{trade_name:tradeName,records:all.length});
  return { tradeName, count: all.length, rows: all.map(r=>[tradeName,...Array.from({length:13},(_,i)=>extractCell(r,i+1))]) };
}
function safeFilename(names){
  const base=names.join("_").replace(/[\\/:*?"<>|]+/g,"_").replace(/\s+/g," ").slice(0,120) || "Результат";
  return `${base}_Росздравнадзор_ЛП.xlsx`;
}
async function createWorkbook(rows){
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("Результаты",{views:[{state:"frozen",ySplit:1}]});
  ws.addRow(COLUMN_NAMES); rows.forEach(r=>ws.addRow(r)); ws.autoFilter={from:"A1",to:`N${Math.max(1,rows.length+1)}`};
  const h=ws.getRow(1); h.font={bold:true};h.alignment={horizontal:"center",vertical:"middle",wrapText:true};h.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFD9EAF7"}};
  [22,12,20,20,13,28,38,18,48,40,42,22,32,34].forEach((w,i)=>ws.getColumn(i+1).width=w);
  ws.eachRow((row,n)=>{if(n>1)row.alignment={vertical:"top",wrapText:true};});
  return wb.xlsx.writeBuffer();
}
exports.handler=async function(event){
  resetDiagnosticLog();
  const invocationId=Math.random().toString(36).slice(2,10);
  log("FUNCTION START",{invocation_id:invocationId,method:event.httpMethod,node:process.version,region:process.env.AWS_REGION||null});
  if(event.httpMethod!=="POST") return {statusCode:405,headers:{"Content-Type":"application/json; charset=utf-8"},body:JSON.stringify({error:"Method not allowed"})};
  try{
    const body=JSON.parse(event.body||"{}");
    const names=[...new Set((body.trade_names||[]).map(v=>String(v).trim()).filter(Boolean))];
    if(!names.length) throw new Error("Укажите минимум одно торговое наименование");
    if(!body.date_from||!body.date_to) throw new Error("Укажите обе даты");
    log("Request accepted",{invocation_id:invocationId,trade_names:names,date_from:body.date_from,date_to:body.date_to});

    const searchResults=[]; let rows=[];
    for(const name of names){
      const result=await loadTradeName(name,body.date_from,body.date_to);
      searchResults.push({trade_name:result.tradeName,count:result.count});
      if(result.count>0) rows.push(...result.rows);
    }
    log("Search summary",{search_results:searchResults,total_rows:rows.length});

    if(rows.length===0){
      log("NO RESULTS FOR ALL TRADE NAMES",{search_results:searchResults});
      return {statusCode:404,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"},body:JSON.stringify({no_results:true,message:"По заданным параметрам сведения не найдены.",search_results:searchResults,invocation_id:invocationId})};
    }

    const seen=new Set(); rows=rows.filter(row=>{const key=JSON.stringify(row);if(seen.has(key))return false;seen.add(key);return true;});
    log("Creating workbook",{rows:rows.length,search_results:searchResults});
    const buffer=await createWorkbook(rows);
    const foundNames=searchResults.filter(x=>x.count>0).map(x=>x.trade_name);
    const filename=safeFilename(foundNames);
    log("FUNCTION SUCCESS",{invocation_id:invocationId,filename,rows:rows.length,bytes:buffer.length,search_results:searchResults});
    return {statusCode:200,isBase64Encoded:true,headers:{
      "Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition":`attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
      "X-Filename":encodeURIComponent(filename),
      "X-Search-Results":encodeURIComponent(JSON.stringify(searchResults)),
      "Cache-Control":"no-store"
    },body:Buffer.from(buffer).toString("base64")};
  }catch(error){
    log("FUNCTION ERROR",{invocation_id:invocationId,name:error.name,message:error.message,stack:error.stack});
    return {statusCode:500,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"},body:JSON.stringify({error:error.message||String(error),invocation_id:invocationId,diagnostic_log:getDiagnosticLog()})};
  }
};
