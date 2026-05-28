const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const genAI = new GoogleGenerativeAI(process.env.GEMINI_KEY);
app.use(express.json());

const usage = {};
const getU = ip => usage[ip] || 0;
const incU = ip => { usage[ip] = (usage[ip] || 0) + 1; };

app.get('/usage',(req,res)=>res.json({remaining:Math.max(0,5-getU(req.ip))}));

app.get('/',(req,res)=>res.send(`<!DOCTYPE html><html><head><title>BankPDF</title><style>*{margin:0;padding:0;box-sizing:border-box;}body{font-family:Arial;background:#0a0f1e;color:white;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:20px;}.box{background:#0d1526;border:1px solid #1e3a5f;border-radius:16px;padding:40px;width:100%;max-width:560px;text-align:center;}h1{color:#38bdf8;font-size:2rem;margin-bottom:8px;}p{color:#94a3b8;margin-bottom:24px;}.tabs{display:flex;gap:8px;margin-bottom:24px;justify-content:center;flex-wrap:wrap;}.tab{padding:8px 18px;border-radius:20px;border:1px solid #1e3a5f;background:transparent;color:#94a3b8;cursor:pointer;font-size:13px;}.tab.active{background:#38bdf8;color:#0a0f1e;font-weight:bold;border-color:#38bdf8;}.panel{display:none;}.panel.active{display:block;}.drop{border:2px dashed #1e3a5f;border-radius:12px;padding:40px;cursor:pointer;color:#38bdf8;margin-bottom:12px;transition:all .2s;}.drop:hover,.drop.over{background:#0a1929;border-color:#38bdf8;}.fname{color:#22c55e;font-size:13px;margin-bottom:12px;min-height:18px;}.btn{width:100%;padding:14px;border-radius:10px;border:none;font-size:1rem;font-weight:bold;cursor:pointer;background:#38bdf8;color:#0a0f1e;margin-bottom:8px;}.result{margin-top:20px;overflow-x:auto;font-size:13px;}.ai-card{background:#0a1929;border:1px solid #1e3a5f;border-radius:8px;padding:12px;margin-bottom:8px;text-align:left;}.ai-label{font-size:10px;color:#94a3b8;text-transform:uppercase;margin-bottom:3px;}.ai-val{color:#38bdf8;font-weight:500;}table{width:100%;border-collapse:collapse;color:white;}th{background:#1e3a5f;padding:8px;font-size:12px;text-align:left;}td{padding:7px;border-bottom:1px solid #1e3a5f;font-size:12px;}.usage-badge{position:fixed;top:16px;right:16px;background:#0d1526;border:1px solid #1e3a5f;padding:6px 14px;border-radius:20px;font-size:12px;color:#94a3b8;}</style></head><body>
<div class="usage-badge">Free uses: <span id="uc">5</span> remaining</div>
<div class="box">
<h1>🏦 BankPDF</h1>
<p>Convert and analyze bank statements instantly</p>
<div class="tabs">
<button class="tab active" onclick="sw('t1',this)">📊 Excel→PDF</button>
<button class="tab" onclick="sw('t2',this)">📄 PDF→Excel</button>
<button class="tab" onclick="sw('t3',this)">🤖 AI Analysis</button>
</div>
<div id="t1" class="panel active">
<div class="drop" id="d1" onclick="document.getElementById('f1').click()">📂 Click or drag Excel file here</div>
<input type="file" id="f1" accept=".xlsx,.xls" style="display:none">
<div class="fname" id="fn1"></div>
<button class="btn" onclick="doE2P()">Convert to PDF</button>
<div class="result" id="r1"></div>
</div>
<div id="t2" class="panel">
<div class="drop" id="d2" onclick="document.getElementById('f2').click()">📂 Click or drag PDF bank statement here</div>
<input type="file" id="f2" accept=".pdf" style="display:none">
<div class="fname" id="fn2"></div>
<button class="btn" onclick="doP2E()">Convert to Excel</button>
<div class="result" id="r2"></div>
</div>
<div id="t3" class="panel">
<div class="drop" id="d3" onclick="document.getElementById('f3').click()">📂 Click or drag bank statement here</div>
<input type="file" id="f3" accept=".pdf,.xlsx,.xls" style="display:none">
<div class="fname" id="fn3"></div>
<button class="btn" onclick="doAI()">Analyze with AI</button>
<div class="result" id="r3"></div>
</div>
</div>
<script>
const files={};
async function loadU(){try{const r=await fetch('/usage');const d=await r.json();document.getElementById('uc').textContent=d.remaining;}catch(e){}}
loadU();
function sw(id,el){document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.getElementById(id).classList.add('active');el.classList.add('active');}
[1,2,3].forEach(i=>{
const d=document.getElementById('d'+i);
const f=document.getElementById('f'+i);
const fn=document.getElementById('fn'+i);
f.onchange=function(){if(this.files[0]){files[i]=this.files[0];fn.textContent='✅ '+this.files[0].name;}};
d.ondragover=e=>{e.preventDefault();d.classList.add('over');};
d.ondragleave=()=>d.classList.remove('over');
d.ondrop=e=>{e.preventDefault();d.classList.remove('over');const f2=e.dataTransfer.files[0];if(f2){files[i]=f2;fn.textContent='✅ '+f2.name;}};
});
let tbl='';
async function doE2P(){
if(!files[1]){alert('Select Excel file!');return;}
document.getElementById('r1').innerHTML='<p style="color:#38bdf8">Converting...</p>';
const fd=new FormData();fd.append('file',files[1]);
const r=await fetch('/excel-to-pdf',{method:'POST',body:fd});
const d=await r.json();loadU();
if(d.error==='FREE_LIMIT'){document.getElementById('r1').innerHTML='<p style="color:#E24B4A">Free limit reached. Upgrade to Pro.</p>';return;}
if(d.success){tbl=d.html;document.getElementById('r1').innerHTML=d.html+'<button class="btn" style="margin-top:12px;background:#22c55e;color:white;" onclick="dlPDF()">⬇️ Download PDF</button>';}
else{document.getElementById('r1').innerHTML='<p style="color:red">Error: '+d.error+'</p>';}
}
function dlPDF(){const w=window.open('','_blank');w.document.write('<html><body style="padding:20px;font-family:Arial;">'+tbl+'</body></html>');w.document.close();w.print();}
async function doP2E(){
if(!files[2]){alert('Select PDF file!');return;}
document.getElementById('r2').innerHTML='<p style="color:#38bdf8">Converting with AI... (15-20 seconds)</p>';
const fd=new FormData();fd.append('file',files[2]);
const r=await fetch('/pdf-to-excel',{method:'POST',body:fd});
loadU();
if(r.headers.get('content-type')&&r.headers.get('content-type').includes('json')){
const d=await r.json();
if(d.error==='FREE_LIMIT'){document.getElementById('r2').innerHTML='<p style="color:#E24B4A">Free limit reached.</p>';return;}
document.getElementById('r2').innerHTML='<p style="color:red">'+d.error+'</p>';
}else{
const blob=await r.blob();const url=URL.createObjectURL(blob);
const a=document.createElement('a');a.href=url;a.download='statement.xlsx';a.click();
document.getElementById('r2').innerHTML='<p style="color:#22c55e">✅ Excel downloaded! Columns properly separated.</p>';
}
}
async function doAI(){
if(!files[3]){alert('Select a file!');return;}
document.getElementById('r3').innerHTML='<p style="color:#38bdf8">AI analyzing... (15 seconds)</p>';
const fd=new FormData();fd.append('file',files[3]);
const r=await fetch('/analyze',{method:'POST',body:fd});
const d=await r.json();loadU();
if(d.error==='FREE_LIMIT'){document.getElementById('r3').innerHTML='<p style="color:#E24B4A">Free limit reached.</p>';return;}
if(d.success){
const a=d.analysis;
document.getElementById('r3').innerHTML=
'<div class="ai-card"><div class="ai-label">Total Credits</div><div class="ai-val">'+a.totalCredits+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Total Debits</div><div class="ai-val">'+a.totalDebits+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Average Balance</div><div class="ai-val">'+a.averageBalance+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Bounced Cheques</div><div class="ai-val">'+a.bouncedCheques+'</div></div>'+
'<div class="ai-card"><div class="ai-label">EMI Load</div><div class="ai-val">'+a.emiLoad+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Loan Eligibility</div><div class="ai-val">'+a.loanEligibility+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Credit Score</div><div class="ai-val">'+a.creditScore+'</div></div>'+
'<div class="ai-card"><div class="ai-label">Summary</div><div class="ai-val" style="color:#94a3b8;">'+a.summary+'</div></div>';
}else{document.getElementById('r3').innerHTML='<p style="color:red">'+d.error+'</p>';}
}
</script></body></html>`));

app.post('/excel-to-pdf',upload.single('file'),(req,res)=>{
if(getU(req.ip)>=5)return res.json({success:false,error:'FREE_LIMIT'});
try{
const wb=xlsx.read(req.file.buffer,{type:'buffer'});
const sh=wb.Sheets[wb.SheetNames[0]];
const data=xlsx.utils.sheet_to_json(sh,{header:1});
let html='<table><thead><tr>';
if(data[0])data[0].forEach(h=>{html+=`<th>${h||''}</th>`;});
html+='</tr></thead><tbody>';
data.slice(1).forEach(row=>{html+='<tr>';row.forEach(c=>{html+=`<td>${c||''}</td>`;});html+='</tr>';});
html+='</tbody></table>';
incU(req.ip);
res.json({success:true,html});
}catch(e){res.json({success:false,error:e.message});}
});

app.post('/pdf-to-excel',upload.single('file'),async(req,res)=>{
if(getU(req.ip)>=5)return res.json({success:false,error:'FREE_LIMIT'});
try{
const pdfLib=require('pdf-parse');
const pdfParse=pdfLib.default||pdfLib;
const parsed=await pdfParse(req.file.buffer);
const text=parsed.text;
const model=genAI.getGenerativeModel({model:'gemini-2.0-flash'});
const prompt=`Extract all bank transactions from this bank statement text and return ONLY a JSON array. Each item must have these exact fields: date, description, debit, credit, balance. If a field is empty use empty string. Return ONLY the JSON array, no other text.
Bank statement text:
${text.substring(0,4000)}`;
const result=await model.generateContent(prompt);
const raw=result.response.text().replace(/\`\`\`json|\`\`\`/g,'').trim();
const transactions=JSON.parse(raw);
const wb=xlsx.utils.book_new();
const wsData=[['Date','Description','Debit (INR)','Credit (INR)','Balance (INR)']];
transactions.forEach(t=>{wsData.push([t.date||'',t.description||'',t.debit||'',t.credit||'',t.balance||'']);});
const ws=xlsx.utils.aoa_to_sheet(wsData);
ws['!cols']=[{wch:12},{wch:40},{wch:15},{wch:15},{wch:15}];
xlsx.utils.book_append_sheet(wb,ws,'Statement');
const buf=xlsx.write(wb,{type:'buffer',bookType:'xlsx'});
incU(req.ip);
res.setHeader('Content-Disposition','attachment; filename=statement.xlsx');
res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
res.send(buf);
}catch(e){res.json({success:false,error:e.message});}
});

app.post('/analyze',upload.single('file'),async(req,res)=>{
if(getU(req.ip)>=5)return res.json({success:false,error:'FREE_LIMIT'});
try{
let text='';
if(req.file.originalname.toLowerCase().endsWith('.pdf')){
const pdfLib=require('pdf-parse');
const pdfParse=pdfLib.default||pdfLib;
const parsed=await pdfParse(req.file.buffer);
text=parsed.text;
}else{
const wb=xlsx.read(req.file.buffer,{type:'buffer'});
const sh=wb.Sheets[wb.SheetNames[0]];
text=xlsx.utils.sheet_to_json(sh,{header:1}).map(r=>r.join(', ')).join('\n');
}
const model=genAI.getGenerativeModel({model:'gemini-2.0-flash'});
const result=await model.generateContent(`Analyze this Indian bank statement. Return ONLY valid JSON:
{"totalCredits":"INR amount","totalDebits":"INR amount","averageBalance":"INR amount","bouncedCheques":"number","emiLoad":"INR amount or Not detected","topExpenses":["cat1","cat2","cat3"],"loanEligibility":"INR amount","creditScore":"Excellent/Good/Fair/Poor","summary":"2 sentences"}
Data: ${text.substring(0,3000)}`);
const raw=result.response.text().replace(/\`\`\`json|\`\`\`/g,'').trim();
incU(req.ip);
res.json({success:true,analysis:JSON.parse(raw)});
}catch(e){res.json({success:false,error:e.message});}
});

app.listen(process.env.PORT||3000,()=>console.log('BankPDF running'));
module.exports=app;
