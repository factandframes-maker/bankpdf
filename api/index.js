const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse').default || require('pdf-parse');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const genAI = new GoogleGenerativeAI('AIzaSyDiMb8UEwhnkDD8uOS1U3rmGmsDqpl9qU0');

app.use(express.static('public'));
app.use(express.json());

const userUsage = {};
function getUsage(ip){ if(!userUsage[ip]) userUsage[ip]=0; return userUsage[ip]; }
function incrementUsage(ip){ if(!userUsage[ip]) userUsage[ip]=0; userUsage[ip]++; }

app.get('/usage',(req,res)=>{ const ip=req.ip; res.json({used:getUsage(ip),limit:5,remaining:Math.max(0,5-getUsage(ip))}); });

app.post('/excel-to-pdf', upload.single('file'), (req,res)=>{
  const ip=req.ip;
  if(getUsage(ip)>=5) return res.json({success:false,error:'FREE_LIMIT'});
  try{
    const workbook=xlsx.read(req.file.buffer,{type:'buffer'});
    const sheet=workbook.Sheets[workbook.SheetNames[0]];
    const data=xlsx.utils.sheet_to_json(sheet,{header:1});
    let html='<table border="1" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:13px;">';
    data.forEach((row,i)=>{
      html+='<tr>';
      row.forEach(cell=>{ const tag=i===0?'th':'td'; html+=`<${tag} style="padding:8px;background:${i===0?'#1e3a5f':'#fff'};color:${i===0?'#fff':'#000'};">${cell||''}</${tag}>`; });
      html+='</tr>';
    });
    html+='</table>';
    incrementUsage(ip);
    res.json({success:true,html,used:getUsage(ip)});
  }catch(err){ res.json({success:false,error:err.message}); }
});

app.post('/pdf-to-excel', upload.single('file'), async (req,res)=>{
  const ip=req.ip;
  if(getUsage(ip)>=5) return res.json({success:false,error:'FREE_LIMIT'});
  try{
    const data=await pdfParse(req.file.buffer);
    const text=data.text;
    const lines=text.split('\n').filter(l=>l.trim());
    const rows=lines.map(line=>line.split(/\s{2,}/).map(c=>c.trim()));
    const workbook=xlsx.utils.book_new();
    const worksheet=xlsx.utils.aoa_to_sheet(rows);
    xlsx.utils.book_append_sheet(workbook,worksheet,'Statement');
    const buffer=xlsx.write(workbook,{type:'buffer',bookType:'xlsx'});
    incrementUsage(ip);
    res.setHeader('Content-Disposition','attachment; filename=statement.xlsx');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  }catch(err){ res.json({success:false,error:err.message}); }
});

app.post('/analyze', upload.single('file'), async (req,res)=>{
  const ip=req.ip;
  if(getUsage(ip)>=5) return res.json({success:false,error:'FREE_LIMIT'});
  try{
    let text='';
    if(req.file.originalname.toLowerCase().endsWith('.pdf')){
      const data=await pdfParse(req.file.buffer);
      text=data.text;
    }else{
      const workbook=xlsx.read(req.file.buffer,{type:'buffer'});
      const sheet=workbook.Sheets[workbook.SheetNames[0]];
      const data=xlsx.utils.sheet_to_json(sheet,{header:1});
      text=data.map(r=>r.join(', ')).join('\n');
    }
    const model=genAI.getGenerativeModel({model:'gemini-1.5-flash'});
    const prompt=`You are a financial analyst. Analyze this Indian bank statement and return ONLY a JSON object with no extra text:
{
  "totalCredits": "total money coming in with rupee symbol",
  "totalDebits": "total money going out with rupee symbol",
  "averageBalance": "average balance with rupee symbol",
  "bouncedCheques": "number of bounced cheques or 0",
  "emiLoad": "estimated monthly EMI or Not detected",
  "topExpenses": ["category1","category2","category3"],
  "loanEligibility": "estimated loan eligibility amount",
  "creditScore": "Excellent or Good or Fair or Poor",
  "summary": "2 sentence plain English summary of financial health"
}
Data:
${text.substring(0,3000)}`;
    const result=await model.generateContent(prompt);
    const raw=result.response.text().replace(/```json|```/g,'').trim();
    const analysis=JSON.parse(raw);
    incrementUsage(ip);
    res.json({success:true,analysis,used:getUsage(ip)});
  }catch(err){ res.json({success:false,error:err.message}); }
});

app.listen(3000,()=>console.log('BankPDF running on http://localhost:3000'));
