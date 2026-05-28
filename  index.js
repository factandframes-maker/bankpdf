const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const genAI = new GoogleGenerativeAI('AIzaSyDiMb8UEwhnkDD8uOS1U3rmGmsDqpl9qU0');
app.use(express.static('public'));
app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const csvText = data.map(r => r.join(', ')).join('\n');
    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });
    const prompt = `Analyze this bank statement and return ONLY raw JSON with no markdown or extra text: {"totalCredits":number,"totalDebits":number,"emiLoad":"High/Medium/Low","bouncedCheques":number,"avgBalance":number,"loanEligibility":"Eligible/Not Eligible/Borderline","summary":"2 sentence summary"}. Data:\n${csvText}`;
    const result = await model.generateContent(prompt);
    const raw = result.response.text().replace(/```json|```/g, '').trim();
    const ai = JSON.parse(raw);
    let t = '<table><thead><tr>';
    if (data[0]) data[0].forEach(h => { t += `<th>${h}</th>`; });
    t += '</tr></thead><tbody>';
    data.slice(1).forEach(row => { t += '<tr>'; row.forEach(cell => { t += `<td>${cell??''}</td>`; }); t += '</tr>'; });
    t += '</tbody></table>';
    const aiHtml = `<div style="margin-top:20px;background:#0f2744;border-radius:12px;padding:20px;text-align:left;"><h3 style="color:#38bdf8;margin-bottom:12px">🤖 AI Analysis</h3><p style="color:#94a3b8;margin-bottom:16px">${ai.summary}</p><div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;"><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">Total Credits</p><p style="color:#22c55e;font-weight:bold">₹${ai.totalCredits?.toLocaleString()}</p></div><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">Total Debits</p><p style="color:#ef4444;font-weight:bold">₹${ai.totalDebits?.toLocaleString()}</p></div><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">EMI Load</p><p style="color:#f59e0b;font-weight:bold">${ai.emiLoad}</p></div><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">Loan Eligibility</p><p style="color:#38bdf8;font-weight:bold">${ai.loanEligibility}</p></div><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">Bounced Cheques</p><p style="color:#ef4444;font-weight:bold">${ai.bouncedCheques}</p></div><div style="background:#1e293b;padding:12px;border-radius:8px"><p style="color:#94a3b8;font-size:12px">Avg Balance</p><p style="color:#22c55e;font-weight:bold">₹${ai.avgBalance?.toLocaleString()}</p></div></div></div>`;
    res.json({ success: true, html: t + aiHtml });
  } catch(e) { res.json({ success: false, error: e.message }); }
});
app.listen(3000, () => console.log('BankPDF running on http://localhost:3000'));