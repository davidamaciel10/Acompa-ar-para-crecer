require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors());
app.use(express.json());

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const PERSONAS_SHEET = 'Personas';
const MONTH_NAMES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

let _sheetsClient = null;
async function getSheets() {
  if (_sheetsClient) return _sheetsClient;
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({ credentials: creds, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const client = await auth.getClient();
  _sheetsClient = google.sheets({ version: 'v4', auth: client });
  return _sheetsClient;
}

function colLetter(idx) {
  let n = idx + 1, s = '';
  while (n > 0) { const r = (n-1)%26; s = String.fromCharCode(65+r)+s; n = Math.floor((n-1)/26); }
  return s;
}
function getSabados(year, month) {
  const result = []; const d = new Date(year, month-1, 1);
  while (d.getMonth()===month-1) { if (d.getDay()===6) result.push(new Date(d)); d.setDate(d.getDate()+1); }
  return result;
}
function fmtDDMM(date) { return String(date.getDate()).padStart(2,'0')+'/'+String(date.getMonth()+1).padStart(2,'0'); }
function sheetNameForMonth(y, m) { return MONTH_NAMES[m-1]+' '+y; }

async function ensureMonthSheet(sheets, year, month) {
  const name = sheetNameForMonth(year, month);
  const r = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = r.data.sheets.map(s => s.properties.title);
  if (!existing.includes(name)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: name } } }] } });
    const pr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: PERSONAS_SHEET+'!A2:E' });
    const personas = (pr.data.values||[]).map(r => [r[0]||'',r[1]||'',r[2]||'',r[4]||'']);
    const sabados = getSabados(year, month).map(fmtDDMM);
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: name+'!A1', valueInputOption: 'RAW', requestBody: { values: [['Nombre','Apellido','DNI','Tipo',...sabados],...personas] } });
  }
  return name;
}

app.get('/', (_req, res) => res.json({ ok: true, service: 'asistencia-api' }));

app.get('/sabados', (_req, res) => {
  const now = new Date(), y = now.getFullYear(), m = now.getMonth()+1;
  res.json({ sabados: getSabados(y,m).map(d => ({ ddmm: fmtDDMM(d), iso: d.toISOString().split('T')[0] })), mes: sheetNameForMonth(y,m) });
});

app.get('/personas', async (req, res) => {
  try {
    const q = (req.query.q||'').toLowerCase().trim();
    const sheets = await getSheets();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: PERSONAS_SHEET+'!A2:E' });
    let personas = (r.data.values||[]).map((row,i) => ({ rowIndex: i+2, nombre: row[0]||'', apellido: row[1]||'', dni: row[2]||'', tipo: row[4]||'' }));
    if (q) personas = personas.filter(p => (p.nombre+' '+p.apellido+' '+p.dni).toLowerCase().includes(q));
    res.json(personas.slice(0,30));
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/asistencia/:fecha', async (req, res) => {
  try {
    const [dd,mm] = req.params.fecha.split('-');
    const fechaCol = dd+'/'+mm, year = new Date().getFullYear(), month = parseInt(mm);
    const sheets = await getSheets();
    const sheetName = await ensureMonthSheet(sheets, year, month);
    const hr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!1:1' });
    const headers = hr.data.values?.[0]||[];
    const colIdx = headers.indexOf(fechaCol);
    if (colIdx===-1) return res.json({ presentes: [], total: 0 });
    const dr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!A2:Z' });
    const rows = dr.data.values||[];
    const presentes = rows.reduce((acc,row,i) => {
      if ((row[colIdx]||'').trim()==='\u2713') acc.push({ rowIndex: i+2, nombre: row[0]||'', apellido: row[1]||'', dni: row[2]||'', tipo: row[3]||'' });
      return acc;
    }, []);
    res.json({ presentes, total: rows.length });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/asistencia', async (req, res) => {
  try {
    const { dni, fecha, marcar } = req.body;
    if (!dni||!fecha) return res.status(400).json({ error: 'Faltan campos: dni, fecha' });
    const [dd,mm] = fecha.split('-');
    const fechaCol = dd+'/'+mm, year = new Date().getFullYear(), month = parseInt(mm);
    const sheets = await getSheets();
    const sheetName = await ensureMonthSheet(sheets, year, month);
    const hr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!1:1' });
    const colIdx = (hr.data.values?.[0]||[]).indexOf(fechaCol);
    if (colIdx===-1) return res.status(400).json({ error: 'Fecha '+fechaCol+' no encontrada' });
    const dr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!C2:C' });
    const offset = (dr.data.values||[]).findIndex(r => String(r[0]).trim()===String(dni).trim());
    if (offset===-1) return res.status(404).json({ error: 'DNI '+dni+' no encontrado' });
    const range = sheetName+'!'+colLetter(colIdx)+(offset+2);
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range, valueInputOption: 'RAW', requestBody: { values: [[marcar?'\u2713':'']] } });
    res.json({ ok: true, range, value: marcar?'\u2713':'' });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

const PORT = process.env.PORT||3000;
app.listen(PORT, () => console.log('asistencia-api en :'+PORT));
