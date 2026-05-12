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

function syntheticId(rowIndex) { return '~' + rowIndex; }
function isSynthetic(dni) { return String(dni).startsWith('~'); }
function rowFromSynthetic(dni) { return parseInt(String(dni).slice(1)); }

// Acepta "YYYY-MM-DD"
function parseFecha(fecha) {
  const [yyyy, mm, dd] = fecha.split('-');
  return { year: parseInt(yyyy), month: parseInt(mm), fechaCol: `${dd}/${mm}` };
}

async function ensureMonthSheet(sheets, year, month) {
  const name = sheetNameForMonth(year, month);
  const r = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = r.data.sheets.map(s => s.properties.title);
  if (!existing.includes(name)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: name } } }] } });
    const pr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: PERSONAS_SHEET+'!A2:E' });
    const personas = (pr.data.values||[]).map((row, i) => [row[0]||'', row[1]||'', row[2]||syntheticId(i+2), row[4]||'']);
    const sabados = getSabados(year, month).map(fmtDDMM);
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: name+'!A1', valueInputOption: 'RAW', requestBody: { values: [['Nombre','Apellido','DNI','Tipo',...sabados],...personas] } });
  }
  return name;
}

app.get('/', (_req, res) => res.json({ ok: true, service: 'asistencia-api' }));

app.get('/personas', async (req, res) => {
  try {
    const q = (req.query.q||'').toLowerCase().trim();
    const sheets = await getSheets();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: PERSONAS_SHEET+'!A2:E' });
    let personas = (r.data.values||[]).map((row, i) => ({
      rowIndex: i+2,
      nombre: row[0]||'',
      apellido: row[1]||'',
      dni: row[2]||syntheticId(i+2),
      actividad: row[3]||'',
      tipo: row[4]||'',
    }));
    if (q) personas = personas.filter(p => (p.nombre+' '+p.apellido+' '+p.dni).toLowerCase().includes(q));
    res.json(personas);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/personas', async (req, res) => {
  try {
    const { nombre, apellido, dni, actividad, tipo } = req.body;
    if (!nombre||!apellido) return res.status(400).json({ error: 'Faltan campos: nombre y apellido' });
    const sheets = await getSheets();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: PERSONAS_SHEET+'!A:E',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[nombre, apellido, dni||'', actividad||'', tipo||'']] },
    });
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/personas/:rowIndex', async (req, res) => {
  try {
    const rowIndex = parseInt(req.params.rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return res.status(400).json({ error: 'rowIndex inválido' });
    const { nombre, apellido, dni, actividad, tipo } = req.body;
    if (!nombre||!apellido) return res.status(400).json({ error: 'Faltan campos: nombre y apellido' });
    const sheets = await getSheets();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERSONAS_SHEET}!A${rowIndex}:E${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[nombre, apellido, dni||'', actividad||'', tipo||'']] },
    });
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Crea la columna de una fecha en el sheet del mes si no existe
app.post('/fecha', async (req, res) => {
  try {
    const { fecha } = req.body; // "YYYY-MM-DD"
    if (!fecha) return res.status(400).json({ error: 'Falta fecha' });
    const { year, month, fechaCol } = parseFecha(fecha);
    const sheets = await getSheets();
    const sheetName = await ensureMonthSheet(sheets, year, month);
    const hr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!1:1' });
    const headers = hr.data.values?.[0] || [];
    if (headers.includes(fechaCol)) return res.json({ ok: true, exists: true });
    const range = sheetName + '!' + colLetter(headers.length) + '1';
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range, valueInputOption: 'RAW', requestBody: { values: [[fechaCol]] } });
    res.json({ ok: true, exists: false });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/personas/:rowIndex', async (req, res) => {
  try {
    const rowIndex = parseInt(req.params.rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return res.status(400).json({ error: 'rowIndex inválido' });
    const sheets = await getSheets();
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheet = meta.data.sheets.find(s => s.properties.title === PERSONAS_SHEET);
    if (!sheet) return res.status(404).json({ error: 'Sheet no encontrado' });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ deleteDimension: { range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex } } }] },
    });
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/asistencia/:fecha', async (req, res) => {
  try {
    const { year, month, fechaCol } = parseFecha(req.params.fecha);
    const sheets = await getSheets();
    const sheetName = await ensureMonthSheet(sheets, year, month);
    const hr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!1:1' });
    const headers = hr.data.values?.[0]||[];
    const colIdx = headers.indexOf(fechaCol);
    if (colIdx===-1) return res.json({ presentes: [], total: 0 });
    const dr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!A2:Z' });
    const rows = dr.data.values||[];
    const presentes = rows.reduce((acc, row, i) => {
      if ((row[colIdx]||'').trim()==='✓') {
        acc.push({ rowIndex: i+2, nombre: row[0]||'', apellido: row[1]||'', dni: row[2]||syntheticId(i+2), tipo: row[3]||'' });
      }
      return acc;
    }, []);
    res.json({ presentes, total: rows.length });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/asistencia', async (req, res) => {
  try {
    const { dni, fecha, marcar } = req.body;
    if (!dni||!fecha) return res.status(400).json({ error: 'Faltan campos: dni, fecha' });
    const { year, month, fechaCol } = parseFecha(fecha);
    const sheets = await getSheets();
    const sheetName = await ensureMonthSheet(sheets, year, month);
    const hr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!1:1' });
    const colIdx = (hr.data.values?.[0]||[]).indexOf(fechaCol);
    if (colIdx===-1) return res.status(400).json({ error: 'Fecha '+fechaCol+' no encontrada en el sheet' });

    let targetRow;
    if (isSynthetic(dni)) {
      targetRow = rowFromSynthetic(dni);
    } else {
      const dr = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: sheetName+'!C2:C' });
      const offset = (dr.data.values||[]).findIndex(r => String(r[0]).trim()===String(dni).trim());
      if (offset===-1) return res.status(404).json({ error: 'DNI '+dni+' no encontrado' });
      targetRow = offset + 2;
    }

    const range = sheetName+'!'+colLetter(colIdx)+targetRow;
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range, valueInputOption: 'RAW', requestBody: { values: [[marcar?'✓':'']] } });
    res.json({ ok: true, range, value: marcar?'✓':'' });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

const PORT = process.env.PORT||3000;
app.listen(PORT, () => console.log('asistencia-api en :'+PORT));
