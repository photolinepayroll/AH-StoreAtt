require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const path    = require('path');
const fs      = require('fs');
const Anthropic = require('@anthropic-ai/sdk');
const ExcelJS   = require('exceljs');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const MASTERLIST_PATH = path.join(__dirname, 'data', 'Attendance_Masterlist.xlsx');

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ── Extract attendance from image ────────────────────────────────────────────
app.post('/api/extract', upload.single('image'), async (req, res) => {
  try {
    const { store = '', date = '' } = req.body;
    const imgBuffer = req.file.buffer;
    const mimeType  = req.file.mimetype;
    const base64    = imgBuffer.toString('base64');

    const prompt = `You are an attendance data extraction expert.
Analyze this attendance photo/sheet and extract ALL employee records.

Return ONLY a valid JSON object with this exact structure:
{
  "meta": { "store": "${store}", "date": "${date}", "notes": "" },
  "rows": [
    {
      "emp_id": "employee ID or empty string",
      "name": "full name",
      "position": "job position/title or empty string",
      "time_in": "time in (e.g. 8:00 AM) or empty string",
      "time_out": "time out (e.g. 5:00 PM) or empty string",
      "status": "Present|Absent|Late|Half Day|Day Off",
      "remarks": "any notes or empty string"
    }
  ]
}

Rules:
- If a field is not visible, use empty string "".
- Status must be exactly one of: Present, Absent, Late, Half Day, Day Off.
- Extract every single person visible in the attendance sheet.
- Do not include markdown formatting, only pure JSON.`;

    const message = await client.messages.create({
      model: 'claude-opus-4-8',
      max_tokens: 4096,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64 } },
          { type: 'text', text: prompt },
        ],
      }],
    });

    const raw = message.content[0].text.trim();
    // strip possible markdown fences
    const jsonStr = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
    const parsed  = JSON.parse(jsonStr);

    res.json(parsed);
  } catch (err) {
    console.error('Extract error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Save / update masterlist Excel ───────────────────────────────────────────
app.post('/api/save', async (req, res) => {
  try {
    const { rows, store, date } = req.body;

    // ensure data dir exists
    fs.mkdirSync(path.dirname(MASTERLIST_PATH), { recursive: true });

    const wb = new ExcelJS.Workbook();
    let ws;

    // load existing or create new
    if (fs.existsSync(MASTERLIST_PATH)) {
      await wb.xlsx.readFile(MASTERLIST_PATH);
      ws = wb.getWorksheet('Masterlist') || wb.addWorksheet('Masterlist');
    } else {
      ws = wb.addWorksheet('Masterlist');

      // Header row styling
      ws.columns = [
        { header: 'No.',         key: 'no',       width: 6  },
        { header: 'Emp ID',      key: 'emp_id',   width: 14 },
        { header: 'Name',        key: 'name',     width: 28 },
        { header: 'Position',    key: 'position', width: 20 },
        { header: 'Store/Branch',key: 'store',    width: 20 },
        { header: 'Date',        key: 'date',     width: 14 },
        { header: 'Time In',     key: 'time_in',  width: 12 },
        { header: 'Time Out',    key: 'time_out', width: 12 },
        { header: 'Status',      key: 'status',   width: 12 },
        { header: 'Remarks',     key: 'remarks',  width: 24 },
      ];

      const headerRow = ws.getRow(1);
      headerRow.eachCell(cell => {
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3C6E' } };
        cell.font   = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
      headerRow.height = 22;
    }

    // determine next row number
    const lastNo = ws.lastRow
      ? (ws.lastRow.getCell('no').value || ws.rowCount - 1)
      : 0;
    let counter = typeof lastNo === 'number' ? lastNo : ws.rowCount - 1;

    const statusColors = {
      'Present':  'FFC8E6C9',
      'Absent':   'FFFFCDD2',
      'Late':     'FFFFF3E0',
      'Half Day': 'FFE1BEE7',
      'Day Off':  'FFE0E0E0',
    };

    rows.forEach(r => {
      counter++;
      const row = ws.addRow({
        no:       counter,
        emp_id:   r.emp_id   || '',
        name:     r.name     || '',
        position: r.position || '',
        store:    store       || '',
        date:     date        || '',
        time_in:  r.time_in  || '',
        time_out: r.time_out || '',
        status:   r.status   || '',
        remarks:  r.remarks  || '',
      });

      const color = statusColors[r.status] || 'FFFFFFFF';
      row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
      row.getCell('status').font = { bold: true };
      row.getCell('name').font   = { bold: true };
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFD0D8E0' } },
          left: { style: 'thin', color: { argb: 'FFD0D8E0' } },
          bottom: { style: 'thin', color: { argb: 'FFD0D8E0' } },
          right: { style: 'thin', color: { argb: 'FFD0D8E0' } },
        };
      });
    });

    // also keep a per-date sheet
    const sheetName = (date || 'NoDate').replace(/\//g, '-').slice(0, 31);
    let dateWs = wb.getWorksheet(sheetName);
    if (!dateWs) {
      dateWs = wb.addWorksheet(sheetName);
      dateWs.columns = ws.columns.map(c => ({ ...c }));
      const dHeader = dateWs.getRow(1);
      dHeader.eachCell(cell => {
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3C6E' } };
        cell.font   = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
    }
    rows.forEach((r, idx) => {
      const row = dateWs.addRow({
        no: idx + 1, emp_id: r.emp_id || '', name: r.name || '',
        position: r.position || '', store: store || '', date: date || '',
        time_in: r.time_in || '', time_out: r.time_out || '',
        status: r.status || '', remarks: r.remarks || '',
      });
      const color = statusColors[r.status] || 'FFFFFFFF';
      row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    });

    // save to disk as well for persistence
    await wb.xlsx.writeFile(MASTERLIST_PATH);

    // send as download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Attendance_Masterlist.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Save error:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AH-StoreAtt running on http://localhost:${PORT}`));
