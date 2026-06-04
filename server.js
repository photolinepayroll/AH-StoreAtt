require('dotenv').config();
const express   = require('express');
const multer    = require('multer');
const path      = require('path');
const Anthropic = require('@anthropic-ai/sdk');
const ExcelJS   = require('exceljs');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '10mb' }));

// ── POST /api/extract  ────────────────────────────────────────────────────────
// Accepts one image, returns extracted attendance rows as JSON
app.post('/api/extract', upload.single('image'), async (req, res) => {
  try {
    const pageNum  = req.body.pageNum || 1;
    const base64   = req.file.buffer.toString('base64');
    const mimeType = req.file.mimetype;

    const prompt = `You are an expert at reading Philippine retail store attendance sheets.

Analyze this attendance photo and extract ALL employee records visible.

Return ONLY a valid JSON object — no markdown, no explanation:
{
  "rows": [
    {
      "branch": "branch/store name exactly as written",
      "name": "FULL NAME as written",
      "d1_15": <number or null>,
      "d16_31": <number or null>,
      "status": "Present"
    }
  ]
}

Rules:
- "branch" — copy the branch name exactly as it appears in the sheet
- "name" — full name in UPPERCASE as shown
- "d1_15" — the attendance count for days 1-15 (first number column). Use null if blank.
- "d16_31" — the attendance count for days 16-31 OR the second number column (labeled 16-31 or 1-31). Use null if blank.
- "status" — always set to "Present" unless the sheet explicitly marks otherwise
- Extract EVERY person — do not skip anyone
- Empty rows between groups should be ignored
- Return pure JSON only`;

    const msg = await client.messages.create({
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

    const raw     = msg.content[0].text.trim();
    const jsonStr = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
    const parsed  = JSON.parse(jsonStr);

    res.json({ rows: parsed.rows || [] });
  } catch (err) {
    console.error('Extract error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/update-masterlist  ──────────────────────────────────────────────
// Accepts the masterlist Excel + extracted rows JSON
// Matches employees, fills 1-15 and 16-31 columns, returns updated Excel as base64
app.post('/api/update-masterlist',
  upload.fields([{ name: 'masterlist', maxCount: 1 }]),
  async (req, res) => {
    try {
      const xlBuf = req.files['masterlist'][0].buffer;
      const rows  = JSON.parse(req.body.rows || '[]');

      // load workbook
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(xlBuf);
      const ws = wb.worksheets[0];

      // build employee index from masterlist
      // columns: A=Last, B=First, C=Middle, D=Branch, E=1-15, F=16-31
      const employees = [];
      ws.eachRow((row, rn) => {
        if (rn < 3) return;
        const last   = cellStr(row, 1);
        const first  = cellStr(row, 2);
        const branch = cellStr(row, 4);
        if (last) employees.push({ rn, last, first, branch });
      });

      const matched   = [];
      const unmatched = [];

      for (const rec of rows) {
        const hit = findEmployee(rec, employees);
        if (hit) {
          const row = ws.getRow(hit.rn);
          if (rec.d1_15  !== null && rec.d1_15  !== undefined) {
            row.getCell(5).value = Number(rec.d1_15);
            styleCell(row.getCell(5));
          }
          if (rec.d16_31 !== null && rec.d16_31 !== undefined) {
            row.getCell(6).value = Number(rec.d16_31);
            styleCell(row.getCell(6));
          }
          row.commit();
          rec._matched   = true;
          rec._masterRow = hit.rn;
          matched.push(rec);
        } else {
          rec._matched = false;
          unmatched.push({ name: rec.name, branch: rec.branch });
        }
      }

      // write to buffer
      const outBuf  = await wb.xlsx.writeBuffer();
      const b64     = outBuf.toString('base64');
      const now     = new Date();
      const stamp   = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}`;
      const filename = `Attendance_Masterlist_${stamp}.xlsx`;

      res.json({ matched: matched.length, unmatched, fileBase64: b64, filename });
    } catch (err) {
      console.error('Update masterlist error:', err.message);
      res.status(500).json({ error: err.message });
    }
  }
);

// ── Helpers ───────────────────────────────────────────────────────────────────
function cellStr(row, col) {
  return ((row.getCell(col).value) || '').toString().trim();
}

function norm(s) {
  return s.toLowerCase()
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\bjr\b\.?/g, '')
    .replace(/\bsr\b\.?/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function branchSim(a, b) {
  const na = norm(a), nb = norm(b);
  if (na === nb) return 1;
  if (na.includes(nb) || nb.includes(na)) return 0.9;
  // token overlap
  const ta = new Set(na.split(' ')), tb = new Set(nb.split(' '));
  const inter = [...ta].filter(t => tb.has(t)).length;
  return inter / Math.max(ta.size, tb.size);
}

function findEmployee(rec, employees) {
  const fullNorm = norm(rec.name || '');
  const words    = fullNorm.split(' ').filter(Boolean);
  if (words.length === 0) return null;

  let best = null, bestScore = 0;

  for (const emp of employees) {
    const ln = norm(emp.last);
    const fn = norm(emp.first);
    if (!ln) continue;

    // last name must appear in the attendance full name
    if (!words.includes(ln) && !fullNorm.includes(ln)) continue;

    // first-name bonus
    const fnWords = fn.split(' ').filter(Boolean);
    const fnMatch = fnWords.length > 0 && fnWords.some(w => fullNorm.includes(w)) ? 1 : 0;

    const bScore  = branchSim(rec.branch || '', emp.branch);
    const score   = (1 + fnMatch) * (0.5 + bScore);

    if (score > bestScore) { bestScore = score; best = emp; }
  }

  // require at least some branch similarity to avoid wrong-person matches
  return bestScore >= 0.8 ? best : null;
}

function styleCell(cell) {
  cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } };
  cell.font      = { bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AH-StoreAtt  →  http://localhost:${PORT}`));
