// Precise update script: each record maps directly to a masterlist row number
// May 2026 attendance data from both Area Head sheets

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const SRC  = '/root/.claude/uploads/d60aa4b3-97db-4d8c-b1f0-fd7f0523581a/daac958e-Masterlist.xlsx';
const DEST = path.join(__dirname, 'data', 'Attendance_Masterlist_May2026.xlsx');

// row = masterlist row number, null = not found in masterlist (new/unmatched employee)
// d1_15 = attendance days 1-15, d16_31 = attendance days 16-31
const RECORDS = [
  // ── DARWIN TOLEDO (Area Head) ─────────────────────────────────────────────
  // SM Grand Central
  { row: 307, name: 'MARIVIN MAQUINAD',      branch: 'SM Grand Central',  d1_15: 13, d16_31: 13 },
  { row: 292, name: 'EMMANUEL MAGNO',        branch: 'SM Grand Central',  d1_15: 13, d16_31: 14 },
  { row: 250, name: 'VINCENT IGLESIA',       branch: 'SM Grand Central',  d1_15: 13, d16_31:  9 },  // masterlist: Iglesias
  { row: null,name: 'JUZZTINE LABROMONTE',   branch: 'SM Grand Central',  d1_15: 12, d16_31: 13 },  // NOT IN MASTERLIST
  // SM Manila
  { row: 460, name: 'ALVIN TIANGCO',         branch: 'SM Manila',         d1_15: 13, d16_31: 14 },
  { row: 115, name: 'CHRISTIAN CAGUIOA',     branch: 'SM Manila',         d1_15: 13, d16_31: 14 },
  { row: 514, name: 'JAMAICA CALLEJA',       branch: 'SM Manila',         d1_15: 13, d16_31: 13 },
  { row: 418, name: 'ELLA JANE SAMSON',      branch: 'SM Manila',         d1_15: 13, d16_31: 14 },
  // Rob Ermita Pro
  { row: 493, name: 'ARMANDO YABA',          branch: 'Rob Ermita Pro',    d1_15: 11, d16_31: 13 },
  { row: 492, name: 'JONAS VIZCARRA',        branch: 'Rob Ermita Pro',    d1_15: 11, d16_31: 11 },
  { row:  99, name: 'RHEA CABACCAN',         branch: 'Rob Ermita Pro',    d1_15: 12, d16_31: 13 },
  { row: 330, name: 'JOSHUA MUNDO',          branch: 'Rob Ermita Pro',    d1_15: 14, d16_31: 12 },
  // Rob Ermita 1
  { row: 118, name: 'RACHELLE CAMUS',        branch: 'Rob Ermita 1',      d1_15: 10, d16_31: 10 },
  { row: 438, name: 'ROMEO SORIANO',         branch: 'Rob Ermita 1',      d1_15: 13, d16_31: 15 },
  { row: 174, name: 'KYLA DELA CRUZ',        branch: 'Rob Ermita 1',      d1_15: 13, d16_31:  0 },
  { row: 351, name: 'JOHN MARK PAGULAYAN',   branch: 'Rob Ermita 1',      d1_15: 14, d16_31: 14 },
  { row: 342, name: 'BABY APRIL OROCEO',     branch: 'Rob Ermita 1',      d1_15:  5, d16_31: 13 },
  // SM Bicutan
  { row: 484, name: 'SILVESTER VINARAO',     branch: 'SM Bicutan',        d1_15: 13, d16_31: 13 },
  { row: 197, name: 'NOAH ESCORO',           branch: 'SM Bicutan',        d1_15: 13, d16_31: 13 },
  { row: 181, name: 'JOSHUA DIAZ',           branch: 'SM Bicutan',        d1_15: 13, d16_31: 11 },
  { row:  63, name: 'RICHE BELLO',           branch: 'SM Bicutan',        d1_15: 11, d16_31: 13 },
  // SM San Lazaro
  { row: 291, name: 'DIANA MAE MAGLEO',      branch: 'SM San Lazaro',     d1_15: 13, d16_31:  8 },
  { row: 109, name: 'VLADIMER CABUYAO',      branch: 'SM San Lazaro',     d1_15: 13, d16_31: 14 },
  { row: 424, name: 'JB SANTOS',             branch: 'SM San Lazaro',     d1_15: 13, d16_31: 14 },
  { row: 512, name: 'MIA CUDAL',             branch: 'SM San Lazaro',     d1_15:  9, d16_31: 12 },
  { row: 287, name: 'KRISTINE MADREDIJO',    branch: 'SM San Lazaro',     d1_15:  8, d16_31: 13 },
  { row: 513, name: 'SAPPHIRE BADILLA',      branch: 'SM San Lazaro',     d1_15:  8, d16_31: 12 },
  // SM Southmall
  { row: 132, name: 'RICHARD CASIMERO',      branch: 'SM Southmall',      d1_15: 13, d16_31: 14 },
  { row: 180, name: 'LOUIE DIAZ',            branch: 'SM Southmall',      d1_15: 13, d16_31: 14 },
  { row: 408, name: 'HAROLD SABATER',        branch: 'SM Southmall',      d1_15: 12, d16_31: 14 },
  // SM Las Piñas
  { row: 212, name: 'ANA FLORENDA',          branch: 'SM Las Pinas',      d1_15: 13, d16_31: 14 },
  { row: 134, name: 'WENDELL CASTILLO',      branch: 'SM Las Pinas',      d1_15: 13, d16_31: 14 },
  { row: 190, name: 'CASSANDRA ELLAREZ',     branch: 'SM Las Pinas',      d1_15:  8, d16_31: 14 },
  { row: 290, name: 'JUNREL MAGALO',         branch: 'SM Las Pinas',      d1_15: 13, d16_31: 14 },
  // Arca South
  { row:  98, name: 'JEFFREY BUSOG',         branch: 'Arca South',        d1_15: 12, d16_31: 13 },
  { row: 119, name: 'MJHAY CAMUS',           branch: 'Arca South',        d1_15: 12, d16_31: 14 },
  { row:  74, name: 'MAY BIGNOTEA',          branch: 'Arca South',        d1_15: 13, d16_31: 11 },

  // ── JECILDA VILLACORTA (Area Head) ────────────────────────────────────────
  // SM Sta Mesa
  { row:  80, name: 'JOHN LADDIE BOQUIREN',         branch: 'SM Sta. Mesa',     d1_15: 13, d16_31: 14 },
  { row: 356, name: 'MARK ANTHONY PALOMO',          branch: 'SM Sta. Mesa',     d1_15: 12, d16_31: 14 },
  { row: 231, name: 'JOSIE GAZZINGAN',              branch: 'SM Sta. Mesa',     d1_15: 13, d16_31: 14 },
  { row: 498, name: 'LEMUEL GONZALES',              branch: 'SM Sta. Mesa',     d1_15: 14, d16_31: 12 },
  // SM Masinag
  { row: 495, name: 'RICA YLANAN',                  branch: 'SM Masinag',       d1_15: 14, d16_31: 14 },
  { row: 279, name: 'ERICA MAE LLANZANA',           branch: 'SM Masinag',       d1_15: 14, d16_31: 14 },
  { row: 289, name: 'JAN MARCIAL MAESTRE',          branch: 'SM Masinag',       d1_15: 13, d16_31: 14 },
  // Rob Metro East
  { row: 347, name: 'MARK PADILLA',                 branch: 'Rob Metro East',   d1_15: 13, d16_31: 14 },
  { row: 150, name: 'RONIEBERT CORTEZ',             branch: 'Rob Metro East',   d1_15: 13, d16_31: 14 },  // masterlist: Cortez Jr.
  { row: 313, name: 'ARLENE MARTIREZ',              branch: 'Rob Metro East',   d1_15: 14, d16_31: 12 },  // masterlist: Martinez
  // SM Marikina
  { row:  82, name: 'MELANIE BORLAGDAN',            branch: 'SM Marikina',      d1_15: 15, d16_31: 11 },
  { row: 487, name: 'JOHN OLIVER VINARAO',          branch: 'SM Marikina',      d1_15:  4, d16_31: 14 },
  { row:  90, name: 'LEO-MAR BUCAD',                branch: 'SM Marikina',      d1_15: 12, d16_31: 14 },
  { row: 504, name: 'MIKE JIM ACHA',                branch: 'SM Marikina',      d1_15: 13, d16_31: 13 },
  // One Ayala
  { row: 485, name: 'LUCY ANN VINARAO',             branch: 'One Ayala',        d1_15: 13, d16_31: 14 },
  { row: 280, name: 'BENBOY LOPESILLO',             branch: 'One Ayala',        d1_15: 13, d16_31: 14 },
  { row:  94, name: 'ARN BUGHAO',                   branch: 'One Ayala',        d1_15: 13, d16_31: 14 },
  // Victory Antipolo
  { row: 225, name: 'MARK GIL GEORGE GAN',          branch: 'Victory Antipolo', d1_15: 15, d16_31: 15 },
  { row: 304, name: 'ROMMEL MANLAPAS',              branch: 'Victory Antipolo', d1_15: 13, d16_31: 14 },
  { row: null,name: 'MONIQUE JAHS MURO',            branch: 'Victory Antipolo', d1_15: null, d16_31: null }, // NOT IN MASTERLIST
  // Rob Antipolo
  { row: 266, name: 'ROBERT LAGUMBAY',              branch: 'Rob Antipolo',     d1_15: 14, d16_31: 13 },
  { row: 123, name: 'JV CANO',                      branch: 'Rob Antipolo',     d1_15: 14, d16_31: 13 },
  { row: 500, name: 'FRANCIS JOHN ARSENUE',         branch: 'Rob Antipolo',     d1_15: 14, d16_31: 13 },
  { row: null,name: 'MAX JAMMINE SALAS',            branch: 'Rob Antipolo',     d1_15:  0, d16_31:  4 },  // NOT IN MASTERLIST
  // SM Taytay
  { row: 332, name: 'MARY JOY NAVARRO',             branch: 'SM Taytay',        d1_15: 13, d16_31: 13 },
  { row:   3, name: 'ALEXANDER CHRISTIAN ABARQUEZ', branch: 'SM Taytay',        d1_15: 14, d16_31: 12 },
  { row: 306, name: 'MARK ARRETH MANZANO',          branch: 'SM Taytay',        d1_15: 13, d16_31: 13 },
  { row: null,name: 'WENDELLEN RADAZA',             branch: 'SM Taytay',        d1_15:  0, d16_31:  4 },  // NOT IN MASTERLIST
];

async function main() {
  fs.mkdirSync(path.dirname(DEST), { recursive: true });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.getWorksheet('MAsterlist');

  const updated  = [];
  const skipped  = [];

  for (const rec of RECORDS) {
    if (rec.row === null) {
      skipped.push(rec);
      continue;
    }
    const row = ws.getRow(rec.row);
    if (rec.d1_15  !== null && rec.d1_15  !== undefined) {
      row.getCell(5).value = rec.d1_15;
      styleAttCell(row.getCell(5));
    }
    if (rec.d16_31 !== null && rec.d16_31 !== undefined) {
      row.getCell(6).value = rec.d16_31;
      styleAttCell(row.getCell(6));
    }
    row.commit();
    updated.push(rec);
  }

  await wb.xlsx.writeFile(DEST);

  console.log(`\n✅ Updated ${updated.length} employee records in masterlist:`);
  updated.forEach(r => console.log(`  ✓ Row ${String(r.row).padStart(3)} | ${r.name.padEnd(35)} | 1-15: ${String(r.d1_15).padStart(2)}  16-31: ${r.d16_31}`));

  if (skipped.length) {
    console.log(`\n⚠️  ${skipped.length} NOT FOUND in masterlist (need to add manually):`);
    skipped.forEach(r => console.log(`  ✗ ${r.name} | ${r.branch} | 1-15: ${r.d1_15}  16-31: ${r.d16_31}`));
  }

  console.log(`\n📁 Updated file saved to:\n   ${DEST}`);
}

function styleAttCell(cell) {
  cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } };
  cell.font      = { bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border    = {
    top:    { style: 'thin', color: { argb: 'FF1A3C6E' } },
    bottom: { style: 'thin', color: { argb: 'FF1A3C6E' } },
    left:   { style: 'thin', color: { argb: 'FF1A3C6E' } },
    right:  { style: 'thin', color: { argb: 'FF1A3C6E' } },
  };
}

main().catch(console.error);
