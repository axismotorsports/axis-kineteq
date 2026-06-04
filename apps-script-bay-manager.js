// ═══════════════════════════════════════════════════════════════════
//  KINETEQ · AXIS BAY MANAGER — Google Apps Script  v3
//  Paste entire file into Code.gs → Deploy as Web App
//  (Execute as: Me · Access: Anyone)
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1a88PFWRrRvwN72CiLbvIYGxyI3GFIqMe58Hzu9Y2CSs';
const TZ       = 'Asia/Bangkok';
const SECRET   = 'K1N3T3Q-AX15-2026'; // ← CHANGE THIS — must match API_SECRET in booking HTML

const COL = {
  timestamp:  0,   // ประทับเวลา
  jobId:      1,   // Job ID
  name:       2,   // First-name
  car:        3,   // Car
  jobType:    4,   // Job Type
  location:   5,   // Location
  date:       6,   // Appointment Date
  startTime:  7,   // Start Time
  endTime:    8,   // End Time
  endDate:    9,   // End Date
  delivDate:  10,  // Delivery Date
  notes:      11,  // Notes
  status:     12,  // Status
  mechanic:   13   // Mechanic (NEW)
};

const HEADERS = [
  'ประทับเวลา',
  'Job ID',
  'First-name (ชื่อจริง)',
  'Car: Brand+Model (ยี่ห้อ+รุ่นรถ)',
  'Job Type',
  'Location',
  'Appointment Date (วันที่จะทำ)',
  'Start Time',
  'End Time',
  'End Date (วันที่เสร็จ)',
  'Delivery Date',
  'Notes',
  'Status',
  'Mechanic'
];

// ── ENTRY POINT ───────────────────────────────────────
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || 'get';

  // ── WRITE OPERATIONS require the secret key ────────
  // GET (read) is public — calendar and live queue need it
  // book, cancel, seed all require ?secret=... matching SECRET above
  if (action === 'book' || action === 'cancel' || action === 'seed') {
    if (params.secret !== SECRET) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  let result;
  try {
    if      (action === 'book')   result = createBooking(params);
    else if (action === 'cancel') result = cancelBooking(params);
    else if (action === 'seed')   result = seedAll(params);
    else                          result = getBookings();
  } catch(err) {
    result = { success: false, message: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET BOOKINGS ──────────────────────────────────────
function getBookings() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    HEADERS.forEach((h, j) => {
      const val = row[j];
      if (val instanceof Date) {
        if (val.getFullYear() <= 1900) {
          // Utilities.formatDate('Asia/Bangkok') uses Bangkok's 1899 historical offset
          // (UTC+6:42, not UTC+7:00), causing an 18-minute drift on time-only values.
          // Fix: shift the Date by +7h manually, then format in UTC to get the correct time.
          const bkkMs   = val.getTime() + 7 * 3600 * 1000;
          const bkkDate = new Date(bkkMs);
          obj[h] = Utilities.formatDate(bkkDate, 'UTC', 'HH:mm');
        } else {
          obj[h] = Utilities.formatDate(val, TZ, 'dd/MM/yyyy');
        }
      } else {
        obj[h] = val;
      }
    });
    obj['_row'] = i + 1;
    rows.push(obj);
  }
  return rows;
}

// ── CREATE BOOKING ────────────────────────────────────
function createBooking(p) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  const name      = p.name      || '';
  const car       = p.car       || '';
  const jobType   = p.job       || '';
  const location  = p.location  || '';
  const dateStr   = p.date      || '';
  const startStr  = p.start     || '';
  const endStr    = p.end       || '';
  const endDate   = p.deliveryDate || '';
  const notes     = p.notes     || '';
  const mechanic  = p.mechanic  || '';

  const isOffsite = (loc) => {
    const l = (loc || '').toLowerCase();
    return l.includes('off-site') || l.includes('offsite') || l.includes('off site');
  };
  const isGallery = (loc) => {
    const l = (loc || '').toLowerCase();
    return l.includes('gallery') || l.includes('rama');
  };
  const isMoto = (loc) => {
    const l = (loc || '').toLowerCase();
    return l.includes('motorsport');
  };

  const newDate  = parseDateStr(dateStr);
  const newEnd   = endDate ? parseDateStr(endDate) : newDate;
  const newStart = parseTimeStr(startStr);
  const newEndT  = parseTimeStr(endStr);

  // ── BAY CONFLICT CHECK (skip for off-site) ─────────
  if (!isOffsite(location)) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[COL.status]) === 'Cancelled') continue;
      const rowLoc = String(row[COL.location] || '');
      if (isOffsite(rowLoc)) continue;

      const newIsGallery = isGallery(location) || (!isMoto(location) && !isOffsite(location));
      const rowIsGallery = isGallery(rowLoc)   || (!isMoto(rowLoc)   && !isOffsite(rowLoc));
      if (newIsGallery !== rowIsGallery) continue;

      const rowDate = parseDateVal(row[COL.date]);
      const rowEndD = row[COL.endDate] ? parseDateVal(row[COL.endDate]) : rowDate;
      if (!rowDate) continue;
      if (newDate > (rowEndD || rowDate) || newEnd < rowDate) continue;

      const rowStart = parseTimeVal(row[COL.startTime]);
      const rowEnd   = parseTimeVal(row[COL.endTime]);
      if (rowStart === null || rowEnd === null) continue;
      if (newStart < rowEnd && newEndT > rowStart) {
        return {
          success: false, conflict: true,
          message: `Bay occupied ${row[COL.startTime]}–${row[COL.endTime]} (${row[COL.car]})`
        };
      }
    }
  }

  // ── MECHANIC CONFLICT CHECK ────────────────────────
  // A mechanic can only be at one place at a time, regardless of location
  const mechanics = mechanic.split(',').map(m => m.trim()).filter(Boolean);
  if (mechanics.length > 0 && newStart !== null && newEndT !== null) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[COL.status]) === 'Cancelled') continue;
      const rowMechs = String(row[COL.mechanic] || '').split(',').map(m => m.trim()).filter(Boolean);
      if (rowMechs.length === 0) continue;
      const shared = mechanics.filter(m => rowMechs.includes(m));
      if (shared.length === 0) continue;

      const rowDate = parseDateVal(row[COL.date]);
      const rowEndD = row[COL.endDate] ? parseDateVal(row[COL.endDate]) : rowDate;
      if (!rowDate || !newDate) continue;
      if (newDate > (rowEndD || rowDate) || newEnd < rowDate) continue;

      const rowStart = parseTimeVal(row[COL.startTime]);
      const rowEnd   = parseTimeVal(row[COL.endTime]);
      if (rowStart === null || rowEnd === null) continue;
      if (newStart < rowEnd && newEndT > rowStart) {
        return {
          success: false, conflict: true, mechanicConflict: true,
          message: `${shared.join(' & ')} already assigned: ${row[COL.name]} · ${row[COL.car]} · ${row[COL.startTime]}–${row[COL.endTime]}`
        };
      }
    }
  }

  // ── GENERATE JOB ID  AXIS-DDMMYY-NNN ──────────────
  const now = new Date();
  const dd  = Utilities.formatDate(now, TZ, 'ddMMYY');
  const existingIds = data.slice(1).map(r => String(r[COL.jobId] || ''));
  let seq = 1;
  while (existingIds.includes(`AXIS-${dd}-${String(seq).padStart(3,'0')}`)) seq++;
  const jobId = `AXIS-${dd}-${String(seq).padStart(3,'0')}`;

  sheet.appendRow([
    Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm:ss'),
    jobId, name, car, jobType, location,
    dateStr, startStr, endStr,
    endDate, endDate,
    notes, 'Active',
    mechanic
  ]);

  return { success: true, jobId };
}

// ── CANCEL BOOKING ────────────────────────────────────
function cancelBooking(p) {
  const row = parseInt(p.row);
  if (!row || isNaN(row)) return { success: false, message: 'No row specified' };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sheet.getRange(row, COL.status + 1).setValue('Cancelled');
  return { success: true };
}

// ── DATE / TIME HELPERS ───────────────────────────────
function parseDateStr(str) {
  if (!str) return null;
  const m = String(str).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  const y = String(str).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (y) return new Date(+y[1], +y[2]-1, +y[3]);
  return null;
}
function parseDateVal(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  return parseDateStr(String(val));
}
function parseTimeStr(str) {
  if (!str) return null;
  const m = String(str).match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}
function parseTimeVal(val) {
  if (!val) return null;
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  return parseTimeStr(String(val));
}

// ═══════════════════════════════════════════════════════════════════
//  SEED DATA
//  ?action=seed&month=may   → seed May only (10 rows)
//  ?action=seed&month=june  → seed June only (3 rows)
//  ?action=seed&month=all   → seed both (13 rows)
// ═══════════════════════════════════════════════════════════════════
function seedAll(params) {
  const month = (params && params.month) ? params.month.toLowerCase() : 'may';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  let count = 0;
  if (month === 'may'  || month === 'all') count += seedMay(sheet);
  if (month === 'june' || month === 'all') count += seedJune(sheet);

  return { success: true, message: `Seeded ${count} rows (month: ${month})` };
}

// ── MAY 2026 — Source: May 2026 Schedule PDF — 10 Jobs ──────────
// Row format: [timestamp, jobId, name, car, jobType, location, date,
//              startTime, endTime, endDate, delivDate, notes, status, mechanic]
function seedMay(sheet) {
  const rows = [
    // K'Tor — Honda S660 — Full Brake + Exhaust — May 5 full day
    ['05/05/2026 09:00:00','AXIS-050526-001',"K'Tor",
     'Honda S660','Intake & Exhaust Install','🔧 AXIS Motorsports',
     '05/05/2026','09:00','18:00','05/05/2026','05/05/2026',
     'Full Brake System + Exhaust','Active',''],

    // K'Kob — Civic FL5 — Engine Build — May 6→31 PM only
    ['06/05/2026 13:00:00','AXIS-060526-002',"K'Kob",
     'Honda Civic FL5','Full Engine Build / Rebuild','🔧 AXIS Motorsports',
     '06/05/2026','13:00','18:00','31/05/2026','31/05/2026',
     'Exhaust · Fuel Pump · Engine Reinstall · Startup','Active',''],

    // K'Lhersak — Civic FE — Group M Air Intake — May 8 AM
    ['08/05/2026 09:00:00','AXIS-080526-004',"K'Lhersak",
     'Honda Civic FE','Intake & Exhaust Install','🔧 AXIS Motorsports',
     '08/05/2026','09:00','12:00','08/05/2026','08/05/2026',
     'Group M Air Intake Installation','Active',''],

    // Waren — GR86 — Coilover — May 12–14 AM
    ['12/05/2026 09:00:00','AXIS-120526-003','Waren',
     'Toyota GR86','Coilover Install','🔧 AXIS Motorsports',
     '12/05/2026','09:00','12:00','14/05/2026','14/05/2026',
     'Coilover Installation','Active',''],

    // Hunter Automotive — FL5 — Off-site — May 26
    ['26/05/2026 10:00:00','AXIS-260526-005','Hunter Automotive',
     'Honda Civic FL5','Bolt-on Performance Parts','Off-site (Hunter Automotive)',
     '26/05/2026','10:00','15:00','26/05/2026','26/05/2026',
     'X-box + Rear Wing + Repower','Active',''],

    // K'Nontakorn — GR86 — Zestek — May 26
    ['26/05/2026 10:00:00','AXIS-250526-007',"K'Nontakorn",
     'Toyota GR86','⚡ ZESTEK — Steering Wheel Application','🔧 AXIS Motorsports',
     '26/05/2026','10:00','12:00','26/05/2026','26/05/2026',
     'Zestek Installation','Active',''],

    // FK Gen1 (Ta) — Civic FK — Zestek — May 27
    ['27/05/2026 12:00:00','AXIS-270526-009','FK Gen1 (Ta)',
     'Honda Civic FK','⚡ ZESTEK — Steering Wheel Application','🔧 AXIS Motorsports',
     '27/05/2026','12:00','14:00','27/05/2026','27/05/2026',
     'Zestek Installation','Active',''],

    // Zestek Giveaway — Off-site Rama9 — May 27
    ['27/05/2026 13:00:00','AXIS-270526-010','Zestek Giveaway',
     '—','⚡ ZESTEK — Steering Wheel Application','Off-site (Rama9)',
     '27/05/2026','13:00','14:00','27/05/2026','27/05/2026',
     'Prize Pickup · Customer Collection','Active',''],

    // K'Thanasak — BMW E-series — Zestek — May 28
    ['28/05/2026 10:00:00','AXIS-280526-006',"K'Thanasak",
     'BMW E-series','⚡ ZESTEK — Steering Wheel Application','🔧 AXIS Motorsports',
     '28/05/2026','10:00','12:00','28/05/2026','28/05/2026',
     "Zestek Installation · Drop-off Cust. 10:00",'Active',''],

    // Monet — BMW F32 420i — Moty's Oil Change — May 29
    ['29/05/2026 13:00:00','AXIS-290526-011','Monet',
     "BMW F32 420i","🛢️ MOTY'S — Engine Oil Change",'🔧 AXIS Motorsports',
     '29/05/2026','13:00','15:00','29/05/2026','29/05/2026',
     "Moty's Oil Change Service",'Active',''],
  ];
  rows.forEach(r => sheet.appendRow(r));
  return rows.length;
}

// ── JUNE 2026 — Source: June 2026 Schedule PDF — 3 Jobs ─────────
function seedJune(sheet) {
  const rows = [
    // K'Kob — FL5 continuation — AXIS Motorsports — Jun 1–2 PM
    ['01/06/2026 13:00:00','AXIS-060526-002',"K'Kob",
     'Honda Civic FL5','Full Engine Build / Rebuild','🔧 AXIS Motorsports',
     '01/06/2026','13:00','18:00','02/06/2026','02/06/2026',
     'Continuation from May · Exhaust + Startup','Active',''],

    // P'Nut (Outsource) — TBD — Off-site Ayutthaya — Jun 4 full day
    ['04/06/2026 09:00:00','AXIS-040626-001',"P'Nut (Outsource)",
     'TBD','Bolt-on Performance Parts','Off-site (Ayutthaya)',
     '04/06/2026','09:00','18:00','04/06/2026','04/06/2026',
     'Uninstall Parts · 3D Scan Preparation','Active',''],

    // Toyota VOXY — TOM'S Full Exterior — AXIS Motorsports — Jun 9 full day
    ['09/06/2026 09:00:00','AXIS-060626-003','Toyota VOXY',
     'Toyota VOXY',"🏎️ TOM'S — Oil Change",'🔧 AXIS Motorsports',
     '09/06/2026','09:00','18:00','09/06/2026','09/06/2026',
     "TOM'S Full Exterior Package",'Active',''],
  ];
  rows.forEach(r => sheet.appendRow(r));
  return rows.length;
}
