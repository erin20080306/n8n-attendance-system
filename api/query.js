const https = require('https');
const crypto = require('crypto');
const nodemailer = require('nodemailer');

// --- Config ---
const SMTP_USER = 'erin20080306@gmail.com';
const SMTP_FROM = '出勤查詢助手 <erin20080306@gmail.com>';
const DEFAULT_EMAIL = 'erin20080306@gmail.com';
const SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly';
const SPREADSHEET_MAP = {
  'TAO1':'1_bhGQdx0YH7lsqPFEq5___6_Nwq_gbelJmIHv0bmaIE',
  'TAO3':'1cffI2jIVZA1uSiAyaLLXXgPzDByhy87xznaN85O7wEE',
  'TAO4':'1tVxQbV0298fn2OXWAF0UqZa7FLbypsatciatxs4YVTU',
  'TAO5':'1jzVXC6gt36hJtlUHoxtTzZLMNj4EtTsd4k8eNB1bdiA',
  'TAO6':'1wwPLSLjl2abfM_OMdTNI9PoiPKo3waCV_y0wmx2DxAE',
  'TAO7':'16nGCqRO8DYDm0PbXFbdt-fiEFZCXxXjlOWjKU67p4LY',
  'TAO10':'1y0w49xdFlHvcVtgtG8fq6zdrF26y8j7HMFh5ujzUyR4',
};
const ALL_WAREHOUSES = Object.keys(SPREADSHEET_MAP);
const INTENT_SHEET_TYPE = { leave_status:'leave', sick_leave_stats:'leave', leave_stats:'leave', attendance_detail:'work', attendance_stats:'work' };
const SHEET_KEYWORDS = { leave:['班表','出勤記錄','假況'], work:['出勤時數','出勤時間','工時'] };
const LEAVE_STATUS_MAP = {
  '病假':'病假','病':'病假','事假':'事假','事':'事假',
  '特休':'特休','特':'特休','曠職':'曠職','曠':'曠職',
  '出勤':'出勤','正常':'出勤','休假':'休假','休':'休假',
  '公假':'公假','公':'公假','產假':'產假','喪假':'喪假','婚假':'婚假',
  '例':'例假','例假':'例假','例休':'例休',
  '未':'未到','離':'離職','離職':'離職',
  '下休(事)':'事假','下休(特休)':'特休','下休(曠)':'曠職',
  '病(無薪)':'病假','無薪病假':'病假',
  '國':'國定假日','國定':'國定假日','國假':'國定假日',
  '補':'補休','補休':'補休',
};
const AV = ['出勤','正常',''];

// --- Google Auth ---
let cachedToken = null, tokenExpiry = 0;

function getSaKey() {
  const raw = process.env.GOOGLE_SA_KEY;
  if (!raw) throw new Error('Missing GOOGLE_SA_KEY env');
  return JSON.parse(raw);
}

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiry) return cachedToken;
  const sa = getSaKey();
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const now = Math.floor(Date.now() / 1000);
  const claim = Buffer.from(JSON.stringify({
    iss: sa.client_email, scope: SCOPE,
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  })).toString('base64url');
  const signInput = `${header}.${claim}`;
  const signature = crypto.createSign('RSA-SHA256').update(signInput).sign(sa.private_key, 'base64url');
  const jwt = `${signInput}.${signature}`;
  return new Promise((resolve, reject) => {
    const postData = `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`;
    const req = https.request({
      hostname: 'oauth2.googleapis.com', path: '/token', method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) },
    }, (res) => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => {
        try {
          const data = JSON.parse(body);
          if (data.access_token) { cachedToken = data.access_token; tokenExpiry = Date.now() + 3500000; resolve(cachedToken); }
          else reject(new Error('Token error: ' + body));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(postData); req.end();
  });
}

function callSheetsApi(path, token) {
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'sheets.googleapis.com', path, method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` },
    }, (res) => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, data: JSON.parse(body) }); }
        catch (e) { resolve({ status: res.statusCode, data: { error: body } }); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// --- Query Parsing ---
function parseQuery(query) {
  const yr = new Date().getFullYear();
  const now = new Date();

  // Date parsing
  let dateInfo;
  let m = query.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s*[~\-到至]\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) { dateInfo = { dateFrom: `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`, dateTo: `${m[4]}-${String(m[5]).padStart(2,'0')}-${String(m[6]).padStart(2,'0')}`, dateMode: 'range' }; }
  else if ((m = query.match(/(\d{1,2})[\/\-](\d{1,2})\s*[~\-到至]\s*(\d{1,2})[\/\-](\d{1,2})/))) {
    dateInfo = { dateFrom: `${yr}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`, dateTo: `${yr}-${String(m[3]).padStart(2,'0')}-${String(m[4]).padStart(2,'0')}`, dateMode: 'range' };
  }
  else if ((m = query.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/))) { dateInfo = { date: `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`, dateMode: 'day' }; }
  else if ((m = query.match(/(\d{1,2})[\/\-](\d{1,2})/))) { dateInfo = { date: `${yr}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`, dateMode: 'day' }; }
  else if ((m = query.match(/(\d{1,2})\s*月/))) { dateInfo = { date: `${yr}-${String(m[1]).padStart(2,'0')}`, dateMode: 'month' }; }
  else { dateInfo = { dateMode: 'all' }; }

  // Intent
  const intentRules = [
    { keywords: ['病假統計'], intent: 'sick_leave_stats' },
    { keywords: ['特休統計'], intent: 'leave_stats' },
    { keywords: ['請假統計'], intent: 'leave_stats' },
    { keywords: ['請假狀況','請假'], intent: 'leave_status' },
    { keywords: ['出勤時數','出勤時間','工時查詢','打卡'], intent: 'attendance_stats' },
    { keywords: ['出勤人員','出勤明細','出勤報表'], intent: 'attendance_detail' },
    { keywords: ['出勤統計'], intent: 'attendance_stats' },
  ];
  let intent = 'leave_status';
  for (const r of intentRules) { for (const k of r.keywords) { if (query.includes(k)) { intent = r.intent; break; } } if (intent !== 'leave_status') break; }

  // Warehouse
  let warehouses;
  if (/全部倉|所有倉|全倉/.test(query)) { warehouses = ALL_WAREHOUSES.slice(); }
  else { const f = []; for (const w of ALL_WAREHOUSES) if (query.toUpperCase().includes(w)) f.push(w); warehouses = f.length ? f : ALL_WAREHOUSES.slice(); }

  // Leave type
  const ltRules = [
    { keywords: ['病假','病'], leaveType: '病假' }, { keywords: ['事假','事'], leaveType: '事假' },
    { keywords: ['特休','特'], leaveType: '特休' }, { keywords: ['曠職','曠'], leaveType: '曠職' },
    { keywords: ['公假'], leaveType: '公假' }, { keywords: ['國定假日','國假'], leaveType: '國定假日' },
  ];
  let leaveType = '';
  for (const r of ltRules) { for (const k of r.keywords) { if (query.includes(k)) { leaveType = r.leaveType; break; } } if (leaveType) break; }

  if (leaveType && ['leave_status','attendance_detail'].includes(intent)) intent = 'leave_stats';

  // Person name extraction: remove known tokens, remaining Chinese chars = name(s)
  let personNames = [];
  let tmp = query;
  // Remove date patterns
  tmp = tmp.replace(/\d{4}[\/-]\d{1,2}[\/-]\d{1,2}/g, '');
  tmp = tmp.replace(/\d{1,2}[\/-]\d{1,2}/g, '');
  tmp = tmp.replace(/\d{1,2}\s*月/g, '');
  tmp = tmp.replace(/[~\-到至]/g, '');
  // Remove warehouse names
  for (const w of ALL_WAREHOUSES) tmp = tmp.replace(new RegExp(w, 'gi'), '');
  // Remove known keywords
  const removeKw = ['全部倉','所有倉','全倉','病假統計','特休統計','請假統計','請假狀況','請假',
    '出勤時數','出勤時間','工時查詢','打卡','出勤人員','出勤明細','出勤報表','出勤統計',
    '病假','事假','特休','曠職','公假','國定假日','國假',
    '假別統計','統計','查詢','的','全部'];
  for (const kw of removeKw) tmp = tmp.replace(new RegExp(kw, 'g'), '');
  tmp = tmp.replace(/[a-zA-Z0-9\s]/g, '').trim();
  // Split by 、,， and filter valid names (2-4 chars)
  if (tmp.length >= 2) {
    const parts = tmp.split(/[、,，]/).map(s => s.trim()).filter(s => s.length >= 2 && s.length <= 4);
    if (parts.length > 0) personNames = parts;
    else if (tmp.length >= 2 && tmp.length <= 4) personNames = [tmp];
  }

  return { ...dateInfo, intent, warehouses, leaveType, personNames, originalQuery: query };
}

// --- Sheet Processing ---
function normalizeH(s) { return s.replace(/[\n\r\s]/g,'').replace(/\(.*?\)/g,'').trim(); }

function findCol(hdr, aliases) {
  for (const a of aliases) { const na = normalizeH(a); const i = hdr.findIndex(z => z && normalizeH(z) === na); if (i !== -1) return i; }
  for (const a of aliases) { const i = hdr.findIndex(z => z && normalizeH(z).includes(normalizeH(a))); if (i !== -1) return i; }
  return -1;
}

function normalizeStatus(v) {
  if (!v || v.trim() === '') return { status: '', ok: true };
  const lv = v.trim();
  if (LEAVE_STATUS_MAP[lv]) return { status: LEAVE_STATUS_MAP[lv], ok: true };
  return { status: lv, ok: false };
}

function processWorkSheet(hdr, raw, warehouse) {
  const cm = {
    date: findCol(hdr, ['日期','Date']),
    department: findCol(hdr, ['組別','部門','Department','組','Group']),
    shift: findCol(hdr, ['班別','Shift','班次']),
    name: findCol(hdr, ['姓名','Name','員工姓名']),
    clockIn: findCol(hdr, ['上班時間','上班','ClockIn']),
    clockOut: findCol(hdr, ['下班時間','下班','ClockOut']),
    workHours: findCol(hdr, ['工作總時數','工時','Hours','總工時']),
    overtimeHours: findCol(hdr, ['加班總時數','OvertimeHours','加班時數','加班']),
    note: findCol(hdr, ['備註','Note','Remark']),
  };
  const rows = [];
  for (const r of raw) {
    const g = f => { const i = cm[f]; return i >= 0 && r[i] !== undefined ? String(r[i]).trim() : ''; };
    if (!g('name')) continue;
    rows.push({ date: g('date'), department: g('department'), shift: g('shift'), name: g('name'),
      clockIn: g('clockIn'), clockOut: g('clockOut'), workHours: g('workHours'), overtimeHours: g('overtimeHours'),
      attendanceStatus: '出勤', leaveType: '', note: g('note'), warehouse });
  }
  return rows;
}

function processLeaveSheet(hdr, raw, warehouse) {
  const ni = findCol(hdr, ['姓名','Name','員工姓名']);
  const di = findCol(hdr, ['組別','部門','Department','組']);
  const si = findCol(hdr, ['班別','Shift','班次']);
  let dsc = -1;
  for (let i = 0; i < hdr.length; i++) { if (hdr[i] && /^\d{1,2}\/\d{1,2}$/.test(hdr[i].trim())) { dsc = i; break; } }
  if (dsc === -1) dsc = 8;
  const dh = hdr.slice(dsc);
  const rows = [];
  for (const r of raw) {
    const pn = ni >= 0 ? String(r[ni] || '').trim() : '';
    const dp = di >= 0 ? String(r[di] || '').trim() : '';
    const sh = si >= 0 ? String(r[si] || '').trim() : '';
    if (!pn) continue;
    for (let i = 0; i < dh.length; i++) {
      const cv = r[dsc + i];
      if (cv === undefined || cv === null || String(cv).trim() === '') continue;
      const s = normalizeStatus(String(cv));
      rows.push({ date: dh[i] ? String(dh[i]).trim() : '', department: dp, shift: sh, name: pn,
        workHours: '', overtimeHours: '', clockIn: '', clockOut: '',
        attendanceStatus: s.status, leaveType: s.status, note: '', warehouse });
    }
  }
  return rows;
}

// --- Stats ---
function toNum(ds) {
  let m = String(ds).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return +m[1] * 10000 + +m[2] * 100 + +m[3];
  m = String(ds).match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) return new Date().getFullYear() * 10000 + +m[1] * 100 + +m[2];
  m = String(ds).match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) return +m[1] * 10000 + +m[2] * 100 + +m[3];
  return 0;
}

function matchDate(rd, parsed) {
  if (!rd) return false;
  rd = String(rd).trim();
  const { dateMode, date, dateFrom, dateTo } = parsed;
  if (dateMode === 'range') { const n = toNum(rd); return n >= toNum(dateFrom) && n <= toNum(dateTo); }
  if (dateMode === 'month') {
    const tM = parseInt((date || '').split('-')[1]);
    let m = rd.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (m) return +m[1] === tM;
    m = rd.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) return +m[2] === tM;
    return false;
  }
  // day
  const p = (date || '').split('-');
  if (p.length !== 3) return false;
  const [tY, tM, tD] = p.map(Number);
  let m = rd.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) return +m[1] === tM && +m[2] === tD;
  m = rd.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return +m[1] === tY && +m[2] === tM && +m[3] === tD;
  return rd === date;
}

function buildDetail(list, showDate) {
  const grouped = {};
  for (const r of list) {
    const t = r.leaveType || r.attendanceStatus || '未分類';
    if (!grouped[t]) grouped[t] = {};
    const key = [r.warehouse, r.department, r.shift, r.name].join('|');
    if (!grouped[t][key]) grouped[t][key] = { info: r, dates: [] };
    if (r.date) grouped[t][key].dates.push(r.date);
  }
  const lines = [];
  for (const [type, people] of Object.entries(grouped)) {
    lines.push('');
    lines.push(`【${type}】${Object.keys(people).length} 人`);
    for (const [, data] of Object.entries(people)) {
      const p = data.info;
      const parts = [p.warehouse, p.department, p.shift, p.name].filter(Boolean);
      let line = '  ' + parts.join(' / ');
      if (showDate && data.dates.length > 0) {
        line += '：' + data.dates.join(', ') + ' 總計：' + data.dates.length + '天';
      }
      lines.push(line);
    }
  }
  return lines.join('\n');
}

function buildWorkDetail(list) {
  const lines = [];
  const byDept = {};
  for (const r of list) { const d = r.department || '未分組'; if (!byDept[d]) byDept[d] = []; byDept[d].push(r); }
  for (const [dept, people] of Object.entries(byDept)) {
    lines.push('');
    lines.push(`【${dept}】${people.length} 筆`);
    for (const p of people) {
      const wh = p.workHours ? p.workHours + 'h' : '-';
      const ot = p.overtimeHours ? '+' + p.overtimeHours + 'h加班' : '';
      const time = (p.clockIn || p.clockOut) ? `（${p.clockIn || '?'}~${p.clockOut || '?'}）` : '';
      const dt = p.date ? p.date + ' ' : '';
      lines.push(`  ${dt}${[p.shift, p.name].filter(Boolean).join('/')} ${wh}${ot} ${time}`);
      if (p.note) lines.push(`    備註：${p.note}`);
    }
  }
  return lines.join('\n');
}

function computeStats(parsed, allRows, wh) {
  const { intent, leaveType: lt, originalQuery: oq, dateMode, date, dateFrom, dateTo } = parsed;
  let dateLabel;
  if (dateMode === 'all') dateLabel = '全部月份';
  else if (dateMode === 'range') dateLabel = `${dateFrom} ~ ${dateTo}`;
  else if (dateMode === 'month') dateLabel = (date || '').replace(/^(\d{4})-0?(\d{1,2})$/, '$1年$2月');
  else dateLabel = date || '';
  const showDate = dateMode !== 'day';

  let df = (dateMode === 'month' || dateMode === 'all') ? allRows : allRows.filter(r => matchDate(r.date, parsed));

  // Person name filter (supports multiple names)
  const pns = parsed.personNames || [];
  if (pns.length > 0) df = df.filter(r => r.name && pns.some(pn => r.name.includes(pn)));

  let res = {};

  switch (intent) {
    case 'leave_status':
    case 'leave_stats': {
      const lr = df.filter(r => { const s = r.attendanceStatus || r.leaveType || ''; return s && !AV.includes(s); });
      const fr = lt ? lr.filter(r => (r.leaveType || r.attendanceStatus) === lt) : lr;
      const detail = buildDetail(fr, showDate);
      const label = lt ? lt + '統計' : '請假統計';
      const uniq = new Set(fr.map(r => r.warehouse + '|' + r.name));
      const nameTag = pns.length ? ` 【${pns.join('、')}】` : '';
      res = { success: true, query: oq, intent, date, dateFrom, dateTo, dateMode,
        summaryText: `${dateLabel} ${label}${nameTag}：共 ${uniq.size} 人（${fr.length} 人次），倉別：${wh.join(', ')}\n${detail}`,
        totals: { totalPeople: uniq.size, totalEntries: fr.length, byWarehouse: {}, byLeaveType: {} }, rows: fr };
      for (const r of fr) { res.totals.byWarehouse[r.warehouse] = (res.totals.byWarehouse[r.warehouse] || 0) + 1; res.totals.byLeaveType[r.leaveType || '未分類'] = (res.totals.byLeaveType[r.leaveType || '未分類'] || 0) + 1; }
      break;
    }
    case 'sick_leave_stats': {
      const sr = df.filter(r => (r.leaveType || r.attendanceStatus || '').trim() === '病假');
      const detail = buildDetail(sr, showDate);
      const uniq = new Set(sr.map(r => r.warehouse + '|' + r.name));
      const nameTag = pns.length ? ` 【${pns.join('、')}】` : '';
      res = { success: true, query: oq, intent, date, dateFrom, dateTo, dateMode,
        summaryText: `${dateLabel} 病假統計${nameTag}：共 ${uniq.size} 人（${sr.length} 人次），倉別：${wh.join(', ')}\n${detail}`,
        totals: { totalPeople: uniq.size, totalEntries: sr.length, byWarehouse: {} }, rows: sr };
      for (const r of sr) res.totals.byWarehouse[r.warehouse] = (res.totals.byWarehouse[r.warehouse] || 0) + 1;
      break;
    }
    case 'attendance_detail': {
      const detail = buildDetail(df, showDate);
      const nameTag = pns.length ? ` 【${pns.join('、')}】` : '';
      res = { success: true, query: oq, intent, date, dateFrom, dateTo, dateMode,
        summaryText: `${dateLabel} 出勤明細${nameTag}：共 ${df.length} 筆，倉別：${wh.join(', ')}\n${detail}`,
        totals: { totalRecords: df.length, byWarehouse: {} }, rows: df };
      for (const r of df) res.totals.byWarehouse[r.warehouse] = (res.totals.byWarehouse[r.warehouse] || 0) + 1;
      break;
    }
    case 'attendance_stats': {
      const detail = buildWorkDetail(df);
      const tw = df.reduce((s, r) => s + (parseFloat(r.workHours) || 0), 0);
      const to = df.reduce((s, r) => s + (parseFloat(r.overtimeHours) || 0), 0);
      const nameTag = pns.length ? ` 【${pns.join('、')}】` : '';
      res = { success: true, query: oq, intent, date, dateFrom, dateTo, dateMode,
        summaryText: `${dateLabel} 出勤時數${nameTag}：${df.length} 人，工時 ${tw.toFixed(1)}h，加班 ${to.toFixed(1)}h，倉別：${wh.join(', ')}\n${detail}`,
        totals: { totalRecords: df.length, totalWorkHours: tw.toFixed(1), totalOvertimeHours: to.toFixed(1), byWarehouse: {} }, rows: df };
      for (const r of df) {
        if (!res.totals.byWarehouse[r.warehouse]) res.totals.byWarehouse[r.warehouse] = { count: 0, wh: 0, ot: 0 };
        res.totals.byWarehouse[r.warehouse].count++; res.totals.byWarehouse[r.warehouse].wh += parseFloat(r.workHours) || 0; res.totals.byWarehouse[r.warehouse].ot += parseFloat(r.overtimeHours) || 0;
      }
      break;
    }
    default: res = { success: false, query: oq, error: '無法識別意圖：' + intent };
  }
  if (res.rows?.length === 0 && res.success) { res.success = false; res.summaryText = `${dateLabel} 查無資料（${wh.join(', ')}）`; }
  res.emailSubject = `[出勤報表] ${dateLabel} ${lt || intent} - ${wh.join(', ')}`;
  res.intent = intent;
  return res;
}

// --- CSV Generation ---
function csvEsc(v) { return '"' + String(v || '').replace(/"/g, '""') + '"'; }

function generateLeaveCsv(rows) {
  if (!rows || rows.length === 0) return '\uFEFF無資料';
  // 按 人+假別 分組，每組列出日期
  const groups = {};
  for (const r of rows) {
    const lt = r.leaveType || r.attendanceStatus || '其他';
    const key = [r.warehouse, r.department, r.shift, r.name, lt].join('|');
    if (!groups[key]) groups[key] = { info: r, leaveType: lt, dates: [] };
    if (r.date) groups[key].dates.push(r.date);
  }
  // 排序日期
  const sortDates = (arr) => arr.sort((a, b) => {
    const pa = a.match(/(\d+)\/(\d+)/), pb = b.match(/(\d+)\/(\d+)/);
    if (pa && pb) return (+pa[1]*100 + +pa[2]) - (+pb[1]*100 + +pb[2]);
    return a.localeCompare(b);
  });
  // 找出最大日期數量（決定欄位數）
  let maxDates = 0;
  for (const g of Object.values(groups)) { sortDates(g.dates); if (g.dates.length > maxDates) maxDates = g.dates.length; }
  const dateHeaders = Array.from({ length: maxDates }, (_, i) => `日期${i + 1}`);
  const headers = ['倉別', '部門', '班別', '姓名', '假別', '天數', ...dateHeaders];
  const csvRows = [headers.map(csvEsc).join(',')];
  for (const [, g] of Object.entries(groups)) {
    const p = g.info;
    const vals = [p.warehouse, p.department, p.shift, p.name, g.leaveType, g.dates.length];
    for (let i = 0; i < maxDates; i++) vals.push(g.dates[i] || '');
    csvRows.push(vals.map(csvEsc).join(','));
  }
  return '\uFEFF' + csvRows.join('\n');
}

function sortDate(a, b) {
  const pa = (a || '').match(/(\d+)\/(\d+)/), pb = (b || '').match(/(\d+)\/(\d+)/);
  if (pa && pb) return (+pa[1]*100 + +pa[2]) - (+pb[1]*100 + +pb[2]);
  return (a || '').localeCompare(b || '');
}

function generateWorkCsv(rows) {
  if (!rows || rows.length === 0) return '\uFEFF無資料';
  // 按姓名 → 倉別 → 日期排序，方便篩選
  const sorted = [...rows].sort((a, b) =>
    (a.name || '').localeCompare(b.name || '') ||
    (a.warehouse || '').localeCompare(b.warehouse || '') ||
    sortDate(a.date, b.date)
  );
  const headers = ['姓名', '日期', '倉別', '部門', '班別', '上班時間', '下班時間', '工時', '加班時數', '備註'];
  const csvRows = [headers.map(csvEsc).join(',')];
  for (const r of sorted) {
    const vals = [r.name, r.date, r.warehouse, r.department, r.shift,
      r.clockIn || '', r.clockOut || '', r.workHours || '', r.overtimeHours || '', r.note || ''];
    csvRows.push(vals.map(csvEsc).join(','));
  }
  return '\uFEFF' + csvRows.join('\n');
}

function generateCsv(rows, intent) {
  if (intent === 'attendance_stats' || intent === 'attendance_detail') return generateWorkCsv(rows);
  return generateLeaveCsv(rows);
}

// --- Email Sending ---
async function sendEmail(subject, summaryText, rows, intent, emails) {
  const smtpPass = process.env.GMAIL_APP_PASSWORD;
  if (!smtpPass) return { success: false, error: '未設定 GMAIL_APP_PASSWORD' };
  const toList = [...new Set([...(emails || []), DEFAULT_EMAIL].filter(Boolean))];
  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com', port: 587, secure: false,
    auth: { user: SMTP_USER, pass: smtpPass },
  });
  // 按倉別排序後產生單一 CSV
  const sorted = [...rows].sort((a, b) => (a.warehouse || '').localeCompare(b.warehouse || ''));
  const csv = generateCsv(sorted, intent);
  const filename = `${subject.replace(/[^\w\u4e00-\u9fff-]/g, '_')}.csv`;
  try {
    await transporter.sendMail({
      from: SMTP_FROM, to: toList.join(','),
      subject, text: summaryText,
      html: `<pre style="font-family:monospace;font-size:14px;white-space:pre-wrap;">${summaryText}</pre>`,
      attachments: [{ filename, content: csv, contentType: 'text/csv; charset=utf-8' }],
    });
    return { success: true, message: `已寄送至 ${toList.join(', ')}` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// --- Main Handler ---
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });

  try {
    const { query, emails } = req.body || {};
    if (!query || !query.trim()) return res.status(400).json({ success: false, error: '請提供查詢指令', hint: '範例：4/2請假狀況、3月特休統計、3/15~4/1特休' });

    const parsed = parseQuery(query.trim());
    const token = await getAccessToken();
    const st = INTENT_SHEET_TYPE[parsed.intent] || 'leave';
    const kw = SHEET_KEYWORDS[st];

    // Determine months to query
    let months = [];
    if (parsed.dateMode === 'range') {
      const mF = parseInt((parsed.dateFrom || '').split('-')[1]) || 0;
      const mT = parseInt((parsed.dateTo || '').split('-')[1]) || 0;
      for (let m = mF; m <= mT; m++) months.push(m);
    } else if (parsed.dateMode !== 'all') {
      const m = parseInt((parsed.date || '').split('-')[1]) || 0;
      if (m) months.push(m);
    }
    months = [...new Set(months)];

    // Fetch data for each warehouse
    const allRows = [];
    const warnings = [];

    for (const wh of parsed.warehouses) {
      const sid = SPREADSHEET_MAP[wh];
      if (!sid) continue;

      // Get sheet list
      const metaRes = await callSheetsApi(`/v4/spreadsheets/${sid}?fields=sheets.properties.title`, token);
      if (metaRes.status !== 200 || !metaRes.data.sheets) { warnings.push(`${wh} 無法取得分頁`); continue; }
      const names = metaRes.data.sheets.map(s => s.properties.title);

      // Find target sheets
      const targets = new Set();
      if (parsed.dateMode === 'all') {
        // 查詢所有符合關鍵字的分頁（各月份）
        for (const name of names) {
          for (const k of kw) { if (name.includes(k)) { targets.add(name); break; } }
        }
        if (targets.size === 0) targets.add(names[0] || 'Sheet1');
      } else {
        for (const mo of months) {
          const monthStr = mo + '月';
          let target = null;
          for (const k of kw) { const f = names.find(n => n.includes(k) && n.includes(monthStr)); if (f) { target = f; break; } }
          if (!target) { for (const k of kw) { const f = names.find(n => n.includes(k)); if (f) { target = f; break; } } }
          if (!target) target = names[0] || 'Sheet1';
          targets.add(target);
        }
      }

      // Fetch and process each sheet
      for (const sheetName of targets) {
        const dataRes = await callSheetsApi(`/v4/spreadsheets/${sid}/values/${encodeURIComponent(sheetName)}?majorDimension=ROWS`, token);
        if (dataRes.status !== 200 || !dataRes.data.values) { warnings.push(`${wh}「${sheetName}」無資料`); continue; }
        const hdr = dataRes.data.values[0] || [];
        const raw = dataRes.data.values.slice(1);
        const rows = st === 'work' ? processWorkSheet(hdr, raw, wh) : processLeaveSheet(hdr, raw, wh);
        allRows.push(...rows);
      }
    }

    const result = computeStats(parsed, allRows, parsed.warehouses);
    if (warnings.length) result.warnings = warnings;

    // Auto-send email report
    if (result.success && result.rows && result.rows.length > 0) {
      try {
        const emailResult = await sendEmail(
          result.emailSubject || '[出勤報表]',
          result.summaryText || '',
          result.rows,
          result.intent || parsed.intent,
          emails || []
        );
        result.email = emailResult;
      } catch (e) {
        result.email = { success: false, error: e.message };
      }
    }

    return res.status(200).json(result);

  } catch (e) {
    return res.status(500).json({ success: false, error: e.message });
  }
};
