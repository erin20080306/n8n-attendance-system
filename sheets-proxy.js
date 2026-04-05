#!/usr/bin/env node
// ============================================================
// Google Sheets API Proxy + Email 報表伺服器
// 啟動：node sheets-proxy.js
// ============================================================
const http = require('http');
const https = require('https');
const crypto = require('crypto');
const url = require('url');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

const PUBLIC_DIR = path.join(__dirname, 'public');
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.css': 'text/css; charset=utf-8',
  '.ico': 'image/x-icon',
};

const PORT = 3456;
const SA_KEY_PATH = '/Users/erin20080306gmail.com/Downloads/erp-glitch-reader-83efc9a7bf5b.json';
const SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly';

// Email 設定（Gmail SMTP）
const EMAIL_CONFIG = {
  to: 'erin20080306@gmail.com',
  from: 'erin20080306@gmail.com',
  smtpHost: 'smtp.gmail.com',
  smtpPort: 587,
  smtpUser: 'erin20080306@gmail.com',
  smtpPass: process.env.GMAIL_APP_PASSWORD || '',
};

const sa = require(SA_KEY_PATH);

// --- Token 快取 ---
let cachedToken = null;
let tokenExpiry = 0;

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiry) return cachedToken;
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
    req.write(postData);
    req.end();
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

// --- CSV 產生 ---
function esc(v) { return '"' + String(v || '').replace(/"/g, '""') + '"'; }

// 假別報表：同人合併一列，日期橫向展開
function generateLeaveCsv(rows) {
  if (!rows || rows.length === 0) return '\uFEFF無資料';
  // 收集所有日期並排序
  const allDates = [...new Set(rows.map(r => r.date).filter(Boolean))];
  allDates.sort((a, b) => {
    const pa = a.match(/(\d+)\/(\d+)/), pb = b.match(/(\d+)\/(\d+)/);
    if (pa && pb) return (+pa[1]*100 + +pa[2]) - (+pb[1]*100 + +pb[2]);
    return a.localeCompare(b);
  });
  // 按人分組
  const people = {};
  for (const r of rows) {
    const key = [r.warehouse, r.department, r.shift, r.name].join('|');
    if (!people[key]) people[key] = { info: r, dates: {} };
    if (r.date) people[key].dates[r.date] = r.leaveType || r.attendanceStatus || 'V';
  }
  const headers = ['倉別', '部門', '班別', '姓名', '假別', '天數', ...allDates];
  const csvRows = [headers.map(esc).join(',')];
  for (const [, data] of Object.entries(people)) {
    const p = data.info;
    const leaveType = p.leaveType || p.attendanceStatus || '';
    const count = Object.keys(data.dates).length;
    const vals = [p.warehouse, p.department, p.shift, p.name, leaveType, count];
    for (const d of allDates) {
      vals.push(data.dates[d] || '');
    }
    csvRows.push(vals.map(esc).join(','));
  }
  return '\uFEFF' + csvRows.join('\n');
}

// 出勤時數報表：含上下班時間、工時、加班
function generateWorkCsv(rows) {
  if (!rows || rows.length === 0) return '\uFEFF無資料';
  const headers = ['日期', '倉別', '部門', '班別', '姓名', '上班時間', '下班時間', '工時', '加班時數', '備註'];
  const csvRows = [headers.map(esc).join(',')];
  for (const r of rows) {
    const vals = [
      r.date, r.warehouse, r.department, r.shift, r.name,
      r.clockIn || '', r.clockOut || '',
      r.workHours || '', r.overtimeHours || '', r.note || '',
    ];
    csvRows.push(vals.map(esc).join(','));
  }
  return '\uFEFF' + csvRows.join('\n');
}

function generateCsv(rows, intent) {
  if (intent === 'attendance_stats' || intent === 'attendance_detail') {
    return generateWorkCsv(rows);
  }
  return generateLeaveCsv(rows);
}

// --- Email 寄送 ---
async function sendEmailReport(subject, summaryText, rows, intent, emails) {
  if (!EMAIL_CONFIG.smtpPass) {
    return { success: false, error: '尚未設定 GMAIL_APP_PASSWORD 環境變數' };
  }
  // 合併預設收件人和自訂收件人
  const toList = [...new Set([...(emails || []), EMAIL_CONFIG.to].filter(Boolean))];
  const transporter = nodemailer.createTransport({
    host: EMAIL_CONFIG.smtpHost, port: EMAIL_CONFIG.smtpPort, secure: false,
    auth: { user: EMAIL_CONFIG.smtpUser, pass: EMAIL_CONFIG.smtpPass },
  });
  const csv = generateCsv(rows, intent);
  const filename = `${subject.replace(/[^\w\u4e00-\u9fff-]/g, '_')}.csv`;
  try {
    await transporter.sendMail({
      from: EMAIL_CONFIG.from, to: toList.join(','),
      subject: subject,
      text: summaryText,
      html: `<pre style="font-family:monospace;font-size:14px;white-space:pre-wrap;">${summaryText}</pre>`,
      attachments: [{ filename, content: csv, contentType: 'text/csv; charset=utf-8' }],
    });
    return { success: true, message: `已寄送至 ${toList.join(', ')}` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// --- 讀取 POST body ---
function readBody(req) {
  return new Promise((resolve) => {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch { resolve({}); } });
  });
}

// --- HTTP Server ---
const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // GET /sheets/meta?spreadsheetId=xxx
  if (parsed.pathname === '/sheets/meta') {
    const sid = parsed.query.spreadsheetId;
    if (!sid) { res.writeHead(400); res.end(JSON.stringify({ error: 'Missing spreadsheetId' })); return; }
    try {
      const token = await getAccessToken();
      const result = await callSheetsApi(`/v4/spreadsheets/${sid}?fields=sheets.properties.title`, token);
      res.writeHead(result.status); res.end(JSON.stringify(result.data));
    } catch (e) { res.writeHead(500); res.end(JSON.stringify({ error: e.message })); }
    return;
  }

  // GET /sheets/data?spreadsheetId=xxx&sheetName=yyy
  if (parsed.pathname === '/sheets/data') {
    const sid = parsed.query.spreadsheetId;
    const sn = parsed.query.sheetName;
    if (!sid || !sn) { res.writeHead(400); res.end(JSON.stringify({ error: 'Missing spreadsheetId or sheetName' })); return; }
    try {
      const token = await getAccessToken();
      const result = await callSheetsApi(`/v4/spreadsheets/${sid}/values/${encodeURIComponent(sn)}?majorDimension=ROWS`, token);
      res.writeHead(result.status); res.end(JSON.stringify(result.data));
    } catch (e) { res.writeHead(500); res.end(JSON.stringify({ error: e.message })); }
    return;
  }

  // POST /email — 寄送報表
  if (parsed.pathname === '/email' && req.method === 'POST') {
    const body = await readBody(req);
    const { subject, summaryText, rows, intent, emails } = body;
    if (!subject || !rows) { res.writeHead(400); res.end(JSON.stringify({ error: 'Missing subject or rows' })); return; }
    const result = await sendEmailReport(subject, summaryText || '', rows, intent || '', emails || []);
    res.writeHead(result.success ? 200 : 500);
    res.end(JSON.stringify(result));
    return;
  }

  // GET /health
  if (parsed.pathname === '/health') {
    res.writeHead(200);
    res.end(JSON.stringify({ status: 'ok', emailConfigured: !!EMAIL_CONFIG.smtpPass }));
    return;
  }

  // 靜態檔案（查詢介面 PWA）
  let filePath = parsed.pathname === '/' ? '/index.html' : parsed.pathname;
  const fullPath = path.join(PUBLIC_DIR, filePath);
  if (fullPath.startsWith(PUBLIC_DIR) && fs.existsSync(fullPath) && fs.statSync(fullPath).isFile()) {
    const ext = path.extname(fullPath);
    const mime = MIME_TYPES[ext] || 'application/octet-stream';
    res.setHeader('Content-Type', mime);
    res.writeHead(200);
    fs.createReadStream(fullPath).pipe(res);
    return;
  }

  res.writeHead(404);
  res.end(JSON.stringify({ error: 'Not found' }));
});

server.listen(PORT, () => {
  console.log(`Sheets Proxy + Email 已啟動：http://localhost:${PORT}`);
  console.log(`Service Account: ${sa.client_email}`);
  console.log(`Email 收件人：${EMAIL_CONFIG.to}`);
  console.log(`Email 狀態：${EMAIL_CONFIG.smtpPass ? '✅ 已設定' : '⚠️  未設定 GMAIL_APP_PASSWORD'}`);
  console.log('端點：');
  console.log('  GET  /sheets/meta?spreadsheetId=xxx');
  console.log('  GET  /sheets/data?spreadsheetId=xxx&sheetName=yyy');
  console.log('  POST /email  { subject, summaryText, rows }');
  console.log('  GET  /health');
  console.log('');
  console.log('📱 查詢介面：http://localhost:' + PORT);
  console.log('   可安裝至手機/電腦桌面（PWA）');
});
