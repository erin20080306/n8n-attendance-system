#!/usr/bin/env node
// 用 Service Account JSON 金鑰產生 Google API Access Token
// 使用方式：node get-token.js /path/to/service-account.json
const crypto = require('crypto');
const https = require('https');

const keyPath = process.argv[2] || '/Users/erin20080306gmail.com/Downloads/erp-glitch-reader-83efc9a7bf5b.json';
const sa = require(keyPath);

// 建立 JWT
const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
const now = Math.floor(Date.now() / 1000);
const claim = Buffer.from(JSON.stringify({
  iss: sa.client_email,
  scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
  aud: 'https://oauth2.googleapis.com/token',
  iat: now,
  exp: now + 3600,
})).toString('base64url');

const signInput = `${header}.${claim}`;
const signature = crypto.createSign('RSA-SHA256').update(signInput).sign(sa.private_key, 'base64url');
const jwt = `${signInput}.${signature}`;

// 換取 access token
const postData = `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`;
const req = https.request({
  hostname: 'oauth2.googleapis.com',
  path: '/token',
  method: 'POST',
  headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) },
}, (res) => {
  let body = '';
  res.on('data', (c) => body += c);
  res.on('end', () => {
    try {
      const data = JSON.parse(body);
      if (data.access_token) {
        process.stdout.write(data.access_token);
      } else {
        process.stderr.write(JSON.stringify(data));
        process.exit(1);
      }
    } catch (e) {
      process.stderr.write(body);
      process.exit(1);
    }
  });
});
req.on('error', (e) => { process.stderr.write(e.message); process.exit(1); });
req.write(postData);
req.end();
