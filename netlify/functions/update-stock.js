const https = require('https');

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = 'Qr Code';
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;

async function getAccessToken() {
  const privateKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n').replace(/\r/g, '');
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  })).toString('base64url');
  const crypto = require('crypto');
  const sign = crypto.createSign('RSA-SHA256');
  sign.update(header + '.' + payload);
  const signature = sign.sign(privateKey, 'base64url');
  const jwt = header + '.' + payload + '.' + signature;
  const body = 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt;
  return new Promise((resolve, reject) => {
    const req = https.request({ hostname: 'oauth2.googleapis.com', path: '/token', method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => { try { resolve(JSON.parse(data).access_token); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getSheetData(token) {
  return new Promise((resolve, reject) => {
    const req = https.request({ hostname: 'sheets.googleapis.com', path: '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(SHEET_NAME + '!A:D'), method: 'GET', headers: { 'Authorization': 'Bearer ' + token } }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.end();
  });
}

async function updateCells(token, updates) {
  const body = JSON.stringify({ valueInputOption: 'RAW', data: updates });
  return new Promise((resolve, reject) => {
    const req = https.request({ hostname: 'sheets.googleapis.com', path: '/v4/spreadsheets/' + SPREADSHEET_ID + '/values:batchUpdate', method: 'POST', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) } }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

exports.handler = async function(event) {
  const headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'POST, OPTIONS' };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: 'Method Not Allowed' };
  try {
    const { items } = JSON.parse(event.body);
    if (!items || !items.length) return { statusCode: 400, headers, body: JSON.stringify({ error: 'No items' }) };
    const token = await getAccessToken();
    const sheetData = await getSheetData(token);
    const rows = sheetData.values || [];
    const updates = [];
    for (const item of items) {
      for (let i = 3; i < rows.length; i++) {
        if (rows[i] && rows[i][0] && rows[i][0].trim() === item.ref.trim()) {
          const currentStock = parseInt(rows[i][3] || '0', 10);
          const newStock = Math.max(0, currentStock - item.qty);
          updates.push({ range: SHEET_NAME + '!D' + (i + 1), values: [[newStock]] });
          console.log('Stock ' + item.ref + ': ' + currentStock + ' -> ' + newStock);
          break;
        }
      }
    }
    if (updates.length > 0) await updateCells(token, updates);
    return { statusCode: 200, headers, body: JSON.stringify({ success: true, updated: updates.length }) };
  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
