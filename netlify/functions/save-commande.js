const https = require('https');

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = '📦 Commandes';
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
    const req = https.request({
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        // >>> On affiche et on vérifie la VRAIE réponse de Google <<<
        console.log('TOKEN endpoint status:', res.statusCode);
        console.log('TOKEN endpoint response:', data);
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) {
            resolve(parsed.access_token);
          } else {
            reject(new Error('TOKEN_REFUSED (' + res.statusCode + '): ' + (parsed.error || '') + ' - ' + (parsed.error_description || data)));
          }
        } catch(e) {
          reject(new Error('TOKEN_PARSE_FAILED: ' + data));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getLastOrderNumber(token) {
  const path = '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(SHEET_NAME + '!A:A');
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'sheets.googleapis.com',
      path: path,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (res.statusCode >= 400 || json.error) {
            const msg = (json.error && json.error.message) ? json.error.message : data;
            return reject(new Error('SHEETS_READ_' + res.statusCode + ': ' + msg));
          }
          const values = json.values || [];
          let lastNum = 0;
          for (let i = 0; i < values.length; i++) {
            const cell = (values[i][0] || '').toString();
            const match = cell.match(/^AMO-(\d+)$/);
            if (match) {
              const n = parseInt(match[1], 10);
              if (n > lastNum) lastNum = n;
            }
          }
          resolve(lastNum);
        } catch(e) { reject(new Error('SHEETS_READ_PARSE: ' + data)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

async function appendRow(token, values) {
  const body = JSON.stringify({ values: [values] });
  const path = '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(SHEET_NAME + '!A5') + ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS';
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'sheets.googleapis.com',
      path: path,
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        console.log('Sheets append status:', res.statusCode);
        console.log('Sheets append response:', data);
        if (res.statusCode >= 400) {
          let msg = data;
          try { const j = JSON.parse(data); if (j.error && j.error.message) msg = j.error.message; } catch(e) {}
          return reject(new Error('SHEETS_APPEND_' + res.statusCode + ': ' + msg));
        }
        try {
          const j = JSON.parse(data);
          if (j.error) return reject(new Error('SHEETS_APPEND_ERR: ' + (j.error.message || data)));
          resolve(j);
        } catch(e) { reject(new Error('SHEETS_APPEND_PARSE: ' + data)); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

exports.handler = async function(event) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: 'Method Not Allowed' };

  try {
    const missing = [];
    if (!SPREADSHEET_ID)  missing.push('GOOGLE_SPREADSHEET_ID');
    if (!CLIENT_EMAIL)    missing.push('GOOGLE_CLIENT_EMAIL');
    if (!process.env.GOOGLE_PRIVATE_KEY) missing.push('GOOGLE_PRIVATE_KEY');
    if (missing.length) throw new Error('ENV_MISSING: ' + missing.join(', '));

    const body = JSON.parse(event.body);
    console.log('Received:', JSON.stringify(body));

    const { prenom, nom, email, produits, quantite, adresse, livraison, comment } = body;

    const produitsFormate = (produits || '').split(' | ').map(function(p, idx) {
      var qties = (quantite || '').split(' | ');
      var q = qties[idx] ? qties[idx] : '';
      return '• ' + p + ' ' + q;
    }).join('\n');

    const token = await getAccessToken();
    console.log('Token VRAIMENT OK');

    const lastNum = await getLastOrderNumber(token);
    const nextNum = lastNum + 1;
    const numCommande = 'AMO-' + String(nextNum).padStart(3, '0');
    console.log('Numéro commande:', numCommande);

    const scanUrl = 'https://amoria-glow-shop.netlify.app/scanner.html?commande=' + numCommande;
    const barcodeFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + encodeURIComponent(scanUrl) + '")';

    await appendRow(token, [
      numCommande,
      (prenom || '') + ' ' + (nom || ''),
      email || '',
      produitsFormate,
      (quantite || '').split(' | ').join('\n'),
      adresse || '',
      livraison === 'Point relais' ? livraison : '',
      '',
      'À préparer',
      '',
      barcodeFormula,
      comment || ''
    ]);

    return { statusCode: 200, headers, body: JSON.stringify({ success: true, numCommande: numCommande }) };

  } catch (err) {
    console.log('ERROR:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ success: false, error: err.message }) };
  }
};
