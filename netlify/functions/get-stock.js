// Lit le stock DEPUIS GOOGLE SHEETS (onglet "Stock")
// Colonne A = Référence | Colonne D = Stock Actuel | données à partir de la ligne 4
// Renvoie au site un objet { "AG-GL-101-001": 2, "AG-BL-102-077": 2, ... }

const https = require('https');

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const SHEET_NAME = 'Stock';

async function getAccessToken() {
  const privateKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n').replace(/\r/g, '');
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
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
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) resolve(parsed.access_token);
          else reject(new Error('TOKEN_REFUSED (' + res.statusCode + '): ' + (parsed.error_description || data)));
        } catch(e) { reject(new Error('TOKEN_PARSE_FAILED: ' + data)); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function readRange(token, range) {
  const path = '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(range);
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
          resolve(json.values || []);
        } catch(e) { reject(new Error('SHEETS_READ_PARSE: ' + data)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

exports.handler = async function(event, context) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    const token = await getAccessToken();

    // Colonne A (réf) et colonne D (stock actuel), à partir de la ligne 4
    const rows = await readRange(token, SHEET_NAME + '!A4:D');

    const stock = {};
    for (let i = 0; i < rows.length; i++) {
      const ref = (rows[i][0] || '').toString().trim();       // colonne A
      const qtyRaw = (rows[i][3] !== undefined ? rows[i][3] : '').toString().trim(); // colonne D
      if (!ref) continue;
      const qty = parseInt(qtyRaw, 10);
      stock[ref] = isNaN(qty) ? 0 : qty;
    }

    console.log('Stock lu:', Object.keys(stock).length, 'produits');
    return { statusCode: 200, headers, body: JSON.stringify(stock) };

  } catch (err) {
    console.log('ERROR get-stock:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
