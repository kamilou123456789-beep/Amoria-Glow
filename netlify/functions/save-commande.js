const https = require('https');

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = '📦 Commandes';
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const FIRST_DATA_ROW = 6; // Les commandes commencent à la ligne 6

// ORDRE DES COLONNES :
// A = Statut | B = ID Commande | C = Nom Client | D = Numéro (téléphone) | E = Email
// F = Produit | G = Quantité | H = Adresse | I = Point Relais | J = Poids
// K = N° Suivi | L = Code-barres | M = Notes

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
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) resolve(parsed.access_token);
          else reject(new Error('TOKEN_REFUSED (' + res.statusCode + '): ' + (parsed.error || '') + ' - ' + (parsed.error_description || data)));
        } catch(e) { reject(new Error('TOKEN_PARSE_FAILED: ' + data)); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Compte les commandes en lisant la COLONNE B (ID Commande)
async function readOrdersColumnB(token) {
  const path = '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(SHEET_NAME + '!B' + FIRST_DATA_ROW + ':B');
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
          let count = 0;
          let lastNum = 0;
          for (let i = 0; i < values.length; i++) {
            const cell = (values[i][0] || '').toString().trim();
            if (cell === '') continue;
            count++;
            const match = cell.match(/^AMO-(\d+)$/);
            if (match) {
              const n = parseInt(match[1], 10);
              if (n > lastNum) lastNum = n;
            }
          }
          resolve({ count: count, lastNum: lastNum });
        } catch(e) { reject(new Error('SHEETS_READ_PARSE: ' + data)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// Écrit de la colonne B à M (on NE touche PAS la colonne A = Statut)
async function writeRowAt(token, rowNumber, values) {
  const range = SHEET_NAME + '!B' + rowNumber + ':M' + rowNumber;
  const body = JSON.stringify({ values: [values] });
  const path = '/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + encodeURIComponent(range) + '?valueInputOption=USER_ENTERED';
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'sheets.googleapis.com',
      path: path,
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        console.log('Sheets write status:', res.statusCode, '-> ligne', rowNumber);
        console.log('Sheets write response:', data);
        if (res.statusCode >= 400) {
          let msg = data;
          try { const j = JSON.parse(data); if (j.error && j.error.message) msg = j.error.message; } catch(e) {}
          return reject(new Error('SHEETS_WRITE_' + res.statusCode + ': ' + msg));
        }
        resolve(data);
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

    const { prenom, nom, telephone, email, produits, quantite, adresse, livraison, comment } = body;

    const produitsFormate = (produits || '').split(' | ').map(function(p, idx) {
      var qties = (quantite || '').split(' | ');
      var q = qties[idx] ? qties[idx] : '';
      return '• ' + p + ' ' + q;
    }).join('\n');

    const token = await getAccessToken();
    console.log('Token OK');

    const info = await readOrdersColumnB(token);
    const targetRow = FIRST_DATA_ROW + info.count;
    const nextNum = info.lastNum + 1;
    const numCommande = 'AMO-' + String(nextNum).padStart(3, '0');
    console.log('Écriture ligne', targetRow, '| numéro', numCommande);

    const scanUrl = 'https://amoria-glow-shop.netlify.app/scanner.html?commande=' + numCommande;
    const barcodeFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + encodeURIComponent(scanUrl) + '")';

    // Colonnes B à M (A = Statut n'est pas touchée)
    await writeRowAt(token, targetRow, [
      numCommande,                                    // B - ID Commande
      (prenom || '') + ' ' + (nom || ''),             // C - Nom Client
      telephone || '',                                // D - Numéro (téléphone)
      email || '',                                    // E - Email
      produitsFormate,                                // F - Produit(s)
      (quantite || '').split(' | ').join('\n'),       // G - Quantité
      adresse || '',                                  // H - Adresse
      livraison === 'Point relais' ? livraison : '',  // I - Point Relais
      '',                                             // J - Poids
      '',                                             // K - N° Suivi
      barcodeFormula,                                 // L - Code-barres
      comment || ''                                   // M - Notes
    ]);

    return { statusCode: 200, headers, body: JSON.stringify({ success: true, numCommande: numCommande, ligne: targetRow }) };

  } catch (err) {
    console.log('ERROR:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ success: false, error: err.message }) };
  }
};
