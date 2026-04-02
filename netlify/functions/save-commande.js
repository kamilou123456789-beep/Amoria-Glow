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
      res.on('end', () => { try { resolve(JSON.parse(data).access_token); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Récupère la dernière ligne de la colonne A pour trouver le dernier numéro AMO-XXX
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
          const values = json.values || [];
          let lastNum = 0;
          // Parcourt toutes les valeurs de la colonne A et cherche le plus grand numéro AMO-XXX
          for (let i = 0; i < values.length; i++) {
            const cell = (values[i][0] || '').toString();
            const match = cell.match(/^AMO-(\d+)$/);
            if (match) {
              const n = parseInt(match[1], 10);
              if (n > lastNum) lastNum = n;
            }
          }
          resolve(lastNum);
        } catch(e) { reject(e); }
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
        console.log('Sheets response:', data);
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
    const body = JSON.parse(event.body);
    console.log('Received:', JSON.stringify(body));

    const { prenom, nom, email, produits, quantite, adresse, livraison, comment } = body;

    const token = await getAccessToken();
    console.log('Token OK');

    // ── Générer le prochain numéro AMO ──────────────────────
    const lastNum = await getLastOrderNumber(token);
    const nextNum = lastNum + 1;
    // Format AMO-001, AMO-002 ... AMO-099 ... AMO-100 etc.
    const numCommande = 'AMO-' + String(nextNum).padStart(3, '0');
    console.log('Numéro commande:', numCommande);

    // ── Formule code-barres Code128 ─────────────────────────
    const barcodeFormula = '=IMAGE("https://barcode.tec-it.com/barcode.ashx?data=' + numCommande + '&code=Code128&translate-esc=on")';

    // ── Colonnes A à K ──────────────────────────────────────
    // A: ID Commande  B: Nom Client  C: Email  D: Produit  E: Quantité
    // F: Adresse      G: Livraison   H: Statut  I: N° Suivi  J: Code-barres  K: Notes
    await appendRow(token, [
      numCommande,                        // A - ID Commande (AMO-001, AMO-002...)
      (prenom || '') + ' ' + (nom || ''), // B - Nom Client
      email || '',                        // C - Email
      produits || '',                     // D - Produit(s)
      quantite || '',                     // E - Quantité
      adresse || '',                      // F - Adresse
      livraison || '',                    // G - Mode livraison
      'À préparer',                       // H - Statut (par défaut)
      '',                                 // I - N° Suivi (à remplir manuellement)
      barcodeFormula,                     // J - Code-barres généré automatiquement
      comment || ''                       // K - Notes
    ]);

    // Retourner le numCommande au front pour l'afficher dans la confirmation
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ success: true, numCommande: numCommande })
    };

  } catch (err) {
    console.log('ERROR:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
