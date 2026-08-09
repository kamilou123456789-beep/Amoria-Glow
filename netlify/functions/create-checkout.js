// Crée une session de paiement Stripe au montant du panier
// Utilise la clé secrète stockée dans Netlify : STRIPE_SECRET_KEY
// Paiement par carte + Apple Pay + Google Pay (automatique via Stripe Checkout)

const https = require('https');
const querystring = require('querystring');

const STRIPE_SECRET_KEY = process.env.STRIPE_SECRET_KEY;
const SITE_URL = 'https://amoria-glow-shop.netlify.app';

// Petit appel à l'API Stripe (form-urlencoded)
function stripeRequest(path, params) {
  const body = querystring.stringify(params);
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.stripe.com',
      path: path,
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + STRIPE_SECRET_KEY,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (res.statusCode >= 400) {
            return reject(new Error('STRIPE_' + res.statusCode + ': ' + (json.error && json.error.message ? json.error.message : data)));
          }
          resolve(json);
        } catch(e) { reject(new Error('STRIPE_PARSE: ' + data)); }
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
    if (!STRIPE_SECRET_KEY) throw new Error('ENV_MISSING: STRIPE_SECRET_KEY');

    const body = JSON.parse(event.body);
    // montant total en euros (ex: 21.52) et un libellé
    const montant = parseFloat(body.montant);
    const libelle = (body.libelle || 'Commande Amoria Glow').substring(0, 120);

    if (!montant || montant <= 0) throw new Error('MONTANT_INVALIDE');

    // Stripe travaille en centimes -> on arrondit
    const montantCents = Math.round(montant * 100);

    // Création de la session Checkout
    const params = {
      'mode': 'payment',
      'success_url': SITE_URL + '/merci.html?paid=1',
      'cancel_url': SITE_URL + '/index.html',
      'line_items[0][price_data][currency]': 'eur',
      'line_items[0][price_data][product_data][name]': libelle,
      'line_items[0][price_data][unit_amount]': String(montantCents),
      'line_items[0][quantity]': '1'
    };

    const session = await stripeRequest('/v1/checkout/sessions', params);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ url: session.url })
    };

  } catch (err) {
    console.log('ERROR:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
