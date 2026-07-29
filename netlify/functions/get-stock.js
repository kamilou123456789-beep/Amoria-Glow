exports.handler = async function(event, context) {
  const SUPABASE_URL = 'https://cfqrdxiutsaxwupaiogn.supabase.co';
  const SUPABASE_KEY = 'sb_publishable_FErqHjDSEAz55WOD13Xf6A_CnlMIdAP';

  try {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/stock?select=ref,qty`, {
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`
      }
    });
    const data = await response.json();
    const stock = {};
    data.forEach(item => { stock[item.ref] = item.qty; });
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(stock)
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
