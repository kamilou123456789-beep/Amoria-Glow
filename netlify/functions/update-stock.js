exports.handler = async function(event, context) {
  const SUPABASE_URL = 'https://cfqrdxiutsaxwupaiogn.supabase.co';
  const SUPABASE_KEY = 'sb_publishable_FErqHjDSEAz55WOD13Xf6A_CnlMIdAP';

  try {
    const body = JSON.parse(event.body);
    const items = body.items;

    for (const item of items) {
      const getRes = await fetch(`${SUPABASE_URL}/rest/v1/stock?ref=eq.${item.ref}&select=qty`, {
        headers: {
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${SUPABASE_KEY}`
        }
      });
      const data = await getRes.json();
      if (data && data[0]) {
        const newQty = Math.max(0, data[0].qty - item.qty);
        await fetch(`${SUPABASE_URL}/rest/v1/stock?ref=eq.${item.ref}`, {
          method: 'PATCH',
          headers: {
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${SUPABASE_KEY}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ qty: newQty })
        });
      }
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: true })
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
