  exports.handler = async function(event, context) {
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby61xxibBlw_VvalSENF96qiCjoda9nOK4ZczqluMcOKsaJq-1EsGa_UIMSzWUZYeSn/exe';

  try {
    const body = event.body;
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: body
    });

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
