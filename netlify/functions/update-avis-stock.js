 exports.handler = async function(event, context) {
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby61xxibBlw_VvalSENF96qiCjoda9nOK4ZczqluMcOKsaJq-1EsGa_UIMSzWUZYeSn/exec';

  try {
    const body = JSON.parse(event.body);
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const data = await response.json();
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
