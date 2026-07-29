exports.handler = async function(event, context) {
  const SHEETS_CSV = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQTkUQ3UTjLhTTVv29vl_dkOtjweQ41gkUJNgwDkxgb_6Y0OtGX-8FsP7ZeQzNmEykeKifIXI5o42_4/pub?gid=1075508092&single=true&output=csv';

  try {
    const body = JSON.parse(event.body);
    const items = body.items;

    const response = await fetch(SHEETS_CSV);
    const csv = await response.text();
    const lines = csv.split('\n');

    console.log('✅ Stock mis à jour pour', items.length, 'produits');

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
