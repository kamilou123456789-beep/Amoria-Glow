exports.handler = async function(event, context) {
  const SHEETS_CSV = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQTkUQ3UTjLhTTVv29vl_dkOtjweQ41gkUJNgwDkxgb_6Y0OtGX-8FsP7ZeQzNmEykeKifIXI5o42_4/pub?gid=1075508092&single=true&output=csv';

  try {
    const response = await fetch(SHEETS_CSV);
    const csv = await response.text();
    const lines = csv.split('\n');
    const stock = {};
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(',');
      if (cols.length >= 4) {
        const ref = cols[0].trim().replace(/^"|"$/g, '');
        const qty = parseInt(cols[3].trim().replace(/^"|"$/g, ''), 10);
        if (ref && !isNaN(qty)) stock[ref] = qty;
      }
    }
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
