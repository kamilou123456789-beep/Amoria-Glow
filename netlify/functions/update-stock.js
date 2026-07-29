const { google } = require('googleapis');

exports.handler = async function(event, context) {
  const SPREADSHEET_ID = '1iiP5phKdHb1DGnfF9w5g6tEXtwSppX2QN6qtG8cEH7E';
  
  try {
    const body = JSON.parse(event.body);
    const items = body.items;

    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: 'amoria-stock@amoria-glow.iam.gserviceaccount.com',
        private_key: '-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCskSkQEwZ/ZG6y\nveYAD7oY64Gfy6qyF0RxOzpnQRN5EilKeQWlLkiIfANQw5Wz9xtNQf340MqX1vpY\nv4/H6aMUC03LPxgeScCsXv63MEZMFHWmWQoiXhVpQoRLlOmzdvzKuOCdG49QBOr+\nf3FWZNLjr8MYWMa/I5d1RVOM9QSHF8rY5MSsMASdKexnVFx7zqf6HmcaNXLbdglB\nKdqgAVbRNDQvrMfnlFgGhPe3TBSOSHekDWlx1z3j1JGCxb9XxF+GnWG2GgI60+lT\nCxA1CFH8+7NOQKOwwWbnhppCWq19JpVGg3Alu35WYa+Z3PKfjDKF1Jevlwnn1t4g\nloicXwLbAgMBAAECggEAMf9s0kdw3oAOwqLafLIRzR6O0+mCb07meZgbd8cXCUEF\nzZn61LzwLvsfSssgGKBDvMKd/vUffZa/ue7mjZlXsnsD8xs4ta3QsSBk1FacR3a2\nD5hEo2h286ReCDgA7gpPe7zM9zgA8cI7A7mQ8OMNZwKJmAhArSh2vXd0maZzxV/Z\nKrT6QHIo5JYm11mlv9DqzIztU01xeV1gecP4Yehjh7aL3p16nH3WM6tzUtviXRH+\nXoloAM+1bFsoOMUTtm2siZCTO0Nn1kKi/h/bTTC0nX43iLmo0NP+8+V4EzoygfZ3\nw1wYyMzd8R0GaggMmnwR/f270/GmcLFG2lmjcyI/wQKBgQDjqpMe/5VgaSWDisgD\nMSha4HaQ98tqD9kfTX45E35BZwCZ/IjS+RLb/FyNntUreAOExsJ+dx/lFvOtrfV6\n+xUxVdSLJwKYzMwFco1a0VdRvwVjztSCrBuC8Zp38QLCCtTdKcDB4cNCgsBEsTRZ\nlm/o/kdGxoH3YH4Mn4Iw968uowKBgQDCCyDh8otOvjIYU1jHLcMCjQNUDhmCGOiN\nGRdMQxrEFr7n8w4LjiEGBe3cgFePts9/08ASJ9vFyh7K3P5DBT88+jGOnr/KJ9YL\nBNrFOvmeYrQSdvCA3rmhoTHPeUYnW/rDQD9ctovMwrX4id9amxYv1OO9SOTT8JDb\nRTY9Qv+2aQKBgFBmJJ6F09LATyctE4VNDttI+ZYobAWAo0SSsUimwaeHIIdAz3Dx\n1N8rN+Qre0xmjZeOOZE/sFvOxy9Gh7JuiQVrMiwSErCzYjlqQtEXrKaJtvWQTSv4\na57Kg6pnynmMKbAQ1qmheLs8QXoAumQI5Gx7n+A2qh8aTGlYyzlPvuXRAoGAaav5\nylJ1vvog+dJZ5I5tRrRYfav4BDtgWYayg1t/9g2VBWf93BkYrtkHwi86gA9ETQ6Z\n6MlADCSYRE25QfJXj/OIjWyycXrkO7f3E7WcPr7t5ahULTodyYGpSJ14sPKMS0xv\ntSPMWkQnKSScOBGBMac0Jt7NjwXRPTgh45ba/xECgYBSeFUr0JFSmLVFjKDEcuZh\nxMHDqcWu9e2Atpl2tfBlgDRlHLWMZibctGrare/QZNGXi37jW/4mP15BgIFi16UD\n1zCuraJbERzMMOwcmnx2ilHa3KyU9nRycDvuy6bb9v3g5gbKlFIygQYW0bPT9mhI\nVZ4eaG7E2RSAHxmlHn3/Iw==\n-----END PRIVATE KEY-----\n'
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });

    const sheets = google.sheets({ version: 'v4', auth });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Stock!A:D'
    });
    
    const rows = response.data.values;
    
    for (const item of items) {
      for (let i = 3; i < rows.length; i++) {
        if (rows[i][0] === item.ref) {
          const currentQty = parseInt(rows[i][3]) || 0;
          const newQty = Math.max(0, currentQty - item.qty);
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Stock!D${i + 1}`,
            valueInputOption: 'RAW',
            requestBody: { values: [[newQty]] }
          });
          break;
        }
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
