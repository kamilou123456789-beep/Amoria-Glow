exports.handler = async function(event, context) {
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxYzOT-40ecSSshLRpPe2pqVqWVtKSNXAA29cisGUXCx0_EGrDEt1YxOV7zk2yCO0pH/exec';

  try {
    const response = await fetch(SCRIPT_URL);
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
};S
