const http = require('http');

const data = {
  "type": "bank",
  "data": {
    "ref_no": "GSE/B/5kW/0001",
    "date": "2026-04-10",
    "kw": "5",
    "client_name": "Test Client",
    "client_address": "123 Test Street",
    "base_cost": "300000",
    "gst": "26700",
    "final_amount": "230700",
    "final_amount_words": "Two Lakh Thirty Thousand Seven Hundred Rupees Only",
    "spec_1": "BIFICAL", "company_1": "Adani", "qty_1": "10 Nos.",
    "spec_2": "5kW", "company_2": "K-Solar", "qty_2": "01 Nos.",
    "spec_3": "GP Paip", "company_3": "Apollo", 
    "spec_4": "Net Meter", "company_4": "Genius/L&T", "qty_4": "1 set",
    "spec_5": "4 mm", "company_5": "Polycab",
    "spec_6": "As Requirement", "company_6": "Aluminium Armand",
    "spec_7": "1 inch Strip", "company_7": "G.I",
    "spec_8": "2 Meter GI", "company_8": "As Per Standard",
    "spec_9": "Requited Made", "company_9": "As Per Standard",
    "spec_10": "As Requirement", "company_10": "As Per Standard",
    "spec_11": "As per standard", "company_11": "As Per Standard"
  }
};

const options = {
  hostname: 'localhost',
  port: 3000,
  path: '/api/download/docx',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  }
};

console.log('Sending request to /api/download/docx...');

const req = http.request(options, (res) => {
  console.log(`STATUS: ${res.statusCode}`);
  res.on('data', () => {}); // consume response
  res.on('end', () => {
    console.log('Request completed perfectly. No errors.');
    process.exit(0);
  });
});

req.on('error', (e) => {
  console.error(`Problem with request: ${e.message}`);
  process.exit(1);
});

req.write(JSON.stringify(data));
req.end();
