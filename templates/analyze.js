const fs = require('fs');
const PizZip = require('pizzip');

function analyze(name, filepath) {
  console.log('\n=== ' + name + ' ===');
  const zip = new PizZip(fs.readFileSync(filepath, 'binary'));
  const xml = zip.files['word/document.xml'].asText();
  const plain = xml.replace(/<[^>]+>/g, '');
  const found = [...new Set([...plain.matchAll(/\{\{([^}]+)\}\}/g)].map(m => m[0]))];
  found.sort().forEach(p => console.log('  ' + p));
  if (xml.includes('<a:ln')) console.log('WARNING: image border tag found');
  else console.log('No image border tags.');
}

analyze('BANK', 'bank-quotation.docx');
analyze('CLIENT', 'client-quotation.docx');
