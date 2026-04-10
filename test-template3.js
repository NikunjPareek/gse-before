const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");

try {
  const content = fs.readFileSync("templates/bank-quotation.docx", "binary");
  const zip = new PizZip(content);
  let xml = zip.files['word/document.xml'].asText();
  
  // Remove proofing and bookmark errors
  xml = xml.replace(/<w:proofErr[^>]*>/g, '');
  xml = xml.replace(/<w:bookmarkStart[^>]*>/g, '');
  xml = xml.replace(/<w:bookmarkEnd[^>]*>/g, '');
  
  zip.file('word/document.xml', xml);

  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  console.log('Template parsed successfully WITHOUT structural rewrites!');
} catch (error) {
  const e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
  }
  console.log(JSON.stringify({error: e}));
  if (error.properties && error.properties.errors instanceof Array) {
      console.log(JSON.stringify(error.properties.errors, null, 2));
  }
}
