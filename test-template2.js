const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");

try {
  const content = fs.readFileSync("c:\\Users\\nikun\\OneDrive\\Documents\\Desktop\\GSE new\\templates\\bank-quotation.docx", "binary");
  const zip = new PizZip(content);
  // Remove regex
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  console.log('Template parsed successfully WITHOUT REGEX!');
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
