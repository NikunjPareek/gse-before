const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");

try {
  const content = fs.readFileSync("templates/bank-quotation.docx", "binary");
  const zip = new PizZip(content);
  const doc = new Docxtemplater();
  doc.loadZip(zip);
  doc.compile();
  console.log("Compiled successfully!");
  console.log(doc.getZip().files['word/document.xml'].asText());
} catch (error) {
  if (error.properties && error.properties.errors instanceof Array) {
      console.log(error.properties.errors);
  }
}
