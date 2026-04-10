const PizZip = require("pizzip");
const { DOMParser } = require("@xmldom/xmldom");
const fs = require("fs");

const content = fs.readFileSync("c:\\Users\\nikun\\OneDrive\\Documents\\Desktop\\GSE new\\templates\\bank-quotation.docx", "binary");
const zip = new PizZip(content);
const xmlContent = zip.files['word/document.xml'].asText();
const doc = new DOMParser().parseFromString(xmlContent, "text/xml");
const paragraphs = doc.getElementsByTagName("w:p");

for (let i = 0; i < paragraphs.length; i++) {
  const p = paragraphs[i];
  if (p.toString().includes('{{')) {
    console.log(`P[${i}]:`, p.toString());
  }
}
