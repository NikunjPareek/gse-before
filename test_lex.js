const PizZip = require("pizzip");
const Lexer = require("docxtemplater/js/lexer.js");
const fs = require("fs");

const content = fs.readFileSync("templates/bank-quotation.docx", "binary");
const zip = new PizZip(content);
const xml = zip.files['word/document.xml'].asText();

const lexer = new Lexer(xml);
const tokens = lexer.parse();
for (let i = 0; i < tokens.length; i++) {
  if (tokens[i].type === 'delimiter' || tokens[i].type === 'tag') {
    if (tokens[i].value && tokens[i].value.includes('}}')) {
       console.log("Found:", tokens[i]);
    }
  }
}
