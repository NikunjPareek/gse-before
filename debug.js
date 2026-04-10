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
  const runs = Array.from(p.getElementsByTagName("w:r"));
  let fullText = "";
  for (const run of runs) {
    const texts = Array.from(run.getElementsByTagName("w:t"));
    for (const t of texts) {
      fullText += t.textContent;
    }
  }
  if (fullText.includes("{{") || fullText.includes("}}")) {
    console.log(`Paragraph ${i} text: ${JSON.stringify(fullText)}`);
  }
}
