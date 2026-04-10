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
  const directRuns = Array.from(p.childNodes).filter(n => n.nodeName === "w:r");
  let hasImage = false;
  for (const run of directRuns) {
    if (run.getElementsByTagName("w:drawing").length > 0 || run.getElementsByTagName("w:pict").length > 0 || run.getElementsByTagName("w:object").length > 0) {
      hasImage = true;
    }
  }
  let fullText = "";
  for (const run of directRuns) {
    const texts = Array.from(run.getElementsByTagName("w:t"));
    for (const t of texts) {
      fullText += t.textContent;
    }
  }
  if (fullText.includes("{{") || fullText.includes("}}")) {
    console.log(`P[${i}] hasImage=${hasImage}: ${fullText}`);
  }
}
