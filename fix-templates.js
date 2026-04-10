const fs = require('fs');
const PizZip = require('pizzip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const templates = [
  'templates/bank-quotation.docx',
  'templates/client-quotation.docx'
];

function fixTemplate(filePath) {
  console.log(`Fixing ${filePath}...`);
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return;
  }
  const content = fs.readFileSync(filePath, 'binary');
  const zip = new PizZip(content);
  let xmlContent = zip.file("word/document.xml").asText();
  
  const doc = new DOMParser().parseFromString(xmlContent, "text/xml");
  const paragraphs = doc.getElementsByTagName("w:p");
  let fixedCount = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    
    // Process direct children in chunks. A chunk is broken by an <w:r> that has an image, or non-r nodes?
    // Actually, we can safely group consecutive w:r and w:proofErr nodes.
    let childNodes = Array.from(p.childNodes);
    let currentChunk = [];
    let chunks = [];
    
    for (const child of childNodes) {
      let isTextNode = false;
      if (child.nodeName === "w:proofErr" || child.nodeName === "w:bookmarkStart" || child.nodeName === "w:bookmarkEnd") {
        isTextNode = true;
      }
      if (child.nodeName === "w:r") {
         const hasImg = child.getElementsByTagName("w:drawing").length > 0 || 
                        child.getElementsByTagName("w:pict").length > 0 || 
                        child.getElementsByTagName("w:object").length > 0 ||
                        child.getElementsByTagName("mc:AlternateContent").length > 0;
         if (!hasImg) {
           isTextNode = true;
         }
      }
      
      if (isTextNode) {
        currentChunk.push(child);
      } else {
        if (currentChunk.length > 0) {
          chunks.push(currentChunk);
          currentChunk = [];
        }
      }
    }
    if (currentChunk.length > 0) {
      chunks.push(currentChunk);
    }
    
    for (const chunk of chunks) {
      // Gather text inside this chunk
      let fullText = "";
      for (const node of chunk) {
        if (node.nodeName === "w:r") {
          const texts = Array.from(node.getElementsByTagName("w:t"));
          for (const t of texts) {
            fullText += t.textContent;
          }
        }
      }
      
      if (fullText.includes("{{") || fullText.includes("}}")) {
        fixedCount++;
        // Find first formatting
        let firstRPr = null;
        for (const node of chunk) {
          if (node.nodeName === "w:r") {
            const rPrs = node.getElementsByTagName("w:rPr");
            if (rPrs.length > 0) {
              firstRPr = rPrs[0].cloneNode(true);
              break;
            }
          }
        }
        
        // Ensure the text runs match what Docxtemplater expects and does not have {{ duplicated or missed
        // Oh wait, if we merge them, the string might legit be correctly fixed.
        
        const newRun = doc.createElement("w:r");
        if (firstRPr) newRun.appendChild(firstRPr);
        const newT = doc.createElement("w:t");
        newT.setAttribute("xml:space", "preserve");
        newT.appendChild(doc.createTextNode(fullText));
        newRun.appendChild(newT);
        
        // Insert newRun before the first element of chunk
        p.insertBefore(newRun, chunk[0]);
        // Remove all old elements in the chunk
        for (const node of chunk) {
          p.removeChild(node);
        }
      }
    }
  }
  
  console.log(`Fixed ${fixedCount} chunks inside paragraphs.`);
  
  // Also apply the a:solidFill fix from server.js to ensure docxtemplater parses properly for pdfs, 
  // although docxtemplater parses without it. It's safe to do.
  let newXml = new XMLSerializer().serializeToString(doc);
  zip.file("word/document.xml", newXml);
  
  const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(filePath, buffer);
}

templates.forEach(fixTemplate);
