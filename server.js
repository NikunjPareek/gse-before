const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { spawnSync } = require('child_process');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const COUNTER_FILE = path.join(__dirname, 'counter.json');
const TEMPLATES_DIR = path.join(__dirname, 'templates');
const OUTPUT_DIR = path.join(__dirname, 'output'); // Temporary dir for PDF conversion

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR);
}

// Ensure counter file exists
if (!fs.existsSync(COUNTER_FILE)) {
  fs.writeFileSync(COUNTER_FILE, JSON.stringify({ count: 1 }));
}

// Helpers
function getCount() {
  const data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf-8'));
  return data.count;
}

function generateDocxBlob(type, data) {
  const templatePath = path.join(TEMPLATES_DIR, type === 'bank' ? 'bank-quotation.docx' : 'client-quotation.docx');
  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);

  // Fix background image border in XML before rendering
  let xml = zip.files['word/document.xml'].asText();
  
  // Find <a:ln ... >
  xml = xml.replace(/<a:ln([^>]*)>/g, (match, attrs) => {
    // If it has a w element, replace it, otherwise add it
    let newAttrs = attrs;
    if (/w="\d+"/.test(newAttrs)) {
        newAttrs = newAttrs.replace(/w="\d+"/, 'w="0"');
    } else {
        newAttrs += ' w="0"';
    }
    return `<a:ln${newAttrs}>`;
  });

  // Remove <a:solidFill> inside <a:ln> nodes.
  // We can do this by just removing any <a:solidFill> that is immediately following <a:ln ...> or inside an a:ln block
  // A simple regex approach works for removing a:solidFill inside a:ln blocks
  xml = xml.replace(/<a:ln([^>]*)>([\s\S]*?)<\/a:ln>/g, (match, attrs, inner) => {
      const fixedInner = inner.replace(/<a:solidFill[\s\S]*?<\/a:solidFill>/g, '').replace(/<a:solidFill[^>]*\/>/g, '');
      return `<a:ln${attrs}>${fixedInner}</a:ln>`;
  });
  
  zip.file('word/document.xml', xml);

  const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
  });

  doc.render(data);
  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function getSafeFilename(clientName, ext) {
    const safeName = (clientName || '').replace(/[^a-zA-Z0-9 .\-]/g, '').trim().replace(/ /g, '_');
    if (!safeName) return `GSE_Quotation_Draft.${ext}`;
    return `${safeName}_GSE_Quotation.${ext}`;
}

// Endpoints
app.get('/api/next-count', (req, res) => {
  res.json({ count: getCount() });
});

app.post('/api/increment-count', (req, res) => {
  let count = getCount();
  count++;
  fs.writeFileSync(COUNTER_FILE, JSON.stringify({ count }));
  res.json({ count });
});

app.post('/api/download/docx', (req, res) => {
  try {
    const { type, data } = req.body;
    const buf = generateDocxBlob(type, data);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${getSafeFilename(data.client_name, 'docx')}"`);
    res.send(buf);
  } catch (error) {
    console.error('Error generating DOCX:', error);
    res.status(500).send({ error: error.message });
  }
});

app.post('/api/download/pdf', (req, res) => {
  try {
    const { type, data } = req.body;
    const docxBuf = generateDocxBlob(type, data);
    
    // Save temp docx
    const tempDocxPath = path.join(OUTPUT_DIR, `temp_${Date.now()}.docx`);
    fs.writeFileSync(tempDocxPath, docxBuf);

    const sofficePath = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
    
    // Convert to PDF
    const result = spawnSync(sofficePath, [
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', OUTPUT_DIR,
        tempDocxPath
    ]);
    
    if (result.status !== 0) {
      console.error('LibreOffice Error:', result.stderr ? result.stderr.toString() : 'Unknown error');
      return res.status(500).json({ error: 'Failed to generate PDF via LibreOffice.' });
    }

    const pdfPath = tempDocxPath.replace(/\.docx$/, '.pdf');
    if (!fs.existsSync(pdfPath)) {
        console.error('LibreOffice did not generate the PDF file.');
        return res.status(500).json({ error: 'PDF file not created.' });
    }

    const pdfBuf = fs.readFileSync(pdfPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${getSafeFilename(data.client_name, 'pdf')}"`);
    res.send(pdfBuf);
    
    // Cleanup
    try {
        fs.unlinkSync(tempDocxPath);
        fs.unlinkSync(pdfPath);
    } catch(e) {}
    
  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).send({ error: error.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
