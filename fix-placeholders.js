/**
 * fix-placeholders.js  v2
 *
 * Problem: In Word, when you type {{ref_no}} the XML splits it across
 * multiple <w:r><w:t>...</w:t></w:r> runs due to spell-check / autocorrect.
 * So docxtemplater never sees a complete {{token}} and leaves it blank.
 *
 * Fix strategy:
 * 1. Within each <w:p> paragraph, collect all <w:t> text fragments.
 * 2. Concatenate them. If the concatenated text contains a bare placeholder
 *    name (like "ref_no" but NOT already "{{ref_no}}"), wrap it.
 * 3. Rebuild the paragraph so the first <w:r> holds the entire corrected
 *    text and subsequent <w:r> runs (that were part of the split) are removed.
 *
 * This is safer than trying to stitch XML nodes together manually.
 */

const PizZip = require('pizzip');
const fs     = require('fs');
const path   = require('path');

// All placeholder names the templates should contain
const PLACEHOLDERS = [
  'ref_no', 'date', 'kw', 'client_name', 'client_address',
  'client_number', 'base_cost', 'gst', 'final_amount', 'final_amount_words',
  'vendor_name', 'vendor_phone', 'Discount_amount',
  ...Array.from({length:11}, (_,i) => `spec_${i+1}`),
  ...Array.from({length:11}, (_,i) => `company_${i+1}`),
  'qty_1', 'qty_2', 'qty_4',
];

// Build a set for fast lookup
const PLACEHOLDER_SET = new Set(PLACEHOLDERS);

/**
 * Given the raw XML of a Word document, merge split runs that contain
 * bare placeholder names and wrap them with {{ }}.
 *
 * Strategy:
 * - Iterate over every <w:p>...</w:p> paragraph.
 * - Collect the concatenated text from all <w:t> elements.
 * - If that text contains a bare placeholder name (not already {{...}}),
 *   rewrite the paragraph: keep the first <w:r> run structure but replace
 *   its <w:t> content with the fully-formed {{placeholder}} text, and
 *   delete the other <w:r> runs that made up the original split.
 */
function fixXml(xml) {
  let totalFixed = 0;

  // Process each paragraph independently
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (para) => {
    // Extract all (run, text) pairs: { runXml, text }
    // A "run" is <w:r...>...</w:r> that contains a <w:t>
    const runRegex = /(<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>)/g;
    const tRegex   = /<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/;

    const runs = [];
    let m;
    while ((m = runRegex.exec(para)) !== null) {
      const runXml = m[1];
      const tMatch = tRegex.exec(runXml);
      if (tMatch) {
        runs.push({ runXml, text: tMatch[1], index: m.index });
      }
    }

    if (runs.length === 0) return para;

    // Concatenate all run texts
    const fullText = runs.map(r => r.text).join('');

    // Check if fullText contains any bare placeholder that needs wrapping
    let newText = fullText;
    let changed = false;

    for (const ph of PLACEHOLDERS) {
      // Already wrapped? Skip.
      if (newText.includes(`{{${ph}}}`)) continue;

      // Is the bare name present?
      // Use word-boundary-like check: not preceded/followed by { or }
      const re = new RegExp(`(?<![{}\\w])${ph.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}(?![{}\\w])`, 'g');
      const replaced = newText.replace(re, `{{${ph}}}`);
      if (replaced !== newText) {
        newText = replaced;
        changed = true;
        totalFixed++;
      }
    }

    if (!changed) return para;

    // Rebuild: keep all runs but replace text only in the FIRST run that
    // carries text, and blank out the others so docxtemplater gets one
    // clean token per paragraph segment.
    //
    // Simpler approach: replace the first run's <w:t> with the full newText
    // and remove all subsequent text-bearing runs' <w:t> content.

    let textAssigned = false;
    let newPara = para;

    // We'll replace each run's <w:t> in order
    // Build a mapping of old runXml -> new runXml
    const replacements = runs.map((run, idx) => {
      if (!textAssigned) {
        textAssigned = true;
        // Replace first run's <w:t>...</w:t> with newText
        const newRunXml = run.runXml.replace(
          /<w:t(?:[^>]*)>[\s\S]*?<\/w:t>/,
          (tTag) => {
            // Preserve xml:space="preserve" if present
            const hasPreserve = /xml:space="preserve"/.test(tTag);
            const attr = hasPreserve ? ' xml:space="preserve"' : '';
            return `<w:t${attr}>${escapeXml(newText)}</w:t>`;
          }
        );
        return { old: run.runXml, new: newRunXml };
      } else {
        // Blank out subsequent runs' text (but keep run for formatting)
        const newRunXml = run.runXml.replace(
          /<w:t(?:[^>]*)>[\s\S]*?<\/w:t>/,
          '<w:t/>'
        );
        return { old: run.runXml, new: newRunXml };
      }
    });

    // Apply replacements (use indexOf to avoid regex issues with XML content)
    for (const rep of replacements) {
      newPara = newPara.replace(rep.old, rep.new);
    }

    return newPara;
  });

  return { xml, totalFixed };
}

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function fixDocx(filePath) {
  console.log(`\nProcessing: ${path.basename(filePath)}`);
  const content = fs.readFileSync(filePath, 'binary');
  const zip     = new PizZip(content);

  const xmlFiles = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/footer1.xml'];
  let grandTotal = 0;

  for (const xmlFile of xmlFiles) {
    if (!zip.files[xmlFile]) continue;
    const original = zip.files[xmlFile].asText();
    const { xml: fixed, totalFixed } = fixXml(original);
    if (totalFixed > 0) {
      zip.file(xmlFile, fixed);
      console.log(`  ${xmlFile}: ${totalFixed} placeholder(s) repaired`);
      grandTotal += totalFixed;
    } else {
      console.log(`  ${xmlFile}: no changes`);
    }
  }

  if (grandTotal > 0) {
    const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(filePath, buf);
    console.log(`  ✅ Saved (${grandTotal} total fixes)`);
  } else {
    console.log(`  ✅ Already clean – nothing saved`);
  }

  return grandTotal;
}

// ── Run ─────────────────────────────────────────────────────────────────────
const templatesDir = path.join(__dirname, 'templates');
fixDocx(path.join(templatesDir, 'bank-quotation.docx'));
fixDocx(path.join(templatesDir, 'client-quotation.docx'));

// ── Verify ───────────────────────────────────────────────────────────────────
console.log('\n── Verification ─────────────────────────────────────────────────');
for (const file of ['bank-quotation.docx', 'client-quotation.docx']) {
  const content = fs.readFileSync(path.join(templatesDir, file), 'binary');
  const zip = new PizZip(content);
  const xml = zip.files['word/document.xml'].asText();
  const found = [...xml.matchAll(/\{\{([^}]+)\}\}/g)]
    .map(m => m[1])
    .filter((v,i,a) => a.indexOf(v) === i)
    .sort();
  console.log(`\n${file} (${found.length} placeholders):`);
  found.forEach(p => console.log(`  {{${p}}}`));
}

// ── Smoke test: render with dummy data ───────────────────────────────────────
console.log('\n── Smoke Test ───────────────────────────────────────────────────');
const Docxtemplater = require('docxtemplater');
const dummyData = {
  ref_no: 'GSE/B/5kW/270426/TEST/0001',
  date: '27/04/2026', kw: '5',
  client_name: 'Test Client', client_address: 'Test Address',
  client_number: '9876543210',
  base_cost: '3,00,000', gst: '26,700', final_amount: '2,08,700',
  final_amount_words: 'Two Lakh Eight Thousand Seven Hundred Rupees Only',
  vendor_name: 'Suraj Singh Rajawat', vendor_phone: '+91 8769151510',
  Discount_amount: '0',
  ...Object.fromEntries(Array.from({length:11},(_,i)=>[ `spec_${i+1}`, `Spec ${i+1}` ])),
  ...Object.fromEntries(Array.from({length:11},(_,i)=>[ `company_${i+1}`, `Co ${i+1}` ])),
  qty_1:'10 No.', qty_2:'1 Nos.', qty_4:'1 set',
};

for (const file of ['bank-quotation.docx', 'client-quotation.docx']) {
  try {
    const content = fs.readFileSync(path.join(templatesDir, file), 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true, linebreaks: true,
      delimiters: { start: '{{', end: '}}' },
      nullGetter: () => '',
    });
    doc.render(dummyData);
    console.log(`  ✅ ${file}: renders without errors`);
  } catch(e) {
    const msgs = e.properties && Array.isArray(e.properties.errors)
      ? e.properties.errors.map(x => x.message).join('; ')
      : e.message;
    console.error(`  ❌ ${file}: ${msgs}`);
  }
}
