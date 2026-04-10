const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");

try {
  const content = fs.readFileSync("c:\\Users\\nikun\\OneDrive\\Documents\\Desktop\\GSE new\\templates\\bank-quotation.docx", "binary");
  const zip = new PizZip(content);
  // Remove a:solidFill things like server.js does
  let xml = zip.files['word/document.xml'].asText();
  xml = xml.replace(/<a:ln([^>]*)>/g, (match, attrs) => {
    let newAttrs = attrs;
    if (/w="\d+"/.test(newAttrs)) {
        newAttrs = newAttrs.replace(/w="\d+"/, 'w="0"');
    } else {
        newAttrs += ' w="0"';
    }
    return `<a:ln${newAttrs}>`;
  });
  xml = xml.replace(/<a:ln([^>]*)>([\s\S]*?)<\/a:ln>/g, (match, attrs, inner) => {
      const fixedInner = inner.replace(/<a:solidFill[\s\S]*?<\/a:solidFill>/g, '').replace(/<a:solidFill[^>]*\/>/g, '');
      return `<a:ln${attrs}>${fixedInner}</a:ln>`;
  });
  zip.file('word/document.xml', xml);

  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  console.log('Template parsed successfully!');
} catch (error) {
  const e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
  }
  console.log(JSON.stringify({error: e}));
  if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors.map(function (error) {
          return error.properties ? error.properties.explanation : error;
      }).join("\n");
      console.log('errorMessages', errorMessages);
      console.log(JSON.stringify(error.properties.errors, null, 2));
  }
}
