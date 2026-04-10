const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

try {
    const content = fs.readFileSync('templates/bank-quotation.docx', 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    console.log("Success!");
} catch (error) {
    if (error.properties) {
        console.log("Error properties:", JSON.stringify(error.properties, null, 2));
    }
    console.log("Error message:", error.message);
}
