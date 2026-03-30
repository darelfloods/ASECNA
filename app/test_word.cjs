const fs = require('fs');
const PizZip = require('./node_modules/pizzip');

function extractTextWithNewlines(filePath) {
    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    let xml = zip.file('word/document.xml').asText();

    // Remplacer la fin d'un paragraphe par un saut de ligne
    xml = xml.replace(/<\/w:p>/gi, '\n');

    // Supprimer toutes les autres balises XML
    const text = xml.replace(/<[^>]+>/g, '');

    // Nettoyer les espaces multiples (optionnel)
    return text.trim();
}

try {
    // Try to find a docx
    const files = fs.readdirSync('c:/Users/DELL/Downloads').filter(f => f.endsWith('.docx'));
    if (files.length > 0) {
        console.log('Testing with: ' + files[0]);
        const text = extractTextWithNewlines('c:/Users/DELL/Downloads/' + files[0]);
        console.log(text.substring(0, 500));
    } else {
        console.log('No docx found to test');
    }
} catch (e) { console.error(e); }
