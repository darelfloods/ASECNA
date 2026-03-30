const PizZip = require('pizzip');
const fs = require('fs');

const templateBuffer = fs.readFileSync('public/FICHE DE MISSION.docx');
const zip = new PizZip(templateBuffer);
let documentXml = zip.file('word/document.xml').asText();

const data = {
  nom: 'DUPONT',
  prenom: 'JEAN',
  matricule: '123456',
  emploi: 'Ingénieur Réseau',
  residence: 'Franceville',
  destination: 'Port-Gentil',
  motif: 'Maintenance antenne',
  dateDepart: '01/04/2026',
  dateRetour: '05/04/2026',
  duree: '5',
  transport: 'Véhicule'
};

function replaceInXml(xml, oldVal, newVal, desc) {
  if (!oldVal) return xml;
  const chars = oldVal.trim().split('');
  let src = '';
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (/^\s$/.test(c)) {
      src += '(?:\s|<[^>]+>)*';
    } else {
      src += c.replace(/[-\/\^$*+?.()|[\]{}]/g, '\$&');
      if (i < chars.length - 1) src += '(?:<[^>]+>)*';
    }
  }
  const regex = new RegExp(src, 'g');
  const orig = xml;
  const safeNew = newVal.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  xml = xml.replace(regex, function(match) {
    // Keep the text content only, strip XML tags from match
    return safeNew;
  });
  console.log(desc + ': ' + (xml !== orig ? 'OK' : 'ECHEC'));
  return xml;
}

function insertValueAfterLabel(xml, label, value, desc) {
  if (!value) return xml;
  const safeValue = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const chars = label.split('');
  let src = '';
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (/^\s$/.test(c)) {
      src += '(?:\s|<[^>]+>)*';
    } else {
      src += c.replace(/[-\/\^$*+?.()|[\]{}]/g, '\$&');
      if (i < chars.length - 1) src += '(?:<[^>]+>)*';
    }
  }
  const pattern = new RegExp('(' + src + '(?:<[^>]+>)*\s*:(?:<[^>]+>)*\s*)(<\/w:t>)', 'gi');
  let changed = false;
  xml = xml.replace(pattern, function(match, before, closeTag) {
    changed = true;
    return before + '  ' + safeValue + closeTag;
  });
  console.log(desc + ' (insertion): ' + (changed ? 'OK' : 'ECHEC'));
  return xml;
}

// Apply replacements
documentXml = replaceInXml(documentXml, 'NDONG', data.nom, 'Nom');
documentXml = replaceInXml(documentXml, '250356', data.matricule, 'Matricule');
documentXml = replaceInXml(documentXml, 'ROBERT', data.prenom, 'Prénom');
if (data.emploi) {
  documentXml = insertValueAfterLabel(documentXml, 'Emploi', data.emploi, 'Emploi');
}
documentXml = replaceInXml(documentXml, ': Libreville', ': ' + data.residence, 'Résidence');
documentXml = replaceInXml(documentXml, 'Port-Gentil', data.destination, 'Destination');
const motifOld = 'Contrôle des occupation domaniale';
documentXml = replaceInXml(documentXml, motifOld, data.motif, 'Motif');
documentXml = replaceInXml(documentXml, '26/03/2026', data.dateDepart, 'Date départ');
documentXml = replaceInXml(documentXml, '30/03/2026', data.dateRetour, 'Date retour');
documentXml = replaceInXml(documentXml, '5 Jours', data.duree + ' Jours', 'Durée');
documentXml = replaceInXml(documentXml, 'Avion', data.transport, 'Transport');

zip.file('word/document.xml', documentXml);
const output = zip.generate({ type: 'nodebuffer' });
fs.writeFileSync('test_output_fiche.docx', output);

// Verify
const vzip = new PizZip(output);
const vxml = vzip.file('word/document.xml').asText();
const paraRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
let m;
console.log('\n=== CONTENU DU DOCUMENT GENERE ===');
while ((m = paraRegex.exec(vxml)) !== null) {
  let text = '';
  const tRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let tm;
  while ((tm = tRegex.exec(m[0])) !== null) text += tm[1];
  if (text.trim()) console.log(text.trim());
}
