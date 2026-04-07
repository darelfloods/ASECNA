const fs = require('fs');
const JSZip = require('jszip');

function findBoxTables(xml) {
  const results = [];
  const gridRegex = /<w:tblGrid>((?:<w:gridCol w:w="\d+"\/>\s*)+)<\/w:tblGrid>/g;
  let match;
  while ((match = gridRegex.exec(xml)) !== null) {
    const cols = match[1].match(/w:w="(\d+)"/g);
    if (!cols) continue;
    const widths = cols.map(c => parseInt(c.match(/\d+/)[0]));
    if (widths[0] >= 200 && widths[0] <= 230 && widths.length >= 3) {
      const tblStart = xml.lastIndexOf('<w:tbl>', match.index);
      let depth = 1, i = tblStart + 7;
      while (depth > 0 && i < xml.length) {
        if (xml.substring(i, i + 7) === '<w:tbl>') depth++;
        if (xml.substring(i, i + 8) === '</w:tbl>') depth--;
        if (depth === 0) break;
        i++;
      }
      results.push({ start: tblStart, end: i + 8 });
    }
  }
  return results;
}

function fillBoxTable(xml, boxes, idx, value) {
  if (idx >= boxes.length || !value) return xml;
  const { start, end } = boxes[idx];
  let tXml = xml.substring(start, end);
  let ci = 0;
  tXml = tXml.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (cm, cc) => {
    if (ci >= value.length) return cm;
    const has = /<w:t[^>]*>[^<\s]+<\/w:t>/.test(cc);
    if (has) return cm;
    const ch = value[ci++];
    const nc = cc.replace(/(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/, (m, ps, pc, pe) =>
      ps + '<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr><w:t>' + ch + '</w:t></w:r>' + pe);
    return '<w:tc>' + nc + '</w:tc>';
  });
  return xml.substring(0, start) + tXml + xml.substring(end);
}

async function fullTest() {
  const buf = fs.readFileSync('public/BON DE COMMANDE A4.docx');
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  const r = [];

  // RIGHT boxes
  let boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 0, '120'); boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 1, '456'); boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 2, '789'); boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 3, '2026'); boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 4, '4567'); boxes = findBoxTables(xml);
  r.push('RIGHT boxes OK');

  // Fournisseur
  const fLabel = 'ADRESSE DU FOURNISSEUR\u00A0: ________________________';
  let fi = xml.indexOf(fLabel);
  if (fi > -1) {
    xml = xml.substring(0, fi) + 'ADRESSE DU FOURNISSEUR\u00A0: SOCIETE ABC SARL' + xml.substring(fi + fLabel.length);
    r.push('Fournisseur: OK');
  } else r.push('Fournisseur: FAIL');

  // Addr lines
  const aLine = '________________________________________________';
  let si = xml.indexOf('ADRESSE DU FOURNISSEUR') + 100;
  fi = xml.indexOf(aLine, si);
  if (fi > -1) { xml = xml.substring(0, fi) + 'BP 1234 Libreville' + xml.substring(fi + aLine.length); si = fi + 20; }
  fi = xml.indexOf(aLine, si);
  if (fi > -1) { xml = xml.substring(0, fi) + 'Gabon' + xml.substring(fi + aLine.length); }
  r.push('Address OK');

  // Articles
  const dIdx = xml.indexOf('DETAIL DE LA COMMANDE');
  const mIdx = xml.indexOf('MONTANT TOTAL EN CHIFFRE');
  const zone = xml.substring(dIdx, mIdx);
  const emptyRows = [];
  const rowRe = /<w:tr [^>]*>[\s\S]*?<\/w:tr>/g;
  let rm;
  while ((rm = rowRe.exec(zone)) !== null) {
    const cells = rm[0].match(/<w:tc>/g);
    const hasText = /<w:t[^>]*>[^\s<]+<\/w:t>/.test(rm[0]);
    const trH = rm[0].match(/w:trHeight w:val="(\d+)"/);
    const h = trH ? parseInt(trH[1]) : 0;
    if (cells && cells.length === 4 && !hasText && h >= 300)
      emptyRows.push({ s: dIdx + rm.index, e: dIdx + rm.index + rm[0].length, x: rm[0] });
  }
  const lignes = [
    { d: 'Fournitures bureau', q: '10', p: '5 000', t: '50 000' },
    { d: 'Cartouches imprimante', q: '5', p: '15 000', t: '75 000' },
  ];
  for (let i = Math.min(lignes.length, emptyRows.length) - 1; i >= 0; i--) {
    const row = emptyRows[i], l = lignes[i], vals = [l.d, l.q, l.p, l.t];
    let ci = 0;
    const filled = row.x.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (cm, cc) => {
      const v = vals[ci++] || '';
      if (!v) return cm;
      const nc = cc.replace(/(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/, (m, ps, pc, pe) =>
        ps + '<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve">' + v + '</w:t></w:r>' + pe);
      return '<w:tc>' + nc + '</w:tc>';
    });
    xml = xml.substring(0, row.s) + filled + xml.substring(row.e);
  }
  r.push('Articles: ' + emptyRows.length + ' rows, ' + lignes.length + ' filled');

  // EN LETTRE
  const elS = 'EN LETTRE\u00A0: ' + '\u2026'.repeat(59) + '...';
  fi = xml.indexOf(elS);
  if (fi > -1) {
    xml = xml.substring(0, fi) + 'EN LETTRE\u00A0: Cent vingt-cinq mille francs CFA' + xml.substring(fi + elS.length);
    r.push('EN LETTRE: OK');
  } else r.push('EN LETTRE: FAIL');

  // DELAI
  const dlS = 'DELAI DE LIVRAISON\u00A0: ' + '\u2026'.repeat(23);
  fi = xml.indexOf(dlS);
  if (fi > -1) {
    xml = xml.substring(0, fi) + 'DELAI DE LIVRAISON\u00A0: 30 jours' + xml.substring(fi + dlS.length);
    r.push('DELAI: OK');
  } else r.push('DELAI: FAIL');

  // Right bottom fields (after NOTA)
  const nota = xml.indexOf('NOTA');

  const insertAfterColon = (label, value) => {
    const lIdx = xml.indexOf('>' + label + '<', nota);
    if (lIdx === -1) { r.push(label + ' right: label not found'); return; }
    const sz = xml.substring(lIdx, lIdx + 800);
    const cm = sz.match(/<w:t[^>]*>\u00A0:[^<]*<\/w:t>/);
    if (!cm) { r.push(label + ' right: no colon'); return; }
    const cAbs = lIdx + cm.index;
    const cR = xml.indexOf('</w:r>', cAbs) + 6;
    xml = xml.substring(0, cR) + '<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve"> ' + value + '</w:t></w:r>' + xml.substring(cR);
    r.push(label + ' right: OK');
  };

  insertAfterColon('COMPTE LIMITATIF', '621');
  insertAfterColon('OPERATION', 'Achat fournitures');
  insertAfterColon("N\u00B0 D'ENGAGEMENT", 'ENG-2026-001');

  // Lieu/date right
  const ldS = ' A\u00A0: ___________ LE : __________';
  fi = xml.indexOf(ldS);
  if (fi > -1) {
    xml = xml.substring(0, fi) + ' A\u00A0: Libreville LE : 03/04/2026' + xml.substring(fi + ldS.length);
    r.push('Lieu/date right: OK');
  } else r.push('Lieu/date right: FAIL');

  // LEFT boxes
  xml = fillBoxTable(xml, findBoxTables(xml), 5, '120');
  xml = fillBoxTable(xml, findBoxTables(xml), 6, '456');
  xml = fillBoxTable(xml, findBoxTables(xml), 7, '789');
  xml = fillBoxTable(xml, findBoxTables(xml), 8, '0042');
  r.push('LEFT boxes OK');

  // BON DE COMMANDE N on left
  const la = xml.indexOf('w:tblpX="380"');
  const bcIdx = xml.indexOf('>BON DE COMMANDE<', la);
  if (bcIdx > -1) {
    const cR = xml.indexOf('</w:r>', bcIdx) + 6;
    xml = xml.substring(0, cR) + '<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve"> N\u00B0 BC-2026-042</w:t></w:r>' + xml.substring(cR);
    r.push('BON N left: OK');
  }

  // Fournisseur left (dots)
  const fDots = '\u2026'.repeat(25);
  fi = xml.indexOf(fDots);
  if (fi > -1) {
    xml = xml.substring(0, fi) + 'SOCIETE ABC SARL' + xml.substring(fi + fDots.length);
    r.push('Fournisseur left: OK');
  } else r.push('Fournisseur left: FAIL');

  // Left amount fields
  const leftFields = [
    ['DISPONIBLE', '275 000'],
    ['CUMUL DES ENGAGEMENTS', '225 000'],
    ['ANTERIEURS', '100 000'],
    ['MONTANT DU BON', '125 000'],
    ['MONTANT A.D.', '500 000'],
    ['COMPTABILITE GENERALE', '4411'],
    ["OPERATION D'EQUIPEMENT", 'Materiel info'],
    ['COMPTE LIMITATIF', '621'],
    ["NUMERO D'ENGAGEMENT", 'ENG-2026-001'],
    ['CODE INIVIDUEL', '0042'],
  ];
  const ins = [];
  for (const [label, value] of leftFields) {
    const idx = xml.indexOf('>' + label + '<', la);
    if (idx === -1) continue;
    const cR = xml.indexOf('</w:r>', idx);
    if (cR === -1) continue;
    ins.push({ pos: cR + 6, value });
  }
  ins.sort((a, b) => b.pos - a.pos);
  for (const i of ins) {
    xml = xml.substring(0, i.pos) + '<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve"> ' + i.value + '</w:t></w:r>' + xml.substring(i.pos);
  }
  r.push('Left fields: ' + ins.length + ' inserted');

  // Date left (boxes 13)
  xml = fillBoxTable(xml, findBoxTables(xml), 13, '030426');
  r.push('Date boxes OK');

  // Lieu left
  fi = xml.indexOf('A __________ LE');
  if (fi > -1) {
    xml = xml.substring(0, fi) + 'A Libreville LE' + xml.substring(fi + 15);
    r.push('Lieu left: OK');
  } else r.push('Lieu left: FAIL');

  // SAVE
  zip.file('word/document.xml', xml);
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync('C:/Users/DELL/Downloads/BON_COMMANDE_TEST.docx', out);
  console.log('=== RESULTS ===');
  r.forEach(x => console.log(x));
  console.log('\nSaved to Downloads/BON_COMMANDE_TEST.docx');
}

fullTest().catch(console.error);
