import * as JSZip from 'jszip';

export interface BonCommandeData {
  cs: string;
  cr: string;
  cc: string;
  article: string;
  exercice: string;
  fournisseurNom: string;
  fournisseurAdresse1: string;
  fournisseurAdresse2: string;
  codeFournisseur: string;
  lignes: { description: string; quantite: string; prixUnitaire: string; total: string }[];
  montantTotalLettres: string;
  delaiLivraison: string;
  lieu: string;
  date: string;
  numeroEngagement: string;
  operation: string;
  numeroSerie: string;
  numeroBon: string;
  codeIndividuel: string;
  compteLimitatif: string;
  operationEquipement: string;
  compteDe: string;
  montantAD: string;
  engagementsAnterieurs: string;
}

// ──────────────────────────────────────────────────────
//  Helpers
// ──────────────────────────────────────────────────────

/** Find all "box tables" (small grid cells of width 200-230) in order */
function findBoxTables(xml: string): { start: number; end: number }[] {
  const results: { start: number; end: number }[] = [];
  const gridRegex = /<w:tblGrid>((?:<w:gridCol w:w="\d+"\/>\s*)+)<\/w:tblGrid>/g;
  let match: RegExpExecArray | null;

  while ((match = gridRegex.exec(xml)) !== null) {
    const cols = match[1].match(/w:w="(\d+)"/g);
    if (!cols) continue;
    const widths = cols.map(c => parseInt(c.match(/\d+/)![0]));
    if (widths[0] >= 200 && widths[0] <= 230 && widths.length >= 3) {
      const tblStart = xml.lastIndexOf('<w:tbl>', match.index);
      let depth = 1;
      let i = tblStart + 7;
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

/** Fill a box-table's empty cells with individual characters */
function fillBoxTable(xml: string, boxTables: { start: number; end: number }[], tableIndex: number, value: string): string {
  if (tableIndex >= boxTables.length || !value) return xml;
  const { start, end } = boxTables[tableIndex];
  let tableXml = xml.substring(start, end);

  let charIdx = 0;
  tableXml = tableXml.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (cellMatch, cellContent: string) => {
    if (charIdx >= value.length) return cellMatch;
    const hasText = /<w:t[^>]*>[^<\s]+<\/w:t>/.test(cellContent);
    if (hasText) return cellMatch;

    const char = value[charIdx++];
    const newCell = cellContent.replace(
      /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/,
      (_m: string, ps: string, _pc: string, pe: string) =>
        `${ps}<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr><w:t>${char}</w:t></w:r>${pe}`
    );
    return `<w:tc>${newCell}</w:tc>`;
  });

  return xml.substring(0, start) + tableXml + xml.substring(end);
}

/** Replace exact text inside a <w:t> tag (tolerates trailing space) */
function replaceText(xml: string, searchText: string, replacement: string): string {
  const escaped = searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`(<w:t[^>]*>)\\s*${escaped}\\s*(</w:t>)`);
  if (regex.test(xml)) {
    return xml.replace(regex, `$1${replacement}$2`);
  }
  return xml;
}

/** Insert a value run right after the </w:r> that contains `label`, searching from `startPos` */
function insertValueAfterLabel(xml: string, label: string, value: string, startPos = 0, fontSize = '13'): string {
  if (!value) return xml;
  // Search for label text, tolerating trailing spaces (template may have "LABEL " instead of "LABEL")
  let idx = xml.indexOf(`>${label}<`, startPos);
  if (idx === -1) idx = xml.indexOf(`>${label} <`, startPos);
  if (idx === -1) idx = xml.indexOf(label, startPos);
  if (idx === -1) return xml;
  const closeR = xml.indexOf('</w:r>', idx);
  if (closeR === -1) return xml;
  const insertPos = closeR + 6;
  const run = `<w:r><w:rPr><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr><w:t xml:space="preserve"> ${value}</w:t></w:r>`;
  return xml.substring(0, insertPos) + run + xml.substring(insertPos);
}

/** Insert value after the colon <w:t> that follows a label. Colon pattern is "\u00A0:" or "\u00A0: ". */
function insertValueAfterColon(xml: string, label: string, value: string, startPos = 0, fontSize = '12'): string {
  if (!value) return xml;
  const labelIdx = xml.indexOf(`>${label}<`, startPos);
  if (labelIdx === -1) return xml;
  // Find the next <w:t> containing "\u00A0:" within 800 chars
  const searchZone = xml.substring(labelIdx, labelIdx + 800);
  const colonMatch = searchZone.match(/<w:t[^>]*>\u00A0:[^<]*<\/w:t>/);
  if (!colonMatch) return xml;
  const colonAbsIdx = labelIdx + colonMatch.index!;
  const closeR = xml.indexOf('</w:r>', colonAbsIdx);
  if (closeR === -1) return xml;
  const insertPos = closeR + 6;
  const run = `<w:r><w:rPr><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr><w:t xml:space="preserve"> ${value}</w:t></w:r>`;
  return xml.substring(0, insertPos) + run + xml.substring(insertPos);
}

// ──────────────────────────────────────────────────────
//  Main generator
// ──────────────────────────────────────────────────────

export async function fillBonCommandeWord(
  templateBuffer: ArrayBuffer,
  data: BonCommandeData
): Promise<Blob> {
  const zip = await JSZip.loadAsync(templateBuffer);
  let xml = await zip.file('word/document.xml')!.async('string');

  // ── Computed values ────────────────────────────────
  const montantTotal = data.lignes.reduce((s, l) => s + (parseFloat(l.total) || 0), 0);
  const montantTotalStr = montantTotal > 0
    ? montantTotal.toLocaleString('fr-FR').replace(/[\u202F\u00A0]/g, ' ')
    : '';
  const ant = parseFloat((data.engagementsAnterieurs || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const ad = parseFloat((data.montantAD || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const cumul = montantTotal + ant;
  const disponible = ad - cumul;
  const fmt = (n: number) => n > 0 ? n.toLocaleString('fr-FR').replace(/[\u202F\u00A0]/g, ' ') : '';

  // ── Index box tables once ──────────────────────────
  // Order in template:
  //  0  CS right       (3 cells)   "CENTRE DE SYNTHESE"
  //  1  CR right       (3 cells)   "CENTRE DE RESPONSABILITE"
  //  2  CC right       (3 cells)   "CENTRE DE COUT"
  //  3  Exercice       (4 cells)   "EXERCICE"
  //  4  Code fourniss. (4 cells)   "CODE FOURNISSEUR"
  //  5  CS left        (3 cells)   "C.S."
  //  6  CR left        (3 cells)   "C.R."
  //  7  CC left        (3 cells)   "C.C."
  //  8  Article left   (4 cells)   near "ARTICLE"
  //  9  (Exercice left or N° Eng)  (4 cells)
  // 10  Compte Limitatif left      (4 cells)
  // 11  Op. Equipement left        (4 cells)
  // 12  Compta. Générale left      (6 cells)
  // 13  Date left                  (6 cells) JJ MM AA
  let boxes = findBoxTables(xml);

  // ================================================================
  //  RIGHT SIDE — BON DE COMMANDE
  // ================================================================

  // Box fields
  xml = fillBoxTable(xml, boxes, 0, data.cs);
  boxes = findBoxTables(xml); // re-index after each modification
  xml = fillBoxTable(xml, boxes, 1, data.cr);
  boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 2, data.cc);
  boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 3, data.exercice);
  boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 4, data.codeFournisseur);
  boxes = findBoxTables(xml);

  // Fournisseur address: line 1 is embedded with the label, lines 2-4 are separate
  // Line 1: "ADRESSE DU FOURNISSEUR\u00A0: ________________________" (non-breaking space before :)
  const shortUnderscores = '________________________';
  const fournLabel = `ADRESSE DU FOURNISSEUR\u00A0: ${shortUnderscores}`;
  const fournIdx = xml.indexOf(fournLabel);
  if (fournIdx > -1) {
    const replacement = `ADRESSE DU FOURNISSEUR\u00A0: ${data.fournisseurNom || shortUnderscores}`;
    xml = xml.substring(0, fournIdx) + replacement + xml.substring(fournIdx + fournLabel.length);
  }
  // Lines 2-3-4: standalone underscore lines
  const addrLine = '________________________________________________';
  const addrIdx = xml.indexOf('ADRESSE DU FOURNISSEUR');
  if (addrIdx > -1) {
    let searchPos = addrIdx + 100;
    for (const val of [data.fournisseurAdresse1, data.fournisseurAdresse2]) {
      const lineIdx = xml.indexOf(addrLine, searchPos);
      if (lineIdx > -1 && val) {
        xml = xml.substring(0, lineIdx) + val + xml.substring(lineIdx + addrLine.length);
        searchPos = lineIdx + val.length;
      } else if (lineIdx > -1) {
        searchPos = lineIdx + addrLine.length;
      }
    }
  }

  // Article table rows (between "DETAIL DE LA COMMANDE" and "MONTANT TOTAL EN CHIFFRE")
  const detailIdx = xml.indexOf('DETAIL DE LA COMMANDE');
  const mtcIdx = xml.indexOf('MONTANT TOTAL EN CHIFFRE');
  if (detailIdx > -1 && mtcIdx > -1) {
    const zone = xml.substring(detailIdx, mtcIdx);
    // Find empty 4-cell rows (article data rows)
    const rowRegex = /<w:tr [^>]*>[\s\S]*?<\/w:tr>/g;
    let rm: RegExpExecArray | null;
    const emptyRows: { absStart: number; absEnd: number; xml: string }[] = [];

    while ((rm = rowRegex.exec(zone)) !== null) {
      const rowXml = rm[0];
      const cells = rowXml.match(/<w:tc>/g);
      const hasText = /<w:t[^>]*>[^\s<]+<\/w:t>/.test(rowXml);
      const trH = rowXml.match(/w:trHeight w:val="(\d+)"/);
      const height = trH ? parseInt(trH[1]) : 0;

      if (cells && cells.length === 4 && !hasText && height >= 300) {
        emptyRows.push({
          absStart: detailIdx + rm.index,
          absEnd: detailIdx + rm.index + rowXml.length,
          xml: rowXml,
        });
      }
    }

    // Fill rows backwards to preserve positions
    for (let i = Math.min(data.lignes.length, emptyRows.length) - 1; i >= 0; i--) {
      const row = emptyRows[i];
      const ligne = data.lignes[i];
      const vals = [ligne.description, ligne.quantite, ligne.prixUnitaire, ligne.total];
      let ci = 0;

      const filled = row.xml.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (cm: string, cc: string) => {
        const v = vals[ci++] || '';
        if (!v) return cm;
        const nc = cc.replace(
          /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/,
          (_m: string, ps: string, _pc: string, pe: string) =>
            `${ps}<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve">${v}</w:t></w:r>${pe}`
        );
        return `<w:tc>${nc}</w:tc>`;
      });

      xml = xml.substring(0, row.absStart) + filled + xml.substring(row.absEnd);
    }
  }

  // Montant total en chiffre: insert in last cell of the MONTANT row
  if (montantTotalStr && mtcIdx > -1) {
    const mtcInXml = xml.indexOf('>MONTANT TOTAL EN CHIFFRE<');
    if (mtcInXml > -1) {
      const rowStart = xml.lastIndexOf('<w:tr ', mtcInXml);
      const rowEnd = xml.indexOf('</w:tr>', mtcInXml) + 7;
      let rowXml = xml.substring(rowStart, rowEnd);
      let cellCount = 0;
      rowXml = rowXml.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (cm: string, cc: string) => {
        cellCount++;
        if (cellCount === 4) {
          const nc = cc.replace(
            /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/,
            (_m: string, ps: string, _pc: string, pe: string) =>
              `${ps}<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="13"/><w:szCs w:val="13"/></w:rPr><w:t xml:space="preserve">${montantTotalStr}</w:t></w:r>${pe}`
          );
          return `<w:tc>${nc}</w:tc>`;
        }
        return cm;
      });
      xml = xml.substring(0, rowStart) + rowXml + xml.substring(rowEnd);
    }
  }

  // EN LETTRE (NBSP before colon: \u00A0, 59 ellipsis chars + "...")
  if (data.montantTotalLettres) {
    const enLettreSearch = 'EN LETTRE\u00A0: ' + '\u2026'.repeat(59) + '...';
    const enLettreIdx = xml.indexOf(enLettreSearch);
    if (enLettreIdx > -1) {
      const repl = `EN LETTRE\u00A0: ${data.montantTotalLettres}`;
      xml = xml.substring(0, enLettreIdx) + repl + xml.substring(enLettreIdx + enLettreSearch.length);
    }
  }

  // DELAI DE LIVRAISON (NBSP before colon, 23 ellipsis chars)
  if (data.delaiLivraison) {
    const delaiSearch = 'DELAI DE LIVRAISON\u00A0: ' + '\u2026'.repeat(23);
    const delaiIdx = xml.indexOf(delaiSearch);
    if (delaiIdx > -1) {
      const repl = `DELAI DE LIVRAISON\u00A0: ${data.delaiLivraison}`;
      xml = xml.substring(0, delaiIdx) + repl + xml.substring(delaiIdx + delaiSearch.length);
    }
  }

  // COMPTE LIMITATIF, OPERATION, N° ENGAGEMENT on right side (after NOTA)
  // These labels have a ":" in a separate <w:t> — find and insert value after it
  // Process bottom-up to preserve positions
  const rightBottom = xml.indexOf('NOTA');
  if (rightBottom > -1) {
    // N° ENGAGEMENT (furthest down, do first to not break positions)
    xml = insertValueAfterColon(xml, "N\u00B0 D\u2019ENGAGEMENT", data.numeroEngagement, rightBottom);
    if (xml.indexOf(">N\u00B0 D\u2019ENGAGEMENT<") === -1) {
      xml = insertValueAfterColon(xml, "N° D'ENGAGEMENT", data.numeroEngagement, rightBottom);
    }
    // OPERATION
    xml = insertValueAfterColon(xml, 'OPERATION', data.operation, rightBottom);
    // COMPTE LIMITATIF
    xml = insertValueAfterColon(xml, 'COMPTE LIMITATIF', data.compteLimitatif, rightBottom);
  }

  // Lieu et date on right side: " A\u00A0: ___________ LE : __________"
  {
    const lieuSearch = ' A\u00A0: ___________ LE : __________';
    const lieuIdx = xml.indexOf(lieuSearch);
    if (lieuIdx > -1) {
      const repl = ` A\u00A0: ${data.lieu || '___________'} LE : ${data.date || '__________'}`;
      xml = xml.substring(0, lieuIdx) + repl + xml.substring(lieuIdx + lieuSearch.length);
    }
  }

  // ================================================================
  //  LEFT SIDE — BON D'ENGAGEMENT
  // ================================================================

  boxes = findBoxTables(xml);

  // CS, CR, CC (left side)
  xml = fillBoxTable(xml, boxes, 5, data.cs);
  boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 6, data.cr);
  boxes = findBoxTables(xml);
  xml = fillBoxTable(xml, boxes, 7, data.cc);
  boxes = findBoxTables(xml);

  // Article
  xml = fillBoxTable(xml, boxes, 8, data.article);
  boxes = findBoxTables(xml);

  // BON DE COMMANDE N° (left side)
  const leftAnchor = xml.indexOf('w:tblpX="380"');
  if (leftAnchor > -1 && (data.numeroBon || data.numeroSerie)) {
    xml = insertValueAfterLabel(xml, 'BON DE COMMANDE', `N° ${data.numeroBon || data.numeroSerie}`, leftAnchor, '12');
  }

  // Fournisseur on left side: replace dots (22 ellipsis chars in template)
  xml = replaceText(xml,
    '\u2026'.repeat(22),
    data.fournisseurNom || '\u2026'.repeat(22)
  );

  // Left side fields: insert value after each label
  // Process from bottom to top to preserve positions
  const leftStart = leftAnchor > -1 ? leftAnchor : Math.floor(xml.length / 2);

  const leftFields: [string, string][] = [
    ['DISPONIBLE', fmt(disponible)],
    ['CUMUL DES ENGAGEMENTS', fmt(cumul)],
    ['ANTERIEURS', data.engagementsAnterieurs],
    ['MONTANT DU BON', montantTotalStr],
    ['MONTANT A.D.', data.montantAD],
    ['COMPTABILITE GENERALE', data.compteDe],
    ['OPERATION D\u2019EQUIPEMENT', data.operationEquipement],
    ['COMPTE LIMITATIF', data.compteLimitatif],
    ['NUMERO D\u2019ENGAGEMENT', data.numeroEngagement],
    ['CODE INIVIDUEL', data.codeIndividuel],
  ];

  // Collect insertions (from bottom of document upward)
  const insertions: { pos: number; run: string }[] = [];
  for (const [label, value] of leftFields) {
    if (!value) continue;
    // Search only in left side (after leftStart) — use label text directly
    // to handle trailing spaces in template text nodes
    const idx = xml.indexOf(label, leftStart);
    if (idx === -1) continue;
    const closeR = xml.indexOf('</w:r>', idx);
    if (closeR === -1) continue;
    insertions.push({
      pos: closeR + 6,
      run: `<w:r><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t xml:space="preserve"> ${value}</w:t></w:r>`,
    });
  }

  // Sort by position descending, then apply
  insertions.sort((a, b) => b.pos - a.pos);
  for (const ins of insertions) {
    xml = xml.substring(0, ins.pos) + ins.run + xml.substring(ins.pos);
  }

  // Date on left side: fill box table 13 (6 cells: JJ MM AA)
  if (data.date) {
    const parts = data.date.split('/');
    if (parts.length === 3) {
      const dateStr = parts[0].padStart(2, '0') + parts[1].padStart(2, '0') + parts[2].slice(-2).padStart(2, '0');
      boxes = findBoxTables(xml);
      xml = fillBoxTable(xml, boxes, 13, dateStr);
    }
  }

  // Lieu on left side
  xml = replaceText(xml, 'A __________ LE', `A ${data.lieu || '__________'} LE`);

  // ── Save ───────────────────────────────────────────
  zip.file('word/document.xml', xml);
  const outBuf = await zip.generateAsync({ type: 'blob' });
  return new Blob([outBuf], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}
