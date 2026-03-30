/**
 * Génère le template Excel "BON DE COMMANDE A4.xlsx" en A4 paysage
 * avec logo ASECNA, bordures, mise en page fidèle au document original.
 *
 * Usage: node create_bon_template_excel.cjs
 */
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function createBonCommandeTemplate() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ASECNA';
  wb.created = new Date();

  // ═══════════════════════════════════════════════════════════════════════════
  // PAGE 1 : BON DE COMMANDE
  // ═══════════════════════════════════════════════════════════════════════════
  const ws = wb.addWorksheet('BON DE COMMANDE', {
    pageSetup: {
      paperSize: 9, // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: {
        left: 0.4, right: 0.4,
        top: 0.3, bottom: 0.3,
        header: 0.2, footer: 0.2
      }
    },
    properties: {
      defaultRowHeight: 15,
      showGridLines: false
    }
  });

  // ── Couleurs & styles ──────────────────────────────────────────────────────
  const DARK_BLUE = '1F3864';
  const BLACK = '000000';
  const LIGHT_GRAY = 'E8E8E8';
  const WHITE = 'FFFFFF';
  const RED = 'CC0000';

  const thinBorder = { style: 'thin', color: { argb: BLACK } };
  const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

  const headerFont = { name: 'Arial', size: 8, bold: true, color: { argb: BLACK } };
  const labelFont = { name: 'Arial', size: 7.5, bold: false, color: { argb: BLACK } };
  const labelFontBold = { name: 'Arial', size: 7.5, bold: true, color: { argb: BLACK } };
  const valueFont = { name: 'Arial', size: 9, bold: true, color: { argb: DARK_BLUE } };
  const valueFontSmall = { name: 'Arial', size: 8, bold: true, color: { argb: DARK_BLUE } };
  const titleFont = { name: 'Arial', size: 14, bold: true, color: { argb: BLACK } };
  const smallFont = { name: 'Arial', size: 6.5, color: { argb: '444444' } };
  const boxFont = { name: 'Arial', size: 9, bold: true, color: { argb: DARK_BLUE } };

  // ── Largeurs de colonnes (14 colonnes pour A4 paysage) ─────────────────────
  // A=2, B=15, C=5, D=5, E=5, F=5, G=5, H=10, I=10, J=10, K=10, L=10, M=10, N=5
  ws.columns = [
    { width: 2 },    // A - marge
    { width: 16 },   // B - labels
    { width: 5.5 },  // C - case 1
    { width: 5.5 },  // D - case 2
    { width: 5.5 },  // E - case 3
    { width: 5.5 },  // F - case 4
    { width: 14 },   // G - détail commande
    { width: 14 },   // H - détail commande
    { width: 10 },   // I - quantité
    { width: 13 },   // J - prix unitaire
    { width: 13 },   // K - total
    { width: 9 },    // L - code fournisseur
    { width: 9 },    // M - code fournisseur
    { width: 2 },    // N - marge
  ];

  // ── Helper: appliquer bordures à une plage ─────────────────────────────────
  function setBorders(startRow, startCol, endRow, endCol) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = ws.getCell(r, c);
        cell.border = allBorders;
      }
    }
  }

  function setOuterBorders(startRow, startCol, endRow, endCol) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = ws.getCell(r, c);
        const border = {};
        if (r === startRow) border.top = thinBorder;
        if (r === endRow) border.bottom = thinBorder;
        if (c === startCol) border.left = thinBorder;
        if (c === endCol) border.right = thinBorder;
        cell.border = border;
      }
    }
  }

  // ── Helper: fusionner et styler ────────────────────────────────────────────
  function mergeAndStyle(row, startCol, endCol, value, font, align, fill, border) {
    ws.mergeCells(row, startCol, row, endCol);
    const cell = ws.getCell(row, startCol);
    cell.value = value;
    cell.font = font || labelFont;
    cell.alignment = { horizontal: align || 'left', vertical: 'middle', wrapText: true };
    if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    if (border !== false) {
      for (let c = startCol; c <= endCol; c++) ws.getCell(row, c).border = allBorders;
    }
  }

  let row = 1;

  // ══════════════════════════════════════════════════════════════════════════
  // EN-TÊTE : Logo + Nom ASECNA | Adresse
  // ══════════════════════════════════════════════════════════════════════════

  // Logo ASECNA
  const logoPath = path.join(__dirname, 'public', 'ASECNA_logo.png');
  if (fs.existsSync(logoPath)) {
    const logoId = wb.addImage({
      filename: logoPath,
      extension: 'png',
    });
    ws.addImage(logoId, {
      tl: { col: 1, row: 0 },
      ext: { width: 55, height: 55 }
    });
  }

  // Ligne 1 : En-tête
  ws.getRow(row).height = 20;
  ws.mergeCells(row, 2, row, 4);
  ws.getCell(row, 2).value = 'AGENCE POUR LA SECURITE';
  ws.getCell(row, 2).font = headerFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle', indent: 4 };

  ws.mergeCells(row, 7, row + 2, 13);
  ws.getCell(row, 7).value = 'B.P. 3144 DAKAR (Sénégal)\nReprésentations de : LIBREVILLE';
  ws.getCell(row, 7).font = { name: 'Arial', size: 7.5, color: { argb: BLACK } };
  ws.getCell(row, 7).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  setBorders(row, 7, row + 2, 13);

  // Bordures gauche en-tête
  setBorders(row, 2, row + 2, 6);

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 4);
  ws.getCell(row, 2).value = 'DE LA NAVIGATION AERIENNE';
  ws.getCell(row, 2).font = headerFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle', indent: 4 };

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 4);
  ws.getCell(row, 2).value = 'EN AFRIQUE ET A MADAGASCAR';
  ws.getCell(row, 2).font = headerFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle', indent: 4 };

  // ── TITRE : BON DE COMMANDE N° ────────────────────────────────────────────
  row++;
  ws.getRow(row).height = 28;
  mergeAndStyle(row, 2, 13, 'BON DE COMMANDE    N°', titleFont, 'center', null, true);
  setBorders(row, 2, row, 13);

  // ══════════════════════════════════════════════════════════════════════════
  // CODES BUDGÉTAIRES
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 3; // Espacement
  row++;

  // CS
  ws.getRow(row).height = 17;
  ws.getCell(row, 2).value = 'CENTRE DE SYNTHESE (C.S)';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
  // 3 cases
  for (let i = 0; i < 3; i++) {
    const cell = ws.getCell(row, 3 + i);
    cell.border = allBorders;
    cell.font = boxFont;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  row++;
  ws.getRow(row).height = 17;
  ws.getCell(row, 2).value = 'CENTRE DE RESPONSABILITE (C.R)';
  ws.getCell(row, 2).font = labelFont;
  for (let i = 0; i < 3; i++) {
    ws.getCell(row, 3 + i).border = allBorders;
    ws.getCell(row, 3 + i).font = boxFont;
    ws.getCell(row, 3 + i).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  row++;
  ws.getRow(row).height = 17;
  ws.getCell(row, 2).value = 'CENTRE DE COUT (C.C)';
  ws.getCell(row, 2).font = labelFont;
  for (let i = 0; i < 3; i++) {
    ws.getCell(row, 3 + i).border = allBorders;
    ws.getCell(row, 3 + i).font = boxFont;
    ws.getCell(row, 3 + i).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  row++;
  ws.getRow(row).height = 17;
  ws.getCell(row, 2).value = 'ARTICLE';
  ws.getCell(row, 2).font = labelFont;
  for (let i = 0; i < 3; i++) {
    ws.getCell(row, 3 + i).border = allBorders;
    ws.getCell(row, 3 + i).font = boxFont;
    ws.getCell(row, 3 + i).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  row++;
  ws.getRow(row).height = 17;
  ws.getCell(row, 2).value = 'EXERCICE';
  ws.getCell(row, 2).font = labelFont;
  for (let i = 0; i < 4; i++) {
    ws.getCell(row, 3 + i).border = allBorders;
    ws.getCell(row, 3 + i).font = boxFont;
    ws.getCell(row, 3 + i).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FOURNISSEUR
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 3; // Espacement
  row++;

  const fournRow = row;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).value = 'ADRESSE DU FOURNISSEUR :';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 9, row, 11);
  ws.getCell(row, 9).value = 'CODE FOURNISSEUR';
  ws.getCell(row, 9).font = labelFont;
  ws.getCell(row, 9).alignment = { horizontal: 'center', vertical: 'middle' };

  // Cases code fournisseur (4 cases sur la même ligne)
  // On va mettre 4 petites cases : L12, L13 + M12, M13 → utilisons col 12,13 de la ligne suivante
  row++;
  ws.getRow(row).height = 18;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).font = valueFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
  // 4 cases code fournisseur
  for (let i = 0; i < 4; i++) {
    ws.getCell(row, 9 + i).border = allBorders;
    ws.getCell(row, 9 + i).font = boxFont;
    ws.getCell(row, 9 + i).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).font = valueFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).font = valueFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  // Bordures section fournisseur
  setOuterBorders(fournRow, 2, row, 8);
  setOuterBorders(fournRow, 9, row, 13);

  // ══════════════════════════════════════════════════════════════════════════
  // AVIS TRES IMPORTANT
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 3;
  row++;

  ws.getRow(row).height = 22;
  mergeAndStyle(row, 2, 13,
    'AVIS TRES IMPORTANT — LE PRESENT BON N\'ENGAGE L\'ASECNA QUE S\'IL COMPORTE LE NUMERO D\'ENGAGEMENT, LE VISA ET LE CACHET DU SERVICE DES ENGAGEMENTS DE L\'ASECNA',
    { name: 'Arial', size: 7, bold: false, color: { argb: BLACK } },
    'left', null, true
  );

  // ══════════════════════════════════════════════════════════════════════════
  // TABLEAU DETAIL DE LA COMMANDE
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 3;
  row++;

  // En-tête tableau
  const detailStartRow = row;
  ws.getRow(row).height = 20;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).value = 'DETAIL DE LA COMMANDE';
  ws.getCell(row, 2).font = labelFontBold;
  ws.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_GRAY } };

  ws.getCell(row, 9).value = 'QUANTITE';
  ws.getCell(row, 9).font = labelFontBold;
  ws.getCell(row, 9).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_GRAY } };

  ws.mergeCells(row, 10, row, 11);
  ws.getCell(row, 10).value = 'PRIX UNITAIRE';
  ws.getCell(row, 10).font = labelFontBold;
  ws.getCell(row, 10).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 10).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_GRAY } };

  ws.mergeCells(row, 12, row, 13);
  ws.getCell(row, 12).value = 'TOTAL';
  ws.getCell(row, 12).font = labelFontBold;
  ws.getCell(row, 12).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 12).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_GRAY } };

  setBorders(row, 2, row, 13);

  // 3 lignes de commande vides
  for (let li = 0; li < 3; li++) {
    row++;
    ws.getRow(row).height = 22;
    ws.mergeCells(row, 2, row, 8);
    ws.getCell(row, 2).font = valueFont;
    ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

    ws.getCell(row, 9).font = valueFontSmall;
    ws.getCell(row, 9).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(row, 10, row, 11);
    ws.getCell(row, 10).font = valueFontSmall;
    ws.getCell(row, 10).alignment = { horizontal: 'right', vertical: 'middle' };

    ws.mergeCells(row, 12, row, 13);
    ws.getCell(row, 12).font = valueFontSmall;
    ws.getCell(row, 12).alignment = { horizontal: 'right', vertical: 'middle' };

    setBorders(row, 2, row, 13);
  }

  // Ligne MONTANT TOTAL EN CHIFFRE
  row++;
  ws.getRow(row).height = 20;
  ws.mergeCells(row, 2, row, 11);
  ws.getCell(row, 2).value = 'MONTANT TOTAL EN CHIFFRE';
  ws.getCell(row, 2).font = labelFontBold;
  ws.getCell(row, 2).alignment = { horizontal: 'right', vertical: 'middle' };

  ws.mergeCells(row, 12, row, 13);
  ws.getCell(row, 12).font = valueFont;
  ws.getCell(row, 12).alignment = { horizontal: 'right', vertical: 'middle' };
  setBorders(row, 2, row, 13);

  // ══════════════════════════════════════════════════════════════════════════
  // MONTANT EN LETTRES + DELAI
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 20;
  ws.mergeCells(row, 2, row, 13);
  ws.getCell(row, 2).value = 'MONTANT TOTAL EN LETTRE : ';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
  setBorders(row, 2, row, 13);

  row++;
  ws.getRow(row).height = 20;
  ws.mergeCells(row, 2, row, 13);
  ws.getCell(row, 2).value = 'DELAI DE LIVRAISON :                    Passé ce délai, l\'ASECNA se réserve le droit de considérer le présent bon comme nul';
  ws.getCell(row, 2).font = { name: 'Arial', size: 7, color: { argb: BLACK } };
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
  setBorders(row, 2, row, 13);

  // ══════════════════════════════════════════════════════════════════════════
  // NOTA
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 22;
  mergeAndStyle(row, 2, 13,
    'NOTA : LES FACTURES, AVEC MENTION DES PRIX UNITAIRES DOIVENT ETRE ADRESSEES EN 4 EXEMPLAIRES ACCOMPAGNEES DU BON DE COMMANDE ORIGINAL ET D\'UN EXEMPLAIRE DU B.L DUMENT DE CHARGE',
    { name: 'Arial', size: 6.5, bold: false, color: { argb: BLACK } },
    'left', null, true
  );

  // ══════════════════════════════════════════════════════════════════════════
  // ZONE DE VALIDATION
  // ══════════════════════════════════════════════════════════════════════════
  row++;
  ws.getRow(row).height = 3;
  row++;

  const valStartRow = row;

  // Ligne 1 validation : Compte limitatif + Lieu date | Lieu date droite
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).value = 'COMPTE LIMITATIF :          A :                    LE :';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 9, row, 13);
  ws.getCell(row, 9).value = 'A :                    LE :';
  ws.getCell(row, 9).font = labelFont;
  ws.getCell(row, 9).alignment = { horizontal: 'left', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).value = 'OPERATION :';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 2, row, 8);
  ws.getCell(row, 2).value = 'N° D\'ENGAGEMENT :';
  ws.getCell(row, 2).font = labelFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 12;
  ws.mergeCells(row, 2, row, 5);
  ws.getCell(row, 2).value = '1er exemplaire original à retourner';
  ws.getCell(row, 2).font = smallFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 6, row, 8);
  ws.getCell(row, 6).value = 'VISA ET CACHET DU SERVICE';
  ws.getCell(row, 6).font = labelFontBold;
  ws.getCell(row, 6).alignment = { horizontal: 'center', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 12;
  ws.mergeCells(row, 2, row, 5);
  ws.getCell(row, 2).value = '2e exemplaire copie à conserver';
  ws.getCell(row, 2).font = smallFont;
  ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 6, row, 8);
  ws.getCell(row, 6).value = 'DES ENGAGEMENTS DE L\'ASECNA';
  ws.getCell(row, 6).font = labelFontBold;
  ws.getCell(row, 6).alignment = { horizontal: 'center', vertical: 'middle' };

  row++;
  ws.getRow(row).height = 14;
  ws.mergeCells(row, 6, row, 8);
  ws.getCell(row, 6).value = 'A :              LE :';
  ws.getCell(row, 6).font = labelFont;
  ws.getCell(row, 6).alignment = { horizontal: 'center', vertical: 'middle' };

  // Espace pour cachet/signature
  row++;
  ws.getRow(row).height = 30;

  row++;
  ws.getRow(row).height = 14;
  ws.mergeCells(row, 9, row, 13);
  ws.getCell(row, 9).value = 'VISA ET CACHET DE L\'ORDONNATEUR';
  ws.getCell(row, 9).font = labelFontBold;
  ws.getCell(row, 9).alignment = { horizontal: 'center', vertical: 'middle' };

  const valEndRow = row;

  // Bordures zone de validation
  setOuterBorders(valStartRow, 2, valEndRow, 8);
  setOuterBorders(valStartRow, 9, valEndRow, 13);
  // Séparation verticale entre gauche et droite
  for (let r = valStartRow; r <= valEndRow; r++) {
    const c8 = ws.getCell(r, 8);
    c8.border = { ...c8.border, right: thinBorder };
    const c9 = ws.getCell(r, 9);
    c9.border = { ...c9.border, left: thinBorder };
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // PAGE 2 : BON D'ENGAGEMENT (coupon interne)
  // ═══════════════════════════════════════════════════════════════════════════
  const ws2 = wb.addWorksheet('BON D\'ENGAGEMENT', {
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: {
        left: 0.5, right: 0.5,
        top: 0.4, bottom: 0.4,
        header: 0.2, footer: 0.2
      }
    },
    properties: {
      defaultRowHeight: 15,
      showGridLines: false
    }
  });

  ws2.columns = [
    { width: 2 },    // A
    { width: 16 },   // B
    { width: 5.5 },  // C
    { width: 5.5 },  // D
    { width: 5.5 },  // E
    { width: 5.5 },  // F
    { width: 12 },   // G
    { width: 12 },   // H
    { width: 12 },   // I
    { width: 12 },   // J
    { width: 12 },   // K
    { width: 12 },   // L
    { width: 2 },    // M
  ];

  row = 1;

  // Titre
  ws2.getRow(row).height = 28;
  ws2.mergeCells(row, 2, row, 12);
  ws2.getCell(row, 2).value = 'BON DE COMMANDE';
  ws2.getCell(row, 2).font = titleFont;
  ws2.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
  setBordersWs(ws2, row, 2, row, 12);

  row++;
  ws2.getRow(row).height = 4;
  row++;

  // Codes comptables (miroir)
  const codesP2 = [
    ['C.S', 3], ['C.R', 3], ['C.C', 3], ['ARTICLE', 1]
  ];
  for (const [label, count] of codesP2) {
    ws2.getRow(row).height = 17;
    ws2.getCell(row, 2).value = label;
    ws2.getCell(row, 2).font = labelFont;
    ws2.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
    for (let i = 0; i < count; i++) {
      ws2.getCell(row, 3 + i).border = allBorders;
      ws2.getCell(row, 3 + i).font = boxFont;
      ws2.getCell(row, 3 + i).alignment = { horizontal: 'center', vertical: 'middle' };
    }
    row++;
  }

  // Fournisseur
  ws2.getRow(row).height = 4;
  row++;
  ws2.getRow(row).height = 18;
  ws2.mergeCells(row, 2, row, 12);
  ws2.getCell(row, 2).value = 'FOURNISSEUR :';
  ws2.getCell(row, 2).font = labelFont;
  ws2.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
  setBordersWs(ws2, row, 2, row, 12);

  // Champs coupon
  row++;
  ws2.getRow(row).height = 4;
  row++;

  const couponFields = [
    'CODE INDIVIDUEL',
    'NUMERO D\'ENGAGEMENT',
    'COMPTE LIMITATIF',
    'OPERATION D\'EQUIPEMENT',
    'COMPTE DE / COMPTABILITE GENERALE',
  ];
  for (const label of couponFields) {
    ws2.getRow(row).height = 18;
    ws2.mergeCells(row, 2, row, 6);
    ws2.getCell(row, 2).value = label;
    ws2.getCell(row, 2).font = labelFont;
    ws2.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };

    ws2.mergeCells(row, 7, row, 12);
    ws2.getCell(row, 7).font = valueFont;
    ws2.getCell(row, 7).alignment = { horizontal: 'left', vertical: 'middle' };
    setBordersWs(ws2, row, 2, row, 12);
    row++;
  }

  // Tableau financier
  row++;
  ws2.getRow(row).height = 20;
  const finHeaders = ['MONTANT A.D.', 'MONTANT DU BON', 'ENG. ANTERIEURS', 'CUMUL ENG.', 'DISPONIBLE'];
  // 5 colonnes : on prend col 2-3, 4-5, 6-7, 8-9, 10-11 (ou 2 cols chacune)
  // Ajustons : col B-C, D-E, F-G, H-I, J-K pour 5 headers, col L reste
  // En fait utilisons col 2 à 12 : chaque header prend ~2.2 cols → arrondissons
  // Mieux : headers dans une seule ligne, valeurs en dessous
  for (let i = 0; i < 5; i++) {
    const colStart = 2 + i * 2;
    const colEnd = colStart + 1;
    if (colEnd <= 12) {
      ws2.mergeCells(row, colStart, row, colEnd);
    }
    ws2.getCell(row, colStart).value = finHeaders[i];
    ws2.getCell(row, colStart).font = { name: 'Arial', size: 7, bold: true, color: { argb: BLACK } };
    ws2.getCell(row, colStart).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    ws2.getCell(row, colStart).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_GRAY } };
    ws2.getCell(row, colStart).border = allBorders;
    if (colEnd <= 12) ws2.getCell(row, colEnd).border = allBorders;
  }
  // Col 12 restante
  ws2.mergeCells(row, 12, row, 12);

  row++;
  ws2.getRow(row).height = 22;
  for (let i = 0; i < 5; i++) {
    const colStart = 2 + i * 2;
    const colEnd = colStart + 1;
    if (colEnd <= 12) {
      ws2.mergeCells(row, colStart, row, colEnd);
    }
    ws2.getCell(row, colStart).font = valueFontSmall;
    ws2.getCell(row, colStart).alignment = { horizontal: 'center', vertical: 'middle' };
    ws2.getCell(row, colStart).border = allBorders;
    if (colEnd <= 12) ws2.getCell(row, colEnd).border = allBorders;
  }

  // Lieu et date + signature
  row += 2;
  ws2.getRow(row).height = 16;
  ws2.mergeCells(row, 2, row, 12);
  ws2.getCell(row, 2).value = 'A __________  LE __________';
  ws2.getCell(row, 2).font = labelFont;
  ws2.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };

  row += 3;
  ws2.getRow(row).height = 16;
  ws2.mergeCells(row, 2, row, 12);
  ws2.getCell(row, 2).value = 'Signature';
  ws2.getCell(row, 2).font = labelFont;
  ws2.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };

  // Helper pour setBorders sur ws2
  function setBordersWs(sheet, startRow, startCol, endRow, endCol) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        sheet.getCell(r, c).border = allBorders;
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SAUVEGARDE
  // ═══════════════════════════════════════════════════════════════════════════
  const outPath = path.join(__dirname, 'public', 'BON DE COMMANDE A4.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log('Template Excel créé :', outPath);
  console.log('Taille :', Math.round(fs.statSync(outPath).size / 1024), 'Ko');
}

createBonCommandeTemplate().catch(err => {
  console.error('Erreur:', err);
  process.exit(1);
});
