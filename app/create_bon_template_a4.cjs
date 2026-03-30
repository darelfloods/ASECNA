/**
 * Crée le template "BON DE COMMANDE A4.docx" en format A4 paysage
 * avec tous les paraIds correspondant à ceux utilisés dans wordParser.ts.
 *
 * Exécuter : node create_bon_template_a4.cjs
 */

const PizZip = require('pizzip');
const fs = require('fs');
const path = require('path');

// ══════════════════════════════════════════════════════════════════════════════
// Dimensions A4 paysage en twips (1 twip = 1/1440 pouce)
// ══════════════════════════════════════════════════════════════════════════════
const PAGE_W = 16838; // 297mm
const PAGE_H = 11906; // 210mm
const MARGIN_TOP = 400;
const MARGIN_BOTTOM = 400;
const MARGIN_LEFT = 600;
const MARGIN_RIGHT = 600;
const CONTENT_W = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT; // 15638

// ══════════════════════════════════════════════════════════════════════════════
// Helpers XML
// ══════════════════════════════════════════════════════════════════════════════

/** Paragraphe avec paraId */
function para(paraId, content, pPrExtra = '') {
  return `<w:p w14:paraId="${paraId}" w14:textId="77777777"><w:pPr>${pPrExtra}</w:pPr>${content}</w:p>`;
}

/** Paragraphe vide (placeholder pour injection) */
function emptyPara(paraId, pPrExtra = '') {
  return `<w:p w14:paraId="${paraId}" w14:textId="77777777"><w:pPr>${pPrExtra}</w:pPr></w:p>`;
}

/** Run texte */
function run(text, rPrExtra = '') {
  const safe = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<w:r><w:rPr>${rPrExtra}</w:rPr><w:t xml:space="preserve">${safe}</w:t></w:r>`;
}

/** rPr pour Arial, taille en demi-points */
function rPr(sizeHalf, bold = false, color = '000000') {
  let s = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="${sizeHalf}"/><w:szCs w:val="${sizeHalf}"/>`;
  if (bold) s += '<w:b/><w:bCs/>';
  if (color !== '000000') s += `<w:color w:val="${color}"/>`;
  return s;
}

/** Propriétés de paragraphe : espacement serré */
function spacing(before = 0, after = 0, line = 240) {
  return `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`;
}

function jc(val) { return `<w:jc w:val="${val}"/>`; }

/** Cellule de tableau avec largeur */
function tc(width, content, tcPrExtra = '') {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPrExtra}</w:tcPr>${content}</w:tc>`;
}

/** Cellule avec bordures */
function tcBordered(width, content, tcPrExtra = '') {
  const borders = '<w:tcBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '</w:tcBorders>';
  return tc(width, content, borders + tcPrExtra);
}

/** Ligne de tableau */
function tr(cells, trPrExtra = '') {
  return `<w:tr>${trPrExtra ? '<w:trPr>' + trPrExtra + '</w:trPr>' : ''}${cells}</w:tr>`;
}

/** Tableau avec propriétés */
function tbl(rows, tblPrExtra = '') {
  const defaultBorders = '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
    '</w:tblBorders>';
  return `<w:tbl><w:tblPr><w:tblW w:w="${CONTENT_W}" w:type="dxa"/>${defaultBorders}<w:tblLook w:val="04A0"/>${tblPrExtra}</w:tblPr>${rows}</w:tbl>`;
}

/** Tableau sans bordures */
function tblNoBorders(rows, width = CONTENT_W, tblPrExtra = '') {
  const noBorders = '<w:tblBorders>' +
    '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '</w:tblBorders>';
  return `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/>${noBorders}<w:tblLook w:val="04A0"/>${tblPrExtra}</w:tblPr>${rows}</w:tbl>`;
}

/** Simple text paragraph (no paraId needed) */
function textPara(text, pPrExtra = '', rPrExtra = '') {
  return `<w:p><w:pPr>${pPrExtra}</w:pPr>${run(text, rPrExtra)}</w:p>`;
}

/** Case individuelle (petite cellule 1 caractère) bordurée, avec paraId */
function charCell(width, paraId) {
  const content = emptyPara(paraId, spacing(0, 0) + jc('center'));
  return tcBordered(width, content, '<w:vAlign w:val="center"/>');
}

/** Label cell (pas de paraId, texte fixe) */
function labelCell(width, text, bold = false) {
  const content = textPara(text, spacing(0, 0), rPr(16, bold));
  return tcBordered(width, content, '<w:vAlign w:val="center"/>');
}

/** Cellule sans bordure avec texte */
function noBorderCell(width, text, pPrExtra = '', rPrExtra = '') {
  const noBorders = '<w:tcBorders>' +
    '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
    '</w:tcBorders>';
  const content = textPara(text, pPrExtra, rPrExtra);
  return tc(width, content, noBorders);
}

// ══════════════════════════════════════════════════════════════════════════════
// Construction du document.xml
// ══════════════════════════════════════════════════════════════════════════════

function buildDocumentXml() {
  const parts = [];

  const S8 = rPr(16);        // 8pt
  const S8B = rPr(16, true); // 8pt bold
  const S9 = rPr(18);        // 9pt
  const S9B = rPr(18, true); // 9pt bold
  const S10 = rPr(20);       // 10pt
  const S10B = rPr(20, true);// 10pt bold
  const S12B = rPr(24, true);// 12pt bold
  const S7 = rPr(14);        // 7pt
  const S6 = rPr(12);        // 6pt

  const CHAR_CELL_W = 400;   // largeur d'une case 1 caractère
  const LABEL_COL_W = 5000;  // largeur colonne labels codes

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 1 : EN-TÊTE
  // ────────────────────────────────────────────────────────────────────────────
  {
    const leftW = Math.floor(CONTENT_W * 0.55);
    const rightW = CONTENT_W - leftW;

    const leftContent =
      textPara('[LOGO ASECNA]', spacing(0, 40), S8B) +
      textPara('AGENCE POUR LA SECURITE', spacing(0, 0), S8B) +
      textPara('DE LA NAVIGATION AERIENNE', spacing(0, 0), S8B) +
      textPara('EN AFRIQUE ET A MADAGASCAR', spacing(0, 40), S8B);

    const rightContent =
      textPara('ADRESSE ASECNA', spacing(0, 0) + jc('right'), S8);

    const noBord = '<w:tcBorders>' +
      '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '</w:tcBorders>';

    const headerRow = tr(
      tc(leftW, leftContent, noBord) +
      tc(rightW, rightContent, noBord)
    );

    parts.push(tblNoBorders(headerRow));

    // Titre BON DE COMMANDE N°
    parts.push(
      textPara('', spacing(80, 0)) // petit espace
    );
    parts.push(
      `<w:p><w:pPr>${spacing(0, 80)}${jc('center')}</w:pPr>${run('BON DE COMMANDE N\u00B0', S12B)}</w:p>`
    );
  }

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 2 : CODES BUDGÉTAIRES
  // ────────────────────────────────────────────────────────────────────────────
  {
    // CS
    const csRow = tr(
      labelCell(LABEL_COL_W, 'CENTRE DE SYNTHESE (C.S)') +
      charCell(CHAR_CELL_W, '278F4863') +
      charCell(CHAR_CELL_W, '3ED8D4CF') +
      charCell(CHAR_CELL_W, '1F68D0AF') +
      tcBordered(CONTENT_W - LABEL_COL_W - 3 * CHAR_CELL_W, emptyPara('00000001', spacing(0, 0)))
    );

    // CR
    const crRow = tr(
      labelCell(LABEL_COL_W, 'CENTRE DE RESPONSABILITE (C.R)') +
      charCell(CHAR_CELL_W, '76A09288') +
      charCell(CHAR_CELL_W, '0922550F') +
      charCell(CHAR_CELL_W, '0525D86D') +
      tcBordered(CONTENT_W - LABEL_COL_W - 3 * CHAR_CELL_W, emptyPara('00000002', spacing(0, 0)))
    );

    // CC
    const ccRow = tr(
      labelCell(LABEL_COL_W, 'CENTRE DE COUT (C.C)') +
      charCell(CHAR_CELL_W, '464D2853') +
      charCell(CHAR_CELL_W, '1910E1CE') +
      charCell(CHAR_CELL_W, '37B353A9') +
      tcBordered(CONTENT_W - LABEL_COL_W - 3 * CHAR_CELL_W, emptyPara('00000003', spacing(0, 0)))
    );

    // Article
    const artRow = tr(
      labelCell(LABEL_COL_W, 'ARTICLE') +
      charCell(CHAR_CELL_W, '2BCFE059') +
      charCell(CHAR_CELL_W, '60DE73FF') +
      charCell(CHAR_CELL_W, '5114EA48') +
      tcBordered(CONTENT_W - LABEL_COL_W - 3 * CHAR_CELL_W, emptyPara('00000004', spacing(0, 0)))
    );

    // Exercice (4 cases)
    const exoRow = tr(
      labelCell(LABEL_COL_W, 'EXERCICE') +
      charCell(CHAR_CELL_W, '2714A278') +
      charCell(CHAR_CELL_W, '5DA999ED') +
      charCell(CHAR_CELL_W, '4E830C1D') +
      charCell(CHAR_CELL_W, '14748769') +
      tcBordered(CONTENT_W - LABEL_COL_W - 4 * CHAR_CELL_W, emptyPara('00000005', spacing(0, 0)))
    );

    parts.push(tbl(csRow + crRow + ccRow + artRow + exoRow));
  }

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 3 : FOURNISSEUR
  // ────────────────────────────────────────────────────────────────────────────
  {
    const leftW = Math.floor(CONTENT_W * 0.65);
    const rightW = CONTENT_W - leftW;

    const fournisseurLeft =
      para('7215E09D', run('ADRESSE DU FOURNISSEUR :', S8B), spacing(20, 0)) +
      emptyPara('4C56EDD5', spacing(0, 0)) +
      emptyPara('3DEED587', spacing(0, 0)) +
      emptyPara('0E826B05', spacing(0, 0));

    // Code fournisseur : label + 4 cases
    const codeFournLabel = textPara('CODE FOURNISSEUR', spacing(20, 40) + jc('center'), S8B);

    // Nested table for code fournisseur cases
    const codeRow = tr(
      charCell(CHAR_CELL_W, '1A6809BF') +
      charCell(CHAR_CELL_W, '244A9EAF') +
      charCell(CHAR_CELL_W, '5DB65F30') +
      charCell(CHAR_CELL_W, '386E9672')
    );
    const codeTable = `<w:tbl><w:tblPr><w:tblW w:w="${4 * CHAR_CELL_W}" w:type="dxa"/>` +
      '<w:tblBorders>' +
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
      '</w:tblBorders>' +
      `<w:jc w:val="center"/><w:tblLook w:val="04A0"/></w:tblPr>${codeRow}</w:tbl>`;

    const fournisseurRight = codeFournLabel + codeTable + textPara('', spacing(0, 0));

    const noBord = '<w:tcBorders>' +
      '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '</w:tcBorders>';

    const fRow = tr(
      tc(leftW, fournisseurLeft, noBord) +
      tc(rightW, fournisseurRight, noBord)
    );

    parts.push(textPara('', spacing(60, 0))); // espace
    parts.push(tblNoBorders(fRow));
  }

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 4 : AVIS TRÈS IMPORTANT
  // ────────────────────────────────────────────────────────────────────────────
  parts.push(textPara('', spacing(60, 0)));
  parts.push(
    `<w:p><w:pPr>${spacing(0, 40)}${jc('center')}</w:pPr>` +
    run('AVIS TRES IMPORTANT \u2014 LE PRESENT BON N\'ENGAGE L\'ASECNA QUE S\'IL COMPORTE LE NUMERO D\'ENGAGEMENT, LE VISA ET LE CACHET DU SERVICE DES ENGAGEMENTS DE L\'ASECNA', S8B) +
    '</w:p>'
  );

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 5 : TABLEAU DÉTAIL COMMANDE
  // ────────────────────────────────────────────────────────────────────────────
  {
    const descW = Math.floor(CONTENT_W * 0.50);
    const qtyW = Math.floor(CONTENT_W * 0.13);
    const prixW = Math.floor(CONTENT_W * 0.18);
    const totalW = CONTENT_W - descW - qtyW - prixW;

    // Header
    const headerRow = tr(
      tcBordered(descW, textPara('DETAIL DE LA COMMANDE', spacing(0, 0) + jc('center'), S8B), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(qtyW, textPara('QUANTITE', spacing(0, 0) + jc('center'), S8B), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(prixW, textPara('PRIX UNITAIRE', spacing(0, 0) + jc('center'), S8B), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(totalW, textPara('TOTAL', spacing(0, 0) + jc('center'), S8B), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>')
    );

    // Ligne 1
    const ligne1 = tr(
      tcBordered(descW, emptyPara('2B8EB34A', spacing(0, 0)), '<w:vAlign w:val="center"/>') +
      tcBordered(qtyW, emptyPara('1C80DE40', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(prixW, emptyPara('16DAF1DA', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(totalW, emptyPara('687CFF4D', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>'),
      '<w:trHeight w:val="400" w:hRule="atLeast"/>'
    );

    // Ligne 2
    const ligne2 = tr(
      tcBordered(descW, emptyPara('274BF247', spacing(0, 0)), '<w:vAlign w:val="center"/>') +
      tcBordered(qtyW, emptyPara('279A4B16', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(prixW, emptyPara('653F54CB', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(totalW, emptyPara('49A4ECBA', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>'),
      '<w:trHeight w:val="400" w:hRule="atLeast"/>'
    );

    // Ligne 3
    const ligne3 = tr(
      tcBordered(descW, emptyPara('2CF4905F', spacing(0, 0)), '<w:vAlign w:val="center"/>') +
      tcBordered(qtyW, emptyPara('3D534AA1', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(prixW, emptyPara('3F62CC38', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(totalW, emptyPara('4C414FFE', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>'),
      '<w:trHeight w:val="400" w:hRule="atLeast"/>'
    );

    // Ligne MONTANT TOTAL EN CHIFFRE
    const totalRow = tr(
      tcBordered(descW + qtyW + prixW,
        textPara('MONTANT TOTAL EN CHIFFRE', spacing(0, 0) + jc('right'), S8B),
        `<w:gridSpan w:val="3"/><w:vAlign w:val="center"/>`) +
      tcBordered(totalW, emptyPara('3A6EA764', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>')
    );

    parts.push(tbl(headerRow + ligne1 + ligne2 + ligne3 + totalRow));
  }

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 6 : MONTANT EN LETTRES / DÉLAI
  // ────────────────────────────────────────────────────────────────────────────
  parts.push(textPara('', spacing(60, 0)));
  parts.push(
    `<w:p><w:pPr>${spacing(0, 20)}</w:pPr>` +
    run('MONTANT TOTAL EN LETTRE : ', S8B) +
    run('\u2026'.repeat(59) + '...', S8) +
    '</w:p>'
  );
  parts.push(
    `<w:p><w:pPr>${spacing(0, 20)}</w:pPr>` +
    run('DELAI DE LIVRAISON : ', S8B) +
    run('\u2026'.repeat(23), S8) +
    run(' Pass\u00E9 ce d\u00E9lai, l\'ASECNA se r\u00E9serve le droit d\'annuler le pr\u00E9sent Bon de Commande', S7) +
    '</w:p>'
  );

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 7 : NOTA
  // ────────────────────────────────────────────────────────────────────────────
  parts.push(textPara('', spacing(40, 0)));
  parts.push(
    `<w:p><w:pPr>${spacing(0, 20)}</w:pPr>` +
    run('NOTA : ', S8B) +
    run('LES FACTURES, AVEC MENTION DES PRIX UNITAIRES DOIVENT ETRE ADRESSEES EN 4 EXEMPLAIRES AU SERVICE DE L\'ENGAGEMENT DES DEPENSES DE L\'ASECNA', S7) +
    '</w:p>'
  );

  // ────────────────────────────────────────────────────────────────────────────
  // SECTION 8 : ZONE DE VALIDATION
  // ────────────────────────────────────────────────────────────────────────────
  {
    const leftW = Math.floor(CONTENT_W * 0.58);
    const rightW = CONTENT_W - leftW;

    const noBord = '<w:tcBorders>' +
      '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
      '</w:tcBorders>';

    const leftContent =
      textPara('COMPTE LIMITATIF', spacing(20, 0), S8B) +
      textPara('A\u00A0: ___________ LE : __________', spacing(0, 0), S8) +
      emptyPara('16A69E91', spacing(0, 0)) +
      textPara('N\u00B0 D\'ENGAGEMENT', spacing(40, 0), S8B) +
      textPara(' :', spacing(0, 0), S8) +
      textPara('   VISA ET CACHET DU SERVICE DES ENGAGEMENTS', spacing(0, 0), S8) +
      textPara('', spacing(0, 0)) +
      textPara('1er exemplaire : Fournisseur', spacing(0, 0), S6) +
      textPara('2\u00E8me exemplaire : Service Engagement des D\u00E9penses', spacing(0, 0), S6) +
      textPara('3\u00E8me exemplaire : Comptabilit\u00E9 G\u00E9n\u00E9rale', spacing(0, 0), S6) +
      textPara('4\u00E8me exemplaire : \u00C0 classer', spacing(0, 0), S6);

    const rightContent =
      textPara('A\u00A0: ___________ LE : __________', spacing(20, 0) + jc('center'), S8) +
      textPara('', spacing(0, 0)) +
      textPara('', spacing(0, 0)) +
      textPara('', spacing(0, 0)) +
      textPara('VISA ET CACHET DE L\'ORDONNATEUR', spacing(0, 0) + jc('center'), S8B);

    const validRow = tr(
      tc(leftW, leftContent, noBord) +
      tc(rightW, rightContent, noBord)
    );

    parts.push(textPara('', spacing(40, 0)));
    parts.push(tblNoBorders(validRow));
  }

  // ────────────────────────────────────────────────────────────────────────────
  // SAUT DE SECTION (nouvelle page)
  // ────────────────────────────────────────────────────────────────────────────
  const sectionBreak = `<w:p><w:pPr><w:sectPr>` +
    `<w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}" w:orient="landscape"/>` +
    `<w:pgMar w:top="${MARGIN_TOP}" w:right="${MARGIN_RIGHT}" w:bottom="${MARGIN_BOTTOM}" w:left="${MARGIN_LEFT}" w:header="400" w:footer="400" w:gutter="0"/>` +
    `<w:cols w:space="708"/>` +
    `<w:docGrid w:linePitch="360"/>` +
    `</w:sectPr></w:pPr></w:p>`;
  parts.push(sectionBreak);

  // ════════════════════════════════════════════════════════════════════════════
  // PAGE 2 : BON D'ENGAGEMENT (coupon interne)
  // ════════════════════════════════════════════════════════════════════════════

  // Titre
  parts.push(
    `<w:p><w:pPr>${spacing(0, 80)}${jc('center')}</w:pPr>` +
    run('BON DE COMMANDE', S12B) +
    '</w:p>'
  );
  // Numéro du bon (paraId pour injection)
  parts.push(emptyPara('7D63CF0B', spacing(0, 60) + jc('center')));

  // Codes comptables miroir (coupon)
  {
    const couponLabelW = 4000;
    const couponCharW = 380;
    const couponRestW = CONTENT_W - couponLabelW - 3 * couponCharW;

    // CS coupon
    const csRow = tr(
      labelCell(couponLabelW, 'C.S') +
      charCell(couponCharW, '039DF871') +
      charCell(couponCharW, '5E50C1CB') +
      charCell(couponCharW, '5B27722B') +
      tcBordered(couponRestW, emptyPara('00000010', spacing(0, 0)))
    );

    // CR coupon
    const crRow = tr(
      labelCell(couponLabelW, 'C.R') +
      charCell(couponCharW, '7BC61F09') +
      charCell(couponCharW, '7AA19F9E') +
      charCell(couponCharW, '25755A46') +
      tcBordered(couponRestW, emptyPara('00000011', spacing(0, 0)))
    );

    // CC coupon
    const ccRow = tr(
      labelCell(couponLabelW, 'C.C') +
      charCell(couponCharW, '7A0E7BE7') +
      charCell(couponCharW, '18B438E4') +
      charCell(couponCharW, '54F752ED') +
      tcBordered(couponRestW, emptyPara('00000012', spacing(0, 0)))
    );

    // Article coupon (1 case)
    const artRow = tr(
      labelCell(couponLabelW, 'ARTICLE') +
      charCell(couponCharW, '4C691C60') +
      tcBordered(couponRestW + 2 * couponCharW, emptyPara('00000013', spacing(0, 0)))
    );

    parts.push(tbl(csRow + crRow + ccRow + artRow));
  }

  // Fournisseur coupon
  parts.push(textPara('', spacing(40, 0)));
  parts.push(emptyPara('3793017D', spacing(0, 40)));

  // Champs coupon dans un tableau vertical
  {
    const labelW = Math.floor(CONTENT_W * 0.40);
    const valueW = CONTENT_W - labelW;

    const rows = [
      // Code individuel
      tr(
        labelCell(labelW, 'CODE INDIVIDUEL') +
        tcBordered(valueW, emptyPara('23CB2604', spacing(0, 0)), '<w:vAlign w:val="center"/>')
      ),
      // N° engagement
      tr(
        labelCell(labelW, 'NUMERO D\'ENGAGEMENT') +
        tcBordered(valueW, emptyPara('1E4115B2', spacing(0, 0)), '<w:vAlign w:val="center"/>')
      ),
      // Compte limitatif
      tr(
        labelCell(labelW, 'COMPTE LIMITATIF') +
        tcBordered(valueW, emptyPara('61CAD576', spacing(0, 0)), '<w:vAlign w:val="center"/>')
      ),
      // Opération d'équipement
      tr(
        labelCell(labelW, 'OPERATION D\'EQUIPEMENT') +
        tcBordered(valueW, emptyPara('63126403', spacing(0, 0)), '<w:vAlign w:val="center"/>')
      ),
      // Compte de
      tr(
        labelCell(labelW, 'COMPTE DE') +
        tcBordered(valueW, emptyPara('78CBC1A4', spacing(0, 0)), '<w:vAlign w:val="center"/>')
      ),
    ];

    parts.push(tbl(rows.join('')));
  }

  // Tableau financier coupon
  parts.push(textPara('', spacing(60, 0)));
  {
    const colW = Math.floor(CONTENT_W / 5);
    const lastColW = CONTENT_W - 4 * colW;

    const headerRow = tr(
      tcBordered(colW, textPara('MONTANT A.D.', spacing(0, 0) + jc('center'), S7), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(colW, textPara('MONTANT DU BON', spacing(0, 0) + jc('center'), S7), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(colW, textPara('ENGAGEMENTS ANTERIEURS', spacing(0, 0) + jc('center'), S7), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(colW, textPara('CUMUL DES ENGAGEMENTS', spacing(0, 0) + jc('center'), S7), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>') +
      tcBordered(lastColW, textPara('DISPONIBLE', spacing(0, 0) + jc('center'), S7), '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>')
    );

    const valueRow = tr(
      tcBordered(colW, emptyPara('3B39980F', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(colW, emptyPara('15911AAF', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(colW, emptyPara('6193C334', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(colW, emptyPara('4C3A7A18', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>') +
      tcBordered(lastColW, emptyPara('1FE6910C', spacing(0, 0) + jc('center')), '<w:vAlign w:val="center"/>'),
      '<w:trHeight w:val="400" w:hRule="atLeast"/>'
    );

    parts.push(tbl(headerRow + valueRow));
  }

  // Lieu et date + signature coupon
  parts.push(textPara('', spacing(60, 0)));
  parts.push(
    `<w:p><w:pPr>${spacing(0, 20)}${jc('right')}</w:pPr>` +
    run('A __________ LE ', S8) +
    '</w:p>'
  );
  parts.push(textPara('', spacing(0, 0)));
  parts.push(textPara('', spacing(0, 0)));
  parts.push(
    `<w:p><w:pPr>${spacing(0, 0)}${jc('right')}</w:pPr>` +
    run('Signature', S8) +
    '</w:p>'
  );

  // ════════════════════════════════════════════════════════════════════════════
  // Assemblage final
  // ════════════════════════════════════════════════════════════════════════════

  const bodyContent = parts.join('');

  // Section properties pour la dernière section (page 2)
  const lastSectPr =
    `<w:sectPr>` +
    `<w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}" w:orient="landscape"/>` +
    `<w:pgMar w:top="${MARGIN_TOP}" w:right="${MARGIN_RIGHT}" w:bottom="${MARGIN_BOTTOM}" w:left="${MARGIN_LEFT}" w:header="400" w:footer="400" w:gutter="0"/>` +
    `<w:cols w:space="708"/>` +
    `<w:docGrid w:linePitch="360"/>` +
    `</w:sectPr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14">
<w:body>
${bodyContent}
${lastSectPr}
</w:body>
</w:document>`;

  return documentXml;
}

// ══════════════════════════════════════════════════════════════════════════════
// Fichiers annexes du DOCX
// ══════════════════════════════════════════════════════════════════════════════

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="16"/>
        <w:szCs w:val="16"/>
        <w:lang w:val="fr-FR" w:eastAsia="fr-FR" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="15" w:type="dxa"/>
        <w:left w:w="30" w:type="dxa"/>
        <w:bottom w:w="15" w:type="dxa"/>
        <w:right w:w="30" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="w14">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>`;

// ══════════════════════════════════════════════════════════════════════════════
// Assemblage et sauvegarde
// ══════════════════════════════════════════════════════════════════════════════

function main() {
  console.log('Generating BON DE COMMANDE A4.docx ...');

  const documentXml = buildDocumentXml();

  const zip = new PizZip();
  zip.file('[Content_Types].xml', contentTypesXml);
  zip.file('_rels/.rels', relsXml);
  zip.file('word/document.xml', documentXml);
  zip.file('word/_rels/document.xml.rels', documentRelsXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/settings.xml', settingsXml);

  const output = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  const outputPath = path.join(__dirname, 'public', 'BON DE COMMANDE A4.docx');
  fs.writeFileSync(outputPath, output);

  console.log(`Template saved to: ${outputPath}`);
  console.log(`Size: ${(output.length / 1024).toFixed(1)} KB`);

  // Verify paraIds
  const verify = new PizZip(output);
  const verifyXml = verify.file('word/document.xml').asText();
  const requiredIds = [
    '278F4863', '3ED8D4CF', '1F68D0AF',  // CS
    '76A09288', '0922550F', '0525D86D',  // CR
    '464D2853', '1910E1CE', '37B353A9',  // CC
    '2BCFE059', '60DE73FF', '5114EA48',  // Article
    '2714A278', '5DA999ED', '4E830C1D', '14748769', // Exercice
    '1A6809BF', '244A9EAF', '5DB65F30', '386E9672', // Code fournisseur
    '7215E09D', '4C56EDD5', '3DEED587', '0E826B05', // Adresse fournisseur
    '2B8EB34A', '1C80DE40', '16DAF1DA', '687CFF4D', // Ligne 1
    '274BF247', '279A4B16', '653F54CB', '49A4ECBA', // Ligne 2
    '2CF4905F', '3D534AA1', '3F62CC38', '4C414FFE', // Ligne 3
    '3A6EA764', // Montant total chiffres
    '16A69E91', // Operation
    '7D63CF0B', // Coupon BON numéro
    '039DF871', '5E50C1CB', '5B27722B', // Coupon CS
    '7BC61F09', '7AA19F9E', '25755A46', // Coupon CR
    '7A0E7BE7', '18B438E4', '54F752ED', // Coupon CC
    '4C691C60', // Coupon Article
    '3793017D', // Coupon Fournisseur
    '23CB2604', // Code individuel
    '1E4115B2', // N° engagement coupon
    '61CAD576', // Compte limitatif coupon
    '63126403', // Opération équipement
    '78CBC1A4', // Compte de
    '3B39980F', // Montant AD
    '15911AAF', // Montant du bon
    '6193C334', // Engagements antérieurs
    '4C3A7A18', // Cumul engagements
    '1FE6910C', // Disponible
  ];

  let allFound = true;
  for (const id of requiredIds) {
    if (!verifyXml.includes(`w14:paraId="${id}"`)) {
      console.error(`  MISSING paraId: ${id}`);
      allFound = false;
    }
  }
  if (allFound) {
    console.log(`All ${requiredIds.length} required paraIds verified OK!`);
  }
}

main();
