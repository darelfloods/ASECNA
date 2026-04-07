import ExcelJS from "exceljs";

export interface ConventionData {
  "NOM DU CLIENT": string;
  "N° CONVENTION": string;
  "OBJET DE LA CONVENTION": string;
  SITE: string;
  "Date de debut": string;
  "Date de fin": string;
  "Durée": string;
  MONTANT: number;
}

// ---------------------------------------------------------------------------
// Conversion d'un nombre en toutes lettres (français)
// ---------------------------------------------------------------------------

const UNITS = [
  "", "UN", "DEUX", "TROIS", "QUATRE", "CINQ", "SIX", "SEPT", "HUIT", "NEUF",
  "DIX", "ONZE", "DOUZE", "TREIZE", "QUATORZE", "QUINZE", "SEIZE",
  "DIX-SEPT", "DIX-HUIT", "DIX-NEUF",
];

const TENS = [
  "", "DIX", "VINGT", "TRENTE", "QUARANTE", "CINQUANTE",
  "SOIXANTE", "SOIXANTE", "QUATRE-VINGTS", "QUATRE-VINGT",
];

function numberToFrenchWords(n: number): string {
  if (n === 0) return "ZÉRO";
  if (n < 0) return "MOINS " + numberToFrenchWords(-n);

  let result = "";

  if (n >= 1_000_000) {
    const millions = Math.floor(n / 1_000_000);
    const millionsWord = millions === 1
      ? "UN MILLION"
      : numberToFrenchWords(millions) + " MILLIONS";
    result += millionsWord + " ";
    n %= 1_000_000;
  }

  if (n >= 1_000) {
    const thousands = Math.floor(n / 1_000);
    if (thousands === 1) {
      result += "MILLE ";
    } else {
      result += numberToFrenchWords(thousands) + " MILLE ";
    }
    n %= 1_000;
  }

  if (n >= 100) {
    const hundreds = Math.floor(n / 100);
    if (hundreds === 1) {
      result += "CENT ";
    } else {
      result += UNITS[hundreds] + " CENT" + (n % 100 === 0 ? "S" : "") + " ";
    }
    n %= 100;
  }

  if (n >= 20) {
    const tensIndex = Math.floor(n / 10);
    const units = n % 10;

    if (tensIndex === 7) {
      // Soixante-dix: 60 + 10..19
      const subVal = 10 + units;
      result += "SOIXANTE-" + UNITS[subVal] + " ";
      n = 0;
    } else if (tensIndex === 9) {
      // Quatre-vingt-dix: 80 + 10..19
      const subVal = 10 + units;
      result += "QUATRE-VINGT-" + UNITS[subVal] + " ";
      n = 0;
    } else {
      // Vingt, Trente, etc.
      let tensWord = TENS[tensIndex];
      if (tensIndex === 8 && units > 0) {
        // Quatre-vingts perd le s quand suivi d'unités
        tensWord = "QUATRE-VINGTS".replace(/S$/, "");
      }
      if (units > 0) {
        const liaison = (tensIndex < 8 && units === 1) ? "-ET-" : "-";
        result += tensWord + liaison + UNITS[units] + " ";
      } else {
        result += tensWord + " ";
      }
      n = 0;
    }
  } else if (n > 0) {
    result += UNITS[n] + " ";
    n = 0;
  }

  return result.trim();
}

/**
 * Formate un montant en toutes lettres avec le format ASECNA.
 * Ex: 5_600_000 → "CINQ MILLIONS SIX CENT MILLE (5 600 000) FRANCS.CFA"
 */
export function formatMontantEnLettres(montant: number): string {
  const words = numberToFrenchWords(Math.round(montant));
  const formatted = montant.toLocaleString("fr-FR").replace(/\s/g, " ");
  return `${words} (${formatted}) FRANCS.CFA`;
}

// ---------------------------------------------------------------------------
// Parsing / incrémentation du numéro de facture
// ---------------------------------------------------------------------------

interface ParsedInvoiceNum {
  prefix: string;   // e.g. "N°2026/"
  num: number;      // e.g. 23
  numWidth: number; // e.g. 3 (for zero-padding)
  suffix: string;   // e.g. "/ASECNA/DGAN/CAF"
}

function parseInvoiceNumber(s: string): ParsedInvoiceNum | null {
  const raw = (s ?? "").trim();
  if (!raw) return null;
  // Exclure les placeholders du template (ex: "????")
  if (/\?/.test(raw)) return null;

  // Matches patterns like:
  //   N°2026/023/ASECNA/DGAN/CAF  → prefix="N°2026/", num=23, suffix="/ASECNA/DGAN/CAF"
  //   N°001                        → prefix="N°", num=1, suffix=""
  const match = raw.match(/^(N°[^0-9]*)([0-9]+)(.*)$/i);
  if (match) {
    const prefix = match[1];
    const numStr = match[2];
    const suffix = match[3];
    return {
      prefix,
      num: parseInt(numStr, 10),
      numWidth: numStr.length,
      suffix,
    };
  }

  // Accepter aussi un numéro "simple" (ex: "7" ou "117")
  const onlyDigits = raw.match(/^([0-9]+)$/);
  if (onlyDigits) {
    const numStr = onlyDigits[1];
    return {
      prefix: "N°",
      num: parseInt(numStr, 10),
      // Utiliser 3 chiffres comme format standard si l'entrée est courte
      numWidth: Math.max(3, numStr.length),
      suffix: "",
    };
  }

  return null;
}

function buildInvoiceNumber(parsed: ParsedInvoiceNum, offset: number): string {
  const n = parsed.num + offset;
  return `${parsed.prefix}${String(n).padStart(parsed.numWidth, "0")}${parsed.suffix}`;
}

/**
 * Structure d'un bloc de facture détecté dans le template.
 * Chaque bloc est une "mini-page" de facture avec ses positions relatives.
 */
interface InvoiceTemplateBlock {
  /** Colonne de départ (1-based) : 1 = bloc gauche (A-H), 9 = bloc droit (I-P) */
  startCol: number;
  /** Ligne de début du bloc (1-based, ligne du "Facture" texte au-dessus) */
  startRow: number;
  /** Ligne de fin du bloc (dernière ligne avec contenu, ex: signature) */
  endRow: number;
  /** Positions relatives des champs (offsets par rapport à startRow) */
  fields: {
    factureLabel:     { rowOffset: number; col: number };  // "Facture" (ex: A3)
    factureNumber:    { rowOffset: number; col: number };  // "Facture N°xxx" (ex: A5)
    clientName:       { rowOffset: number; col: number };  // Nom du client (ex: F5)
    clientCity:       { rowOffset: number; col: number };  // Ville du client (ex: F7)
    site:             { rowOffset: number; col: number };  // "Site: xxx" (ex: A10)
    periode:          { rowOffset: number; col: number };  // "Du xx au xx" (ex: A14)
    designation:      { rowOffset: number; col: number };  // Description (ex: B15)
    conventionNum:    { rowOffset: number; col: number };  // N° convention (ex: B16)
    montant:          { rowOffset: number; col: number };  // Montant (ex: H17)
    montantEnLettres: { rowOffset: number; col: number };  // Montant en lettres (ex: A24)
    montantHT:        { rowOffset: number; col: number };  // Montant HT récap (ex: B21)
    taxe:             { rowOffset: number; col: number };  // Taxe récap (ex: D21)
    acompte:          { rowOffset: number; col: number };  // Acompte récap (ex: F21)
    montantTTC:       { rowOffset: number; col: number };  // Montant TTC récap (ex: G21)
    solde:            { rowOffset: number; col: number };  // Solde récap (ex: H21)
  };
  /** Nombre total de lignes du bloc (hauteur) */
  height: number;
}

/**
 * Configuration pour la feuille "Conventions" basée sur l'analyse du template V1.
 * Il y a 2 colonnes de blocs (gauche A-H, droite I-P) et les blocs
 * sont empilés verticalement par paires.
 * 
 * Paire 1: lignes 3 à ~38 (bloc gauche + bloc droit)
 * Paire 2: lignes 50 à ~85 (bloc gauche + bloc droit)
 * 
 * Chaque "paire" fait ~47 lignes (avec espace entre les paires).
 */

// Bloc gauche (colonnes A-H) — basé sur le template analysé (v3 07032026)
const LEFT_BLOCK_TEMPLATE: Omit<InvoiceTemplateBlock, 'startRow'> = {
  startCol: 1,   // Colonne A
  endRow: 0,     // Sera calculé
  height: 36,    // Lignes 3-38 = 36 lignes par bloc
  fields: {
    factureLabel:   { rowOffset: 0,  col: 1 },  // A3  -> "Facture"
    factureNumber:  { rowOffset: 2,  col: 1 },  // A5  -> "Facture N°xxx"
    clientName:     { rowOffset: 2,  col: 6 },  // F5  -> Nom client (même ligne que N° facture)
    clientCity:     { rowOffset: 4,  col: 6 },  // F7  -> Ville client
    site:           { rowOffset: 7,  col: 1 },  // A10 -> "Site: xxx"
    periode:        { rowOffset: 11, col: 1 },  // A14 -> "Du xx au xx"
    designation:    { rowOffset: 12, col: 2 },  // B15 -> Désignation
    conventionNum:  { rowOffset: 13, col: 2 },  // B16 -> N° convention
    montant:        { rowOffset: 14, col: 8 },  // H17 -> Montant principal
    montantEnLettres: { rowOffset: 21, col: 1 }, // A24 -> Montant en lettres
    montantHT:      { rowOffset: 18, col: 2 },  // B21 -> Montant HT
    taxe:           { rowOffset: 18, col: 4 },  // D21 -> Taxe
    acompte:        { rowOffset: 18, col: 6 },  // F21 -> Acompte
    montantTTC:     { rowOffset: 18, col: 7 },  // G21 -> Montant TTC
    solde:          { rowOffset: 18, col: 8 },  // H21 -> Solde
  }
};

// Bloc droit (colonnes I-P) — basé sur le template analysé (v3 07032026)
const RIGHT_BLOCK_TEMPLATE: Omit<InvoiceTemplateBlock, 'startRow'> = {
  startCol: 9,   // Colonne I
  endRow: 0,
  height: 36,
  fields: {
    factureLabel:   { rowOffset: 0,  col: 9  },  // I3  -> "Facture"
    factureNumber:  { rowOffset: 2,  col: 9  },  // I5  -> "Facture N°xxx"
    clientName:     { rowOffset: 2,  col: 14 },  // N5  -> Nom client (même ligne que N° facture)
    clientCity:     { rowOffset: 4,  col: 14 },  // N7  -> Ville client
    site:           { rowOffset: 7,  col: 9  },  // I10 -> "Site: xxx"
    periode:        { rowOffset: 11, col: 9  },  // I14 -> "Du xx au xx"
    designation:    { rowOffset: 12, col: 10 },  // J15 -> Désignation
    conventionNum:  { rowOffset: 13, col: 10 },  // J16 -> N° convention
    montant:        { rowOffset: 14, col: 16 },  // P17 -> Montant principal
    montantEnLettres: { rowOffset: 21, col: 9 }, // I24 -> Montant en lettres
    montantHT:      { rowOffset: 18, col: 10 },  // J21 -> Montant HT
    taxe:           { rowOffset: 18, col: 12 },  // L21 -> Taxe
    acompte:        { rowOffset: 18, col: 14 },  // N21 -> Acompte
    montantTTC:     { rowOffset: 18, col: 15 },  // O21 -> Montant TTC
    solde:          { rowOffset: 18, col: 16 },  // P21 -> Solde
  }
};

/** Espace entre deux "rangées" de blocs (paires haut/bas) */
const ROW_SPACING = 47; // lignes 3->50 = 47 lignes

/** Ligne de départ de la première paire */
const FIRST_PAIR_START_ROW = 3;


/**
 * Sauvegarde le style complet d'une cellule
 */
function preserveCellStyle(cell: ExcelJS.Cell): any {
  return {
    font: cell.font ? { ...cell.font } : undefined,
    alignment: cell.alignment ? { ...cell.alignment } : undefined,
    border: cell.border ? { ...cell.border } : undefined,
    fill: cell.fill ? { ...cell.fill } : undefined,
    numFmt: cell.numFmt,
  };
}

/**
 * Définir la valeur d'une cellule en préservant son style
 */
function setCellValuePreservingStyle(
  cell: ExcelJS.Cell,
  value: string | number
): void {
  const style = preserveCellStyle(cell);
  cell.value = value;

  if (style.font) cell.font = style.font;
  if (style.alignment) cell.alignment = style.alignment;
  if (style.border) cell.border = style.border;
  if (style.fill) cell.fill = style.fill;
  if (style.numFmt) cell.numFmt = style.numFmt;
}

/**
 * Copie le style et les bordures d'une cellule source vers une cellule destination
 */
function copyCellStyle(source: ExcelJS.Cell, dest: ExcelJS.Cell): void {
  if (source.font) dest.font = { ...source.font };
  if (source.alignment) dest.alignment = { ...source.alignment };
  if (source.border) dest.border = { ...source.border };
  if (source.fill) dest.fill = { ...source.fill };
  if (source.numFmt) dest.numFmt = source.numFmt;
}

/**
 * Copie complète d'une cellule (valeur + style) d'une source vers une destination
 * NOTE: Cette fonction n'est plus utilisée pour éviter la duplication des données.
 * Utilisez copyCellStyle à la place pour copier uniquement le style.
 */
function copyCell(source: ExcelJS.Cell, dest: ExcelJS.Cell): void {
  // Copier la valeur
  dest.value = source.value;
  // Copier le style
  copyCellStyle(source, dest);
}

/**
 * Liste des positions de cellules contenant des données VARIABLES (à ne pas copier).
 * Ces positions sont relatives au début du bloc (offset par rapport à startRow).
 * Format: { rowOffset, cols: [colonnes concernées] }
 */
const VARIABLE_DATA_POSITIONS_LEFT = [
  { rowOffset: 2,  cols: [1] },             // Facture N°xxx (A5)
  { rowOffset: 2,  cols: [6, 7, 8] },       // Nom client (F5-H5) — même ligne que N° facture
  { rowOffset: 4,  cols: [6, 7, 8] },       // Ville client (F7-H7)
  { rowOffset: 7,  cols: [1] },             // Site (A10)
  { rowOffset: 11, cols: [1] },             // Période (A14)
  { rowOffset: 12, cols: [2, 3, 4, 5, 6, 7] }, // Désignation (B15-G15)
  { rowOffset: 13, cols: [2, 3, 4, 5, 6, 7] }, // N° Convention (B16-G16)
  { rowOffset: 14, cols: [8] },             // Montant (H17)
  { rowOffset: 18, cols: [2, 3, 4, 6, 7, 8] }, // Récap: HT, Taxe, Acompte, TTC, Solde (ligne 21)
  { rowOffset: 21, cols: [1, 2, 3, 4, 5, 6, 7, 8] }, // Montant en lettres (A24)
];

const VARIABLE_DATA_POSITIONS_RIGHT = [
  { rowOffset: 2,  cols: [9] },              // Facture N°xxx (I5)
  { rowOffset: 2,  cols: [14, 15, 16] },     // Nom client (N5-P5) — même ligne que N° facture
  { rowOffset: 4,  cols: [14, 15, 16] },     // Ville client (N7-P7)
  { rowOffset: 7,  cols: [9] },              // Site (I10)
  { rowOffset: 11, cols: [9] },              // Période (I14)
  { rowOffset: 12, cols: [10, 11, 12, 13, 14, 15] }, // Désignation (J15-O15)
  { rowOffset: 13, cols: [10, 11, 12, 13, 14, 15] }, // N° Convention (J16-O16)
  { rowOffset: 14, cols: [16] },             // Montant (P17)
  { rowOffset: 18, cols: [10, 11, 12, 14, 15, 16] }, // Récap: HT, Taxe, Acompte, TTC, Solde
  { rowOffset: 21, cols: [9, 10, 11, 12, 13, 14, 15, 16] }, // Montant en lettres (I24)
];

/**
 * Vérifie si une cellule contient des données variables (à ne pas copier)
 */
function isVariableDataCell(
  rowOffset: number,
  col: number,
  isLeftBlock: boolean
): boolean {
  const positions = isLeftBlock ? VARIABLE_DATA_POSITIONS_LEFT : VARIABLE_DATA_POSITIONS_RIGHT;
  
  for (const pos of positions) {
    if (pos.rowOffset === rowOffset && pos.cols.includes(col)) {
      return true;
    }
  }
  return false;
}

/**
 * Récupère les informations sur les merges d'une feuille
 * Retourne un Set de clés "row,col" pour les cellules qui font partie d'un merge
 * mais ne sont PAS la cellule en haut à gauche (master)
 */
function getMergedSlaveCells(sheet: ExcelJS.Worksheet): Set<string> {
  const slaveCells = new Set<string>();
  const merges = (sheet as any)._merges || {};
  
  for (const mergeKey of Object.keys(merges)) {
    const mergeRef = merges[mergeKey];
    if (mergeRef && mergeRef.model) {
      const { top, left, bottom, right } = mergeRef.model;
      // Toutes les cellules sauf la première (top-left) sont des "slaves"
      for (let r = top; r <= bottom; r++) {
        for (let c = left; c <= right; c++) {
          if (r !== top || c !== left) {
            slaveCells.add(`${r},${c}`);
          }
        }
      }
    }
  }
  
  return slaveCells;
}

/**
 * Si (row,col) appartient à une cellule fusionnée, retourner la cellule "maître"
 * (top-left) du merge pour que la valeur soit bien prise en compte par ExcelJS.
 */
function resolveMergedCellToMaster(
  sheet: ExcelJS.Worksheet,
  row: number,
  col: number
): { row: number; col: number } {
  const merges = (sheet as any)._merges || {};
  for (const mergeKey of Object.keys(merges)) {
    const mergeRef = merges[mergeKey];
    if (!mergeRef?.model) continue;
    const { top, left, bottom, right } = mergeRef.model;
    const inside =
      row >= top && row <= bottom && col >= left && col <= right;
    if (!inside) continue;
    return { row: top, col: left };
  }
  return { row, col };
}

/**
 * Copie un "bloc" du template (une facture complète) depuis une position source
 * vers une position destination dans la même feuille.
 * IMPORTANT: Cette fonction copie la structure, le style ET les valeurs statiques (labels),
 * mais EXCLUT les données variables (nom client, montants, etc.)
 * Pour les cellules fusionnées, seule la cellule maître reçoit la valeur.
 * @param sheet - La feuille Excel
 * @param srcStartRow - Ligne de départ du bloc source (1-based)
 * @param destStartRow - Ligne de départ de la destination (1-based)
 * @param startCol - Colonne de départ (1-based)
 * @param endCol - Colonne de fin (1-based)
 * @param height - Nombre de lignes à copier
 */
function copyBlockRegion(
  sheet: ExcelJS.Worksheet,
  srcStartRow: number,
  destStartRow: number,
  startCol: number,
  endCol: number,
  height: number
): void {
  // Récupérer les cellules "esclaves" des merges (celles qui ne doivent pas avoir de valeur)
  const slaveCells = getMergedSlaveCells(sheet);
  
  for (let rowOffset = 0; rowOffset < height; rowOffset++) {
    const srcRow = srcStartRow + rowOffset;
    const destRow = destStartRow + rowOffset;

    // Copier la hauteur de ligne
    const srcRowObj = sheet.getRow(srcRow);
    const destRowObj = sheet.getRow(destRow);
    if (srcRowObj.height) {
      destRowObj.height = srcRowObj.height;
    }

    for (let col = startCol; col <= endCol; col++) {
      const srcCell = sheet.getCell(srcRow, col);
      const destCell = sheet.getCell(destRow, col);
      
      // Toujours copier le style
      copyCellStyle(srcCell, destCell);
      
      // Déterminer si c'est le bloc gauche ou droit
      const isLeftBlock = col <= 8;
      
      // Vérifier si la cellule SOURCE fait partie d'un merge mais n'est pas la cellule maître
      const isSrcSlave = slaveCells.has(`${srcRow},${col}`);
      
      // Vérifier si c'est une cellule de données variables
      if (isVariableDataCell(rowOffset, col, isLeftBlock)) {
        // Ne pas copier la valeur (données variables)
        destCell.value = null;
      } else if (isSrcSlave) {
        // Cellule esclave d'un merge - ne pas copier la valeur (elle sera fusionnée)
        destCell.value = null;
      } else {
        // Copier la valeur (labels statiques, formules, etc.)
        destCell.value = srcCell.value;
      }
    }
  }
}

/**
 * Copie les merges (cellules fusionnées) d'un bloc source vers un bloc destination
 */
function copyMergesForBlock(
  sheet: ExcelJS.Worksheet,
  srcStartRow: number,
  destStartRow: number,
  startCol: number,
  endCol: number,
  height: number
): void {
  // ExcelJS stocke les merges avec un objet model contenant top, left, bottom, right
  const merges = (sheet as any)._merges || {};
  const newMerges: string[] = [];

  console.log(`copyMergesForBlock: src=${srcStartRow}, dest=${destStartRow}, cols=${startCol}-${endCol}, height=${height}`);
  console.log(`Nombre de merges trouvés: ${Object.keys(merges).length}`);

  for (const mergeKey of Object.keys(merges)) {
    const mergeRef = merges[mergeKey];
    
    try {
      let top: number, left: number, bottom: number, right: number;
      
      // ExcelJS stocke les merges avec un objet model
      if (mergeRef && mergeRef.model) {
        top = mergeRef.model.top;
        left = mergeRef.model.left;
        bottom = mergeRef.model.bottom;
        right = mergeRef.model.right;
      } else {
        // Fallback: essayer de parser la clé comme "A1:B2"
        const parts = mergeKey.split(':');
        if (parts.length !== 2) continue;
        
        const topLeftAddr = parseCellAddress(parts[0]);
        const bottomRightAddr = parseCellAddress(parts[1]);
        
        if (!topLeftAddr || !bottomRightAddr) continue;
        
        top = topLeftAddr.row;
        left = topLeftAddr.col;
        bottom = bottomRightAddr.row;
        right = bottomRightAddr.col;
      }

      // Vérifier si ce merge est dans la zone source
      if (top >= srcStartRow &&
          bottom < srcStartRow + height &&
          left >= startCol &&
          right <= endCol) {
        
        // Calculer les nouvelles positions (décalage)
        const newTop = top - srcStartRow + destStartRow;
        const newBottom = bottom - srcStartRow + destStartRow;

        const newTopLeft = encodeCellAddress(newTop, left);
        const newBottomRight = encodeCellAddress(newBottom, right);
        const newMergeRange = `${newTopLeft}:${newBottomRight}`;

        console.log(`  Merge trouvé: ${mergeKey} (${top},${left}):(${bottom},${right}) -> ${newMergeRange}`);
        newMerges.push(newMergeRange);
      }
    } catch (e) {
      console.error(`Erreur parsing merge ${mergeKey}:`, e);
    }
  }

  console.log(`Nouveaux merges à créer: ${newMerges.length}`);

  // Appliquer les nouveaux merges
  for (const merge of newMerges) {
    try {
      sheet.mergeCells(merge);
      console.log(`  Merge créé: ${merge}`);
    } catch (e) {
      console.error(`  Erreur création merge ${merge}:`, e);
    }
  }
}

/**
 * Parse une adresse de cellule comme "A1" en {row, col} (1-based)
 */
function parseCellAddress(addr: string): { row: number; col: number } | null {
  const match = addr.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  const colLetters = match[1];
  const row = parseInt(match[2]);

  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }

  return { row, col };
}

/**
 * Encode une position {row, col} (1-based) en adresse de cellule "A1"
 */
function encodeCellAddress(row: number, col: number): string {
  let colStr = '';
  let c = col;
  while (c > 0) {
    const remainder = (c - 1) % 26;
    colStr = String.fromCharCode(65 + remainder) + colStr;
    c = Math.floor((c - 1) / 26);
  }
  return `${colStr}${row}`;
}

/**
 * Remplit un bloc de facture "Convention" avec les données d'une convention
 */
function fillConventionBlock(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  blockTemplate: Omit<InvoiceTemplateBlock, 'startRow'>,
  convention: ConventionData,
  invoiceNumberStr: string
): void {
  const fields = blockTemplate.fields;

  try {
    const safeInvoiceNumberStr = /\?/.test(invoiceNumberStr) ? 'N°001' : invoiceNumberStr;

    // Numéro de facture
    const rawNumRow = startRow + fields.factureNumber.rowOffset;
    const rawNumCol = fields.factureNumber.col;
    const masterPos = resolveMergedCellToMaster(sheet, rawNumRow, rawNumCol);
    const numCell = sheet.getCell(masterPos.row, masterPos.col);
    setCellValuePreservingStyle(numCell, `Facture ${safeInvoiceNumberStr}`);

    // Nom du client
    const clientCell = sheet.getCell(startRow + fields.clientName.rowOffset, fields.clientName.col);
    setCellValuePreservingStyle(clientCell, convention["NOM DU CLIENT"]);

    // Ville du client — on utilise le site brut comme ville
    const cityCell = sheet.getCell(startRow + fields.clientCity.rowOffset, fields.clientCity.col);
    setCellValuePreservingStyle(cityCell, convention.SITE);

    // Site
    const siteCell = sheet.getCell(startRow + fields.site.rowOffset, fields.site.col);
    setCellValuePreservingStyle(siteCell, `Site: ${convention.SITE}`);

    // Période
    const periodeCell = sheet.getCell(startRow + fields.periode.rowOffset, fields.periode.col);
    setCellValuePreservingStyle(periodeCell, `Du ${convention["Date de debut"]} au ${convention["Date de fin"]}`);

    // Désignation (objet de la convention)
    const designationCell = sheet.getCell(startRow + fields.designation.rowOffset, fields.designation.col);
    setCellValuePreservingStyle(designationCell, convention["OBJET DE LA CONVENTION"]);

    // N° Convention
    const convCell = sheet.getCell(startRow + fields.conventionNum.rowOffset, fields.conventionNum.col);
    setCellValuePreservingStyle(convCell, convention["N° CONVENTION"]);

    // Montant principal
    const montantValue = typeof convention.MONTANT === 'number'
      ? convention.MONTANT
      : parseFloat(String(convention.MONTANT).replace(/[^\d.-]/g, ''));

    if (!isNaN(montantValue)) {
      const montantCell = sheet.getCell(startRow + fields.montant.rowOffset, fields.montant.col);
      const montantStyle = preserveCellStyle(montantCell);
      montantCell.value = montantValue;
      if (montantStyle.font) montantCell.font = montantStyle.font;
      if (montantStyle.numFmt) montantCell.numFmt = montantStyle.numFmt;

      // Montant en lettres
      const lettresCell = sheet.getCell(startRow + fields.montantEnLettres.rowOffset, fields.montantEnLettres.col);
      setCellValuePreservingStyle(lettresCell, formatMontantEnLettres(montantValue));

      // Remplir aussi la ligne récapitulatif
      const htCell = sheet.getCell(startRow + fields.montantHT.rowOffset, fields.montantHT.col);
      const htStyle = preserveCellStyle(htCell);
      htCell.value = montantValue;
      if (htStyle.font) htCell.font = htStyle.font;
      if (htStyle.numFmt) htCell.numFmt = htStyle.numFmt;

      const taxeCell = sheet.getCell(startRow + fields.taxe.rowOffset, fields.taxe.col);
      const taxeStyle = preserveCellStyle(taxeCell);
      taxeCell.value = 0;
      if (taxeStyle.font) taxeCell.font = taxeStyle.font;

      const acompteCell = sheet.getCell(startRow + fields.acompte.rowOffset, fields.acompte.col);
      const acompteStyle = preserveCellStyle(acompteCell);
      acompteCell.value = 0;
      if (acompteStyle.font) acompteCell.font = acompteStyle.font;

      const ttcCell = sheet.getCell(startRow + fields.montantTTC.rowOffset, fields.montantTTC.col);
      const ttcStyle = preserveCellStyle(ttcCell);
      ttcCell.value = montantValue;
      if (ttcStyle.font) ttcCell.font = ttcStyle.font;
      if (ttcStyle.numFmt) ttcCell.numFmt = ttcStyle.numFmt;

      const soldeCell = sheet.getCell(startRow + fields.solde.rowOffset, fields.solde.col);
      const soldeStyle = preserveCellStyle(soldeCell);
      soldeCell.value = montantValue;
      if (soldeStyle.font) soldeCell.font = soldeStyle.font;
      if (soldeStyle.numFmt) soldeCell.numFmt = soldeStyle.numFmt;
    }
  } catch (error) {
    console.error("Erreur lors du remplissage du bloc convention:", error);
  }
}

/**
 * Efface les données variables d'un bloc (remet les champs à vide)
 * mais conserve les labels statiques et le style
 */
function clearBlockData(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  blockTemplate: Omit<InvoiceTemplateBlock, 'startRow'>
): void {
  const fields = blockTemplate.fields;

  const fieldsToClear = [
    fields.factureNumber,
    fields.clientName,
    fields.clientCity,
    fields.site,
    fields.periode,
    fields.designation,
    fields.conventionNum,
    fields.montant,
    fields.montantEnLettres,
    fields.montantHT,
    fields.taxe,
    fields.acompte,
    fields.montantTTC,
    fields.solde,
  ];

  for (const field of fieldsToClear) {
    const cell = sheet.getCell(startRow + field.rowOffset, field.col);
    const style = preserveCellStyle(cell);
    cell.value = null;
    if (style.font) cell.font = style.font;
    if (style.alignment) cell.alignment = style.alignment;
    if (style.border) cell.border = style.border;
    if (style.fill) cell.fill = style.fill;
    if (style.numFmt) cell.numFmt = style.numFmt;
  }
}

/**
 * Crée une nouvelle paire de blocs (gauche + droite) en copiant la structure
 * d'une paire existante dans le template.
 * @param sheet - La feuille Excel
 * @param sourcePairStartRow - Ligne de début de la paire source
 * @param destPairStartRow - Ligne de début de la nouvelle paire
 * @param pairHeight - Hauteur d'une paire complète
 */
function duplicatePair(
  sheet: ExcelJS.Worksheet,
  sourcePairStartRow: number,
  destPairStartRow: number,
  pairHeight: number
): void {
  // Copier toute la région de la paire (colonnes A à P, 1 à 16)
  copyBlockRegion(sheet, sourcePairStartRow, destPairStartRow, 1, 16, pairHeight);

  // Copier les merges
  copyMergesForBlock(sheet, sourcePairStartRow, destPairStartRow, 1, 16, pairHeight);
}

/**
 * Normalise le nom d'un site pour le regroupement
 */
function normalizeSiteName(site: string): string {
  if (!site || site.trim() === '') return 'AUTRES';
  
  const normalized = site.trim().toUpperCase()
    .replace(/^SITE:\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  
  // Ignorer les lignes de titre/total qui ne sont pas des sites valides
  if (normalized.includes('TABLEAU') || 
      normalized.includes('TOTAL') || 
      normalized.includes('CONVENTION') ||
      normalized.includes('FACTURER') ||
      normalized.includes('SUITE') ||
      normalized.length > 20) {
    return 'IGNORER'; // Marqueur spécial pour filtrer ces lignes
  }
  
  // Mapper les variantes courantes (codes OACI et noms complets)
  const siteMapping: Record<string, string> = {
    // Port-Gentil
    'POG': 'PORT-GENTIL',
    'PORT GENTIL': 'PORT-GENTIL',
    'PORT-GENTIL': 'PORT-GENTIL',
    'FOOG': 'PORT-GENTIL',
    
    // Libreville
    'LBV': 'LIBREVILLE',
    'LIBREVILLE': 'LIBREVILLE',
    'LEON MBA': 'LIBREVILLE',
    
    // Mvengue / Franceville
    'MVG': 'MVENGUE',
    'MVENGUE': 'MVENGUE',
    "M'VENGUE": 'MVENGUE',
    'MVENGUÉ': 'MVENGUE',
    'FCV': 'FRANCEVILLE',
    'FRANCEVILLE': 'FRANCEVILLE',
    'FOON': 'MVENGUE',
    
    // Oyem
    'OYE': 'OYEM',
    'OYM': 'OYEM',
    'OYEM': 'OYEM',
    'FOOY': 'OYEM',
    
    // Makokou
    'MKU': 'MAKOKOU',
    'MAKOKOU': 'MAKOKOU',
    'FOOK': 'MAKOKOU',
    
    // Lambaréné
    'LBQ': 'LAMBARENE',
    'LAMBARENE': 'LAMBARENE',
    'LAMBARÉNÉ': 'LAMBARENE',
    'LAMBARENÉ': 'LAMBARENE',
    'L/RENE': 'LAMBARENE',
    'L/RÉNÉ': 'LAMBARENE',
    'L/RENÉ': 'LAMBARENE',
    'L / RENE': 'LAMBARENE',
    'L / RÉNÉ': 'LAMBARENE',
    'LRENE': 'LAMBARENE',
    'LRÉNÉ': 'LAMBARENE',
    'FOOL': 'LAMBARENE', // Code OACI
    
    // Tchibanga
    'TCH': 'TCHIBANGA',
    'TCHIBANGA': 'TCHIBANGA',
    'FOOT': 'TCHIBANGA',
    
    // Bitam
    'BTM': 'BITAM',
    'BMM': 'BITAM',
    'BITAM': 'BITAM',
    'FOOB': 'BITAM',
    
    // Moanda
    'MFF': 'MOANDA',
    'MOANDA': 'MOANDA',
    'FOOM': 'MOANDA',
    
    // Mouila
    'MJL': 'MOUILA',
    'MOUILA': 'MOUILA',
    
    // Koulamoutou
    'KOU': 'KOULAMOUTOU',
    'KOULAMOUTOU': 'KOULAMOUTOU',
  };
  
  return siteMapping[normalized] || normalized || 'AUTRES';
}

/**
 * Regroupe les conventions par site
 * Filtre automatiquement les lignes de titre/total (marquées 'IGNORER')
 */
function groupConventionsBySite(conventions: ConventionData[]): Map<string, ConventionData[]> {
  const groups = new Map<string, ConventionData[]>();
  
  for (const convention of conventions) {
    const siteName = normalizeSiteName(convention.SITE);
    
    // Ignorer les lignes de titre/total
    if (siteName === 'IGNORER') {
      console.log(`Convention ignorée (ligne de titre): "${convention["NOM DU CLIENT"]}"`);
      continue;
    }
    
    if (!groups.has(siteName)) {
      groups.set(siteName, []);
    }
    groups.get(siteName)!.push(convention);
  }
  
  return groups;
}

/**
 * Génère les factures pour un groupe de conventions sur une feuille donnée
 */
async function generateInvoicesOnSheet(
  sheet: ExcelJS.Worksheet,
  conventions: ConventionData[],
  startInvoiceNumber: number,
  parsedFirstNum: ParsedInvoiceNum | null
): Promise<number> {
  const nbConventions = conventions.length;

  if (nbConventions === 0) {
    return startInvoiceNumber;
  }

  // Configuration : 2 blocs par paire (gauche + droite)
  const BLOCKS_PER_PAIR = 2;

  // Nombre de paires dans le template original (analysé: 2 paires = 4 blocs)
  const TEMPLATE_PAIR_COUNT = 2;

  // Nombre de paires nécessaires pour toutes les conventions
  const neededPairs = Math.ceil(nbConventions / BLOCKS_PER_PAIR);

  // Si on a besoin de plus de paires que le template en contient, dupliquer des paires
  if (neededPairs > TEMPLATE_PAIR_COUNT) {
    const extraPairs = neededPairs - TEMPLATE_PAIR_COUNT;
    const sourcePairStartRow = FIRST_PAIR_START_ROW;
    const pairHeight = ROW_SPACING;

    for (let i = 0; i < extraPairs; i++) {
      const newPairIndex = TEMPLATE_PAIR_COUNT + i;
      const newPairStartRow = FIRST_PAIR_START_ROW + (newPairIndex * ROW_SPACING);
      duplicatePair(sheet, sourcePairStartRow, newPairStartRow, pairHeight);
    }
  }

  // Remplir chaque bloc avec les données des conventions
  let invoiceNumber = startInvoiceNumber;

  for (let i = 0; i < nbConventions; i++) {
    const convention = conventions[i];
    const pairIndex = Math.floor(i / BLOCKS_PER_PAIR);
    const posInPair = i % BLOCKS_PER_PAIR;
    const pairStartRow = FIRST_PAIR_START_ROW + (pairIndex * ROW_SPACING);
    const blockTemplate = posInPair === 0 ? LEFT_BLOCK_TEMPLATE : RIGHT_BLOCK_TEMPLATE;

    // Construire le numéro de facture en string
    const invoiceNumStr = parsedFirstNum
      ? buildInvoiceNumber(parsedFirstNum, invoiceNumber - 1)
      : `N°${String(invoiceNumber).padStart(3, "0")}`;

    fillConventionBlock(sheet, pairStartRow, blockTemplate, convention, invoiceNumStr);
    invoiceNumber++;
  }

  // Effacer les blocs excédentaires dans le template (si conventions < blocs template)
  const TEMPLATE_BLOCK_COUNT = TEMPLATE_PAIR_COUNT * BLOCKS_PER_PAIR;
  if (nbConventions < TEMPLATE_BLOCK_COUNT) {
    for (let i = nbConventions; i < TEMPLATE_BLOCK_COUNT; i++) {
      const pairIndex = Math.floor(i / BLOCKS_PER_PAIR);
      const posInPair = i % BLOCKS_PER_PAIR;
      const pairStartRow = FIRST_PAIR_START_ROW + (pairIndex * ROW_SPACING);
      const blockTemplate = posInPair === 0 ? LEFT_BLOCK_TEMPLATE : RIGHT_BLOCK_TEMPLATE;
      clearBlockData(sheet, pairStartRow, blockTemplate);
    }
  }

  return invoiceNumber;
}

/**
 * Génère un fichier Excel avec toutes les factures de conventions remplies.
 * Les factures sont classées par site/localisation sur des feuilles séparées.
 * 
 * La logique:
 * 1. Charge le template Excel
 * 2. Regroupe les conventions par site
 * 3. Pour chaque site, crée une copie de la feuille "Conventions" nommée avec le site
 * 4. Remplit les factures sur chaque feuille
 * 5. Les conventions sans site sont placées dans une feuille "AUTRES"
 */
/**
 * Pré-clone la feuille "Conventions" pour chaque site supplémentaire directement
 * dans le ZIP du template. Cela préserve 100% de la mise en forme, des bordures,
 * des fusions et des en-têtes/logos (VML).
 */
async function precloneConventionsSheet(
  templateBuffer: ArrayBuffer,
  siteNames: string[],
  JSZipClass: any
): Promise<ArrayBuffer> {
  const zip = await JSZipClass.loadAsync(templateBuffer);

  const wbXml: string = await zip.file('xl/workbook.xml')!.async('string');
  const wbRelsXml: string = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');

  // Trouver la feuille "Conventions" dans workbook.xml.rels
  const sheetMatch = wbXml.match(/<sheet[^>]+name="Conventions"[^>]*>/i);
  if (!sheetMatch) return templateBuffer;

  const ridMatch = sheetMatch[0].match(/r:id="([^"]+)"/i);
  if (!ridMatch) return templateBuffer;
  const convRid = ridMatch[1];

  const relMatch = wbRelsXml.match(new RegExp(`<Relationship[^>]+Id="${convRid}"[^>]+>`, 'i'));
  if (!relMatch) return templateBuffer;
  const targetMatch = relMatch[0].match(/Target="([^"]+)"/i);
  if (!targetMatch) return templateBuffer;

  const convTarget = targetMatch[1]; // ex: "worksheets/sheet2.xml"
  const convFilename = `xl/${convTarget}`;
  const convBaseName = convTarget.split('/').pop()!; // "sheet2.xml"
  const convRelsFilename = `xl/worksheets/_rels/${convBaseName}.rels`;

  const convXml: string = await zip.file(convFilename)!.async('string');
  const convRelsXml: string | null = zip.file(convRelsFilename)
    ? await zip.file(convRelsFilename)!.async('string')
    : null;

  // Trouver le numéro max de sheet existant
  const sheetNums = Object.keys(zip.files)
    .map((f) => f.match(/^xl\/worksheets\/sheet(\d+)\.xml$/))
    .filter(Boolean)
    .map((m) => parseInt(m![1]));
  let maxSheetNum = sheetNums.length > 0 ? Math.max(...sheetNums) : 4;

  // Trouver le rId max dans workbook.xml.rels
  const rIdNums = [...wbRelsXml.matchAll(/Id="rId(\d+)"/gi)].map((m) => parseInt(m[1]));
  let maxRid = rIdNums.length > 0 ? Math.max(...rIdNums) : 8;

  // Trouver le sheetId max dans workbook.xml
  const sheetIdNums = [...wbXml.matchAll(/sheetId="(\d+)"/gi)].map((m) => parseInt(m[1]));
  let maxSheetId = sheetIdNums.length > 0 ? Math.max(...sheetIdNums) : 4;

  let newWbXml = wbXml;
  let newWbRelsXml = wbRelsXml;
  let ctXml: string = await zip.file('[Content_Types].xml')!.async('string');

  // Renommer "Conventions" → premier site dans workbook.xml
  newWbXml = newWbXml.replace(
    /name="Conventions"/i,
    `name="${siteNames[0].replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`
  );

  // Dupliquer la feuille pour chaque site supplémentaire
  for (let i = 1; i < siteNames.length; i++) {
    const newSheetNum = ++maxSheetNum;
    const newRid = `rId${++maxRid}`;
    const newSheetId = ++maxSheetId;
    const newSheetFilename = `xl/worksheets/sheet${newSheetNum}.xml`;
    const newSheetRelsFilename = `xl/worksheets/_rels/sheet${newSheetNum}.xml.rels`;
    const newTarget = `worksheets/sheet${newSheetNum}.xml`;
    const escapedName = siteNames[i].replace(/&/g, '&amp;').replace(/"/g, '&quot;');

    // Ajouter l'entrée de feuille dans workbook.xml (avant </sheets>)
    newWbXml = newWbXml.replace(
      '</sheets>',
      `<sheet name="${escapedName}" sheetId="${newSheetId}" r:id="${newRid}"/></sheets>`
    );

    // Ajouter la relation dans workbook.xml.rels
    newWbRelsXml = newWbRelsXml.replace(
      '</Relationships>',
      `<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newTarget}"/></Relationships>`
    );

    // Cloner le XML de la feuille Conventions
    zip.file(newSheetFilename, convXml);
    if (convRelsXml) {
      zip.file(newSheetRelsFilename, convRelsXml);
    }

    // Ajouter l'entrée dans Content_Types.xml si besoin
    const partName = `/xl/worksheets/sheet${newSheetNum}.xml`;
    if (!ctXml.includes(partName)) {
      ctXml = ctXml.replace(
        '</Types>',
        `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
      );
    }
  }

  // Supprimer calcChain.xml qui référence des formules/feuilles potentiellement supprimées
  if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml');
  ctXml = ctXml.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
  // Retirer la relation calcChain dans workbook.xml.rels
  newWbRelsXml = newWbRelsXml.replace(
    /<Relationship[^>]*Type="[^"]*calcChain[^"]*"[^>]*\/>/gi, ''
  );

  zip.file('xl/workbook.xml', newWbXml);
  zip.file('xl/_rels/workbook.xml.rels', newWbRelsXml);
  zip.file('[Content_Types].xml', ctXml);

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/**
 * Génère UN SEUL fichier Excel avec UNE SEULE facture sur UNE SEULE page.
 * - Seul le bloc gauche (colonnes A-H) est utilisé
 * - Les colonnes I-P sont masquées
 * - La 2e paire de blocs (lignes 50+) est supprimée
 */
export async function generateSingleInvoiceFile(
  convention: ConventionData,
  templateSource: string | ArrayBuffer,
  invoiceNumber?: string
): Promise<ArrayBuffer> {
  // Charger le template
  let arrayBuffer: ArrayBuffer;
  if (typeof templateSource === 'string') {
    const response = await fetch(encodeURI(templateSource));
    if (!response.ok) throw new Error(`Impossible de charger le template: ${response.status}`);
    const ct = response.headers.get('content-type') || '';
    if (ct.includes('text/html')) throw new Error('Le template retourné est du HTML, pas un fichier Excel. Vérifiez le chemin du template.');
    arrayBuffer = await response.arrayBuffer();
  } else {
    arrayBuffer = templateSource;
  }

  const JSZip = (await import("jszip")).default;

  // Phase 1 : JSZip — renommer la feuille Conventions avec le nom du site
  const siteName = normalizeSiteName(convention.SITE) || convention.SITE || "FACTURE";
  const preparedBuffer = await precloneConventionsSheet(arrayBuffer, [siteName], JSZip);

  // Phase 2 : ExcelJS — remplir le bloc gauche et masquer le côté droit
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(preparedBuffer);

  const sheet = workbook.getWorksheet(siteName);
  if (!sheet) throw new Error(`Feuille "${siteName}" introuvable`);

  // Parsing du numéro de facture
  const parsedNum: ParsedInvoiceNum | null = invoiceNumber
    ? parseInvoiceNumber(invoiceNumber)
    : null;
  const invoiceNumStr = parsedNum
    ? buildInvoiceNumber(parsedNum, 0)
    : `N°001`;

  // Remplir le bloc gauche (premier bloc)
  fillConventionBlock(sheet, FIRST_PAIR_START_ROW, LEFT_BLOCK_TEMPLATE, convention, invoiceNumStr);

  // Vider le bloc droit (ne pas afficher de facture vide à côté)
  clearBlockData(sheet, FIRST_PAIR_START_ROW, RIGHT_BLOCK_TEMPLATE);

  // Masquer les colonnes I-P (9-16) — bloc droit
  for (let col = 9; col <= 16; col++) {
    const column = sheet.getColumn(col);
    column.hidden = true;
  }

  // Supprimer le contenu des lignes de la 2e paire (à partir de la ligne 39)
  // pour n'avoir qu'une seule facture visible
  const lastContentRow = FIRST_PAIR_START_ROW + LEFT_BLOCK_TEMPLATE.height + 10;
  for (let r = lastContentRow; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      cell.value = null;
    });
    row.hidden = true;
  }

  // Écrire le buffer
  const excelJsRaw = await workbook.xlsx.writeBuffer({ useStyles: true });
  // Normaliser en ArrayBuffer pour compatibilité navigateur avec JSZip
  const excelJsBuffer = (excelJsRaw instanceof ArrayBuffer
    ? excelJsRaw
    : (excelJsRaw as Uint8Array).buffer.slice(
        (excelJsRaw as Uint8Array).byteOffset,
        (excelJsRaw as Uint8Array).byteOffset + (excelJsRaw as Uint8Array).byteLength
      )) as ArrayBuffer;

  // Phase 3 : JSZip — restaurer VML/en-têtes/marges dans le bon ordre OOXML
  try {
    const origZip = await JSZip.loadAsync(arrayBuffer.slice(0));
    const genZip  = await JSZip.loadAsync(excelJsBuffer);

    // Supprimer calcChain.xml si présent (évite l'erreur de réparation Excel)
    if (genZip.file('xl/calcChain.xml')) genZip.remove('xl/calcChain.xml');
    const genCtFile = genZip.file('[Content_Types].xml');
    if (genCtFile) {
      let genCtXml = await genCtFile.async('string');
      genCtXml = genCtXml.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
      genZip.file('[Content_Types].xml', genCtXml);
    }

    // Copier médias + printerSettings
    for (const filename of Object.keys(origZip.files)) {
      if (
        (filename.startsWith("xl/media/") || filename.startsWith("xl/printerSettings/")) &&
        origZip.file(filename) && !genZip.file(filename)
      ) {
        genZip.file(filename, await origZip.file(filename)!.async("arraybuffer"));
      }
    }

    // Récupérer la feuille Conventions originale
    const origWbXml     = await origZip.file("xl/workbook.xml")!.async("string");
    const origWbRelsXml = await origZip.file("xl/_rels/workbook.xml.rels")!.async("string");
    const convSheetRow  = origWbXml.match(/<sheet[^>]+name="Conventions"[^>]*>/i);
    const convRidM      = convSheetRow ? convSheetRow[0].match(/r:id="([^"]+)"/i) : null;
    const convRid       = convRidM ? convRidM[1] : null;
    const convRelM      = convRid
      ? origWbRelsXml.match(new RegExp(`<Relationship[^>]+Id="${convRid}"[^>]+>`, "i"))
      : null;
    const convTargetM   = convRelM ? convRelM[0].match(/Target="([^"]+)"/i) : null;
    const convSheetFile = convTargetM ? `xl/${convTargetM[1].replace(/^\//, "")}` : null;

    if (!convSheetFile || !origZip.file(convSheetFile)) {
      return await genZip.generateAsync({ type: "arraybuffer" });
    }

    const convSheetXml = await origZip.file(convSheetFile)!.async("string");
    const convRelsFile = `xl/worksheets/_rels/${convSheetFile.split("/").pop()}.rels`;
    const convRelsXml  = origZip.file(convRelsFile)
      ? await origZip.file(convRelsFile)!.async("string")
      : null;

    // Tags à restaurer
    const tPageMargins  = convSheetXml.match(/<pageMargins\b[^>]*\/>/i)?.[0] || null;
    const tHeaderFooter = (() => {
      const m = convSheetXml.match(/<headerFooter\b[^>]*>[\s\S]*?<\/headerFooter>/i);
      return m ? m[0] : null;
    })();

    // r:id attendus par le template (pour aligner header/logo VML)
    const templateLegacyDrawingRid =
      convSheetXml.match(/<legacyDrawingHF\b[^>]*\br:id="(rId\d+)"[^>]*\/>/i)?.[1] || null;
    const templatePageSetupRid =
      convSheetXml.match(/<pageSetup\b[^>]*\br:id="(rId\d+)"[^>]*\/?>/i)?.[1] || null;

    // VML
    const vmlRelM    = convRelsXml
      ? convRelsXml.match(/<Relationship[^>]+Type="[^"]*vmlDrawing[^"]*"[^>]+>/i)
      : null;
    const vmlTarget  = vmlRelM ? (vmlRelM[0].match(/Target="([^"]+)"/i) || [])[1] : null;
    const vmlBaseName = vmlTarget ? vmlTarget.replace(/^.*\//, "") : null;
    const vmlContent  = vmlBaseName && origZip.file(`xl/drawings/${vmlBaseName}`)
      ? await origZip.file(`xl/drawings/${vmlBaseName}`)!.async("string")
      : null;
    const vmlRelsContent = vmlBaseName && origZip.file(`xl/drawings/_rels/${vmlBaseName}.rels`)
      ? await origZip.file(`xl/drawings/_rels/${vmlBaseName}.rels`)!.async("string")
      : null;

    const printerRelM  = convRelsXml
      ? convRelsXml.match(/<Relationship[^>]+Type="[^"]*printerSettings[^"]*"[^>]+>/i)
      : null;
    const printerTarget = printerRelM
      ? (printerRelM[0].match(/Target="([^"]+)"/i) || [])[1]
      : null;

    const existingVmlNums = Object.keys(origZip.files)
      .map(f => { const m = f.match(/xl\/drawings\/vmlDrawing(\d+)\.vml$/); return m ? parseInt(m[1]) : 0; })
      .filter(Boolean);
    let vmlCounter = existingVmlNums.length > 0 ? Math.max(...existingVmlNums) : 3;

    // Identifier la feuille du site (clone de Conventions) — ne PAS toucher les autres
    const genWbXmlS = await genZip.file("xl/workbook.xml")!.async("string");
    const genWbRelsXmlS = await genZip.file("xl/_rels/workbook.xml.rels")!.async("string");
    const siteSheetFilesS = new Set<string>();
    const escapedSiteS = siteName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const sheetRowS = genWbXmlS.match(new RegExp(`<sheet[^>]+name="${escapedSiteS}"[^>]+r:id="(rId\\d+)"`, "i"));
    if (sheetRowS) {
      const relS = genWbRelsXmlS.match(new RegExp(`Id="${sheetRowS[1]}"[^>]+Target="([^"]+)"`, "i"));
      if (relS) siteSheetFilesS.add(`xl/${relS[1].replace(/^\//, "")}`);
    }

    for (const genName of [...siteSheetFilesS]) {
      const genFile = genZip.file(genName);
      if (!genFile) continue;

      let genXml = await genFile.async("string");
      let legacyDrawingHFTag: string | null = null;
      let pageSetupTag: string | null = null;
      let ridVml = "rId2";
      let ridPrinter = "rId1";

      if (vmlContent) {
        const sheetNum  = parseInt(genName.match(/sheet(\d+)\.xml$/)![1]);
        const newVmlNum = ++vmlCounter;
        const newSpid   = sheetNum * 1024 + 1;
        const updatedVml = vmlContent
          .replace(/<o:idmap v:ext="edit" data="\d+"\/>/gi, `<o:idmap v:ext="edit" data="${sheetNum}"/>`)
          .replace(/o:spid="_x0000_s\d+"/gi, `o:spid="_x0000_s${newSpid}"`);
        genZip.file(`xl/drawings/vmlDrawing${newVmlNum}.vml`, updatedVml);
        if (vmlRelsContent) genZip.file(`xl/drawings/_rels/vmlDrawing${newVmlNum}.vml.rels`, vmlRelsContent);

        const genRelFile = `xl/worksheets/_rels/${genName.split("/").pop()}.rels`;
        const existingRels = genZip.file(genRelFile)
          ? await genZip.file(genRelFile)!.async("string")
          : null;

        if (existingRels) {
          // Aligner les r:id avec ceux attendus par le template
          ridVml = templateLegacyDrawingRid || ridVml;
          ridPrinter = templatePageSetupRid || ridPrinter;

          let merged = existingRels;
          // Retirer toute ancienne entrée printer/vml pour éviter chevauchement
          merged = merged.replace(
            /<Relationship\b[^>]*Type="[^"]*printerSettings[^"]*"[^>]*\/>/gi,
            ""
          );
          merged = merged.replace(
            /<Relationship\b[^>]*Type="[^"]*vmlDrawing[^"]*"[^>]*\/>/gi,
            ""
          );

          const printerEntry = printerTarget
            ? `<Relationship Id="${ridPrinter}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="${printerTarget}"/>`
            : "";
          const vmlEntry = `<Relationship Id="${ridVml}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${newVmlNum}.vml"/>`;

          merged = merged.replace(
            "</Relationships>",
            `${printerEntry}${vmlEntry}</Relationships>`
          );
          genZip.file(genRelFile, merged);
        } else {
          const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
            + (printerTarget ? `<Relationship Id="${ridPrinter}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="${printerTarget}"/>` : "")
            + `<Relationship Id="${ridVml}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${newVmlNum}.vml"/>`
            + `</Relationships>`;
          genZip.file(genRelFile, rels);
        }

        legacyDrawingHFTag = `<legacyDrawingHF r:id="${ridVml}"/>`;
        pageSetupTag = printerTarget
          ? `<pageSetup paperSize="9" scale="85" orientation="portrait" r:id="${ridPrinter}"/>`
          : `<pageSetup paperSize="9" scale="85" orientation="portrait"/>`;
      }

      // Supprimer les éléments existants et les réinsérer dans le bon ordre OOXML
      for (const t of ["pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","legacyDrawingHF"]) {
        genXml = genXml.replace(new RegExp(`\\s*<${t}\\b[^>]*/>`, "gi"), "");
        genXml = genXml.replace(new RegExp(`\\s*<${t}\\b[^>]*>[\\s\\S]*?<\\/${t}>`, "gi"), "");
      }
      const tail = [tPageMargins, pageSetupTag, tHeaderFooter, legacyDrawingHFTag]
        .filter(Boolean).join("");
      genXml = genXml.replace("</worksheet>", tail + "</worksheet>");
      genZip.file(genName, genXml);
    }

    console.log("Phase 3 (single) : VML/en-têtes restaurés avec succès");
    return await genZip.generateAsync({ type: "arraybuffer" });
  } catch (err) {
    console.error("Phase 3 (single) ERREUR — fallback sans VML :", err);
    return excelJsBuffer;
  }
}

export async function generateMultiInvoiceFile(
  conventions: ConventionData[],
  templateSource: string | ArrayBuffer,
  firstInvoiceNumber?: string
): Promise<ArrayBuffer> {
  // Charger le template
  let arrayBuffer: ArrayBuffer;
  if (typeof templateSource === 'string') {
    const response = await fetch(encodeURI(templateSource));
    if (!response.ok) {
      throw new Error(`Impossible de charger le template: ${response.status}`);
    }
    const ct = response.headers.get('content-type') || '';
    if (ct.includes('text/html')) throw new Error('Le template retourné est du HTML, pas un fichier Excel. Vérifiez le chemin du template.');
    arrayBuffer = await response.arrayBuffer();
  } else {
    arrayBuffer = templateSource;
  }

  // Regrouper les conventions par site
  const conventionsBySite = groupConventionsBySite(
    conventions.length > 0 ? conventions : []
  );

  // Trier les sites ("AUTRES" à la fin)
  const siteNames = Array.from(conventionsBySite.keys()).sort((a, b) => {
    if (a === 'AUTRES') return 1;
    if (b === 'AUTRES') return -1;
    return a.localeCompare(b);
  });

  // Parsing du numéro de facture de départ
  const parsedFirstNum: ParsedInvoiceNum | null =
    firstInvoiceNumber ? parseInvoiceNumber(firstInvoiceNumber) : null;

  if (siteNames.length === 0) {
    console.warn("Aucune convention à traiter");
    return arrayBuffer;
  }

  console.log(`Sites trouvés: ${siteNames.join(', ')}`);

  // --- Helpers XML réutilisables ---
  function extractTagBlock(xml: string, tagName: string): string | null {
    const re = new RegExp(`<${tagName}\\b[^>]*>[\\s\\S]*?<\\/${tagName}>`, "i");
    const m = xml.match(re);
    return m ? m[0] : null;
  }
  function extractSelfClosing(xml: string, tagName: string): string | null {
    const re = new RegExp(`<${tagName}\\b[^>]*/>`, "i");
    const m = xml.match(re);
    return m ? m[0] : null;
  }
  function ensureTagBlock(genXml: string, tagName: string, block: string | null): string {
    if (!block) return genXml;
    const re = new RegExp(`<${tagName}\\b[^>]*>[\\s\\S]*?<\\/${tagName}>`, "i");
    if (re.test(genXml)) return genXml.replace(re, block);
    return genXml.replace("</worksheet>", `${block}</worksheet>`);
  }
  function ensureSelfClosing(genXml: string, tagName: string, tag: string | null): string {
    if (!tag) return genXml;
    const re = new RegExp(`<${tagName}\\b[^>]*/>`, "i");
    if (re.test(genXml)) return genXml.replace(re, tag);
    return genXml.replace("</worksheet>", `${tag}</worksheet>`);
  }

  // Phase 1 — JSZip : dupliquer la feuille "Conventions" pour chaque site
  // (préserve 100% de la mise en forme, bordures, fusions et VML d'origine)
  const JSZip = (await import("jszip")).default;
  const preparedBuffer = await precloneConventionsSheet(arrayBuffer, siteNames, JSZip);

  // Phase 2 — ExcelJS : charger le classeur pré-cloné et remplir les données
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(preparedBuffer);

  let invoiceNumber = 1;
  for (const siteName of siteNames) {
    const siteConventions = conventionsBySite.get(siteName)!;
    console.log(`\nRemplissage "${siteName}" — ${siteConventions.length} convention(s)`);

    const sheet = workbook.getWorksheet(siteName);
    if (!sheet) {
      console.warn(`Feuille "${siteName}" introuvable dans le classeur pré-cloné`);
      continue;
    }
    invoiceNumber = await generateInvoicesOnSheet(sheet, siteConventions, invoiceNumber, parsedFirstNum);
  }

  // Phase 3 — JSZip post-processing : restaurer VML/en-têtes/marges
  // (ExcelJS supprime ces informations lors de l'écriture)
  const excelJsRaw = await workbook.xlsx.writeBuffer({ useStyles: true });
  // Normaliser en ArrayBuffer pour compatibilité navigateur avec JSZip
  const excelJsBuffer = (excelJsRaw instanceof ArrayBuffer
    ? excelJsRaw
    : (excelJsRaw as Uint8Array).buffer.slice(
        (excelJsRaw as Uint8Array).byteOffset,
        (excelJsRaw as Uint8Array).byteOffset + (excelJsRaw as Uint8Array).byteLength
      )) as ArrayBuffer;

  try {
    const origZip = await JSZip.loadAsync(arrayBuffer.slice(0));
    const genZip  = await JSZip.loadAsync(excelJsBuffer);

    // Supprimer calcChain.xml si présent (évite l'erreur de réparation Excel)
    if (genZip.file('xl/calcChain.xml')) genZip.remove('xl/calcChain.xml');
    const genCtFile2 = genZip.file('[Content_Types].xml');
    if (genCtFile2) {
      let genCtXml2 = await genCtFile2.async('string');
      genCtXml2 = genCtXml2.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
      genZip.file('[Content_Types].xml', genCtXml2);
    }

    // 3a. Copier médias + printerSettings depuis le template original
    for (const filename of Object.keys(origZip.files)) {
      if (
        (filename.startsWith("xl/media/") || filename.startsWith("xl/printerSettings/")) &&
        origZip.file(filename) && !genZip.file(filename)
      ) {
        genZip.file(filename, await origZip.file(filename)!.async("arraybuffer"));
      }
    }

    // 3b. Localiser la feuille "Conventions" dans le template pour extraire VML + tags
    const origWbXml     = await origZip.file("xl/workbook.xml")!.async("string");
    const origWbRelsXml = await origZip.file("xl/_rels/workbook.xml.rels")!.async("string");

    const convSheetRow   = origWbXml.match(/<sheet[^>]+name="Conventions"[^>]*>/i);
    const convRidM       = convSheetRow ? convSheetRow[0].match(/r:id="([^"]+)"/i) : null;
    const convRid        = convRidM ? convRidM[1] : null;
    const convRelM       = convRid
      ? origWbRelsXml.match(new RegExp(`<Relationship[^>]+Id="${convRid}"[^>]+>`, "i"))
      : null;
    const convTargetM    = convRelM ? convRelM[0].match(/Target="([^"]+)"/i) : null;
    const convSheetFile  = convTargetM
      ? `xl/${convTargetM[1].replace(/^\//, "")}`
      : null;

    if (!convSheetFile || !origZip.file(convSheetFile)) {
      return await genZip.generateAsync({ type: "arraybuffer" });
    }

    const convSheetXml  = await origZip.file(convSheetFile)!.async("string");
    const convRelsFile  = `xl/worksheets/_rels/${convSheetFile.split("/").pop()}.rels`;
    const convRelsXml   = origZip.file(convRelsFile)
      ? await origZip.file(convRelsFile)!.async("string")
      : null;

    // Localiser le VML drawing de la feuille Conventions
    const vmlRelM       = convRelsXml
      ? convRelsXml.match(/<Relationship[^>]+Type="[^"]*vmlDrawing[^"]*"[^>]+>/i)
      : null;
    const vmlTargetM    = vmlRelM ? vmlRelM[0].match(/Target="([^"]+)"/i) : null;
    const vmlBaseName   = vmlTargetM ? vmlTargetM[1].replace(/^.*\//, "") : null; // "vmlDrawing2.vml"
    const vmlFileInZip  = vmlBaseName ? `xl/drawings/${vmlBaseName}` : null;
    const vmlRelsInZip  = vmlBaseName ? `xl/drawings/_rels/${vmlBaseName}.rels` : null;

    const vmlContent    = vmlFileInZip && origZip.file(vmlFileInZip)
      ? await origZip.file(vmlFileInZip)!.async("string")
      : null;
    const vmlRelsContent = vmlRelsInZip && origZip.file(vmlRelsInZip)
      ? await origZip.file(vmlRelsInZip)!.async("string")
      : null;

    // Tags XML à restaurer sur chaque feuille
    const tHeaderFooter  = extractTagBlock(convSheetXml, "headerFooter");
    const tPageMargins   = extractSelfClosing(convSheetXml, "pageMargins")
                        || extractTagBlock(convSheetXml, "pageMargins");
    const tRowBreaks     = extractTagBlock(convSheetXml, "rowBreaks");
    const tColBreaks     = extractTagBlock(convSheetXml, "colBreaks");

    // Récupérer les r:id attendus par le template (souvent rId2 pour le logo VML)
    const templateLegacyDrawingRid =
      convSheetXml.match(/<legacyDrawingHF\b[^>]*\br:id="(rId\d+)"[^>]*\/>/i)?.[1] || null;
    const templatePageSetupRid =
      convSheetXml.match(/<pageSetup\b[^>]*\br:id="(rId\d+)"[^>]*\/?>/i)?.[1] || null;

    // printerSettings reference (rId1 dans les rels d'origine)
    const printerRelM   = convRelsXml
      ? convRelsXml.match(/<Relationship[^>]+Type="[^"]*printerSettings[^"]*"[^>]+>/i)
      : null;
    const printerTarget = printerRelM
      ? (printerRelM[0].match(/Target="([^"]+)"/i) || [])[1]
      : null; // e.g. "../printerSettings/printerSettings2.bin"

    // Trouver le compteur VML max déjà présent dans le template original
    const existingVmlNums = Object.keys(origZip.files)
      .map((f) => { const m = f.match(/xl\/drawings\/vmlDrawing(\d+)\.vml$/); return m ? parseInt(m[1]) : 0; })
      .filter(Boolean);
    let vmlCounter = existingVmlNums.length > 0 ? Math.max(...existingVmlNums) : 3;

    // Helper : supprime toutes les occurrences d'un tag dans le XML
    function stripTag(xml: string, tagName: string): string {
      // Balise auto-fermante : <tag ... />
      xml = xml.replace(new RegExp(`\\s*<${tagName}\\b[^>]*/>`, "gi"), "");
      // Balise ouvrante/fermante : <tag ...>...</tag>
      xml = xml.replace(new RegExp(`\\s*<${tagName}\\b[^>]*>[\\s\\S]*?<\\/${tagName}>`, "gi"), "");
      return xml;
    }

    // Helper : insère tous les éléments de la queue dans le BON ORDRE OOXML
    // Ordre requis : pageMargins → pageSetup → headerFooter → rowBreaks → colBreaks
    //               → drawing (ExcelJS gère) → legacyDrawingHF
    function applyWorksheetTail(
      xml: string,
      pageMargins: string | null,
      pageSetup: string | null,
      headerFooter: string | null,
      rowBreaks: string | null,
      colBreaks: string | null,
      legacyDrawingHF: string | null
    ): string {
      // 1. Supprimer les occurrences existantes de chaque élément
      const tags = ["pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "legacyDrawingHF"];
      for (const t of tags) xml = stripTag(xml, t);

      // 2. Construire la queue dans l'ordre correct
      const tail = [
        pageMargins,
        pageSetup,
        headerFooter,
        rowBreaks,
        colBreaks,
        legacyDrawingHF,
      ].filter(Boolean).join("");

      // 3. Insérer juste avant </worksheet>
      return xml.replace("</worksheet>", tail + "</worksheet>");
    }

    // 3c. Identifier les feuilles de SITE (clones de Conventions) — ne PAS toucher Bandes, BK, Recap
    const genWbXml = await genZip.file("xl/workbook.xml")!.async("string");
    const genWbRelsXml = await genZip.file("xl/_rels/workbook.xml.rels")!.async("string");
    const siteSheetFiles = new Set<string>();
    for (const site of siteNames) {
      const escapedSite = site.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const sheetRow = genWbXml.match(new RegExp(`<sheet[^>]+name="${escapedSite}"[^>]+r:id="(rId\\d+)"`, "i"));
      if (!sheetRow) continue;
      const rel = genWbRelsXml.match(new RegExp(`Id="${sheetRow[1]}"[^>]+Target="([^"]+)"`, "i"));
      if (rel) siteSheetFiles.add(`xl/${rel[1].replace(/^\//, "")}`);
    }
    console.log(`Phase 3c: feuilles site à traiter:`, [...siteSheetFiles]);

    for (const genName of [...siteSheetFiles]) {
      const genFile = genZip.file(genName);
      if (!genFile) continue;

      let genXml = await genFile.async("string");

      let ridVml     = "rId2";
      let ridPrinter = "rId1";
      let legacyDrawingHFTag: string | null = null;
      let pageSetupTag: string | null = null;

      if (vmlContent) {
        const sheetNum  = parseInt(genName.match(/sheet(\d+)\.xml$/)![1]);
        const newVmlNum = ++vmlCounter;
        const newVmlFile = `xl/drawings/vmlDrawing${newVmlNum}.vml`;
        const newVmlRels = `xl/drawings/_rels/vmlDrawing${newVmlNum}.vml.rels`;

        // Mettre à jour l'idmap et le spid pour ce numéro de feuille
        const newSpid = sheetNum * 1024 + 1;
        const updatedVml = vmlContent
          .replace(/<o:idmap v:ext="edit" data="\d+"\/>/gi, `<o:idmap v:ext="edit" data="${sheetNum}"/>`)
          .replace(/o:spid="_x0000_s\d+"/gi, `o:spid="_x0000_s${newSpid}"`);
        genZip.file(newVmlFile, updatedVml);
        if (vmlRelsContent) genZip.file(newVmlRels, vmlRelsContent);

        // Construire / fusionner le fichier rels sans conflit de rId
        const genRelFile = `xl/worksheets/_rels/${genName.split("/").pop()}.rels`;
        const existingGenRels = genZip.file(genRelFile)
          ? await genZip.file(genRelFile)!.async("string")
          : null;

        if (existingGenRels) {
          // On force les mêmes r:id que le template pour éviter les décalages
          ridVml = templateLegacyDrawingRid || ridVml;
          ridPrinter = templatePageSetupRid || ridPrinter;

          // Retirer toute ancienne entrée printer/vml pour éviter les conflits / chevauchements
          let mergedRels = existingGenRels;
          mergedRels = mergedRels.replace(
            /<Relationship\b[^>]*Type="[^"]*printerSettings[^"]*"[^>]*\/>/gi,
            ""
          );
          mergedRels = mergedRels.replace(
            /<Relationship\b[^>]*Type="[^"]*vmlDrawing[^"]*"[^>]*\/>/gi,
            ""
          );

          const printerEntry = printerTarget
            ? `<Relationship Id="${ridPrinter}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="${printerTarget}"/>`
            : "";
          const vmlEntry = `<Relationship Id="${ridVml}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${newVmlNum}.vml"/>`;

          mergedRels = mergedRels.replace(
            "</Relationships>",
            `${printerEntry}${vmlEntry}</Relationships>`
          );
          genZip.file(genRelFile, mergedRels);
        } else {
          const printerEntry = printerTarget
            ? `<Relationship Id="${ridPrinter}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="${printerTarget}"/>`
            : "";
          const newRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
            + printerEntry
            + `<Relationship Id="${ridVml}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${newVmlNum}.vml"/>`
            + `</Relationships>`;
          genZip.file(genRelFile, newRels);
        }

        legacyDrawingHFTag = `<legacyDrawingHF r:id="${ridVml}"/>`;
        pageSetupTag = printerTarget
          ? `<pageSetup paperSize="9" scale="85" orientation="portrait" r:id="${ridPrinter}"/>`
          : `<pageSetup paperSize="9" scale="85" orientation="portrait"/>`;
      }

      // Appliquer tous les éléments dans le bon ordre OOXML
      genXml = applyWorksheetTail(
        genXml,
        tPageMargins,
        pageSetupTag,
        tHeaderFooter,
        tRowBreaks,
        tColBreaks,
        legacyDrawingHFTag
      );

      genZip.file(genName, genXml);
    }

    const finalBuffer = await genZip.generateAsync({ type: "arraybuffer" });
    console.log(`Fichier généré — ${siteNames.length} feuille(s) site`);
    return finalBuffer;
  } catch (error) {
    console.error("Erreur post-processing (fallback) :", error);
    return excelJsBuffer;
  }
}
