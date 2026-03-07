import * as XLSX from "xlsx";
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

/**
 * Mapping des sites vers les feuilles Excel
 */
const SITE_TO_SHEET: Record<string, string> = {
  // Par défaut, le nouveau modèle utilise "Bandes" ou "Conventions"
  "MVENGUE": "Bandes",
  "M'VENGUE": "Bandes",
  "MVENGE": "Bandes",
  "PORT-GENTIL": "Bandes",
  "PORT GENTIL": "Bandes",
  "POG": "Bandes",
  "LIBREVILLE": "Conventions",
  "LBV": "Conventions",
  "BITAM": "Conventions",
};

/**
 * Normalise le nom du site
 */
function normalizeSite(site: string): string {
  return site.toUpperCase().replace(/['']/g, "'").trim();
}

function getSheetNameForSite(site: string): string {
  const normalized = normalizeSite(site);
  return SITE_TO_SHEET[normalized] || "Bandes"; // Par défaut "Bandes"
}

/**
 * Structure d'un bloc de facture détecté
 */
interface DetectedBlock {
  col: number; // Colonne où commence le bloc
  factureRow: number; // Ligne où se trouve "Facture N°"
}

/**
 * Détecte les blocs de factures dans une feuille
 */
function detectInvoiceBlocks(sheet: ExcelJS.Worksheet): DetectedBlock[] {
  const blocks: DetectedBlock[] = [];

  // Dans le modèle 2026, les factures sont empilées verticalement
  // Chercher "Facture N°" dans les 300 premières lignes, colonnes A ou B (1 ou 2)
  for (let row = 1; row <= 300; row++) {
    for (let col = 1; col <= 5; col++) {
      const cell = sheet.getCell(row, col);
      const value = cell.value;

      let textValue = "";
      if (typeof value === "string") textValue = value;
      else if (value && typeof value === 'object' && 'richText' in value && Array.isArray(value.richText)) {
        textValue = value.richText.map((r: any) => r.text).join("");
      }

      if (textValue && textValue.match(/Facture\s*N°/i)) {
        // Éviter les doublons (même ligne ou ligne très proche)
        const existingBlock = blocks.find(b => Math.abs(b.factureRow - row) < 5);
        if (!existingBlock) {
          blocks.push({
            col: col, // Colonne exacte du texte "Facture N°..."
            factureRow: row,
          });
        }
      }
    }
  }

  // Trier les blocs par ligne (de haut en bas)
  blocks.sort((a, b) => a.factureRow - b.factureRow);

  return blocks;
}

/**
 * Sauvegarder le style d'une cellule avant modification
 */
function preserveCellStyle(cell: ExcelJS.Cell): any {
  return {
    font: cell.font,
    alignment: cell.alignment,
    border: cell.border,
    fill: cell.fill,
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

  // Restaurer le style
  if (style.font) cell.font = style.font;
  if (style.alignment) cell.alignment = style.alignment;
  if (style.border) cell.border = style.border;
  if (style.fill) cell.fill = style.fill;
  if (style.numFmt) cell.numFmt = style.numFmt;
}

/**
 * Remplit un bloc de facture avec les données d'une convention
 */
function fillInvoiceBlock(
  sheet: ExcelJS.Worksheet,
  block: DetectedBlock,
  convention: ConventionData,
  invoiceNumber: number
): void {
  const baseCol = block.col;
  const baseRow = block.factureRow;
  const isBandesInfo = sheet.name === "Bandes" || baseCol === 2; // "Bandes" starts at B7

  try {
    // Numéro de facture
    const textFacture = `Facture N°${String(invoiceNumber).padStart(3, "0")}`;
    const factureCell = sheet.getCell(baseRow, baseCol);
    setCellValuePreservingStyle(factureCell, textFacture);

    if (isBandesInfo) {
      // Configuration "Bandes"
      // Client en F10 (baseRow+3, col F=6)
      const clientCell = sheet.getCell(baseRow + 3, 6);
      setCellValuePreservingStyle(clientCell, convention["NOM DU CLIENT"]);

      // Site en C14 (baseRow+7, col C=3)
      const siteCell = sheet.getCell(baseRow + 7, 3);
      setCellValuePreservingStyle(siteCell, convention.SITE);

      // Période en A19 (baseRow+12, col A=1)
      const periodeCell = sheet.getCell(baseRow + 12, 1);
      setCellValuePreservingStyle(periodeCell, `Du ${convention["Date de debut"]} au ${convention["Date de fin"]}`);

      // Série (N° Convention) en C15 (baseRow+8, col C=3)
      const serieCell = sheet.getCell(baseRow + 8, 3);
      setCellValuePreservingStyle(serieCell, convention["N° CONVENTION"]);

      // Désignation en B20 (baseRow+13, col B=2)
      const designationCell = sheet.getCell(baseRow + 13, 2);
      setCellValuePreservingStyle(designationCell, convention["OBJET DE LA CONVENTION"]);

      // Montant en H22 (baseRow+15, col H=8)
      const montantCell = sheet.getCell(baseRow + 15, 8);
      const montantStyle = preserveCellStyle(montantCell);
      const montantValue = typeof convention.MONTANT === 'number'
        ? convention.MONTANT
        : parseFloat(String(convention.MONTANT).replace(/[^\d.-]/g, ''));
      if (!isNaN(montantValue)) {
        montantCell.value = montantValue;
        if (montantStyle.font) montantCell.font = montantStyle.font;
        if (montantStyle.numFmt) montantCell.numFmt = montantStyle.numFmt;
      }
    } else {
      // Configuration "Conventions" (Facture en A4)
      // Client en F4 (baseRow, col F=6)
      const clientCell = sheet.getCell(baseRow, 6);
      setCellValuePreservingStyle(clientCell, convention["NOM DU CLIENT"]);

      // Site en B9 (baseRow+5, col B=2)
      const siteCell = sheet.getCell(baseRow + 5, 2);
      setCellValuePreservingStyle(siteCell, `Site: ${convention.SITE}`);

      // Période en A13 (baseRow+9, col A=1)
      const periodeCell = sheet.getCell(baseRow + 9, 1);
      setCellValuePreservingStyle(periodeCell, `Du ${convention["Date de debut"]} au ${convention["Date de fin"]}`);

      // Désignation en B14 (baseRow+10, col B=2)
      const designationCell = sheet.getCell(baseRow + 10, 2);
      setCellValuePreservingStyle(designationCell, convention["OBJET DE LA CONVENTION"]);

      // N° Convention en B15 (baseRow+11, col B=2)
      const covCell = sheet.getCell(baseRow + 11, 2);
      setCellValuePreservingStyle(covCell, convention["N° CONVENTION"]);

      // Montant en H16 (baseRow+12, col H=8)
      const montantCell = sheet.getCell(baseRow + 12, 8);
      const montantStyle = preserveCellStyle(montantCell);
      const montantValue = typeof convention.MONTANT === 'number'
        ? convention.MONTANT
        : parseFloat(String(convention.MONTANT).replace(/[^\d.-]/g, ''));
      if (!isNaN(montantValue)) {
        montantCell.value = montantValue;
        if (montantStyle.font) montantCell.font = montantStyle.font;
        if (montantStyle.numFmt) montantCell.numFmt = montantStyle.numFmt;
      }
    }

    // Agrandir la colonne H pour l'affichage correct des montants dépassant la largeur
    const colH = sheet.getColumn(8);
    if (!colH.width || colH.width < 18) {
      colH.width = 18;
    }

  } catch (error) {
    console.error("Erreur lors du remplissage du bloc:", error);
  }
}

/**
 * Génère un fichier Excel avec plusieurs factures remplies
 */
export async function generateMultiInvoiceFile(
  conventions: ConventionData[],
  templateSource: string | ArrayBuffer
): Promise<ArrayBuffer> {
  // Charger le template
  let arrayBuffer: ArrayBuffer;
  if (typeof templateSource === 'string') {
    const response = await fetch(templateSource);
    arrayBuffer = await response.arrayBuffer();
  } else {
    arrayBuffer = templateSource; // Utiliser directement le buffer fourni
  }

  // IMPORTANT: Charger avec l'option pour préserver les images
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  // Grouper les conventions par site
  const conventionsBySite = new Map<string, ConventionData[]>();

  for (const convention of conventions) {
    const sheetName = getSheetNameForSite(convention.SITE);
    if (!conventionsBySite.has(sheetName)) {
      conventionsBySite.set(sheetName, []);
    }
    conventionsBySite.get(sheetName)!.push(convention);
  }

  console.log("Conventions groupées par feuille:", Array.from(conventionsBySite.entries()).map(([k, v]) => `${k}: ${v.length}`));

  // Pour chaque feuille, remplir les blocs disponibles
  let globalInvoiceNumber = 1;

  for (const [sheetName, siteConventions] of conventionsBySite.entries()) {
    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) {
      console.warn(`Feuille ${sheetName} non trouvée dans le template`);
      continue;
    }

    // Vérifier les images présentes dans la feuille
    console.log(`Feuille ${sheetName}: ${sheet.getImages ? sheet.getImages().length : 0} images détectées`);

    // Détecter les blocs de factures dans cette feuille
    const blocks = detectInvoiceBlocks(sheet);
    console.log(`Feuille ${sheetName}: ${blocks.length} blocs détectés`);

    // Remplir autant de blocs que possible
    const nbToFill = Math.min(siteConventions.length, blocks.length);

    for (let i = 0; i < nbToFill; i++) {
      console.log(`Remplissage bloc ${i + 1}/${nbToFill} pour ${siteConventions[i]["NOM DU CLIENT"]}`);
      fillInvoiceBlock(sheet, blocks[i], siteConventions[i], globalInvoiceNumber);
      globalInvoiceNumber++;
    }

    if (siteConventions.length > blocks.length) {
      console.warn(`${siteConventions.length - blocks.length} conventions non traitées sur la feuille ${sheetName} (pas assez de blocs)`);
    }
  }

  // Retourner le fichier généré avec toutes les propriétés préservées
  const buffer = await workbook.xlsx.writeBuffer({
    // Préserver les propriétés du workbook original
    useStyles: true,
    useSharedStrings: true,
  });

  return buffer;
}
