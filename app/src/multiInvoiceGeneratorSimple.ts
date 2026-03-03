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
  "MVENGUE": "FVC",
  "M'VENGUE": "FVC",
  "M'VENGUE": "FVC",
  "MVENGE": "FVC",
  "PORT-GENTIL": "POG",
  "PORT GENTIL": "POG",
  "POG": "POG",
  "LIBREVILLE": "AVOIR",
  "LBV": "AVOIR",
};

/**
 * Normalise le nom du site
 */
function normalizeSite(site: string): string {
  return site.toUpperCase().replace(/['']/g, "'").trim();
}

/**
 * Obtient le nom de la feuille pour un site donné
 */
function getSheetNameForSite(site: string): string {
  const normalized = normalizeSite(site);
  return SITE_TO_SHEET[normalized] || "FVC"; // Par défaut FVC
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
  
  // Chercher "Facture N°" dans les 10 premières lignes et toutes les colonnes
  for (let row = 1; row <= 10; row++) {
    for (let col = 1; col <= 25; col++) {
      const cell = sheet.getCell(row, col);
      const value = cell.value;
      
      if (value && typeof value === "string" && value.match(/Facture\s*N°/i)) {
        // Vérifier si on n'a pas déjà ce bloc (éviter les doublons)
        const existingBlock = blocks.find(b => Math.abs(b.col - col) < 5);
        if (!existingBlock) {
          blocks.push({
            col: Math.max(1, col - 1), // Colonne de départ (1 avant le "Facture N°")
            factureRow: row,
          });
        }
      }
    }
  }
  
  // Trier les blocs par colonne
  blocks.sort((a, b) => a.col - b.col);
  
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
  
  try {
    // Numéro de facture (sur la ligne détectée, décalé de 1 colonne)
    const factureCell = sheet.getCell(baseRow, baseCol + 1);
    setCellValuePreservingStyle(factureCell, `Facture N°${String(invoiceNumber).padStart(3, "0")}`);
    
    // Client (généralement 3-4 lignes plus bas, décalé de 4-5 colonnes)
    const clientCell = sheet.getCell(baseRow + 3, baseCol + 5);
    const clientStyle = preserveCellStyle(clientCell);
    clientCell.value = convention["NOM DU CLIENT"];
    // Préserver le style mais ajouter le wrap si nécessaire
    if (clientStyle.alignment) {
      clientCell.alignment = { ...clientStyle.alignment, wrapText: true };
    } else {
      clientCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
    }
    if (clientStyle.font) clientCell.font = clientStyle.font;
    if (clientStyle.border) clientCell.border = clientStyle.border;
    if (clientStyle.fill) clientCell.fill = clientStyle.fill;
    
    // Site (généralement 7-8 lignes plus bas, colonne de base + 1)
    const siteCell = sheet.getCell(baseRow + 7, baseCol + 1);
    setCellValuePreservingStyle(siteCell, `Site: ${convention.SITE}`);
    
    // Période (généralement 12 lignes plus bas, colonne de base)
    const periodeCell = sheet.getCell(baseRow + 12, baseCol);
    const periodeStyle = preserveCellStyle(periodeCell);
    periodeCell.value = `Du ${convention["Date de debut"]} au ${convention["Date de fin"]}`;
    if (periodeStyle.alignment) {
      periodeCell.alignment = { ...periodeStyle.alignment, wrapText: true };
    }
    if (periodeStyle.font) periodeCell.font = periodeStyle.font;
    if (periodeStyle.border) periodeCell.border = periodeStyle.border;
    if (periodeStyle.fill) periodeCell.fill = periodeStyle.fill;
    
    // Désignation / Objet (généralement 13-14 lignes plus bas, colonne de base + 1)
    const designationCell = sheet.getCell(baseRow + 14, baseCol + 1);
    const designationStyle = preserveCellStyle(designationCell);
    designationCell.value = convention["OBJET DE LA CONVENTION"];
    if (designationStyle.alignment) {
      designationCell.alignment = { ...designationStyle.alignment, wrapText: true };
    }
    if (designationStyle.font) designationCell.font = designationStyle.font;
    if (designationStyle.border) designationCell.border = designationStyle.border;
    if (designationStyle.fill) designationCell.fill = designationStyle.fill;
    
    // Numéro de convention (généralement 15-16 lignes plus bas, colonne de base + 1)
    const convCell = sheet.getCell(baseRow + 15, baseCol + 1);
    setCellValuePreservingStyle(convCell, convention["N° CONVENTION"]);
    
    // Montant HT (généralement 15-16 lignes plus bas, colonne de base + 7)
    const montantCell = sheet.getCell(baseRow + 15, baseCol + 7);
    const montantStyle = preserveCellStyle(montantCell);
    
    // S'assurer que c'est bien un nombre
    const montantValue = typeof convention.MONTANT === 'number' 
      ? convention.MONTANT 
      : parseFloat(String(convention.MONTANT).replace(/[^\d.-]/g, ''));
    
    if (!isNaN(montantValue)) {
      montantCell.value = montantValue;
      // Restaurer le style incluant le format numérique
      if (montantStyle.font) montantCell.font = montantStyle.font;
      if (montantStyle.alignment) montantCell.alignment = montantStyle.alignment;
      if (montantStyle.border) montantCell.border = montantStyle.border;
      if (montantStyle.fill) montantCell.fill = montantStyle.fill;
      if (montantStyle.numFmt) {
        montantCell.numFmt = montantStyle.numFmt;
      } else {
        montantCell.numFmt = '#,##0';
      }
    } else {
      console.warn(`Montant invalide pour ${convention["NOM DU CLIENT"]}: ${convention.MONTANT}`);
      montantCell.value = 0;
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
  templatePath: string
): Promise<ArrayBuffer> {
  // Charger le template
  const response = await fetch(templatePath);
  const arrayBuffer = await response.arrayBuffer();
  
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
