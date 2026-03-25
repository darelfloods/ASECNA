import * as XLSX from "xlsx";
import { InvoiceBlock, getBlocksForSite, normalizeSite } from "./invoiceMapping";

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
 * Génère un fichier Excel avec plusieurs factures remplies à partir des conventions
 * @param conventions - Liste des conventions à facturer
 * @param templatePath - Chemin vers le fichier template de facturation
 * @returns Le workbook Excel rempli
 */
export async function generateMultiInvoiceFile(
  conventions: ConventionData[],
  templatePath: string
): Promise<XLSX.WorkBook> {
  // Charger le template
  const response = await fetch(templatePath);
  const data = await response.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });

  // Grouper les conventions par site
  const conventionsBySite = groupConventionsBySite(conventions);

  // Pour chaque site, remplir les blocs de factures
  for (const [site, siteConventions] of Object.entries(conventionsBySite)) {
    const blocks = getBlocksForSite(site);
    
    // Remplir autant de factures que possible dans les blocs disponibles
    for (let i = 0; i < Math.min(siteConventions.length, blocks.length); i++) {
      const convention = siteConventions[i];
      const block = blocks[i];
      
      // Remplir le bloc avec les données de la convention
      fillInvoiceBlock(workbook, block, convention, i);
    }
  }

  return workbook;
}

/**
 * Groupe les conventions par site
 */
function groupConventionsBySite(
  conventions: ConventionData[]
): Record<string, ConventionData[]> {
  const groups: Record<string, ConventionData[]> = {};

  for (const convention of conventions) {
    const site = normalizeSite(convention.SITE);
    if (!groups[site]) {
      groups[site] = [];
    }
    groups[site].push(convention);
  }

  return groups;
}

/**
 * Remplit un bloc de facture avec les données d'une convention
 */
function fillInvoiceBlock(
  workbook: XLSX.WorkBook,
  block: InvoiceBlock,
  convention: ConventionData,
  invoiceIndex: number
): void {
  const sheet = workbook.Sheets[block.sheetName];
  if (!sheet) {
    console.error(`Feuille ${block.sheetName} non trouvée`);
    return;
  }

  const mapping = block.cellMapping;

  // Remplir le numéro de facture
  setCellValue(
    sheet,
    mapping.numeroFacture.row,
    mapping.numeroFacture.col,
    `Facture N°${String(invoiceIndex + 1).padStart(3, "0")}`
  );

  // Remplir le client
  setCellValue(
    sheet,
    mapping.client.row,
    mapping.client.col,
    convention["NOM DU CLIENT"]
  );

  // Remplir le site
  const siteValue = `Site: ${convention.SITE}`;
  setCellValue(sheet, mapping.site.row, mapping.site.col, siteValue);

  // Remplir la période
  const periode = `Du ${convention["Date de debut"]} au ${convention["Date de fin"]}`;
  setCellValue(sheet, mapping.periode.row, mapping.periode.col, periode);

  // Remplir la désignation
  setCellValue(
    sheet,
    mapping.designation.row,
    mapping.designation.col,
    convention["OBJET DE LA CONVENTION"]
  );

  // Remplir le numéro de convention
  setCellValue(
    sheet,
    mapping.numeroConvention.row,
    mapping.numeroConvention.col,
    convention["N° CONVENTION"]
  );

  // Remplir le montant HT
  setCellValue(
    sheet,
    mapping.montantHT.row,
    mapping.montantHT.col,
    convention.MONTANT
  );

  // Remplir le montant TTC (même valeur que HT pour l'instant)
  setCellValue(
    sheet,
    mapping.montantTTC.row,
    mapping.montantTTC.col,
    convention.MONTANT
  );
}

/**
 * Définit la valeur d'une cellule dans une feuille
 */
function setCellValue(
  sheet: XLSX.WorkSheet,
  row: number,
  col: number,
  value: string | number
): void {
  const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
  
  // Si la cellule existe déjà, conserver son format
  const existingCell = sheet[cellAddress];
  
  if (existingCell) {
    sheet[cellAddress] = {
      ...existingCell,
      v: value,
      t: typeof value === "number" ? "n" : "s",
    };
  } else {
    sheet[cellAddress] = {
      v: value,
      t: typeof value === "number" ? "n" : "s",
    };
  }
}
