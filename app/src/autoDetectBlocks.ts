import * as XLSX from "xlsx";
import { InvoiceBlock } from "./invoiceMapping";

/**
 * Détecte automatiquement les blocs de factures dans une feuille Excel
 * en cherchant les patterns "Facture N°" dans les premières lignes
 */
export function autoDetectInvoiceBlocks(
  sheet: XLSX.WorkSheet,
  sheetName: string
): InvoiceBlock[] {
  const blocks: InvoiceBlock[] = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");

  // Chercher "Facture N°" dans les 15 premières lignes
  const facturePositions: Array<{ row: number; col: number }> = [];

  for (let row = 0; row < Math.min(15, range.e.r); row++) {
    for (let col = 0; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[cellAddress];

      if (cell && cell.v && String(cell.v).match(/Facture\s*N°/i)) {
        facturePositions.push({ row, col });
      }
    }
  }

  console.log(`Feuille ${sheetName}: ${facturePositions.length} blocs détectés`, facturePositions);

  // Pour chaque position de facture détectée, créer un bloc
  facturePositions.forEach((pos, index) => {
    // Déterminer la colonne de départ du bloc (généralement 1 colonne avant ou la même)
    const startCol = Math.max(0, pos.col - 1);

    blocks.push({
      sheetName,
      blockIndex: index,
      startCol,
      cellMapping: {
        // Ces positions sont relatives à la structure observée
        // Vous devrez peut-être les ajuster selon vos besoins
        numeroFacture: { row: pos.row, col: pos.col },
        client: { row: pos.row + 3, col: pos.col + 4 }, // Estimation
        site: { row: pos.row + 7, col: startCol + 1 }, // Estimation
        serie: { row: pos.row + 9, col: startCol + 1 }, // Estimation
        periode: { row: pos.row + 12, col: startCol }, // Estimation
        designation: { row: pos.row + 13, col: startCol + 1 }, // Estimation
        numeroConvention: { row: pos.row + 15, col: startCol + 1 }, // Estimation
        montantHT: { row: pos.row + 15, col: pos.col + 6 }, // Estimation
        montantTTC: { row: pos.row + 15, col: pos.col + 6 }, // Estimation
      },
    });
  });

  return blocks;
}

/**
 * Obtient les blocs disponibles pour une feuille donnée
 */
export async function getAvailableBlocks(
  templatePath: string,
  sheetName: string
): Promise<InvoiceBlock[]> {
  const response = await fetch(templatePath);
  const data = await response.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    console.error(`Feuille ${sheetName} non trouvée`);
    return [];
  }

  return autoDetectInvoiceBlocks(sheet, sheetName);
}
