// Mapping des sites vers les feuilles Excel
export const SITE_TO_SHEET: Record<string, string> = {
  "MVENGUE": "FVC",
  "M'VENGUE": "FVC",
  "PORT-GENTIL": "POG",
  "PORT GENTIL": "POG",
  "POG": "POG",
  "LIBREVILLE": "AVOIR",
  // Ajoutez d'autres mappings si nécessaire
};

// Configuration des blocs de facture par feuille
// Chaque bloc représente une facture (ensemble de colonnes)
export interface InvoiceBlock {
  sheetName: string;
  blockIndex: number;
  startCol: number; // Index de la colonne de départ (0-based)
  // Mapping des champs vers les cellules relatives au bloc
  cellMapping: {
    numeroFacture: { row: number; col: number }; // Ex: row 2, col 1 (B3)
    client: { row: number; col: number };
    site: { row: number; col: number };
    serie: { row: number; col: number };
    periode: { row: number; col: number };
    designation: { row: number; col: number };
    numeroConvention: { row: number; col: number };
    montantHT: { row: number; col: number };
    montantTTC: { row: number; col: number };
  };
}

// Configuration des blocs pour la feuille FVC
export const FVC_BLOCKS: InvoiceBlock[] = [
  {
    sheetName: "FVC",
    blockIndex: 0,
    startCol: 0, // Colonnes A-H
    cellMapping: {
      numeroFacture: { row: 2, col: 1 }, // B3
      client: { row: 5, col: 5 }, // F6
      site: { row: 9, col: 1 }, // B10
      serie: { row: 11, col: 1 }, // B12
      periode: { row: 14, col: 0 }, // A15
      designation: { row: 15, col: 1 }, // B16
      numeroConvention: { row: 17, col: 1 }, // B18 (exemple)
      montantHT: { row: 17, col: 7 }, // H18
      montantTTC: { row: 17, col: 7 }, // H18 (même valeur)
    },
  },
  {
    sheetName: "FVC",
    blockIndex: 1,
    startCol: 8, // Colonnes I-P
    cellMapping: {
      numeroFacture: { row: 2, col: 9 }, // J3
      client: { row: 5, col: 13 }, // N6 (estimation)
      site: { row: 10, col: 10 }, // K11
      serie: { row: 11, col: 9 }, // J12 (estimation)
      periode: { row: 14, col: 8 }, // I15
      designation: { row: 16, col: 9 }, // J17
      numeroConvention: { row: 17, col: 9 }, // J18
      montantHT: { row: 20, col: 9 }, // J21 (estimation)
      montantTTC: { row: 20, col: 9 }, // J21
    },
  },
  {
    sheetName: "FVC",
    blockIndex: 2,
    startCol: 16, // Colonnes Q-X
    cellMapping: {
      numeroFacture: { row: 3, col: 17 }, // R4 (estimation basée sur la structure)
      client: { row: 5, col: 21 }, // V6 (estimation)
      site: { row: 11, col: 17 }, // R12 (estimation)
      serie: { row: 11, col: 17 }, // R12
      periode: { row: 14, col: 16 }, // Q15 (estimation)
      designation: { row: 16, col: 17 }, // R17
      numeroConvention: { row: 17, col: 17 }, // R18
      montantHT: { row: 17, col: 22 }, // W18 (estimation)
      montantTTC: { row: 17, col: 22 }, // W18
    },
  },
];

// Configuration des blocs pour la feuille POG
export const POG_BLOCKS: InvoiceBlock[] = [
  {
    sheetName: "POG",
    blockIndex: 0,
    startCol: 0,
    cellMapping: {
      numeroFacture: { row: 6, col: 1 }, // B7
      client: { row: 9, col: 5 }, // F10
      site: { row: 13, col: 2 }, // C14
      serie: { row: 14, col: 2 }, // C15
      periode: { row: 18, col: 0 }, // A19
      designation: { row: 19, col: 1 }, // B20
      numeroConvention: { row: 20, col: 1 }, // B21 (estimation)
      montantHT: { row: 21, col: 7 }, // H22 (estimation)
      montantTTC: { row: 21, col: 7 }, // H22
    },
  },
  {
    sheetName: "POG",
    blockIndex: 1,
    startCol: 8,
    cellMapping: {
      numeroFacture: { row: 6, col: 8 }, // I7
      client: { row: 9, col: 14 }, // O10 (estimation)
      site: { row: 14, col: 9 }, // J15
      serie: { row: 15, col: 9 }, // J16 (estimation)
      periode: { row: 18, col: 8 }, // I19
      designation: { row: 19, col: 9 }, // J20
      numeroConvention: { row: 20, col: 9 }, // J21
      montantHT: { row: 21, col: 15 }, // P22 (estimation)
      montantTTC: { row: 21, col: 15 }, // P22
    },
  },
  {
    sheetName: "POG",
    blockIndex: 2,
    startCol: 16,
    cellMapping: {
      numeroFacture: { row: 6, col: 16 }, // Q7
      client: { row: 9, col: 22 }, // W10 (estimation)
      site: { row: 14, col: 16 }, // Q15
      serie: { row: 15, col: 16 }, // Q16
      periode: { row: 18, col: 16 }, // Q19
      designation: { row: 19, col: 17 }, // R20
      numeroConvention: { row: 20, col: 17 }, // R21
      montantHT: { row: 21, col: 23 }, // X22 (estimation)
      montantTTC: { row: 21, col: 23 }, // X22
    },
  },
];

// Fonction pour obtenir les blocs disponibles pour un site
export function getBlocksForSite(site: string): InvoiceBlock[] {
  const sheetName = SITE_TO_SHEET[site.toUpperCase()];
  
  if (sheetName === "FVC") {
    return FVC_BLOCKS;
  } else if (sheetName === "POG") {
    return POG_BLOCKS;
  }
  
  // Par défaut, utiliser FVC
  return FVC_BLOCKS;
}

// Fonction pour normaliser le site
export function normalizeSite(site: string): string {
  return site.toUpperCase().replace(/'/g, "'").trim();
}
