// Mapping : colonnes de la convention (ou champs calculés) → cellules du modèle de facture.
// Ajustez les adresses (B4, H15, etc.) en fonction de votre fichier de facture Excel.

export const CELL_MAPPING: Record<string, string> = {
  // Champs convention (extraits du fichier Excel)
  "NOM DU CLIENT": "B4",
  "N° CONVENTION": "B5",
  "OBJET DE LA CONVENTION": "B6",
  SITE: "B7",
  "Date de debut": "B8",
  "Date de fin": "B9",
  "Durée": "B10",
  MONTANT: "B11",
  // Champs facture : Montant HT = MONTANT, Taxe = 0, Accompte = 0, TTC = HT, Solde = MONTANT
  "Montant HT": "H12",
  Taxe: "H13",
  Accompte: "H14",
  "Montant TTC": "H15",
  Solde: "H16",
};
