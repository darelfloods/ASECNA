// Mapping : colonnes de la convention (ou champs calculés) → cellules du modèle de facture.
// Ajustez les adresses (B4, H15, etc.) en fonction de votre fichier de facture Excel.

export const CELL_MAPPING: Record<string, string> = {
  // Nouvelles coordonnées pour la feuille "Bandes" (1er bloc)
  "NOM DU CLIENT": "F10",
  "N° CONVENTION": "C15", // Série N°
  "OBJET DE LA CONVENTION": "B20", // Désignation
  SITE: "C14",
  "Date de debut": "A19", // Format fusionné
  "Date de fin": "A19", // Mappé sur la même cellule
  "Durée": "A19",
  MONTANT: "H22",
  // Champs facture : Montant HT = MONTANT, Taxe = 0, Accompte = 0, TTC = HT, Solde = MONTANT
  "Montant HT": "B27",
  Taxe: "D27",
  Accompte: "F27",
  "Montant TTC": "H27", // Solde
  Solde: "H27",
};
