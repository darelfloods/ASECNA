import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';

export interface BonCommandeData {
  cs: string;
  cr: string;
  cc: string;
  article: string;
  exercice: string;
  fournisseurNom: string;
  fournisseurAdresse1: string;
  fournisseurAdresse2: string;
  codeFournisseur: string;
  lignes: { description: string; quantite: string; prixUnitaire: string; total: string }[];
  montantTotalLettres: string;
  delaiLivraison: string;
  lieu: string;
  date: string;
  numeroEngagement: string;
  operation: string;
  numeroSerie: string;
  numeroBon: string;
  codeIndividuel: string;
  compteLimitatif: string;
  operationEquipement: string;
  compteDe: string;
  montantAD: string;
  engagementsAnterieurs: string;
}

export async function fillBonCommandePdf(
  templatePdf: ArrayBuffer,
  data: BonCommandeData
): Promise<Blob> {
  const doc = await PDFDocument.load(templatePdf);
  const page = doc.getPages()[0];
  const { width: W, height: H } = page.getSize();
  const rot = page.getRotation().angle;

  console.log(`[BonCommande PDF] width=${W} height=${H} rotation=${rot}`);

  const font = await doc.embedFont(StandardFonts.Helvetica);
  const boldFont = await doc.embedFont(StandardFonts.HelveticaBold);
  const color = rgb(0, 0, 0);

  // --- DEBUG : afficher les dimensions en petit dans le coin bas-gauche ---
  page.drawText(`${Math.round(W)}x${Math.round(H)} r=${rot}`, {
    x: 5, y: 5, size: 5, font, color: rgb(0.7, 0.7, 0.7),
  });

  // --- Helpers ---
  const cleanText = (s: string): string => {
    if (!s) return '';
    return String(s)
      .replace(/[\u202F\u00A0]/g, ' ')  // Espaces insécables -> espaces normaux
      .replace(/[^\x20-\x7E\u00C0-\u00FF]/g, ''); // Garder ASCII étendu + accents français
  };

  const draw = (s: string, x: number, y: number, size = 8, bold = false) => {
    if (!s) return;
    const cleaned = cleanText(s);
    page.drawText(cleaned, { x, y, size, font: bold ? boldFont : font, color });
  };

  const charBoxes = (s: string, x: number, y: number, boxW = 14, size = 9) => {
    for (let i = 0; i < s.length; i++) {
      draw(s[i], x + i * boxW + 3, y, size);
    }
  };

  // --- Valeurs calculees ---
  const montantTotal = data.lignes.reduce((s, l) => s + (parseFloat(l.total) || 0), 0);
  const montantTotalStr = montantTotal.toLocaleString('fr-FR').replace(/[\u202F\u00A0]/g, ' ');
  const ant = parseFloat((data.engagementsAnterieurs || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const ad = parseFloat((data.montantAD || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const cumul = montantTotal + ant;
  const disponible = ad - cumul;

  // ================================================================
  //  DROITE — BON DE COMMANDE (A4 Paysage 842x595 pt)
  //  Coordonnées recalibrées basées sur le template réel
  // ================================================================

  // --- En-tête : CS, CR, CC, Exercice (cases à droite) ---
  charBoxes(data.cs,       760, H - 83,  13.5);   // CS
  charBoxes(data.cr,       760, H - 111, 13.5);   // CR
  charBoxes(data.cc,       760, H - 139, 13.5);   // CC
  charBoxes(data.exercice, 760, H - 195, 13.5);   // Exercice

  // --- Adresse du fournisseur ---
  draw(data.fournisseurNom,      475, H - 244, 8);
  draw(data.fournisseurAdresse1, 475, H - 258, 8);
  draw(data.fournisseurAdresse2, 475, H - 272, 8);

  // --- Code fournisseur (4 cases) ---
  charBoxes(data.codeFournisseur, 740, H - 327, 13.5);

  // --- Tableau : Detail de la commande ---
  const tblStartY = H - 453;  // Première ligne du tableau
  const rowH = 17.5;  // Hauteur entre les lignes
  for (let i = 0; i < data.lignes.length && i < 8; i++) {
    const l = data.lignes[i];
    const y = tblStartY - i * rowH;
    draw(l.description,  340, y, 7);   // Colonne Description
    draw(l.quantite,     565, y, 7);   // Colonne Quantité
    draw(l.prixUnitaire, 645, y, 7);   // Colonne Prix unitaire
    draw(l.total,        765, y, 7);   // Colonne Total
  }

  // --- Montant total en chiffre ---
  draw(montantTotalStr, 765, H - 594, 9, true);

  // --- Montant total en lettre ---
  draw(data.montantTotalLettres, 488, H - 615, 7);

  // --- Délai de livraison ---
  draw(data.delaiLivraison, 465, H - 635, 7);

  // --- Section basse ---
  draw(data.compteLimitatif,  385, H - 678, 7);  // Compte limitatif
  draw(data.operation,        385, H - 699, 7);  // Opération
  draw(data.numeroEngagement, 385, H - 719, 7);  // N° engagement

  // Lieu et date — VISA Service Engagements
  draw(data.lieu, 535, H - 678, 7);
  draw(data.date, 625, H - 678, 7);

  // Lieu et date — VISA Ordonnateur
  draw(data.lieu, 685, H - 678, 7);
  draw(data.date, 775, H - 678, 7);

  // ================================================================
  //  GAUCHE — BON D'ENGAGEMENT
  // ================================================================

  // --- CS / Article / CR / CC ---
  charBoxes(data.cs,      88, H - 267, 11.5);   // CS
  charBoxes(data.article, 258, H - 267, 11.5);  // Article

  // Numéro du bon de commande
  draw(data.numeroBon || data.numeroSerie, 208, H - 301, 7);

  charBoxes(data.cr, 88, H - 314, 11.5);  // CR
  charBoxes(data.cc, 88, H - 359, 11.5);  // CC

  // --- Fournisseur ---
  draw(data.fournisseurNom, 158, H - 389, 7);

  // --- Code individuel ---
  draw(data.codeIndividuel, 158, H - 415, 7);

  // --- Numéro d'engagement ---
  draw(data.numeroEngagement, 158, H - 453, 7);

  // --- Compte limitatif ---
  draw(data.compteLimitatif, 158, H - 482, 7);

  // --- Opération d'équipement ---
  draw(data.operationEquipement, 158, H - 511, 7);

  // --- Compte de comptabilité générale ---
  draw(data.compteDe, 158, H - 540, 7);

  // --- Montant A.D. ---
  draw(data.montantAD, 158, H - 573, 7);

  // --- Montant du bon ---
  draw(montantTotalStr, 158, H - 586, 7);

  // --- Engagements antérieurs ---
  draw(data.engagementsAnterieurs, 158, H - 618, 7);

  // --- Cumul des engagements ---
  if (cumul > 0) draw(cumul.toLocaleString('fr-FR').replace(/[\u202F\u00A0]/g, ' '), 158, H - 632, 7);

  // --- Disponible ---
  if (ad > 0) draw(disponible.toLocaleString('fr-FR').replace(/[\u202F\u00A0]/g, ' '), 158, H - 669, 7);

  // --- Lieu et date (bas gauche) ---
  draw(data.lieu, 63, H - 720, 7);
  const dateParts = (data.date || '').split('/');
  if (dateParts.length === 3) {
    charBoxes(dateParts[0], 123, H - 740, 11.5);  // JOUR
    charBoxes(dateParts[1], 171, H - 740, 11.5);  // MOIS
    charBoxes(dateParts[2], 220, H - 740, 11.5);  // ANNEE
  }

  const pdfBytes = await doc.save();
  return new Blob([new Uint8Array(pdfBytes) as unknown as ArrayBuffer], { type: 'application/pdf' });
}
