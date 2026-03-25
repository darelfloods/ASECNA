/**
 * SCRIPT D'ANALYSE — Diff entre facture originale (PDF) et facture générée (BandesModule)
 * Exécuter : node scripts/analyse-diff-facture-bandes.js
 */

const DIFF = [
  // ─────────────────────────────────────────────────────────────────
  // SECTION 1 : EN-TÊTE
  // ─────────────────────────────────────────────────────────────────
  {
    id: 1,
    section: 'EN-TÊTE',
    ligne: 'L1 — Siège ASECNA',
    fichier: 'BandesModule.tsx',
    codeRef: 'ligne ~705',
    statut: '❌ INCORRECT',
    original: 'Agence pour la Sécurité de la Navigation Aérienne en Afrique et à Madagascar\nSiège Social : 32 – 38 Av. Thierno Seydou Nourou TALL - B.P. 3144 - Dakar – Sénégal Tél : 33 849 68 08 / 33 849 88 62 - Site web : www.asecna.aero',
    genere:   'Agence pour la Sécurité de la Navigation Aérienne en Afrique et à Madagascar — Siège Social : 32-38 Av. Jean Jaurès DAKAR B.P. 3144 — site web : www.asecna.aero',
    delta: [
      'Adresse : "Jean Jaurès" → "Thierno Seydou Nourou TALL"',
      'Tél manquant : ajouter "Tél : 33 849 68 08 / 33 849 88 62"',
    ],
  },
  {
    id: 2,
    section: 'EN-TÊTE',
    ligne: 'L2 — Délégation Gabon',
    fichier: 'BandesModule.tsx',
    codeRef: 'ligne ~711',
    statut: '❌ INCORRECT',
    original: 'Délégation pour la gestion des Activités Aéronautiques Nationales du Gabon\nSiège Social : Aéroport International Léon MBA, Libreville - B.P. : 2252 -Tél. : +241 65 04 28 62',
    genere:   'Délégation aux Activités Aéronautiques Nationales du Gabon — BP: 2252 Libreville, Tél.65.04.28.63',
    delta: [
      '"aux" → "pour la gestion des"',
      'Manque "Siège Social : Aéroport International Léon MBA, Libreville"',
      'Tél : "65.04.28.63" → "+241 65 04 28 62"',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 2 : NUMÉRO DE FACTURE
  // ─────────────────────────────────────────────────────────────────
  {
    id: 3,
    section: 'NUMÉRO DE FACTURE',
    ligne: 'L4 — Format du N° de facture',
    fichier: 'BandesModule.tsx',
    codeRef: 'ligne ~724',
    statut: '❌ FORMAT INCORRECT',
    original: 'N°2026/020/ASECNA/DGAN/CAF  (gras, aligné gauche, sans préfixe "Facture")',
    genere:   'Facture N°003  (avec préfixe "Facture")',
    delta: [
      'Format officiel : N°{ANNÉE}/{NUM_PADDED}/ASECNA/DGAN/CAF',
      'Supprimer le mot "Facture" devant le N°',
      'Numéro en gras avec le N° surligné/coloré (optionnel)',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 3 : BLOC COMPAGNIE
  // ─────────────────────────────────────────────────────────────────
  {
    id: 4,
    section: 'BLOC COMPAGNIE',
    ligne: 'L4-L6 — Bloc compagnie (coin haut-droite)',
    fichier: 'BandesModule.tsx',
    codeRef: 'lignes ~726-728',
    statut: '❌ INCOMPLET',
    original: 'Rectangle avec bordure contenant :\n  - Nom compagnie (GRAS + ITALIQUE)\n  - BP + Tél (gras italique)\n  - Ville (souligné)',
    genere:   'Juste ws.getCell("F4").value = facture.compagnie (pas de bordure, pas d\'adresse)',
    delta: [
      'Ajouter bordure autour du bloc compagnie (merge F4:I6 avec border)',
      'Ajouter adresse_compagnie (BP + Tél) en F5 (gras+italique)',
      'Ajouter ville_compagnie en F6 (souligné)',
      'Appliquer font italic + bold sur le nom compagnie',
    ],
    champsAjouter: [
      'adresse_compagnie: string  // ex: "BP:13025 Tél:011 44 40 15"',
      'ville_compagnie: string    // ex: "Oyem"',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 4 : SOUS-HEADERS DU TABLEAU
  // ─────────────────────────────────────────────────────────────────
  {
    id: 5,
    section: 'TABLEAU — SOUS-HEADERS',
    ligne: 'L10 — Nbre H / CU / Total 1 / Nbre Ann / CU / Total 2',
    fichier: 'BandesModule.tsx',
    codeRef: 'lignes ~786-793',
    statut: '❌ FOND GRIS MANQUANT',
    original: 'Fond gris (#D9D9D9) sur toute la ligne Nbre H | CU | Total 1 | Nbre Ann | CU | Total 2',
    genere:   'Pas de fond (variable "grey" définie mais non appliquée sur L10)',
    delta: [
      'Ajouter ws.getCell(col).fill = grey; pour chaque cellule B10 à G10',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 5 : TABLEAU — LIGNE VALEURS (annonces = 0)
  // ─────────────────────────────────────────────────────────────────
  {
    id: 6,
    section: 'TABLEAU — VALEURS',
    ligne: 'L11 — Annonces point "I" quand = 0',
    fichier: 'BandesModule.tsx',
    codeRef: 'lignes ~801-803',
    statut: '⚠️ AFFICHAGE CARET',
    original: 'Affiche " - " (tiret) quand nombre_annonces = 0 et CU annonces = "-"',
    genere:   'Affiche la valeur numérique 0 et 3500',
    delta: [
      'Si nombre_annonces === 0, afficher "-" au lieu de 0 pour E11',
      'Si nombre_annonces === 0, afficher "-" au lieu de 3500 pour F11',
      'Si nombre_annonces === 0, afficher "-" pour G11 (Total 2)',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 6 : SIGNATURE
  // ─────────────────────────────────────────────────────────────────
  {
    id: 7,
    section: 'SIGNATURE',
    ligne: 'L21-L24 — Bloc signature',
    fichier: 'BandesModule.tsx',
    codeRef: 'lignes ~886-889',
    statut: '⚠️ ALIGNEMENT',
    original: 'Centré sur la moitié droite de la page, texte en gras\n  "Le Délégué du Directeur Général de l\'ASECNA"\n  "pour la Gestion des Activités Aéronautiques"\n  "Nationales du Gabon par intérim"\n  [espace]\n  "Brice Thierry Arist SOKI" (gras, souligné)',
    genere:   'Colonne D uniquement, pas de soulignement sur SOKI',
    delta: [
      'Aligner la signature à droite (colonne E ou F)',
      'Ajouter alignment: { horizontal: "center" } sur les lignes signature',
      'Ajouter underline: true sur "Brice Thierry Arist SOKI"',
    ],
  },

  // ─────────────────────────────────────────────────────────────────
  // SECTION 7 : DONNÉES MANQUANTES
  // ─────────────────────────────────────────────────────────────────
  {
    id: 8,
    section: 'DONNÉES',
    ligne: 'Interface Facture — champs manquants',
    fichier: 'BandesModule.tsx',
    codeRef: 'lignes ~44-66 (interface Facture)',
    statut: '❌ CHAMPS MANQUANTS',
    original: 'Bloc compagnie contient : nom + adresse (BP+Tél) + ville',
    genere:   'Interface Facture n\'a pas adresse_compagnie ni ville_compagnie',
    delta: [
      'Ajouter dans interface Facture : adresse_compagnie?: string',
      'Ajouter dans interface Facture : ville_compagnie?: string',
      'Ajouter dans FacturationForm : champs de saisie pour ces 2 nouveaux champs',
      'Ajouter dans le POST /api/factures-bandes (server) : ces 2 colonnes',
    ],
  },
];

// ─────────────────────────────────────────────────────────────────
// RÉCAPITULATIF
// ─────────────────────────────────────────────────────────────────
const RESUME = {
  total: DIFF.length,
  erreurs_critiques: DIFF.filter(d => d.statut.startsWith('❌')).length,
  avertissements:    DIFF.filter(d => d.statut.startsWith('⚠️')).length,
  fichiersPrincipaux: ['BandesModule.tsx', 'server/standalone-commonjs.js'],
};

// ─────────────────────────────────────────────────────────────────
// AFFICHAGE
// ─────────────────────────────────────────────────────────────────
console.log('\n══════════════════════════════════════════════════════');
console.log('   ANALYSE DIFF — Facture Bandes (PDF vs Généré)');
console.log('══════════════════════════════════════════════════════\n');

for (const d of DIFF) {
  console.log(`┌── [#${d.id}] ${d.statut}  |  ${d.section}`);
  console.log(`│   Ligne     : ${d.ligne}`);
  console.log(`│   Fichier   : ${d.fichier} (${d.codeRef})`);
  console.log(`│   Original  : ${d.original.split('\n').join('\n│              ')}`);
  console.log(`│   Généré    : ${d.genere}`);
  console.log(`│   À changer :`);
  for (const delta of d.delta) console.log(`│     → ${delta}`);
  if (d.champsAjouter) {
    console.log(`│   Champs à ajouter :`);
    for (const c of d.champsAjouter) console.log(`│     + ${c}`);
  }
  console.log('└' + '─'.repeat(56) + '\n');
}

console.log('══════════════════════════════════════════════════════');
console.log('   RÉCAPITULATIF');
console.log('══════════════════════════════════════════════════════');
console.log(`  Total points : ${RESUME.total}`);
console.log(`  ❌ Erreurs critiques : ${RESUME.erreurs_critiques}`);
console.log(`  ⚠️  Avertissements   : ${RESUME.avertissements}`);
console.log(`  Fichiers impactés   : ${RESUME.fichiersPrincipaux.join(', ')}`);
console.log('\n  ORDRE DE CORRECTION RECOMMANDÉ :');
console.log('  1. Textes en-tête (L1, L2)          → changement de string, 2 min');
console.log('  2. Format N° facture                 → format string, 2 min');
console.log('  3. Fond gris sous-headers            → .fill = grey sur B10-G10, 1 min');
console.log('  4. Affichage "-" si annonces = 0     → condition JS, 3 min');
console.log('  5. Soulignement signature SOKI       → underline: true, 1 min');
console.log('  6. Bloc compagnie (adresse + bordure)→ interface + UI + Excel, 15 min');
console.log('══════════════════════════════════════════════════════\n');
