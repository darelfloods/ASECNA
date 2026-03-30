import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';

export interface FicheMissionData {
  nom: string;
  matricule: string;
  prenom: string;
  emploi: string;
  residence: string;
  destination: string;
  motif: string;
  dateDepart: string;
  dateRetour: string;
  duree: string;
  transport: string;
  templateType?: 'original' | 'ndong'; // Type de template détecté
}

export interface OrdreMissionData extends FicheMissionData {
  // Champs budgétaires
  cs: string;          // CS : 050
  eng: string;         // Eng : 13
  cr: string;          // CR : EN3
  cc: string;          // CC : 100
  cl: string;          // CL : 621

  // Montants financiers
  autorisationDep: string;   // Autorisation de dép : 16 250 000
  montantEngage: string;     // Montant engagé : 128 000
  engagementAnt: string;     // Engagement Ant : 15 879 600
  disponible: string;        // Disponible : 242 400

  // Lieu et date de signature
  lieuSignature: string;     // Libreville
  dateSignature: string;     // 06/02/2025
}

/**
 * Détecte le type de template utilisé dans le fichier Word
 */
function detectTemplateType(text: string): 'original' | 'ndong' {
  // Le template NDONG a "Monsieur :" ou "Madame :" et "Est autorisé à se rendre à"
  // Le template original a "Nom :" et "Prénom :" séparés et "Se rendra à"
  
  const hasMonsieurMadame = /(?:Monsieur|Madame)\s*:/i.test(text);
  const hasEstAutorise = /Est autorisé à se rendre à/i.test(text);
  const hasFonction = /Fonction\s*:/i.test(text);
  const hasDateDepartPrevue = /Date de départ prévue/i.test(text);
  
  // Si on a au moins 2 de ces indicateurs, c'est le template NDONG
  const ndongScore = [hasMonsieurMadame, hasEstAutorise, hasFonction, hasDateDepartPrevue].filter(Boolean).length;
  
  console.log('Détection de template:', {
    hasMonsieurMadame,
    hasEstAutorise,
    hasFonction,
    hasDateDepartPrevue,
    ndongScore,
    type: ndongScore >= 2 ? 'ndong' : 'original'
  });
  
  return ndongScore >= 2 ? 'ndong' : 'original';
}

/**
 * Parse un fichier Word FICHE DE MISSION et extrait les données
 */
export async function parseFicheMission(file: File): Promise<FicheMissionData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const content = e.target?.result as ArrayBuffer;
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip);

        // Lire le XML directement pour conserver les retours à la ligne (paragraphes et cellules de tableau)
        const xml = zip.file('word/document.xml')?.asText() || '';

        // Remplacer les fins de paragraphes, de lignes de tableau et de cellules par des sauts de ligne
        let textWithNewlines = xml
          .replace(/<w:p[^>]*>/g, '\n')
          .replace(/<\/w:p>/g, '\n')
          .replace(/<w:br[^>]*>/g, '\n')
          .replace(/<\/w:tc>/g, '\n')
          .replace(/<[^>]+>/g, '');

        // Nettoyer les espaces multiples et sauts de ligne multiples
        textWithNewlines = textWithNewlines.replace(/\n\s*\n/g, '\n');

        // Remplacer les entités HTML (si présentes)
        const text = textWithNewlines
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .trim();

        // Console log pour debug
        console.log("Texte extrait avec retours à la ligne :\n", text);

        // Détecter le type de template
        const templateType = detectTemplateType(text);
        console.log('Type de template détecté:', templateType);

        // Parser selon le type de template
        let data: FicheMissionData;
        
        if (templateType === 'ndong') {
          // Format NDONG: "Monsieur : NDONG Robert"
          const nomComplet = extractValue(text, /(?:Monsieur|Madame)\s*:\s*([^\n]+)/i) || '';
          const nomParts = nomComplet.trim().split(/\s+/);
          const nom = nomParts.length > 0 ? nomParts[0] : '';
          const prenom = nomParts.length > 1 ? nomParts.slice(1).join(' ') : '';
          
          data = {
            nom: nom,
            matricule: extractValue(text, /Matricule\s*:\s*(\d+)/i) || '',
            prenom: prenom,
            emploi: extractValue(text, /Fonction\s*:\s*([^\n]+)/i) || '',
            residence: extractValue(text, /Résidence Administrative\s*:\s*([^\n]+)/i) || '',
            destination: extractValue(text, /Est autorisé à se rendre à\s*:\s*([^\n]+)/i) || '',
            motif: extractValue(text, /Motif de la mission\s*([^\n]+)/i) || '',
            dateDepart: extractValue(text, /Date de départ prévue\s*:\s*(\d{2}\/\d{2}\/\d{4})/i) || '',
            dateRetour: extractValue(text, /Date de retour prévue\s*:\s*(\d{2}\/\d{2}\/\d{4})/i) || '',
            duree: extractValue(text, /Durée du séjour\s*:\s*(\d+)\s*jours/i) || '',
            transport: extractValue(text, /Moyen de transport utilisé\s*:\s*([^\n]+)/i) || '',
            templateType: 'ndong'
          };
        } else {
          // Format original: "Nom :" et "Prénom :" séparés
          data = {
            nom: extractValue(text, /Nom\s*:\s*([^,\n]+)/i) || '',
            matricule: extractValue(text, /matricule\s*:\s*(\d+)/i) || '',
            prenom: extractValue(text, /Prénom\s*:\s*([^\n]+)/i) || '',
            emploi: extractValue(text, /Emploi\s*:\s*([^\n]+)/i) || '',
            residence: extractValue(text, /Résidence Administrative\s*:\s*([^\n]+)/i) || '',
            destination: extractValue(text, /Se rendra à\s*:\s*([^\n]+)/i) || '',
            motif: extractValue(text, /Motif du déplacement\s*:\s*([^\n]+)/i) || '',
            dateDepart: extractValue(text, /Date de départ\s*:\s*(\d{2}\/\d{2}\/\d{4})/i) || '',
            dateRetour: extractValue(text, /Retour\s*:\s*(\d{2}\/\d{2}\/\d{4})/i) || '',
            duree: extractValue(text, /Durée prévue\s*:\s*(\d+)\s*Jours/i) || '',
            transport: extractValue(text, /Moyen de transport\s*:\s*([^\n]+)/i) || '',
            templateType: 'original'
          };
        }

        resolve(data);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erreur de lecture du fichier'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Extrait une valeur avec une regex
 */
function extractValue(text: string, pattern: RegExp): string | null {
  const match = text.match(pattern);
  if (match && match[1]) {
    return match[1].trim();
  }
  return null;
}

// ─── BON DE COMMANDE ──────────────────────────────────────────────────────────

export interface BonCommandeData {
  // Codes comptables
  cs: string;
  cr: string;
  cc: string;
  article: string;
  exercice: string;

  // Fournisseur
  fournisseurNom: string;      // Remplace "________________________" sur la ligne 1
  fournisseurAdresse1: string; // Remplace la 1ère ligne "________________________________________________"
  fournisseurAdresse2: string; // Remplace la 2ème ligne "________________________________________________"
  codeFournisseur: string;

  // Lignes de commande (jusqu'à 3)
  lignes: {
    description: string;
    quantite: string;
    prixUnitaire: string;
    total: string;
  }[];

  // Totaux
  montantTotalChiffres: string;
  montantTotalLettres: string;
  delaiLivraison: string;

  // Validation
  lieu: string;
  date: string;
  numeroEngagement: string;
  operation: string;
  numeroSerie: string;

  // Bon d'engagement (coupon interne — ne pas envoyer au fournisseur)
  numeroBon: string;             // Numéro du bon de commande (ex: BC-2026-0028)
  codeIndividuel: string;        // CODE INDIVIDUEL
  compteLimitatif: string;       // COMPTE LIMITATIF
  operationEquipement: string;   // OPERATION D'EQUIPEMENT
  compteDe: string;              // COMPTE DE / COMPTABILITE GENERALE
  montantAD: string;             // MONTANT A.D. (autorisation de dépenses)
  engagementsAnterieurs: string; // ENGAGEMENTS ANTERIEURS
}

/**
 * Convertit le format de page B4 paysage → A4 paysage avec mise à l'échelle proportionnelle.
 * Ratio : 16838 / 20636 ≈ 0.8158
 */
/**
 * Génère un BON DE COMMANDE en remplissant le template Word
 */
export async function generateBonCommande(data: BonCommandeData): Promise<Blob> {
  const response = await fetch('/BON DE COMMANDE A4.docx');
  if (!response.ok) throw new Error('Template BON DE COMMANDE A4.docx introuvable dans /public');
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);
  let documentXml = zip.file('word/document.xml')!.asText();

  // On conserve le format B4 paysage d'origine du template (pas de conversion A4)

  const BOLD_RPR = '<w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/></w:rPr>';

  // ── Injecteur de texte dans un paragraphe vide via paraId ─────────────────
  const inject = (paraId: string, text: string) => {
    if (!text) return;
    const safe = text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    const re = new RegExp(
      `(<w:p[^>]*w14:paraId="${paraId}"[^>]*>)([\\s\\S]*?)(<\\/w:p>)`,
      'g'
    );
    documentXml = documentXml.replace(re, `$1$2<w:r>${BOLD_RPR}<w:t xml:space="preserve">${safe}</w:t></w:r>$3`);
  };

  // ── Injecteur pour cases à 1 caractère (préserve l'indent, centre le texte) ─
  const injectChar = (paraId: string, char: string) => {
    if (!char?.trim()) return;
    const safe = char.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const re = new RegExp(
      `(<w:p[^>]*w14:paraId="${paraId}"[^>]*>)([\\s\\S]*?)(<\\/w:p>)`,
      'g'
    );
    documentXml = documentXml.replace(re, (_m: string, open: string, inner: string, close: string) => {
      // Préserver TOUTE la pPr existante (avec w:ind) — juste ajouter w:jc center si absent
      const fixed = inner.includes('<w:jc ')
        ? inner
        : inner.replace(/<\/w:pPr>/, '<w:jc w:val="center"/></w:pPr>');
      return `${open}${fixed}<w:r>${BOLD_RPR}<w:t xml:space="preserve">${safe}</w:t></w:r>${close}`;
    });
  };

  // ── Injecteur centré (quantité / prix / total du tableau) ─────────────────
  const injectCentered = (paraId: string, text: string) => {
    if (!text) return;
    const safe = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const re = new RegExp(`(<w:p[^>]*w14:paraId="${paraId}"[^>]*>)([\\s\\S]*?)(<\\/w:p>)`, 'g');
    documentXml = documentXml.replace(re, (_m: string, open: string, inner: string, close: string) => {
      let body = inner;
      if (/<w:pPr[\s>]/.test(body)) {
        if (!body.includes('<w:jc ')) body = body.replace(/<\/w:pPr>/, '<w:jc w:val="center"/></w:pPr>');
      } else {
        body = '<w:pPr><w:jc w:val="center"/></w:pPr>' + body;
      }
      return `${open}${body}<w:r>${BOLD_RPR}<w:t xml:space="preserve">${safe}</w:t></w:r>${close}`;
    });
  };

  // ── Codes comptables — 1 caractère par case ─────────────────────────────────
  const spreadChars = (ids: string[], value: string) => {
    const chars = (value || '').split('');
    ids.forEach((id, i) => injectChar(id, chars[i] || ''));
  };

  spreadChars(['278F4863', '3ED8D4CF', '1F68D0AF'], data.cs);                    // CS  (3 cases)
  spreadChars(['76A09288', '0922550F', '0525D86D'], data.cr);                    // CR  (3 cases)
  spreadChars(['464D2853', '1910E1CE', '37B353A9'], data.cc);                    // CC  (3 cases)
  spreadChars(['2BCFE059', '60DE73FF', '5114EA48'], data.article);               // Article (3 cases — template v2)
  spreadChars(['2714A278', '5DA999ED', '4E830C1D', '14748769'], data.exercice);  // Exercice (4 cases)

  // ── Code fournisseur (4 cases) ─────────────────────────────────────────────
  spreadChars(['1A6809BF', '244A9EAF', '5DB65F30', '386E9672'], data.codeFournisseur);

  // ── Adresse fournisseur ─────────────────────────────────────────────────────
  // Remplace le contenu d'un paragraphe (tirets) par la valeur saisie
  const replaceParaContent = (paraId: string, newContent: string, bold = false) => {
    if (!newContent) return;
    const safe = newContent.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const rPr = bold ? BOLD_RPR : '';
    const re = new RegExp(`(<w:p[^>]*w14:paraId="${paraId}"[^>]*>)([\\s\\S]*?)(<\\/w:p>)`, 'g');
    documentXml = documentXml.replace(re, (_m: string, open: string, inner: string, close: string) => {
      const pPrMatch = inner.match(/(<w:pPr>[\s\S]*?<\/w:pPr>)/);
      const pPr = pPrMatch ? pPrMatch[1] : '';
      return `${open}${pPr}<w:r>${rPr}<w:t xml:space="preserve">${safe}</w:t></w:r>${close}`;
    });
  };
  // En-tête : label fixe (pas en gras)
  replaceParaContent('7215E09D', 'ADRESSE DU FOURNISSEUR :');
  // Lignes fournisseur : valeurs saisies en gras
  replaceParaContent('4C56EDD5', data.fournisseurNom,      true);
  replaceParaContent('3DEED587', data.fournisseurAdresse1, true);
  replaceParaContent('0E826B05', data.fournisseurAdresse2, true);

  // ── Lignes de commande (3 lignes dans le tableau DETAIL) ──────────────────
  const ligneIds = [
    { desc: '2B8EB34A', qty: '1C80DE40', prix: '16DAF1DA', total: '687CFF4D' },
    { desc: '274BF247', qty: '279A4B16', prix: '653F54CB', total: '49A4ECBA' },
    { desc: '2CF4905F', qty: '3D534AA1', prix: '3F62CC38', total: '4C414FFE' },
  ];
  const fmtNum = (v: string) => {
    const n = parseFloat((v || '').replace(/\s/g, '').replace(/,/g, '.'));
    return isNaN(n) ? v : n.toLocaleString('fr-FR');
  };
  data.lignes.slice(0, 3).forEach((ligne, i) => {
    const ids = ligneIds[i];
    inject(ids.desc, ligne.description);
    injectCentered(ids.qty,   fmtNum(ligne.quantite));
    injectCentered(ids.prix,  fmtNum(ligne.prixUnitaire));
    injectCentered(ids.total, fmtNum(ligne.total));
  });

  // ── Montant total en chiffres (cellule TOTAL de la ligne MONTANT TOTAL) ───
  injectCentered('3A6EA764', data.montantTotalChiffres);

  // ── Injecteur inline : coupe un run au niveau du placeholder et insère un run bold ──
  const injectInlineBold = (placeholder: string, value: string) => {
    if (!value) return;
    const safe = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // Le placeholder est en fin de <w:t> : on coupe le run ici et on ouvre un run bold
    documentXml = documentXml.replace(
      `${placeholder}</w:t>`,
      `</w:t></w:r><w:r>${BOLD_RPR}<w:t xml:space="preserve">${safe}</w:t>`
    );
  };

  // ── Montant total en lettres ───────────────────────────────────────────────
  injectInlineBold('\u2026'.repeat(59) + '...', data.montantTotalLettres);

  // ── Délai de livraison ─────────────────────────────────────────────────────
  injectInlineBold('\u2026'.repeat(23), data.delaiLivraison);

  // ── Lieu et date ──────────────────────────────────────────────────────────
  if (data.lieu || data.date) {
    const safeLieu = (data.lieu || '___________').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const safeDate = (data.date || '__________').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    documentXml = documentXml.replace(
      /A\u00A0: ___________ LE : __________<\/w:t>/g,
      `A\u00A0: </w:t></w:r><w:r>${BOLD_RPR}<w:t xml:space="preserve">${safeLieu}</w:t></w:r>` +
      `<w:r><w:t xml:space="preserve"> LE : </w:t></w:r><w:r>${BOLD_RPR}<w:t xml:space="preserve">${safeDate}</w:t>`
    );
  }

  // ── N° d'engagement : insérer entre " :" et "VISA..." dans le même paragraphe
  if (data.numeroEngagement) {
    const safe = data.numeroEngagement.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // Dans le para 111DD25B : Run0="N° D'ENGAGEMENT" Run1=" :" Run2="   VISA ET CACHET..."
    // On insère un run bold+bleu entre Run1 (2ème </w:t></w:r>) et Run2
    const engIdx = documentXml.indexOf("N\u00B0 D'ENGAGEMENT");
    if (engIdx >= 0) {
      const CLOSE_RUN = '</w:t></w:r>';
      const firstEnd = documentXml.indexOf(CLOSE_RUN, engIdx);
      if (firstEnd >= 0) {
        const secondEnd = documentXml.indexOf(CLOSE_RUN, firstEnd + CLOSE_RUN.length);
        if (secondEnd >= 0) {
          const insertAt = secondEnd + CLOSE_RUN.length;
          const boldRun = `<w:r>${BOLD_RPR}<w:t xml:space="preserve"> ${safe}</w:t></w:r>`;
          documentXml = documentXml.slice(0, insertAt) + boldRun + documentXml.slice(insertAt);
        }
      }
    }
  }

  // ── Opération : append bold run au paragraphe ──────────────────────────────
  if (data.operation) inject('16A69E91', ` ${data.operation}`);

  // ── BON D'ENGAGEMENT (coupon interne) ──────────────────────────────────────
  // Numéro du bon de commande dans le coupon (après le label "BON DE COMMANDE")
  if (data.numeroBon) inject('7D63CF0B', ` ${data.numeroBon}`);

  // Codes comptables (miroir du formulaire principal)
  spreadChars(['039DF871', '5E50C1CB', '5B27722B'], data.cs);  // CS
  spreadChars(['7BC61F09', '7AA19F9E', '25755A46'], data.cr);  // CR
  spreadChars(['7A0E7BE7', '18B438E4', '54F752ED'], data.cc);  // CC
  injectChar('4C691C60', (data.article || '').charAt(0));       // Article (1 case dans le coupon)

  // Fournisseur dans le coupon
  replaceParaContent('3793017D', `FOURNISSEUR : ${data.fournisseurNom}`, true);

  // Code individuel — cellule valeur standalone (23CB2604 uniquement)
  // NE PAS injecter dans 7E7C8CF6 : c'est la cellule fusionnée (gridSpan=9)
  // qui contient les labels NUMERO D'ENGAGEMENT etc.
  if (data.codeIndividuel) {
    inject('23CB2604', data.codeIndividuel);
  }

  // N° d'engagement (coupon)
  if (data.numeroEngagement) inject('1E4115B2', data.numeroEngagement);

  // Compte limitatif (coupon)
  if (data.compteLimitatif) inject('61CAD576', data.compteLimitatif);

  // Opération d'équipement
  if (data.operationEquipement) inject('63126403', data.operationEquipement);

  // Compte de / Comptabilité générale
  if (data.compteDe) inject('78CBC1A4', data.compteDe);

  // Montant A.D. (autorisation de dépenses) — cellule principale (439 twips)
  if (data.montantAD) injectCentered('3B39980F', data.montantAD);

  // Montant du bon (= montant total calculé) — cellule principale
  if (data.montantTotalChiffres) injectCentered('15911AAF', data.montantTotalChiffres);

  // Engagements antérieurs — cellule principale
  if (data.engagementsAnterieurs) injectCentered('6193C334', data.engagementsAnterieurs);

  // Lieu et date du BON D'ENGAGEMENT (format : "A __________ LE ")
  // La cellule 39D89933 a déjà un run avec ce texte — on remplace directement dans le XML
  if (data.lieu || data.date) {
    const safeLieuBE = (data.lieu || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const safeDateBE = (data.date || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // Le run contient exactement "A __________ LE " — on remplace le texte du run
    documentXml = documentXml.replace(
      'A __________ LE </w:t>',
      `A ${safeLieuBE} LE ${safeDateBE}</w:t>`
    );
  }

  // Cumul des engagements (calculé : montant bon + engagements antérieurs)
  const montantBonNum = parseFloat((data.montantTotalChiffres || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const engAntNum    = parseFloat((data.engagementsAnterieurs  || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
  const cumul        = montantBonNum + engAntNum;
  const disponible   = (parseFloat((data.montantAD || '').replace(/\s/g, '').replace(/,/g, '.')) || 0) - cumul;
  if (cumul > 0)     injectCentered('4C3A7A18', cumul.toLocaleString('fr-FR'));
  if (disponible)    injectCentered('1FE6910C', disponible.toLocaleString('fr-FR'));

  // ── Tampons numériques ─────────────────────────────────────────────────────
  if (data.numeroSerie) {
    const safeSerial = data.numeroSerie.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    // Paragraphe minimal (hauteur 0) pour ne pas perturber la mise en page
    const ZERO_PARA_OPEN = `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr><w:r><w:pict>`;

    // Tampon 1 : cercle avec N° de série — positions B4 d'origine
    const circularStamp =
      ZERO_PARA_OPEN +
      `<v:oval xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
      `style="position:absolute;left:0;text-align:left;margin-left:430pt;margin-top:680pt;width:67pt;height:67pt;` +
      `z-index:251659264;visibility:visible;mso-wrap-style:square;` +
      `mso-position-horizontal-relative:page;mso-position-vertical-relative:page" ` +
      `strokecolor="#CC0000" strokeweight="1.5pt" o:allowoverlap="t" filled="f">` +
      `<v:textbox inset="4pt,14pt,4pt,4pt" style="overflow:hidden">` +
      `<w:txbxContent>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="CC0000"/><w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr><w:t>N\u00B0 S\u00C9RIE</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="CC0000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">${safeSerial}</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox></v:oval>` +
      `</w:pict></w:r></w:p>`;

    // Tampon 2 : rectangle « APPROUVÉ » — positions B4 d'origine
    const approvedStamp =
      ZERO_PARA_OPEN +
      `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
      `style="position:absolute;left:0;text-align:left;margin-left:375pt;margin-top:755pt;width:132pt;height:46pt;` +
      `z-index:251659265;visibility:visible;mso-wrap-style:square;` +
      `mso-position-horizontal-relative:page;mso-position-vertical-relative:page;rotation:355" ` +
      `strokecolor="#CC0000" strokeweight="2.5pt" o:allowoverlap="t" filled="f">` +
      `<v:textbox inset="4pt,2pt,4pt,2pt" style="overflow:hidden">` +
      `<w:txbxContent>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="CC0000"/><w:sz w:val="46"/><w:szCs w:val="46"/></w:rPr><w:t>APPROUV\u00C9</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox></v:rect>` +
      `</w:pict></w:r></w:p>`;

    // Insérer les tampons AVANT le premier <w:sectPr> (saut de section)
    // pour rester dans la section BON DE COMMANDE et ne pas créer de page supplémentaire.
    const sectPrIdx = documentXml.indexOf('<w:sectPr');
    if (sectPrIdx >= 0) {
      // Trouver le début du paragraphe contenant ce sectPr
      const pBefore = Math.max(
        documentXml.lastIndexOf('<w:p ', sectPrIdx),
        documentXml.lastIndexOf('<w:p>', sectPrIdx)
      );
      const insertAt = pBefore > 0 ? pBefore : sectPrIdx;
      documentXml = documentXml.slice(0, insertAt) + circularStamp + approvedStamp + documentXml.slice(insertAt);
    } else {
      documentXml = documentXml.replace('</w:body>', circularStamp + approvedStamp + '</w:body>');
    }
  }

  zip.file('word/document.xml', documentXml);
  return zip.generate({ type: 'blob' });
}

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Génère un ORDRE DE MISSION à partir des données et du template
 */
export async function generateOrdreMission(
  data: OrdreMissionData,
  numeroOrdre: string
): Promise<Blob> {
  // Charger le template
  const response = await fetch('/ORDRE DE MISSION.docx');
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip);

  // Récupérer le texte original
  let text = doc.getFullText();

  // Remplacer les valeurs
  text = text.replace(/N°\d+\/\d+/, `N°${numeroOrdre}`);
  text = replaceValue(text, /Nom\s*:\s*[^\n]+/, `Nom : ${data.nom}`);
  text = replaceValue(text, /Prénom\s*:\s*[^\n]+/, `Prénom : ${data.prenom}`);
  text = replaceValue(text, /Emploi\s*:\s*[^\n]+/, `Emploi : ${data.emploi}`);
  text = replaceValue(text, /Résidence Administrative\s*:\s*[^\n]+/, `Résidence Administrative : ${data.residence}`);
  text = replaceValue(text, /Se rendra à\s*:\s*[^\n]+/, `Se rendra à : ${data.destination}`);
  text = replaceValue(text, /Motif du déplacement\s*:\s*[^\n]+/, `Motif du déplacement : ${data.motif}`);
  text = replaceValue(text, /Date de départ\s*:\s*\d{2}\/\d{2}\/\d{4}/, `Date de départ : ${data.dateDepart}`);
  text = replaceValue(text, /Retour\s*:\s*\d{2}\/\d{2}\/\d{4}/, `Retour : ${data.dateRetour}`);
  text = replaceValue(text, /Durée prévue\s*:\s*\d+\s*Jours/, `Durée prévue : ${data.duree} Jours`);
  text = replaceValue(text, /Moyen de transport\s*:\s*[^\n]+/, `Moyen de transport : ${data.transport}`);

  // Cette méthode ne marchera pas bien, il faut une approche différente
  // Je vais utiliser une méthode qui modifie directement le XML

  return generateWordFromTemplate(data, numeroOrdre);
}

function replaceValue(text: string, pattern: RegExp, replacement: string): string {
  return text.replace(pattern, replacement);
}

/**
 * Génère un Word en modifiant directement le XML du template
 */
async function generateWordFromTemplate(
  data: OrdreMissionData,
  numeroOrdre: string
): Promise<Blob> {
  const response = await fetch('/ORDRE DE MISSION.docx');
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);

  // Lire le document.xml
  let documentXml = zip.file('word/document.xml')!.asText();

  // Remplacer les valeurs personnelles dans le XML
  documentXml = replaceInXml(documentXml, 'BOUTE BOU KISSEMBA', data.nom, 'Nom');
  documentXml = replaceInXml(documentXml, 'Chérubin', data.prenom, 'Prénom');
  documentXml = replaceInXml(documentXml, 'Commandant d\u2019a\u00e9rodrome PI', data.emploi, 'Emploi');
  documentXml = replaceInXml(documentXml, 'Oyem', data.residence, 'Résidence');
  documentXml = replaceInXml(documentXml, 'Bitam', data.destination, 'Destination');
  documentXml = replaceInXml(documentXml, 'Obsèques de Feu ENGOUROU MOTO Sylvain, ex Observateur météo', data.motif, 'Motif');
  // Les dates nécessitent de faire gaffe à pas confondre la date de départ et la date signature, mais le template utilise 06/02/2025 pour les deux initialement
  // "Libreville 06/02/2025" sera traité à part pour la signature
  documentXml = replaceInXml(documentXml, 'Libreville 06/02/2025', `${data.lieuSignature} ${data.dateSignature}`, 'Lieu/Date Signature');

  documentXml = replaceInXml(documentXml, '06/02/2025', data.dateDepart, 'Date Départ');
  documentXml = replaceInXml(documentXml, '07/02/2025', data.dateRetour, 'Date Retour');
  documentXml = replaceInXml(documentXml, '02 Jours', `${data.duree} Jours`, 'Durée');
  documentXml = replaceInXml(documentXml, 'voiture', data.transport, 'Transport');
  documentXml = replaceInXml(documentXml, 'N°011/2025', `N°${numeroOrdre}`, 'Numéro Ordre');

  // Remplacer les champs budgétaires
  documentXml = replaceInXml(documentXml, 'CS : 050', `CS : ${data.cs}`, 'CS');
  documentXml = replaceInXml(documentXml, 'Eng : 13', `Eng : ${data.eng}`, 'Eng');
  documentXml = replaceInXml(documentXml, 'CR : EN3', `CR : ${data.cr}`, 'CR');
  documentXml = replaceInXml(documentXml, 'CC : 100', `CC : ${data.cc}`, 'CC');
  // Éviter de remplacer accidentellement si c'était fragmenté
  documentXml = replaceInXml(documentXml, 'CL : 621', `CL : ${data.cl}`, 'CL');

  // Remplacer les montants financiers
  documentXml = replaceInXml(documentXml, '16 250 000', data.autorisationDep.replace(/\s/g, ' '), 'Autorisation Dep');
  documentXml = replaceInXml(documentXml, '128 000', data.montantEngage.replace(/\s/g, ' '), 'Montant Engagé');
  documentXml = replaceInXml(documentXml, '15 879 600', data.engagementAnt.replace(/\s/g, ' '), 'Engagement Ant');
  documentXml = replaceInXml(documentXml, '242 400', data.disponible.replace(/\s/g, ' '), 'Disponible');

  // Uniformiser la taille de police de toutes les valeurs saisies (24 = 12pt)
  const valeursANormaliser = [
    data.nom, data.prenom, data.emploi, data.residence, data.destination,
    data.motif, data.dateDepart, data.dateRetour, `${data.duree} Jours`,
    data.transport, `N°${numeroOrdre}`,
    `CS : ${data.cs}`, `Eng : ${data.eng}`, `CR : ${data.cr}`,
    `CC : ${data.cc}`, `CL : ${data.cl}`,
    `${data.lieuSignature} ${data.dateSignature}`,
    data.autorisationDep.replace(/\s/g, ' '),
    data.montantEngage.replace(/\s/g, ' '),
    data.engagementAnt.replace(/\s/g, ' '),
    data.disponible.replace(/\s/g, ' '),
  ];
  for (const val of valeursANormaliser) {
    if (val && val.trim()) documentXml = normalizeRunFormat(documentXml, val.trim());
  }

  // Remettre le XML modifié dans le zip
  zip.file('word/document.xml', documentXml);

  // Générer le nouveau fichier Word
  const outputBlob = zip.generate({ type: 'blob' });
  return outputBlob;
}

/**
 * Génère une FICHE DE MISSION modifiée à partir des données
 * @param data - Données modifiées à insérer
 * @param originalFile - Fichier Word original importé (optionnel, utilise le template par défaut si absent)
 */
export async function generateFicheMission(
  data: FicheMissionData,
  _originalFile?: File  // conservé pour la compatibilité de signature, non utilisé comme base
): Promise<Blob> {
  const templateType = data.templateType || 'original';

  // Toujours utiliser le template officiel comme base (jamais le fichier importé)
  const templateFile = templateType === 'ndong' ? '/FICHE DE MISSION NDONG.docx' : '/FICHE DE MISSION.docx';
  const response = await fetch(templateFile);
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);

  // Lire le document.xml
  let documentXml = zip.file('word/document.xml')!.asText();

  console.log('=== GÉNÉRATION FICHE DE MISSION ===');
  console.log('Type de template:', templateType);
  console.log('Données à insérer:', data);

  if (templateType === 'ndong') {
    // Format NDONG : extraction dynamique des valeurs courantes puis remplacement
    let cleanText = documentXml
      .replace(/<\/w:p>/g, '\n')
      .replace(/<w:br[^>]*>/g, '\n')
      .replace(/<\/w:tr>/g, '\n')
      .replace(/<\/w:tc>/g, '\t')
      .replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, '$1')
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&apos;/g, "'").replace(/&quot;/g, '"')
      .replace(/[ \t]+/g, ' ').replace(/\n\s*\n/g, '\n').trim();

    const extractValue = (pattern: RegExp, defaultValue: string): string => {
      const match = cleanText.match(pattern);
      return match ? match[1].trim() : defaultValue;
    };

    const currentNomComplet = extractValue(/(?:Monsieur|Madame)\s*:\s*([^\n]+?)(?=\s*Matricule|\n)/i, 'NDONG Robert');
    const parts = currentNomComplet.trim().split(/\s+/);
    const currentNom = parts[0] || 'NDONG';
    const currentPrenom = parts.slice(1).join(' ') || 'Robert';
    const currentMatricule = extractValue(/Matricule\s*:\s*(\d+)/i, '250356');
    const currentEmploi = extractValue(/Fonction\s*:\s*([^\n]+?)(?=\s*Résidence|\n)/i, 'Chargé Administration et des Finances Pi');
    const currentResidence = extractValue(/Résidence Administrative\s*:\s*([^\n]+?)(?=\s*Est autorisé|\n)/i, 'Libreville');
    const currentDestination = extractValue(/Est autorisé à se rendre à\s*:\s*([^\n]+?)(?=\s*Motif|\n)/i, 'Port-Gentil');
    const currentMotif = extractValue(/Motif de la mission\s*([^\n]+?)(?=\s*Date de départ|\n)/i, "Demande d'occupation de la société SGEPP");
    const currentDateDepart = extractValue(/Date de départ prévue\s*:\s*(\d{2}\/\d{2}\/\d{4})/i, '');
    const currentDateRetour = extractValue(/Date de retour prévue\s*:\s*(\d{2}\/\d{2}\/\d{4})/i, '');
    const currentDuree = extractValue(/Durée du séjour\s*:\s*(\d+\s*jours)/i, '03 jours');
    const currentTransport = extractValue(/Moyen de transport utilisé\s*:\s*([^\n]+?)(?=\s*Mention|\n|$)/i, 'Avion');

    const safeReplace = (oldVal: string, newVal: string, desc: string) => {
      documentXml = replaceInXml(documentXml, oldVal, newVal, desc);
    };

    safeReplace(`${currentNom} ${currentPrenom}`, `${data.nom} ${data.prenom}`, 'Nom complet');
    safeReplace(currentMatricule, data.matricule, 'Matricule');
    safeReplace(currentEmploi, data.emploi, 'Fonction');
    safeReplace(currentResidence, data.residence, 'Résidence');
    safeReplace(currentDestination, data.destination, 'Destination');
    safeReplace(currentMotif, data.motif, 'Motif');

    if (currentDateDepart) {
      safeReplace(currentDateDepart, data.dateDepart, 'Date départ');
    } else {
      documentXml = insertValueAfterLabel(documentXml, 'Date de départ prévue', data.dateDepart, 'Date départ');
    }
    if (currentDateRetour) {
      safeReplace(currentDateRetour, data.dateRetour, 'Date retour');
    } else {
      documentXml = insertValueAfterLabel(documentXml, 'Date de retour prévue', data.dateRetour, 'Date retour');
    }

    safeReplace(currentDuree, `${data.duree} jours`, 'Durée');
    safeReplace(currentTransport, data.transport, 'Transport');

  } else {
    // Format original : remplacement direct des valeurs connues du template
    // Nom, Prénom, Matricule sur lignes séparées
    documentXml = replaceInXml(documentXml, 'NDONG', data.nom, 'Nom');
    documentXml = replaceInXml(documentXml, '250356', data.matricule, 'Matricule');
    documentXml = replaceInXml(documentXml, 'ROBERT', data.prenom, 'Prénom');
    // Emploi est vide dans le template, insertion après le label
    if (data.emploi) {
      documentXml = insertValueAfterLabel(documentXml, 'Emploi', data.emploi, 'Emploi');
    }
    // Préfixe ": " pour éviter de remplacer "Libreville" dans le footer
    documentXml = replaceInXml(documentXml, ': Libreville', `: ${data.residence}`, 'Résidence');
    documentXml = replaceInXml(documentXml, 'Port-Gentil', data.destination, 'Destination');
    documentXml = replaceInXml(documentXml, 'Contrôle des occupation domaniale', data.motif, 'Motif');
    // Date de départ et Retour sur lignes séparées
    documentXml = replaceInXml(documentXml, '26/03/2026', data.dateDepart, 'Date départ');
    documentXml = replaceInXml(documentXml, '30/03/2026', data.dateRetour, 'Date retour');
    documentXml = replaceInXml(documentXml, '5 Jours', `${data.duree} Jours`, 'Durée');
    documentXml = replaceInXml(documentXml, 'Avion', data.transport, 'Transport');

    // Normaliser le formatage (supprimer gras/italique, uniformiser la taille)
    for (const val of [data.nom, data.prenom, data.matricule, data.emploi, data.destination, data.motif, data.transport, data.dateDepart, data.dateRetour, `${data.duree} Jours`]) {
      if (val) documentXml = normalizeRunFormat(documentXml, val);
    }
  }

  // Remettre le XML modifié dans le zip
  zip.file('word/document.xml', documentXml);

  // Générer le nouveau fichier Word
  const outputBlob = zip.generate({ type: 'blob' });
  return outputBlob;
}

/**
 * Fonction pour insérer une valeur après un label dans le XML Word
 * Utilisée quand le champ est vide dans le document original
 * Recherche le pattern "Label :" et insère la valeur après les deux-points
 */
function insertValueAfterLabel(documentXml: string, label: string, value: string, description: string): string {
  if (!value) return documentXml;

  const safeValue = value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Stratégie 1: Chercher le label suivi de ":" dans un même élément <w:t> ou fragmenté
  // Pattern pour trouver "Date de départ prévue" potentiellement fragmenté
  const labelChars = label.split('');
  let patternSource = '';
  for (let i = 0; i < labelChars.length; i++) {
    const char = labelChars[i];
    if (/^\s$/.test(char)) {
      patternSource += '(?:\\s|<[^>]+>)*';
    } else {
      patternSource += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (i < labelChars.length - 1) {
        patternSource += '(?:<[^>]+>)*';
      }
    }
  }
  
  // Pattern complet: label + ":" + espaces optionnels + fin de w:t ou contenu vide
  // Capture le groupe pour pouvoir ajouter la valeur après
  const pattern1 = new RegExp(
    `(${patternSource}(?:<[^>]+>)*\\s*:(?:<[^>]+>)*\\s*)(<\/w:t>)`,
    'gi'
  );
  
  let changed = false;
  let resultXml = documentXml.replace(pattern1, (match, before, closeTag) => {
    changed = true;
    // Ajouter un espace et la valeur avant la fermeture </w:t>
    return `${before}  ${safeValue}${closeTag}`;
  });

  // Stratégie 2: Si pas trouvé, chercher le ":" seul après le label dans un élément séparé
  if (!changed) {
    // Chercher un pattern où ":" est dans son propre <w:t>:</w:t> après le label
    const pattern2 = new RegExp(
      `(${patternSource}(?:</w:t>)?(?:<[^>]*>)*<w:t[^>]*>\\s*:\\s*)(<\/w:t>)`,
      'gi'
    );
    
    resultXml = documentXml.replace(pattern2, (match, before, closeTag) => {
      changed = true;
      return `${before}  ${safeValue}${closeTag}`;
    });
  }

  // Stratégie 3: Chercher le label et ajouter après le prochain </w:t> qui suit le ":"
  if (!changed) {
    // Pattern plus permissif: trouve le label, puis cherche le premier </w:t> après un ":"
    const pattern3 = new RegExp(
      `(${patternSource}[^<]*:[^<]*)(<\/w:t>)`,
      'gi'
    );
    
    resultXml = documentXml.replace(pattern3, (match, before, closeTag) => {
      if (!changed) { // Ne remplacer que la première occurrence
        changed = true;
        return `${before}  ${safeValue}${closeTag}`;
      }
      return match;
    });
  }

  console.log(`${description} (insertion): "${label}" + "${value}" (changements: ${changed})`);
  return resultXml;
}

/**
 * Fonction utilitaire globale pour remplacer du texte en ignorant la fragmentation XML par Word
 */
function replaceInXml(documentXml: string, oldValue: string, newValue: string, description: string): string {
  if (!oldValue) return documentXml;

  const searchChars = oldValue.trim().split('');
  let patternSource = '';

  for (let i = 0; i < searchChars.length; i++) {
    const char = searchChars[i];
    if (/^\s$/.test(char)) {
      patternSource += '(?:\\s|<[^>]+>|&[^;]+;)*';
    } else {
      patternSource += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (i < searchChars.length - 1) {
        patternSource += '(?:\\s|<[^>]+>|&[^;]+;)*';
      }
    }
  }

  const regex = new RegExp(patternSource, 'g');

  const safeNewValue = newValue
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  let changed = false;
  const resultXml = documentXml.replace(regex, (match) => {
    changed = true;
    let result = '';
    let isInsideTag = false;
    let remainingReplace: string | null = safeNewValue;

    for (let i = 0; i < match.length; i++) {
      const char = match[i];
      if (char === '<') {
        isInsideTag = true;
        result += char;
      } else if (char === '>') {
        isInsideTag = false;
        result += char;
      } else if (!isInsideTag) {
        if (remainingReplace !== null) {
          result += remainingReplace;
          remainingReplace = null;
        }
      } else {
        result += char;
      }
    }

    if (remainingReplace !== null) {
      result += remainingReplace;
    }
    return result;
  });

  console.log(`${description}: "${oldValue}" -> "${newValue}" (changements: ${changed})`);
  return resultXml;
}

/**
 * Normalise le formatage d'un run Word contenant une valeur donnée :
 * supprime le gras et l'italique, fixe la taille de police à targetSize (24 = 12pt par défaut).
 *
 * Approche string-search : remonte depuis le <w:t> contenant la valeur jusqu'au <w:r> parent
 * pour éviter les faux positifs du regex qui peut sauter des éléments XML entiers.
 */
function normalizeRunFormat(documentXml: string, textValue: string, targetSize: string = '24'): string {
  if (!textValue || !textValue.trim()) return documentXml;

  const needle = textValue.trim();
  let result = documentXml;
  let searchFrom = 0;

  while (true) {
    // Trouver le prochain <w:t> contenant notre valeur
    const wtIdx = result.indexOf('>' + needle + '<', searchFrom);
    if (wtIdx < 0) break;

    // Vérifier que c'est bien dans un <w:t>
    const tStart = result.lastIndexOf('<w:t', wtIdx);
    const tClose = result.indexOf('</w:t>', wtIdx);
    if (tStart < 0 || tClose < 0) { searchFrom = wtIdx + 1; continue; }

    // Vérifier que le texte est entièrement contenu dans ce <w:t>
    const tTagClose = result.indexOf('>', tStart); // fin de <w:t ...>
    if (tTagClose < 0 || tTagClose > wtIdx) { searchFrom = wtIdx + 1; continue; }

    // Remonter au <w:r> parent : chercher le dernier <w:r> ou <w:r > avant tStart
    // en s'assurant que ce n'est pas <w:rPr>, <w:rFonts>, etc.
    let rStart = tStart;
    while (rStart > 0) {
      const prev = result.lastIndexOf('<w:r', rStart - 1);
      if (prev < 0) break;
      // Vérifier que le tag est bien <w:r> ou <w:r ...> (et pas <w:rPr>, <w:rFonts>, etc.)
      const charAfterR = result[prev + 4]; // char juste après '<w:r'
      if (charAfterR === '>' || charAfterR === ' ') {
        rStart = prev;
        break;
      }
      rStart = prev;
    }
    if (rStart <= 0) { searchFrom = wtIdx + 1; continue; }

    // Trouver la fin du tag <w:r ...>
    const rTagEnd = result.indexOf('>', rStart);
    if (rTagEnd < 0) { searchFrom = wtIdx + 1; continue; }

    // Trouver le </w:r> correspondant (immédiatement après le </w:t>)
    const rClose = result.indexOf('</w:r>', tClose) + 6;
    if (rClose < 6) { searchFrom = wtIdx + 1; continue; }

    // Extraire le contenu du run
    const runContent = result.slice(rStart, rClose);

    // Vérifier que le run ne contient pas d'éléments complexes (ex: <w:drawing>)
    if (runContent.includes('<w:drawing>') || runContent.includes('<w:drawing ')) {
      searchFrom = rClose;
      continue;
    }

    // Trouver le <w:rPr> du run en comptant les niveaux imbriqués
    const rPrStartIdx = runContent.indexOf('<w:rPr>');
    let newRun: string;

    if (rPrStartIdx >= 0) {
      // Trouver le </w:rPr> externe en comptant la profondeur d'imbrication
      let depth = 0;
      let rPrEndIdx = -1;
      for (let i = rPrStartIdx; i < runContent.length - 7; i++) {
        if (runContent.slice(i, i + 7) === '<w:rPr>') depth++;
        else if (runContent.slice(i, i + 8) === '</w:rPr>') {
          depth--;
          if (depth === 0) { rPrEndIdx = i + 8; break; }
        }
      }

      if (rPrEndIdx > 0) {
        const oldRPr = runContent.slice(rPrStartIdx, rPrEndIdx);
        const fontMatch = oldRPr.match(/<w:rFonts[^>]*\/>/);
        const fontTag = fontMatch ? fontMatch[0] : '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>';
        const newRPr = `<w:rPr>${fontTag}<w:b w:val="0"/><w:bCs w:val="0"/><w:sz w:val="${targetSize}"/><w:szCs w:val="${targetSize}"/></w:rPr>`;
        newRun = runContent.slice(0, rPrStartIdx) + newRPr + runContent.slice(rPrEndIdx);
      } else {
        newRun = runContent;
      }
    } else {
      // Pas de rPr : en ajouter un après le tag d'ouverture du run
      const insertAt = rTagEnd - rStart + 1;
      const newRPr = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b w:val="0"/><w:bCs w:val="0"/><w:sz w:val="${targetSize}"/><w:szCs w:val="${targetSize}"/></w:rPr>`;
      newRun = runContent.slice(0, insertAt) + newRPr + runContent.slice(insertAt);
    }

    result = result.slice(0, rStart) + newRun + result.slice(rClose);
    searchFrom = rStart + newRun.length;
    console.log(`normalizeRunFormat: formatage normalisé pour "${textValue}"`);
    break; // Une seule occurrence par valeur
  }

  return result;
}
