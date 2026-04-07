/**
 * Générateur de factures pour bandes d'enregistrement
 *
 * APPROCHE : JSZip uniquement — on injecte les données directement dans le XML
 * du template, sans passer par ExcelJS qui détruit le logo, l'en-tête et le VML.
 *
 * Structure du template Excel :
 *  - Feuille "BK"  → worksheets/sheet3.xml
 *  - Relations BK  → worksheets/_rels/sheet3.xml.rels
 *  - VML logo      → drawings/vmlDrawing3.vml  (image3.png via rId1)
 *  - PrintSettings → printerSettings/printerSettings3.bin
 */

export interface BandeFactureData {
  numero_facture: string;
  date_facture: string;       // ISO: YYYY-MM-DD
  compagnie: string;
  adresse_compagnie: string;
  ville_compagnie: string;
  site: string;
  serie_bandes: string;
  periode_debut: string;      // ISO: YYYY-MM-DD
  periode_fin: string;        // ISO: YYYY-MM-DD
  nombre_heures: number;
  tarif_horaire: number;
  total_heures: number;
  nombre_annonces: number;
  tarif_annonce: number;
  total_annonces: number;
  montant_ht: number;
  total_pax: number;
  taxes: number;
  acompte: number;
  solde: number;
  montant_en_lettres: string;
}

// ── Helpers de formatage ────────────────────────────────────────────────────

function fmtDateShort(iso: string): string {
  if (!iso) return '';
  const [, m, d] = iso.split('-');
  return `${d}/${m}`;
}

function fmtDateFull(iso: string): string {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function fmtDateLibreville(iso: string): string {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

/** Échappe les caractères spéciaux XML */
function escXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ── Manipulation des shared strings ─────────────────────────────────────────

/**
 * Ajoute une chaîne dans les sharedStrings et retourne son index.
 * Si la chaîne existe déjà, retourne l'index existant.
 */
function addSharedString(ssXml: string, value: string): { xml: string; index: number } {
  // Chercher si la valeur existe déjà
  const escapedValue = escXml(value);
  const searchPattern = `<si><t>${escapedValue}</t></si>`;

  // Compter les <si> existants pour déterminer le prochain index
  const allSi = [...ssXml.matchAll(/<si>/g)];
  const count = allSi.length;

  // Chercher si déjà présent
  const existingIdx = ssXml.indexOf(searchPattern);
  if (existingIdx !== -1) {
    // Trouver l'index dans les si
    const before = ssXml.substring(0, existingIdx);
    const idx = [...before.matchAll(/<si>/g)].length;
    return { xml: ssXml, index: idx };
  }

  // Ajouter la nouvelle chaîne avant </sst>
  const newSi = `<si><t>${escapedValue}</t></si>`;
  // Utiliser un lookbehind négatif pour ne pas matcher "uniqueCount" quand on veut "count"
  const newXml = ssXml
    .replace(/(<sst[^>]*\b)count="(\d+)"/, (m, prefix, countStr) => {
      const newCount = parseInt(countStr) + 1;
      return `${prefix}count="${newCount}"`;
    })
    .replace(/(<sst[^>]*)uniqueCount="(\d+)"/, (m, prefix, countStr) => {
      const newCount = parseInt(countStr) + 1;
      return `${prefix}uniqueCount="${newCount}"`;
    })
    .replace('</sst>', newSi + '</sst>');

  return { xml: newXml, index: count };
}

/**
 * Remplace la valeur d'une cellule de type string (t="s") dans le XML de la feuille.
 * La cellule garde son style (attribut s="...").
 */
function replaceCellString(sheetXml: string, cellRef: string, ssIndex: number): string {
  // Remplacer <c r="CELLREF" s="..." t="s"><v>OLDIDX</v></c>
  return sheetXml.replace(
    new RegExp(`(<c r="${cellRef}"[^>]*t="s"[^>]*>)<v>\\d+</v>(</c>)`, 'g'),
    `$1<v>${ssIndex}</v>$2`
  );
}

/**
 * Remplace la valeur d'une cellule numérique dans le XML de la feuille.
 * Supprime les formules et met la valeur directement.
 */
function replaceCellNumber(sheetXml: string, cellRef: string, value: number): string {
  // Cas 1 : cellule avec formule <f>...</f><v>...</v>
  let result = sheetXml.replace(
    new RegExp(`(<c r="${cellRef}"[^>]*>)<f>[^<]*</f><v>[^<]*</v>(</c>)`, 'g'),
    `$1<v>${value}</v>$2`
  );
  // Cas 2 : cellule numérique simple <v>...</v>
  result = result.replace(
    new RegExp(`(<c r="${cellRef}"(?![^>]*t=")[^>]*>)<v>[^<]*</v>(</c>)`, 'g'),
    `$1<v>${value}</v>$2`
  );
  return result;
}

// ── Génération de facture unique ─────────────────────────────────────────────

/**
 * Génère une facture unique en utilisant JSZip uniquement pour préserver
 * 100% du template (logo, en-tête, styles, VML).
 */
export async function generateSingleBandeInvoice(
  facture: BandeFactureData,
  templatePath: string
): Promise<ArrayBuffer> {
  // Charger le template
  const response = await fetch(encodeURI(templatePath));
  if (!response.ok) throw new Error(`Impossible de charger le template: ${response.status}`);
  const ct = response.headers.get('content-type') || '';
  if (ct.includes('text/html')) throw new Error('Le template retourné est du HTML, pas un fichier Excel.');
  const arrayBuffer = await response.arrayBuffer();

  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(arrayBuffer);

  // ── 1. Lire les fichiers nécessaires ──────────────────────────────────────
  const sheetFile    = zip.file('xl/worksheets/sheet3.xml')!;
  const ssFile       = zip.file('xl/sharedStrings.xml')!;
  const wbFile       = zip.file('xl/workbook.xml')!;
  const wbRelsFile   = zip.file('xl/_rels/workbook.xml.rels')!;
  const ctFile       = zip.file('[Content_Types].xml')!;

  let sheetXml = await sheetFile.async('string');
  let ssXml    = await ssFile.async('string');
  let wbXml    = await wbFile.async('string');
  let wbRels   = await wbRelsFile.async('string');
  let ctXml    = await ctFile.async('string');

  // ── 2. Préparer les valeurs ────────────────────────────────────────────────
  const periodeStr = (facture.periode_debut && facture.periode_fin)
    ? `Du ${fmtDateShort(facture.periode_debut)} au ${fmtDateFull(facture.periode_fin)}`
    : '';

  const dateLibreville = `Libreville, le ${fmtDateLibreville(facture.date_facture)}`;
  const numeroFacture  = `Facture N°${facture.numero_facture}`;
  const serieStr       = `Série N°:${facture.serie_bandes || ''}`;

  // ── 3. Ajouter toutes les valeurs dans sharedStrings ─────────────────────
  const strings: Record<string, string> = {
    B4:  numeroFacture,
    F7:  facture.compagnie,
    A8:  dateLibreville,
    F8:  facture.adresse_compagnie || '',
    F9:  facture.ville_compagnie || facture.site,
    C11: facture.site,
    B12: serieStr,
    A16: periodeStr,
    A27: facture.montant_en_lettres,
  };

  const ssIndices: Record<string, number> = {};
  for (const [cell, value] of Object.entries(strings)) {
    const result = addSharedString(ssXml, value);
    ssXml = result.xml;
    ssIndices[cell] = result.index;
  }

  // ── 4. Injecter les valeurs texte dans le XML de la feuille ──────────────
  for (const [cell, idx] of Object.entries(ssIndices)) {
    sheetXml = replaceCellString(sheetXml, cell, idx);
  }

  // ── 5. Injecter les valeurs numériques ────────────────────────────────────
  sheetXml = replaceCellNumber(sheetXml, 'B19', facture.nombre_heures);
  sheetXml = replaceCellNumber(sheetXml, 'C19', facture.tarif_horaire);
  sheetXml = replaceCellNumber(sheetXml, 'D19', facture.total_heures);
  sheetXml = replaceCellNumber(sheetXml, 'E19', facture.nombre_annonces);
  sheetXml = replaceCellNumber(sheetXml, 'F19', facture.nombre_annonces === 0 ? 0 : facture.tarif_annonce);
  sheetXml = replaceCellNumber(sheetXml, 'G19', facture.total_annonces);
  sheetXml = replaceCellNumber(sheetXml, 'H19', facture.montant_ht);
  sheetXml = replaceCellNumber(sheetXml, 'A24', facture.total_pax);
  sheetXml = replaceCellNumber(sheetXml, 'B24', facture.montant_ht);
  sheetXml = replaceCellNumber(sheetXml, 'D24', facture.taxes);
  sheetXml = replaceCellNumber(sheetXml, 'F24', facture.acompte);
  sheetXml = replaceCellNumber(sheetXml, 'G24', facture.montant_ht);
  sheetXml = replaceCellNumber(sheetXml, 'H24', facture.solde);

  // ── 6. Vider le bloc 2 (colonnes I-P) ────────────────────────────────────
  // On remplace les valeurs des cellules du bloc 2 par des chaînes vides
  const bloc2StringCells = ['J4','N7','I8','N8','N9','K11','J12','I16','I27'];
  for (const cell of bloc2StringCells) {
    const result = addSharedString(ssXml, '');
    ssXml = result.xml;
    sheetXml = replaceCellString(sheetXml, cell, result.index);
  }
  // Vider les cellules numériques du bloc 2
  const bloc2NumCells = ['J19','K19','L19','M19','N19','O19','P19','I24','J24','L24','N24','O24','P24'];
  for (const cell of bloc2NumCells) {
    sheetXml = replaceCellNumber(sheetXml, cell, 0);
  }

  // ── 7. Supprimer les feuilles inutiles et ne garder que BK ───────────────
  // Garder uniquement rId3 (BK = sheet3.xml)
  wbXml = wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/, 
    '<sheets><sheet name="BK" sheetId="4" r:id="rId3"/></sheets>'
  );

  // Supprimer les relations des autres feuilles du workbook
  // Garder : rId3 (sheet3=BK), rId7 (sharedStrings), styles, theme
  wbRels = wbRels
    .replace(/<Relationship[^>]*rId1[^>]*\/>/g, '')  // Bandes
    .replace(/<Relationship[^>]*rId2[^>]*\/>/g, '')  // Conventions
    .replace(/<Relationship[^>]*rId4[^>]*\/>/g, '')  // Recap
    .replace(/<Relationship[^>]*rId8[^>]*\/>/g, ''); // calcChain

  // ── 8. Mettre à jour le Content_Types si nécessaire ──────────────────────
  // S'assurer que les types VML et PNG sont présents
  if (!ctXml.includes('image/png')) {
    ctXml = ctXml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  if (!ctXml.includes('vmlDrawing')) {
    ctXml = ctXml.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>');
  }

  // ── 9. Supprimer les feuilles inutiles du ZIP ────────────────────────────
  // Supprimer sheet1.xml (Bandes), sheet2.xml (Conventions), sheet4.xml (Recap)
  // et leurs relations + calcChain qui référence des formules/feuilles supprimées
  ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet4.xml',
   'xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/_rels/sheet2.xml.rels',
   'xl/worksheets/_rels/sheet4.xml.rels',
   'xl/calcChain.xml'].forEach(f => {
    if (zip.file(f)) zip.remove(f);
  });

  // ── 10. Nettoyer Content_Types.xml : retirer les Override des fichiers supprimés
  ctXml = ctXml
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet1\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet2\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet4\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');

  // ── 11. Mettre à jour les fichiers dans le ZIP ───────────────────────────
  zip.file('xl/worksheets/sheet3.xml', sheetXml);
  zip.file('xl/sharedStrings.xml', ssXml);
  zip.file('xl/workbook.xml', wbXml);
  zip.file('xl/_rels/workbook.xml.rels', wbRels);
  zip.file('[Content_Types].xml', ctXml);

  console.log('✅ Facture générée avec préservation complète du template');
  return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}


// ── Génération multiple (groupée par site) ───────────────────────────────────

function normalizeSiteName(site: string): string {
  return (site || 'AUTRES')
    .toUpperCase()
    .replace(/[^A-Z0-9\-]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .substring(0, 31);
}

function groupFacturesBySite(factures: BandeFactureData[]): Map<string, BandeFactureData[]> {
  const groups = new Map<string, BandeFactureData[]>();
  for (const f of factures) {
    const site = f.site || 'AUTRES';
    if (!groups.has(site)) groups.set(site, []);
    groups.get(site)!.push(f);
  }
  return groups;
}

/**
 * Génère un seul fichier Excel avec plusieurs factures,
 * une feuille par site (classées par site).
 *
 * Approche : on part du template, on garde la feuille BK comme base pour
 * chaque site — on duplique sheet3.xml avec un nouveau nom de feuille pour
 * chaque site, puis on met à jour workbook.xml avec toutes les feuilles.
 */
export async function generateMultiBandesInvoices(
  factures: BandeFactureData[],
  templatePath: string
): Promise<ArrayBuffer> {
  const response = await fetch(encodeURI(templatePath));
  if (!response.ok) throw new Error(`Impossible de charger le template: ${response.status}`);
  const ct = response.headers.get('content-type') || '';
  if (ct.includes('text/html')) throw new Error('Le template retourné est du HTML, pas un fichier Excel.');
  const templateBuffer = await response.arrayBuffer();

  const JSZip = (await import('jszip')).default;

  // Grouper par site
  const facturesBySite = groupFacturesBySite(factures);
  const siteNames = Array.from(facturesBySite.keys()).sort((a, b) => {
    if (a === 'AUTRES') return 1;
    if (b === 'AUTRES') return -1;
    return a.localeCompare(b);
  });

  /**
   * Construire la liste des feuilles à générer :
   * - Si un site a 1 seule facture → 1 feuille nommée "SITE"
   * - Si un site a plusieurs factures → 1 feuille par compagnie nommée "SITE-COMPAGNIE"
   *   (si 2 compagnies identiques sur le même site → on ajoute un numéro)
   */
  interface SheetEntry { sheetLabel: string; facture: BandeFactureData; }
  const sheetsToGenerate: SheetEntry[] = [];

  for (const siteName of siteNames) {
    const siteFactures = facturesBySite.get(siteName)!;

    if (siteFactures.length === 1) {
      // Cas simple : 1 facture → 1 feuille nommée par le site
      sheetsToGenerate.push({
        sheetLabel: normalizeSiteName(siteName).substring(0, 31),
        facture: siteFactures[0],
      });
    } else {
      // Plusieurs factures → regrouper par compagnie dans ce site
      const byCompagnie = new Map<string, BandeFactureData[]>();
      for (const f of siteFactures) {
        const key = f.compagnie || 'AUTRES';
        if (!byCompagnie.has(key)) byCompagnie.set(key, []);
        byCompagnie.get(key)!.push(f);
      }

      const compagnieNames = Array.from(byCompagnie.keys()).sort((a, b) => {
        if (a === 'AUTRES') return 1;
        if (b === 'AUTRES') return -1;
        return a.localeCompare(b);
      });

      for (const compagnieName of compagnieNames) {
        const compagnieFactures = byCompagnie.get(compagnieName)!;
        // Nom de feuille : "SITE-COMPAGNIE" tronqué à 31 chars
        const siteNorm = normalizeSiteName(siteName);
        const compNorm = normalizeSiteName(compagnieName);
        const label = `${siteNorm}-${compNorm}`.substring(0, 31);

        // S'il y a plusieurs factures pour la même compagnie sur ce site, on prend la première
        // (cas très rare — on pourrait les fusionner dans le futur)
        sheetsToGenerate.push({ sheetLabel: label, facture: compagnieFactures[0] });
      }
    }
  }

  // Dédupliquer les noms de feuilles si besoin
  const usedLabels = new Map<string, number>();
  for (const entry of sheetsToGenerate) {
    const base = entry.sheetLabel;
    if (usedLabels.has(base)) {
      const count = usedLabels.get(base)! + 1;
      usedLabels.set(base, count);
      entry.sheetLabel = base.substring(0, 28) + `-${count}`;
    } else {
      usedLabels.set(base, 1);
    }
  }

  console.log(`Génération: ${sheetsToGenerate.map(s => s.sheetLabel).join(', ')}`);

  // On charge le template une seule fois pour avoir la base
  const zip = await JSZip.loadAsync(templateBuffer);

  // Lire les fichiers partagés du template
  let ssXml  = await zip.file('xl/sharedStrings.xml')!.async('string');
  let wbXml  = await zip.file('xl/workbook.xml')!.async('string');
  let wbRels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  let ctXml  = await zip.file('[Content_Types].xml')!.async('string');

  // XML de la feuille BK originale (sheet3.xml) — servira de template pour chaque feuille
  const bkTemplateXml = await zip.file('xl/worksheets/sheet3.xml')!.async('string');

  // Supprimer les feuilles inutiles (Bandes, Conventions, Recap) + calcChain du ZIP
  ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet4.xml',
   'xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/_rels/sheet2.xml.rels',
   'xl/worksheets/_rels/sheet4.xml.rels',
   'xl/calcChain.xml'].forEach(f => {
    if (zip.file(f)) zip.remove(f);
  });

  // Nettoyer Content_Types.xml des feuilles supprimées
  ctXml = ctXml
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet1\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet2\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet4\.xml"[^>]*\/>/g, '')
    .replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');

  // Préparer les entrées pour workbook.xml et workbook.xml.rels
  const sheetEntries: string[] = [];
  const relsEntries: string[]  = [];

  let sheetIdCounter = 10;
  let rIdCounter     = 10;

  for (const { sheetLabel, facture } of sheetsToGenerate) {
    // Copier le XML du template BK et injecter les données
    let sheetXml = bkTemplateXml;

    // Préparer les valeurs texte
    const periodeStr     = (facture.periode_debut && facture.periode_fin)
      ? `Du ${fmtDateShort(facture.periode_debut)} au ${fmtDateFull(facture.periode_fin)}` : '';
    const dateLibreville = `Libreville, le ${fmtDateLibreville(facture.date_facture)}`;
    const numeroFacture  = `Facture N°${facture.numero_facture}`;
    const serieStr       = `Série N°:${facture.serie_bandes || ''}`;

    const strings: Record<string, string> = {
      B4:  numeroFacture,
      F7:  facture.compagnie,
      A8:  dateLibreville,
      F8:  facture.adresse_compagnie || '',
      F9:  facture.ville_compagnie || facture.site,
      C11: facture.site,
      B12: serieStr,
      A16: periodeStr,
      A27: facture.montant_en_lettres,
    };

    // Ajouter dans sharedStrings et injecter dans le XML de la feuille
    for (const [cell, value] of Object.entries(strings)) {
      const result = addSharedString(ssXml, value);
      ssXml = result.xml;
      sheetXml = replaceCellString(sheetXml, cell, result.index);
    }

    // Valeurs numériques
    sheetXml = replaceCellNumber(sheetXml, 'B19', facture.nombre_heures);
    sheetXml = replaceCellNumber(sheetXml, 'C19', facture.tarif_horaire);
    sheetXml = replaceCellNumber(sheetXml, 'D19', facture.total_heures);
    sheetXml = replaceCellNumber(sheetXml, 'E19', facture.nombre_annonces);
    sheetXml = replaceCellNumber(sheetXml, 'F19', facture.nombre_annonces === 0 ? 0 : facture.tarif_annonce);
    sheetXml = replaceCellNumber(sheetXml, 'G19', facture.total_annonces);
    sheetXml = replaceCellNumber(sheetXml, 'H19', facture.montant_ht);
    sheetXml = replaceCellNumber(sheetXml, 'A24', facture.total_pax);
    sheetXml = replaceCellNumber(sheetXml, 'B24', facture.montant_ht);
    sheetXml = replaceCellNumber(sheetXml, 'D24', facture.taxes);
    sheetXml = replaceCellNumber(sheetXml, 'F24', facture.acompte);
    sheetXml = replaceCellNumber(sheetXml, 'G24', facture.montant_ht);
    sheetXml = replaceCellNumber(sheetXml, 'H24', facture.solde);

    // Vider le bloc 2 (colonnes I-P)
    const bloc2StringCells = ['J4','N7','I8','N8','N9','K11','J12','I16','I27'];
    for (const cell of bloc2StringCells) {
      const result = addSharedString(ssXml, '');
      ssXml = result.xml;
      sheetXml = replaceCellString(sheetXml, cell, result.index);
    }
    ['J19','K19','L19','M19','N19','O19','P19','I24','J24','L24','N24','O24','P24'].forEach(c => {
      sheetXml = replaceCellNumber(sheetXml, c, 0);
    });

    // Nom de la feuille dans le fichier : sheetNN.xml
    const sheetFileName = `sheet${sheetIdCounter}.xml`;
    const rId           = `rId${rIdCounter}`;
    const sheetName     = sheetLabel;

    // Ajouter la feuille dans le ZIP
    zip.file(`xl/worksheets/${sheetFileName}`, sheetXml);

    // Copier les relations de BK (sheet3.xml.rels) pour cette nouvelle feuille
    const bkRels = await zip.file('xl/worksheets/_rels/sheet3.xml.rels')?.async('string');
    if (bkRels) {
      zip.file(`xl/worksheets/_rels/${sheetFileName}.rels`, bkRels);
    }

    // Préparer les entrées workbook
    sheetEntries.push(`<sheet name="${escXml(sheetName)}" sheetId="${sheetIdCounter}" r:id="${rId}"/>`);
    relsEntries.push(
      `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetFileName}"/>`
    );

    // Ajouter le PartName de cette feuille dans Content_Types
    if (!ctXml.includes(`/xl/worksheets/${sheetFileName}`)) {
      ctXml = ctXml.replace('</Types>',
        `<Override PartName="/xl/worksheets/${sheetFileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
      );
    }

    console.log(`✅ Feuille "${sheetName}" créée`);
    sheetIdCounter++;
    rIdCounter++;
  }

  // Supprimer la feuille BK originale du ZIP (on a créé des copies)
  if (zip.file('xl/worksheets/sheet3.xml')) zip.remove('xl/worksheets/sheet3.xml');
  if (zip.file('xl/worksheets/_rels/sheet3.xml.rels')) zip.remove('xl/worksheets/_rels/sheet3.xml.rels');
  // Retirer l'Override de sheet3.xml dans Content_Types
  ctXml = ctXml.replace(/<Override[^>]*PartName="\/xl\/worksheets\/sheet3\.xml"[^>]*\/>/g, '');

  // Mettre à jour workbook.xml avec toutes les feuilles
  wbXml = wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/,
    `<sheets>${sheetEntries.join('')}</sheets>`
  );

  // Mettre à jour les relations du workbook
  // Supprimer les anciennes relations de feuilles
  wbRels = wbRels
    .replace(/<Relationship[^>]*rId1[^>]*\/>/g, '')
    .replace(/<Relationship[^>]*rId2[^>]*\/>/g, '')
    .replace(/<Relationship[^>]*rId3[^>]*\/>/g, '')
    .replace(/<Relationship[^>]*rId4[^>]*\/>/g, '')
    .replace(/<Relationship[^>]*rId8[^>]*\/>/g, '');
  // Ajouter les nouvelles
  wbRels = wbRels.replace('</Relationships>', relsEntries.join('') + '</Relationships>');

  // S'assurer que les types PNG/VML sont présents
  if (!ctXml.includes('image/png')) {
    ctXml = ctXml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  if (!ctXml.includes('vmlDrawing')) {
    ctXml = ctXml.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>');
  }

  // Écrire tous les fichiers modifiés
  zip.file('xl/sharedStrings.xml', ssXml);
  zip.file('xl/workbook.xml', wbXml);
  zip.file('xl/_rels/workbook.xml.rels', wbRels);
  zip.file('[Content_Types].xml', ctXml);

  console.log(`✅ Fichier multi-feuilles généré avec ${siteNames.length} feuille(s)`);
  return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
