export type DownloadType = 'docx' | 'xlsx' | 'zip';

export interface PendingDownload {
  blob: Blob;
  filename: string;
  type: DownloadType;
  bonCommande?: { data: any; numero: string };
}

// ── Charge le logo ASECNA en base64 pour l'embarquer dans le HTML ────────────
async function fetchLogoBase64(): Promise<string> {
  try {
    const res = await fetch('/ASECNA_logo.png');
    if (!res.ok) return '';
    const buf = await res.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = '';
    bytes.forEach(b => { binary += String.fromCharCode(b); });
    return 'data:image/png;base64,' + btoa(binary);
  } catch {
    return '';
  }
}

// ── Renderer HTML fidèle pour le BON DE COMMANDE (A4 paysage) ────────────────
function renderBonCommandeHtml(data: any, numero: string, logoSrc: string): string {
  // Boîte individuelle pour CS/CR/CC/Article/Exercice/Code fournisseur
  const box = (c: string) =>
    `<span style="display:inline-block;border:1px solid #000;min-width:20px;height:17px;` +
    `text-align:center;font-weight:bold;color:#1F3864;font-size:8pt;` +
    `line-height:17px;padding:0 2px;">${c || '\u00A0'}</span>`;

  const boxes = (val: string, n: number) =>
    Array.from({ length: n }, (_, i) => box((val || '')[i] || '')).join('');

  // Ligne CS/CR/CC/Article/Exercice : label à gauche, boîtes à droite
  const codeRow = (label: string, val: string, n: number) =>
    `<tr>
      <td style="padding:1px 6px 1px 0;white-space:nowrap;font-size:7.5pt;">${label}</td>
      <td style="padding:1px 0;">${boxes(val, n)}</td>
    </tr>`;

  const lignesHtml = ((data.lignes || []) as any[]).map((l: any) => `
    <tr>
      <td style="border:1px solid #000;padding:2px 5px;color:#1F3864;font-weight:bold;">${l.description || ''}</td>
      <td style="border:1px solid #000;padding:2px 4px;text-align:center;color:#1F3864;font-weight:bold;">${l.quantite || ''}</td>
      <td style="border:1px solid #000;padding:2px 4px;text-align:right;color:#1F3864;font-weight:bold;">${l.prixUnitaire || ''}</td>
      <td style="border:1px solid #000;padding:2px 4px;text-align:right;color:#1F3864;font-weight:bold;">${l.total || ''}</td>
    </tr>`).join('');

  const montantTotal = ((data.lignes || []) as any[])
    .reduce((s: number, l: any) => s + (parseFloat(String(l.total || '').replace(/\s/g, '').replace(',', '.')) || 0), 0)
    .toLocaleString('fr-FR');

  const logoHtml = logoSrc
    ? `<img src="${logoSrc}" style="width:52px;height:52px;object-fit:contain;display:block;" alt="ASECNA"/>`
    : `<div style="width:52px;height:52px;background:#1a4a8a;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:bold;font-size:8pt;flex-shrink:0;">ASECNA</div>`;

  // Tampons numériques (cercle N° série + rectangle APPROUVÉ)
  const stampHtml = data.numeroSerie ? `
    <div style="display:flex;flex-direction:column;align-items:center;gap:5px;margin-top:6px;">
      <div style="border:2px solid #CC0000;border-radius:50%;width:58px;height:58px;
                  display:flex;flex-direction:column;align-items:center;justify-content:center;
                  color:#CC0000;line-height:1.3;text-align:center;flex-shrink:0;">
        <div style="font-size:6pt;font-weight:bold;text-transform:uppercase;">N° Série</div>
        <div style="font-size:8pt;font-weight:bold;">${data.numeroSerie}</div>
      </div>
      <div style="border:2.5px solid #CC0000;padding:2px 8px;color:#CC0000;
                  font-size:15pt;font-weight:900;transform:rotate(-5deg);letter-spacing:1px;
                  white-space:nowrap;">APPROUVÉ</div>
    </div>` : '';

  return `<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>BON DE COMMANDE ${numero}</title>
  <style>
    @page { size: A4 landscape; margin: 1.8cm 2cm; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; font-size: 8pt; color: #000; }
    table { border-collapse: collapse; width: 100%; }
    .val { color: #1F3864; font-weight: bold; }
    .sec { margin-bottom: 4px; }
    .bar { background:#1a4a8a;color:#fff;padding:6px 14px;border-radius:5px;margin-bottom:8px;
           display:flex;align-items:center;justify-content:space-between; }
    @media print { .bar { display:none!important; } }
  </style>
</head>
<body>
  <div class="bar">
    BON DE COMMANDE ${numero} &mdash; <em>Imprimer &rarr; Enregistrer en PDF</em>
    <button onclick="window.print()"
      style="background:#fff;color:#1a4a8a;border:none;padding:4px 14px;
             border-radius:4px;font-weight:700;cursor:pointer;font-size:10pt;">Imprimer</button>
  </div>

  <!-- EN-TÊTE ─────────────────────────────────────────────────────── -->
  <table class="sec" style="border:1px solid #000;">
    <colgroup>
      <col style="width:58%"/>
      <col style="width:42%"/>
    </colgroup>
    <tr>
      <td style="border:1px solid #000;padding:5px 8px;vertical-align:middle;">
        <div style="display:flex;align-items:center;gap:10px;">
          ${logoHtml}
          <div>
            <div style="font-weight:bold;font-size:8pt;">AGENCE POUR LA SECURITE</div>
            <div style="font-weight:bold;font-size:8pt;">DE LA NAVIGATION AERIENNE</div>
            <div style="font-weight:bold;font-size:8pt;">EN AFRIQUE ET A MADAGASCAR</div>
          </div>
        </div>
      </td>
      <td style="border:1px solid #000;padding:5px 8px;text-align:center;vertical-align:middle;font-size:7.5pt;">
        ADRESSE ASECNA
      </td>
    </tr>
    <tr>
      <td colspan="2" style="border:1px solid #000;padding:5px;text-align:center;font-weight:bold;font-size:12pt;">
        BON DE COMMANDE &nbsp;&nbsp; N°&nbsp;${numero}
      </td>
    </tr>
  </table>

  <!-- CODES BUDGÉTAIRES ───────────────────────────────────────────── -->
  <table class="sec" style="width:auto;">
    <tbody>
      ${codeRow('CENTRE DE SYNTHESE (C.S)', data.cs, 3)}
      ${codeRow('CENTRE DE RESPONSABILITE (C.R)', data.cr, 3)}
      ${codeRow('CENTRE DE COUT (C.C)', data.cc, 3)}
      ${codeRow('ARTICLE', data.article, 3)}
      ${codeRow('EXERCICE', data.exercice, 4)}
    </tbody>
  </table>

  <!-- FOURNISSEUR ─────────────────────────────────────────────────── -->
  <table class="sec" style="border:1px solid #000;">
    <colgroup>
      <col style="width:66%"/>
      <col style="width:34%"/>
    </colgroup>
    <tr>
      <td style="border:1px solid #000;padding:4px 7px;vertical-align:top;">
        <div style="font-size:7pt;margin-bottom:3px;">ADRESSE DU FOURNISSEUR :</div>
        <div class="val">${data.fournisseurNom || ''}</div>
        <div class="val">${data.fournisseurAdresse1 || ''}</div>
        <div class="val">${data.fournisseurAdresse2 || ''}</div>
      </td>
      <td style="border:1px solid #000;padding:4px 7px;vertical-align:top;">
        <div style="font-size:7pt;margin-bottom:4px;">CODE FOURNISSEUR</div>
        <div>${boxes(data.codeFournisseur, 4)}</div>
      </td>
    </tr>
  </table>

  <!-- AVIS ────────────────────────────────────────────────────────── -->
  <div class="sec" style="border:1px solid #000;padding:3px 7px;font-size:7pt;">
    <strong>AVIS TRES IMPORTANT</strong> &mdash; LE PRESENT BON N'ENGAGE L'ASECNA QUE S'IL COMPORTE
    LE NUMERO D'ENGAGEMENT, LE VISA ET LE CACHET DU SERVICE DES ENGAGEMENTS DE L'ASECNA
  </div>

  <!-- TABLEAU DE DÉTAIL ───────────────────────────────────────────── -->
  <table class="sec" style="border:1px solid #000;">
    <colgroup>
      <col style="width:46%"/>
      <col style="width:14%"/>
      <col style="width:21%"/>
      <col style="width:19%"/>
    </colgroup>
    <thead>
      <tr style="background:#e8e8e8;">
        <th style="border:1px solid #000;padding:3px 5px;text-align:left;font-size:7.5pt;">DETAIL DE LA COMMANDE</th>
        <th style="border:1px solid #000;padding:3px 5px;text-align:center;font-size:7.5pt;">QUANTITE</th>
        <th style="border:1px solid #000;padding:3px 5px;text-align:center;font-size:7.5pt;">PRIX UNITAIRE</th>
        <th style="border:1px solid #000;padding:3px 5px;text-align:center;font-size:7.5pt;">TOTAL</th>
      </tr>
    </thead>
    <tbody>
      ${lignesHtml}
      <tr>
        <td colspan="3" style="border:1px solid #000;padding:3px 5px;font-weight:bold;font-size:7.5pt;">MONTANT TOTAL EN CHIFFRE</td>
        <td style="border:1px solid #000;padding:3px 5px;text-align:right;" class="val">${montantTotal}</td>
      </tr>
    </tbody>
  </table>

  <!-- MONTANT EN LETTRES / DÉLAI ──────────────────────────────────── -->
  <table class="sec" style="border:1px solid #000;">
    <tr>
      <td style="border:1px solid #000;padding:3px 7px;">
        <span style="font-size:7.5pt;">MONTANT TOTAL EN LETTRE :</span>
        <span class="val"> ${data.montantTotalLettres || ''}</span>
      </td>
    </tr>
    <tr>
      <td style="border:1px solid #000;padding:3px 7px;">
        <span style="font-size:7.5pt;">DELAI DE LIVRAISON :</span>
        <span class="val"> ${data.delaiLivraison || ''}</span>
        <span style="font-size:7pt;color:#333;margin-left:6px;">Passé ce délai, l'ASECNA se réserve le droit de considérer le présent bon comme nul</span>
      </td>
    </tr>
  </table>

  <!-- NOTA ────────────────────────────────────────────────────────── -->
  <div class="sec" style="border:1px solid #000;padding:3px 7px;font-size:7pt;">
    <strong>NOTA</strong> : LES FACTURES, AVEC MENTION DES PRIX UNITAIRES DOIVENT ETRE ADRESSEES EN 4 EXEMPLAIRES
    ACCOMPAGNEES DU BON DE COMMANDE ORIGINAL ET D'UN EXEMPLAIRE DU B.L DUMENT DE CHARGE
  </div>

  <!-- ZONE DE VALIDATION ──────────────────────────────────────────── -->
  <table style="border:1px solid #000;table-layout:fixed;">
    <colgroup>
      <col style="width:73%"/>
      <col style="width:27%"/>
    </colgroup>
    <tr>
      <!-- Compte limitatif + engagement (gauche) -->
      <td style="border:1px solid #000;padding:5px 8px;vertical-align:top;">
        <div style="display:flex;gap:4px;flex-wrap:wrap;align-items:baseline;margin-bottom:3px;">
          <span style="font-size:7.5pt;white-space:nowrap;">COMPTE LIMITATIF :</span>
          <span style="font-size:7.5pt;white-space:nowrap;">A :&nbsp;<span class="val">${data.lieu || ''}</span></span>
          <span style="font-size:7.5pt;white-space:nowrap;">LE :&nbsp;<span class="val">${data.date || ''}</span></span>
        </div>
        <div style="margin-bottom:3px;font-size:7.5pt;">OPERATION : <span class="val">${data.operation || ''}</span></div>
        <div style="margin-bottom:6px;font-size:7.5pt;">N° D'ENGAGEMENT : <span class="val">${data.numeroEngagement || ''}</span></div>
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div>
            <div style="font-size:6.5pt;color:#444;">1er exemplaire original à retourner</div>
            <div style="font-size:6.5pt;color:#444;">2e exemplaire copie à conserver</div>
          </div>
          <div style="text-align:center;">
            <div style="font-size:7.5pt;font-weight:bold;">VISA ET CACHET DU SERVICE</div>
            <div style="font-size:7.5pt;font-weight:bold;">DES ENGAGEMENTS DE L'ASECNA</div>
            <div style="font-size:7pt;margin-top:3px;">
              A :&nbsp;<span class="val">${data.lieu || ''}</span>
              &nbsp;&nbsp;LE :&nbsp;<span class="val">${data.date || ''}</span>
            </div>
            ${stampHtml}
          </div>
        </div>
      </td>
      <!-- Visa ordonnateur (droite) -->
      <td style="border:1px solid #000;padding:5px 8px;vertical-align:top;">
        <div style="font-size:7.5pt;margin-bottom:4px;">
          A :&nbsp;<span class="val">${data.lieu || ''}</span>
          &nbsp;&nbsp;LE :&nbsp;<span class="val">${data.date || ''}</span>
        </div>
        <div style="height:50px;"></div>
        <div style="font-size:7.5pt;font-weight:bold;">VISA ET CACHET DE L'ORDONNATEUR</div>
      </td>
    </tr>
  </table>

  <script>
    window.addEventListener('load', function() { setTimeout(function() { window.print(); }, 700); });
  </script>
</body>
</html>`;
}

// ── Point d'entrée principal ──────────────────────────────────────────────────
export async function downloadAsPdf(pending: PendingDownload): Promise<void> {
  if (!pending.bonCommande) return; // Seul le bon de commande passe par ici

  const logoSrc = await fetchLogoBase64();
  const html = renderBonCommandeHtml(pending.bonCommande.data, pending.bonCommande.numero, logoSrc);

  const pw = window.open('', '_blank', 'width=1100,height=800');
  if (!pw) {
    alert('Pop-ups bloqués. Autorisez les pop-ups pour ce site puis réessayez.');
    return;
  }
  pw.document.open();
  pw.document.write(html);
  pw.document.close();
}
