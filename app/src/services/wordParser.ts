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

        // Parser le texte pour extraire les valeurs
        const data: FicheMissionData = {
          nom: extractValue(text, /Nom\s*:\s*([^,]+),/i) || '',
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
        };

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

/**
 * Génère un ORDRE DE MISSION à partir des données et du template
 */
export async function generateOrdreMission(
  data: FicheMissionData,
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
  data: FicheMissionData,
  numeroOrdre: string
): Promise<Blob> {
  const response = await fetch('/ORDRE DE MISSION.docx');
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);

  // Lire le document.xml
  let documentXml = zip.file('word/document.xml')!.asText();

  // Remplacer les valeurs dans le XML (recherche et remplacement brut)
  documentXml = documentXml.replace(/BOUTE BOU KISSEMBA/g, data.nom);
  documentXml = documentXml.replace(/Chérubin/g, data.prenom);
  documentXml = documentXml.replace(/Commandant d'aérodrome PI/g, data.emploi);
  documentXml = documentXml.replace(/Oyem(?!-)/g, data.residence);
  documentXml = documentXml.replace(/Bitam/g, data.destination);
  documentXml = documentXml.replace(/Obsèques de Feu ENGOUROU MOTO Sylvain, ex Observateur météo/g, data.motif);
  documentXml = documentXml.replace(/06\/02\/2025/g, data.dateDepart);
  documentXml = documentXml.replace(/07\/02\/2025/g, data.dateRetour);
  documentXml = documentXml.replace(/02 Jours/g, `${data.duree} Jours`);
  documentXml = documentXml.replace(/voiture/g, data.transport);
  documentXml = documentXml.replace(/N°011\/2025/g, `N°${numeroOrdre}`);

  // Remettre le XML modifié dans le zip
  zip.file('word/document.xml', documentXml);

  // Générer le nouveau fichier Word
  const outputBlob = zip.generate({ type: 'blob' });
  return outputBlob;
}

/**
 * Génère une FICHE DE MISSION modifiée à partir des données
 */
export async function generateFicheMission(
  data: FicheMissionData
): Promise<Blob> {
  const response = await fetch('/FICHE DE MISSION.docx');
  const templateBuffer = await response.arrayBuffer();

  const zip = new PizZip(templateBuffer);

  // Lire le document.xml
  let documentXml = zip.file('word/document.xml')!.asText();

  // Remplacer les valeurs dans le XML basées sur le template FICHE DE MISSION
  // IMPORTANT: On échappe les caractères spéciaux pour le XML
  const escapeXml = (str: string) => {
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');
  };

  // Nom - chercher avec l'apostrophe échappée aussi
  documentXml = documentXml.replace(/ASSEKO EBOZO'O/g, data.nom);
  documentXml = documentXml.replace(/ASSEKO EBOZO&apos;O/g, data.nom);
  
  // Matricule
  documentXml = documentXml.replace(/26072005/g, data.matricule);
  
  // Prénom
  documentXml = documentXml.replace(/Darel/g, data.prenom);
  
  // Emploi - attention aux espaces
  documentXml = documentXml.replace(/Chef Unité Météo\s*/g, data.emploi);
  
  // Résidence
  documentXml = documentXml.replace(/Libreville/g, data.residence);
  
  // Destination
  documentXml = documentXml.replace(/Oyem- Bitam/g, data.destination);
  
  // Motif - peut avoir des espaces en début
  documentXml = documentXml.replace(/\s*Obsèques de Feu ENGOUROU MOTO Sylvain/g, data.motif);
  
  // Date de départ
  documentXml = documentXml.replace(/26\/07\/2005/g, data.dateDepart.replace(/\//g, '/'));
  
  // Date de retour - cherchons le pattern exact
  documentXml = documentXml.replace(/09\/02\/2025/g, data.dateRetour.replace(/\//g, '/'));
  
  // Durée - essayons plusieurs formats
  documentXml = documentXml.replace(/04Jours/g, `${data.duree}Jours`);
  documentXml = documentXml.replace(/04\s*Jours/g, `${data.duree}Jours`);
  
  // Transport
  documentXml = documentXml.replace(/Avion/g, data.transport);
  
  // Debug: Sauvegarder le XML modifié pour vérification
  console.log('=== MODIFICATIONS APPLIQUÉES ===');
  console.log('Nom:', data.nom);
  console.log('Matricule:', data.matricule);
  console.log('Dates:', data.dateDepart, '->', data.dateRetour);
  console.log('Durée:', data.duree);

  // Remettre le XML modifié dans le zip
  zip.file('word/document.xml', documentXml);

  // Générer le nouveau fichier Word
  const outputBlob = zip.generate({ type: 'blob' });
  return outputBlob;
}
