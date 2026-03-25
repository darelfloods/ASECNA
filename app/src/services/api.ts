const API_BASE_URL = 'http://localhost:3002/api';

export interface HistoryEntry {
  id?: number;
  date: string;
  type: 'facture' | 'fiche-mission' | 'ordre-mission' | 'bon-commande';
  fileName: string;
  nbConventions: number;
  status: 'success' | 'error';
  details?: string;
  createdAt?: number;
  actorEmail?: string;
  actorName?: string;
  actorRole?: 'admin' | 'user' | 'viewer';
  action?: string;
  hasDocument?: boolean; // true si le fichier est stocké en base SQLite
}

export function getFilePreviewUrl(fileName: string): string {
  return `${API_BASE_URL}/files/${encodeURIComponent(fileName)}`;
}

// ── Historique ────────────────────────────────────────────────────────────────
export async function getHistory(type?: string, _userEmail?: string, _userRole?: string): Promise<HistoryEntry[]> {
  try {
    const params = new URLSearchParams();
    if (type && type !== 'all') params.append('type', type);
    const response = await fetch(`${API_BASE_URL}/history?${params.toString()}`);
    const data = await response.json();
    if (!data.success) throw new Error(data.error || 'Erreur lors de la récupération de l\'historique');
    return data.data;
  } catch (error) {
    console.error('Erreur API getHistory:', error);
    return [];
  }
}

export async function addHistoryEntry(entry: Omit<HistoryEntry, 'id' | 'createdAt'>): Promise<number | null> {
  try {
    const response = await fetch(`${API_BASE_URL}/history`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(entry),
    });
    const data = await response.json();
    if (!data.success) throw new Error(data.error || 'Erreur lors de l\'ajout à l\'historique');
    return data.id;
  } catch (error) {
    console.error('Erreur API addHistoryEntry:', error);
    return null;
  }
}

export async function deleteHistoryEntry(id: number): Promise<boolean> {
  try {
    const response = await fetch(`${API_BASE_URL}/history/${id}`, { method: 'DELETE' });
    const data = await response.json();
    return data.success;
  } catch (error) {
    console.error('Erreur API deleteHistoryEntry:', error);
    return false;
  }
}

// ── Stockage permanent des documents générés ─────────────────────────────────
/**
 * Envoie le blob d'un document généré au serveur pour stockage permanent en SQLite.
 * À appeler juste après addHistoryEntry() pour les entrées status='success'.
 */
export async function storeDocument(historyId: number, blob: Blob, fileName: string): Promise<boolean> {
  try {
    const buffer = await blob.arrayBuffer();
    const response = await fetch(`${API_BASE_URL}/documents/${historyId}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/octet-stream',
        'X-File-Name': encodeURIComponent(fileName),
      },
      body: buffer,
    });
    const data = await response.json();
    return data.success === true;
  } catch (error) {
    console.error('Erreur API storeDocument:', error);
    return false;
  }
}

/** URL de téléchargement direct depuis SQLite (par historyId) */
export function getDocumentDownloadUrl(historyId: number): string {
  return `${API_BASE_URL}/documents/${historyId}/download`;
}

/** URL de prévisualisation HTML depuis SQLite (par historyId) */
export function getDocumentPreviewUrl(historyId: number): string {
  return `${API_BASE_URL}/documents/${historyId}/preview-html`;
}

// ── Santé API ─────────────────────────────────────────────────────────────────
export async function checkAPIHealth(): Promise<boolean> {
  try {
    const response = await fetch(`${API_BASE_URL}/health`);
    const data = await response.json();
    return data.status === 'ok';
  } catch { return false; }
}
