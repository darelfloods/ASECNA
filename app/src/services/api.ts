const API_BASE_URL = 'http://localhost:3002/api';

export interface HistoryEntry {
  id?: number;
  date: string;
  type: 'facture';
  fileName: string;
  nbConventions: number;
  status: 'success' | 'error';
  details?: string;
  createdAt?: number;
}

// Récupérer l'historique
export async function getHistory(type?: string): Promise<HistoryEntry[]> {
  try {
    const params = new URLSearchParams();
    if (type && type !== 'all') {
      params.append('type', type);
    }
    
    const response = await fetch(`${API_BASE_URL}/history?${params.toString()}`);
    const data = await response.json();
    
    if (!data.success) {
      throw new Error(data.error || 'Erreur lors de la récupération de l\'historique');
    }
    
    return data.data;
  } catch (error) {
    console.error('Erreur API getHistory:', error);
    // Retourner un tableau vide si l'API n'est pas disponible
    return [];
  }
}

// Ajouter une entrée à l'historique
export async function addHistoryEntry(entry: Omit<HistoryEntry, 'id' | 'createdAt'>): Promise<number | null> {
  try {
    const response = await fetch(`${API_BASE_URL}/history`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(entry),
    });
    
    const data = await response.json();
    
    if (!data.success) {
      throw new Error(data.error || 'Erreur lors de l\'ajout à l\'historique');
    }
    
    return data.id;
  } catch (error) {
    console.error('Erreur API addHistoryEntry:', error);
    return null;
  }
}

// Supprimer une entrée de l'historique
export async function deleteHistoryEntry(id: number): Promise<boolean> {
  try {
    const response = await fetch(`${API_BASE_URL}/history/${id}`, {
      method: 'DELETE',
    });
    
    const data = await response.json();
    return data.success;
  } catch (error) {
    console.error('Erreur API deleteHistoryEntry:', error);
    return false;
  }
}

// Vérifier si l'API est disponible
export async function checkAPIHealth(): Promise<boolean> {
  try {
    const response = await fetch(`${API_BASE_URL}/health`);
    const data = await response.json();
    return data.status === 'ok';
  } catch (error) {
    return false;
  }
}
