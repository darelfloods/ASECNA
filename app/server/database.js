import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { readFileSync, writeFileSync, existsSync } from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const DB_FILE = join(__dirname, 'history.json');

// Initialiser le fichier JSON s'il n'existe pas
if (!existsSync(DB_FILE)) {
  writeFileSync(DB_FILE, JSON.stringify({ history: [], nextId: 1 }, null, 2));
}

// Lire la base de données
function readDB() {
  try {
    const data = readFileSync(DB_FILE, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    return { history: [], nextId: 1 };
  }
}

// Écrire dans la base de données
function writeDB(data) {
  writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
}

// Fonction pour ajouter une entrée à l'historique
export function addHistoryEntry(entry) {
  const db = readDB();
  const id = db.nextId;
  
  const newEntry = {
    id,
    date: entry.date,
    type: entry.type,
    fileName: entry.fileName,
    nbConventions: entry.nbConventions,
    status: entry.status,
    details: entry.details || null,
    actorEmail: entry.actorEmail || null,
    actorName: entry.actorName || null,
    actorRole: entry.actorRole || null,
    action: entry.action || null,
    createdAt: Date.now()
  };
  
  db.history.unshift(newEntry); // Ajouter au début
  db.nextId += 1;
  
  writeDB(db);
  return id;
}

// Fonction pour récupérer l'historique
export function getHistory(filters = {}) {
  const db = readDB();
  let history = db.history;
  
  // Filtrer par type si spécifié
  if (filters.type) {
    history = history.filter(item => item.type === filters.type);
  }
  
  // Limiter les résultats si spécifié
  if (filters.limit) {
    history = history.slice(0, filters.limit);
  }
  
  return history;
}

// Fonction pour supprimer une entrée
export function deleteHistoryEntry(id) {
  const db = readDB();
  db.history = db.history.filter(item => item.id !== id);
  writeDB(db);
  return { changes: 1 };
}

// Fonction pour nettoyer l'historique ancien (optionnel)
export function cleanOldHistory(daysOld = 90) {
  const db = readDB();
  const cutoffDate = Date.now() - (daysOld * 24 * 60 * 60 * 1000);
  const oldCount = db.history.length;
  
  db.history = db.history.filter(item => item.createdAt >= cutoffDate);
  writeDB(db);
  
  return { changes: oldCount - db.history.length };
}

// Fonction pour vider complètement l'historique
export function clearHistory() {
  const db = readDB();
  const count = db.history.length;
  db.history = [];
  writeDB(db);
  return { changes: count };
}

// Fonction pour vider tout l'historique
export function clearAllHistory() {
  const db = readDB();
  const count = db.history.length;
  db.history = [];
  writeDB(db);
  return { changes: count };
}
