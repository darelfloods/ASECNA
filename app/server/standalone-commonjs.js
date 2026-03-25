const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const os = require('os');
const nodemailer = require('nodemailer');
const Database = require('better-sqlite3');

// ── Base de données SQLite ────────────────────────────────────────────────────
const DB_PATH = path.join(__dirname, 'asecna.db');
const db = new Database(DB_PATH);

db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    date        TEXT    NOT NULL,
    type        TEXT    NOT NULL,
    fileName    TEXT    NOT NULL,
    nbConventions INTEGER NOT NULL DEFAULT 0,
    status      TEXT    NOT NULL DEFAULT 'success',
    details     TEXT,
    actorEmail  TEXT,
    actorName   TEXT,
    actorRole   TEXT,
    action      TEXT,
    createdAt   INTEGER NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS documents (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    historyId   INTEGER NOT NULL,
    fileName    TEXT    NOT NULL,
    fileData    BLOB    NOT NULL,
    fileSize    INTEGER NOT NULL,
    createdAt   INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (historyId) REFERENCES history(id) ON DELETE CASCADE
  );
`);

// ── Tables Bandes d'enregistrement ───────────────────────────────────────────
db.exec(`
  CREATE TABLE IF NOT EXISTS fiches_bandes (
    id                      TEXT PRIMARY KEY,
    numero_fiche            TEXT NOT NULL,
    assistant               TEXT,
    numero_vol              TEXT NOT NULL,
    type_vol                TEXT NOT NULL DEFAULT 'irregulier',
    date_saisie             TEXT NOT NULL,
    site                    TEXT NOT NULL,
    compagnie_assistee      TEXT NOT NULL,
    immatricule_aeronef     TEXT,
    vol_national            INTEGER DEFAULT 0,
    vol_regional            INTEGER DEFAULT 0,
    vol_international       INTEGER DEFAULT 0,
    banques_depart          TEXT DEFAULT '[]',
    nombre_banques_depart   INTEGER DEFAULT 0,
    heure_ouverture_depart  TEXT,
    date_ouverture_depart   TEXT,
    heure_cloture_depart    TEXT,
    date_cloture_depart     TEXT,
    pax_prevu_depart        INTEGER DEFAULT 0,
    banques_arrivee         TEXT DEFAULT '[]',
    nombre_banques_arrivee  INTEGER DEFAULT 0,
    heure_ouverture_arrivee TEXT,
    date_ouverture_arrivee  TEXT,
    heure_cloture_arrivee   TEXT,
    date_cloture_arrivee    TEXT,
    duree_comptoirs_minutes INTEGER DEFAULT 0,
    pax_arrives             INTEGER DEFAULT 0,
    pax_departs             INTEGER DEFAULT 0,
    pax_transit             INTEGER DEFAULT 0,
    duree_heures_decimal    REAL DEFAULT 0,
    statut                  TEXT DEFAULT 'saisie',
    createdAt               INTEGER NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS factures_bandes (
    id                  TEXT PRIMARY KEY,
    numero_facture      TEXT NOT NULL,
    date_facture        TEXT NOT NULL,
    compagnie           TEXT NOT NULL,
    adresse_compagnie   TEXT DEFAULT '',
    ville_compagnie     TEXT DEFAULT '',
    site                TEXT NOT NULL,
    serie_bandes        TEXT,
    periode_debut       TEXT,
    periode_fin         TEXT,
    fiches_ids          TEXT DEFAULT '[]',
    nombre_heures       REAL DEFAULT 0,
    tarif_horaire       INTEGER DEFAULT 10000,
    total_heures        INTEGER DEFAULT 0,
    nombre_annonces     INTEGER DEFAULT 0,
    tarif_annonce       INTEGER DEFAULT 3500,
    total_annonces      INTEGER DEFAULT 0,
    montant_ht          INTEGER DEFAULT 0,
    taxes               INTEGER DEFAULT 0,
    acompte             INTEGER DEFAULT 0,
    solde               INTEGER DEFAULT 0,
    total_pax           INTEGER DEFAULT 0,
    montant_en_lettres  TEXT,
    statut              TEXT DEFAULT 'brouillon',
    createdAt           INTEGER NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS config_bandes (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
  );
`);

// Migration : ajout des colonnes adresse_compagnie / ville_compagnie si absentes
try { db.prepare("ALTER TABLE factures_bandes ADD COLUMN adresse_compagnie TEXT DEFAULT ''").run(); } catch {}
try { db.prepare("ALTER TABLE factures_bandes ADD COLUMN ville_compagnie TEXT DEFAULT ''").run(); } catch {}

// Valeurs par défaut config
const configDefaults = [
  ['tarif_horaire', '10000'],
  ['tarif_annonce', '3500'],
  ['next_fiche_numero', '3001'],
  ['next_facture_numero', '1'],
  ['next_bordereau_numero', '1'],
];
const insertConfig = db.prepare('INSERT OR IGNORE INTO config_bandes (key, value) VALUES (?, ?)');
for (const [k, v] of configDefaults) insertConfig.run(k, v);

// ── Migration depuis history.json (si existant) ───────────────────────────────
const LEGACY_DB = path.join(__dirname, 'history.json');
if (fs.existsSync(LEGACY_DB)) {
  try {
    const legacy = JSON.parse(fs.readFileSync(LEGACY_DB, 'utf8'));
    if (legacy.history && legacy.history.length > 0) {
      const existing = db.prepare('SELECT COUNT(*) as cnt FROM history').get();
      if (existing.cnt === 0) {
        const ins = db.prepare(`
          INSERT INTO history (id, date, type, fileName, nbConventions, status, details,
            actorEmail, actorName, actorRole, action, createdAt)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);
        const migrate = db.transaction((entries) => {
          for (const e of entries) {
            ins.run(
              e.id, e.date, e.type, e.fileName, e.nbConventions,
              e.status, e.details || null, e.actorEmail || null,
              e.actorName || null, e.actorRole || null,
              e.action || null, e.createdAt || Date.now()
            );
          }
        });
        migrate(legacy.history);
        console.log(`✅ Migration history.json → SQLite : ${legacy.history.length} entrée(s)`);
      }
    }
    fs.renameSync(LEGACY_DB, LEGACY_DB + '.bak');
  } catch (e) {
    console.error('⚠️  Migration history.json échouée :', e.message);
  }
}

// ── Chargement de la configuration email ─────────────────────────────────────
const EMAIL_CONFIG_FILE = path.join(__dirname, 'email-config.json');

function getEmailConfig() {
  try {
    return JSON.parse(fs.readFileSync(EMAIL_CONFIG_FILE, 'utf8'));
  } catch { return null; }
}

function isEmailConfigured(config) {
  return config &&
    config.auth &&
    config.auth.user &&
    config.auth.user !== 'votre-email@gmail.com' &&
    config.auth.pass &&
    config.auth.pass !== 'votre-mot-de-passe-application';
}

const app = express();
const PORT = process.env.PORT || 3002;

const DOWNLOADS_DIR = path.join(os.homedir(), 'Downloads');

// ── Middlewares ───────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());

// ── Helpers SQLite ────────────────────────────────────────────────────────────
const stmtInsertHistory = db.prepare(`
  INSERT INTO history (date, type, fileName, nbConventions, status, details,
    actorEmail, actorName, actorRole, action, createdAt)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);

const stmtSelectHistory = db.prepare(`
  SELECT h.*,
    CASE WHEN d.id IS NOT NULL THEN 1 ELSE 0 END AS hasDocument
  FROM history h
  LEFT JOIN (
    SELECT MIN(id) AS id, historyId FROM documents GROUP BY historyId
  ) d ON d.historyId = h.id
  ORDER BY h.createdAt DESC
`);

const stmtSelectHistoryByType = db.prepare(`
  SELECT h.*,
    CASE WHEN d.id IS NOT NULL THEN 1 ELSE 0 END AS hasDocument
  FROM history h
  LEFT JOIN (
    SELECT MIN(id) AS id, historyId FROM documents GROUP BY historyId
  ) d ON d.historyId = h.id
  WHERE h.type = ?
  ORDER BY h.createdAt DESC
`);

// ── Routes API – Historique ───────────────────────────────────────────────────
app.get('/api/history', (req, res) => {
  try {
    const { type } = req.query;
    const rows = (type && type !== 'all')
      ? stmtSelectHistoryByType.all(type)
      : stmtSelectHistory.all();
    res.json({ success: true, data: rows });
  } catch (err) {
    console.error('Erreur GET /api/history :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/history', (req, res) => {
  try {
    const { date, type, fileName, nbConventions, status, details,
            actorEmail, actorName, actorRole, action } = req.body;

    if (!date || !type || !fileName || nbConventions === undefined || !status) {
      return res.status(400).json({ success: false, error: 'Champs manquants' });
    }

    const result = stmtInsertHistory.run(
      date, type, fileName, nbConventions, status, details || null,
      actorEmail || null, actorName || null, actorRole || null,
      action || null, Date.now()
    );

    res.json({ success: true, id: result.lastInsertRowid });
  } catch (err) {
    console.error('Erreur POST /api/history :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/history/:id', (req, res) => {
  try {
    db.prepare('DELETE FROM history WHERE id = ?').run(parseInt(req.params.id));
    res.json({ success: true });
  } catch (err) {
    console.error('Erreur DELETE /api/history/:id :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Routes API – Documents (stockage permanent des fichiers générés) ───────────
// Dépôt d'un document (corps binaire brut)
app.post('/api/documents/:historyId',
  express.raw({ type: '*/*', limit: '100mb' }),
  (req, res) => {
    try {
      const historyId = parseInt(req.params.historyId);
      const fileName  = decodeURIComponent(req.headers['x-file-name'] || 'document');
      const fileData  = req.body; // Buffer

      if (!fileData || !fileData.length) {
        return res.status(400).json({ success: false, error: 'Aucun fichier reçu' });
      }

      const hist = db.prepare('SELECT id FROM history WHERE id = ?').get(historyId);
      if (!hist) {
        return res.status(404).json({ success: false, error: 'Entrée historique introuvable' });
      }

      // Remplacer l'éventuel document déjà stocké pour cette entrée
      db.prepare('DELETE FROM documents WHERE historyId = ?').run(historyId);
      const result = db.prepare(`
        INSERT INTO documents (historyId, fileName, fileData, fileSize, createdAt)
        VALUES (?, ?, ?, ?, ?)
      `).run(historyId, fileName, fileData, fileData.length, Date.now());

      res.json({ success: true, id: result.lastInsertRowid });
    } catch (err) {
      console.error('Erreur POST /api/documents/:historyId :', err);
      res.status(500).json({ success: false, error: err.message });
    }
  }
);

// Téléchargement d'un document stocké
app.get('/api/documents/:historyId/download', (req, res) => {
  try {
    const historyId = parseInt(req.params.historyId);
    const doc = db.prepare('SELECT * FROM documents WHERE historyId = ?').get(historyId);

    if (!doc) {
      return res.status(404).json({ success: false, error: 'Document non trouvé en base' });
    }

    const ext = path.extname(doc.fileName).toLowerCase();
    const mimes = {
      '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.zip':  'application/zip',
    };
    const contentType = mimes[ext] || 'application/octet-stream';

    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', `attachment; filename="${doc.fileName}"`);
    res.send(doc.fileData);
  } catch (err) {
    console.error('Erreur GET /api/documents/:historyId/download :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Prévisualisation HTML d'un document stocké
app.get('/api/documents/:historyId/preview-html', async (req, res) => {
  try {
    const historyId = parseInt(req.params.historyId);
    const doc = db.prepare('SELECT * FROM documents WHERE historyId = ?').get(historyId);

    if (!doc) {
      return res.status(404).json({ success: false, error: 'Document non trouvé en base' });
    }

    const ext = path.extname(doc.fileName).toLowerCase();

    if (ext === '.docx') {
      const mammoth = require('mammoth');
      const result  = await mammoth.convertToHtml({ buffer: Buffer.from(doc.fileData) });
      return res.json({ success: true, html: result.value, type: 'docx', fileName: doc.fileName });
    }

    if (ext === '.xlsx' || ext === '.xls') {
      const XLSX     = require('xlsx');
      const workbook = XLSX.read(doc.fileData, { type: 'buffer' });
      const sheet    = workbook.SheetNames[0];
      const html     = XLSX.utils.sheet_to_html(workbook.Sheets[sheet], { id: 'preview-table' });
      return res.json({ success: true, html, type: 'xlsx', fileName: doc.fileName });
    }

    if (ext === '.zip') {
      return res.json({ success: true, html: null, type: 'zip', fileName: doc.fileName });
    }

    return res.status(400).json({ success: false, error: 'Format non supporté pour la prévisualisation' });
  } catch (err) {
    console.error('Erreur GET /api/documents/:historyId/preview-html :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Routes API – Fichiers depuis le dossier Téléchargements (fallback) ────────
app.get('/api/files/:fileName/preview-html', async (req, res) => {
  try {
    const safeName = path.basename(req.params.fileName);
    const filePath = path.join(DOWNLOADS_DIR, safeName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'Fichier introuvable dans le dossier Téléchargements' });
    }

    const ext = path.extname(safeName).toLowerCase();

    if (ext === '.docx') {
      const mammoth = require('mammoth');
      const result  = await mammoth.convertToHtml({ path: filePath });
      return res.json({ success: true, html: result.value, type: 'docx', fileName: safeName });
    }

    if (ext === '.xlsx' || ext === '.xls') {
      const XLSX     = require('xlsx');
      const workbook = XLSX.readFile(filePath);
      const sheet    = workbook.SheetNames[0];
      const html     = XLSX.utils.sheet_to_html(workbook.Sheets[sheet], { id: 'preview-table' });
      return res.json({ success: true, html, type: 'xlsx', fileName: safeName });
    }

    if (ext === '.zip') {
      return res.json({ success: true, html: null, type: 'zip', fileName: safeName });
    }

    return res.status(400).json({ success: false, error: 'Format non supporté pour la prévisualisation' });
  } catch (err) {
    console.error('Erreur prévisualisation fichier :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/files/:fileName/download', (req, res) => {
  try {
    const safeName = path.basename(req.params.fileName);
    const filePath = path.join(DOWNLOADS_DIR, safeName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'Fichier introuvable dans le dossier Téléchargements' });
    }

    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
    res.sendFile(filePath);
  } catch (err) {
    console.error('Erreur téléchargement fichier :', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/files/:fileName', (req, res) => {
  try {
    const safeName = path.basename(req.params.fileName);
    const filePath = path.join(DOWNLOADS_DIR, safeName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'Fichier introuvable dans le dossier Téléchargements' });
    }

    const ext = path.extname(safeName).toLowerCase();
    const mimes = {
      '.pdf':  'application/pdf',
      '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    };
    res.setHeader('Content-Type', mimes[ext] || 'application/octet-stream');
    res.setHeader('Content-Disposition', `inline; filename="${safeName}"`);
    res.sendFile(filePath);
  } catch (err) {
    console.error('Erreur prévisualisation fichier :', err);
    res.status(500).json({ success: false, error: 'Erreur lors de la prévisualisation du fichier' });
  }
});

// ── Routes Bandes d'enregistrement — Fiches ──────────────────────────────────
function getConfigVal(key) {
  const row = db.prepare('SELECT value FROM config_bandes WHERE key = ?').get(key);
  return row ? row.value : null;
}
function setConfigVal(key, value) {
  db.prepare('INSERT OR REPLACE INTO config_bandes (key, value) VALUES (?, ?)').run(key, String(value));
}

app.get('/api/fiches-bandes', (req, res) => {
  try {
    const { compagnie, site, statut } = req.query;
    let sql = 'SELECT * FROM fiches_bandes WHERE 1=1';
    const params = [];
    if (compagnie) { sql += ' AND compagnie_assistee = ?'; params.push(compagnie); }
    if (site) { sql += ' AND site = ?'; params.push(site); }
    if (statut) { sql += ' AND statut = ?'; params.push(statut); }
    sql += ' ORDER BY createdAt DESC';
    const rows = db.prepare(sql).all(...params);
    const parsed = rows.map(r => ({
      ...r,
      banques_depart: JSON.parse(r.banques_depart || '[]'),
      banques_arrivee: JSON.parse(r.banques_arrivee || '[]'),
      vol_national: !!r.vol_national,
      vol_regional: !!r.vol_regional,
      vol_international: !!r.vol_international,
    }));
    res.json({ success: true, data: parsed });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/fiches-bandes/:id', (req, res) => {
  try {
    const row = db.prepare('SELECT * FROM fiches_bandes WHERE id = ?').get(req.params.id);
    if (!row) return res.status(404).json({ success: false, error: 'Fiche introuvable' });
    res.json({ success: true, data: { ...row,
      banques_depart: JSON.parse(row.banques_depart || '[]'),
      banques_arrivee: JSON.parse(row.banques_arrivee || '[]'),
      vol_national: !!row.vol_national, vol_regional: !!row.vol_regional, vol_international: !!row.vol_international,
    }});
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/fiches-bandes', (req, res) => {
  try {
    const d = req.body;
    const nextNum = parseInt(getConfigVal('next_fiche_numero') || '3001');
    const numero_fiche = String(nextNum).padStart(7, '0');
    const id = `fiche-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    db.prepare(`INSERT INTO fiches_bandes
      (id, numero_fiche, assistant, numero_vol, type_vol, date_saisie, site, compagnie_assistee,
       immatricule_aeronef, vol_national, vol_regional, vol_international,
       banques_depart, nombre_banques_depart, heure_ouverture_depart, date_ouverture_depart,
       heure_cloture_depart, date_cloture_depart, pax_prevu_depart,
       banques_arrivee, nombre_banques_arrivee, heure_ouverture_arrivee, date_ouverture_arrivee,
       heure_cloture_arrivee, date_cloture_arrivee, duree_comptoirs_minutes,
       pax_arrives, pax_departs, pax_transit, duree_heures_decimal, statut, createdAt)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(
      id, numero_fiche, d.assistant || '', d.numero_vol, d.type_vol || 'irregulier',
      d.date_saisie, d.site, d.compagnie_assistee, d.immatricule_aeronef || '',
      d.vol_national ? 1 : 0, d.vol_regional ? 1 : 0, d.vol_international ? 1 : 0,
      JSON.stringify(d.banques_depart || []), d.nombre_banques_depart || 0,
      d.heure_ouverture_depart || '', d.date_ouverture_depart || '',
      d.heure_cloture_depart || '', d.date_cloture_depart || '', d.pax_prevu_depart || 0,
      JSON.stringify(d.banques_arrivee || []), d.nombre_banques_arrivee || 0,
      d.heure_ouverture_arrivee || '', d.date_ouverture_arrivee || '',
      d.heure_cloture_arrivee || '', d.date_cloture_arrivee || '', d.duree_comptoirs_minutes || 0,
      d.pax_arrives || 0, d.pax_departs || 0, d.pax_transit || 0,
      d.duree_heures_decimal || 0, 'saisie', Date.now()
    );
    setConfigVal('next_fiche_numero', nextNum + 1);
    res.json({ success: true, id, numero_fiche });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.put('/api/fiches-bandes/:id', (req, res) => {
  try {
    const d = req.body;
    db.prepare(`UPDATE fiches_bandes SET
      assistant=?, numero_vol=?, type_vol=?, date_saisie=?, site=?, compagnie_assistee=?,
      immatricule_aeronef=?, vol_national=?, vol_regional=?, vol_international=?,
      banques_depart=?, nombre_banques_depart=?, heure_ouverture_depart=?, date_ouverture_depart=?,
      heure_cloture_depart=?, date_cloture_depart=?, pax_prevu_depart=?,
      banques_arrivee=?, nombre_banques_arrivee=?, heure_ouverture_arrivee=?, date_ouverture_arrivee=?,
      heure_cloture_arrivee=?, date_cloture_arrivee=?, duree_comptoirs_minutes=?,
      pax_arrives=?, pax_departs=?, pax_transit=?, duree_heures_decimal=?, statut=?
      WHERE id=?`).run(
      d.assistant || '', d.numero_vol, d.type_vol || 'irregulier', d.date_saisie, d.site,
      d.compagnie_assistee, d.immatricule_aeronef || '',
      d.vol_national ? 1 : 0, d.vol_regional ? 1 : 0, d.vol_international ? 1 : 0,
      JSON.stringify(d.banques_depart || []), d.nombre_banques_depart || 0,
      d.heure_ouverture_depart || '', d.date_ouverture_depart || '',
      d.heure_cloture_depart || '', d.date_cloture_depart || '', d.pax_prevu_depart || 0,
      JSON.stringify(d.banques_arrivee || []), d.nombre_banques_arrivee || 0,
      d.heure_ouverture_arrivee || '', d.date_ouverture_arrivee || '',
      d.heure_cloture_arrivee || '', d.date_cloture_arrivee || '', d.duree_comptoirs_minutes || 0,
      d.pax_arrives || 0, d.pax_departs || 0, d.pax_transit || 0,
      d.duree_heures_decimal || 0, d.statut || 'saisie', req.params.id
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/fiches-bandes/:id', (req, res) => {
  try {
    db.prepare('DELETE FROM fiches_bandes WHERE id = ?').run(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Routes Bandes d'enregistrement — Factures ─────────────────────────────────
app.get('/api/factures-bandes', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM factures_bandes ORDER BY createdAt DESC').all();
    const parsed = rows.map(r => ({ ...r, fiches_ids: JSON.parse(r.fiches_ids || '[]') }));
    res.json({ success: true, data: parsed });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/factures-bandes/:id', (req, res) => {
  try {
    const row = db.prepare('SELECT * FROM factures_bandes WHERE id = ?').get(req.params.id);
    if (!row) return res.status(404).json({ success: false, error: 'Facture introuvable' });
    res.json({ success: true, data: { ...row, fiches_ids: JSON.parse(row.fiches_ids || '[]') } });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/factures-bandes', (req, res) => {
  try {
    const d = req.body;
    const year = new Date().getFullYear();
    const nextNum = parseInt(getConfigVal('next_facture_numero') || '1');
    const numero_facture = String(nextNum).padStart(4, '0');
    const id = `fact-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    db.prepare(`INSERT INTO factures_bandes
      (id, numero_facture, date_facture, compagnie, adresse_compagnie, ville_compagnie, site, serie_bandes, periode_debut, periode_fin,
       fiches_ids, nombre_heures, tarif_horaire, total_heures, nombre_annonces, tarif_annonce,
       total_annonces, montant_ht, taxes, acompte, solde, total_pax, montant_en_lettres, statut, createdAt)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(
      id, numero_facture, d.date_facture, d.compagnie, d.adresse_compagnie || '', d.ville_compagnie || '',
      d.site, d.serie_bandes || '', d.periode_debut || '', d.periode_fin || '', JSON.stringify(d.fiches_ids || []),
      d.nombre_heures || 0, d.tarif_horaire || 10000, d.total_heures || 0,
      d.nombre_annonces || 0, d.tarif_annonce || 3500, d.total_annonces || 0,
      d.montant_ht || 0, d.taxes || 0, d.acompte || 0, d.solde || 0,
      d.total_pax || 0, d.montant_en_lettres || '', d.statut || 'brouillon', Date.now()
    );
    // Marquer les fiches comme facturées
    if (d.fiches_ids && d.fiches_ids.length > 0) {
      const upd = db.prepare("UPDATE fiches_bandes SET statut='facturee' WHERE id=?");
      for (const fid of d.fiches_ids) upd.run(fid);
    }
    setConfigVal('next_facture_numero', nextNum + 1);
    res.json({ success: true, id, numero_facture });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.put('/api/factures-bandes/:id', (req, res) => {
  try {
    const d = req.body;
    db.prepare(`UPDATE factures_bandes SET
      date_facture=?, compagnie=?, adresse_compagnie=?, ville_compagnie=?, site=?, serie_bandes=?, periode_debut=?, periode_fin=?,
      fiches_ids=?, nombre_heures=?, tarif_horaire=?, total_heures=?, nombre_annonces=?,
      tarif_annonce=?, total_annonces=?, montant_ht=?, taxes=?, acompte=?, solde=?,
      total_pax=?, montant_en_lettres=?, statut=?
      WHERE id=?`).run(
      d.date_facture, d.compagnie, d.adresse_compagnie || '', d.ville_compagnie || '',
      d.site, d.serie_bandes || '', d.periode_debut || '', d.periode_fin || '',
      JSON.stringify(d.fiches_ids || []), d.nombre_heures || 0, d.tarif_horaire || 10000,
      d.total_heures || 0, d.nombre_annonces || 0, d.tarif_annonce || 3500, d.total_annonces || 0,
      d.montant_ht || 0, d.taxes || 0, d.acompte || 0, d.solde || 0,
      d.total_pax || 0, d.montant_en_lettres || '', d.statut || 'brouillon', req.params.id
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/factures-bandes/:id', (req, res) => {
  try {
    const fact = db.prepare('SELECT fiches_ids FROM factures_bandes WHERE id = ?').get(req.params.id);
    if (fact) {
      const ids = JSON.parse(fact.fiches_ids || '[]');
      const upd = db.prepare("UPDATE fiches_bandes SET statut='saisie' WHERE id=?");
      for (const fid of ids) upd.run(fid);
    }
    db.prepare('DELETE FROM factures_bandes WHERE id = ?').run(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Config Bandes ─────────────────────────────────────────────────────────────
app.get('/api/config-bandes', (req, res) => {
  try {
    const rows = db.prepare('SELECT key, value FROM config_bandes').all();
    const config = Object.fromEntries(rows.map(r => [r.key, r.value]));
    res.json({ success: true, data: config });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Stats Bandes ──────────────────────────────────────────────────────────────
app.get('/api/stats-bandes', (req, res) => {
  try {
    const now = new Date();
    const monthStart = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-01`;
    const fichesTotal = db.prepare("SELECT COUNT(*) as c FROM fiches_bandes WHERE date_saisie >= ?").get(monthStart)?.c || 0;
    const fichesSaisie = db.prepare("SELECT COUNT(*) as c FROM fiches_bandes WHERE statut='saisie'").get()?.c || 0;
    const facturesMois = db.prepare("SELECT COALESCE(SUM(montant_ht),0) as total FROM factures_bandes WHERE date_facture >= ?").get(monthStart)?.total || 0;
    const facturesAttente = db.prepare("SELECT COUNT(*) as c FROM factures_bandes WHERE statut='emise'").get()?.c || 0;
    res.json({ success: true, data: { fichesTotal, fichesSaisie, facturesMois, facturesAttente } });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Route santé ───────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'API ASECNA opérationnelle' });
});

// ── Statut et envoi d'email ───────────────────────────────────────────────────
app.get('/api/email-status', (req, res) => {
  const config = getEmailConfig();
  res.json({ configured: isEmailConfigured(config) });
});

app.post('/api/send-approval-email', async (req, res) => {
  const { email, prenom, nom } = req.body;

  if (!email || !prenom || !nom) {
    return res.status(400).json({ success: false, error: 'Paramètres manquants (email, prenom, nom)' });
  }

  const config = getEmailConfig();

  if (!isEmailConfigured(config)) {
    console.warn('⚠️  Email non configuré — aucun email envoyé à', email);
    return res.json({
      success: false,
      warning: 'Email non configuré. Éditez server/email-config.json pour activer les notifications.',
      email
    });
  }

  try {
    const transporter = nodemailer.createTransport({
      host: config.host, port: config.port, secure: config.secure,
      auth: { user: config.auth.user, pass: config.auth.pass }
    });

    const logoPath = path.join(__dirname, '../public/ASECNA_logo.png');
    const hasLogo  = fs.existsSync(logoPath);

    const html = `
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:20px;}
  .container{max-width:600px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);}
  .header{background:#1a5276;padding:28px 30px 20px;text-align:center;}
  .header img{height:72px;width:auto;margin-bottom:12px;display:block;margin-left:auto;margin-right:auto;}
  .header h1{color:#fff;margin:0 0 4px;font-size:20px;letter-spacing:1px;}
  .header p{color:#aed6f1;margin:0;font-size:13px;}
  .body{padding:30px;color:#333;}
  .body h2{color:#1a5276;margin-top:0;}
  .badge{display:inline-block;background:#27ae60;color:#fff;padding:6px 20px;border-radius:20px;font-size:14px;font-weight:bold;margin:10px 0;}
  .info-box{background:#eaf4fb;border-left:4px solid #1a5276;padding:15px 20px;border-radius:4px;margin:20px 0;}
  .footer{background:#f0f0f0;padding:15px 30px;text-align:center;font-size:12px;color:#888;}
</style></head><body>
<div class="container">
  <div class="header">
    ${hasLogo ? '<img src="cid:asecna-logo" alt="ASECNA Logo"/>' : ''}
    <h1>ASECNA</h1>
    <p>Service Budget et Facturation — Délégation du Gabon</p>
  </div>
  <div class="body">
    <h2>Bonjour ${prenom} ${nom},</h2>
    <p>Nous avons le plaisir de vous informer que votre compte a été <strong>validé</strong> par un administrateur.</p>
    <div style="text-align:center;margin:20px 0;"><span class="badge">✓ Compte activé</span></div>
    <div class="info-box">
      <p style="margin:0;">Vous pouvez maintenant vous connecter à l'application <strong>ASECNA Facturation</strong> avec votre adresse email et votre mot de passe.</p>
    </div>
    <p>Si vous rencontrez des difficultés pour vous connecter, veuillez contacter votre administrateur.</p>
    <p>Cordialement,<br><strong>L'équipe ASECNA — Service Budget et Facturation</strong></p>
  </div>
  <div class="footer">© ${new Date().getFullYear()} ASECNA Délégation du Gabon — Usage interne uniquement</div>
</div></body></html>`;

    await transporter.sendMail({
      from: config.from,
      to: email,
      subject: 'Votre compte ASECNA Facturation a été activé',
      html,
      attachments: hasLogo ? [{ filename: 'ASECNA_logo.png', path: logoPath, cid: 'asecna-logo' }] : []
    });

    console.log(`✅ Email d'approbation envoyé à ${email}`);
    res.json({ success: true, message: `Email envoyé à ${email}` });
  } catch (err) {
    console.error('❌ Erreur envoi email :', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Fichiers statiques frontend ───────────────────────────────────────────────
const isDev      = !process.env.NODE_ENV || process.env.NODE_ENV === 'development';
const isElectron = process.versions && process.versions.electron;

let distPath, publicPath;

if (isElectron && !isDev) {
  const resourcesPath = process.resourcesPath;
  distPath   = path.join(resourcesPath, 'app.asar.unpacked', 'dist');
  publicPath = path.join(resourcesPath, 'app.asar.unpacked', 'public');
} else {
  distPath   = path.join(__dirname, '../dist');
  publicPath = path.join(__dirname, '../public');
}

console.log('📁 Chemin dist :', distPath);
console.log('📁 Chemin public :', publicPath);

app.use(express.static(distPath));
app.use(express.static(publicPath));

app.get('*', (req, res) => {
  if (!req.url.startsWith('/api')) {
    const indexPath = path.join(distPath, 'index.html');
    if (fs.existsSync(indexPath)) {
      res.sendFile(indexPath);
    } else {
      res.status(404).send('index.html not found at: ' + indexPath);
    }
  }
});

// ── Démarrage ─────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log('╔════════════════════════════════════════════╗');
  console.log('║   ASECNA - Service Budget et Facturation   ║');
  console.log('╚════════════════════════════════════════════╝');
  console.log('');
  console.log(`✅ Serveur démarré sur http://localhost:${PORT}`);
  console.log(`📦 Base de données : ${DB_PATH}`);
  console.log('');
});

module.exports = app;
