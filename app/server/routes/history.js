import express from 'express';
import { addHistoryEntry, getHistory, deleteHistoryEntry, clearAllHistory } from '../database.js';

const router = express.Router();

// GET /api/history - Récupérer l'historique
router.get('/', (req, res) => {
  try {
    const { type, limit, userEmail, userRole } = req.query;
    const filters = {};

    if (type && type !== 'all') {
      filters.type = type;
    }

    if (limit) {
      filters.limit = parseInt(limit);
    }

    let history = getHistory(filters);

    // Masquer les infos des autres utilisateurs pour les non-admins
    if (userRole !== 'admin' && userEmail) {
      history = history.map(entry => {
        if (entry.actorEmail === userEmail) {
          return entry;
        }
        return {
          ...entry,
          actorEmail: null,
          actorName: null,
          actorRole: null
        };
      });
    }

    res.json({ success: true, data: history });
  } catch (error) {
    console.error('Erreur récupération historique:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// POST /api/history - Ajouter une entrée
router.post('/', (req, res) => {
  try {
    const { date, type, fileName, nbConventions, status, details, actorEmail, actorName, actorRole, action } = req.body;

    if (!date || !type || !fileName || nbConventions === undefined || !status) {
      return res.status(400).json({
        success: false,
        error: 'Champs manquants'
      });
    }

    const id = addHistoryEntry({
      date,
      type,
      fileName,
      nbConventions,
      status,
      details,
      actorEmail,
      actorName,
      actorRole,
      action
    });
    
    res.json({ success: true, id });
  } catch (error) {
    console.error('Erreur ajout historique:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// DELETE /api/history/clear - Vider tout l'historique (DOIT être avant /:id)
router.delete('/clear', (req, res) => {
  try {
    const result = clearAllHistory();
    res.json({ success: true, deletedCount: result.changes });
  } catch (error) {
    console.error('Erreur vidage historique:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// DELETE /api/history/:id - Supprimer une entrée
router.delete('/:id', (req, res) => {
  try {
    const { id } = req.params;
    deleteHistoryEntry(parseInt(id));
    res.json({ success: true });
  } catch (error) {
    console.error('Erreur suppression historique:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// DELETE /api/history - Vider tout l'historique
router.delete('/', (req, res) => {
  try {
    const { clearHistory } = require('../database.js');
    clearHistory();
    res.json({ success: true, message: 'Historique vidé' });
  } catch (error) {
    console.error('Erreur vidage historique:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

export default router;
