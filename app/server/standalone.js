import express from 'express';
import cors from 'cors';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import historyRouter from './routes/history.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3001;

// Middlewares
app.use(cors());
app.use(express.json());

// Routes API
app.use('/api/history', historyRouter);

// Route de test
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'API ASECNA opérationnelle' });
});

// Servir les fichiers statiques du frontend (après build)
const distPath = join(__dirname, '../dist');
app.use(express.static(distPath));

// Rediriger toutes les autres routes vers index.html (pour le routing React)
app.get('*', (req, res) => {
  if (!req.url.startsWith('/api')) {
    res.sendFile(join(distPath, 'index.html'));
  }
});

// Démarrer le serveur
app.listen(PORT, () => {
  console.log('╔════════════════════════════════════════════╗');
  console.log('║   ASECNA - Service Budget et Facturation   ║');
  console.log('╚════════════════════════════════════════════╝');
  console.log('');
  console.log(`✅ Serveur démarré sur http://localhost:${PORT}`);
  console.log(`📊 Base de données prête`);
  console.log(`📁 Fichiers Excel dans le dossier public/`);
  console.log('');
  console.log(`👉 Ouvrez votre navigateur sur http://localhost:${PORT}`);
  console.log('');
  console.log('Appuyez sur Ctrl+C pour arrêter le serveur');
});
