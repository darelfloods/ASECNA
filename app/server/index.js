import express from 'express';
import cors from 'cors';
import historyRouter from './routes/history.js';

const app = express();
const PORT = 3001;

// Middlewares
app.use(cors());
app.use(express.json());

// Routes
app.use('/api/history', historyRouter);

// Route de test
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'API ASECNA opérationnelle' });
});

// Démarrer le serveur
app.listen(PORT, () => {
  console.log(`✅ Serveur API démarré sur http://localhost:${PORT}`);
  console.log(`📊 Base de données SQLite prête`);
});
