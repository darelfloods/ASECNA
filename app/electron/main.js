const { app, BrowserWindow } = require('electron');
const path = require('path');

let mainWindow;

// Démarrer le serveur Node.js dans le même process qu'Electron
function startServer() {
  const serverPath = path.join(__dirname, '../server/standalone-commonjs.js');

  try {
    // Utilise le runtime Node intégré à Electron (pas besoin de Node installé sur la machine)
    require(serverPath);
    console.log('✅ Serveur ASECNA démarré depuis le process Electron');
  } catch (err) {
    console.error('Erreur serveur:', err);
  }
}

// Créer la fenêtre principale
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1000,
    minHeight: 700,
    title: 'ASECNA - Service Budget et Facturation',
    icon: path.join(__dirname, '../public/ASECNA_logo.png'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
    autoHideMenuBar: true,
    show: false
  });

  // Attendre que le serveur soit prêt avant de charger
  setTimeout(() => {
    mainWindow.loadURL('http://localhost:3001');
    mainWindow.show();
  }, 2000);

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Quand Electron est prêt
app.whenReady().then(() => {
  startServer();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// Quitter proprement
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('will-quit', () => {
});
