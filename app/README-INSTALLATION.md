# ASECNA - Application de Génération de Factures

## Installation et Utilisation

### Prérequis
- **Node.js** version 18 ou supérieure
  - Télécharger depuis : https://nodejs.org/
  - Pendant l'installation, cochez "Automatically install the necessary tools"

### Installation (Première utilisation)

1. **Installer Node.js** (si pas déjà installé)
   - Téléchargez Node.js depuis https://nodejs.org/
   - Installez avec les options par défaut

2. **Lancer l'application**
   - Double-cliquez sur `start-windows.bat`
   - L'application va :
     - Installer automatiquement les dépendances (première fois seulement)
     - Construire l'application (première fois seulement)
     - Démarrer le serveur
     - Ouvrir automatiquement votre navigateur

3. **Utiliser l'application**
   - L'application s'ouvre automatiquement dans votre navigateur
   - Adresse : http://localhost:3001
   - Si le navigateur ne s'ouvre pas, ouvrez manuellement cette adresse

### Utilisation quotidienne

- **Démarrer** : Double-cliquez sur `start-windows.bat`
- **Arrêter** : Fermez la fenêtre du terminal ou appuyez sur Ctrl+C

### Structure des fichiers

```
app/
├── start-windows.bat          # Script de démarrage (double-cliquez ici)
├── server/                    # Serveur backend
│   ├── standalone.js          # Serveur principal
│   ├── database.js            # Gestion de l'historique
│   └── history.json           # Base de données de l'historique
├── dist/                      # Application web (générée automatiquement)
└── public/                    # Fichiers Excel (templates de factures)
```

### Fonctionnalités

1. **Génération de factures**
   - Charger un fichier Excel de conventions
   - Générer des factures individuelles ou multiples
   - Téléchargement automatique des fichiers générés

2. **Historique**
   - Consultation de toutes les factures générées
   - Filtrage par type
   - Persistance des données (même après redémarrage)

### Support

Pour toute question ou problème :
- Vérifiez que Node.js est bien installé : `node --version` dans le terminal
- Vérifiez que le port 3001 n'est pas déjà utilisé
- Consultez les logs dans la fenêtre du terminal

### Mise à jour

Pour mettre à jour l'application :
1. Remplacez tous les fichiers par la nouvelle version
2. Supprimez le dossier `node_modules/`
3. Supprimez le dossier `dist/`
4. Relancez `start-windows.bat`
