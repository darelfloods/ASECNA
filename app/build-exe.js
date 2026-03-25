import { exec } from 'child_process';
import { promisify } from 'util';
import { copyFileSync, mkdirSync, existsSync } from 'fs';
import { join } from 'path';

const execPromise = promisify(exec);

async function buildExecutable() {
  console.log('🔨 Construction de l\'exécutable Windows...');
  
  // Créer le dossier de distribution
  const distDir = 'dist';
  if (!existsSync(distDir)) {
    mkdirSync(distDir);
  }
  
  try {
    // 1. Build du frontend avec Vite
    console.log('📦 Build du frontend...');
    await execPromise('npm run build');
    
    // 2. Créer un serveur qui sert à la fois l'API et le frontend
    console.log('🔧 Préparation du serveur...');
    
    // 3. Packager avec pkg
    console.log('📦 Création de l\'exécutable...');
    await execPromise('pkg server/index.js --target node18-win-x64 --output dist/asecna-facturation.exe');
    
    console.log('✅ Exécutable créé avec succès : dist/asecna-facturation.exe');
    console.log('📁 Copiez le dossier dist/ pour distribuer l\'application');
    
  } catch (error) {
    console.error('❌ Erreur lors de la construction:', error);
  }
}

buildExecutable();
