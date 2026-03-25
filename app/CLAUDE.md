# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

**ASECNA Facturation** — Application desktop (Electron) + web pour la génération de factures, fiches de mission et ordres de mission pour l'ASECNA (Agence pour la Sécurité de la Navigation Aérienne). Interface en français.

## Commandes

```bash
# Développement (frontend + backend en parallèle)
npm run dev

# Frontend seul (port 5173)
npm run dev:frontend

# Backend seul (port 3001)
npm run server

# Build frontend
npm run build

# Electron en dev
npm run electron:dev

# Build installateur Windows
npm run build:win
```

Pas de tests automatisés configurés dans ce projet.

## Architecture

**Stack :**
- Frontend : React 18 + TypeScript, bundlé avec Vite (SWC)
- Backend : Express.js sur le port 3001
- Desktop : Electron 40
- Persistence : fichier JSON (`server/history.json`), pas de base de données SQL
- Documents : ExcelJS/XLSX pour Excel, docxtemplater/PizzIP pour Word

**Organisation :**
```
src/          → Frontend React/TypeScript
server/       → Backend Express (standalone-commonjs.js = point d'entrée)
electron/     → Processus principal Electron
public/       → Templates Excel/Word et assets statiques
```

**Flux de données :**
1. Le frontend appelle le backend via `src/services/api.ts`
2. Le backend expose `/api/history` (routes dans `server/routes/history.js`)
3. La génération de documents se fait côté frontend : les templates publics sont chargés, remplis, puis téléchargés en ZIP
4. L'authentification est gérée côté client dans `src/services/authService.ts` avec persistance localStorage

**Points critiques :**
- `src/App.tsx` (~3000 lignes) contient l'essentiel de la logique applicative et de l'UI — c'est le composant central
- `src/multiInvoiceGeneratorSimple.ts` (~1000 lignes) gère la génération batch de factures Excel
- `src/services/wordParser.ts` parse les champs des templates Word pour substitution
- L'authentification est **client-side uniquement** (localStorage + hachage côté client) — non adapté à la production

**Types de documents générés :**
- `facture` — Factures Excel depuis template `Facturation bandes d'enregistrements de 2026-V1.xlsx`
- `fiche-mission` — Depuis `FICHE DE MISSION.docx`
- `ordre-mission` — Depuis `ORDRE DE MISSION.docx`

**Roles utilisateur :** `admin`, `user`, `viewer` — gérés dans `src/components/UserManagement.tsx`

## Modèles de données clés

```typescript
// Utilisateur
{ id, email, nom, prenom, role, status, matricule, service, createdAt, lastLogin }

// Entrée historique
{ id, date, type, fileName, nbConventions, status, details, createdAt, actorInfo }

// Convention (données facture)
{ clientName, conventionNumber, object, site, startDate, endDate, duration, amount }
```

## Notes importantes

- Le fichier `IDENTIFIANTS_ADMIN.txt` contient des identifiants en clair — ne jamais committer dans un repo public
- `server/standalone-commonjs.js` est le point d'entrée serveur réel (pas `index.js`)
- Le build Electron produit un installateur NSIS dans `release/`
- L'application est entièrement localisée en français
