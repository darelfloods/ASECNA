import React, { useCallback, useMemo, useState, useEffect } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import { Auth } from "./components/Auth";
import { UserManagement } from "./components/UserManagement";
import { BandesModule } from "./components/BandesModule";
import { User, getAuthState, logout as authLogout, getPendingUsers } from "./services/authService";

/**
 * Calcule la durée en jours entre deux dates au format JJ/MM/AAAA
 * @param dateDepart Date de départ au format JJ/MM/AAAA
 * @param dateRetour Date de retour au format JJ/MM/AAAA
 * @returns Nombre de jours ou null si les dates sont invalides
 */
function calculateDuration(dateDepart: string, dateRetour: string): number | null {
  try {
    // Valider le format JJ/MM/AAAA
    const dateRegex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const matchDepart = dateDepart.match(dateRegex);
    const matchRetour = dateRetour.match(dateRegex);

    if (!matchDepart || !matchRetour) {
      return null;
    }

    // Extraire jour, mois, année
    const [, jourDepart, moisDepart, anneeDepart] = matchDepart;
    const [, jourRetour, moisRetour, anneeRetour] = matchRetour;

    // Créer les objets Date (mois - 1 car JavaScript compte les mois de 0 à 11)
    const dateD = new Date(parseInt(anneeDepart), parseInt(moisDepart) - 1, parseInt(jourDepart));
    const dateR = new Date(parseInt(anneeRetour), parseInt(moisRetour) - 1, parseInt(jourRetour));

    // Calculer la différence en millisecondes
    const diffMs = dateR.getTime() - dateD.getTime();

    // Convertir en jours (+1 car l'ASECNA compte le jour de départ)
    const diffJours = Math.ceil(diffMs / (1000 * 60 * 60 * 24)) + 1;

    // Retourner null si la durée est négative ou nulle
    if (diffJours <= 0) {
      return null;
    }

    return diffJours;
  } catch (error) {
    console.error('Erreur calcul durée:', error);
    return null;
  }
}
import { CELL_MAPPING } from "./mapping";
import { generateMultiInvoiceFile, generateSingleInvoiceFile, ConventionData } from "./multiInvoiceGeneratorSimple";
import { getHistory, addHistoryEntry, deleteHistoryEntry, storeDocument, getDocumentDownloadUrl, getDocumentPreviewUrl, HistoryEntry, checkAPIHealth, getFilePreviewUrl } from "./services/api";

type Row = Record<string, unknown>;

/** Champs facture dérivés : Montant HT = MONTANT, Taxe = 0, Accompte = 0, TTC = HT, Solde = MONTANT */
function buildInvoicePayload(formValues: Row): Row {
  const montant = formValues["MONTANT"];
  return {
    ...formValues,
    "Montant HT": montant,
    Taxe: 0,
    Accompte: 0,
    "Montant TTC": montant,
    Solde: montant,
  };
}

const NUMERIC_KEYS = new Set([
  "MONTANT",
  "Montant HT",
  "Taxe",
  "Accompte",
  "Montant TTC",
  "Solde",
]);

type Status =
  | { type: "idle" }
  | { type: "parsing" }
  | { type: "ready" }
  | { type: "error"; message: string }
  | { type: "generating" }
  | { type: "done" };

// ─── Composant : saisie caractère par caractère dans des cases individuelles ──
interface CharBoxInputProps {
  value: string;
  onChange: (val: string) => void;
  maxLength: number;
  style?: React.CSSProperties;
}

const CharBoxInput: React.FC<CharBoxInputProps> = ({ value, onChange, maxLength, style }) => {
  const inputRefs = React.useRef<(HTMLInputElement | null)[]>([]);
  // Tableau de maxLength cases, chacune contient UN seul caractère
  const chars = Array.from({ length: maxLength }, (_, i) => value[i] || '');

  const moveFocus = (idx: number) => {
    setTimeout(() => inputRefs.current[idx]?.focus(), 0);
  };

  const handleChange = (i: number, e: React.ChangeEvent<HTMLInputElement>) => {
    // Récupère uniquement le dernier caractère saisi (en cas de remplacement)
    const raw = e.target.value;
    if (!raw) return;
    const ch = raw.slice(-1).toUpperCase();
    const newChars = [...chars];
    newChars[i] = ch;
    onChange(newChars.join(''));
    if (i < maxLength - 1) moveFocus(i + 1);
  };

  const handleKeyDown = (i: number, e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Backspace') {
      e.preventDefault();
      const newChars = [...chars];
      if (newChars[i]) {
        newChars[i] = '';
        onChange(newChars.join(''));
      } else if (i > 0) {
        newChars[i - 1] = '';
        onChange(newChars.join(''));
        moveFocus(i - 1);
      }
    } else if (e.key === 'Delete') {
      e.preventDefault();
      const newChars = [...chars];
      newChars[i] = '';
      onChange(newChars.join(''));
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      if (i > 0) moveFocus(i - 1);
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      if (i < maxLength - 1) moveFocus(i + 1);
    }
  };

  const handlePaste = (i: number, e: React.ClipboardEvent<HTMLInputElement>) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text').replace(/\s/g, '');
    const newChars = [...chars];
    let lastIdx = i;
    for (let j = 0; j < text.length && i + j < maxLength; j++) {
      newChars[i + j] = text[j].toUpperCase();
      lastIdx = i + j;
    }
    onChange(newChars.join(''));
    moveFocus(Math.min(lastIdx + 1, maxLength - 1));
  };

  return (
    <div style={{ display: 'inline-flex', gap: '4px', alignItems: 'center', ...style }}>
      {chars.map((char, i) => (
        <input
          key={i}
          ref={el => { inputRefs.current[i] = el; }}
          type="text"
          maxLength={1}
          value={char}
          onChange={e => handleChange(i, e)}
          onKeyDown={e => handleKeyDown(i, e)}
          onPaste={e => handlePaste(i, e)}
          onFocus={e => e.currentTarget.select()}
          onClick={() => inputRefs.current[i]?.select()}
          style={{
            width: '30px',
            height: '34px',
            textAlign: 'center',
            border: '2px solid #4b5563',
            borderRadius: '4px',
            fontSize: '16px',
            fontWeight: 700,
            fontFamily: 'monospace',
            padding: 0,
            background: char ? '#eef3ff' : '#fff',
            cursor: 'text',
            color: '#1a2e4a',
            outline: 'none',
            boxSizing: 'border-box',
          }}
        />
      ))}
    </div>
  );
};
// ─────────────────────────────────────────────────────────────────────────────

const App: React.FC = () => {
  // État d'authentification
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [currentUser, setCurrentUser] = useState<User | null>(null);

  function getHistoryTimestamp(entry: { createdAt?: number; date: string }): number {
    if (typeof entry.createdAt === "number") return entry.createdAt;
    // Fallback pour les anciennes entrées stockées en "fr-FR" (ex: "11/03/2026 14:05:12")
    const m = entry.date.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      const [, dd, mm, yyyy, hh = "00", min = "00", ss = "00"] = m;
      const d = new Date(
        Number(yyyy),
        Number(mm) - 1,
        Number(dd),
        Number(hh),
        Number(min),
        Number(ss)
      );
      return d.getTime();
    }
    const parsed = Date.parse(entry.date);
    return Number.isNaN(parsed) ? 0 : parsed;
  }

  function formatHistoryDate(ts: number): string {
    const d = new Date(ts);
    return d.toLocaleString("fr-FR", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    }).replace(",", "");
  }

  function makeHistoryEntry(params: {
    type: "facture" | "fiche-mission" | "ordre-mission" | "bon-commande" | "bandes-enregistrement";
    fileName: string;
    nbConventions: number;
    status: "success" | "error";
    details?: string;
    action: string;
  }) {
    const now = Date.now();
    return {
      date: new Date(now).toISOString(), // ISO pour parsing fiable
      type: params.type,
      fileName: params.fileName,
      nbConventions: params.nbConventions,
      status: params.status,
      details: params.details,
      createdAt: now,
      action: params.action,
      actorEmail: currentUser?.email,
      actorName: currentUser
        ? currentUser.role === "admin"
          ? "admin"
          : `${currentUser.prenom} ${currentUser.nom}`.trim()
        : undefined,
      actorRole: currentUser?.role,
    };
  }

  // Vérifier l'authentification au démarrage + nettoyage des clés de session orphelines
  useEffect(() => {
    // Supprimer les clés legacy si elles existent (migration vers asecna_token/asecna_user)
    if (localStorage.getItem('token')) localStorage.removeItem('token');
    if (localStorage.getItem('user')) localStorage.removeItem('user');
    const authState = getAuthState();
    setIsAuthenticated(authState.isAuthenticated);
    setCurrentUser(authState.user);
  }, []);

  // Callback de connexion réussie
  const handleAuthSuccess = useCallback((user: User) => {
    setIsAuthenticated(true);
    setCurrentUser(user);
  }, []);

  // Callback de déconnexion
  const handleLogout = useCallback(() => {
    authLogout();
    setIsAuthenticated(false);
    setCurrentUser(null);
  }, []);

  const [file, setFile] = useState<File | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null);
  const [formValues, setFormValues] = useState<Row>({});
  const [status, setStatus] = useState<Status>({ type: "idle" });
  const [isDragging, setIsDragging] = useState(false);
  // firstInvoiceNumber est maintenant géré via formValues["N° ORDRE"]
  const [activeTab, setActiveTab] = useState<"dashboard" | "factures" | "fiches-mission" | "ordres-mission" | "bon-commande" | "historique" | "parametres" | "utilisateurs" | "bandes-enregistrement">("dashboard");
  const [pendingUsersCount, setPendingUsersCount] = useState(0);

  // Mettre à jour le compteur de demandes en attente
  useEffect(() => {
    if (isAuthenticated && currentUser?.role === 'admin') {
      const updatePendingCount = () => {
        setPendingUsersCount(getPendingUsers().length);
      };
      updatePendingCount();
      // Rafraîchir toutes les 30 secondes
      const interval = setInterval(updatePendingCount, 30000);
      return () => clearInterval(interval);
    }
  }, [isAuthenticated, currentUser, activeTab]);

  // États pour les fiches de mission
  const [ficheMissionData, setFicheMissionData] = useState<any>(null);
  const [ficheMissionStatus, setFicheMissionStatus] = useState<Status>({ type: "idle" });
  const [ficheMissionFile, setFicheMissionFile] = useState<File | null>(null);

  // États pour les ordres de mission
  const [ordreMissionData, setOrdreMissionData] = useState<any>(null);
  const [ordreMissionStatus, setOrdreMissionStatus] = useState<Status>({ type: "idle" });
  const [ordreMissionFile, setOrdreMissionFile] = useState<File | null>(null);

  // États pour les bons de commande
  const defaultLignes = [
    { description: '', quantite: '', prixUnitaire: '', total: '' },
    { description: '', quantite: '', prixUnitaire: '', total: '' },
    { description: '', quantite: '', prixUnitaire: '', total: '' },
  ];

  // Numéro de série du bon de commande (persisté en localStorage)
  const getNextBonNumero = (): string => {
    const year = new Date().getFullYear();
    const stored = localStorage.getItem('bc_last_numero');
    const last = stored ? parseInt(stored, 10) : 0;
    return `BC-${year}-${String(last + 1).padStart(4, '0')}`;
  };
  const [bonNumero, setBonNumero] = useState<string>(getNextBonNumero);
  const [showBonConfirm, setShowBonConfirm] = useState(false);

  const [bonCommandeData, setBonCommandeData] = useState({
    cs: '', cr: '', cc: '', article: '', exercice: '',
    fournisseurNom: '', fournisseurAdresse1: '', fournisseurAdresse2: '', codeFournisseur: '',
    lignes: defaultLignes,
    montantTotalChiffres: '', montantTotalLettres: '',
    delaiLivraison: '',
    lieu: 'Libreville', date: new Date().toLocaleDateString('fr-FR'),
    numeroEngagement: '', operation: '', numeroSerie: '',
    // Bon d'engagement
    numeroBon: '', codeIndividuel: '', compteLimitatif: '', operationEquipement: '',
    compteDe: '', montantAD: '', engagementsAnterieurs: '',
  });
  const [bonCommandeStatus, setBonCommandeStatus] = useState<Status>({ type: "idle" });

  // État pour le choix de format au téléchargement

  // État pour l'historique
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [apiAvailable, setApiAvailable] = useState<boolean>(false);
  const [previewModal, setPreviewModal] = useState<{
    open: boolean;
    fileName: string;
    historyId?: number;
    hasDocument?: boolean;
    html: string | null;
    type: 'docx' | 'xlsx' | 'zip' | null;
    loading: boolean;
    error: string | null;
  }>({ open: false, fileName: '', html: null, type: null, loading: false, error: null });

  const openPreview = async (item: HistoryEntry) => {
    setPreviewModal({ open: true, fileName: item.fileName, historyId: item.id, hasDocument: item.hasDocument, html: null, type: null, loading: true, error: null });
    try {
      const url = (item.hasDocument && item.id)
        ? getDocumentPreviewUrl(item.id)
        : `http://localhost:3002/api/files/${encodeURIComponent(item.fileName)}/preview-html`;
      const res = await fetch(url);
      const data = await res.json();
      if (data.success) {
        setPreviewModal(prev => ({ ...prev, html: data.html, type: data.type, loading: false }));
      } else {
        setPreviewModal(prev => ({ ...prev, error: data.error || 'Erreur de chargement', loading: false }));
      }
    } catch {
      setPreviewModal(prev => ({ ...prev, error: 'Impossible de contacter le serveur', loading: false }));
    }
  };

  const downloadFile = (item: HistoryEntry | { fileName: string; id?: number; hasDocument?: boolean }) => {
    const a = document.createElement('a');
    a.href = (item.hasDocument && item.id)
      ? getDocumentDownloadUrl(item.id)
      : `http://localhost:3002/api/files/${encodeURIComponent(item.fileName)}/download`;
    a.download = item.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // Fonction réutilisable pour recharger l'historique
  const refreshHistory = useCallback(async () => {
    if (apiAvailable) {
      const historyData = await getHistory(undefined, currentUser?.email, currentUser?.role);
      setHistory(historyData);
    }
  }, [apiAvailable, currentUser]);

  // Charger l'historique au démarrage
  useEffect(() => {
    async function loadHistory() {
      const isAvailable = await checkAPIHealth();
      setApiAvailable(isAvailable);

      if (isAvailable) {
        const historyData = await getHistory(undefined, currentUser?.email, currentUser?.role);
        setHistory(historyData);
      }
    }
    loadHistory();
  }, [currentUser]);

  // Mettre à jour le titre de l'onglet selon la page active
  useEffect(() => {
    const titles: Record<string, string> = {
      dashboard: 'ASECNA — Dashboard',
      factures: 'ASECNA — Factures',
      'fiche-mission': 'ASECNA — Fiches de mission',
      'ordre-mission': 'ASECNA — Ordres de mission',
      'bon-commande': 'ASECNA — Bons de commande',
      historique: 'ASECNA — Historique',
      parametres: 'ASECNA — Paramètres',
      users: 'ASECNA — Utilisateurs',
    };
    document.title = titles[activeTab] || 'ASECNA — Budget & Facturation';
  }, [activeTab]);

  // Recharger l'historique quand on change d'onglet
  useEffect(() => {
    if (activeTab === "historique") {
      refreshHistory();
    }
  }, [activeTab, refreshHistory]);

  // Réinitialiser les états lors d'un changement d'onglet
  useEffect(() => {
    // Réinitialisation Factures
    setFile(null);
    setRows([]);
    setSelectedIndex(null);
    setFormValues({});
    setStatus({ type: "idle" });

    // Réinitialisation Fiches de mission
    setFicheMissionData(null);
    setFicheMissionFile(null);
    setFicheMissionStatus({ type: "idle" });

    // Réinitialisation Ordres de mission
    setOrdreMissionData(null);
    setOrdreMissionFile(null);
    setOrdreMissionStatus({ type: "idle" });

    // Réinitialisation Bon de commande
    setBonCommandeStatus({ type: "idle" });
  }, [activeTab]);

  const handleFile = useCallback(async (f: File) => {
    setStatus({ type: "parsing" });
    try {
      const data = await f.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheet];

      // Lire les données brutes pour construire les en-têtes corrects
      const rawData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      if (rawData.length < 3) {
        throw new Error("Le fichier ne contient pas assez de lignes");
      }

      // Ligne 2 (index 1) contient les en-têtes principaux
      // Ligne 3 (index 2) contient les sous-en-têtes
      const mainHeaders = rawData[1]; // ["N° ORDRE", "NOM DU CLIENT", ..., "PERIODE DE VALIDITE", null, null, "MONTANT", ...]
      const subHeaders = rawData[2];  // [null, null, ..., "Date de debut", "Date de fin", "Durée", null, ...]

      // Construire les en-têtes finaux en fusionnant les deux lignes
      const headers: string[] = mainHeaders.map((header: any, index: number) => {
        const subHeader = subHeaders[index];

        // Si on a un sous-en-tête, l'utiliser
        if (subHeader && subHeader !== "") {
          return String(subHeader);
        }

        // Sinon, utiliser l'en-tête principal
        if (header && header !== "") {
          return String(header);
        }

        // Si les deux sont vides, ignorer cette colonne
        return "";
      });

      // Convertir les lignes de données (à partir de la ligne 4, index 3) en objets
      const dataRows = rawData.slice(3); // Ignorer les 3 premières lignes

      // Liste des champs qui sont des dates
      const dateFields = new Set(["Date de debut", "Date de fin"]);

      const json: Row[] = dataRows
        .filter((row) => row && row.length > 0 && row[0]) // Ignorer les lignes vides
        .map((row) => {
          const obj: Row = {};
          headers.forEach((header, index) => {
            if (header && header !== "") {
              let value = row[index] ?? "";

              // Convertir les numéros Excel en dates lisibles
              if (dateFields.has(header) && typeof value === "number") {
                // Excel stocke les dates comme nombre de jours depuis le 01/01/1900
                const excelEpoch = new Date(1900, 0, 1);
                const date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
                value = date.toLocaleDateString("fr-FR"); // Format: JJ/MM/AAAA
              }

              obj[header] = value;
            }
          });
          return obj;
        });

      setRows(json);
      setSelectedIndex(json.length > 0 ? 0 : null);
      setFormValues(json[0] ?? {});
      setStatus({ type: "ready" });
    } catch (err) {
      console.error(err);
      setRows([]);
      setSelectedIndex(null);
      setFormValues({});
      setStatus({
        type: "error",
        message:
          "Impossible de lire le fichier. Vérifiez qu'il s'agit bien d'un fichier Excel valide (.xlsx).",
      });
    }
  }, []);

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0];
      if (!f) return;
      setFile(f);
      void handleFile(f);
    },
    [handleFile]
  );

  const onDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setIsDragging(false);
      const f = e.dataTransfer.files?.[0];
      if (!f) return;
      setFile(f);
      void handleFile(f);
    },
    [handleFile]
  );

  const onDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const onRowClick = useCallback(
    (index: number) => {
      const row = rows[index];
      setSelectedIndex(index);
      setFormValues(row);
    },
    [rows]
  );

  const handleFieldChange = useCallback(
    (key: string, value: string) => {
      setFormValues((prev) => ({
        ...prev,
        [key]: value,
      }));
    },
    []
  );

  const handleGenerate = useCallback(async () => {
    if (!Object.keys(formValues).length) {
      setStatus({
        type: "error",
        message: "Aucune donnée à injecter. Veuillez d'abord charger un fichier.",
      });
      return;
    }

    setStatus({ type: "generating" });

    try {
      // Formater la donnée unique en tableau
      const montantRaw = formValues["MONTANT"];
      const montant =
        typeof montantRaw === "number"
          ? montantRaw
          : Number(String(montantRaw ?? 0).replace(/[^\d.-]/g, "")) || 0;

      const singleConvention: ConventionData[] = [{
        "NOM DU CLIENT": String(formValues["NOM DU CLIENT"] ?? "").trim(),
        "N° CONVENTION": String(formValues["N° CONVENTION"] ?? "").trim(),
        "OBJET DE LA CONVENTION": String(formValues["OBJET DE LA CONVENTION"] ?? "").trim(),
        SITE: String(formValues["SITE"] ?? formValues["Site"] ?? "").trim(),
        "Date de debut": String(formValues["Date de debut"] ?? "").trim(),
        "Date de fin": String(formValues["Date de fin"] ?? "").trim(),
        "Durée": String(formValues["Durée"] ?? "").trim(),
        MONTANT: montant,
      }];

      // Utiliser le N° ORDRE comme numéro de facture
      const invoiceNum = String(formValues["N° ORDRE"] ?? "").trim() || undefined;

      // Générer une facture unique sur une seule page (sans les blocs vides)
      const buffer = await generateSingleInvoiceFile(
        singleConvention[0],
        "/Facturation bandes d'enregistrements de 2026-V1.xlsx",
        invoiceNum
      );

      // Créer le Blob et télécharger
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const baseName =
        typeof formValues["NOM DU CLIENT"] === "string" && formValues["NOM DU CLIENT"]
          ? String(formValues["NOM DU CLIENT"]).replace(/[\\/:*?"<>|]/g, "_")
          : "facture-asecna";

      saveAs(blob, `${baseName}.xlsx`);
      setStatus({ type: "done" });

      // Ajouter à l'historique local/distant
      const historyEntry = makeHistoryEntry({
        type: "facture",
        fileName: `${baseName}.xlsx`,
        nbConventions: 1,
        status: "success",
        action: "generate_facture_single",
      });

      if (apiAvailable) {
        const entryId = await addHistoryEntry(historyEntry);
        if (entryId) await storeDocument(entryId, blob, historyEntry.fileName);
        const updatedHistory = await getHistory(undefined, currentUser?.email, currentUser?.role);
        setHistory(updatedHistory);
      } else {
        setHistory(prev => [{ ...historyEntry, id: Date.now() }, ...prev]);
      }

      // Garder l'état "done" visible pour afficher le bouton "Importer un autre fichier"
    } catch (err) {
      console.error(err);
      setStatus({
        type: "error",
        message:
          "Impossible de générer la facture. Vérifiez que le modèle Excel est présent.",
      });
    }
  }, [formValues, apiAvailable, currentUser]);

  const handleGenerateAll = useCallback(async () => {
    if (rows.length === 0) {
      setStatus({
        type: "error",
        message: "Aucune donnée à traiter. Veuillez d'abord charger un fichier.",
      });
      return;
    }

    console.log(`Début de la génération de ${rows.length} factures`);
    setStatus({ type: "generating" });

    try {
      // Charger le modèle de facture
      const response = await fetch(
        "/Facturation bandes d'enregistrements de 2026-V1.xlsx"
      );
      if (!response.ok) {
        throw new Error("Modèle de facture introuvable");
      }
      const templateBuffer = await response.arrayBuffer();
      console.log("Modèle de facture chargé");

      // Créer un ZIP pour toutes les factures
      const zip = new JSZip();

      // Générer une facture pour chaque ligne
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        console.log(`Génération facture ${i + 1}/${rows.length} pour:`, row["NOM DU CLIENT"]);

        const payload = buildInvoicePayload(row);

        // Charger une nouvelle copie du modèle pour chaque facture
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(templateBuffer);

        const sheet = workbook.worksheets[0];
        if (!sheet) {
          throw new Error("Aucune feuille dans le modèle de facture");
        }

        // Remplir les cellules
        Object.entries(CELL_MAPPING).forEach(([sourceKey, cellAddress]) => {
          const value = payload[sourceKey];
          if (value === undefined || value === null) return;
          const cell = sheet.getCell(cellAddress);
          if (NUMERIC_KEYS.has(sourceKey)) {
            const n = typeof value === "number" ? value : Number(value);
            if (!Number.isNaN(n)) {
              cell.value = n;
              return;
            }
          }
          cell.value = String(value);
        });

        // Générer le buffer Excel
        const out = await workbook.xlsx.writeBuffer();

        // Nom du fichier basé sur le nom du client ou numéro de convention
        let baseName =
          typeof row["NOM DU CLIENT"] === "string" && row["NOM DU CLIENT"]
            ? String(row["NOM DU CLIENT"]).replace(/[\\/:*?"<>|]/g, "_").trim()
            : typeof row["N° CONVENTION"] === "string" && row["N° CONVENTION"]
              ? String(row["N° CONVENTION"]).replace(/[\\/:*?"<>|]/g, "_").trim()
              : `facture-${i + 1}`;

        // Éviter les doublons en ajoutant un numéro si nécessaire
        let finalName = baseName;
        let counter = 1;
        while (zip.file(`${finalName}.xlsx`)) {
          finalName = `${baseName}_${counter}`;
          counter++;
        }

        console.log(`Ajout au ZIP: ${finalName}.xlsx`);

        // Ajouter au ZIP
        zip.file(`${finalName}.xlsx`, out);
      }

      console.log(`${Object.keys(zip.files).length} fichiers dans le ZIP`);

      // Générer le ZIP et le télécharger
      const zipBlob = await zip.generateAsync({ type: "blob" });
      const now = new Date();
      const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
      const zipFileName = `factures-asecna-${dateStr}.zip`;
      saveAs(zipBlob, zipFileName);

      console.log("ZIP téléchargé avec succès");

      // Historique: tracer qui a exporté le ZIP
      const historyEntry = makeHistoryEntry({
        type: "facture",
        fileName: zipFileName,
        nbConventions: rows.length,
        status: "success",
        action: "generate_facture_zip",
      });

      if (apiAvailable) {
        const entryId = await addHistoryEntry(historyEntry);
        if (entryId) await storeDocument(entryId, zipBlob, zipFileName);
        await refreshHistory();
      } else {
        setHistory(prev => [{ ...historyEntry, id: Date.now() }, ...prev]);
      }

      setStatus({ type: "done" });

      // Reset après 3 secondes
      setTimeout(() => {
        setStatus({ type: "ready" });
      }, 3000);
    } catch (err) {
      console.error("Erreur lors de la génération:", err);

      const historyEntry = makeHistoryEntry({
        type: "facture",
        fileName: file?.name || "ZIP factures",
        nbConventions: rows.length,
        status: "error",
        details: err instanceof Error ? err.message : "Erreur inconnue",
        action: "generate_facture_zip",
      });

      if (apiAvailable) {
        await addHistoryEntry(historyEntry);
        await refreshHistory();
      } else {
        setHistory(prev => [{ ...historyEntry, id: Date.now() }, ...prev]);
      }

      setStatus({
        type: "error",
        message:
          "Impossible de générer les factures. Vérifiez que le modèle Excel est présent.",
      });
    }
  }, [rows, apiAvailable, refreshHistory, file, currentUser]);

  const handleGenerateMulti = useCallback(async () => {
    if (rows.length === 0) {
      setStatus({
        type: "error",
        message: "Aucune donnée à traiter. Veuillez d'abord charger un fichier.",
      });
      return;
    }

    console.log(`Génération d'un fichier unique avec ${rows.length} factures`);
    setStatus({ type: "generating" });

    try {
      const conventions: ConventionData[] = rows.map((row) => {
        const montantRaw = row["MONTANT"];
        const montant =
          typeof montantRaw === "number"
            ? montantRaw
            : Number(String(montantRaw ?? 0).replace(/[^\d.-]/g, "")) || 0;

        return {
          "NOM DU CLIENT": String(row["NOM DU CLIENT"] ?? "").trim(),
          "N° CONVENTION": String(row["N° CONVENTION"] ?? "").trim(),
          "OBJET DE LA CONVENTION": String(row["OBJET DE LA CONVENTION"] ?? "").trim(),
          SITE: String(row["SITE"] ?? row["Site"] ?? "").trim(),
          "Date de debut": String(row["Date de debut"] ?? "").trim(),
          "Date de fin": String(row["Date de fin"] ?? "").trim(),
          "Durée": String(row["Durée"] ?? "").trim(),
          MONTANT: montant,
        };
      });

      // Utiliser le N° ORDRE de la première convention comme numéro de départ
      const firstOrdre = String(rows[0]?.["N° ORDRE"] ?? formValues["N° ORDRE"] ?? "").trim() || undefined;

      // Générer le fichier avec toutes les factures
      const buffer = await generateMultiInvoiceFile(
        conventions,
        "/Facturation bandes d'enregistrements de 2026-V1.xlsx",
        firstOrdre
      );

      // Télécharger le fichier
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const now = new Date();
      const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
      const fileName = `factures-asecna-${dateStr}.xlsx`;
      saveAs(blob, fileName);

      console.log("Fichier unique téléchargé avec succès");

      // Ajouter à l'historique (API ou local)
      const historyEntry = {
        ...makeHistoryEntry({
          type: "facture",
          fileName,
          nbConventions: rows.length,
          status: "success",
          action: "generate_facture_multi",
        }),
      };

      if (apiAvailable) {
        const entryId = await addHistoryEntry(historyEntry);
        if (entryId) await storeDocument(entryId, blob, fileName);
        await refreshHistory();
      } else {
        // Fallback local si API non disponible
        setHistory(prev => [{
          ...historyEntry,
          id: Date.now(),
        }, ...prev]);
      }

      setStatus({ type: "done" });

      // Reset après 3 secondes
      setTimeout(() => {
        setStatus({ type: "ready" });
      }, 3000);
    } catch (err) {
      console.error("Erreur lors de la génération multi-factures:", err);

      // Ajouter l'erreur à l'historique (API ou local)
      const historyEntry = {
        ...makeHistoryEntry({
          type: "facture",
          fileName: file?.name || "Fichier inconnu",
          nbConventions: rows.length,
          status: "error",
          details: err instanceof Error ? err.message : "Erreur inconnue",
          action: "generate_facture_multi",
        }),
      };

      if (apiAvailable) {
        await addHistoryEntry(historyEntry);
        await refreshHistory();
      } else {
        // Fallback local si API non disponible
        setHistory(prev => [{
          ...historyEntry,
          id: Date.now(),
        }, ...prev]);
      }

      setStatus({
        type: "error",
        message: "Impossible de générer le fichier multi-factures.",
      });
    }
  }, [rows, file, formValues, apiAvailable, refreshHistory, currentUser]);

  const fields = useMemo(() => Object.keys(formValues), [formValues]);

  // Calculer l'étape actuelle du stepper
  const currentStep = useMemo(() => {
    if (status.type === "idle" || status.type === "error") return 1;
    if (status.type === "parsing") return 2;
    if (status.type === "ready") return 2;
    if (status.type === "generating") return 3;
    if (status.type === "done") return 3;
    return 1;
  }, [status.type]);

  // Calculer l'étape actuelle du stepper Fiches
  const currentStepFiche = useMemo(() => {
    if (ficheMissionStatus.type === "idle" || ficheMissionStatus.type === "error") return 1;
    if (ficheMissionStatus.type === "parsing") return 2;
    if (ficheMissionStatus.type === "ready") return 2;
    if (ficheMissionStatus.type === "generating") return 3;
    if (ficheMissionStatus.type === "done") return 3;
    return 1;
  }, [ficheMissionStatus.type]);

  // Calculer l'étape actuelle du stepper Ordres
  const currentStepOrdre = useMemo(() => {
    if (ordreMissionStatus.type === "idle" || ordreMissionStatus.type === "error") return 1;
    if (ordreMissionStatus.type === "parsing") return 2;
    if (ordreMissionStatus.type === "ready") return 2;
    if (ordreMissionStatus.type === "generating") return 3;
    if (ordreMissionStatus.type === "done") return 3;
    return 1;
  }, [ordreMissionStatus.type]);

  // Statistiques pour l'état idle
  const stats = useMemo(() => {
    const thisMonth = history.filter(h => {
      const entryDate = new Date(getHistoryTimestamp(h));
      const now = new Date();
      return entryDate.getMonth() === now.getMonth() &&
        entryDate.getFullYear() === now.getFullYear();
    });

    const lastEntry = history.length > 0 ? history[0] : null;

    return {
      thisMonthCount: thisMonth.length,
      totalConventions: thisMonth.reduce((acc, h) => acc + (h.nbConventions || 0), 0),
      lastImportDate: lastEntry ? new Date(getHistoryTimestamp(lastEntry)).toISOString() : null,
      lastFileName: lastEntry?.fileName || null
    };
  }, [history]);

  // Fonction pour ouvrir le dossier des fichiers générés (Téléchargements)
  const handleOpenFilesFolder = useCallback(() => {
    // En mode Electron, on peut utiliser shell.openPath
    // En mode web, on ouvre une nouvelle fenêtre avec le protocole file://
    try {
      // Vérifier si on est dans Electron
      const isElectron = typeof window !== 'undefined' && 
        (window as any).process?.type === 'renderer' || 
        (window as any).electronAPI;
      
      if (isElectron && (window as any).electronAPI?.openDownloadsFolder) {
        // Utiliser l'API Electron si disponible
        (window as any).electronAPI.openDownloadsFolder();
      } else {
        // En mode web, afficher un message d'aide
        alert(
          "📂 Dossier des fichiers générés\n\n" +
          "Les fichiers générés sont téléchargés dans votre dossier Téléchargements par défaut.\n\n" +
          "Pour y accéder :\n" +
          "• Windows : Ouvrez l'Explorateur de fichiers → Téléchargements\n" +
          "• Mac : Ouvrez le Finder → Téléchargements\n\n" +
          "Astuce : Vous pouvez aussi cliquer sur le fichier téléchargé dans votre navigateur pour l'ouvrir directement."
        );
      }
    } catch (error) {
      console.error("Erreur lors de l'ouverture du dossier:", error);
      alert("Impossible d'ouvrir le dossier. Vérifiez votre dossier Téléchargements manuellement.");
    }
  }, []);

  // Fonction pour vider l'historique
  const handleClearHistory = useCallback(async () => {
    if (!confirm("Êtes-vous sûr de vouloir supprimer tout l'historique ?")) {
      return;
    }

    try {
      if (apiAvailable) {
        // Supprimer toutes les entrées via l'API
        for (const item of history) {
          if (item.id) {
            await deleteHistoryEntry(item.id);
          }
        }
        // Recharger l'historique
        const updatedHistory = await getHistory(undefined, currentUser?.email, currentUser?.role);
        setHistory(updatedHistory);
      } else {
        // Vider l'historique local
        setHistory([]);
      }
    } catch (error) {
      console.error("Erreur lors de la suppression de l'historique:", error);
      alert("Erreur lors de la suppression de l'historique");
    }
  }, [history, apiAvailable]);

  // Si non authentifié, afficher la page de connexion
  if (!isAuthenticated) {
    return <Auth onAuthSuccess={handleAuthSuccess} />;
  }

  // Composant Sidebar
  const Sidebar = () => (
    <div className="sidebar">
      <div className="sidebar-logo">
        <div className="sidebar-logo-icon">
          <img src="/ASECNA_logo.png" alt="ASECNA Logo" />
        </div>
        <span className="sidebar-logo-text">ASECNA</span>
      </div>
      {/* Informations utilisateur */}
      {currentUser && (
        <div className="sidebar-user-info">
          <p className="sidebar-user-name">{currentUser.prenom} {currentUser.nom}</p>
          <p className="sidebar-user-email">{currentUser.email}</p>
          <span className="sidebar-user-role">{currentUser.role}</span>
        </div>
      )}
      <nav className="sidebar-nav">
        <div
          className={`sidebar-nav-item ${activeTab === "dashboard" ? "active" : ""}`}
          onClick={() => setActiveTab("dashboard")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <rect x="3" y="3" width="7" height="7" />
            <rect x="14" y="3" width="7" height="7" />
            <rect x="14" y="14" width="7" height="7" />
            <rect x="3" y="14" width="7" height="7" />
          </svg>
          <span>Dashboard</span>
        </div>
        <div
          className={`sidebar-nav-item ${activeTab === "factures" ? "active" : ""}`}
          onClick={() => setActiveTab("factures")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <line x1="16" y1="13" x2="8" y2="13" />
            <line x1="16" y1="17" x2="8" y2="17" />
            <polyline points="10 9 9 9 8 9" />
          </svg>
          <span>Factures</span>
        </div>
        <div
          className={`sidebar-nav-item ${activeTab === "fiches-mission" ? "active" : ""}`}
          onClick={() => setActiveTab("fiches-mission")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2" />
            <rect x="9" y="3" width="6" height="4" rx="1" />
            <line x1="9" y1="12" x2="15" y2="12" />
            <line x1="9" y1="16" x2="15" y2="16" />
          </svg>
          <span>Fiches de mission</span>
        </div>
        <div
          className={`sidebar-nav-item ${activeTab === "ordres-mission" ? "active" : ""}`}
          onClick={() => setActiveTab("ordres-mission")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <path d="M9 15l2 2 4-4" />
          </svg>
          <span>Ordres de mission</span>
        </div>
        {/* Bons de commande masqués pour le MVP
        <div
          className={`sidebar-nav-item ${activeTab === "bon-commande" ? "active" : ""}`}
          onClick={() => setActiveTab("bon-commande")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <rect x="3" y="3" width="18" height="18" rx="2" />
            <path d="M3 9h18M9 21V9" />
          </svg>
          <span>Bons de commande</span>
        </div>
        */}
        <div
          className={`sidebar-nav-item ${activeTab === "bandes-enregistrement" ? "active" : ""}`}
          onClick={() => setActiveTab("bandes-enregistrement")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <rect x="2" y="7" width="20" height="10" rx="2" />
            <path d="M6 7V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v2" />
            <line x1="8" y1="12" x2="16" y2="12" />
          </svg>
          <span>Bandes d'enreg.</span>
        </div>
        <div
          className={`sidebar-nav-item ${activeTab === "historique" ? "active" : ""}`}
          onClick={() => setActiveTab("historique")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10" />
            <polyline points="12 6 12 12 16 14" />
          </svg>
          <span>Historique</span>
        </div>
        {/* Onglet Utilisateurs - visible uniquement pour les admins */}
        {currentUser?.role === 'admin' && (
          <div
            className={`sidebar-nav-item ${activeTab === "utilisateurs" ? "active" : ""}`}
            onClick={() => setActiveTab("utilisateurs")}
          >
            <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
              <circle cx="9" cy="7" r="4" />
              <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
              <path d="M16 3.13a4 4 0 0 1 0 7.75" />
            </svg>
            <span>Utilisateurs</span>
            {pendingUsersCount > 0 && (
              <span className="sidebar-badge">{pendingUsersCount}</span>
            )}
          </div>
        )}
        <div
          className={`sidebar-nav-item ${activeTab === "parametres" ? "active" : ""}`}
          onClick={() => setActiveTab("parametres")}
        >
          <svg className="sidebar-nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="3" />
            <path d="M12 1v6m0 6v6m9-9h-6m-6 0H3" />
          </svg>
          <span>Paramètres</span>
        </div>
      </nav>
      {/* Bouton de déconnexion */}
      <div className="sidebar-logout">
        <button className="sidebar-logout-btn" onClick={handleLogout}>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" strokeLinecap="round" strokeLinejoin="round" />
            <polyline points="16,17 21,12 16,7" strokeLinecap="round" strokeLinejoin="round" />
            <line x1="21" y1="12" x2="9" y2="12" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
          Déconnexion
        </button>
      </div>
    </div>
  );

  // Vue Fiches de Mission
  if (activeTab === "fiches-mission") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
          <div className="header-buttons">
            <button
              className="header-button header-button-secondary"
              onClick={handleOpenFilesFolder}
              title="Ouvrir le dossier des fichiers générés"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '16px', height: '16px', marginRight: '6px' }}>
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Voir les fichiers
            </button>
            <button
              className="header-button"
              onClick={() => {
                // Reset
                setFicheMissionData(null);
                setFicheMissionFile(null);
                setFicheMissionStatus({ type: "idle" });
              }}
            >
              Nouvelle fiche
            </button>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="welcome-header">
              <h1 className="welcome-title">Modifier des fiches de mission</h1>
              <p className="welcome-subtitle">
                Importez une fiche de mission existante au format Word (.docx), modifiez les informations et générez une nouvelle version mise à jour.
              </p>
            </div>

            {/* Stepper Fiches */}
            <div className="stepper">
              <div className={`stepper-step ${currentStepFiche >= 1 ? 'active' : ''} ${currentStepFiche > 1 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepFiche > 1 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '1'}
                </div>
                <div className="stepper-step-label">Importer</div>
              </div>
              <div className={`stepper-line ${currentStepFiche >= 2 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStepFiche >= 2 ? 'active' : ''} ${currentStepFiche > 2 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepFiche > 2 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '2'}
                </div>
                <div className="stepper-step-label">Vérifier</div>
              </div>
              <div className={`stepper-line ${currentStepFiche >= 3 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStepFiche >= 3 ? 'active' : ''} ${currentStepFiche > 3 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepFiche > 3 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '3'}
                </div>
                <div className="stepper-step-label">Générer</div>
              </div>
            </div>

            {!ficheMissionData ? (
              <>

                {/* Zone de drag & drop améliorée */}
                <div className="upload-main-section">
                  <div className="upload-primary">
                    <div className="upload-section">
                      <div className={`dropzone-compact ${isDragging ? "drag-over" : ""}`}
                        onDrop={(e) => {
                          e.preventDefault();
                          setIsDragging(false);
                          const file = e.dataTransfer.files?.[0];
                          if (file && file.name.endsWith('.docx')) {
                            setFicheMissionStatus({ type: "parsing" });
                            (async () => {
                              try {
                                const { parseFicheMission } = await import('./services/wordParser');
                                const data = await parseFicheMission(file);
                                console.log('Données extraites:', data);

                                // Calculer automatiquement la durée si les dates sont valides
                                if (data.dateDepart && data.dateRetour) {
                                  const duree = calculateDuration(data.dateDepart, data.dateRetour);
                                  if (duree !== null) {
                                    data.duree = String(duree);
                                    console.log('Durée calculée automatiquement:', duree, 'jours');
                                  }
                                }

                                setFicheMissionData(data);
                                setFicheMissionFile(file);
                                setFicheMissionStatus({ type: "ready" });
                              } catch (err: any) {
                                console.error('Erreur parsing:', err);
                                setFicheMissionStatus({ type: "error", message: err.message });
                              }
                            })();
                          }
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                          setIsDragging(true);
                        }}
                        onDragLeave={(e) => {
                          e.preventDefault();
                          setIsDragging(false);
                        }}
                        onClick={() => document.getElementById("fiche-file-input")?.click()}
                      >
                        {/* Illustration accueillante */}
                        <div className="dropzone-illustration">
                          <svg viewBox="0 0 120 120" fill="none">
                            <circle cx="60" cy="60" r="50" fill="#f5f5f5" stroke="#e5e5e5" strokeWidth="1" />
                            <path d="M45 65L60 50L75 65M60 50V80" stroke="#1a1a1a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M70 75H80C82.2091 75 84 73.2091 84 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                            <path d="M50 75H40C37.7909 75 36 73.2091 36 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                          </svg>
                        </div>
                        <p className="dropzone-text-compact">Prêt à modifier une fiche de mission ?</p>
                        <p className="dropzone-subtext-compact">Déposez votre fichier Word ou cliquez pour parcourir</p>
                        <button className="dropzone-button-compact" type="button" onClick={(e) => {
                          e.stopPropagation();
                          document.getElementById("fiche-file-input")?.click();
                        }}>
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '18px', height: '18px', marginRight: '8px' }}>
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                          Sélectionner un fichier
                        </button>
                        <button
                          className="dropzone-info-toggle"
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            const button = e.currentTarget;
                            const dropzone = button.closest('.dropzone-compact');
                            const details = dropzone?.querySelector('.dropzone-details');
                            if (details) {
                              details.classList.toggle('visible');
                            }
                          }}
                        >
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '14px', height: '14px' }}>
                            <circle cx="12" cy="12" r="10" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M12 16v-4M12 8h.01" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                          Formats acceptés
                        </button>
                        <div className="dropzone-details">
                          <p>Format accepté : <strong>.docx</strong> (Word)</p>
                          <p>Importez une fiche de mission existante au format Word.</p>
                        </div>
                        <input
                          type="file"
                          accept=".docx"
                          style={{ display: 'none' }}
                          id="fiche-file-input"
                          onChange={async (e: any) => {
                            const file = e.target?.files?.[0];
                            if (file) {
                              setFicheMissionStatus({ type: "parsing" });
                              try {
                                const { parseFicheMission } = await import('./services/wordParser');
                                const data = await parseFicheMission(file);
                                console.log('Données extraites:', data);

                                // Calculer automatiquement la durée si les dates sont valides
                                if (data.dateDepart && data.dateRetour) {
                                  const duree = calculateDuration(data.dateDepart, data.dateRetour);
                                  if (duree !== null) {
                                    data.duree = String(duree);
                                    console.log('Durée calculée automatiquement:', duree, 'jours');
                                  }
                                }

                                setFicheMissionData(data);
                                setFicheMissionFile(file);
                                setFicheMissionStatus({ type: "ready" });
                              } catch (err: any) {
                                console.error('Erreur parsing:', err);
                                setFicheMissionStatus({ type: "error", message: err.message });
                              }
                            }
                            // Reset file input
                            e.target.value = '';
                          }}
                        />
                      </div>
                    </div>
                  </div>
                </div>

                {ficheMissionStatus.type === "parsing" && (
                  <div className="status-message">
                    <div className="spinner"></div>
                    Lecture du fichier en cours...
                  </div>
                )}

                {ficheMissionStatus.type === "error" && (
                  <div className="status-message error">
                    {ficheMissionStatus.message}
                  </div>
                )}
              </>
            ) : (
              <>
                <div className="data-section">
                  <div className="section-header">
                    <h2 className="section-title">Modification de la fiche de mission</h2>
                    <div className="section-badge">
                      <span className="badge-dot"></span>
                      Prêt à modifier
                    </div>
                  </div>

                  <p className="section-subtitle">Modifiez les informations de la fiche avant de la générer</p>

                  {/* Formulaire de modification */}
                  <div className="data-table-wrapper">
                    <table className="data-table">
                      <tbody>
                        <tr>
                          <td className="table-label">Nom</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.nom}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, nom: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Matricule</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.matricule}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, matricule: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Prénom</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.prenom}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, prenom: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Emploi</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.emploi}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, emploi: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Résidence administrative</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.residence}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, residence: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Destination</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.destination}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, destination: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Motif du déplacement</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.motif}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, motif: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Date de départ</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.dateDepart}
                              onChange={(e) => {
                                const newData = { ...ficheMissionData, dateDepart: e.target.value };
                                // Calculer automatiquement la durée si les deux dates sont valides
                                if (newData.dateDepart && newData.dateRetour) {
                                  const duree = calculateDuration(newData.dateDepart, newData.dateRetour);
                                  if (duree !== null) newData.duree = String(duree);
                                }
                                setFicheMissionData(newData);
                              }}
                              placeholder="JJ/MM/AAAA"
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Date de retour</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.dateRetour}
                              onChange={(e) => {
                                const newData = { ...ficheMissionData, dateRetour: e.target.value };
                                // Calculer automatiquement la durée si les deux dates sont valides
                                if (newData.dateDepart && newData.dateRetour) {
                                  const duree = calculateDuration(newData.dateDepart, newData.dateRetour);
                                  if (duree !== null) newData.duree = String(duree);
                                }
                                setFicheMissionData(newData);
                              }}
                              placeholder="JJ/MM/AAAA"
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Durée (jours)</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.duree}
                              readOnly
                              style={{ backgroundColor: '#F1F5F9', cursor: 'not-allowed' }}
                              title="Calculé automatiquement à partir des dates"
                            />
                            <span style={{ fontSize: '11px', color: '#64748B', marginLeft: '8px' }}>
                              ⚡ Calculé automatiquement
                            </span>
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Moyen de transport</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ficheMissionData.transport}
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, transport: e.target.value })}
                            />
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Boutons d'action */}
                <div className="actions-section">
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'stretch', justifyContent: 'center' }}>
                    <div
                      className="action-card primary compact"
                      style={{
                        minWidth: '200px',
                        maxWidth: '200px',
                        display: 'flex',
                        flexDirection: 'column',
                        justifyContent: 'center'
                      }}
                      onClick={async () => {
                        setFicheMissionStatus({ type: "generating" });
                        try {
                          const { generateFicheMission } = await import('./services/wordParser');
                          const blob = await generateFicheMission(ficheMissionData, ficheMissionFile || undefined);
                          const { saveAs: save } = await import('file-saver');
                          save(blob, `Fiche_Mission_${ficheMissionData.nom}.docx`);

                          // Ajouter à l'historique
                          const entry = makeHistoryEntry({
                            type: "fiche-mission",
                            fileName: `Fiche_Mission_${ficheMissionData.nom}.docx`,
                            nbConventions: 1,
                            status: "success",
                            action: "generate_fiche_mission",
                          });

                          if (apiAvailable) {
                            const entryId = await addHistoryEntry(entry);
                            if (entryId) await storeDocument(entryId, blob, entry.fileName);
                            await refreshHistory();
                          } else {
                            setHistory(prev => [{ ...entry, id: Date.now() }, ...prev]);
                          }

                          setFicheMissionStatus({ type: "done" });
                          setTimeout(() => setFicheMissionStatus({ type: "ready" }), 3000);
                        } catch (err: any) {
                          console.error('Erreur génération:', err);
                          setFicheMissionStatus({ type: "error", message: err.message });
                        }
                      }}
                    >
                      <div className="action-card-title">Générer la fiche</div>
                      <div className="action-card-subtitle">Document modifié</div>
                    </div>
                    <button
                      className="header-button"
                      style={{
                        background: '#64748B',
                        borderColor: '#64748B',
                        minWidth: '200px',
                        padding: '12px 20px',
                        fontSize: '14px',
                        fontWeight: '600'
                      }}
                      onClick={() => {
                        setFicheMissionData(null);
                        setFicheMissionFile(null);
                        setFicheMissionStatus({ type: "idle" });
                      }}
                    >
                      Annuler la génération
                    </button>
                  </div>
                </div>

                {ficheMissionStatus.type === "done" && (
                  <div className="status-message success">
                    ✓ Fiche de mission générée et téléchargée avec succès
                  </div>
                )}

                {ficheMissionStatus.type === "error" && (
                  <div className="status-message error">
                    {ficheMissionStatus.message}
                  </div>
                )}
              </>
            )}

            <div className="footer">
              ASECNA — Service Budget et Facturation • Usage interne uniquement
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue Ordres de Mission
  if (activeTab === "ordres-mission") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
          <div className="header-buttons">
            <button
              className="header-button header-button-secondary"
              onClick={handleOpenFilesFolder}
              title="Ouvrir le dossier des fichiers générés"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '16px', height: '16px', marginRight: '6px' }}>
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Voir les fichiers
            </button>
            <button
              className="header-button"
              onClick={() => {
                setOrdreMissionData(null);
                setOrdreMissionStatus({ type: "idle" });
              }}
            >
              Nouvel ordre
            </button>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="welcome-header">
              <h1 className="welcome-title">Générer des ordres de mission</h1>
              <p className="welcome-subtitle">
                Importez une fiche de mission au format Word (.docx) pour générer automatiquement un ordre de mission avec un numéro unique.
              </p>
            </div>

            {/* Stepper Ordres */}
            <div className="stepper">
              <div className={`stepper-step ${currentStepOrdre >= 1 ? 'active' : ''} ${currentStepOrdre > 1 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepOrdre > 1 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '1'}
                </div>
                <div className="stepper-step-label">Importer</div>
              </div>
              <div className={`stepper-line ${currentStepOrdre >= 2 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStepOrdre >= 2 ? 'active' : ''} ${currentStepOrdre > 2 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepOrdre > 2 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '2'}
                </div>
                <div className="stepper-step-label">Vérifier</div>
              </div>
              <div className={`stepper-line ${currentStepOrdre >= 3 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStepOrdre >= 3 ? 'active' : ''} ${currentStepOrdre > 3 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStepOrdre > 3 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '3'}
                </div>
                <div className="stepper-step-label">Générer</div>
              </div>
            </div>

            {!ordreMissionData ? (
              <>

                {/* Zone de drag & drop améliorée */}
                <div className="upload-main-section">
                  <div className="upload-primary">
                    <div className="upload-section">
                      <div className={`dropzone-compact ${isDragging ? "drag-over" : ""}`}
                        onDrop={(e) => {
                          e.preventDefault();
                          setIsDragging(false);
                          const file = e.dataTransfer.files?.[0];
                          if (file && file.name.endsWith('.docx')) {
                            setOrdreMissionStatus({ type: "parsing" });
                            (async () => {
                              try {
                                const { parseFicheMission } = await import('./services/wordParser');
                                const data = await parseFicheMission(file);
                                console.log('Données extraites:', data);

                                // Calculer automatiquement la durée si les dates sont valides
                                if (data.dateDepart && data.dateRetour) {
                                  const duree = calculateDuration(data.dateDepart, data.dateRetour);
                                  if (duree !== null) {
                                    data.duree = String(duree);
                                    console.log('Durée calculée automatiquement:', duree, 'jours');
                                  }
                                }

                                // Ajouter les champs spécifiques à l'ordre de mission avec valeurs par défaut
                                const ordreMissionDataComplete = {
                                  ...data,
                                  cs: '050',
                                  eng: '13',
                                  cr: 'EN3',
                                  cc: '100',
                                  cl: '621',
                                  autorisationDep: '16 250 000',
                                  montantEngage: '128 000',
                                  engagementAnt: '15 879 600',
                                  disponible: '242 400',
                                  lieuSignature: 'Libreville',
                                  dateSignature: new Date().toLocaleDateString('fr-FR')
                                };

                                setOrdreMissionData(ordreMissionDataComplete);
                                setOrdreMissionFile(file);
                                setOrdreMissionStatus({ type: "ready" });
                              } catch (err: any) {
                                console.error('Erreur parsing:', err);
                                setOrdreMissionStatus({ type: "error", message: err.message });
                              }
                            })();
                          }
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                          setIsDragging(true);
                        }}
                        onDragLeave={(e) => {
                          e.preventDefault();
                          setIsDragging(false);
                        }}
                        onClick={() => document.getElementById("ordre-file-input")?.click()}
                      >
                        {/* Illustration accueillante */}
                        <div className="dropzone-illustration">
                          <svg viewBox="0 0 120 120" fill="none">
                            <circle cx="60" cy="60" r="50" fill="#f5f5f5" stroke="#e5e5e5" strokeWidth="1" />
                            <path d="M45 65L60 50L75 65M60 50V80" stroke="#1a1a1a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M70 75H80C82.2091 75 84 73.2091 84 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                            <path d="M50 75H40C37.7909 75 36 73.2091 36 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                          </svg>
                        </div>
                        <p className="dropzone-text-compact">Prêt à générer un ordre de mission ?</p>
                        <p className="dropzone-subtext-compact">Déposez votre fiche de mission Word ou cliquez pour parcourir</p>
                        <button className="dropzone-button-compact" type="button" onClick={(e) => {
                          e.stopPropagation();
                          document.getElementById("ordre-file-input")?.click();
                        }}>
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '18px', height: '18px', marginRight: '8px' }}>
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                          Sélectionner un fichier
                        </button>
                        <button
                          className="dropzone-info-toggle"
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            const button = e.currentTarget;
                            const dropzone = button.closest('.dropzone-compact');
                            const details = dropzone?.querySelector('.dropzone-details');
                            if (details) {
                              details.classList.toggle('visible');
                            }
                          }}
                        >
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '14px', height: '14px' }}>
                            <circle cx="12" cy="12" r="10" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M12 16v-4M12 8h.01" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                          Formats acceptés
                        </button>
                        <div className="dropzone-details">
                          <p>Format accepté : <strong>.docx</strong> (Word)</p>
                          <p>Importez une fiche de mission existante au format Word.</p>
                        </div>
                        <input
                          type="file"
                          accept=".docx"
                          style={{ display: 'none' }}
                          id="ordre-file-input"
                          onChange={async (e: any) => {
                            const file = e.target?.files?.[0];
                            if (file) {
                              setOrdreMissionStatus({ type: "parsing" });
                              try {
                                const { parseFicheMission } = await import('./services/wordParser');
                                const data = await parseFicheMission(file);
                                console.log('Données extraites:', data);

                                // Calculer automatiquement la durée si les dates sont valides
                                if (data.dateDepart && data.dateRetour) {
                                  const duree = calculateDuration(data.dateDepart, data.dateRetour);
                                  if (duree !== null) {
                                    data.duree = String(duree);
                                    console.log('Durée calculée automatiquement:', duree, 'jours');
                                  }
                                }

                                // Ajouter les champs spécifiques à l'ordre de mission avec valeurs par défaut
                                const ordreMissionDataComplete = {
                                  ...data,
                                  cs: '050',
                                  eng: '13',
                                  cr: 'EN3',
                                  cc: '100',
                                  cl: '621',
                                  autorisationDep: '16 250 000',
                                  montantEngage: '128 000',
                                  engagementAnt: '15 879 600',
                                  disponible: '242 400',
                                  lieuSignature: 'Libreville',
                                  dateSignature: new Date().toLocaleDateString('fr-FR')
                                };

                                setOrdreMissionData(ordreMissionDataComplete);
                                setOrdreMissionFile(file);
                                setOrdreMissionStatus({ type: "ready" });
                              } catch (err: any) {
                                console.error('Erreur parsing:', err);
                                setOrdreMissionStatus({ type: "error", message: err.message });
                              }
                            }
                            // Reset file input
                            e.target.value = '';
                          }}
                        />
                      </div>
                    </div>
                  </div>
                </div>

                {ordreMissionStatus.type === "parsing" && (
                  <div className="status-message">
                    <div className="spinner"></div>
                    Lecture du fichier en cours...
                  </div>
                )}

                {ordreMissionStatus.type === "error" && (
                  <div className="status-message error">
                    {ordreMissionStatus.message}
                  </div>
                )}
              </>
            ) : (
              <>
                <div className="data-section">
                  <div className="section-header">
                    <h2 className="section-title">Données de la fiche de mission</h2>
                    <div className="section-badge">
                      <span className="badge-dot"></span>
                      Prêt à générer
                    </div>
                  </div>

                  <p className="section-subtitle">Vérifiez les informations avant de générer l'ordre de mission</p>

                  {/* Tableau de données */}
                  <div className="data-table-wrapper">
                    <table className="data-table">
                      <tbody>
                        <tr>
                          <td className="table-label">Nom</td>
                          <td>{ordreMissionData.nom}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Matricule</td>
                          <td>{ordreMissionData.matricule}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Prénom</td>
                          <td>{ordreMissionData.prenom}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Emploi</td>
                          <td>{ordreMissionData.emploi}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Résidence administrative</td>
                          <td>{ordreMissionData.residence}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Destination</td>
                          <td>{ordreMissionData.destination}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Motif du déplacement</td>
                          <td>{ordreMissionData.motif}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Date de départ</td>
                          <td>{ordreMissionData.dateDepart}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Date de retour</td>
                          <td>{ordreMissionData.dateRetour}</td>
                        </tr>
                        <tr>
                          <td className="table-label">Durée</td>
                          <td>{ordreMissionData.duree} jours</td>
                        </tr>
                        <tr>
                          <td className="table-label">Moyen de transport</td>
                          <td>{ordreMissionData.transport}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Section Informations Budgétaires */}
                  <div className="section-header" style={{ marginTop: '32px' }}>
                    <h3 className="section-title" style={{ fontSize: '16px' }}>Informations budgétaires</h3>
                  </div>

                  <div className="data-table-wrapper">
                    <table className="data-table">
                      <tbody>
                        <tr>
                          <td className="table-label">CS</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.cs || '050'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, cs: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Eng</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.eng || '13'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, eng: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">CR</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.cr || 'EN3'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, cr: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">CC</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.cc || '100'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, cc: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">CL</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.cl || '621'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, cl: e.target.value })}
                            />
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Section Montants Financiers */}
                  <div className="section-header" style={{ marginTop: '32px' }}>
                    <h3 className="section-title" style={{ fontSize: '16px' }}>Montants financiers</h3>
                  </div>

                  <div className="data-table-wrapper">
                    <table className="data-table">
                      <tbody>
                        <tr>
                          <td className="table-label">Autorisation de dép</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.autorisationDep || '16 250 000'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, autorisationDep: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Montant engagé</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.montantEngage || '128 000'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, montantEngage: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Engagement Ant</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.engagementAnt || '15 879 600'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, engagementAnt: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Disponible</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.disponible || '242 400'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, disponible: e.target.value })}
                            />
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Section Signature */}
                  <div className="section-header" style={{ marginTop: '32px' }}>
                    <h3 className="section-title" style={{ fontSize: '16px' }}>Signature</h3>
                  </div>

                  <div className="data-table-wrapper">
                    <table className="data-table">
                      <tbody>
                        <tr>
                          <td className="table-label">Lieu de signature</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.lieuSignature || 'Libreville'}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, lieuSignature: e.target.value })}
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="table-label">Date de signature</td>
                          <td>
                            <input
                              type="text"
                              className="form-input"
                              value={ordreMissionData.dateSignature || new Date().toLocaleDateString('fr-FR')}
                              onChange={(e) => setOrdreMissionData({ ...ordreMissionData, dateSignature: e.target.value })}
                              placeholder="JJ/MM/AAAA"
                            />
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Boutons d'action */}
                <div className="actions-section">
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'stretch', justifyContent: 'center' }}>
                    <div
                      className="action-card primary compact"
                      style={{
                        minWidth: '200px',
                        maxWidth: '200px',
                        display: 'flex',
                        flexDirection: 'column',
                        justifyContent: 'center'
                      }}
                      onClick={async () => {
                        setOrdreMissionStatus({ type: "generating" });
                        try {
                          const { generateOrdreMission } = await import('./services/wordParser');
                          const year = new Date().getFullYear();
                          const numero = `${String(Math.floor(Math.random() * 999) + 1).padStart(3, '0')}/${year}`;

                          const blob = await generateOrdreMission(ordreMissionData, numero);
                          const { saveAs: save } = await import('file-saver');
                          save(blob, `Ordre_Mission_${ordreMissionData.nom}_${numero.replace('/', '-')}.docx`);

                          // Ajouter à l'historique
                          const entry = makeHistoryEntry({
                            type: "ordre-mission",
                            fileName: `Ordre_Mission_${ordreMissionData.nom}.docx`,
                            nbConventions: 1,
                            status: "success",
                            action: "generate_ordre_mission",
                          });

                          if (apiAvailable) {
                            const entryId = await addHistoryEntry(entry);
                            if (entryId) await storeDocument(entryId, blob, entry.fileName);
                            await refreshHistory();
                          } else {
                            setHistory(prev => [{ ...entry, id: Date.now() }, ...prev]);
                          }

                          setOrdreMissionStatus({ type: "done" });
                          setTimeout(() => setOrdreMissionStatus({ type: "ready" }), 3000);
                        } catch (err: any) {
                          console.error('Erreur génération:', err);
                          setOrdreMissionStatus({ type: "error", message: err.message });
                        }
                      }}
                    >
                      <div className="action-card-title">Générer l'ordre</div>
                      <div className="action-card-subtitle">Avec numéro unique</div>
                    </div>
                    <button
                      className="header-button"
                      style={{
                        background: '#64748B',
                        borderColor: '#64748B',
                        minWidth: '200px',
                        padding: '12px 20px',
                        fontSize: '14px',
                        fontWeight: '600'
                      }}
                      onClick={() => {
                        setOrdreMissionData(null);
                        setOrdreMissionStatus({ type: "idle" });
                      }}
                    >
                      Annuler la génération
                    </button>
                  </div>
                </div>

                {ordreMissionStatus.type === "done" && (
                  <div className="status-message success">
                    ✓ Ordre de mission généré et téléchargé avec succès
                  </div>
                )}

                {ordreMissionStatus.type === "error" && (
                  <div className="status-message error">
                    {ordreMissionStatus.message}
                  </div>
                )}
              </>
            )}

            <div className="footer">
              ASECNA — Service Budget et Facturation • Usage interne uniquement
            </div>
          </div>
        </div>
      </>
    );
  }

  // ── Vue Bon de commande ────────────────────────────────────────────────────
  if (activeTab === "bon-commande") {
    const setLigne = (index: number, field: string, value: string) => {
      setBonCommandeData(prev => ({
        ...prev,
        lignes: prev.lignes.map((l, i) => {
          if (i !== index) return l;
          const updated = { ...l, [field]: value };
          if (field === 'quantite' || field === 'prixUnitaire') {
            const qty = parseFloat(field === 'quantite' ? value : l.quantite) || 0;
            const prix = parseFloat(field === 'prixUnitaire' ? value : l.prixUnitaire) || 0;
            updated.total = qty > 0 && prix > 0 ? String(Math.round(qty * prix)) : '';
          }
          return updated;
        }),
      }));
    };

    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
          <div className="header-buttons">
            <button
              className="header-button header-button-secondary"
              onClick={handleOpenFilesFolder}
              title="Ouvrir le dossier des fichiers générés"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '16px', height: '16px', marginRight: '6px' }}>
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Voir les fichiers
            </button>
            <button
              className="header-button"
              onClick={() => {
                setBonCommandeData({
                  cs: '', cr: '', cc: '', article: '', exercice: '',
                  fournisseurNom: '', fournisseurAdresse1: '', fournisseurAdresse2: '', codeFournisseur: '',
                  lignes: [
                    { description: '', quantite: '', prixUnitaire: '', total: '' },
                    { description: '', quantite: '', prixUnitaire: '', total: '' },
                    { description: '', quantite: '', prixUnitaire: '', total: '' },
                  ],
                  montantTotalChiffres: '', montantTotalLettres: '',
                  delaiLivraison: '',
                  lieu: 'Libreville', date: new Date().toLocaleDateString('fr-FR'),
                  numeroEngagement: '', operation: '', numeroSerie: '',
                  numeroBon: '', codeIndividuel: '', compteLimitatif: '', operationEquipement: '',
                  compteDe: '', montantAD: '', engagementsAnterieurs: '',
                });
                setBonCommandeStatus({ type: "idle" });
              }}
            >
              Nouveau bon
            </button>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="welcome-header">
              <h1 className="welcome-title">Générer un bon de commande</h1>
              <p className="welcome-subtitle">
                Remplissez le formulaire ci-dessous pour générer un bon de commande ASECNA au format Word.
              </p>
            </div>

            <div className="data-section">
              {/* Numéro de série */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '20px', padding: '10px 14px', background: '#f0f4ff', borderRadius: '8px', border: '1px solid #c7d4f0' }}>
                <span style={{ fontWeight: 600, fontSize: '13px', color: '#2a4a8a', whiteSpace: 'nowrap' }}>N° du bon :</span>
                <input
                  type="text"
                  className="form-input"
                  value={bonNumero}
                  onChange={e => setBonNumero(e.target.value)}
                  style={{ width: '180px', fontWeight: 600, letterSpacing: '0.5px' }}
                />
                <span style={{ fontSize: '12px', color: '#6b7280' }}>Numéro séquentiel — modifiable si nécessaire</span>
              </div>

              {/* Section 1 : Codes comptables */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Codes comptables</h2>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Centre de Synthèse (C.S)</td>
                      <td><CharBoxInput maxLength={3} value={bonCommandeData.cs} onChange={val => setBonCommandeData({ ...bonCommandeData, cs: val })} /></td>
                      <td className="table-label">Centre de Responsabilité (C.R)</td>
                      <td><CharBoxInput maxLength={3} value={bonCommandeData.cr} onChange={val => setBonCommandeData({ ...bonCommandeData, cr: val })} /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Centre de Coût (C.C)</td>
                      <td><CharBoxInput maxLength={3} value={bonCommandeData.cc} onChange={val => setBonCommandeData({ ...bonCommandeData, cc: val })} /></td>
                      <td className="table-label">Article</td>
                      <td><CharBoxInput maxLength={3} value={bonCommandeData.article} onChange={val => setBonCommandeData({ ...bonCommandeData, article: val })} /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Exercice</td>
                      <td><CharBoxInput maxLength={4} value={bonCommandeData.exercice} onChange={val => setBonCommandeData({ ...bonCommandeData, exercice: val })} /></td>
                      <td className="table-label"></td>
                      <td></td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Section 2 : Fournisseur */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Fournisseur</h2>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Nom / Raison sociale</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.fournisseurNom} onChange={e => setBonCommandeData({ ...bonCommandeData, fournisseurNom: e.target.value })} placeholder="Nom ou raison sociale du fournisseur" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Adresse (ligne 1)</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.fournisseurAdresse1} onChange={e => setBonCommandeData({ ...bonCommandeData, fournisseurAdresse1: e.target.value })} placeholder="Ex: Quartier, BP, Rue..." /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Adresse (ligne 2)</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.fournisseurAdresse2} onChange={e => setBonCommandeData({ ...bonCommandeData, fournisseurAdresse2: e.target.value })} placeholder="Ex: Ville, Code postal..." /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Code fournisseur</td>
                      <td><CharBoxInput maxLength={4} value={bonCommandeData.codeFournisseur} onChange={val => setBonCommandeData({ ...bonCommandeData, codeFournisseur: val })} /></td>
                      <td></td><td></td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Section 3 : Détail de la commande */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Détail de la commande</h2>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th style={{ textAlign: 'left', padding: '8px', fontWeight: 600 }}>Description</th>
                      <th style={{ textAlign: 'center', padding: '8px', fontWeight: 600, width: '110px' }}>Quantité</th>
                      <th style={{ textAlign: 'center', padding: '8px', fontWeight: 600, width: '140px' }}>Prix unitaire</th>
                      <th style={{ textAlign: 'center', padding: '8px', fontWeight: 600, width: '140px' }}>Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bonCommandeData.lignes.map((ligne, i) => (
                      <tr key={i}>
                        <td style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <input type="text" className="form-input" style={{ width: '100%' }} value={ligne.description} onChange={e => setLigne(i, 'description', e.target.value)} placeholder={`Article / prestation ${i + 1}`} />
                          {bonCommandeData.lignes.length > 1 && (
                            <button type="button" onClick={() => setBonCommandeData(prev => ({ ...prev, lignes: prev.lignes.filter((_, idx) => idx !== i) }))} style={{ flexShrink: 0, width: '24px', height: '24px', border: 'none', background: '#fee2e2', color: '#dc2626', borderRadius: '4px', cursor: 'pointer', fontSize: '16px', lineHeight: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }} title="Supprimer cette ligne">×</button>
                          )}
                        </td>
                        <td><input type="text" className="form-input" style={{ textAlign: 'center' }} value={ligne.quantite} onChange={e => setLigne(i, 'quantite', e.target.value)} placeholder="0" /></td>
                        <td><input type="text" className="form-input" style={{ textAlign: 'right' }} value={ligne.prixUnitaire} onChange={e => setLigne(i, 'prixUnitaire', e.target.value)} placeholder="0" /></td>
                        <td><input type="text" className="form-input" readOnly style={{ textAlign: 'right', background: '#f8fafc', color: '#374151' }} value={ligne.total} placeholder="Auto" /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <button
                  type="button"
                  onClick={() => setBonCommandeData(prev => ({ ...prev, lignes: [...prev.lignes, { description: '', quantite: '', prixUnitaire: '', total: '' }] }))}
                  style={{ marginTop: '8px', padding: '6px 14px', background: '#eff6ff', color: '#1d4ed8', border: '1px dashed #93c5fd', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: 500 }}
                >
                  + Ajouter une ligne
                </button>
              </div>

              {/* Section 4 : Montants et délai */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Montants et délai</h2>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Montant total (en chiffres)</td>
                      <td colSpan={3}>
                        <input
                          type="text"
                          className="form-input"
                          style={{ width: '100%', background: '#f8fafc', fontWeight: 600 }}
                          readOnly
                          value={bonCommandeData.lignes
                            .reduce((sum, l) => sum + (parseFloat(l.total) || 0), 0)
                            .toLocaleString('fr-FR')}
                          placeholder="Calculé automatiquement"
                        />
                      </td>
                    </tr>
                    <tr>
                      <td className="table-label">Montant total (en lettres)</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.montantTotalLettres} onChange={e => setBonCommandeData({ ...bonCommandeData, montantTotalLettres: e.target.value })} placeholder="ex: Cinq cent mille francs CFA" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Délai de livraison</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.delaiLivraison} onChange={e => setBonCommandeData({ ...bonCommandeData, delaiLivraison: e.target.value })} placeholder="ex: 30 jours" /></td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Section 5 : Validation */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Validation</h2>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Lieu</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.lieu} onChange={e => setBonCommandeData({ ...bonCommandeData, lieu: e.target.value })} placeholder="ex: Libreville" /></td>
                      <td className="table-label">Date</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.date} onChange={e => setBonCommandeData({ ...bonCommandeData, date: e.target.value })} placeholder="JJ/MM/AAAA" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">N° d'engagement</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.numeroEngagement} onChange={e => setBonCommandeData({ ...bonCommandeData, numeroEngagement: e.target.value })} placeholder="ex: ENG-2026-001" /></td>
                      <td className="table-label">Opération</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.operation} onChange={e => setBonCommandeData({ ...bonCommandeData, operation: e.target.value })} placeholder="ex: Achat fournitures" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">N° de série (tampon)</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.numeroSerie} onChange={e => setBonCommandeData({ ...bonCommandeData, numeroSerie: e.target.value })} placeholder="ex: BC-2026-042" /></td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Section 6 : Bon d'engagement */}
              <div className="section-header" style={{ marginBottom: '12px' }}>
                <h2 className="section-title">Bon d'engagement</h2>
                <p style={{ fontSize: '12px', color: '#6b7280', marginTop: '4px' }}>
                  Coupon interne — ne pas envoyer au fournisseur. Les codes comptables et le fournisseur sont repris automatiquement.
                </p>
              </div>
              <div className="data-table-wrapper" style={{ marginBottom: '24px' }}>
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Code individuel</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.codeIndividuel} onChange={e => setBonCommandeData({ ...bonCommandeData, codeIndividuel: e.target.value })} placeholder="ex: 0042" /></td>
                      <td className="table-label">Compte limitatif</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.compteLimitatif} onChange={e => setBonCommandeData({ ...bonCommandeData, compteLimitatif: e.target.value })} placeholder="ex: 621" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Opération d'équipement</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.operationEquipement} onChange={e => setBonCommandeData({ ...bonCommandeData, operationEquipement: e.target.value })} placeholder="ex: Achat matériel informatique" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Compte de comptabilité générale</td>
                      <td colSpan={3}><input type="text" className="form-input" style={{ width: '100%' }} value={bonCommandeData.compteDe} onChange={e => setBonCommandeData({ ...bonCommandeData, compteDe: e.target.value })} placeholder="ex: 4411" /></td>
                    </tr>
                    <tr>
                      <td className="table-label">Montant A.D.</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.montantAD} onChange={e => setBonCommandeData({ ...bonCommandeData, montantAD: e.target.value })} placeholder="Autorisation de dépenses" /></td>
                      <td className="table-label">Engagements antérieurs</td>
                      <td><input type="text" className="form-input" value={bonCommandeData.engagementsAnterieurs} onChange={e => setBonCommandeData({ ...bonCommandeData, engagementsAnterieurs: e.target.value })} placeholder="ex: 0" /></td>
                    </tr>
                    <tr>
                      <td className="table-label" style={{ color: '#6b7280' }}>Cumul des engagements</td>
                      <td>
                        <input type="text" className="form-input" readOnly style={{ background: '#f8fafc', color: '#374151' }}
                          value={(() => {
                            const bon = parseFloat((bonCommandeData.lignes.reduce((s, l) => s + (parseFloat(l.total) || 0), 0)).toString()) || 0;
                            const ant = parseFloat((bonCommandeData.engagementsAnterieurs || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
                            return bon + ant > 0 ? (bon + ant).toLocaleString('fr-FR') : '';
                          })()} placeholder="Calculé automatiquement" />
                      </td>
                      <td className="table-label" style={{ color: '#6b7280' }}>Disponible</td>
                      <td>
                        <input type="text" className="form-input" readOnly style={{ background: '#f8fafc', color: '#374151' }}
                          value={(() => {
                            const ad  = parseFloat((bonCommandeData.montantAD || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
                            const bon = parseFloat((bonCommandeData.lignes.reduce((s, l) => s + (parseFloat(l.total) || 0), 0)).toString()) || 0;
                            const ant = parseFloat((bonCommandeData.engagementsAnterieurs || '').replace(/\s/g, '').replace(/,/g, '.')) || 0;
                            const dispo = ad - (bon + ant);
                            return ad > 0 ? dispo.toLocaleString('fr-FR') : '';
                          })()} placeholder="Calculé automatiquement" />
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {/* Bouton de génération */}
            <div className="actions-section">
              <div style={{ display: 'flex', gap: '16px', alignItems: 'stretch', justifyContent: 'center' }}>
                <div
                  className="action-card primary compact"
                  style={{ minWidth: '220px', maxWidth: '220px', display: 'flex', flexDirection: 'column', justifyContent: 'center', cursor: 'pointer' }}
                  onClick={() => setShowBonConfirm(true)}
                >
                  <div className="action-card-title">Générer le bon</div>
                  <div className="action-card-subtitle">Vérifier et télécharger</div>
                </div>
              </div>
            </div>

            {bonCommandeStatus.type === "done" && (
              <div className="status-message success">
                ✓ Bon de commande {bonNumero} généré et téléchargé avec succès
              </div>
            )}
            {bonCommandeStatus.type === "error" && (
              <div className="status-message error">
                {(bonCommandeStatus as any).message}
              </div>
            )}

            <div className="footer">
              ASECNA — Service Budget et Facturation • Usage interne uniquement
            </div>
          </div>
        </div>

        {/* Modale de confirmation */}
        {showBonConfirm && (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
            <div style={{ background: '#fff', borderRadius: '12px', padding: '28px 32px', maxWidth: '480px', width: '90%', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }}>
              <h2 style={{ margin: '0 0 20px', fontSize: '18px', fontWeight: 700, color: '#1a2e4a' }}>Confirmer le bon de commande</h2>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px', marginBottom: '20px' }}>
                <tbody>
                  <tr style={{ borderBottom: '1px solid #e5e7eb' }}>
                    <td style={{ padding: '8px 0', color: '#6b7280', width: '45%' }}>N° du bon</td>
                    <td style={{ padding: '8px 0', fontWeight: 600, color: '#1a2e4a' }}>{bonNumero}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e5e7eb' }}>
                    <td style={{ padding: '8px 0', color: '#6b7280' }}>Fournisseur</td>
                    <td style={{ padding: '8px 0', fontWeight: 600 }}>{bonCommandeData.fournisseurNom || '—'}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e5e7eb' }}>
                    <td style={{ padding: '8px 0', color: '#6b7280' }}>Montant (chiffres)</td>
                    <td style={{ padding: '8px 0' }}>{bonCommandeData.montantTotalChiffres || '—'}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e5e7eb' }}>
                    <td style={{ padding: '8px 0', color: '#6b7280' }}>Montant (lettres)</td>
                    <td style={{ padding: '8px 0' }}>{bonCommandeData.montantTotalLettres || '—'}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e5e7eb' }}>
                    <td style={{ padding: '8px 0', color: '#6b7280' }}>Délai livraison</td>
                    <td style={{ padding: '8px 0' }}>{bonCommandeData.delaiLivraison || '—'}</td>
                  </tr>
                  <tr>
                    <td style={{ padding: '8px 0', color: '#6b7280' }}>Lieu / Date</td>
                    <td style={{ padding: '8px 0' }}>{bonCommandeData.lieu} — {bonCommandeData.date}</td>
                  </tr>
                </tbody>
              </table>
              <p style={{ margin: '0 0 20px', fontSize: '13px', color: '#6b7280' }}>
                Veuillez vérifier les informations ci-dessus. Une fois confirmé, le document sera généré et téléchargé.
              </p>
              <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
                <button
                  style={{ padding: '9px 20px', borderRadius: '7px', border: '1px solid #d1d5db', background: '#f9fafb', cursor: 'pointer', fontSize: '14px' }}
                  onClick={() => setShowBonConfirm(false)}
                >
                  Annuler
                </button>
                <button
                  style={{ padding: '9px 20px', borderRadius: '7px', border: 'none', background: '#1a4a8a', color: '#fff', cursor: 'pointer', fontSize: '14px', fontWeight: 600 }}
                  onClick={async () => {
                    setShowBonConfirm(false);
                    setBonCommandeStatus({ type: "generating" });
                    try {
                      const { fillBonCommandeWord } = await import('./bonCommandeWordGenerator');
                      // Charger le template Word A4
                      const response = await fetch('/BON DE COMMANDE A4.docx');
                      if (!response.ok) throw new Error('Impossible de charger le template Word');
                      const docxBuffer = await response.arrayBuffer();

                      // Calculer le montant total
                      const montantCalcule = bonCommandeData.lignes
                        .reduce((sum, l) => sum + (parseFloat(l.total) || 0), 0)
                        .toLocaleString('fr-FR');

                      // Remplir le template Word avec les données du formulaire
                      const blob = await fillBonCommandeWord(docxBuffer, {
                        ...bonCommandeData,
                        numeroBon: bonNumero,
                        montantTotalChiffres: montantCalcule,
                      } as any);

                      const fournisseur = bonCommandeData.fournisseurNom.trim() || 'Fournisseur';
                      const dateStr = new Date().toISOString().slice(0, 10);
                      saveAs(blob, `Bon_Commande_${bonNumero}_${fournisseur}_${dateStr}.docx`);

                      // Incrémenter et sauvegarder le numéro de série
                      const numMatch = bonNumero.match(/(\d+)$/);
                      if (numMatch) {
                        localStorage.setItem('bc_last_numero', numMatch[1]);
                      }
                      // Préparer le prochain numéro
                      setBonNumero(getNextBonNumero());

                      const entry = makeHistoryEntry({
                        type: "bon-commande",
                        fileName: `Bon_Commande_${bonNumero}_${fournisseur}.docx`,
                        nbConventions: 1,
                        status: "success",
                        action: "generate_bon_commande",
                      });
                      if (apiAvailable) {
                        const entryId = await addHistoryEntry(entry);
                        if (entryId) await storeDocument(entryId, blob, entry.fileName);
                        await refreshHistory();
                      } else {
                        setHistory(prev => [{ ...entry, id: Date.now() }, ...prev]);
                      }
                      setBonCommandeStatus({ type: "done" });
                      setTimeout(() => setBonCommandeStatus({ type: "idle" }), 4000);
                    } catch (err: any) {
                      console.error('Erreur génération bon de commande:', err);
                      setBonCommandeStatus({ type: "error", message: err.message });
                    }
                  }}
                >
                  Confirmer et télécharger
                </button>
              </div>
            </div>
          </div>
        )}
      </>
    );
  }

  // Vue Historique
  if (activeTab === "historique") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Historique des activités</span>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="data-section">
              <div className="section-header">
                <h2 className="section-title">Historique</h2>
                <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                  <div
                    className={`section-badge ${history.length > 0 ? "section-badge--info" : "section-badge--muted"
                      }`}
                  >
                    {history.length > 0 ? (
                      <>
                        <span className="badge-dot"></span>
                        {history.length} activité{history.length > 1 ? "s" : ""}
                      </>
                    ) : (
                      "Aucune activité"
                    )}
                  </div>
                  {history.length > 0 && (
                    <>
                      <button
                        className="header-button"
                        style={{ fontSize: '13px', padding: '6px 12px' }}
                        onClick={() => {
                          const latest = history[0];
                          alert(`Dernier import:\n\nDate: ${latest.date}\nType: ${latest.type}\nFichier: ${latest.fileName}\nConventions: ${latest.nbConventions}\nStatut: ${latest.status}`);
                        }}
                      >
                        Voir le dernier
                      </button>
                      <button
                        className="header-button"
                        style={{
                          fontSize: '13px',
                          padding: '6px 12px',
                          borderColor: '#EF4444',
                          color: '#EF4444'
                        }}
                        onClick={async () => {
                          console.log('Bouton vider cliqué');
                          console.log('Historique actuel:', history.length, 'entrées');
                          console.log('API disponible:', apiAvailable);

                          if (!confirm('Êtes-vous sûr de vouloir vider tout l\'historique ? Cette action est irréversible.')) {
                            console.log('Annulé par l\'utilisateur');
                            return;
                          }

                          try {
                            if (apiAvailable) {
                              console.log('Suppression via API...');
                              // Supprimer toutes les entrées via l'API
                              let deleted = 0;
                              for (const item of history) {
                                if (item.id) {
                                  const success = await deleteHistoryEntry(item.id);
                                  if (success) deleted++;
                                  console.log(`Suppression ID ${item.id}:`, success);
                                }
                              }
                              console.log(`${deleted} entrées supprimées`);

                              // Recharger l'historique
                              const updatedHistory = await getHistory(undefined, currentUser?.email, currentUser?.role);
                              console.log('Nouvel historique:', updatedHistory.length, 'entrées');
                              setHistory(updatedHistory);
                            } else {
                              console.log('Suppression locale...');
                              // Vider l'historique local
                              setHistory([]);
                            }

                            alert('Historique vidé avec succès');
                          } catch (error) {
                            console.error('Erreur lors du vidage:', error);
                            alert('Erreur lors du vidage de l\'historique');
                          }
                        }}
                      >
                        Vider l'historique
                      </button>
                    </>
                  )}
                </div>
              </div>

              <p className="section-subtitle">
                Consultez l'historique de toutes vos générations de factures
              </p>

              {/* Filtre par type */}
              <div className="convention-select-wrapper">
                <label className="convention-select-label">Filtrer par type</label>
                <select
                  className="convention-select"
                  onChange={(e) => {
                    const type = e.target.value;
                    if (type === "all") {
                      getHistory(undefined, currentUser?.email, currentUser?.role).then(setHistory).catch(console.error);
                    } else {
                      getHistory(undefined, currentUser?.email, currentUser?.role).then(allHistory => {
                        setHistory(allHistory.filter(h => h.type === type));
                      }).catch(console.error);
                    }
                  }}
                >
                  <option value="all">Tous les types</option>
                  <option value="facture">Factures</option>
                  <option value="fiche-mission">Fiches de mission</option>
                  <option value="ordre-mission">Ordres de mission</option>
                  <option value="bandes-enregistrement">Bandes d'enregistrement</option>
                  {/* <option value="bon-commande">Bons de commande</option> */}
                </select>
              </div>

              {/* Tableau d'historique */}
              {history.length === 0 ? (
                <div style={{
                  textAlign: "center",
                  padding: "60px 20px",
                  color: "#64748B",
                  fontSize: "15px"
                }}>
                  <svg
                    style={{
                      width: "48px",
                      height: "48px",
                      margin: "0 auto 16px",
                      color: "#CBD5E1"
                    }}
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                  >
                    <circle cx="12" cy="12" r="10" />
                    <polyline points="12 6 12 12 16 14" />
                  </svg>
                  <p>Aucune activité enregistrée pour le moment</p>
                  <p style={{ fontSize: "13px", marginTop: "8px" }}>
                    Générez des factures pour voir l'historique ici
                  </p>
                </div>
              ) : (
                <div className="data-table-wrapper" style={{ overflowX: 'auto' }}>
                  <table className="data-table" style={{ minWidth: '700px' }}>
                    <thead>
                      <tr>
                        <th style={{ whiteSpace: 'nowrap' }}>Date</th>
                        <th>Utilisateur</th>
                        <th>Type</th>
                        <th>Fichier</th>
                        <th style={{ whiteSpace: 'nowrap' }}>Conv.</th>
                        <th style={{ whiteSpace: 'nowrap' }}>Statut</th>
                        <th style={{ whiteSpace: 'nowrap' }}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {history.map((item) => (
                        <tr key={item.id}>
                          <td>{formatHistoryDate(getHistoryTimestamp(item))}</td>
                          <td>
                            {currentUser?.role === 'admin'
                              ? (item.actorName || item.actorEmail || "—")
                              : item.actorEmail
                                ? "Vous"
                                : "—"
                            }
                            {currentUser?.role === 'admin' && item.action ? (
                              <div style={{ fontSize: "11px", color: "#64748B", marginTop: "2px" }}>
                                {item.action === "generate_bon_commande" ? "Génération bon de commande"
                                  : item.action === "generate_facture" ? "Génération facture"
                                  : item.action === "generate_fiche_mission" ? "Génération fiche de mission"
                                  : item.action === "generate_ordre_mission" ? "Génération ordre de mission"
                                  : item.action === "generate_facture_bandes" ? "Génération facture bandes"
                                  : item.action === "generate_factures_bandes_multi" ? "Génération factures bandes (multi)"
                                  : item.action === "generate_bordereau_bandes" ? "Génération bordereau bandes"
                                  : item.action}
                              </div>
                            ) : null}
                          </td>
                          <td>
                            <span style={{
                              background: item.type === "facture" ? "#EFF6FF" : item.type === "fiche-mission" ? "#F0FDF4" : item.type === "bon-commande" ? "#FFF7ED" : item.type === "bandes-enregistrement" ? "#FDF4FF" : "#FEF2F2",
                              color: item.type === "facture" ? "#1E40AF" : item.type === "fiche-mission" ? "#166534" : item.type === "bon-commande" ? "#C2410C" : item.type === "bandes-enregistrement" ? "#7C3AED" : "#991B1B",
                              padding: "4px 10px",
                              borderRadius: "4px",
                              fontSize: "12px",
                              fontWeight: "500"
                            }}>
                              {item.type === "facture" ? "Facture" : item.type === "fiche-mission" ? "Fiche de mission" : item.type === "bon-commande" ? "Bon de commande" : item.type === "bandes-enregistrement" ? "Bandes d'enreg." : "Ordre de mission"}
                            </span>
                          </td>
                          <td>{item.fileName}</td>
                          <td>{item.nbConventions}</td>
                          <td>
                            {item.status === "success" ? (
                              <span style={{
                                background: "#ECFDF5",
                                color: "#059669",
                                padding: "4px 10px",
                                borderRadius: "4px",
                                fontSize: "12px",
                                fontWeight: "500"
                              }}>
                                ✓ Succès
                              </span>
                            ) : (
                              <span style={{
                                background: "#FEF2F2",
                                color: "#DC2626",
                                padding: "4px 10px",
                                borderRadius: "4px",
                                fontSize: "12px",
                                fontWeight: "500"
                              }}>
                                ✗ Erreur
                              </span>
                            )}
                          </td>
                          <td>
                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                              <button
                                className="hist-action-btn hist-preview-btn"
                                disabled={item.status === 'error'}
                                title={item.status === 'error' ? 'Aucun fichier généré' : 'Prévisualiser'}
                                onClick={() => openPreview(item)}
                              >
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                                  <circle cx="12" cy="12" r="3" />
                                </svg>
                                Voir
                              </button>
                              <button
                                className="hist-action-btn hist-download-btn"
                                disabled={item.status === 'error'}
                                title={item.hasDocument ? 'Retélécharger depuis la base de données' : 'Télécharger depuis Téléchargements'}
                                onClick={() => downloadFile(item)}
                              >
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                                  <polyline points="7 10 12 15 17 10" />
                                  <line x1="12" y1="15" x2="12" y2="3" />
                                </svg>
                                {item.hasDocument ? 'Retélécharger' : 'Télécharger'}
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <div className="footer">
              ASECNA — Service Budget et Facturation • Usage interne uniquement
            </div>
          </div>
        </div>

        {/* Modal de prévisualisation */}
        {previewModal.open && (
          <div className="hist-modal-overlay" onClick={() => setPreviewModal(p => ({ ...p, open: false }))}>
            <div className="hist-modal" onClick={e => e.stopPropagation()}>
              <div className="hist-modal-header">
                <div className="hist-modal-title">
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                  </svg>
                  <span>{previewModal.fileName}</span>
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  {!previewModal.loading && !previewModal.error && previewModal.type !== 'zip' && (
                    <button
                      className="hist-modal-action-btn"
                      onClick={() => {
                        const win = window.open('', '_blank');
                        if (win) {
                          win.document.write(`<html><head><title>${previewModal.fileName}</title><style>body{font-family:Arial,sans-serif;padding:20px;}table{border-collapse:collapse;width:100%;}td,th{border:1px solid #ccc;padding:6px 10px;}</style></head><body>${previewModal.html}</body></html>`);
                          win.document.close();
                          win.print();
                        }
                      }}
                    >
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                        <polyline points="6 9 6 2 18 2 18 9" />
                        <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
                        <rect x="6" y="14" width="12" height="8" />
                      </svg>
                      Imprimer / PDF
                    </button>
                  )}
                  <button
                    className="hist-modal-action-btn"
                    onClick={() => downloadFile({ fileName: previewModal.fileName, id: previewModal.historyId, hasDocument: previewModal.hasDocument })}
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                      <polyline points="7 10 12 15 17 10" />
                      <line x1="12" y1="15" x2="12" y2="3" />
                    </svg>
                    Télécharger
                  </button>
                  <button
                    className="hist-modal-close"
                    onClick={() => setPreviewModal(p => ({ ...p, open: false }))}
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18">
                      <line x1="18" y1="6" x2="6" y2="18" />
                      <line x1="6" y1="6" x2="18" y2="18" />
                    </svg>
                  </button>
                </div>
              </div>

              <div className="hist-modal-body">
                {previewModal.loading && (
                  <div className="hist-modal-state">
                    <div className="hist-modal-spinner" />
                    <p>Chargement de la prévisualisation...</p>
                  </div>
                )}
                {previewModal.error && (
                  <div className="hist-modal-state">
                    <svg viewBox="0 0 24 24" fill="none" stroke="#EF4444" strokeWidth="2" width="40" height="40">
                      <circle cx="12" cy="12" r="10" />
                      <line x1="12" y1="8" x2="12" y2="12" />
                      <line x1="12" y1="16" x2="12.01" y2="16" />
                    </svg>
                    <p style={{ color: '#EF4444', marginTop: '12px' }}>{previewModal.error}</p>
                    {!previewModal.hasDocument && (
                      <p style={{ fontSize: '13px', color: '#64748B' }}>
                        Le fichier doit être présent dans le dossier Téléchargements
                      </p>
                    )}
                  </div>
                )}
                {!previewModal.loading && !previewModal.error && previewModal.type === 'zip' && (
                  <div className="hist-modal-state">
                    <svg viewBox="0 0 24 24" fill="none" stroke="#94A3B8" strokeWidth="2" width="48" height="48">
                      <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                    </svg>
                    <p style={{ marginTop: '12px', color: '#475569' }}>Les archives ZIP ne peuvent pas être prévisualisées.</p>
                    <p style={{ fontSize: '13px', color: '#64748B' }}>Utilisez le bouton Télécharger pour l'ouvrir.</p>
                  </div>
                )}
                {!previewModal.loading && !previewModal.error && previewModal.html && (
                  <div
                    className={`hist-modal-content ${previewModal.type === 'xlsx' ? 'xlsx-preview' : 'docx-preview'}`}
                    dangerouslySetInnerHTML={{ __html: previewModal.html }}
                  />
                )}
              </div>
            </div>
          </div>
        )}
      </>
    );
  }



  // Vue Utilisateurs (admin)
  if (activeTab === "utilisateurs") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Gestion des utilisateurs</span>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <UserManagement currentUser={currentUser} />
          </div>
        </div>
      </>
    );
  }


  // Vue Paramètres
  if (activeTab === "parametres") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Paramètres</span>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="data-section">
              <div className="section-header">
                <h2 className="section-title">Configuration de l'application</h2>
              </div>

              <p className="section-subtitle">
                Gérez les paramètres de l'application ASECNA Facturation
              </p>

              <div className="data-table-wrapper">
                <table className="data-table">
                  <tbody>
                    <tr>
                      <td className="table-label">Version de l'application</td>
                      <td>1.0.0</td>
                    </tr>
                    <tr>
                      <td className="table-label">API disponible</td>
                      <td>{apiAvailable ? "✓ Connectée" : "✗ Non disponible"}</td>
                    </tr>
                    <tr>
                      <td className="table-label">Entrées dans l'historique</td>
                      <td>{history.length}</td>
                    </tr>
                    <tr>
                      <td className="table-label">Dernier fichier généré</td>
                      <td>{history.length > 0 ? history[0].fileName : file?.name || "Aucun"}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div className="actions-section" style={{ marginTop: '32px' }}>
                <div className="actions-grid">
                  <div
                    className="action-card secondary"
                    onClick={handleClearHistory}
                  >
                    <div className="action-card-title">Vider l'historique</div>
                    <div className="action-card-subtitle">Supprimer toutes les entrées</div>
                  </div>
                  <div
                    className="action-card secondary"
                    onClick={() => window.location.reload()}
                  >
                    <div className="action-card-title">Réinitialiser l'application</div>
                    <div className="action-card-subtitle">Redémarrer depuis zéro</div>
                  </div>
                </div>
              </div>
            </div>

            <div className="footer">
              ASECNA — SERVICE BUDGET ET FACTURATION • Usage interne uniquement
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue Bandes d'enregistrement
  if (activeTab === "bandes-enregistrement") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
        </header>
        <div className="app-container">
          <div className="content-wrapper" style={{ maxWidth: 1200 }}>
            <div style={{ padding: '20px 24px 0', borderBottom: '1px solid #E2E8F0' }}>
              <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#1A2B4A' }}>
                Gestion des Comptoirs d'Enregistrement
              </h2>
              <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748B' }}>
                Saisie des fiches, génération de factures et bordereau d'émission
              </p>
            </div>
            <BandesModule onHistoryAdd={async (fileName: string, action: string, nbConventions?: number) => {
              const entry = makeHistoryEntry({
                type: "bandes-enregistrement",
                fileName,
                nbConventions: nbConventions ?? 1,
                status: "success",
                action,
              });
              if (apiAvailable) {
                await addHistoryEntry(entry);
                const updated = await getHistory(undefined, currentUser?.email, currentUser?.role);
                setHistory(updated);
              } else {
                setHistory(prev => [{ ...entry, id: Date.now() }, ...prev]);
              }
            }} />
          </div>
        </div>
      </>
    );
  }

  // Vue Dashboard
  if (activeTab === "dashboard") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Dashboard</span>
          </div>
          <div className="header-buttons">
            <button
              className="header-button header-button-secondary"
              onClick={handleOpenFilesFolder}
              title="Ouvrir le dossier des fichiers générés"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '16px', height: '16px', marginRight: '6px' }}>
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Voir les fichiers
            </button>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            <div className="welcome-header">
              <h1 className="welcome-title">Bienvenue sur ASECNA Facturation</h1>
              <p className="welcome-subtitle">
                Gérez vos factures, fiches de mission et ordres de mission en toute simplicité
              </p>
            </div>

            {history.length > 0 && (
              <div className="stats-section">
                <div className="stats-header">
                  <h2 className="stats-title">Statistiques</h2>
                </div>
                <div className="stats-cards">
                  <div className="stat-card">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value">{stats.thisMonthCount}</div>
                      <div className="stat-card-label">Fichiers ce mois</div>
                    </div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value">{stats.totalConventions}</div>
                      <div className="stat-card-label">Conventions traitées</div>
                    </div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value">{history.length}</div>
                      <div className="stat-card-label">Total activités</div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            <div className="actions-section" style={{ marginTop: '40px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: 600, color: '#1a1a1a', marginBottom: '20px' }}>Actions rapides</h3>
              <div className="actions-grid">
                <div className="action-card primary" onClick={() => setActiveTab("factures")}>
                  <div className="action-card-title">Générer des factures</div>
                  <div className="action-card-subtitle">Importer un fichier Excel</div>
                </div>
                <div className="action-card secondary" onClick={() => setActiveTab("fiches-mission")}>
                  <div className="action-card-title">Créer une fiche de mission</div>
                  <div className="action-card-subtitle">Importer un document Word</div>
                </div>
                <div className="action-card secondary" onClick={() => setActiveTab("historique")}>
                  <div className="action-card-title">Consulter l'historique</div>
                  <div className="action-card-subtitle">{history.length} entrée{history.length > 1 ? 's' : ''}</div>
                </div>
              </div>
            </div>

            <div className="footer">
              ASECNA — SERVICE BUDGET ET FACTURATION • Usage interne uniquement
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue principale : dropzone
  if ((status.type === "idle" || (status.type === "error" && !file)) && activeTab === "factures") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
          <div className="header-buttons">
            <button
              className="header-button header-button-secondary"
              onClick={handleOpenFilesFolder}
              title="Ouvrir le dossier des fichiers générés"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '16px', height: '16px', marginRight: '6px' }}>
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              Voir les fichiers
            </button>
            <button className="header-button" onClick={() => window.location.reload()}>
              Nouveau fichier
            </button>
          </div>
        </header>

        <div className="app-container">
          <div className="content-wrapper">
            {/* En-tête d'accueil */}
            <div className="welcome-header">
              <h1 className="welcome-title">Générer une facture</h1>
              <p className="welcome-subtitle">
                Déposez votre fichier Excel de conventions pour générer automatiquement vos factures ASECNA
              </p>
            </div>

            {/* Stepper */}
            <div className="stepper">
              <div className={`stepper-step ${currentStep >= 1 ? 'active' : ''} ${currentStep > 1 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStep > 1 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '1'}
                </div>
                <div className="stepper-step-label">Importer</div>
              </div>
              <div className={`stepper-line ${currentStep >= 2 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStep >= 2 ? 'active' : ''} ${currentStep > 2 ? 'completed' : ''}`}>
                <div className="stepper-step-number">
                  {currentStep > 2 ? (
                    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                      <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                    </svg>
                  ) : '2'}
                </div>
                <div className="stepper-step-label">Vérifier</div>
              </div>
              <div className={`stepper-line ${currentStep >= 3 ? 'active' : ''}`}></div>
              <div className={`stepper-step ${currentStep >= 3 ? 'active' : ''}`}>
                <div className="stepper-step-number">3</div>
                <div className="stepper-step-label">Générer</div>
              </div>
            </div>

            {/* Statistiques du système */}
            {history.length > 0 && (
              <div className="stats-section">
                <div className="stats-header">
                  <h2 className="stats-title">Vue d'ensemble</h2>
                  <button className="clear-history-btn" onClick={handleClearHistory}>
                    <svg viewBox="0 0 20 20" fill="currentColor">
                      <path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm5-1a1 1 0 00-1 1v6a1 1 0 102 0V8a1 1 0 00-1-1z" clipRule="evenodd" />
                    </svg>
                    Vider l'historique
                  </button>
                </div>
                <div className="stats-cards">
                  <div className="stat-card">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value">{stats.thisMonthCount}</div>
                      <div className="stat-card-label">Fichiers ce mois</div>
                    </div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value">{stats.totalConventions}</div>
                      <div className="stat-card-label">Conventions traitées</div>
                    </div>
                  </div>
                  <div className="stat-card stat-card-wide">
                    <div className="stat-card-icon">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <div className="stat-card-content">
                      <div className="stat-card-value-small">
                        {stats.lastImportDate ? new Date(stats.lastImportDate).toLocaleDateString('fr-FR', {
                          day: 'numeric',
                          month: 'short'
                        }) : '—'}
                      </div>
                      <div className="stat-card-label">Dernière importation</div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Zone principale avec drop zone et fichiers récents */}
            <div className="upload-main-section">
              <div className="upload-primary">
                <div className="upload-section">
                  <div
                    className={`dropzone-compact ${isDragging ? "drag-over" : ""}`}
                    onDrop={onDrop}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onClick={() => document.getElementById("file-input")?.click()}
                  >
                    {/* Illustration accueillante */}
                    <div className="dropzone-illustration">
                      <svg viewBox="0 0 120 120" fill="none">
                        <circle cx="60" cy="60" r="50" fill="#f5f5f5" stroke="#e5e5e5" strokeWidth="1" />
                        <path d="M45 65L60 50L75 65M60 50V80" stroke="#1a1a1a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                        <path d="M70 75H80C82.2091 75 84 73.2091 84 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                        <path d="M50 75H40C37.7909 75 36 73.2091 36 71V60" stroke="#1a1a1a" strokeWidth="2" strokeLinecap="round" />
                      </svg>
                    </div>
                    <p className="dropzone-text-compact">Prêt à traiter vos factures du mois ?</p>
                    <p className="dropzone-subtext-compact">Déposez votre fichier Excel ou cliquez pour parcourir</p>
                    <button className="dropzone-button-compact" type="button" onClick={(e) => {
                      e.stopPropagation();
                      document.getElementById("file-input")?.click();
                    }}>
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '18px', height: '18px', marginRight: '8px' }}>
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      Sélectionner un fichier
                    </button>
                    <button
                      className="dropzone-info-toggle"
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        const button = e.currentTarget;
                        const dropzone = button.closest('.dropzone-compact');
                        const details = dropzone?.querySelector('.dropzone-details');
                        if (details) {
                          details.classList.toggle('visible');
                        }
                      }}
                    >
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '14px', height: '14px' }}>
                        <circle cx="12" cy="12" r="10" strokeLinecap="round" strokeLinejoin="round" />
                        <path d="M12 16v-4M12 8h.01" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      Formats acceptés
                    </button>
                    <div className="dropzone-details">
                      <p>Format accepté : <strong>.xlsx</strong> (Excel 2007+)</p>
                      <p>Le fichier doit contenir les colonnes standards des conventions ASECNA</p>
                    </div>
                    <input
                      type="file"
                      accept=".xlsx"
                      onChange={onFileChange}
                      className="dropzone-input"
                      id="file-input"
                    />
                  </div>
                </div>

                {/* Raccourcis contextuels */}
                <div className="context-shortcuts">
                  <button
                    className="shortcut-chip"
                    onClick={() => setActiveTab("historique")}
                    disabled={history.length === 0}
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    Voir le dernier import
                  </button>
                  <a
                    href="/CONVENTIONS DOMANIALES ACTUALISEES 2025 (Enregistré automatiquement).xlsx"
                    download
                    className="shortcut-chip"
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    Télécharger modèle Excel
                  </a>
                  <button className="shortcut-chip" onClick={() => setActiveTab("parametres")}>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" strokeLinecap="round" strokeLinejoin="round" />
                      <path d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    Paramètres
                  </button>
                </div>
              </div>

              {/* Section fichiers récents — Factures uniquement */}
              {history.filter(h => h.type === 'facture').length > 0 && (
                <div className="recent-files-section">
                  <h3 className="recent-files-title">Fichiers récents</h3>
                  <div className="recent-files-list">
                    {history.filter(h => h.type === 'facture').slice(0, 5).map((item) => (
                      <div
                        key={item.id}
                        className="recent-file-item"
                        onClick={() => {
                          // Rediriger vers l'historique pour voir les détails
                          setActiveTab("historique");
                        }}
                        title="Cliquez pour voir dans l'historique"
                      >
                        <div className="recent-file-icon">
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </div>
                        <div className="recent-file-content">
                          <div className="recent-file-name">{item.fileName}</div>
                          <div className="recent-file-meta">
                            {item.nbConventions} convention{item.nbConventions > 1 ? 's' : ''} • {
                              new Date(getHistoryTimestamp(item)).toLocaleDateString('fr-FR', {
                                day: 'numeric',
                                month: 'short',
                                hour: '2-digit',
                                minute: '2-digit'
                              })
                            }
                          </div>
                        </div>
                        <div className={`recent-file-status ${item.status}`}>
                          {item.status === 'success' ? (
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                              <path d="M5 13l4 4L19 7" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          ) : (
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                              <path d="M6 18L18 6M6 6l12 12" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {status.type === "error" && (
              <div className="status-message error">
                {status.message}
              </div>
            )}

            <div className="footer">
              ASECNA — SERVICE BUDGET ET FACTURATION • Usage interne uniquement
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue de chargement
  if (status.type === "parsing") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
          </div>
        </header>
        <div className="app-container">
          <div className="content-wrapper">
            <div className="status-message generating">
              <div className="spinner"></div>
              <span style={{ marginLeft: "8px" }}>Lecture du fichier en cours...</span>
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue de génération
  if (status.type === "generating") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Génération de factures</span>
          </div>
        </header>
        <div className="app-container">
          <div className="content-wrapper">
            <div className="status-message generating">
              <div className="spinner"></div>
              <div style={{ marginLeft: "8px" }}>
                <div>
                  {rows.length > 1
                    ? `Génération de ${rows.length} factures en cours...`
                    : "Génération de la facture Excel..."}
                </div>
                {rows.length > 1 && (
                  <div style={{ fontSize: "12px", opacity: 0.8, marginTop: "4px" }}>
                    Cela peut prendre quelques instants, veuillez patienter...
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </>
    );
  }

  // Vue principale avec données
  return (
    <>
      <Sidebar />
      <header className="institutional-header">
        <div className="header-left">
          <span className="header-title">SERVICE BUDGET ET FACTURATION</span>
        </div>
        <button
          className="header-button"
          onClick={() => {
            setFile(null);
            setRows([]);
            setSelectedIndex(null);
            setFormValues({});
            setStatus({ type: "idle" });
          }}
        >
          Nouveau fichier
        </button>
      </header>

      <div className="app-container">
        <div className="content-wrapper">
          <div className="welcome-header">
            <h1 className="welcome-title">Générer une facture</h1>
            <p className="welcome-subtitle">
              Déposez votre fichier Excel de conventions pour générer automatiquement vos factures ASECNA
            </p>
          </div>

          {/* Stepper Factures (Vue données/génération) */}
          <div className="stepper">
            <div className={`stepper-step ${currentStep >= 1 ? 'active' : ''} ${currentStep > 1 ? 'completed' : ''}`}>
              <div className="stepper-step-number">
                {currentStep > 1 ? (
                  <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                ) : '1'}
              </div>
              <div className="stepper-step-label">Importer</div>
            </div>
            <div className={`stepper-line ${currentStep >= 2 ? 'active' : ''}`}></div>
            <div className={`stepper-step ${currentStep >= 2 ? 'active' : ''} ${currentStep > 2 ? 'completed' : ''}`}>
              <div className="stepper-step-number">
                {currentStep > 2 ? (
                  <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                ) : '2'}
              </div>
              <div className="stepper-step-label">Vérifier</div>
            </div>
            <div className={`stepper-line ${currentStep >= 3 ? 'active' : ''}`}></div>
            <div className={`stepper-step ${currentStep >= 3 ? 'active' : ''} ${currentStep > 3 ? 'completed' : ''}`}>
              <div className="stepper-step-number">
                {currentStep > 3 ? (
                  <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '16px', height: '16px' }}>
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                ) : '3'}
              </div>
              <div className="stepper-step-label">Générer</div>
            </div>
          </div>

          <div className="data-section">
            <div className="section-header">
              <h2 className="section-title">Convention chargée</h2>
              <div className="section-badge section-badge--info">
                <span className="badge-dot"></span>
                {rows.length} convention{rows.length > 1 ? "s" : ""}
              </div>
            </div>

            <p className="section-subtitle">{file?.name}</p>

            {/* Sélection de convention si plusieurs */}
            {rows.length > 1 && (
              <div className="convention-select-wrapper">
                <label className="convention-select-label">Sélectionner une convention</label>
                <select
                  className="convention-select"
                  value={selectedIndex ?? 0}
                  onChange={(e) => onRowClick(Number(e.target.value))}
                >
                  {rows.map((row, index) => (
                    <option key={index} value={index}>
                      Convention {index + 1} - {String(row["NOM DU CLIENT"] ?? row["N° CONVENTION"] ?? `Ligne ${index + 1}`)}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Tableau de données */}
            <div className="data-table-wrapper">
              <table className="data-table">
                <tbody>
                  {fields.map((key) => (
                    <tr key={key}>
                      <td className="table-label">{key}</td>
                      <td>
                        {key === "N° ORDRE" ? (
                          <input
                            type="text"
                            className="form-input"
                            style={{ width: '100%', fontWeight: 600, fontSize: '13px' }}
                            value={String(formValues[key] ?? "")}
                            onChange={e => handleFieldChange(key, e.target.value)}
                            placeholder="ex: N°2026/023/ASECNA/DGAN/CAF"
                            title="Numéro de facture (les suivants seront incrémentés automatiquement)"
                          />
                        ) : (
                          String(formValues[key] ?? "—")
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Boutons de génération */}
          <div className="actions-section">
            <div style={{ display: 'flex', gap: '16px', alignItems: 'stretch', justifyContent: 'center' }}>
              <div
                className="action-card primary compact"
                style={{
                  minWidth: '200px',
                  maxWidth: '200px',
                  display: 'flex',
                  flexDirection: 'column',
                  justifyContent: 'center'
                }}
                onClick={() => void handleGenerate()}
              >
                <div className="action-card-title">Générer la facture</div>
              </div>

              {rows.length > 1 && (
                <div
                  className="action-card primary compact"
                  style={{
                    minWidth: '200px',
                    maxWidth: '200px',
                    display: 'flex',
                    flexDirection: 'column',
                    justifyContent: 'center'
                  }}
                  onClick={() => void handleGenerateMulti()}
                >
                  <div className="action-card-title">Toutes les factures</div>
                </div>
              )}

              <button
                className="header-button"
                style={{
                  background: '#64748B',
                  borderColor: '#64748B',
                  minWidth: '200px',
                  padding: '12px 20px',
                  fontSize: '14px',
                  fontWeight: '600'
                }}
                onClick={() => {
                  setFile(null);
                  setRows([]);
                  setSelectedIndex(null);
                  setFormValues({});
                  setStatus({ type: "idle" });
                }}
              >
                Annuler la génération
              </button>
            </div>
          </div>

          {/* Message de succès */}
          {status.type === "done" && (
            <div className="status-message success">
              ✓ Facture générée et téléchargée avec succès
            </div>
          )}

          {/* Message d'erreur */}
          {status.type === "error" && (
            <div className="status-message error">
              {status.message}
            </div>
          )}

          <div className="footer">
            ASECNA — SERVICE BUDGET ET FACTURATION • Usage interne uniquement
          </div>
        </div>
      </div>
    </>
  );
};

export default App;

