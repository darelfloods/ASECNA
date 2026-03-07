import React, { useCallback, useMemo, useState, useEffect } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import { CELL_MAPPING } from "./mapping";
import { generateMultiInvoiceFile, ConventionData } from "./multiInvoiceGeneratorSimple";
import { getHistory, addHistoryEntry, deleteHistoryEntry, HistoryEntry, checkAPIHealth } from "./services/api";

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

const App: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null);
  const [formValues, setFormValues] = useState<Row>({});
  const [status, setStatus] = useState<Status>({ type: "idle" });
  const [isDragging, setIsDragging] = useState(false);
  const [activeTab, setActiveTab] = useState<"dashboard" | "factures" | "fiches-mission" | "ordres-mission" | "historique" | "parametres">("factures");

  // États pour les fiches de mission
  const [ficheMissionData, setFicheMissionData] = useState<any>(null);
  const [ficheMissionStatus, setFicheMissionStatus] = useState<Status>({ type: "idle" });

  // États pour les ordres de mission
  const [ordreMissionData, setOrdreMissionData] = useState<any>(null);
  const [ordreMissionStatus, setOrdreMissionStatus] = useState<Status>({ type: "idle" });

  // État pour l'historique
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [apiAvailable, setApiAvailable] = useState<boolean>(false);

  // Charger l'historique au démarrage
  useEffect(() => {
    async function loadHistory() {
      const isAvailable = await checkAPIHealth();
      setApiAvailable(isAvailable);

      if (isAvailable) {
        const historyData = await getHistory();
        setHistory(historyData);
      }
    }
    loadHistory();
  }, []);

  // Recharger l'historique quand on change d'onglet
  useEffect(() => {
    if (activeTab === "historique" && apiAvailable) {
      async function refreshHistory() {
        const historyData = await getHistory();
        setHistory(historyData);
      }
      refreshHistory();
    }
  }, [activeTab, apiAvailable]);

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

      // Générer le fichier via le générateur centralisé
      const buffer = await generateMultiInvoiceFile(
        singleConvention,
        "/Facturation bandes d'enregistrements de 2026.xlsx"
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
      const historyEntry = {
        date: new Date().toLocaleString('fr-FR'),
        type: "facture" as const,
        fileName: `${baseName}.xlsx`,
        nbConventions: 1,
        status: "success" as const
      };

      if (apiAvailable) {
        await addHistoryEntry(historyEntry);
        const updatedHistory = await getHistory();
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
  }, [formValues, apiAvailable]);

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
        "/Facturation bandes d'enregistrements de 2026.xlsx"
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
      saveAs(zipBlob, `factures-asecna-${dateStr}.zip`);

      console.log("ZIP téléchargé avec succès");
      setStatus({ type: "done" });

      // Reset après 3 secondes
      setTimeout(() => {
        setStatus({ type: "ready" });
      }, 3000);
    } catch (err) {
      console.error("Erreur lors de la génération:", err);
      setStatus({
        type: "error",
        message:
          "Impossible de générer les factures. Vérifiez que le modèle Excel est présent.",
      });
    }
  }, [rows]);

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

      // Générer le fichier avec toutes les factures
      const buffer = await generateMultiInvoiceFile(
        conventions,
        "/Facturation bandes d'enregistrements de 2026.xlsx"
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
        date: new Date().toLocaleString('fr-FR'),
        type: "facture" as const,
        fileName: fileName,
        nbConventions: rows.length,
        status: "success" as const
      };

      if (apiAvailable) {
        await addHistoryEntry(historyEntry);
        // Recharger l'historique depuis l'API
        const updatedHistory = await getHistory();
        setHistory(updatedHistory);
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
        date: new Date().toLocaleString('fr-FR'),
        type: "facture" as const,
        fileName: file?.name || "Fichier inconnu",
        nbConventions: rows.length,
        status: "error" as const,
        details: err instanceof Error ? err.message : "Erreur inconnue"
      };

      if (apiAvailable) {
        await addHistoryEntry(historyEntry);
        const updatedHistory = await getHistory();
        setHistory(updatedHistory);
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
  }, [rows, file, apiAvailable]);

  const fields = useMemo(() => Object.keys(formValues), [formValues]);

  // Calculer l'étape actuelle du stepper
  const currentStep = useMemo(() => {
    if (status.type === "idle" || status.type === "error") return 1;
    if (status.type === "loading") return 2;
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
      const entryDate = new Date(h.date);
      const now = new Date();
      return entryDate.getMonth() === now.getMonth() &&
        entryDate.getFullYear() === now.getFullYear();
    });

    const lastEntry = history.length > 0 ? history[0] : null;

    return {
      thisMonthCount: thisMonth.length,
      totalConventions: thisMonth.reduce((acc, h) => acc + (h.nbConventions || 0), 0),
      lastImportDate: lastEntry?.date || null,
      lastFileName: lastEntry?.fileName || null
    };
  }, [history]);

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
        const updatedHistory = await getHistory();
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

  // Composant Sidebar
  const Sidebar = () => (
    <div className="sidebar">
      <div className="sidebar-logo">
        <div className="sidebar-logo-icon">
          <img src="/75664_O.jpg" alt="ASECNA Logo" />
        </div>
        <span className="sidebar-logo-text">ASECNA</span>
      </div>
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
          <button
            className="header-button"
            onClick={() => {
              // Reset
              setFicheMissionData(null);
              setFicheMissionStatus({ type: "idle" });
            }}
          >
            Nouvelle fiche
          </button>
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
                                setFicheMissionData(data);
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
                        <button className="dropzone-button-compact" type="button" onClick={(e) => e.stopPropagation()}>
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
                                setFicheMissionData(data);
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
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, dateDepart: e.target.value })}
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
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, dateRetour: e.target.value })}
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
                              onChange={(e) => setFicheMissionData({ ...ficheMissionData, duree: e.target.value })}
                            />
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
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center', justifyContent: 'center' }}>
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
                        setFicheMissionStatus({ type: "idle" });
                      }}
                    >
                      Annuler
                    </button>
                    <div
                      className="action-card primary compact"
                      style={{
                        minWidth: '200px',
                        maxWidth: '200px'
                      }}
                      onClick={async () => {
                        setFicheMissionStatus({ type: "generating" });
                        try {
                          const { generateFicheMission } = await import('./services/wordParser');
                          const blob = await generateFicheMission(ficheMissionData);
                          const { saveAs } = await import('file-saver');
                          saveAs(blob, `Fiche_Mission_${ficheMissionData.nom}.docx`);

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
          <button
            className="header-button"
            onClick={() => {
              setOrdreMissionData(null);
              setOrdreMissionStatus({ type: "idle" });
            }}
          >
            Nouvel ordre
          </button>
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
                                setOrdreMissionData(data);
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
                        <button className="dropzone-button-compact" type="button" onClick={(e) => e.stopPropagation()}>
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
                                setOrdreMissionData(data);
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
                </div>

                {/* Boutons d'action */}
                <div className="actions-section">
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center', justifyContent: 'center' }}>
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
                      Annuler
                    </button>
                    <div
                      className="action-card primary compact"
                      style={{
                        minWidth: '200px',
                        maxWidth: '200px'
                      }}
                      onClick={async () => {
                        setOrdreMissionStatus({ type: "generating" });
                        try {
                          const { generateOrdreMission } = await import('./services/wordParser');
                          const year = new Date().getFullYear();
                          const numero = `${String(Math.floor(Math.random() * 999) + 1).padStart(3, '0')}/${year}`;

                          const blob = await generateOrdreMission(ordreMissionData, numero);
                          const { saveAs } = await import('file-saver');
                          saveAs(blob, `Ordre_Mission_${ordreMissionData.nom}_${numero.replace('/', '-')}.docx`);

                          // Ajouter à l'historique
                          if (apiAvailable) {
                            await addHistoryEntry({
                              date: new Date().toLocaleString('fr-FR'),
                              type: 'facture',
                              fileName: `Ordre_Mission_${ordreMissionData.nom}.docx`,
                              nbConventions: 1,
                              status: 'success'
                            });
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
                              const updatedHistory = await getHistory();
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
                <select className="convention-select">
                  <option value="all">Tous les types</option>
                  <option value="facture">Facture</option>
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
                <div className="data-table-wrapper">
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Date</th>
                        <th>Type</th>
                        <th>Fichier</th>
                        <th>Conventions</th>
                        <th>Statut</th>
                      </tr>
                    </thead>
                    <tbody>
                      {history.map((item) => (
                        <tr key={item.id}>
                          <td>{item.date}</td>
                          <td>
                            <span style={{
                              background: "#EFF6FF",
                              color: "#1E40AF",
                              padding: "4px 10px",
                              borderRadius: "4px",
                              fontSize: "12px",
                              fontWeight: "500"
                            }}>
                              Facture
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
                      <td className="table-label">Dernier fichier importé</td>
                      <td>{file?.name || "Aucun"}</td>
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

  // Vue Dashboard
  if (activeTab === "dashboard") {
    return (
      <>
        <Sidebar />
        <header className="institutional-header">
          <div className="header-left">
            <span className="header-title">Dashboard</span>
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
          <button className="header-button" onClick={() => window.location.reload()}>
            Nouveau fichier
          </button>
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
                    <button className="dropzone-button-compact" type="button" onClick={(e) => e.stopPropagation()}>
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

              {/* Section fichiers récents */}
              {history.length > 0 && (
                <div className="recent-files-section">
                  <h3 className="recent-files-title">Fichiers récents</h3>
                  <div className="recent-files-list">
                    {history.slice(0, 5).map((item) => (
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
                              new Date(item.date).toLocaleDateString('fr-FR', {
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
                      <td>{String(formValues[key] ?? "—")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Boutons de génération */}
          <div className="actions-section">
            <div className="actions-grid">
              <div
                className="action-card primary"
                onClick={() => void handleGenerate()}
              >
                <div className="action-card-title">Générer cette facture</div>
              </div>

              {rows.length > 1 && (
                <div
                  className="action-card secondary"
                  onClick={() => void handleGenerateMulti()}
                >
                  <div className="action-card-title">Générer toutes les factures</div>
                </div>
              )}
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

