import React, { useCallback, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { CELL_MAPPING } from "./mapping";

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

  const handleFile = useCallback(async (f: File) => {
    setStatus({ type: "parsing" });
    try {
      const data = await f.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheet];
      const json: Row[] = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
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
      const response = await fetch(
        "/Facturation bandes d'enregistrements de 2014-2025-22-10-25 (Enregistré automatiquement).xlsx"
      );
      if (!response.ok) {
        throw new Error("Modèle de facture introuvable");
      }
      const buf = await response.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buf);

      const sheet = workbook.worksheets[0];
      if (!sheet) {
        throw new Error("Aucune feuille dans le modèle de facture");
      }

      const payload = buildInvoicePayload(formValues);

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

      const out = await workbook.xlsx.writeBuffer();
      const blob = new Blob([out], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const baseName =
        typeof formValues["NOM DU CLIENT"] === "string" && formValues["NOM DU CLIENT"]
          ? String(formValues["NOM DU CLIENT"]).replace(/[\\/:*?"<>|]/g, "_")
          : "facture-asecna";

      saveAs(blob, `${baseName}.xlsx`);
      setStatus({ type: "done" });
      
      // Reset après 3 secondes
      setTimeout(() => {
        setStatus({ type: "ready" });
      }, 3000);
    } catch (err) {
      console.error(err);
      setStatus({
        type: "error",
        message:
          "Impossible de générer la facture. Vérifiez que le modèle Excel est présent.",
      });
    }
  }, [formValues]);

  const fields = useMemo(() => Object.keys(formValues), [formValues]);

  // Vue principale : dropzone
  if (status.type === "idle" || (status.type === "error" && !file)) {
    return (
      <div className="app-container">
        <header className="app-header">
          <h1 className="app-title">Génération de Facture ASECNA</h1>
          <p className="app-subtitle">Secrétariat administratif</p>
        </header>

        <main className="app-main">
          <div
            className={`dropzone-card ${isDragging ? "dropzone-card-dragging" : ""}`}
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
          >
            <svg className="dropzone-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="17 8 12 3 7 8" />
              <line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            <h2 className="dropzone-title">Déposez votre convention domaniale ici</h2>
            <p className="dropzone-text">ou cliquez pour sélectionner un fichier</p>
            <p className="dropzone-hint">Format accepté : .xlsx uniquement</p>
            <input
              type="file"
              accept=".xlsx"
              onChange={onFileChange}
              className="dropzone-input"
              id="file-input"
            />
            <label htmlFor="file-input" className="dropzone-button">
              Parcourir les fichiers
            </label>
          </div>

          {status.type === "error" && (
            <div className="alert alert-error">
              <svg className="alert-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="8" x2="12" y2="12" />
                <line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
              <span>{status.message}</span>
            </div>
          )}
        </main>

        <footer className="app-footer">
          <p>Application locale - Aucune donnée n'est envoyée sur Internet</p>
        </footer>
      </div>
    );
  }

  // Vue de chargement
  if (status.type === "parsing") {
    return (
      <div className="app-container">
        <header className="app-header">
          <h1 className="app-title">Génération de Facture ASECNA</h1>
        </header>
        <main className="app-main">
          <div className="loading-card">
            <div className="spinner"></div>
            <p className="loading-text">Lecture du fichier en cours...</p>
          </div>
        </main>
      </div>
    );
  }

  // Vue de génération
  if (status.type === "generating") {
    return (
      <div className="app-container">
        <header className="app-header">
          <h1 className="app-title">Génération de Facture ASECNA</h1>
        </header>
        <main className="app-main">
          <div className="loading-card">
            <div className="spinner"></div>
            <p className="loading-text">Génération de la facture Excel...</p>
          </div>
        </main>
      </div>
    );
  }

  // Vue principale avec données
  return (
    <div className="app-container">
      <header className="app-header">
        <h1 className="app-title">Génération de Facture ASECNA</h1>
        <button 
          className="header-reset-button"
          onClick={() => {
            setFile(null);
            setRows([]);
            setSelectedIndex(null);
            setFormValues({});
            setStatus({ type: "idle" });
          }}
        >
          ← Nouveau fichier
        </button>
      </header>

      <main className="app-main">
        {/* Carte récapitulative */}
        <div className="summary-card">
          <div className="summary-header">
            <div>
              <h2 className="summary-title">Données extraites</h2>
              <p className="summary-subtitle">
                {file?.name} • {rows.length} convention{rows.length > 1 ? "s" : ""} trouvée{rows.length > 1 ? "s" : ""}
              </p>
            </div>
            <div className="badge-success">
              <svg className="badge-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="20 6 9 17 4 12" />
              </svg>
              Fichier chargé
            </div>
          </div>

          {/* Sélection de convention si plusieurs */}
          {rows.length > 1 && (
            <div className="convention-selector">
              <label className="selector-label">Sélectionnez une convention :</label>
              <select
                className="selector-select"
                value={selectedIndex ?? 0}
                onChange={(e) => onRowClick(Number(e.target.value))}
              >
                {rows.map((row, index) => (
                  <option key={index} value={index}>
                    Convention {index + 1} - {String(row["NOM DU CLIENT"] ?? row["Nº CONVENTION"] ?? `Ligne ${index + 1}`)}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Grille de données */}
          <div className="data-grid">
            {fields.map((key) => (
              <div key={key} className="data-field">
                <label className="data-label">{key}</label>
                <input
                  className="data-input"
                  value={String(formValues[key] ?? "")}
                  onChange={(e) => handleFieldChange(key, e.target.value)}
                />
              </div>
            ))}
          </div>
        </div>

        {/* Bouton de génération */}
        <button
          className="generate-button"
          onClick={() => void handleGenerate()}
        >
          <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
            <polyline points="7 10 12 15 17 10" />
            <line x1="12" y1="15" x2="12" y2="3" />
          </svg>
          Générer la facture Excel
        </button>

        {/* Message de succès */}
        {status.type === "done" && (
          <div className="alert alert-success">
            <svg className="alert-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="20 6 9 17 4 12" />
            </svg>
            <span>Facture générée et téléchargée avec succès !</span>
          </div>
        )}

        {/* Message d'erreur */}
        {status.type === "error" && (
          <div className="alert alert-error">
            <svg className="alert-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
            <span>{status.message}</span>
          </div>
        )}
      </main>

      <footer className="app-footer">
        <p>Application locale - Aucune donnée n'est envoyée sur Internet</p>
      </footer>
    </div>
  );
};

export default App;

