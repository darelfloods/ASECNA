import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

interface InvoiceBlock {
  startCol: number;
  endCol: number;
  label: string;
  sampleCells: Record<string, any>;
}

export const InvoiceLayoutAnalyzer: React.FC = () => {
  const [analysis, setAnalysis] = useState<string>("Chargement...");

  useEffect(() => {
    async function analyze() {
      try {
        const response = await fetch(
          "/Facturation bandes d'enregistrements de 2014-2025-22-10-25 (Enregistré automatiquement).xlsx"
        );
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        
        let output = `=== ANALYSE DE LA STRUCTURE DES FACTURES ===\n\n`;
        
        // Analyser la première feuille (FVC)
        const sheetName = "FVC";
        const sheet = workbook.Sheets[sheetName];

        if (!sheet || !sheet["!ref"]) {
          throw new Error(`Feuille "${sheetName}" introuvable ou référence !ref manquante`);
        }

        const range = XLSX.utils.decode_range(sheet["!ref"]);
        
        output += `Feuille analysée: ${sheetName}\n`;
        output += `Plage totale: ${sheet["!ref"]}\n\n`;
        
        // Détecter les blocs de facture (colonnes)
        // En regardant les données, on voit que les factures sont en colonnes:
        // - Facture 1: colonnes A-H (0-7)
        // - Facture 2: colonnes I-P (8-15)
        // - Facture 3: colonnes Q-X (16-23)
        
        output += `=== STRUCTURE DÉTECTÉE ===\n\n`;
        output += `Les factures semblent être organisées en colonnes :\n`;
        output += `- Bloc 1 (colonnes A-H): Facture N°117\n`;
        output += `- Bloc 2 (colonnes I-P): Facture N°036\n`;
        output += `- Bloc 3 (colonnes Q-X): Facture N°003\n\n`;
        
        // Analyser où se trouvent les informations clés
        output += `=== EMPLACEMENTS DES DONNÉES CLÉS ===\n\n`;
        
        // Bloc 1 (colonnes A-H)
        output += `BLOC 1 (colonnes A-H):\n`;
        output += `- Numéro facture: B3 = "${sheet["B3"]?.v}"\n`;
        output += `- Client: F6 = "${sheet["F6"]?.v}"\n`;
        output += `- Site: B10 = "${sheet["B10"]?.v}"\n`;
        output += `- Série: B12 = "${sheet["B12"]?.v}"\n`;
        output += `- Période: A15 = "${sheet["A15"]?.v}"\n`;
        output += `- Montant: H18 = "${sheet["H18"]?.v}"\n\n`;
        
        // Bloc 2 (colonnes I-P)
        output += `BLOC 2 (colonnes I-P):\n`;
        output += `- Numéro facture: J3 = "${sheet["J3"]?.v}"\n`;
        output += `- Site: K11 = "${sheet["K11"]?.v}"\n`;
        output += `- Période: I15 = "${sheet["I15"]?.v}"\n`;
        output += `- Désignation: J17 = "${sheet["J17"]?.v}"\n`;
        output += `- N° Convention: J18 = "${sheet["J18"]?.v}"\n\n`;
        
        // Analyser les lignes suivantes pour voir s'il y a d'autres factures
        output += `=== ANALYSE DES LIGNES SUIVANTES ===\n\n`;
        output += `Recherche de nouvelles factures dans les lignes 30-100...\n\n`;
        
        for (let row = 30; row <= Math.min(100, range.e.r); row++) {
          const cellB = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })]; // Colonne B
          const cellJ = sheet[XLSX.utils.encode_cell({ r: row, c: 9 })]; // Colonne J
          
          if (cellB?.v && String(cellB.v).includes("Facture")) {
            output += `Ligne ${row + 1}: Colonne B = "${cellB.v}"\n`;
          }
          if (cellJ?.v && String(cellJ.v).includes("Facture")) {
            output += `Ligne ${row + 1}: Colonne J = "${cellJ.v}"\n`;
          }
        }
        
        setAnalysis(output);
      } catch (err: any) {
        setAnalysis(`Erreur: ${err.message}`);
      }
    }
    analyze();
  }, []);

  return (
    <div style={{ padding: "20px", fontFamily: "monospace" }}>
      <h1>Analyse de la structure des factures</h1>
      <pre style={{ whiteSpace: "pre-wrap", fontSize: "12px" }}>
        {analysis}
      </pre>
    </div>
  );
};
