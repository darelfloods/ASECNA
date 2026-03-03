import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

export const AnalyzeFile: React.FC = () => {
  const [analysis, setAnalysis] = useState<string>("Chargement...");

  useEffect(() => {
    async function analyze() {
      try {
        const response = await fetch(
          "/CONVENTIONS DOMANIALES ACTUALISEES 2025 (Enregistré automatiquement).xlsx"
        );
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheet = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheet];

        let output = `=== Feuille: ${firstSheet} ===\n`;
        output += `Plage: ${sheet["!ref"]}\n\n`;

        // Lire les lignes brutes
        output += "=== Premières lignes brutes (cellules) ===\n";
        const range = XLSX.utils.decode_range(sheet["!ref"]);
        for (let row = 0; row <= Math.min(14, range.e.r); row++) {
          const cells = [];
          for (let col = 0; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = sheet[cellAddress];
            cells.push(cell ? cell.v : null);
          }
          output += `Ligne ${row + 1}: ${JSON.stringify(cells)}\n`;
        }

        // Lire comme JSON
        output += "\n=== Données JSON (avec defval='') ===\n";
        const json: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        output += `Nombre de lignes: ${json.length}\n\n`;

        if (json.length > 0) {
          output += "Clés trouvées:\n";
          output += JSON.stringify(Object.keys(json[0]), null, 2) + "\n\n";

          output += "Premières 3 lignes:\n";
          output += JSON.stringify(json.slice(0, 3), null, 2) + "\n";
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
      <h1>Analyse du fichier Excel</h1>
      <pre style={{ whiteSpace: "pre-wrap", fontSize: "12px" }}>
        {analysis}
      </pre>
    </div>
  );
};
