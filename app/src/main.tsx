import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { AnalyzeFile } from "./AnalyzeFile";
import "./styles.css";

// Temporaire: afficher l'analyse
const showAnalysis = window.location.search.includes("analyze");

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    {showAnalysis ? <AnalyzeFile /> : <App />}
  </React.StrictMode>
);
