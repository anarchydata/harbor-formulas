import "../styles.css";
import { configureMonacoEnvironment } from "./setup/monacoEnv.js";
import { initializeApp } from "./bootstrap.js";

configureMonacoEnvironment();
document.addEventListener("DOMContentLoaded", () => {
  initializeApp().catch((error) => {
    console.error("Failed to initialize Harbor Formulas:", error);
  });
});
