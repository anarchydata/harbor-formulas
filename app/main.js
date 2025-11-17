import "../styles.css";
// import { configureMonacoEnvironment } from "./setup/monacoEnv.js"; // Will be added back when Monaco is needed
import { initializeApp } from "./bootstrap.ts";

// configureMonacoEnvironment(); // Will be added back when Monaco is needed
document.addEventListener("DOMContentLoaded", () => {
  initializeApp().catch((error) => {
    console.error("Failed to initialize Harbor Formulas:", error);
  });
});
