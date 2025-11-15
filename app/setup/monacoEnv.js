export function configureMonacoEnvironment() {
  self.MonacoEnvironment = {
    getWorkerUrl(moduleId, label) {
      const base = import.meta.env.BASE_URL || "/";
      if (label === "json") {
        return `${base}node_modules/monaco-editor/esm/vs/language/json/json.worker.js`;
      }
      if (label === "css" || label === "scss" || label === "less") {
        return `${base}node_modules/monaco-editor/esm/vs/language/css/css.worker.js`;
      }
      if (label === "html" || label === "handlebars" || label === "razor") {
        return `${base}node_modules/monaco-editor/esm/vs/language/html/html.worker.js`;
      }
      if (label === "typescript" || label === "javascript") {
        return `${base}node_modules/monaco-editor/esm/vs/language/typescript/ts.worker.js`;
      }
      return `${base}node_modules/monaco-editor/esm/vs/editor/editor.worker.js`;
    }
  };
}
