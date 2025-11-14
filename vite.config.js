import { defineConfig } from 'vite';
import monacoEditorPlugin from 'vite-plugin-monaco-editor';

export default defineConfig({
  server: {
    port: 5500,
    open: true
  },
  optimizeDeps: {
    include: ['monaco-editor']
  },
  plugins: [
    monacoEditorPlugin.default({})
  ]
});

