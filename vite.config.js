import { defineConfig } from 'vite';
import monacoEditorPlugin from 'vite-plugin-monaco-editor';

const GOOGLE_TAG_ID = 'G-PZ9LE2EGD4';

const googleAnalyticsFirstPlugin = {
  name: 'inject-google-analytics-first',
  enforce: 'pre',
  transformIndexHtml() {
    const inlineConfig = `
      window.dataLayer = window.dataLayer || [];
      function gtag(){dataLayer.push(arguments);}
      gtag('js', new Date());
      gtag('config', '${GOOGLE_TAG_ID}');
    `.trim();

    return [
      {
        tag: 'script',
        attrs: {
          async: true,
          src: `https://www.googletagmanager.com/gtag/js?id=${GOOGLE_TAG_ID}`
        },
        injectTo: 'head-prepend'
      },
      {
        tag: 'script',
        children: inlineConfig,
        injectTo: 'head-prepend'
      }
    ];
  }
};

export default defineConfig({
  server: {
    port: 5500,
    open: true
  },
  optimizeDeps: {
    include: ['monaco-editor']
  },
  plugins: [
    googleAnalyticsFirstPlugin,
    monacoEditorPlugin.default({})
  ]
});

