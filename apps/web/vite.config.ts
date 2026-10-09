import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { createHash } from 'crypto';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const require = createRequire(import.meta.url);
const spreadsheetCodecSource = require('fs')
  .readFileSync(require.resolve('xlsx-js-style/dist/xlsx.min.js'), 'utf8')
  .replace(/\n\/\/#[#]? sourceMappingURL=.*$/u, '');
const spreadsheetCodecHash = createHash('sha256')
  .update(spreadsheetCodecSource)
  .digest('hex')
  .slice(0, 12);
const spreadsheetCodecFileName = `assets/xlsx-codec-${spreadsheetCodecHash}.js`;

/**
 * xlsx-js-style is distributed as one browser bundle. Emit it as a lazy ESM asset instead of
 * asking Rollup to inline that codec into both the app fallback and the Excel worker.
 */
function emitSpreadsheetCodec() {
  return {
    name: 'emit-spreadsheet-codec',
    generateBundle() {
      this.emitFile({
        type: 'asset',
        fileName: spreadsheetCodecFileName,
        source: `${spreadsheetCodecSource}\nexport { XLSX as default };\n`,
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), emitSpreadsheetCodec()],
  define: {
    __XLSX_CODEC_ASSET__: JSON.stringify(`./${spreadsheetCodecFileName.replace(/^assets\//u, '')}`),
  },
  server: {
    port: 5173,
    host: true,
  },
  resolve: {
    alias: {
      '@excel-agent/engine': path.resolve(__dirname, '../../packages/engine/src/index.ts'),
      '@excel-agent/agent': path.resolve(__dirname, '../../packages/agent/src/index.ts'),
    },
  },
  build: {
    chunkSizeWarningLimit: 700,
    rollupOptions: {
      output: {
        manualChunks(id: string) {
          if (id.includes('node_modules/react')) return 'react';
          if (id.includes('node_modules/framer-motion')) return 'motion';
          if (id.includes('node_modules/lucide-react')) return 'icons';
          if (id.includes('node_modules/@supabase')) return 'supabase';
          if (id.includes('node_modules/@tanstack')) return 'tanstack';
          if (id.includes('/packages/engine/src/')) return 'engine';
          if (id.includes('/packages/agent/src/')) return 'agent';
          return undefined;
        },
      },
    },
  },
});
