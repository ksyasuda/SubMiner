import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  base: './',
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          const normalized = id.replaceAll('\\', '/');

          if (
            normalized.includes('/node_modules/react-dom/') ||
            normalized.includes('/node_modules/react/')
          ) {
            return 'react-vendor';
          }

          if (normalized.includes('/node_modules/recharts/')) {
            return 'charts-vendor';
          }

          return undefined;
        },
      },
    },
  },
});
