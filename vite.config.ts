import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  define: {
    // Inject process.env.API_KEY into the client-side bundle
    // The value will be stringified, so process.env.API_KEY will become a literal string
    'process.env.API_KEY': JSON.stringify(process.env.API_KEY),
  },
  build: {
    rollupOptions: {
      // Mark these as external to avoid bundling them, as they are loaded via importmap
      external: ['react', 'react-dom', '@google/genai', 'pptxgenjs'],
      output: {
        // You might want to explore manualChunks here later to address the chunk size warning
      },
    },
  },
});