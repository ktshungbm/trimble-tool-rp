import { defineConfig } from 'vite'

// https://vitejs.dev/config/
export default defineConfig({
  base: './', // CRITICAL: This allows GitHub Pages to load assets correctly within any sub-directory
  build: {
    outDir: 'dist',
    sourcemap: true,
  }
})
