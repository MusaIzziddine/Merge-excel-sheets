import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// base: './' makes the build work under GitHub Pages' /<repo-name>/ path.
export default defineConfig({
  plugins: [react()],
  base: './',
});
