import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'path';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    https: true,
  },
  build: {
    outDir: 'docs', // 将打包输出目录改为 docs，方便 GitHub Pages 识别
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'taskpane.html'),
      },
    },
  },
});
