import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig} from 'vite';
import dotenv from 'dotenv';

// Load environment variables from .env
dotenv.config();

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss()
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
        // === THÊM DÒNG NÀY ĐỂ TRIỆT TIÊU LỖI ROLLUP ===
        // Ép Vite hiểu pdfjs-dist là một đối tượng trống, chặn đứng Rollup quét tìm file vật lý
        'pdfjs-dist': 'identity-obj-proxy', 
      },
    },
    server: {
      hmr: process.env.DISABLE_HMR !== 'true',
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
    optimizeDeps: {
      exclude: ['pdfjs-dist']
    },
    build: {
      rollupOptions: {
        external: ['pdfjs-dist'],
      },
    },
  };
});
