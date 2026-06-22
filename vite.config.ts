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
      },
    },
    server: {
      hmr: process.env.DISABLE_HMR !== 'true',
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
    // === CHÈN CHÍNH XÁC ĐOẠN NÀY VÀO BÊN TRONG RETURN ===
    optimizeDeps: {
      exclude: ['pdfjs-dist'] // Chặn Vite quét thư viện này lúc chạy nội bộ
    },
    build: {
      rollupOptions: {
        external: ['pdfjs-dist'], // Khóa tính năng bóc tách tệp này để Vercel không báo lỗi khi build
      },
    },
    // ===================================================
  };
});
