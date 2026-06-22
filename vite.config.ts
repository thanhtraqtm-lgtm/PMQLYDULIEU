import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import { defineConfig } from 'vite';
import dotenv from 'dotenv';

// Tải các biến môi trường từ tệp .env
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
    // === ĐOẠN 4: CẤU HÌNH BẮT BUỘC ĐỂ VITE 6 BIÊN DỊCH ĐƯỢC PDFJS-DIST ===
    optimizeDeps: {
      include: ['pdfjs-dist'], 
      esbuildOptions: {
        supported: {
          'top-level-await': true // Cho phép chạy tính năng await đặc thù của file worker pdfjs
        }
      }
    },
    build: {
      target: 'es2022' // Nâng cấp chuẩn đầu ra biên dịch để tương thích với cấu trúc của pdfjs-dist
    }
    // ===================================================================
  };
});
