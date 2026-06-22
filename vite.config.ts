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
    
    // === CHÈN CHÍNH XÁC VÀO ĐÂY (BÊN TRONG KHỐI RETURN) ===
    optimizeDeps: {
      exclude: ['pdfjs-dist'] // Bỏ qua quét thư viện này ở môi trường phát triển local
    },
    build: {
      rollupOptions: {
        external: ['pdfjs-dist'], // Lệnh ép buộc Rollup bỏ qua để không báo lỗi biên dịch trên Vercel
      },
    },
    // =====================================================
  };
});
