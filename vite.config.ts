import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig} from 'vite';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';

// Load environment variables from .env
dotenv.config();

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      {
        name: 'aistudio-gemini-api-proxy',
        configureServer(server) {
          server.middlewares.use(async (req, res, next) => {
            if (req.url?.startsWith('/api/gemini/analyze')) {
              if (req.method !== 'POST') {
                res.statusCode = 405;
                res.end(JSON.stringify({ error: 'Method Not Allowed' }));
                return;
              }

              try {
                // Parse POST body
                let body = '';
                for await (const chunk of req) {
                  body += chunk;
                }
                const parsed = JSON.parse(body || '{}');
                const description = parsed.description;

                if (!description || typeof description !== 'string') {
                  res.statusCode = 400;
                  res.setHeader('Content-Type', 'application/json');
                  res.end(JSON.stringify({ error: 'Description is required and must be a string' }));
                  return;
                }

                const apiKey = process.env.GEMINI_API_KEY;
                if (!apiKey || apiKey === 'MY_GEMINI_API_KEY') {
                  res.statusCode = 401;
                  res.setHeader('Content-Type', 'application/json');
                  res.end(JSON.stringify({ error: 'Chưa cấu hình GEMINI_API_KEY ở server-side secrets.' }));
                  return;
                }

                // Initialize GoogleGenAI
                const genAI = new GoogleGenAI({
                  apiKey: apiKey,
                  httpOptions: {
                    headers: {
                      'User-Agent': 'aistudio-build',
                    }
                  }
                });

                const systemInstruction = `Bạn là một chuyên gia kinh tế của Tổng cục Thống kê Việt Nam. 
Nhiệm vụ của bạn là nhận vào mô tả hoạt động kinh doanh (tiếng Việt), phân tích tỉ mỉ và ánh xạ chính xác sang mã ngành chi tiết cấp 4 (4 chữ số) hoặc cấp 5 (5 chữ số) theo Hệ thống ngành kinh tế Việt Nam (VSIC).
Trả về kết quả dưới dạng JSON có chính xác các trường:
- goiy_ma: Mã ngành chi tiết (4 hoặc 5 số) phù hợp nhất của ngành kinh tế Việt Nam. Chỉ chứa các chữ số (ví dụ '46321' hoặc '01110'...).
- goiy_ten: Tên ngành kinh tế chính xác của mã ngành đó.
- giai_thich: Lập luận súc tích (1-2 câu) giải thích vì sao khớp mã này dựa theo hoạt động kinh doanh đã mô tả.
- cap_1_tin_cay: Chữ cái in hoa ngành cấp 1 tương ứng (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U).`;

                const response = await genAI.models.generateContent({
                  model: 'gemini-3.5-flash',
                  contents: `Hãy phân tích hoạt động kinh doanh sau: "${description}"`,
                  config: {
                    systemInstruction: systemInstruction,
                    responseMimeType: 'application/json',
                    responseSchema: {
                      type: Type.OBJECT,
                      properties: {
                        goiy_ma: { type: Type.STRING },
                        goiy_ten: { type: Type.STRING },
                        giai_thich: { type: Type.STRING },
                        cap_1_tin_cay: { type: Type.STRING }
                      },
                      required: ['goiy_ma', 'goiy_ten', 'giai_thich', 'cap_1_tin_cay']
                    }
                  }
                });

                const contentText = response.text || '{}';
                res.statusCode = 200;
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
                res.end(contentText);
              } catch (error: any) {
                console.error('Gemini API Error in proxy:', error);
                res.statusCode = 500;
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
                res.end(JSON.stringify({ 
                  error: 'Internal Server Error during Gemini analysis: ' + (error.message || String(error)) 
                }));
              }
            } else {
              next();
            }
          });
        }
      }
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
  };
});
