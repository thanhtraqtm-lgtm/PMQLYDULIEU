import express from "express";
import path from "path";
import fs from "fs";
import dotenv from "dotenv";

// Nạp cấu hình từ .env nếu có (không gây lỗi nếu không có file .env)
dotenv.config();

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  // Hỗ trợ parse JSON body
  app.use(express.json());

  // API Route: Kiểm tra trạng thái máy chủ (Cloud Run & Internal probes)
  app.get("/health", (req, res) => {
    res.json({ status: "ok" });
  });
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // API Route: Trợ lý AI tạo biểu thức điều kiện tra cứu/kiểm tra logic (GSO syntax)
  app.post("/api/ai/generate-expression", async (req, res) => {
    try {
      const { prompt, columns, customApiKey } = req.body;
      const apiKey = customApiKey || process.env.VITE_GEMINI_API_KEY || "AQ.Ab8RN6KLde-z8EMqDKTLwQqUG4GHLtbKg6-Isumal41h6SHiMg";

      if (!prompt || typeof prompt !== "string") {
        return res.status(400).json({ error: "Thiếu mô tả yêu cầu tra cứu!" });
      }

      const columnList = Array.isArray(columns) ? columns : [];
      const systemPrompt = `Bạn là chuyên gia tra cứu dữ liệu thống kê GSO. Nhiệm vụ của bạn là đọc yêu cầu tra cứu hoặc kiểm tra logic bằng tiếng Việt và chuyển thành biểu thức điều kiện chính xác.

Danh sách các cột thực tế trong bảng dữ liệu:
[${columnList.map((c: string) => `"${c}"`).join(", ")}]

Quy tắc cú pháp biểu thức bắt buộc:
1. Tên cột luôn nằm trong dấu ngoặc vuông: [TênCột]. Phải khớp chính xác tên cột trong danh sách trên.
2. Các phép so sánh: = (bằng), <> (khác), <, <=, >, >=
3. Các phép toán tử logic:
   - và (AND)
   - hoặc (OR)
4. Phép toán chuỗi và rỗng:
   - có chứa (ví dụ: [TenDoanhNghiep] có chứa 'Hưng Yên')
   - không có giá trị (ví dụ: [MaSoThue] không có giá trị)
   - có giá trị (ví dụ: [MaSoThue] có giá trị)
5. Toán tử số học: +, -, *, /, (, )

Yêu cầu người dùng:
"${prompt.trim()}"

Trả về duy nhất 1 đối tượng JSON (KHÔNG bọc markdown, KHÔNG viết thêm lời giải thích ngoài JSON):
{
  "expression": "biểu thức chuẩn theo cú pháp trên",
  "explanation": "giải thích ngắn gọn súc tích bằng tiếng Việt"
}`;

      const candidateModels = ["gemini-2.5-flash", "gemini-3-flash-preview", "gemini-2.5-pro"];
      let candidateText = "";
      let lastError = "";

      for (const model of candidateModels) {
        try {
          const googleRes = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                contents: [{ parts: [{ text: systemPrompt }] }],
                generationConfig: { temperature: 0.1 }
              })
            }
          );

          const data: any = await googleRes.json();
          if (googleRes.ok && data?.candidates?.[0]?.content?.parts?.[0]?.text) {
            candidateText = data.candidates[0].content.parts[0].text;
            break;
          } else {
            lastError = data?.error?.message || `Lỗi từ ${model}`;
          }
        } catch (e: any) {
          lastError = e?.message || `Lỗi kết nối ${model}`;
        }
      }

      if (!candidateText) {
        return res.status(500).json({ error: lastError || "Không thể nhận phản hồi từ mô hình AI" });
      }

      const cleanedJson = candidateText.replace(/```json/gi, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(cleanedJson);

      return res.json({ success: true, expression: parsed.expression, explanation: parsed.explanation });
    } catch (err: any) {
      console.error("[AI Expression API Error]:", err);
      return res.status(500).json({ error: err?.message || "Lỗi xử lý yêu cầu AI" });
    }
  });

  // API Route: Trợ lý AI tạo quy tắc Logic tự động (Server-side bypass CORS & Browser 403)
  app.post("/api/ai/generate-rule", async (req, res) => {
    try {
      const { prompt, columns, customApiKey } = req.body;
      const apiKey = customApiKey || process.env.VITE_GEMINI_API_KEY || "AQ.Ab8RN6KLde-z8EMqDKTLwQqUG4GHLtbKg6-Isumal41h6SHiMg";

      if (!prompt || typeof prompt !== "string") {
        return res.status(400).json({ error: "Thiếu mô tả yêu cầu kiểm tra logic!" });
      }

      const columnList = Array.isArray(columns) ? columns : [];
      const systemPrompt = `Bạn là chuyên gia phân tích dữ liệu thống kê. Nhiệm vụ của bạn là đọc yêu cầu kiểm tra logic bằng tiếng Việt và chuyển thành cấu hình quy tắc logic JSON.

Danh sách các cột thực tế trong bảng dữ liệu:
[${columnList.map((c: string) => `"${c}"`).join(", ")}]

Yêu cầu kiểm tra của người dùng:
"${prompt.trim()}"

Hãy chọn toán tử phù hợp nhất trong danh sách sau:
1. "not_empty": Không được để trống.
2. "is_number": Phải là dạng số.
3. "greater_than": Lớn hơn giá trị số trong "param".
4. "less_than": Nhỏ hơn giá trị số trong "param".
5. "equals": Bằng giá trị chuỗi trong "param".
6. "not_equals": Khác giá trị chuỗi trong "param".
7. "len_equals": Độ dài chuỗi ký tự bằng số trong "param".
8. "compare_col_gt": Cột colA phải lớn hơn cột colB (ví dụ: Tổng số > Chi tiết).
9. "compare_col_lt": Cột colA phải nhỏ hơn cột colB.
10. "compare_col_eq": Cột colA phải bằng cột colB.
11. "col_a_sum_col_b_c": Cột colA = colB + colC (cân đối tổng).
12. "invalid_tax_code": Mã số thuế sai 10 hoặc 13 số.
13. "invalid_vsic": Mã ngành không thuộc danh mục VSIC.
14. "conflict_dt_ld": Mâu thuẫn giữa Doanh thu và Lao động.

Trả về duy nhất 1 đối tượng JSON, KHÔNG có giải thích, KHÔNG bọc mã markdown:
{
  "name": "Tên quy tắc ngắn gọn, dễ hiểu",
  "colA": "Tên cột chính khớp trong danh sách cột",
  "operator": "toán tử ở trên",
  "colB": "Tên cột phụ thứ 2 nếu dùng so sánh giữa 2 cột (hoặc chuỗi rỗng)",
  "colC": "Tên cột phụ thứ 3 nếu dùng phép cộng colA = colB + colC (hoặc chuỗi rỗng)",
  "param": "Tham số số hoặc chuỗi nếu có (hoặc chuỗi rỗng)",
  "severity": "error" hoặc "warning",
  "description": "Giải thích ngắn điều kiện kiểm tra"
}`;

      // Danh sách các mô hình dự phòng nếu 2.5-flash bị quá tải hoặc đạt hạn mức quota
      const candidateModels = ["gemini-2.5-flash", "gemini-3-flash-preview", "gemini-2.5-pro"];
      let candidateText = "";
      let lastError = "";

      for (const model of candidateModels) {
        try {
          const googleRes = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                contents: [{ parts: [{ text: systemPrompt }] }],
                generationConfig: { temperature: 0.1 }
              })
            }
          );

          const data: any = await googleRes.json();
          if (googleRes.ok && data?.candidates?.[0]?.content?.parts?.[0]?.text) {
            candidateText = data.candidates[0].content.parts[0].text;
            break;
          } else {
            lastError = data?.error?.message || `Lỗi từ ${model}`;
          }
        } catch (e: any) {
          lastError = e?.message || `Lỗi kết nối ${model}`;
        }
      }

      if (!candidateText) {
        return res.status(500).json({ error: lastError || "Không thể nhận phản hồi từ mô hình AI" });
      }

      const cleanedJson = candidateText.replace(/```json/gi, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(cleanedJson);

      return res.json({ success: true, rule: parsed });
    } catch (err: any) {
      console.error("[AI API Error]:", err);
      return res.status(500).json({ error: err?.message || "Lỗi xử lý yêu cầu AI máy chủ" });
    }
  });

  // API Route: Proxy chung cho các yêu cầu Gemini khác nếu cần
  app.post("/api/ai/proxy", async (req, res) => {
    try {
      const { model = "gemini-2.5-flash", contents, generationConfig, customApiKey } = req.body;
      const apiKey = customApiKey || process.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || "AQ.Ab8RN6KLde-z8EMqDKTLwQqUG4GHLtbKg6-Isumal41h6SHiMg";

      const googleRes = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contents, generationConfig })
        }
      );

      const data: any = await googleRes.json();
      if (!googleRes.ok) {
        return res.status(googleRes.status).json({ error: data?.error?.message || "Lỗi dịch vụ Gemini" });
      }

      return res.json(data);
    } catch (err: any) {
      return res.status(500).json({ error: err?.message || "Lỗi máy chủ proxy AI" });
    }
  });

  // API Route: Hỗ trợ phân tích gợi ý mã ngành từ mô tả
  app.post("/api/gemini/analyze", async (req, res) => {
    try {
      const { description, customApiKey } = req.body;
      const apiKey = customApiKey || process.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || "AQ.Ab8RN6KLde-z8EMqDKTLwQqUG4GHLtbKg6-Isumal41h6SHiMg";

      if (!description || typeof description !== "string") {
        return res.status(400).json({ error: "Thiếu mô tả hoạt động sản xuất kinh doanh!" });
      }

      const prompt = `Bạn là chuyên gia phân loại ngành kinh tế Việt Nam (VSIC).
Hãy xác định mã ngành VSIC 4 hoặc 5 chữ số phù hợp nhất cho mô tả hoạt động: "${description}".
Trả về duy nhất 1 JSON hợp lệ:
{
  "goiy_ma": "mã ngành",
  "goiy_ten": "tên ngành",
  "cap_1_tin_cay": "ngành cấp 1",
  "giai_thich": "giải thích ngắn gọn"
}`;

      const googleRes = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.1 }
          })
        }
      );

      const data: any = await googleRes.json();
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
      const cleaned = text.replace(/```json/gi, "").replace(/```/g, "").trim();
      if (cleaned) {
        return res.json(JSON.parse(cleaned));
      }
      return res.status(500).json({ error: "Không phân tích được phản hồi từ AI" });
    } catch (err: any) {
      return res.status(500).json({ error: err?.message || "Lỗi xử lý gợi ý mã ngành" });
    }
  });

  // Xác định chế độ môi trường:
  // - Nếu là bundle CJS (dist/server.cjs) hoặc NODE_ENV=production -> Chế độ production phục vụ file tĩnh từ dist/
  // - Mặc định khi chạy `npm run dev` (tsx server.ts) -> Chế độ development tích hợp Vite dev middleware
  const isCjsBundle = typeof __filename !== "undefined" && (__filename.endsWith(".cjs") || __filename.includes("dist"));
  const isProduction = process.env.NODE_ENV === "production" || isCjsBundle;

  if (!isProduction) {
    try {
      const { createServer: createViteServer } = await import("vite");
      const vite = await createViteServer({
        server: { middlewareMode: true },
        appType: "spa",
      });
      app.use(vite.middlewares);
      console.log("[Vite] Dev middleware mounted successfully.");
    } catch (viteErr) {
      console.warn("[Vite Middleware Startup Warning]:", viteErr);
    }
  } else {
    // Phục vụ các tệp tĩnh ở môi trường production (Cloud Run)
    const distPath =
      typeof __dirname !== "undefined" && fs.existsSync(path.join(__dirname, "index.html"))
        ? __dirname
        : path.resolve(process.cwd(), "dist");

    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      const indexPath = path.join(distPath, "index.html");
      if (fs.existsSync(indexPath)) {
        res.sendFile(indexPath);
      } else {
        const rootIndexPath = path.resolve(process.cwd(), "dist", "index.html");
        if (fs.existsSync(rootIndexPath)) {
          res.sendFile(rootIndexPath);
        } else {
          res.status(200).send("<!doctype html><html><body><div id='root'>Hệ thống đang khởi động...</div></body></html>");
        }
      }
    });
  }

  const server = app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server is running on http://0.0.0.0:${PORT} [mode: ${!isProduction ? "development" : "production"}]`);
  });

  server.on("error", (err: any) => {
    if (err.code === "EADDRINUSE") {
      console.error(`[Server Error]: Cổng ${PORT} đang được sử dụng bởi tiến trình khác.`);
    } else {
      console.error("[Server Error]:", err);
    }
  });
}

startServer();
