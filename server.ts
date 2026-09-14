import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Hỗ trợ parse JSON body
  app.use(express.json());

  // API Route: Kiểm tra trạng thái máy chủ
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // Tích hợp Vite middleware ở môi trường phát triển (Development)
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Phục vụ các tệp tĩnh ở môi trường production
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server is running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
