import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import agoraTokenPkg from "agora-token";

const { RtcTokenBuilder, RtcRole } = agoraTokenPkg;

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Hỗ trợ parse JSON body
  app.use(express.json());

  // API Route: Tạo Agora Token động từ App ID và App Certificate
  app.post("/api/generate-token", (req: express.Request, res: express.Response) => {
    try {
      let { appId, appCertificate, channelName, uid } = req.body;

      // Ưu tiên dùng Biến môi trường hệ thống nếu Client không truyền lên, hoặc truyền giá trị tượng trưng
      if (!appId || appId === "YOUR_AGORA_APP_ID" || appId === "") {
        appId = process.env.VITE_AGORA_APP_ID || process.env.AGORA_APP_ID || "";
      }
      if (!appCertificate || appCertificate === "system-configured" || appCertificate === "") {
        appCertificate = process.env.AGORA_APP_CERTIFICATE || process.env.VITE_AGORA_CERTIFICATE || "";
      }
      
      if (!appId || !appCertificate || !channelName) {
        return res.status(400).json({ error: "Thiếu tham số bắt buộc: appId, appCertificate, channelName" });
      }

      const role = RtcRole.PUBLISHER;
      // Token có hiệu lực trong 2 giờ
      const expirationTimeInSeconds = 7200;
      const currentTimestamp = Math.floor(Date.now() / 1000);
      const privilegeExpiredTs = currentTimestamp + expirationTimeInSeconds;

      let token = "";
      
      // Kiểm tra xem UID là dạng số hay chuỗi kí tự
      const isNumeric = !isNaN(Number(uid)) && uid !== "" && uid !== null && uid !== undefined;
      const numericUid = isNumeric ? Number(uid) : 0;

      if (isNumeric) {
        token = RtcTokenBuilder.buildTokenWithUid(
          appId,
          appCertificate,
          channelName,
          numericUid,
          role,
          privilegeExpiredTs,
          privilegeExpiredTs
        );
      } else {
        const stringUid = uid ? String(uid) : "0";
        token = RtcTokenBuilder.buildTokenWithUserAccount(
          appId,
          appCertificate,
          channelName,
          stringUid,
          role,
          privilegeExpiredTs,
          privilegeExpiredTs
        );
      }

      return res.json({ token });
    } catch (err: any) {
      console.error("Lỗi khi sinh Agora Token:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ nội bộ" });
    }
  });

  // API Route: Lấy cấu hình hệ thống mặc định (Cấu hình toàn cục từ file .env)
  app.get("/api/agora-global-status", (req: express.Request, res: express.Response) => {
    const appId = process.env.VITE_AGORA_APP_ID || process.env.AGORA_APP_ID || "";
    const hasCertificate = !!(process.env.AGORA_APP_CERTIFICATE || process.env.VITE_AGORA_CERTIFICATE);
    
    return res.json({
      hasGlobalConfig: !!(appId && hasCertificate),
      appId: appId || null
    });
  });

  // Lưu trữ cấu hình phòng họp Agora trong bộ nhớ RAM máy chủ để đồng bộ tức thì cho tất cả các thiết bị kết nối chung phòng
  const agoraRoomConfigs = new Map<string, { appId: string; appCertificate: string; token: string }>();
  // Lưu trữ mật khẩu phòng họp động trực tuyến trên máy chủ đám mây
  const agoraRoomPasswords = new Map<string, string>();

  // API Route: Lấy mật khẩu của phòng họp
  app.get("/api/get-room-password", (req: express.Request, res: express.Response) => {
    try {
      const channelName = req.query.channelName as string;
      if (!channelName) {
        return res.status(400).json({ error: "Thiếu tên phòng channelName" });
      }
      const password = agoraRoomPasswords.get(channelName);
      if (password) {
        return res.json({ found: true, password });
      } else {
        return res.json({ found: false });
      }
    } catch (err: any) {
      console.error("Lỗi khi lấy mật khẩu phòng:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ" });
    }
  });

  // API Route: Thiết lập mật khẩu cho phòng họp
  app.post("/api/set-room-password", (req: express.Request, res: express.Response) => {
    try {
      const { channelName, password } = req.body;
      if (!channelName || !password) {
        return res.status(400).json({ error: "Thiếu channelName hoặc password" });
      }
      agoraRoomPasswords.set(channelName, password);
      console.log(`[Agora Password] Đã thiết lập mật khẩu mới cho phòng [${channelName}]: ${password}`);
      return res.json({ success: true, message: "Thiết lập mật khẩu thành công!" });
    } catch (err: any) {
      console.error("Lỗi khi đặt mật khẩu phòng:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ" });
    }
  });

  // API Route: Xóa mật khẩu phòng họp (đặt lại phòng về trạng thái tự do)
  app.post("/api/clear-room-password", (req: express.Request, res: express.Response) => {
    try {
      const { channelName } = req.body;
      if (!channelName) {
        return res.status(400).json({ error: "Thiếu channelName" });
      }
      agoraRoomPasswords.delete(channelName);
      console.log(`[Agora Password] Đã xóa mật khẩu bảo mật của phòng [${channelName}]`);
      return res.json({ success: true, message: "Đã xóa mật khẩu phòng họp!" });
    } catch (err: any) {
      console.error("Lỗi khi xóa mật khẩu phòng:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ" });
    }
  });

  // API Route: Lưu cấu hình phòng họp của một đơn vị/kênh họp
  app.post("/api/save-agora-config", (req: express.Request, res: express.Response) => {
    try {
      const { channelName, appId, appCertificate, token } = req.body;
      if (!channelName) {
        return res.status(400).json({ error: "Thiếu tên phòng channelName" });
      }
      
      agoraRoomConfigs.set(channelName, {
        appId: appId || "",
        appCertificate: appCertificate || "",
        token: token || ""
      });

      console.log(`[Agora Sync] Đã lưu cấu hình mới cho phòng: ${channelName}`);
      return res.json({ success: true, message: "Đã đồng bộ cấu hình phòng thành công lên máy chủ đám mây!" });
    } catch (err: any) {
      console.error("Lỗi khi lưu cấu hình Agora:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ" });
    }
  });

  // API Route: Lấy cấu hình phòng họp của một đơn vị/kênh họp
  app.get("/api/get-agora-config", (req: express.Request, res: express.Response) => {
    try {
      const channelName = req.query.channelName as string;
      if (!channelName) {
        return res.status(400).json({ error: "Thiếu tham số channelName" });
      }

      const config = agoraRoomConfigs.get(channelName);
      if (config) {
        return res.json({ found: true, ...config });
      } else {
        return res.json({ found: false });
      }
    } catch (err: any) {
      console.error("Lỗi khi lấy cấu hình Agora:", err);
      return res.status(500).json({ error: err.message || "Lỗi máy chủ" });
    }
  });

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
