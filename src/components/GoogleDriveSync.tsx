import React, { useState, useEffect } from "react";
import { useAuth } from "../context/AuthContext";
import { 
  Cloud, 
  CloudLightning, 
  CloudOff, 
  Download, 
  Upload, 
  FileSpreadsheet, 
  RefreshCw, 
  Trash2, 
  Check, 
  AlertCircle, 
  Loader2,
  Calendar,
  FileJson,
  Database
} from "lucide-react";
import * as XLSX from "xlsx";

interface GoogleDriveSyncProps {
  mainData: any[];
  rawImportedData: any[];
  columns: any[];
  fileName: string;
  mapping: any;
  customColConfigs: any[];
  dataMode: "corp" | "individual";
  onRestore: (restoredState: {
    mainData: any[];
    rawImportedData: any[];
    columns: any[];
    fileName: string;
    mapping: any;
    customColConfigs: any[];
  }) => void;
}

interface BackupFile {
  id: string;
  name: string;
  createdTime: string;
  size?: string;
}

export const GoogleDriveSync: React.FC<GoogleDriveSyncProps> = ({
  mainData,
  rawImportedData,
  columns,
  fileName,
  mapping,
  customColConfigs,
  dataMode,
  onRestore
}) => {
  const { signInWithGoogle, googleAccessToken, logout, user } = useAuth();
  
  const [backups, setBackups] = useState<BackupFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [syncingExcel, setSyncingExcel] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [showConfirmRestoreId, setShowConfirmRestoreId] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // Tự động tải danh sách tệp lưu trữ khi đã có Access Token
  useEffect(() => {
    if (googleAccessToken) {
      fetchBackups();
    } else {
      setBackups([]);
    }
  }, [googleAccessToken, dataMode]);

  const fetchBackups = async () => {
    setLoading(true);
    setErrorMessage(null);
    try {
      // Tìm các tệp sao lưu JSON có tên chứa 'VSIC_Backup'
      const query = `name contains 'VSIC_Backup' and trashed = false`;
      const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id, name, createdTime, size)&orderBy=createdTime desc`;
      
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${googleAccessToken}`
        }
      });

      if (!response.ok) {
        throw new Error("Không thể kết nối danh sách tệp Google Drive. Token có thể đã hết hạn.");
      }

      const data = await response.json();
      // Lọc các bản sao lưu phù hợp với chế độ dữ liệu hiện tại để tránh nhầm lẫn
      const filterKeyword = dataMode === "corp" ? "DoanhNghiep" : "CaThe";
      const filteredFiles = (data.files || []).filter((f: any) => f.name.includes(filterKeyword));
      
      setBackups(filteredFiles);
    } catch (error: any) {
      console.error("Lỗi lấy danh sách sao lưu:", error);
      setErrorMessage(error.message || "Lỗi tải tệp sao lưu từ Google Drive.");
    } finally {
      setLoading(false);
    }
  };

  const handleLogin = async () => {
    setIsLoggingIn(true);
    setErrorMessage(null);
    try {
      await signInWithGoogle();
      setStatusMessage("Kết nối Google Drive thành công!");
      setTimeout(() => setStatusMessage(null), 3000);
    } catch (err: any) {
      console.error("Đăng nhập thất bại:", err);
      setErrorMessage("Kết nối tài khoản Google thất bại. Hãy thử lại.");
    } finally {
      setIsLoggingIn(false);
    }
  };

  // Tạo tệp sao lưu cấu hình toàn bộ phiên (JSON Backup)
  const handleBackup = async () => {
    if (!mainData || mainData.length === 0) {
      alert("Không có dữ liệu trong trạm làm việc để sao lưu!");
      return;
    }

    const confirmed = window.confirm(
      `Xác nhận tải toàn bộ cấu hình & ${mainData.length} dòng dữ liệu chế độ ${dataMode === "corp" ? "Doanh Nghiệp" : "Cá Thể"} lên Google Drive của bạn?`
    );
    if (!confirmed) return;

    setSyncing(true);
    setErrorMessage(null);
    setStatusMessage("Đang chuẩn bị dữ liệu sao lưu...");

    try {
      const backupState = {
        mainData,
        rawImportedData,
        columns,
        fileName,
        mapping,
        customColConfigs,
        exportedAt: new Date().toISOString(),
        dataMode
      };

      const modeStr = dataMode === "corp" ? "DoanhNghiep" : "CaThe";
      const now = new Date();
      const dateStr = now.toLocaleDateString("vi-VN").replace(/\//g, "-");
      const timeStr = now.toLocaleTimeString("vi-VN", { hour12: false }).replace(/:/g, "-");
      const driveFileName = `VSIC_Backup_${modeStr}_${dateStr}_${timeStr}.json`;

      setStatusMessage("Đang tạo tệp trên Google Drive...");
      
      // 1. Tạo file metadata
      const metaResponse = await fetch("https://www.googleapis.com/drive/v3/files", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${googleAccessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          name: driveFileName,
          mimeType: "application/json",
          description: `VSIC Data Manager Backup State. Mode: ${dataMode}. Records: ${mainData.length}`
        })
      });

      if (!metaResponse.ok) {
        throw new Error("Lỗi khởi tạo tệp mới trên Google Drive.");
      }

      const meta = await metaResponse.json();
      const fileId = meta.id;

      setStatusMessage("Đang truyền tải dữ liệu...");

      // 2. Upload nội dung tệp
      const uploadResponse = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${googleAccessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(backupState)
      });

      if (!uploadResponse.ok) {
        throw new Error("Lỗi truyền dữ liệu tệp sao lưu.");
      }

      setStatusMessage("Đã sao lưu thành công tệp lên Google Drive!");
      await fetchBackups(); // Tải lại danh sách
      setTimeout(() => setStatusMessage(null), 4000);
    } catch (error: any) {
      console.error("Lỗi sao lưu:", error);
      setErrorMessage(error.message || "Lỗi lưu trữ dữ liệu lên Google Drive.");
    } finally {
      setSyncing(false);
    }
  };

  // Lưu file Excel (.xlsx) trực tiếp lên Google Drive
  const handleExportExcelToDrive = async () => {
    if (!mainData || mainData.length === 0) {
      alert("Không có dữ liệu trong trạm làm việc để xuất Excel!");
      return;
    }

    const confirmed = window.confirm(
      `Xác nhận tạo tệp Excel (.xlsx) chứa ${mainData.length} dòng dữ liệu chế độ ${dataMode === "corp" ? "Doanh Nghiệp" : "Cá Thể"} và lưu trực tiếp lên Google Drive?`
    );
    if (!confirmed) return;

    setSyncingExcel(true);
    setErrorMessage(null);
    setStatusMessage("Đang kết xuất bảng dữ liệu sang Excel...");

    try {
      // Convert to excel
      const ws = XLSX.utils.json_to_sheet(mainData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "VSIC_Clean_Data");
      
      const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const excelBlob = new Blob([excelBuffer], { 
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
      });

      const modeStr = dataMode === "corp" ? "DoanhNghiep" : "CaThe";
      const dateStr = new Date().toISOString().slice(0, 10);
      const excelFileName = `VSIC_Cleaned_Export_${modeStr}_${dateStr}.xlsx`;

      setStatusMessage("Đang tạo tệp Excel trên Google Drive...");

      // 1. Tạo file metadata
      const metaResponse = await fetch("https://www.googleapis.com/drive/v3/files", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${googleAccessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          name: excelFileName,
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          description: `VSIC Cleaned Data Export. Mode: ${dataMode}. Rows: ${mainData.length}`
        })
      });

      if (!metaResponse.ok) {
        throw new Error("Lỗi khởi tạo tệp Excel trên Google Drive.");
      }

      const meta = await metaResponse.json();
      const fileId = meta.id;

      setStatusMessage("Đang tải tệp Excel lên...");

      // 2. Upload media
      const uploadResponse = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${googleAccessToken}`,
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        },
        body: excelBlob
      });

      if (!uploadResponse.ok) {
        throw new Error("Lỗi truyền dữ liệu tệp Excel.");
      }

      setStatusMessage(`Đã xuất và lưu tệp Excel "${excelFileName}" thành công lên Google Drive!`);
      setTimeout(() => setStatusMessage(null), 5000);
    } catch (error: any) {
      console.error("Lỗi xuất Excel lên Drive:", error);
      setErrorMessage(error.message || "Lỗi xuất và lưu tệp Excel lên Google Drive.");
    } finally {
      setSyncingExcel(false);
    }
  };

  // Phục hồi dữ liệu từ bản sao lưu Google Drive
  const handleRestore = async (file: BackupFile) => {
    setSyncing(true);
    setStatusMessage("Đang tải dữ liệu từ Google Drive...");
    setErrorMessage(null);

    try {
      const response = await fetch(`https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`, {
        headers: {
          Authorization: `Bearer ${googleAccessToken}`
        }
      });

      if (!response.ok) {
        throw new Error("Không thể tải nội dung tệp sao lưu.");
      }

      const backupState = await response.json();

      if (!backupState.mainData || !Array.isArray(backupState.mainData)) {
        throw new Error("Định dạng tệp sao lưu không hợp lệ hoặc bị hỏng.");
      }

      // Xác minh chế độ dữ liệu
      if (backupState.dataMode && backupState.dataMode !== dataMode) {
        const acceptModeMismatch = window.confirm(
          `Cảnh báo: Bản sao lưu này thuộc Chế độ ${backupState.dataMode === "corp" ? "Doanh Nghiệp" : "Cá Thể"} nhưng bạn đang ở Chế độ ${dataMode === "corp" ? "Doanh Nghiệp" : "Cá Thể"}. Bạn có chắc muốn khôi phục không?`
        );
        if (!acceptModeMismatch) {
          setSyncing(false);
          setStatusMessage(null);
          return;
        }
      }

      // Gọi hàm restore lên App.tsx
      onRestore({
        mainData: backupState.mainData,
        rawImportedData: backupState.rawImportedData || backupState.mainData,
        columns: backupState.columns || [],
        fileName: backupState.fileName || file.name,
        mapping: backupState.mapping || { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" },
        customColConfigs: backupState.customColConfigs || []
      });

      setStatusMessage(`Đã khôi phục thành công tệp sao lưu! Đã nạp ${backupState.mainData.length} dòng dữ liệu.`);
      setShowConfirmRestoreId(null);
      setTimeout(() => setStatusMessage(null), 4000);
    } catch (error: any) {
      console.error("Lỗi phục hồi:", error);
      setErrorMessage(error.message || "Lỗi phục hồi dữ liệu từ đám mây.");
    } finally {
      setSyncing(false);
    }
  };

  // Xóa tệp sao lưu khỏi Google Drive
  const handleDeleteBackup = async (fileId: string, fileName: string) => {
    const confirmed = window.confirm(`Bạn có chắc muốn xóa vĩnh viễn tệp sao lưu "${fileName}" khỏi Google Drive của bạn? Thao tác này không thể hoàn tác.`);
    if (!confirmed) return;

    setLoading(true);
    setErrorMessage(null);
    try {
      const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${googleAccessToken}`
        }
      });

      if (!response.ok) {
        throw new Error("Không thể xóa tệp trên Google Drive.");
      }

      setStatusMessage("Xóa tệp sao lưu thành công!");
      await fetchBackups();
      setTimeout(() => setStatusMessage(null), 3000);
    } catch (error: any) {
      console.error("Lỗi xóa tệp sao lưu:", error);
      setErrorMessage(error.message || "Không thể xóa tệp sao lưu.");
    } finally {
      setLoading(false);
    }
  };

  const formatSize = (bytesStr?: string) => {
    if (!bytesStr) return "N/A";
    const bytes = parseInt(bytesStr);
    if (isNaN(bytes)) return "N/A";
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
  };

  const formatTime = (isoString: string) => {
    try {
      const d = new Date(isoString);
      return d.toLocaleString("vi-VN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit"
      });
    } catch {
      return isoString;
    }
  };

  return (
    <div className="bg-white border border-indigo-100 rounded-2xl p-5 space-y-4 shadow-sm animate-fade-in font-sans">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-indigo-50/60 pb-3">
        <div className="flex items-center gap-2.5">
          <div className="p-2 bg-indigo-50 rounded-xl text-indigo-600">
            <Cloud className="w-5 h-5 animate-pulse" />
          </div>
          <div>
            <h4 className="text-sm font-extrabold text-slate-800 uppercase tracking-tight flex items-center gap-1.5">
              Đồng bộ đám mây Google Drive
            </h4>
            <p className="text-[11px] text-slate-500 mt-0.5">Sao lưu, phục hồi và lưu file Excel trực tuyến chủ động</p>
          </div>
        </div>
        
        {googleAccessToken && user && (
          <div className="flex items-center gap-2 text-xs">
            <div className="text-right">
              <span className="font-bold text-slate-700 block text-[11px]">{user.displayName || "Google User"}</span>
              <span className="text-[10px] text-slate-400 block">{user.email}</span>
            </div>
            {user.email && (
              <div className="w-8 h-8 rounded-full bg-indigo-600/10 text-indigo-700 font-bold flex items-center justify-center text-xs border border-indigo-100 uppercase shadow-inner">
                {user.email.slice(0, 2)}
              </div>
            )}
            <button 
              onClick={logout}
              className="text-[10px] text-rose-500 hover:text-rose-600 hover:underline cursor-pointer ml-1.5"
            >
              Đăng xuất
            </button>
          </div>
        )}
      </div>

      {/* THÔNG BÁO TRẠNG THÁI */}
      {statusMessage && (
        <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs rounded-xl flex items-center gap-2 animate-fade-in">
          <Check className="w-4 h-4 text-emerald-600 shrink-0" />
          <span className="font-semibold">{statusMessage}</span>
        </div>
      )}

      {errorMessage && (
        <div className="p-3 bg-rose-50 border border-rose-150 text-rose-800 text-xs rounded-xl flex items-center gap-2 animate-fade-in">
          <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
          <span className="font-semibold">{errorMessage}</span>
        </div>
      )}

      {/* CHƯA KẾT NỐI GOOGLE DRIVE */}
      {!googleAccessToken ? (
        <div className="bg-slate-50 border border-slate-200/60 rounded-xl p-5 text-center space-y-4">
          <div className="max-w-md mx-auto space-y-2">
            <CloudOff className="w-8 h-8 text-slate-400 mx-auto" />
            <p className="text-xs font-bold text-slate-700">Chưa kích hoạt Đồng bộ hóa đám mây</p>
            <p className="text-[11px] text-slate-500 leading-relaxed font-normal">
              Dữ liệu của bạn hiện đang được lưu trữ an toàn trong <strong>IndexedDB (Bộ nhớ trình duyệt địa phương)</strong> của bạn.
              Đăng nhập bằng tài khoản Google để có thể lưu trữ bản sao lưu trực tiếp vào <strong>Google Drive</strong> của bạn, dễ dàng phục hồi trên mọi máy tính và thiết bị khác nhau!
            </p>
          </div>

          <div className="flex justify-center pt-1.5">
            <button
              onClick={handleLogin}
              disabled={isLoggingIn}
              className="gsi-material-button flex items-center justify-center gap-2.5 bg-white border border-slate-300 hover:border-indigo-400 rounded-xl px-5 py-2.5 transition text-xs font-bold shadow-sm text-slate-700 cursor-pointer disabled:opacity-50"
            >
              {isLoggingIn ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin text-indigo-600" />
                  Đang ủy quyền...
                </>
              ) : (
                <>
                  <div className="gsi-material-button-icon w-4 h-4 shrink-0">
                    <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" style={{ display: "block" }}>
                      <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                      <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                      <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                      <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                    </svg>
                  </div>
                  <span className="gsi-material-button-contents">Kích hoạt &amp; Đăng nhập Google Drive</span>
                </>
              )}
            </button>
          </div>
        </div>
      ) : (
        /* ĐÃ KẾT NỐI GOOGLE DRIVE */
        <div className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {/* THAO TÁC ĐỒNG BỘ */}
            <div className="bg-slate-50/80 border border-slate-200/50 rounded-xl p-4 flex flex-col justify-between space-y-3">
              <div>
                <span className="text-[10px] font-extrabold text-indigo-600 uppercase tracking-wider block">
                  Công cụ xuất trực tuyến
                </span>
                <h5 className="text-xs font-bold text-slate-800 mt-1">Lưu trữ từ trạm làm việc</h5>
                <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                  Lưu trữ dữ liệu phân tích hiện tại trực tiếp vào tài khoản Google Drive cá nhân để phục hồi sau này hoặc chia sẻ liên ngành.
                </p>
              </div>

              <div className="flex flex-col sm:flex-row gap-2 pt-1">
                <button
                  onClick={handleBackup}
                  disabled={syncing || mainData.length === 0}
                  className="flex-1 flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white text-[11px] font-extrabold px-3 py-2.5 rounded-xl cursor-pointer transition active:scale-95 disabled:opacity-50"
                >
                  {syncing ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Upload className="w-3.5 h-3.5" />
                  )}
                  SAO LƯU PHIÊN LÀM VIỆC
                </button>

                <button
                  onClick={handleExportExcelToDrive}
                  disabled={syncingExcel || mainData.length === 0}
                  className="flex-1 flex items-center justify-center gap-2 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-[11px] font-extrabold px-3 py-2.5 rounded-xl cursor-pointer transition active:scale-95 disabled:opacity-50"
                >
                  {syncingExcel ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-700" />
                  ) : (
                    <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                  )}
                  LƯU FILE EXCEL (.XLSX)
                </button>
              </div>
            </div>

            {/* THÔNG TIN TỔNG QUAN */}
            <div className="bg-slate-50/80 border border-slate-200/50 rounded-xl p-4 flex flex-col justify-between space-y-2">
              <div>
                <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-wider block">
                  Trạng thái dữ liệu hiện tại
                </span>
                <h5 className="text-xs font-bold text-slate-800 mt-1 flex items-center gap-1.5">
                  <Database className="w-4 h-4 text-indigo-500" />
                  Chế độ {dataMode === "corp" ? "Doanh nghiệp" : "Cá thể"}
                </h5>
                <div className="grid grid-cols-2 gap-3 mt-2 text-[11px]">
                  <div className="p-2 bg-white rounded-lg border border-slate-200/60">
                    <span className="text-slate-400 block">Dữ liệu chính:</span>
                    <strong className="text-slate-800 text-xs font-bold">{mainData.length} dòng</strong>
                  </div>
                  <div className="p-2 bg-white rounded-lg border border-slate-200/60">
                    <span className="text-slate-400 block">Dữ liệu thô:</span>
                    <strong className="text-slate-800 text-xs font-bold">{rawImportedData.length} dòng</strong>
                  </div>
                </div>
              </div>

              <div className="text-[10px] text-slate-400 font-sans italic leading-normal pt-1 flex items-center gap-1">
                <AlertCircle className="w-3 h-3 text-slate-400 shrink-0" />
                <span>Mỗi chế độ phân tích dữ liệu có tệp sao lưu riêng độc lập.</span>
              </div>
            </div>
          </div>

          {/* DANH SÁCH BẢN SAO LƯU */}
          <div className="space-y-2.5 pt-1">
            <div className="flex items-center justify-between border-b border-slate-100 pb-1.5">
              <h5 className="text-xs font-bold text-slate-800 uppercase tracking-tight flex items-center gap-1.5">
                <FileJson className="w-4 h-4 text-indigo-500" />
                Bản sao lưu có sẵn ({dataMode === "corp" ? "Doanh nghiệp" : "Cá thể"})
              </h5>
              <button
                onClick={fetchBackups}
                disabled={loading}
                className="text-indigo-600 hover:text-indigo-700 text-[11px] font-extrabold flex items-center gap-1 cursor-pointer hover:underline disabled:opacity-50"
              >
                <RefreshCw className={`w-3 h-3 ${loading ? "animate-spin" : ""}`} />
                LÀM MỚI
              </button>
            </div>

            {loading ? (
              <div className="text-center py-8 flex flex-col items-center justify-center space-y-2 text-slate-400">
                <Loader2 className="w-6 h-6 animate-spin text-indigo-500" />
                <span className="text-xs font-medium">Đang tìm các bản sao lưu trên Google Drive...</span>
              </div>
            ) : backups.length === 0 ? (
              <div className="text-center py-6 bg-slate-50/50 rounded-xl border border-dashed border-slate-200/80 text-slate-400 text-xs flex flex-col items-center justify-center space-y-1">
                <CloudLightning className="w-6 h-6 text-slate-300" />
                <p className="font-semibold text-slate-500 text-[11px]">Chưa có bản sao lưu nào được lưu trên Google Drive</p>
                <p className="text-[10px] text-slate-400">Ấn nút "Sao lưu phiên làm việc" ở trên để lưu trữ bản đầu tiên.</p>
              </div>
            ) : (
              <div className="max-h-56 overflow-y-auto space-y-2 pr-1.5">
                {backups.map((file) => (
                  <div
                    key={file.id}
                    className="p-3 bg-slate-50/80 hover:bg-slate-50 border border-slate-200/65 hover:border-slate-300 rounded-xl flex items-center justify-between gap-3 transition"
                  >
                    <div className="flex items-center gap-2.5 overflow-hidden">
                      <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg shrink-0">
                        <FileJson className="w-4 h-4" />
                      </div>
                      <div className="overflow-hidden">
                        <p className="text-xs font-bold text-slate-800 truncate" title={file.name}>
                          {file.name}
                        </p>
                        <p className="text-[10px] text-slate-400 mt-0.5 flex items-center gap-2">
                          <span className="flex items-center gap-1 font-mono">
                            <Calendar className="w-3 h-3" /> {formatTime(file.createdTime)}
                          </span>
                          <span>• {formatSize(file.size)}</span>
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-1.5 shrink-0">
                      {showConfirmRestoreId === file.id ? (
                        <div className="flex items-center gap-1 animate-fade-in">
                          <button
                            onClick={() => handleRestore(file)}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white text-[10px] font-bold px-2.5 py-1.5 rounded-lg cursor-pointer transition active:scale-95"
                          >
                            Xác nhận tải
                          </button>
                          <button
                            onClick={() => setShowConfirmRestoreId(null)}
                            className="bg-slate-200 hover:bg-slate-300 text-slate-700 text-[10px] font-bold px-2 py-1.5 rounded-lg cursor-pointer transition"
                          >
                            Hủy
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => setShowConfirmRestoreId(file.id)}
                          className="p-1.5 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 rounded-lg transition cursor-pointer flex items-center gap-1 text-[11px] font-extrabold"
                          title="Tải bản sao lưu này về trạm dữ liệu hiện tại"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>Khôi phục</span>
                        </button>
                      )}

                      <button
                        onClick={() => handleDeleteBackup(file.id, file.name)}
                        className="p-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-600 rounded-lg transition cursor-pointer"
                        title="Xóa tệp khỏi Google Drive"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
