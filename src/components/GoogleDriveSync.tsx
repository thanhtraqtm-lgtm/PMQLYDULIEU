import React, { useState } from "react";
import { Cloud, CloudUpload, CloudDownload, HardDrive, CheckCircle2, AlertCircle, RefreshCw, Database } from "lucide-react";

export interface GoogleDriveSyncProps {
  mainData: any[];
  rawImportedData: any[];
  columns: string[];
  fileName: string;
  mapping: any;
  customColConfigs: any;
  dataMode: any;
  embedded?: boolean;
  onRestore: (state: {
    mainData: any[];
    rawImportedData: any[];
    columns: string[];
    fileName: string;
    mapping: any;
    customColConfigs: any;
  }) => void;
}

export const GoogleDriveSync: React.FC<GoogleDriveSyncProps> = ({
  mainData,
  rawImportedData,
  columns,
  fileName,
  mapping,
  customColConfigs,
  dataMode,
  embedded = false,
  onRestore,
}) => {
  const [syncStatus, setSyncStatus] = useState<string | null>(null);
  const [backupHistory, setBackupHistory] = useState<
    { id: string; date: string; rowsCount: number; fileName: string; snapshot: any }[]
  >(() => {
    try {
      const saved = localStorage.getItem("vtong_backups_list");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const handleSaveSnapshot = () => {
    if (mainData.length === 0) {
      alert("Không có dữ liệu trong bộ nhớ để tạo bản sao lưu!");
      return;
    }
    const snapshot = {
      mainData,
      rawImportedData,
      columns,
      fileName,
      mapping,
      customColConfigs,
    };
    const newEntry = {
      id: "backup_" + Date.now(),
      date: new Date().toLocaleString("vi-VN"),
      rowsCount: mainData.length,
      fileName: fileName || "DuLieu_VTong.xlsx",
      snapshot,
    };
    const updated = [newEntry, ...backupHistory.slice(0, 4)];
    setBackupHistory(updated);
    try {
      localStorage.setItem("vtong_backups_list", JSON.stringify(updated));
    } catch (e) {
      console.warn("Storage quota limit reached for local snapshot list");
    }
    setSyncStatus("Đã tạo điểm sao lưu dữ liệu cục bộ thành công!");
    setTimeout(() => setSyncStatus(null), 3000);
  };

  const handleExportJSON = () => {
    if (mainData.length === 0) {
      alert("Không có dữ liệu để xuất gói sao lưu!");
      return;
    }
    const bundle = {
      version: "1.0",
      exportTime: new Date().toISOString(),
      fileName,
      mapping,
      customColConfigs,
      dataMode,
      columns,
      rowCount: mainData.length,
      mainData,
      rawImportedData,
    };
    const blob = new Blob([JSON.stringify(bundle, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `VTong_Backup_${(fileName || "Data").replace(/\.[^/.]+$/, "")}_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportJSON = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const bundle = JSON.parse(evt.target?.result as string);
        if (!bundle.mainData || !Array.isArray(bundle.mainData)) {
          throw new Error("Tệp không đúng định dạng bản sao lưu hợp lệ.");
        }
        onRestore({
          mainData: bundle.mainData,
          rawImportedData: bundle.rawImportedData || bundle.mainData,
          columns: bundle.columns || Object.keys(bundle.mainData[0] || {}),
          fileName: bundle.fileName || file.name,
          mapping: bundle.mapping || {},
          customColConfigs: bundle.customColConfigs || [],
        });
        alert(`Đã khôi phục thành công ${bundle.mainData.length} dòng dữ liệu từ bản sao lưu!`);
      } catch (err: any) {
        alert("Lỗi đọc bản sao lưu: " + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div className={embedded ? "space-y-4 font-sans" : "bg-white rounded-xl shadow-sm border border-slate-200 p-6 space-y-6 font-sans"}>
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-200/70">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-sky-100 flex items-center justify-center text-sky-700">
            <Cloud className="w-4 h-4" />
          </div>
          <div>
            <h4 className="font-bold text-slate-800 text-xs">Đồng Bộ & Bản Sao Lưu</h4>
            <p className="text-[11px] text-slate-500">
              Tạo điểm khôi phục nhanh hoặc xuất tệp sao lưu định dạng JSON.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={handleSaveSnapshot}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-[#0284c7] hover:bg-[#0369a1] text-white text-xs font-bold rounded-none shadow-xs transition-colors cursor-pointer whitespace-nowrap border-0"
          >
            <CloudUpload className="w-3.5 h-3.5" />
            Tạo Điểm Lưu Trữ
          </button>
          <button
            onClick={handleExportJSON}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-[#4f46e5] hover:bg-[#4338ca] text-white text-xs font-bold rounded-none shadow-xs transition-colors cursor-pointer whitespace-nowrap border-0"
          >
            <CloudDownload className="w-3.5 h-3.5" />
            Tải File JSON (.json)
          </button>
          <label className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-[#7c3aed] hover:bg-[#6d28d9] text-white text-xs font-bold rounded-none shadow-xs transition-colors cursor-pointer whitespace-nowrap border-0">
            <HardDrive className="w-3.5 h-3.5" />
            Khôi Phục Từ File
            <input type="file" accept=".json" onChange={handleImportJSON} className="hidden" />
          </label>
        </div>
      </div>

      {syncStatus && (
        <div className="p-2.5 bg-emerald-50 border border-emerald-300 rounded-none flex items-center gap-2 text-emerald-800 text-xs font-bold">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{syncStatus}</span>
        </div>
      )}

      {/* Snapshot History Table */}
      <div className="space-y-2.5">
        <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
          <Database className="w-3.5 h-3.5 text-sky-700" />
          Các Điểm Sao Lưu Gần Đây Trong Phiên Làm Việc
        </h4>

        {backupHistory.length === 0 ? (
          <div className="py-6 text-center text-slate-500 bg-sky-50/40 rounded-none border border-dashed border-sky-200 text-xs">
            Chưa có điểm sao lưu nào được ghi nhận. Bấm "Tạo Điểm Lưu Trữ" để lưu trạng thái hiện tại.
          </div>
        ) : (
          <div className="divide-y divide-slate-100 border border-slate-200 rounded-none overflow-hidden bg-white">
            {backupHistory.map((item, idx) => (
              <div key={item.id} className="p-3 flex items-center justify-between hover:bg-sky-50/50 transition-colors">
                <div>
                  <p className="text-xs font-bold text-slate-800">{item.fileName}</p>
                  <p className="text-[11px] text-slate-500">
                    Thời gian: {item.date} • Quy mô: {item.rowsCount.toLocaleString("vi-VN")} dòng
                  </p>
                </div>
                <button
                  onClick={() => {
                    if (confirm(`Bạn có chắc muốn khôi phục lại điểm lưu lúc ${item.date}?`)) {
                      onRestore(item.snapshot);
                    }
                  }}
                  className="px-2.5 py-1 bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs rounded-none shadow-2xs transition-colors cursor-pointer border-0"
                >
                  Khôi Phục Bản Này
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
export default GoogleDriveSync;
