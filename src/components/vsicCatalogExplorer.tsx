import React, { useState, useMemo, useRef, useEffect } from "react";
import { 
  Search, 
  Database, 
  Layers, 
  Filter, 
  CheckCircle2, 
  ChevronRight, 
  Upload, 
  Trash2, 
  Download, 
  AlertTriangle, 
  Check, 
  FileSpreadsheet,
  HelpCircle,
  ToggleLeft
} from "lucide-react";
import { vsicRawData, clearAllSectorsInVSIC, loadSectorsIntoVSIC, vsicParentMap, clearAllParentsInVSIC, loadParentsIntoVSIC } from "../data/vsic";
import * as XLSX from "xlsx";

const VsicCatalogExplorer = React.memo(function VsicCatalogExplorer() {
  const [search, setSearch] = useState("");
  const [level, setLevel] = useState("ALL");
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [isDragging, setIsDragging] = useState(false);
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{
    success?: boolean;
    message?: string;
    addedCount?: number;
  }>({});

  const [pureMode, setPureMode] = useState<boolean>(() => {
    try {
      const stored = localStorage.getItem("custom_vsic_is_pure");
      return stored === "true";
    } catch {
      return false;
    }
  });

  const [customCatalog, setCustomCatalog] = useState<{ [key: string]: string }>(() => {
    try {
      const stored = localStorage.getItem("custom_vsic_data");
      return stored ? JSON.parse(stored) : {};
    } catch {
      return {};
    }
  });

  const [editingCode, setEditingCode] = useState<string | null>(null);
  const [editingName, setEditingName] = useState<string>("");

  const handleSaveEdit = (code: string) => {
    if (!editingName.trim()) return;
    const updated = { ...customCatalog, [code]: editingName.trim() };
    setCustomCatalog(updated);
    try {
      localStorage.setItem("custom_vsic_data", JSON.stringify(updated));
    } catch (e) {
      console.warn("localStorage error:", e);
    }
    vsicRawData[code] = editingName.trim();
    setEditingCode(null);
  };

  useEffect(() => {
    if (pureMode) {
      clearAllSectorsInVSIC();
    }
    Object.assign(vsicRawData, customCatalog);
  }, [pureMode, customCatalog]);

  const vsicList = useMemo(() => {
    return Object.entries(vsicRawData).map(([code, name]) => {
      let cap = 5;
      if (/^[A-Z]$/.test(code)) cap = 1;
      else if (code.length === 2) cap = 2;
      else if (code.length === 3) cap = 3;
      else if (code.length === 4) cap = 4;
      return { code, name, cap };
    });
  }, [customCatalog, pureMode]);

  const quickSearches = useMemo(() => {
    const rawKeys = Object.keys(vsicRawData);
    if (rawKeys.length === 0) return [];
    
    const candidates = ["56101", "56302", "47521", "68104", "Bán lẻ", "Nhà hàng", "Cà phê", "Sản xuất"];
    const valSample = Object.values(vsicRawData).slice(0, 500);
    
    return candidates.filter(cand => {
      if (vsicRawData[cand]) return true;
      if (isNaN(Number(cand))) {
        const lowerCand = cand.toLowerCase();
        return valSample.some(name => name.toLowerCase().includes(lowerCand));
      }
      return false;
    });
  }, [customCatalog, pureMode]);

  const filteredList = useMemo(() => {
    const query = search.trim().toLowerCase();
    return vsicList.filter((item) => {
      if (level !== "ALL" && String(item.cap) !== level) {
        return false;
      }
      if (query) {
        const matchCode = item.code.toLowerCase().includes(query);
        const matchName = item.name.toLowerCase().includes(query);
        return matchCode || matchName;
      }
      return true;
    });
  }, [search, level, vsicList]);

  const displayLimit = 150;
  const displayedItems = filteredList.slice(0, displayLimit);

  const handleDownloadTemplate = (format: "csv" | "xlsx") => {
    const headers = ["NganhCap1", "NganhCap2", "NganhCap3", "NganhCap4", "NganhCap5", "TenNganh"];
    const rows = [
      ["A", "", "", "", "", "Nông nghiệp, lâm nghiệp và thủy sản"],
      ["A", "01", "", "", "", "Nông nghiệp và hoạt động dịch vụ có liên quan"],
      ["A", "01", "011", "", "", "Trồng cây hàng năm"],
      ["A", "01", "011", "0111", "", "Trồng lúa"],
      ["A", "01", "011", "0111", "01110", "Trồng lúa"],
      ["A", "01", "011", "0112", "", "Trồng ngô và cây lương thực có hạt khác"],
      ["A", "01", "011", "0112", "01120", "Trồng ngô và cây lương thực có hạt khác"]
    ];

    if (format === "csv") {
      let csvContent = "data:text/csv;charset=utf-8,\uFEFF";
      csvContent += headers.join(",") + "\n";
      rows.forEach(r => {
        const formattedRow = r.map(val => `"${val.replace(/"/g, '""')}"`);
        csvContent += formattedRow.join(",") + "\n";
      });
      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", "Mau_Danh_Muc_Nganh_5_Cap.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } else {
      const wsData = [headers, ...rows];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      XLSX.utils.book_append_sheet(wb, ws, "Danh_Muc_Mau_5_Cap");
      XLSX.writeFile(wb, "Mau_Danh_Muc_Nganh_5_Cap.xlsx");
    }
  };

  const handleProcessFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        const rawRows = XLSX.utils.sheet_to_json<any>(sheet, { header: 1 });
        if (rawRows.length < 2) {
          setUploadStatus({
            success: false,
            message: "Tệp tải lên trống hoặc không có dòng tiêu đề (Dòng 1)."
          });
          return;
        }

        const headerRow = rawRows[0].map((h: any) => String(h || "").trim().toLowerCase());
        
        let cap1Idx = -1;
        let cap2Idx = -1;
        let cap3Idx = -1;
        let cap4Idx = -1;
        let cap5Idx = -1;
        let tenNganhIdx = -1;

        let codeIdxSimple = -1;
        let nameIdxSimple = -1;

        headerRow.forEach((name: string, index: number) => {
          if (name.includes("nganhcap1") || name.includes("cấp 1") || name === "cap1") cap1Idx = index;
          else if (name.includes("nganhcap2") || name.includes("cấp 2") || name === "cap2") cap2Idx = index;
          else if (name.includes("nganhcap3") || name.includes("cấp 3") || name === "cap3") cap3Idx = index;
          else if (name.includes("nganhcap4") || name.includes("cấp 4") || name === "cap4") cap4Idx = index;
          else if (name.includes("nganhcap5") || name.includes("cấp 5") || name === "cap5") cap5Idx = index;
          else if (name.includes("tennganh") || name.includes("tên ngành") || name.includes("tên gọi") || name.includes("mô tả")) tenNganhIdx = index;

          if (name.includes("manganh") || name.includes("mã ngành") || name === "code" || name === "ma_nganh") codeIdxSimple = index;
          if (name.includes("tennganh") || name.includes("tên ngành") || name === "name" || name === "ten_nganh") nameIdxSimple = index;
        });

        const customSectors: { [key: string]: string } = {};
        const customParents: { [key: string]: string } = {};
        let added = 0;

        for (let i = 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          let finalCode = "";
          let finalName = "";

          const hasDetailedCols = (cap1Idx !== -1 || cap2Idx !== -1 || cap3Idx !== -1 || cap4Idx !== -1 || cap5Idx !== -1) && tenNganhIdx !== -1;

          if (hasDetailedCols) {
            const c5Val = cap5Idx !== -1 && row[cap5Idx] ? String(row[cap5Idx]).trim() : "";
            const c4Val = cap4Idx !== -1 && row[cap4Idx] ? String(row[cap4Idx]).trim() : "";
            const c3Val = cap3Idx !== -1 && row[cap3Idx] ? String(row[cap3Idx]).trim() : "";
            const c2Val = cap2Idx !== -1 && row[cap2Idx] ? String(row[cap2Idx]).trim() : "";
            const c1Val = cap1Idx !== -1 && row[cap1Idx] ? String(row[cap1Idx]).trim() : "";

            finalCode = c5Val || c4Val || c3Val || c2Val || c1Val;
            finalName = String(row[tenNganhIdx] || "").trim();

            if (finalCode && finalName) {
              const hierarchy = [c1Val, c2Val, c3Val, c4Val, c5Val].filter(Boolean);
              for (let hIdx = 1; hIdx < hierarchy.length; hIdx++) {
                const child = hierarchy[hIdx];
                const parent = hierarchy[hIdx - 1];
                if (child && parent) {
                  customParents[child] = parent;
                }
              }
            }
          } else if (codeIdxSimple !== -1 && nameIdxSimple !== -1) {
            finalCode = String(row[codeIdxSimple] || "").trim();
            finalName = String(row[nameIdxSimple] || "").trim();
          }

          if (finalCode && finalName) {
            customSectors[finalCode] = finalName;
            added++;
          }
        }

        if (added === 0) {
          setUploadStatus({
            success: false,
            message: "Không lọc được dòng dữ liệu ngành nào hợp lệ. Vui lòng kiểm tra lại cấu trúc tiêu đề cột của file."
          });
          return;
        }

        const mergedCatalog = { ...customCatalog, ...customSectors };
        setCustomCatalog(mergedCatalog);

        try {
          localStorage.setItem("custom_vsic_data", JSON.stringify(mergedCatalog));
          localStorage.setItem("custom_vsic_parents", JSON.stringify({ ...vsicParentMap, ...customParents }));
        } catch (storageErr) {
          console.warn("Storage write limit reached, loaded in-memory only.", storageErr);
        }

        Object.assign(vsicRawData, customSectors);
        Object.assign(vsicParentMap, customParents);

        setUploadStatus({
          success: true,
          message: `Khớp dữ liệu thành công! Đã giải mã và nạp bổ sung ${added} mã ngành kinh tế chi tiết vào bộ nhớ.`,
          addedCount: added
        });
      } catch (err: any) {
        setUploadStatus({
          success: false,
          message: `Lỗi phân tích cú pháp tệp Excel: ${err.message || err}`
        });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleProcessFile(file);
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleProcessFile(file);
  };

  const executeResetCatalog = () => {
    localStorage.removeItem("custom_vsic_data");
    localStorage.removeItem("custom_vsic_parents");
    localStorage.setItem("custom_vsic_is_pure", "true");
    setCustomCatalog({});
    setPureMode(true);
    clearAllSectorsInVSIC();
    clearAllParentsInVSIC();
    setShowResetConfirm(false);
    try {
      alert("Đã xóa sạch danh mục cũ trong bộ nhớ! Hệ thống đang ở chế độ chờ, xin mời bạn tải lên tệp Excel danh mục của bạn.");
    } catch (e) {
      console.warn("Alert blocked:", e);
    }
    window.location.reload();
  };

  const handlePureModeToggle = (e: React.ChangeEvent<HTMLInputElement>) => {
    const isChecked = e.target.checked;
    setPureMode(isChecked);
    localStorage.setItem("custom_vsic_is_pure", isChecked ? "true" : "false");
    
    if (isChecked) {
      alert("Đã kích hoạt Chế độ Thuần khiết! Hệ thống sẽ chỉ sử dụng duy nhất tệp danh mục Excel do quý khách tự nạp.");
      window.location.reload();
    } else {
      alert("Đã tắt Chế độ Thuần khiết! Hệ thống sẽ phối hợp danh mục của quý khách cùng bộ khung mặc định.");
      window.location.reload();
    }
  };

  const customKeys = Object.keys(customCatalog);

  return (
    <div className="space-y-6 font-sans text-slate-800">
      {/* 1. KHU VỰC TẢI LÊN DANH MỤC TRỰC QUAN */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* Nạp tệp mới */}
        <div className="lg:col-span-2 bg-white border border-slate-200 rounded-2xl p-6 space-y-4 shadow-xl">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-50 p-2 rounded-lg border border-indigo-100">
              <Upload className="w-5 h-5 text-indigo-600" />
            </div>
            <div>
              <h4 className="text-sm font-bold text-slate-900 uppercase tracking-wider">Nạp bổ sung danh mục ngành của bạn (Excel / CSV)</h4>
              <p className="text-xs text-slate-500">Tải lên danh mục riêng để ghép mã thắt nút, giải phóng hoàn toàn lỗi định dạng danh mục mặc định của bạn.</p>
            </div>
          </div>

          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all ${
              isDragging 
                ? "border-indigo-500 bg-indigo-50" 
                : "border-slate-300 hover:border-slate-400 bg-slate-50/50 hover:bg-slate-50"
            }`}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".csv,.xlsx,.xls"
              className="hidden"
            />
            <FileSpreadsheet className="w-10 h-10 text-slate-400 mx-auto mb-3" />
            <p className="text-sm text-slate-800 font-semibold">Tải lên tệp CSV hoặc Excel chứa danh mục mở rộng</p>
            <p className="text-xs text-slate-500 mt-1">Kéo thả file vào đây hoặc click để duyệt từ máy tính</p>
            <p className="text-[10px] text-indigo-700 font-mono mt-2 italic">*Hệ thống tự đối chiếu thông minh cột [NganhCap1] đến [NganhCap5] thành Mã chuẩn</p>
          </div>

          {uploadStatus.message && (
            <div className={`p-4 rounded-xl text-xs border flex items-start gap-2.5 ${
              uploadStatus.success 
                ? "bg-emerald-50 border-emerald-200 text-emerald-800" 
                : "bg-rose-50 border-rose-200 text-rose-800"
            }`}>
              {uploadStatus.success ? <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600" /> : <AlertTriangle className="w-4 h-4 shrink-0 text-rose-600" />}
              <div>
                <p className="font-bold">{uploadStatus.success ? "NẠP THÀNH CÔNG" : "THẤT BẠI KHI NẠP"}</p>
                <p className="mt-0.5 leading-relaxed">{uploadStatus.message}</p>
              </div>
            </div>
          )}
        </div>

        {/* Thư viện quản lý bộ mẫu chuẩn */}
        <div className="bg-white border border-slate-200 rounded-2xl p-6 flex flex-col justify-between shadow-xl space-y-4">
          <div className="space-y-4">
            <h4 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
              <Download className="w-4 h-4 text-indigo-650" /> TẢI FILE MẪU BAN ĐẦU
            </h4>
            <p className="text-xs text-slate-500 leading-normal">
              Bạn chưa rõ cấu trúc điền file thế nào? Hãy tải về file mẫu của chúng tôi, mở trên Excel, copy danh sách mã của bạn dán vào rồi nạp tệp 5 cấp!
            </p>
            <div className="grid grid-cols-2 gap-3 pt-2">
              <button 
                onClick={() => handleDownloadTemplate("xlsx")}
                className="flex items-center justify-center gap-1.5 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-700 py-2.5 px-3 rounded-lg text-xs font-bold cursor-pointer transition-all shadow-sm"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" /> File mẫu Excel
              </button>
              <button 
                onClick={() => handleDownloadTemplate("csv")}
                className="flex items-center justify-center gap-1.5 bg-sky-50 hover:bg-sky-100 border border-sky-200 text-sky-700 py-2.5 px-3 rounded-lg text-xs font-bold cursor-pointer transition-all shadow-sm"
              >
                <Download className="w-3.5 h-3.5" /> File mẫu CSV
              </button>
            </div>
          </div>

          <div className="border-t border-slate-100 pt-4 space-y-3">
            <div className="flex items-center justify-between bg-purple-50 border border-purple-100 rounded-xl p-3 mb-2">
              <div>
                <p className="text-xs font-bold text-purple-900">Chế độ Danh mục Riêng</p>
                <p className="text-[10px] text-slate-500 mt-0.5 leading-tight">Mở tính năng xóa danh mục hệ thống mặc định; chỉ xử lý duy nhất danh mục 5 cấp quý khách nạp.</p>
              </div>
              <label className="relative inline-flex items-center cursor-pointer select-none">
                <input 
                  type="checkbox" 
                  checked={pureMode} 
                  onChange={handlePureModeToggle} 
                  className="sr-only peer" 
                />
                <div className="w-9 h-5 bg-slate-250 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-purple-600"></div>
              </label>
            </div>

            <div className="flex items-center justify-between text-xs text-slate-600">
              <span>Đã nạp bổ sung:</span>
              <strong className="text-purple-750 font-mono text-sm">{customKeys.length} mã ngành</strong>
            </div>
            
            {(customKeys.length > 0 || pureMode) && (
              <div className="space-y-2">
                {!showResetConfirm ? (
                  <button
                    onClick={() => setShowResetConfirm(true)}
                    className="w-full flex items-center justify-center gap-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-700 py-2 rounded-lg text-xs font-bold cursor-pointer transition-all shadow-sm"
                    id="trigger-reset-catalog-btn"
                  >
                    <Trash2 className="w-3.5 h-3.5" /> Xóa sạch bộ nhớ &amp; Nạp tệp mới
                  </button>
                ) : (
                  <div className="bg-rose-50 border border-rose-200 rounded-xl p-2.5 space-y-2">
                    <p className="text-[10px] text-rose-800 text-center font-bold leading-normal uppercase">Xác nhận xóa sạch danh mục tiêu chuẩn khỏi trình duyệt?</p>
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        onClick={executeResetCatalog}
                        className="bg-rose-600 hover:bg-rose-700 text-white rounded py-1 px-2 text-[10px] font-bold transition-all cursor-pointer text-center shadow-sm"
                        id="confirm-reset-catalog-btn"
                      >
                        Có, Xóa hết
                      </button>
                      <button
                        onClick={() => setShowResetConfirm(false)}
                        className="bg-slate-100 hover:bg-slate-200 text-slate-700 rounded py-1 px-2 text-[10px] font-bold transition-all cursor-pointer text-center border border-slate-250"
                        id="cancel-reset-catalog-btn"
                      >
                        Hủy bỏ
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

      </div>

      {/* 2. KHU VỰC TRA CỨU DANH MỤC TRỰC QUAN */}
      <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-xl">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-5">
          <div className="space-y-1">
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Database className="w-5 h-5 text-indigo-600" /> DANH MỤC HỆ THỐNG PHÂN CẤP HOẠT ĐỘNG (ĐÃ HỢP NHẤT)
            </h3>
            <p className="text-xs text-slate-500">
              Dữ liệu tích hợp gồm tất cả các mã ngành chuẩn quốc gia và mã ngành tùy chọn bạn đã nạp vào hệ thống để tra cứu, tìm kiếm.
            </p>
          </div>
          <div className="bg-indigo-50 border border-indigo-150 text-indigo-800 font-mono text-xs px-3 py-1.5 rounded-lg flex items-center gap-1.5 self-start md:self-center">
            <span className="relative flex h-2 w-2">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-455 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-indigo-600"></span>
            </span>
            Sẵn sàng kết nối: {vsicList.length} mã chuẩn
          </div>
        </div>

        {/* Tra cứu nhanh */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="md:col-span-2 relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
              <Search className="w-4 h-4" />
            </span>
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Nhập mã ngành (ví dụ: 5610) hoặc từ khóa (bán lẻ, cà phê...)"
              className="w-full bg-slate-50 border border-slate-300 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl pl-9 pr-4 py-2.5 text-sm text-slate-800 placeholder-slate-450 outline-none transition-all font-medium"
              id="vsic-catalog-search-input"
            />
          </div>

          <div className="relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
              <Filter className="w-4 h-4" />
            </span>
            <select
              value={level}
              onChange={(e) => setLevel(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl pl-9 pr-4 py-2.5 text-sm text-slate-800 outline-none transition-all cursor-pointer font-medium"
              id="vsic-catalog-level-select"
            >
              <option value="ALL">Tất cả Phân cấp (Cấp 1 - 5)</option>
              <option value="1">Cấp 1 (Phân khu thư mục lớn)</option>
              <option value="2">Cấp 2 (Nhóm chính)</option>
              <option value="3">Cấp 3 (Nhóm cấp dưới)</option>
              <option value="4">Cấp 4 (Nhóm tiểu ngành)</option>
              <option value="5">Cấp 5 (Nhóm kinh tế chi tiết nhất)</option>
            </select>
          </div>

          <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 flex items-center justify-between">
            <div className="text-slate-500 text-xs font-semibold">Phù hợp:</div>
            <div className="text-right">
              <span className="text-sm font-black text-emerald-700">{filteredList.length}</span>
              <span className="text-slate-500 text-[10px] ml-1">nhóm</span>
            </div>
          </div>
        </div>

        {/* Gợi ý từ khóa tìm kiếm nhanh */}
        <div className="flex flex-wrap items-center gap-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
          <span className="text-xs text-slate-500 font-bold select-none">Từ khóa tìm nhanh:</span>
          {quickSearches.map((qs) => (
            <button
              key={qs}
              onClick={() => setSearch(qs)}
              className="bg-white hover:bg-indigo-50 hover:text-indigo-700 border border-slate-300 hover:border-indigo-300 px-2.5 py-1 rounded-lg text-xs font-mono text-slate-700 transition-all cursor-pointer font-medium shadow-sm"
              id={`quick-search-${qs}`}
            >
              {qs}
            </button>
          ))}
          {search && (
            <button
              onClick={() => setSearch("")}
              className="text-xs text-rose-600 hover:text-rose-700 ml-auto font-bold px-2"
              id="clear-search-btn"
            >
              Xóa tìm kiếm
            </button>
          )}
        </div>

        {/* Danh sách bảng kết quả cứu */}
        <div className="border border-slate-250 rounded-xl overflow-hidden bg-white shadow-sm">
          <div className="max-h-[500px] overflow-y-auto">
            <table className="w-full text-left border-collapse" id="vsic-catalog-table">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-250 sticky top-0 z-10 text-slate-700 font-mono text-xs">
                  <th className="p-3.5 text-center w-[90px] font-bold">MÃ NGÀNH</th>
                  <th className="p-3.5 w-[110px] text-center font-bold">BẬC PHÂN CẤP</th>
                  <th className="p-3.5 text-left font-bold">MÔ TẢ TÊN GỌI CHUẨN VSIC QUỐC GIA</th>
                  <th className="p-3.5 text-center w-[120px] font-bold">NGUỒN MÃ</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-slate-800 text-[11.5px]">
                {displayedItems.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="p-8 text-center text-sm text-slate-400 italic">
                      Không tìm thấy mã VSIC hoạt động nào khớp với bộ lọc và điều kiện tìm kiếm của bạn. Hãy thử nhập lại!
                    </td>
                  </tr>
                ) : (
                  displayedItems.map((item, idx) => {
                    let capBadgeColor = "";
                    switch (item.cap) {
                      case 1:
                        capBadgeColor = "bg-rose-50 text-rose-700 border border-rose-200";
                        break;
                      case 2:
                        capBadgeColor = "bg-orange-50 text-orange-700 border border-orange-200";
                        break;
                      case 3:
                        capBadgeColor = "bg-amber-50 text-amber-700 border border-amber-200";
                        break;
                      case 4:
                        capBadgeColor = "bg-teal-50 text-teal-700 border border-teal-200";
                        break;
                      default:
                        capBadgeColor = "bg-sky-50 text-sky-700 border border-sky-200";
                    }

                    const isCustom = !!customCatalog[item.code];

                    return (
                      <tr
                        key={item.code}
                        className="hover:bg-slate-50 transition-colors group cursor-pointer border-b border-slate-100"
                        onClick={() => setSearch(item.code)}
                      >
                        <td className="p-3.5 text-center font-mono font-bold text-amber-850 text-sm group-hover:text-indigo-700 transition-colors">
                          {item.code}
                        </td>
                        <td className="p-3.5 text-center">
                          <span className={`inline-block py-1 px-2.5 rounded-lg text-[10px] font-bold tracking-wider ${capBadgeColor}`}>
                            CẤP {item.cap}
                          </span>
                        </td>
                        <td className="p-3.5 text-sm text-slate-800 font-sans group-hover:text-slate-900 transition-colors leading-relaxed">
                          {editingCode === item.code ? (
                            <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                              <input
                                type="text"
                                value={editingName}
                                onChange={(e) => setEditingName(e.target.value)}
                                className="flex-1 bg-white border border-indigo-500 rounded-lg px-3 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                                autoFocus
                                onKeyDown={(e) => {
                                  if (e.key === "Enter") handleSaveEdit(item.code);
                                  else if (e.key === "Escape") setEditingCode(null);
                                }}
                              />
                              <button
                                onClick={() => handleSaveEdit(item.code)}
                                className="p-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 rounded-lg cursor-pointer transition-all flex items-center justify-center shadow-sm"
                                title="Lưu"
                              >
                                <Check className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setEditingCode(null)}
                                className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-250 rounded-lg cursor-pointer transition-all flex items-center justify-center"
                                title="Hủy"
                              >
                                <span className="text-xs font-bold px-1 select-none">Hủy</span>
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-between gap-4">
                              <span className="font-medium text-[#0f172a]">{item.name}</span>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setEditingCode(item.code);
                                  setEditingName(item.name);
                                }}
                                className="opacity-0 group-hover:opacity-100 hover:text-indigo-600 p-1.5 rounded bg-slate-100 hover:bg-slate-200 border border-slate-200 transition-all flex items-center justify-center cursor-pointer"
                                title="Sửa tên ngành"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="lucide lucide-pencil"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
                              </button>
                            </div>
                          )}
                        </td>
                        <td className="p-3.5 text-center">
                          {isCustom ? (
                            <span className="inline-block py-0.5 px-2 bg-purple-50 text-purple-700 border border-purple-200 rounded-md text-[10px] font-bold">
                              Do Tự Nạp
                            </span>
                          ) : (
                            <span className="inline-block py-0.5 px-2 bg-slate-100 text-slate-500 rounded-md text-[10px] border border-slate-200 font-medium">
                              Hệ thống
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
          {filteredList.length > displayLimit && (
            <div className="p-3.5 bg-slate-50 border-t border-slate-250 text-center text-xs text-slate-400 italic">
              Hiển thị tối đa {displayLimit} kết quả lọc tiêu biểu nhất. Hãy thu hẹp tìm kiếm của bạn bằng cách nhập cụ thể từ khóa!
            </div>
          )}
        </div>

        {/* Box hướng dẫn lý giải cấp 5 học */}
        <div className="bg-purple-50 rounded-xl p-5 border border-purple-100 space-y-3">
          <h4 className="text-xs font-bold text-purple-900 uppercase tracking-widest flex items-center gap-1.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-600" /> BẬT MÍ: HƯỚNG DẪN QUY NẠP HỢP LỆ THEO BẬC LOGIC
          </h4>
          <div className="text-xs text-slate-700 leading-normal space-y-1.5">
            <p>
              Mẫu dữ liệu thực tế do Doanh nghiệp / Hộ kinh doanh nhập đôi khi sử dụng định dạng mã rút gọn (4 số thay vì 5 số).
              Bạn không cần lo lắng vì phần mềm của chúng tôi:
            </p>
            <ul className="list-disc list-inside space-y-1 pl-1 text-slate-650">
              <li>Tự động <strong className="text-amber-800">bẻ tách thông minh</strong> mã ngành có dấu "." hoặc ký tự thừa liên quan.</li>
              <li>Có sẵn công nghệ <strong className="text-purple-900">Quy nạp từ mã cấp con</strong> lên mã cha gần nhất (e.g., nhập <code className="font-mono bg-white border border-slate-200 px-1 py-0.5 rounded text-rose-600">56101</code> không có sẽ được khớp với cha của chúng là <code className="font-mono bg-white border border-slate-200 px-1 py-0.5 rounded text-emerald-700">56100</code> hoặc <code className="font-mono bg-white border border-slate-200 px-1 py-0.5 rounded text-sky-700">5610</code>).</li>
              <li>Bảo lưu an toàn 100% cột ngành của bạn, ghi chú trực tiếp trạng thái <span className="text-amber-700 font-bold">"Khớp quy nạp cấp học"</span> tại tệp kết toán tải về!</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
});

VsicCatalogExplorer.displayName = "VsicCatalogExplorer";

export default VsicCatalogExplorer;
