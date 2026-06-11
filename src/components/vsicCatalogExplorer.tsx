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

export default function VsicCatalogExplorer() {
  const [search, setSearch] = useState("");
  const [level, setLevel] = useState("ALL");
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  // Trạng thái file tải lên
  const [isDragging, setIsDragging] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{
    success?: boolean;
    message?: string;
    addedCount?: number;
  }>({});

  // Chế độ thuần khiết (Chỉ dùng danh mục nạp vào, xóa sạch danh mục mặc định của app)
  const [pureMode, setPureMode] = useState<boolean>(() => {
    try {
      const stored = localStorage.getItem("custom_vsic_is_pure");
      return stored !== "false";
    } catch {
      return true;
    }
  });

  // Danh mục bổ sung của người dùng lưu trong localStorage
  const [customCatalog, setCustomCatalog] = useState<{ [key: string]: string }>(() => {
    try {
      const stored = localStorage.getItem("custom_vsic_data");
      return stored ? JSON.parse(stored) : {};
    } catch {
      return {};
    }
  });

  // Trạng thái cập nhật, chỉnh sửa thủ công trực tiếp tên ngành
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

  // Hồi phục và hợp nhất dữ liệu gốc tùy biến theo chế độ pureMode
  useEffect(() => {
    if (pureMode) {
      clearAllSectorsInVSIC();
    }
    Object.assign(vsicRawData, customCatalog);
  }, [pureMode, customCatalog]);

  // Hợp nhất dữ liệu gốc và dữ liệu tùy biến người dùng nạp vào để xem/tìm kiếm
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

  const quickSearches = ["56101", "56302", "47521", "68104", "Bán lẻ", "Nhà hàng", "Cà phê", "Sản xuất"];

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

  // 1. CHỨC NĂNG TẢI FILE MẪU BAN ĐẦU (Theo đúng cấu trúc 5 cấp người dùng yêu cầu)
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
      let csvContent = "data:text/csv;charset=utf-8,\uFEFF"; // Thêm BOM
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
      // XLSX Format
      const wsData = [headers, ...rows];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      XLSX.utils.book_append_sheet(wb, ws, "Danh_Muc_Mau_5_Cap");
      XLSX.writeFile(wb, "Mau_Danh_Muc_Nganh_5_Cap.xlsx");
    }
  };

  // 2. CHỨC NĂNG PARSE FILE EXCEL/CSV (Hỗ trợ cấu trúc 5 cấp chi tiết + 2 cột thông thường)
  const handleProcessFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        // Đọc dưới dạng mảng 2 chiều
        const rawRows = XLSX.utils.sheet_to_json<any>(sheet, { header: 1 });
        if (rawRows.length < 2) {
          setUploadStatus({
            success: false,
            message: "Tệp tải lên trống hoặc không có dòng tiêu đề (Dòng 1)."
          });
          return;
        }

        // Tìm chỉ số các cột trong hàng tiêu đề
        const headerRow = rawRows[0].map((h: any) => String(h || "").trim().toLowerCase());
        
        // Dò tìm cho cấu trúc 5 cấp
        let cap1Idx = -1;
        let cap2Idx = -1;
        let cap3Idx = -1;
        let cap4Idx = -1;
        let cap5Idx = -1;
        let tenNganhIdx = -1;

        // Dò tìm cho trường hợp tệp 2 cột đơn giản
        let codeIdxSimple = -1;
        let nameIdxSimple = -1;

        for (let i = 0; i < headerRow.length; i++) {
          const h = headerRow[i];
          // 5 cấp
          if (h.includes("cap1") || h.includes("cấp 1") || h.includes("cap 1") || h.includes("nganhcap1")) cap1Idx = i;
          else if (h.includes("cap2") || h.includes("cấp 2") || h.includes("cap 2") || h.includes("nganhcap2")) cap2Idx = i;
          else if (h.includes("cap3") || h.includes("cấp 3") || h.includes("cap 3") || h.includes("nganhcap3")) cap3Idx = i;
          else if (h.includes("cap4") || h.includes("cấp 4") || h.includes("cap 4") || h.includes("nganhcap4")) cap4Idx = i;
          else if (h.includes("cap5") || h.includes("cấp 5") || h.includes("cap 5") || h.includes("nganhcap5")) cap5Idx = i;
          else if (h.includes("tên ngành") || h.includes("tennganh") || h.includes("tên gọi") || h.includes("tên") || h.includes("ten") || h.includes("name") || h.includes("mô tả") || h.includes("mota")) {
            tenNganhIdx = i;
          }

          // 2 cột thông thường fallback
          if (h.includes("mã") || h.includes("ma") || h.includes("code") || h.includes("sector") || h.includes("id")) {
            if (codeIdxSimple === -1) codeIdxSimple = i;
          }
        }

        const is5LevelStructure = (cap1Idx !== -1 || cap2Idx !== -1 || cap3Idx !== -1 || cap4Idx !== -1 || cap5Idx !== -1) && tenNganhIdx !== -1;

        const newCustomDict: { [key: string]: string } = {};
        const newParents: { [key: string]: string } = {};
        let added = 0;

        if (is5LevelStructure) {
          // Bóc tách theo logic 5 cấp
          for (let r = 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || row.length === 0) continue;

            const val1 = cap1Idx !== -1 && row[cap1Idx] !== undefined && row[cap1Idx] !== null ? String(row[cap1Idx]).trim() : "";
            const val2 = cap2Idx !== -1 && row[cap2Idx] !== undefined && row[cap2Idx] !== null ? String(row[cap2Idx]).trim() : "";
            const val3 = cap3Idx !== -1 && row[cap3Idx] !== undefined && row[cap3Idx] !== null ? String(row[cap3Idx]).trim() : "";
            const val4 = cap4Idx !== -1 && row[cap4Idx] !== undefined && row[cap4Idx] !== null ? String(row[cap4Idx]).trim() : "";
            const val5 = cap5Idx !== -1 && row[cap5Idx] !== undefined && row[cap5Idx] !== null ? String(row[cap5Idx]).trim() : "";
            const nameVal = tenNganhIdx !== -1 && row[tenNganhIdx] !== undefined && row[tenNganhIdx] !== null ? String(row[tenNganhIdx]).trim() : "";

            if (!nameVal) continue;

            // Normalize các mã cấp
            let c1 = val1.trim().toUpperCase();
            let c2 = val2.replace(/\D/g, "");
            let c3 = val3.replace(/\D/g, "");
            let c4 = val4.replace(/\D/g, "");
            let c5 = val5.replace(/\D/g, "");

            // Khôi phục số 0 đầu bị Excel cắt đi
            if (c2) c2 = c2.padStart(2, "0");
            if (c3) c3 = c3.padStart(3, "0");
            if (c4) c4 = c4.padStart(4, "0");
            if (c5) c5 = c5.padStart(5, "0");

            // Lưu sơ đồ ánh xạ về mã Cấp 1 (c1)
            if (c1 && /^[A-U]$/i.test(c1)) {
              if (c2) newParents[c2] = c1;
              if (c3) newParents[c3] = c1;
              if (c4) newParents[c4] = c1;
              if (c5) newParents[c5] = c1;
            }

            // Lựa chọn mã chuẩn từ sâu nhất đến khái quát nhất (Cấp 5 -> Cấp 1) làm đại diện cho dòng này để lấy tên ngành
            let cleanCode = "";
            if (c5) cleanCode = c5;
            else if (c4) cleanCode = c4;
            else if (c3) cleanCode = c3;
            else if (c2) cleanCode = c2;
            else if (c1) cleanCode = c1;

            if (cleanCode) {
              newCustomDict[cleanCode] = nameVal;
              added++;
            }
          }
        } else {
          // Tệp 2 cột đơn giản fallback
          if (codeIdxSimple === -1) codeIdxSimple = 0;
          if (nameIdxSimple === -1) nameIdxSimple = headerRow.length > 1 ? 1 : 0;
          if (codeIdxSimple === nameIdxSimple && rawRows[0].length > 1) nameIdxSimple = 1;

          for (let r = 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || row.length === 0) continue;

            let rawCode = row[codeIdxSimple];
            let rawName = row[nameIdxSimple];

            if (rawCode === undefined || rawCode === null) continue;

            let cleanCode = String(rawCode).trim().replace(/\D/g, "");
            if (!cleanCode) {
              const letterCode = String(rawCode).trim().toUpperCase();
              if (/^[A-U]$/.test(letterCode)) {
                cleanCode = letterCode;
              }
            }

            // Tự động suy luận cấp để khôi phục số 0 đầu bị mất trong định dạng 2 cột đơn giản
            if (cleanCode && /^\d+$/.test(cleanCode)) {
              const valNum = parseInt(cleanCode, 10);
              if (cleanCode.length === 1) {
                // ví dụ "1" -> cấp 2 "01"
                cleanCode = cleanCode.padStart(2, "0");
              } else if (cleanCode.length === 2 && valNum < 10) {
                // ví dụ "11" -> cấp 3 "011"
                cleanCode = cleanCode.padStart(3, "0");
              } else if (cleanCode.length === 3 && valNum < 100) {
                cleanCode = cleanCode.padStart(4, "0");
              } else if (cleanCode.length === 4 && valNum < 1000) {
                cleanCode = cleanCode.padStart(5, "0");
              }
            }

            let cleanName = rawName ? String(rawName).trim() : "";
            if (cleanCode && cleanName) {
              newCustomDict[cleanCode] = cleanName;
              added++;
            }
          }
        }

        if (added === 0) {
          setUploadStatus({
            success: false,
            message: "Hệ thống không tìm thấy hàng dữ liệu ngành hợp lệ nào từ file của bạn. Hãy đảm bảo file đúng mẫu cấu trúc 5 cấp hoặc 2 cột."
          });
          return;
        }

        // Lưu trữ
        let merged = {};
        let mergedParents = {};

        let existingParents = {};
        try {
          const stored = localStorage.getItem("custom_vsic_parents");
          if (stored) existingParents = JSON.parse(stored);
        } catch (e) {}

        if (pureMode) {
          // Xóa hết cũ, chỉ sử dụng tệp nạp vào làm danh mục
          merged = newCustomDict;
          mergedParents = newParents;
          clearAllSectorsInVSIC();
          clearAllParentsInVSIC();
        } else {
          // Sáp nhập đồng thời
          merged = { ...customCatalog, ...newCustomDict };
          mergedParents = { ...existingParents, ...newParents };
        }

        setCustomCatalog(merged);
        localStorage.setItem("custom_vsic_data", JSON.stringify(merged));
        localStorage.setItem("custom_vsic_parents", JSON.stringify(mergedParents));
        
        Object.assign(vsicRawData, merged);
        Object.assign(vsicParentMap, mergedParents);

        setUploadStatus({
          success: true,
          message: `Nạp thành công ${added} danh mục cấp nhóm thành công! Hệ thống đã tự động liên kết dữ liệu theo phân cấp.`,
          addedCount: added
        });

        if (fileInputRef.current) fileInputRef.current.value = "";
      } catch (err: any) {
        setUploadStatus({
          success: false,
          message: "Lỗi phân tích file Excel/CSV: " + err.message
        });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleProcessFile(e.target.files[0]);
    }
  };

  // 3. KÉO THẢ TỆP TIN
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleProcessFile(e.dataTransfer.files[0]);
    }
  };

  // 4. XÓA DANH MỤC TRONG BỘ NHỚ ĐỂ NẠP LẠI
  const handleResetCatalog = () => {
    if (window.confirm("Bạn có chắc chắn muốn xóa sạch toàn bộ danh mục mã ngành hiện tại trong bộ nhớ để nạp lại tệp Excel mới cho chuẩn và đồng nhất không?")) {
      localStorage.removeItem("custom_vsic_data");
      localStorage.removeItem("custom_vsic_parents");
      localStorage.setItem("custom_vsic_is_pure", "true");
      setCustomCatalog({});
      setPureMode(true);
      clearAllSectorsInVSIC();
      clearAllParentsInVSIC();
      alert("Đã xóa sạch danh mục cũ trong bộ nhớ! Hệ thống đang ở chế độ chờ, xin mời bạn tải lên tệp Excel danh mục của bạn.");
      window.location.reload();
    }
  };

  // Xử lý thay đổi chế độ PureMode
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
    <div className="space-y-6">
      {/* 1. KHU VỰC TẢI LÊN DANH MỤC TRỰC QUAN */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* Nạp tệp mới */}
        <div className="lg:col-span-2 bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4 shadow-xl">
          <div className="flex items-center gap-3">
            <div className="bg-purple-500/10 p-2 rounded-lg">
              <Upload className="w-5 h-5 text-purple-400" />
            </div>
            <div>
              <h4 className="text-sm font-bold text-white uppercase tracking-wider">Nạp bổ sung danh mục ngành của bạn (Excel / CSV)</h4>
              <p className="text-xs text-gray-400">Tải lên danh mục riêng để ghép mã thắt nút, giải phóng hoàn toàn lỗi định dạng danh mục mặc định của bạn.</p>
            </div>
          </div>

          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all ${
              isDragging 
                ? "border-purple-400 bg-purple-950/10" 
                : "border-gray-700 hover:border-gray-600 bg-[#111827]/40 hover:bg-[#111827]/65"
            }`}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".csv,.xlsx,.xls"
              className="hidden"
            />
            <FileSpreadsheet className="w-10 h-10 text-gray-400 mx-auto mb-3" />
            <p className="text-sm text-gray-200 font-semibold">Tải lên tệp CSV hoặc Excel chứa danh mục mở rộng</p>
            <p className="text-xs text-gray-500 mt-1">Kéo thả file vào đây hoặc click để duyệt từ máy tính</p>
            <p className="text-[10px] text-purple-400/80 font-mono mt-2 italic">*Hệ thống tự đối chiếu thông minh cột [NganhCap1] đến [NganhCap5] thành Mã chuẩn</p>
          </div>

          {/* Hiển thị kết quả nạp */}
          {uploadStatus.message && (
            <div className={`p-4 rounded-xl text-xs border flex items-start gap-2.5 ${
              uploadStatus.success 
                ? "bg-emerald-950/20 border-emerald-500/20 text-emerald-400" 
                : "bg-rose-950/20 border-rose-500/20 text-rose-400"
            }`}>
              {uploadStatus.success ? <CheckCircle2 className="w-4 h-4 shrink-0" /> : <AlertTriangle className="w-4 h-4 shrink-0" />}
              <div>
                <p className="font-semibold">{uploadStatus.success ? "NẠP THÀNH CÔNG" : "THẤT BẠI KHI NẠP"}</p>
                <p className="mt-0.5 leading-relaxed">{uploadStatus.message}</p>
              </div>
            </div>
          )}
        </div>

        {/* Thư viện quản lý bộ mẫu chuẩn */}
        <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 flex flex-col justify-between shadow-xl space-y-4">
          <div className="space-y-4">
            <h4 className="text-sm font-bold text-white uppercase tracking-wider flex items-center gap-2">
              <Download className="w-4 h-4 text-amber-400" /> TẢI FILE MẪU BAN ĐẦU
            </h4>
            <p className="text-xs text-gray-400 leading-normal">
              Bạn chưa rõ cấu trúc điền file thế nào? Hãy tải về file mẫu của chúng tôi, mở trên Excel, copy danh sách mã của bạn dán vào rồi nạp tệp 5 cấp!
            </p>
            <div className="grid grid-cols-2 gap-3 pt-2">
              <button 
                onClick={() => handleDownloadTemplate("xlsx")}
                className="flex items-center justify-center gap-1.5 bg-emerald-600/10 hover:bg-emerald-600/20 border border-emerald-500/20 text-emerald-400 py-2.5 px-3 rounded-lg text-xs font-semibold cursor-pointer transition-all"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" /> File mẫu Excel
              </button>
              <button 
                onClick={() => handleDownloadTemplate("csv")}
                className="flex items-center justify-center gap-1.5 bg-cyan-600/10 hover:bg-cyan-600/20 border border-cyan-500/20 text-cyan-400 py-2.5 px-3 rounded-lg text-xs font-semibold cursor-pointer transition-all"
              >
                <Download className="w-3.5 h-3.5" /> File mẫu CSV
              </button>
            </div>
          </div>

          <div className="border-t border-gray-800 pt-4 space-y-3">
            <div className="flex items-center justify-between bg-purple-950/20 border border-purple-500/20 rounded-xl p-3 mb-2">
              <div>
                <p className="text-xs font-bold text-purple-300">Chế độ Danh mục Riêng</p>
                <p className="text-[10px] text-gray-400 mt-0.5 leading-tight">Mở tính năng xóa danh mục hệ thống mặc định; chỉ xử lý duy nhất danh mục 5 cấp quý khách nạp.</p>
              </div>
              <label className="relative inline-flex items-center cursor-pointer select-none">
                <input 
                  type="checkbox" 
                  checked={pureMode} 
                  onChange={handlePureModeToggle} 
                  className="sr-only peer" 
                />
                <div className="w-9 h-5 bg-gray-700 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-purple-500"></div>
              </label>
            </div>

            <div className="flex items-center justify-between text-xs text-gray-400">
              <span>Đã nạp bổ sung:</span>
              <strong className="text-purple-400 font-mono text-sm">{customKeys.length} mã ngành</strong>
            </div>
            
            {(customKeys.length > 0 || pureMode) && (
              <button
                onClick={handleResetCatalog}
                className="w-full flex items-center justify-center gap-1.5 bg-red-950/25 hover:bg-red-950/40 border border-red-500/20 text-red-400 py-2 rounded-lg text-xs font-medium cursor-pointer transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" /> Xóa sạch bộ nhớ &amp; Nạp tệp mới
              </button>
            )}
          </div>
        </div>

      </div>

      {/* 2. CHƯC NĂNG TRA CỨU DANH MỤC TRỰC QUAN */}
      <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6 shadow-xl">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-gray-800 pb-5">
          <div className="space-y-1">
            <h3 className="text-lg font-bold text-white flex items-center gap-2">
              <Database className="w-5 h-5 text-amber-500" /> DANH MỤC HỆ THỐNG PHÂN CẤP HOẠT ĐỘNG (ĐÃ HỢP NHẤT)
            </h3>
            <p className="text-xs text-gray-400">
              Dữ liệu tích hợp gồm tất cả các mã ngành chuẩn quốc gia và mã ngành tùy chọn bạn đã nạp vào hệ thống để tra cứu, tìm kiếm.
            </p>
          </div>
          <div className="bg-amber-500/10 border border-amber-500/20 text-amber-400 font-mono text-xs px-3 py-1.5 rounded-lg flex items-center gap-1.5 self-start md:self-center">
            <span className="relative flex h-2 w-2">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-amber-500"></span>
            </span>
            Sẵn sàng kết nối: {vsicList.length} mã chuẩn
          </div>
        </div>

        {/* Tra cứu nhanh */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          {/* Tìm kiếm */}
          <div className="md:col-span-2 relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-gray-400">
              <Search className="w-4 h-4" />
            </span>
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Nhập mã ngành (ví dụ: 5610) hoặc từ khóa (bán lẻ, cà phê...)"
              className="w-full bg-[#111827] border border-[#374151] focus:border-purple-500 focus:ring-1 focus:ring-purple-500 rounded-xl pl-9 pr-4 py-2.5 text-sm text-white placeholder-gray-500 outline-none transition-all"
              id="vsic-catalog-search-input"
            />
          </div>

          {/* Lọc cấp độ */}
          <div className="relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-gray-400">
              <Filter className="w-4 h-4" />
            </span>
            <select
              value={level}
              onChange={(e) => setLevel(e.target.value)}
              className="w-full bg-[#111827] border border-[#374151] focus:border-purple-500 focus:ring-1 focus:ring-purple-500 rounded-xl pl-9 pr-4 py-2.5 text-sm text-white outline-none transition-all cursor-pointer"
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

          {/* Tổng quan kết quả lọc */}
          <div className="bg-[#111827] border border-gray-800 rounded-xl px-4 py-2 flex items-center justify-between">
            <div className="text-gray-400 text-xs">Phù hợp:</div>
            <div className="text-right">
              <span className="text-sm font-bold text-emerald-400">{filteredList.length}</span>
              <span className="text-gray-500 text-[10px] ml-1">nhóm</span>
            </div>
          </div>
        </div>

        {/* Gợi ý từ khóa tìm kiếm nhanh */}
        <div className="flex flex-wrap items-center gap-2 bg-[#111827]/40 p-3 rounded-xl border border-gray-850">
          <span className="text-xs text-gray-400 font-medium select-none">Từ khóa tìm nhanh:</span>
          {quickSearches.map((qs) => (
            <button
              key={qs}
              onClick={() => setSearch(qs)}
              className="bg-[#1f2937] hover:bg-purple-600/15 hover:text-purple-400 border border-gray-800 hover:border-purple-500/20 px-2.5 py-1 rounded-lg text-xs font-mono text-gray-300 transition-all cursor-pointer"
              id={`quick-search-${qs}`}
            >
              {qs}
            </button>
          ))}
          {search && (
            <button
              onClick={() => setSearch("")}
              className="text-xs text-red-400 hover:text-red-300 ml-auto font-semibold px-2"
              id="clear-search-btn"
            >
              Xóa tìm kiếm
            </button>
          )}
        </div>

        {/* Danh sách bảng kết quả cứu */}
        <div className="border border-gray-800 rounded-xl overflow-hidden bg-[#111827]/60">
          <div className="max-h-[500px] overflow-y-auto">
            <table className="w-full text-left border-collapse" id="vsic-catalog-table">
              <thead>
                <tr className="bg-[#111827] border-b border-gray-800 sticky top-0 z-10 text-gray-400 font-mono text-xs">
                  <th className="p-3.5 text-center w-[90px]">MÃ NGÀNH</th>
                  <th className="p-3.5 w-[110px] text-center">BẬC PHÂN CẤP</th>
                  <th className="p-3.5 text-left">MÔ TẢ TÊN GỌI CHUẨN VSIC QUỐC GIA</th>
                  <th className="p-3.5 text-center w-[120px]">NGUỒN MÃ</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-850/50">
                {displayedItems.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="p-8 text-center text-sm text-gray-500 italic">
                      Không tìm thấy mã VSIC hoạt động nào khớp với bộ lọc và điều kiện tìm kiếm của bạn. Hãy thử nhập lại!
                    </td>
                  </tr>
                ) : (
                  displayedItems.map((item, idx) => {
                    let capBadgeColor = "";
                    switch (item.cap) {
                      case 1:
                        capBadgeColor = "bg-rose-950/40 text-rose-400 border border-rose-500/10";
                        break;
                      case 2:
                        capBadgeColor = "bg-orange-950/40 text-orange-400 border border-orange-500/10";
                        break;
                      case 3:
                        capBadgeColor = "bg-amber-950/40 text-amber-400 border border-amber-500/10";
                        break;
                      case 4:
                        capBadgeColor = "bg-teal-950/40 text-teal-400 border border-teal-500/10";
                        break;
                      default:
                        capBadgeColor = "bg-sky-950/40 text-sky-400 border border-sky-500/10";
                    }

                    const isCustom = !!customCatalog[item.code];

                    return (
                      <tr
                        key={item.code}
                        className="hover:bg-[#1d2636]/40 transition-colors group cursor-pointer"
                        onClick={() => setSearch(item.code)}
                      >
                        <td className="p-3.5 text-center font-mono font-bold text-amber-400 text-sm group-hover:text-purple-400 transition-colors">
                          {item.code}
                        </td>
                        <td className="p-3.5 text-center">
                          <span className={`inline-block py-1 px-2.5 rounded-lg text-[10px] font-bold tracking-wider ${capBadgeColor}`}>
                            CẤP {item.cap}
                          </span>
                        </td>
                        <td className="p-3.5 text-sm text-gray-200 font-sans group-hover:text-white transition-colors leading-relaxed">
                          {editingCode === item.code ? (
                            <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                              <input
                                type="text"
                                value={editingName}
                                onChange={(e) => setEditingName(e.target.value)}
                                className="flex-1 bg-[#111827] border border-purple-500 rounded-lg px-3 py-1.5 text-sm text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                                autoFocus
                                onKeyDown={(e) => {
                                  if (e.key === "Enter") handleSaveEdit(item.code);
                                  else if (e.key === "Escape") setEditingCode(null);
                                }}
                              />
                              <button
                                onClick={() => handleSaveEdit(item.code)}
                                className="p-1.5 bg-emerald-600/20 hover:bg-emerald-600/45 text-emerald-400 border border-emerald-500/20 rounded-lg cursor-pointer transition-all flex items-center justify-center"
                                title="Lưu"
                              >
                                <Check className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setEditingCode(null)}
                                className="p-1.5 bg-gray-800 hover:bg-gray-750 text-gray-400 border border-gray-700 rounded-lg cursor-pointer transition-all flex items-center justify-center"
                                title="Hủy"
                              >
                                <span className="text-xs font-bold px-1 select-none">Hủy</span>
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-between gap-4">
                              <span>{item.name}</span>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setEditingCode(item.code);
                                  setEditingName(item.name);
                                }}
                                className="opacity-0 group-hover:opacity-100 hover:text-purple-400 p-1.5 rounded bg-gray-800/50 hover:bg-gray-800 border border-transparent hover:border-gray-700 transition-all flex items-center justify-center cursor-pointer"
                                title="Sửa tên ngành"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="lucide lucide-pencil"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
                              </button>
                            </div>
                          )}
                        </td>
                        <td className="p-3.5 text-center">
                          {isCustom ? (
                            <span className="inline-block py-0.5 px-2 bg-purple-500/15 text-purple-400 border border-purple-500/25 rounded-md text-[10px] font-bold">
                              Do Tự Nạp
                            </span>
                          ) : (
                            <span className="inline-block py-0.5 px-2 bg-gray-800 text-gray-400 rounded-md text-[10px]">
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
            <div className="p-3.5 bg-[#111827] border-t border-gray-800 text-center text-xs text-gray-500 italic">
              Hiển thị tối đa {displayLimit} kết quả lọc tiêu biểu nhất. Hãy thu hẹp tìm kiếm của bạn bằng cách nhập cụ thể từ khóa!
            </div>
          )}
        </div>

        {/* Box hướng dẫn lý giải cấp 5 học */}
        <div className="bg-[#111827] rounded-xl p-5 border border-purple-500/15 space-y-3">
          <h4 className="text-xs font-bold text-purple-400 uppercase tracking-widest flex items-center gap-1.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-400" /> BẬT MÍ: HƯỚNG DẪN QUY NẠP HỢP LỆ THEO BẬC LOGIC
          </h4>
          <div className="text-xs text-gray-300 leading-normal space-y-1.5">
            <p>
              Mẫu dữ liệu thực tế do Doanh nghiệp / Hộ kinh doanh nhập đôi khi sử dụng định dạng mã rút gọn (4 số thay vì 5 số).
              Bạn không cần lo lắng vì phần mềm của chúng tôi:
            </p>
            <ul className="list-disc list-inside space-y-1 pl-1 text-gray-400">
              <li>Tự động <strong className="text-yellow-400">bẻ tách thông minh</strong> mã ngành có dấu "." hoặc ký tự thừa liên quan.</li>
              <li>Có sẵn công nghệ <strong className="text-purple-400">Quy nạp từ mã cấp con</strong> lên mã cha gần nhất (e.g., nhập <code className="font-mono bg-gray-900 px-1 py-0.5 rounded text-red-400">56101</code> không có sẽ được khớp với cha của chúng là <code className="font-mono bg-gray-900 px-1 py-0.5 rounded text-emerald-400">56100</code> hoặc <code className="font-mono bg-gray-900 px-1 py-0.5 rounded text-cyan-400">5610</code>).</li>
              <li>Bảo lưu an toàn 100% cột ngành của bạn, ghi chú trực tiếp trạng thái <span className="text-amber-400">"Khớp quy nạp cấp học"</span> tại tệp kết toán tải về!</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
