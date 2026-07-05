import React, { useState, useMemo, useEffect } from "react";
import { 
  CheckSquare, 
  AlertTriangle, 
  XCircle, 
  CheckCircle2, 
  Sliders, 
  HelpCircle, 
  ArrowRightLeft, 
  Download, 
  Search, 
  Filter, 
  Info, 
  RefreshCw,
  FileSpreadsheet
} from "lucide-react";
import * as XLSX from "xlsx";
import { vsicRawData, normalizeSectorCode, smartSuggestSectorByDescription } from "../data/vsic";

function normalizeTextToCompare(text: string): string {
  if (!text) return "";
  let clean = text.toString().toLowerCase().trim();
  clean = clean.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  clean = clean.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  clean = clean.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  clean = clean.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  clean = clean.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  clean = clean.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  clean = clean.replace(/đ/g, "d");
  clean = clean.replace(/[^a-z0-9\s]/g, " ");
  return clean.replace(/\s+/g, " ").trim();
}

interface DescriptorMatchScannerProps {
  mainData: any[];
  columns: string[];
  mapping?: {
    mota: string;
    manganh: string;
    xa: string;
    doanhthu: string;
    laodong: string;
    idCol: string;
  };
}

const DescriptorMatchScanner = React.memo(function DescriptorMatchScanner({ mainData, columns, mapping }: DescriptorMatchScannerProps) {
  const [colManganh, setColManganh] = useState<string>("");
  const [colMotaDtv, setColMotaDtv] = useState<string>("");

  const [uiError, setUiError] = useState<string | null>(null);
  const [uiSuccess, setUiSuccess] = useState<string | null>(null);

  const [searchTerm, setSearchTerm] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<string>("ALL");

  const [smartFuzzyMode, setSmartFuzzyMode] = useState<boolean>(true);
  const [fuzzyThreshold, setFuzzyThreshold] = useState<number>(35);

  const [scanResults, setScanResults] = useState<any[]>([]);
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [hasScanned, setHasScanned] = useState<boolean>(false);

  useEffect(() => {
    if (columns && columns.length > 0) {
      if (mapping && mapping.manganh && columns.includes(mapping.manganh)) {
        setColManganh(mapping.manganh);
      } else {
        setColManganh("");
      }

      if (mapping && mapping.mota && columns.includes(mapping.mota)) {
        setColMotaDtv(mapping.mota);
      } else {
        setColMotaDtv("");
      }
    }
  }, [columns, mapping]);

  const handleRunMatchScan = () => {
    setUiError(null);
    setUiSuccess(null);

    if (!colManganh) {
      setUiError("Vui lòng chọn 'Cột mã ngành' từ tệp trước khi bắt đầu so sánh.");
      return;
    }
    if (!colMotaDtv) {
      setUiError("Vui lòng chọn 'Cột mô tả/tên ngành thực tế của ĐTV' từ tệp trước khi bắt đầu so sánh.");
      return;
    }

    setIsScanning(true);
    setHasScanned(false);

    setTimeout(() => {
      try {
        const analyzed = mainData.map((row, index) => {
          if (!row || typeof row !== "object") {
            return {
              index: index + 1,
              originalRow: {},
              codeVal: "",
              descDtv: "",
              standardName: "Dòng trống",
              isMatch: false,
              compareResult: "Trống",
              status: "CRITICAL"
            };
          }

          const codeVal = String(row[colManganh] || "").trim();
          const descDtv = String(row[colMotaDtv] || "").trim();

          const normalizedCode = normalizeSectorCode(codeVal);
          const standardName = vsicRawData[normalizedCode] || "";

          let status: "SAFE" | "CRITICAL" = "SAFE";
          let compareResult = "";
          let isMatch = false;
          let suggestedCode = "";
          let suggestedName = "";
          let suggestedScore = 0;

          const dtvNorm = normalizeTextToCompare(descDtv);

          if (!standardName) {
            status = "CRITICAL";
            compareResult = `Không tìm thấy mã ngành '${normalizedCode}' trong danh mục đã nạp!`;
            isMatch = false;
          } else {
            const stdNorm = normalizeTextToCompare(standardName);

            if (dtvNorm === stdNorm) {
              status = "SAFE";
              compareResult = "Trùng khớp hoàn toàn ngữ nghĩa";
              isMatch = true;
            } else {
              if (smartFuzzyMode) {
                const dtvWords = dtvNorm.split(" ").filter(w => w.length >= 2);
                const stdWords = stdNorm.split(" ").filter(w => w.length >= 2);
                let similarity = 0;
                if (dtvWords.length > 0 && stdWords.length > 0) {
                  let overlapCount = 0;
                  const matchedSet = new Set<string>();
                  stdWords.forEach(w => {
                    if (dtvWords.includes(w) && !matchedSet.has(w)) {
                      overlapCount++;
                      matchedSet.add(w);
                    }
                  });
                  similarity = Math.round((2 * overlapCount) / (dtvWords.length + stdWords.length) * 100);
                }

                if (similarity >= fuzzyThreshold) {
                  status = "SAFE";
                  compareResult = `Khớp từ khóa thông minh (Tương đồng: ${similarity}%)`;
                  isMatch = true;
                } else {
                  status = "CRITICAL";
                  isMatch = false;
                  if (!dtvNorm) {
                    compareResult = "Tên ngành thực tế ĐTV mô tả bị để trống hoàn toàn";
                  } else {
                    compareResult = `Sai lệch chữ (Chỉ trùng: ${similarity}%). Khác biệt câu từ quá lớn.`;
                  }
                }
              } else {
                status = "CRITICAL";
                isMatch = false;
                if (!dtvNorm) {
                  compareResult = "Tên ngành thực tế ĐTV mô tả bị để trống hoàn toàn";
                } else {
                  compareResult = "Có sự khác khác biệt về mặt câu chữ giữa mô tả ĐTV và tên ngành chuẩn";
                }
              }
            }
          }

          // If not a perfect/fuzzy match or code doesn't exist, search the whole VSIC memory for a better matching code recommendation
          if (!isMatch && dtvNorm) {
            const suggestion = smartSuggestSectorByDescription(descDtv);
            if (suggestion) {
              suggestedCode = suggestion.ma;
              suggestedName = suggestion.ten;
              suggestedScore = Math.round(Math.min(100, Math.sqrt(suggestion.diem) * 100));
              
              if (!standardName) {
                compareResult = `Mã '${normalizedCode}' không tồn tại. Gợi ý mã đúng: [${suggestedCode}] - ${suggestedName} (Tin cậy: ${suggestedScore}%)`;
              } else {
                compareResult = `${compareResult} Gợi ý thay thế: [${suggestedCode}] - ${suggestedName} (Tin cậy: ${suggestedScore}%)`;
              }
            }
          }

          return {
            index: index + 1,
            originalRow: row,
            codeVal,
            descDtv,
            standardName: standardName || "(Không tìm thấy tên ngành chuẩn)",
            isMatch,
            compareResult,
            status,
            suggestedCode,
            suggestedName,
            suggestedScore
          };
        });

        setScanResults(analyzed);
        setHasScanned(true);
        setUiSuccess(`Đã so sánh đối chiếu khớp danh mục thành công ${analyzed.length} dòng dữ liệu!`);
      } catch (err: any) {
        setUiError("Lỗi trong quá trình so sánh: " + err.message);
      } finally {
        setIsScanning(false);
      }
    }, 100);
  };

  const stats = useMemo(() => {
    if (!hasScanned) return { total: 0, safe: 0, critical: 0 };
    const total = scanResults.length;
    const safe = scanResults.filter(r => r.status === "SAFE").length;
    const critical = scanResults.filter(r => r.status === "CRITICAL").length;
    return { total, safe, critical };
  }, [scanResults, hasScanned]);

  const filteredScanItems = useMemo(() => {
    if (!hasScanned) return [];
    
    return scanResults.filter(item => {
      if (statusFilter !== "ALL") {
        if (statusFilter === "SAFE" && item.status !== "SAFE") return false;
        if (statusFilter === "CRITICAL" && item.status !== "CRITICAL") return false;
      }
      if (searchTerm.trim()) {
        const query = searchTerm.toLowerCase();
        return (
          item.codeVal.toLowerCase().includes(query) ||
          item.descDtv.toLowerCase().includes(query) ||
          item.standardName.toLowerCase().includes(query) ||
          item.compareResult.toLowerCase().includes(query)
        );
      }
      return true;
    });
  }, [scanResults, hasScanned, statusFilter, searchTerm]);

  const displayLimit = 250;
  const displayedScanItems = filteredScanItems.slice(0, displayLimit);

  const handleExportMatchReport = () => {
    setUiError(null);
    setUiSuccess(null);

    if (scanResults.length === 0) {
      setUiError("Chưa có dữ liệu so sánh nào được tính toán thành công.");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      const sheetData = scanResults.map(item => {
        return {
          ...item.originalRow,
          "Mã ngành": item.codeVal,
          "Mô tả ngành của ĐTV": item.descDtv,
          "Tên ngành chuẩn VSIC": item.standardName,
          "Kết quả so sánh đối chiếu": item.isMatch ? "TRÙNG KHỚP" : "CÓ SAI LỆCH",
          "Chi tiết chênh lệch": item.compareResult,
          "Mã ngành gợi ý thay thế (AI)": item.suggestedCode || "",
          "Tên ngành gợi ý thay thế (AI)": item.suggestedName || "",
          "Độ tin cậy gợi ý (%)": item.suggestedScore || ""
        };
      });

      const ws = XLSX.utils.json_to_sheet(sheetData);
      XLSX.utils.book_append_sheet(wb, ws, "Doi_Chieu_Khop_Nganh");
      XLSX.writeFile(wb, "Bao_Cao_Doi_Chieu_Khop_Nganh.xlsx");
      
      setUiSuccess("Đã xuất tệp Excel báo cáo đối sánh 'Bao_Cao_Doi_Chieu_Khop_Nganh.xlsx' thành công!");
    } catch (e: any) {
      setUiError("Lỗi xuất Excel: " + e.message);
    }
  };

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-xl font-sans text-slate-800">
      
      {/* Tab Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-5">
        <div className="space-y-1">
          <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
            <ArrowRightLeft className="w-5 h-5 text-indigo-600" /> KHỚP NGÀNH & SO SÁNH ĐÁNH DẤU SAI LỆCH
          </h3>
          <p className="text-xs text-slate-500">
            Hệ thống đối chiếu mã ngành trong tệp với Danh mục bạn đã nạp, tự động đặt Tên ngành chuẩn cạnh Tên mô tả của ĐTV nhập để chỉ ra điểm sai khác.
          </p>
        </div>
      </div>

      {/* Configuration Card */}
      <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 space-y-4">
        <h4 className="text-xs font-bold text-slate-500 uppercase tracking-widest flex items-center gap-1.5 pb-2 border-b border-slate-200">
          <Sliders className="w-4 h-4 text-indigo-600" /> THIẾT LẬP CỘT ĐỐI CHIẾU
        </h4>

        {/* UI Notices */}
        {uiError && (
          <div className="p-3.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-800 text-xs flex items-center gap-2 font-medium">
            <XCircle className="w-4.5 h-4.5 shrink-0 text-rose-600" />
            <span>{uiError}</span>
          </div>
        )}
        {uiSuccess && (
          <div className="p-3.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs flex items-center gap-2 font-medium">
            <CheckCircle2 className="w-4.5 h-4.5 shrink-0 text-emerald-600" />
            <span>{uiSuccess}</span>
          </div>
        )}
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          
          {/* Code Selection */}
          <div className="space-y-1.5">
            <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-amber-500"></span>
              Chọn Cột chứa Mã Ngành cần tra cứu:
            </label>
            <select
              value={colManganh}
              onChange={(e) => setColManganh(e.target.value)}
              className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2.5 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
            >
              <option value="">-- Chọn cột --</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
            <p className="text-[10px] text-slate-500">Mã ngành trong file Excel sẽ dùng để so khớp trực tiếp với Danh mục chuẩn đã nạp.</p>
          </div>

          {/* DTV Description Selection */}
          <div className="space-y-1.5">
            <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              Chọn Cột chứa Mô tả/Tên ngành thực tế của ĐTV:
            </label>
            <select
              value={colMotaDtv}
              onChange={(e) => setColMotaDtv(e.target.value)}
              className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2.5 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
            >
              <option value="">-- Chọn cột --</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
            <p className="text-[10px] text-slate-500">Chữ mô tả thực tế của điều tra viên dùng để so sánh chênh lệch với tên ngành chuẩn.</p>
          </div>

        </div>

        {/* Smart Match Parameters Row */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4 border-t border-slate-200/80">
          
          {/* Smart Fuzzy Match Toggle */}
          <div className="flex items-center justify-between bg-white border border-slate-200 p-4 rounded-xl shadow-sm">
            <div className="space-y-0.5">
              <label className="text-xs text-slate-800 font-bold block">
                🧠 So khớp từ vựng tương tự thông minh
              </label>
              <span className="text-[10px] text-slate-500 block leading-relaxed max-w-[280px]">
                Tự động bỏ qua lỗi chính tả nhỏ, dấu câu, hoa thường để giảm tải lỗi ảo đối chiếu chéo.
              </span>
            </div>
            <label className="relative inline-flex items-center cursor-pointer select-none ml-4 shrink-0">
              <input 
                type="checkbox" 
                checked={smartFuzzyMode} 
                onChange={(e) => setSmartFuzzyMode(e.target.checked)} 
                className="sr-only peer" 
              />
              <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-indigo-600"></div>
            </label>
          </div>

          {/* Similarity Threshold Slider */}
          <div className="bg-white border border-slate-200 p-4 rounded-xl shadow-sm flex flex-col justify-between space-y-2">
            <div className="flex items-center justify-between">
              <label className="text-xs text-slate-800 font-bold">
                🎯 Ngưỡng tương đồng tối thiểu
              </label>
              <span className="font-mono text-[10px] font-extrabold text-indigo-700 bg-indigo-50 border border-indigo-150 rounded px-1.5 py-0.5 shrink-0">
                {fuzzyThreshold}%
              </span>
            </div>
            <input
              type="range"
              min="10"
              max="90"
              step="5"
              value={fuzzyThreshold}
              disabled={!smartFuzzyMode}
              onChange={(e) => setFuzzyThreshold(parseInt(e.target.value, 10))}
              className="w-full h-1.5 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-indigo-650 disabled:opacity-40"
            />
            <span className="text-[9px] text-slate-400 block leading-none">
              Duyệt là ĐẠT nếu tỷ lệ số lượng từ trùng khớp giữa ĐTV viết và tên VSIC &ge; {fuzzyThreshold}%.
            </span>
          </div>

        </div>

        {/* Action Controls */}
        <div className="flex flex-col sm:flex-row gap-3 pt-2">
          <button
            onClick={handleRunMatchScan}
            disabled={isScanning}
            className="flex-1 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-700 hover:to-indigo-800 text-white font-bold text-xs py-3 px-6 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
          >
            {isScanning ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin text-white" /> ĐANG ĐỐI CHIẾU MÃ NGÀNH VÀ DANH MỤC...
              </>
            ) : (
              <>
                <CheckSquare className="w-4 h-4 text-indigo-200" /> CHẠY KHỚP NGÀNH & SO SÁNH CHÊNH LỆCH
              </>
            )}
          </button>

          {hasScanned && (
            <button
              onClick={handleExportMatchReport}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3 px-6 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
            >
              <Download className="w-4 h-4" /> XUẤT BÁO CÁO ĐỐI CHIẾU (.XLSX)
            </button>
          )}
        </div>
      </div>

      {/* Analysis Metrics */}
      {hasScanned && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">Tổng số dòng</span>
              <p className="text-xl font-black text-slate-900 font-mono">{stats.total.toLocaleString()}</p>
            </div>
            <div className="bg-indigo-50 p-2.5 rounded-xl border border-indigo-100">
              <Info className="w-5 h-5 text-indigo-650" />
            </div>
          </div>

          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-emerald-700 font-bold uppercase tracking-wider">Trùng khớp hoàn toàn</span>
              <p className="text-xl font-black text-emerald-800 font-mono">
                {stats.safe.toLocaleString()} 
                <span className="text-xs text-slate-500 ml-1 font-normal">({stats.total > 0 ? Math.round(stats.safe / stats.total * 100) : 0}%)</span>
              </p>
            </div>
            <div className="bg-emerald-50 p-2.5 rounded-xl border border-emerald-100">
              <CheckCircle2 className="w-5 h-5 text-emerald-600" />
            </div>
          </div>

          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-rose-700 font-bold uppercase tracking-wider">Có Sự Sai Lệch / Chưa tìm thấy</span>
              <p className="text-xl font-black text-rose-800 font-mono">
                {stats.critical.toLocaleString()} 
                <span className="text-xs text-slate-500 ml-1 font-normal">({stats.total > 0 ? Math.round(stats.critical / stats.total * 100) : 0}%)</span>
              </p>
            </div>
            <div className="bg-rose-50 p-2.5 rounded-xl border border-rose-100">
              <XCircle className="w-5 h-5 text-rose-600" />
            </div>
          </div>

        </div>
      )}

      {/* Grid of parsed rows with filters */}
      {hasScanned && (
        <div className="space-y-4">
          
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            
            {/* Status Tabs */}
            <div className="flex flex-wrap items-center gap-1.5 bg-slate-100 p-1.5 rounded-xl border border-slate-200">
              <button
                onClick={() => setStatusFilter("ALL")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  statusFilter === "ALL" ? "bg-indigo-650 text-white shadow-sm" : "text-slate-600 hover:text-slate-800"
                }`}
              >
                Tất cả ({stats.total})
              </button>
              <button
                onClick={() => setStatusFilter("SAFE")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                  statusFilter === "SAFE" ? "bg-emerald-600 text-white shadow-sm" : "text-slate-600 hover:text-emerald-700"
                }`}
              >
                ✓ Trùng khớp ({stats.safe})
              </button>
              <button
                onClick={() => setStatusFilter("CRITICAL")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                  statusFilter === "CRITICAL" ? "bg-rose-600 text-white shadow-sm" : "text-slate-600 hover:text-rose-700"
                }`}
              >
                ✗ Sai lệch / Không tìm thấy ({stats.critical})
              </button>
            </div>

            {/* Quick Filter Search */}
            <div className="relative w-full md:w-[280px]">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                <Search className="w-4 h-4" />
              </span>
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Tìm mã hoặc tên ngành..."
                className="w-full bg-slate-50 border border-slate-300 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl pl-9 pr-4 py-2 text-xs text-slate-800 placeholder-slate-450 outline-none font-medium"
              />
            </div>

          </div>

          {/* Results Table */}
          <div className="border border-slate-200 rounded-2xl overflow-hidden bg-white shadow-sm">
            <div className="max-h-[550px] overflow-y-auto w-full">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-250 sticky top-0 z-10 text-slate-700 font-mono text-[10px] tracking-wider uppercase font-bold">
                    <th className="p-3.5 text-center w-[75px]">STT</th>
                    <th className="p-3.5 w-[110px] text-center">MÃ NGÀNH</th>
                    <th className="p-3.5 min-w-[220px]">TÊN MÔ TẢ NGÀNH THỰC TẾ (ĐTV)</th>
                    <th className="p-3.5 min-w-[220px]">TÊN NGÀNH CHUẨN TRONG DANH MỤC</th>
                    <th className="p-3.5 text-center w-[120px]">TRẠNG THÁI</th>
                    <th className="p-3.5 min-w-[200px]">CHI TIẾT ĐỐI CHIẾU</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs text-slate-800">
                  {displayedScanItems.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="p-10 text-center text-slate-400 italic">
                        Không tìm thấy bất kỳ dòng đối sánh hay sai lỗi nào khớp với bộ lọc hiện thời.
                      </td>
                    </tr>
                  ) : (
                    displayedScanItems.map((item) => {
                      const trBg = item.status === "SAFE" ? "" : "bg-rose-50/40 hover:bg-rose-100/30";
                      return (
                        <tr key={item.index} className={`${trBg} transition-colors hover:bg-slate-50 border-b border-slate-100`}>
                          <td className="p-3.5 text-center font-mono font-bold text-slate-400">
                            {item.index}
                          </td>
                          <td className="p-3.5 text-center font-mono font-bold text-amber-850">
                            {item.codeVal}
                          </td>
                          <td className="p-3.5 font-semibold text-slate-800">
                            {item.descDtv}
                          </td>
                          <td className="p-3.5 font-semibold text-emerald-800">
                            {item.standardName}
                          </td>
                          <td className="p-3.5 text-center">
                            {item.isMatch ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-lg text-[9px] font-bold">
                                ✓ TRÙNG KHỚP
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-rose-50 text-rose-700 border border-rose-200 rounded-lg text-[9px] font-bold">
                                ✗ SAI LỆCH
                              </span>
                            )}
                          </td>
                          <td className="p-3.5 text-slate-500 text-[11px] font-medium">
                            {item.compareResult}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
            
            {filteredScanItems.length > displayLimit && (
              <div className="p-4 bg-slate-50 border-t border-slate-200 text-center text-xs text-slate-400 italic font-semibold">
                Đã giới hạn hiển thị tối đa {displayLimit} trên trình duyệt để tránh lag. Bản Excel sẽ xuất đầy đủ {filteredScanItems.length} dòng dữ liệu đối kết.
              </div>
            )}
          </div>

        </div>
      )}

      {/* Guide details when not scanned */}
      {!hasScanned && (
        <div className="bg-indigo-50/50 rounded-xl p-6 border border-indigo-100 space-y-4">
          <div className="flex items-center gap-2">
            <HelpCircle className="w-5 h-5 text-indigo-650" />
            <h4 className="text-sm font-bold text-indigo-900 uppercase tracking-wider">Nguyên lý hoạt động đối sánh & Chênh lệch tên ngành:</h4>
          </div>
          
          <div className="space-y-4 text-xs text-slate-700 leading-relaxed">
            <p className="text-slate-650">
              Quy trình đối sánh hoạt động trực tiếp dựa trên tệp Danh mục Chuẩn được bạn cấu hình/nạp tại tab <strong>Tra Cứu / Nạp Danh Mục</strong>:
            </p>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-white p-4 rounded-xl border border-indigo-100/50 shadow-sm">
              <div>
                <span className="text-emerald-800 font-bold">1. Lục tìm mã trùng khớp:</span>
                <p className="mt-1 text-slate-500">
                  Hệ thống lấy mã ngành trong mỗi hàng, gỡ các ký hiệu phân cách rỗng, tra cứu chính xác mã số đó trong Catalog. Sau đó lấy tên chuẩn bê đặt bên cạnh tên mô tả khai báo thực tế của ĐTV nhập để chỉ ra điểm sai khác.
                </p>
              </div>
              
              <div>
                <span className="text-amber-800 font-bold">2. Nhận diện chênh lệch chữ:</span>
                <p className="mt-1 text-slate-500">
                  Phần mềm so sánh nguyên mẫu chuỗi ký tự không dấu (đã loại bỏ khoảng trắng rỗng và chữ in hoa thường) giữa 2 cột. Nếu có sự khác biệt về nội dung hoặc mục đích đăng ký, phần mềm sẽ đánh dấu <span className="text-rose-700 font-bold">SAI LỆCH</span> ngay lập tức để bôi đỏ.
                </p>
              </div>
            </div>

            <p className="font-mono text-[10px] text-indigo-700 font-bold italic">
              * Không sử dụng AI hay các giả thuyết phỏng đoán ngẫu nhiên. Mọi kết quả khớp đều đối chiếu cơ sở dữ liệu gốc do bạn trực tiếp kiểm soát!
            </p>
          </div>
        </div>
      )}

    </div>
  );
});

DescriptorMatchScanner.displayName = "DescriptorMatchScanner";

export default DescriptorMatchScanner;
