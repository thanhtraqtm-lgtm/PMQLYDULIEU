import React, { useState, useMemo, useRef } from "react";
import {
  GitMerge,
  Upload,
  ArrowRight,
  CheckCircle2,
  AlertCircle,
  Sparkles,
  Sliders,
  Check,
  X,
  Search,
  FileSpreadsheet,
  Download,
  RefreshCw,
  HelpCircle,
  ChevronRight,
  Layers,
  TableProperties,
  ChevronLeft,
  Filter,
  CheckSquare,
  Square
} from "lucide-react";
import * as XLSX from "xlsx";
import { parse2DArrayWithSmartHeader } from "../utils/sharedHelpers";
import {
  autoDetectMatchingColumns,
  executeSmartLookup,
  SmartLookupResult,
  cleanForComparison
} from "../utils/fuzzyCompare";

interface FileMergerProps {
  mainData: any[];
  fileName?: string;
  setMainData?: (data: any[]) => void;
  setRawImportedData?: (data: any[]) => void;
  setColumns?: (cols: string[]) => void;
  setFileName?: (name: string) => void;
  setCustomColConfigs?: (configs: any[]) => void;
  setMapping?: (mapping: any) => void;
  setLoading?: (loading: boolean) => void;
  setProgress?: (progress: number) => void;
  setStatusMessage?: (msg: string) => void;
  onExportExcel?: () => void;
}

export const FileMerger: React.FC<FileMergerProps> = ({
  mainData = [],
  fileName = "",
  setMainData,
  setRawImportedData,
  setColumns,
  setFileName,
}) => {
  const [incomingFile, setIncomingFile] = useState<File | null>(null);
  const [incomingData, setIncomingData] = useState<any[]>([]);
  const [incomingColumns, setIncomingColumns] = useState<string[]>([]);
  const [mergeMode, setMergeMode] = useState<"join" | "union">("join"); // Mặc định VLOOKUP mềm

  // Cột khóa
  const [mainKeyCol, setMainKeyCol] = useState<string>("");
  const [incomingKeyCol, setIncomingKeyCol] = useState<string>("");

  // Cột cần lấy từ tệp 2
  const [selectedColsToPull, setSelectedColsToPull] = useState<string[]>([]);
  const [prefixColName, setPrefixColName] = useState<string>("");

  // Tùy chọn làm sạch & so khớp bỏ qua những thứ không ảnh hưởng
  const [ignoreDiacritics, setIgnoreDiacritics] = useState<boolean>(true);
  const [ignoreSpaces, setIgnoreSpaces] = useState<boolean>(true);
  const [ignorePunctuation, setIgnorePunctuation] = useState<boolean>(true);
  const [ignoreCase, setIgnoreCase] = useState<boolean>(true);
  const [similarityThreshold, setSimilarityThreshold] = useState<number>(0.90); // 90%

  // Trạng thái quét thử & chẩn đoán
  const [previewResult, setPreviewResult] = useState<SmartLookupResult | null>(null);
  const [previewTab, setPreviewTab] = useState<"matched" | "unmatched">("matched");
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [mergeStatus, setMergeStatus] = useState<string | null>(null);

  // BẢNG DỮ LIỆU KẾT QUẢ SAU KHI THỰC HIỆN GHÉP (LIVE RESULT TABLE)
  const [mergedResultData, setMergedResultData] = useState<any[] | null>(null);
  const [mergedResultCols, setMergedResultCols] = useState<string[]>([]);
  const [mergedResultType, setMergedResultType] = useState<"join" | "union" | null>(null);
  const [mergedNewCols, setMergedNewCols] = useState<string[]>([]);
  const [originalMainRowCount, setOriginalMainRowCount] = useState<number>(0);
  const [mergedNewRowCount, setMergedNewRowCount] = useState<number>(0);

  // Phân trang & Tìm kiếm trên bảng kết quả
  const [resultSearchTerm, setResultSearchTerm] = useState<string>("");
  const [resultFilterMode, setResultFilterMode] = useState<"all" | "matched" | "unmatched" | "incoming_rows">("all");
  const [resultPage, setResultPage] = useState<number>(1);
  const [resultPageSize, setResultPageSize] = useState<number>(30);

  const resultTableRef = useRef<HTMLDivElement>(null);

  const mainColumns = useMemo(() => {
    return mainData.length > 0 ? Object.keys(mainData[0]) : [];
  }, [mainData]);

  // Tự động phát hiện các cặp cột có tên giống nhau / đồng nghĩa
  const detectedColumnPairs = useMemo(() => {
    if (mainColumns.length === 0 || incomingColumns.length === 0) return [];
    return autoDetectMatchingColumns(mainColumns, incomingColumns);
  }, [mainColumns, incomingColumns]);

  // Xử lý nạp file thứ hai
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIncomingFile(file);
    setPreviewResult(null);
    setMergeStatus(null);
    setMergedResultData(null);

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wsName], { header: 1, defval: "" });
        const parsed = parse2DArrayWithSmartHeader(rawRows);
        setIncomingData(parsed.data);
        setIncomingColumns(parsed.columns);

        // Tự động gợi ý cột khóa tốt nhất
        const pairs = autoDetectMatchingColumns(mainColumns, parsed.columns);
        if (pairs.length > 0) {
          setMainKeyCol(pairs[0].col1);
          setIncomingKeyCol(pairs[0].col2);
        } else if (parsed.columns.length > 0) {
          setIncomingKeyCol(parsed.columns[0]);
          if (mainColumns.length > 0) setMainKeyCol(mainColumns[0]);
        }

        // Mặc định chọn tất cả các cột của tệp 2 để lấy sang (trừ cột khóa)
        const defaultCols = parsed.columns.filter(c => c !== (pairs.length > 0 ? pairs[0].col2 : parsed.columns[0]));
        setSelectedColsToPull(defaultCols);
      } catch (err: any) {
        alert("Lỗi đọc tệp ghép: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // Áp dụng cặp cột được hệ thống AI gợi ý
  const handleApplySuggestedPair = (col1: string, col2: string) => {
    setMainKeyCol(col1);
    setIncomingKeyCol(col2);
    setSelectedColsToPull(prev => prev.filter(c => c !== col2));
    setPreviewResult(null);
  };

  // Chuyển đổi chọn/bỏ chọn cột cần lấy sang
  const toggleColToPull = (col: string) => {
    setSelectedColsToPull(prev =>
      prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
    );
    setPreviewResult(null);
  };

  const handleSelectAllCols = () => {
    setSelectedColsToPull(incomingColumns.filter(c => c !== incomingKeyCol));
    setPreviewResult(null);
  };

  const handleDeselectAllCols = () => {
    setSelectedColsToPull([]);
    setPreviewResult(null);
  };

  // Quét thử đối chiếu & Chẩn đoán lỗi
  const handleRunDiagnosticPreview = () => {
    if (!mainKeyCol || !incomingKeyCol) {
      alert("Vui lòng chọn cột khóa ở cả Tệp 1 và Tệp 2 để đối chiếu!");
      return;
    }
    if (selectedColsToPull.length === 0) {
      alert("Vui lòng chọn ít nhất một cột từ Tệp 2 để lấy sang!");
      return;
    }

    setIsAnalyzing(true);
    setTimeout(() => {
      try {
        const result = executeSmartLookup({
          mainData,
          incomingData,
          mainKeyCol,
          incomingKeyCol,
          columnsToPull: selectedColsToPull,
          similarityThreshold,
          ignoreDiacritics,
          ignoreSpaces,
          ignorePunctuation,
          ignoreCase,
          prefixColName,
        });

        setPreviewResult(result);
        if (result.stats.unmatched > 0 && result.stats.exactMatches === 0 && result.stats.fuzzyMatches === 0) {
          setPreviewTab("unmatched");
        } else {
          setPreviewTab("matched");
        }
      } catch (err: any) {
        alert("Lỗi khi quét đối chiếu: " + err.message);
      } finally {
        setIsAnalyzing(false);
      }
    }, 50);
  };

  // Xuất file Excel danh sách các dòng KHÔNG KHỚP để người dùng rà soát lỗi
  const handleExportUnmatched = () => {
    if (!previewResult || previewResult.sampleUnmatched.length === 0) {
      alert("Không có dòng lỗi nào để xuất!");
      return;
    }

    const cleanOpts = { ignoreDiacritics, ignoreSpaces, ignorePunctuation, ignoreCase };
    const exactSet = new Set<string>();
    incomingData.forEach(row => {
      const k = cleanForComparison(row[incomingKeyCol], cleanOpts);
      if (k) exactSet.add(k);
    });

    const exportRows: any[] = [];
    mainData.forEach((row, idx) => {
      const rawVal = row[mainKeyCol];
      const cleaned = cleanForComparison(rawVal, cleanOpts);
      if (!cleaned || !exactSet.has(cleaned)) {
        exportRows.push({
          "Dòng số": idx + 1,
          [`Giá trị khóa tệp chính (${mainKeyCol})`]: rawVal || "(Rỗng)",
          "Trạng thái": !cleaned ? "Ô trống / Rỗng" : "Không tìm thấy mã khớp bên Tệp 2",
          ...row
        });
      }
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DanhSach_KhongKhop");
    XLSX.writeFile(wb, `DanhSach_KhongKhop_${mainKeyCol}.xlsx`);
  };

  // Xuất file Excel BẢNG KẾT QUẢ ĐÃ GHÉP
  const handleExportMergedExcel = () => {
    if (!mergedResultData || mergedResultData.length === 0) {
      alert("Chưa có bảng kết quả để xuất!");
      return;
    }

    const ws = XLSX.utils.json_to_sheet(mergedResultData);
    const wb = XLSX.utils.book_new();
    const sheetName = mergedResultType === "join" ? "KetQua_Vlookup" : "KetQua_GhepDoc";
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const fileNameExport = mergedResultType === "join"
      ? `${(fileName || "DuLieu").replace(/\.[^/.]+$/, "")}_DaGhepCot_Vlookup.xlsx`
      : `${(fileName || "DuLieu").replace(/\.[^/.]+$/, "")}_DaGhepDoc_Union.xlsx`;
    XLSX.writeFile(wb, fileNameExport);
  };

  // THỰC HIỆN GHÉP CHÍNH THỨC & HIỂN THỊ BẢNG KẾT QUẢ
  const handleExecuteMerge = () => {
    if (incomingData.length === 0) {
      alert("Vui lòng tải lên tệp thứ hai để thực hiện ghép!");
      return;
    }

    if (mergeMode === "union") {
      // 1. GHÉP NỐI TIẾP DÒNG (UNION)
      const prevMainCount = mainData.length;
      const merged = [...mainData, ...incomingData];
      const allCols = Array.from(new Set([...mainColumns, ...incomingColumns]));

      if (setMainData) setMainData(merged);
      if (setRawImportedData) setRawImportedData(merged);
      if (setColumns) setColumns(allCols);
      if (setFileName && fileName) setFileName(`${fileName.replace(/\.[^/.]+$/, "")}_DaGhep.xlsx`);

      // Cập nhật bảng kết quả trực tiếp hiển thị tại trang này!
      setMergedResultData(merged);
      setMergedResultCols(allCols);
      setMergedResultType("union");
      setOriginalMainRowCount(prevMainCount);
      setMergedNewRowCount(incomingData.length);
      setMergedNewCols(incomingColumns.filter(c => !mainColumns.includes(c)));
      setResultPage(1);
      setResultFilterMode("all");

      setMergeStatus(
        `Ghép nối hàng dọc thành công! Đã nối thêm ${incomingData.length.toLocaleString("vi-VN")} dòng từ Tệp 2 vào Tệp 1. Tổng bảng mới: ${merged.length.toLocaleString("vi-VN")} dòng, ${allCols.length} cột.`
      );

      setTimeout(() => {
        resultTableRef.current?.scrollIntoView({ behavior: "smooth" });
      }, 100);
      return;
    }

    // 2. GHÉP CỘT VLOOKUP MỀM (JOIN)
    if (!mainKeyCol || !incomingKeyCol) {
      alert("Vui lòng chọn cột khóa ở cả hai tệp để ghép dữ liệu!");
      return;
    }

    if (selectedColsToPull.length === 0) {
      alert("Vui lòng chọn ít nhất một cột cần lấy từ Tệp 2 sang!");
      return;
    }

    const result = executeSmartLookup({
      mainData,
      incomingData,
      mainKeyCol,
      incomingKeyCol,
      columnsToPull: selectedColsToPull,
      similarityThreshold,
      ignoreDiacritics,
      ignoreSpaces,
      ignorePunctuation,
      ignoreCase,
      prefixColName,
    });

    const newColumns = Array.from(new Set([...mainColumns, ...result.addedColumns]));

    if (setMainData) setMainData(result.mergedData);
    if (setRawImportedData) setRawImportedData(result.mergedData);
    if (setColumns) setColumns(newColumns);
    if (setFileName && fileName) {
      setFileName(`${fileName.replace(/\.[^/.]+$/, "")}_Vlookup.xlsx`);
    }

    // Cập nhật bảng kết quả trực tiếp hiển thị tại trang này!
    setMergedResultData(result.mergedData);
    setMergedResultCols(newColumns);
    setMergedResultType("join");
    setMergedNewCols(result.addedColumns);
    setOriginalMainRowCount(mainData.length);
    setMergedNewRowCount(0);
    setResultPage(1);
    setResultFilterMode("all");
    setPreviewResult(result);

    setMergeStatus(
      `Ghép cột VLOOKUP thành công! Đã bổ sung ${result.addedColumns.length} cột mới [${result.addedColumns.join(", ")}]. Khớp ${result.stats.exactMatches.toLocaleString("vi-VN")} dòng (100%) và ${result.stats.fuzzyMatches.toLocaleString("vi-VN")} dòng (gần đúng ≥ ${(similarityThreshold * 100).toFixed(0)}%). Tỷ lệ khớp: ${result.stats.matchRate}%.`
    );

    setTimeout(() => {
      resultTableRef.current?.scrollIntoView({ behavior: "smooth" });
    }, 100);
  };

  // Lọc dữ liệu trên bảng kết quả
  const filteredResultData = useMemo(() => {
    if (!mergedResultData) return [];

    let list = mergedResultData;

    // 1. Lọc theo trạng thái khớp hoặc theo nguồn dòng
    if (mergedResultType === "join" && mergedNewCols.length > 0) {
      const checkCol = mergedNewCols[0];
      if (resultFilterMode === "matched") {
        list = list.filter(row => row[checkCol] !== undefined && row[checkCol] !== null && String(row[checkCol]).trim() !== "");
      } else if (resultFilterMode === "unmatched") {
        list = list.filter(row => row[checkCol] === undefined || row[checkCol] === null || String(row[checkCol]).trim() === "");
      }
    } else if (mergedResultType === "union") {
      if (resultFilterMode === "incoming_rows") {
        list = list.slice(originalMainRowCount);
      }
    }

    // 2. Lọc theo từ khóa tìm kiếm
    if (resultSearchTerm.trim()) {
      const q = resultSearchTerm.toLowerCase().trim();
      list = list.filter(row => {
        return Object.values(row).some(v => String(v || "").toLowerCase().includes(q));
      });
    }

    return list;
  }, [mergedResultData, mergedResultType, mergedNewCols, resultFilterMode, resultSearchTerm, originalMainRowCount]);

  // Phân trang
  const totalResultPages = Math.max(1, Math.ceil(filteredResultData.length / resultPageSize));
  const paginatedResultRows = useMemo(() => {
    const start = (resultPage - 1) * resultPageSize;
    return filteredResultData.slice(start, start + resultPageSize);
  }, [filteredResultData, resultPage, resultPageSize]);

  return (
    <div className="bg-white rounded-none border border-slate-200 p-4 sm:p-5 space-y-5 text-slate-800 font-sans shadow-2xs">
      {/* Tiêu đề thanh công cụ */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-200">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-none bg-[#286e42] flex items-center justify-center text-white shrink-0">
            <GitMerge className="w-4 h-4" />
          </div>
          <div>
            <h3 className="font-bold text-slate-900 text-sm sm:text-base flex items-center gap-2">
              GHÉP NỐI & VLOOKUP THÔNG MINH GIỮA 2 TỆP DỮ LIỆU
              <span className="text-[11px] font-normal bg-indigo-100 text-indigo-800 px-2 py-0.5 border border-indigo-200 rounded-none">
                Bỏ qua khoảng trắng • Bỏ dấu tiếng Việt • Khớp tương đồng ≥90%
              </span>
            </h3>
            <p className="text-xs text-slate-500">
              Tự động nhận diện cột đồng nghĩa, đối chiếu mềm không phân biệt định dạng, nhặt cột từ bảng này sang bảng khác và hiển thị trực tiếp bảng kết quả.
            </p>
          </div>
        </div>

        {/* Nút chuyển chế độ */}
        <div className="flex items-center gap-1 border border-slate-300 p-0.5 bg-slate-50 rounded-none text-xs">
          <button
            type="button"
            onClick={() => {
              setMergeMode("join");
              setMergedResultData(null);
            }}
            className={`px-3 py-1 font-bold rounded-none transition-colors cursor-pointer border-0 ${
              mergeMode === "join"
                ? "bg-[#0284c7] text-white shadow-2xs"
                : "text-slate-600 hover:bg-slate-200"
            }`}
          >
            🔍 VLOOKUP Mềm (Lấy cột)
          </button>
          <button
            type="button"
            onClick={() => {
              setMergeMode("union");
              setMergedResultData(null);
            }}
            className={`px-3 py-1 font-bold rounded-none transition-colors cursor-pointer border-0 ${
              mergeMode === "union"
                ? "bg-[#0284c7] text-white shadow-2xs"
                : "text-slate-600 hover:bg-slate-200"
            }`}
          >
            📋 Ghép nối dòng (Gộp dọc)
          </button>
        </div>
      </div>

      {/* Thông báo kết quả thực hiện */}
      {mergeStatus && (
        <div className="p-3 bg-emerald-50 border border-emerald-300 rounded-none text-emerald-900 text-xs font-semibold flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>{mergeStatus}</span>
          </div>
          {mergedResultData && (
            <button
              type="button"
              onClick={handleExportMergedExcel}
              className="px-2.5 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded-none text-[11px] font-bold flex items-center gap-1 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Tải file Excel kết quả
            </button>
          )}
        </div>
      )}

      {/* KHUNG THÔNG TIN 2 TỆP */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
        {/* Tệp 1 (Chính) */}
        <div className="p-3 bg-slate-50 border border-slate-200 rounded-none space-y-1">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-700 flex items-center gap-1">
              <TableProperties className="w-3.5 h-3.5 text-slate-500" /> TỆP 1 (Bảng dữ liệu chính đang mở):
            </span>
            <span className="text-[11px] font-mono bg-white px-2 py-0.5 border border-slate-200 text-slate-600">
              {mainData.length.toLocaleString("vi-VN")} dòng • {mainColumns.length} cột
            </span>
          </div>
          <p className="text-xs font-bold text-[#286e42] truncate font-mono">
            {fileName || "Dữ liệu hiện hành trong hệ thống"}
          </p>
        </div>

        {/* Tệp 2 (Cần nạp để lấy dữ liệu) */}
        <div className="p-3 bg-sky-50/60 border border-sky-200 rounded-none space-y-1.5 flex flex-col justify-between">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-sky-900 flex items-center gap-1">
              <FileSpreadsheet className="w-3.5 h-3.5 text-sky-600" /> TỆP 2 (Bảng phụ / Đối chiếu cần lấy cột):
            </span>
            {incomingData.length > 0 && (
              <span className="text-[11px] font-mono bg-white px-2 py-0.5 border border-sky-200 text-sky-700 font-bold">
                {incomingData.length.toLocaleString("vi-VN")} dòng • {incomingColumns.length} cột
              </span>
            )}
          </div>

          <div className="flex items-center justify-between gap-2">
            <p className="text-xs font-medium text-slate-700 truncate font-mono">
              {incomingFile ? incomingFile.name : "(Chưa tải tệp thứ 2 lên)"}
            </p>
            <label className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-[#0284c7] hover:bg-sky-700 text-white text-xs font-bold rounded-none transition-colors cursor-pointer border-0 shrink-0">
              <Upload className="w-3.5 h-3.5" />
              {incomingFile ? "Đổi tệp Excel khác" : "Chọn tệp Excel / CSV thứ 2"}
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>
        </div>
      </div>

      {/* GỢI Ý CÁC CỘT TƯƠNG ĐỒNG / ĐỒNG NGHĨA PHÁT HIỆN TỰ ĐỘNG */}
      {incomingData.length > 0 && detectedColumnPairs.length > 0 && (
        <div className="bg-amber-50/80 border border-amber-200 p-2.5 rounded-none space-y-1.5">
          <div className="flex items-center gap-1.5 text-xs font-bold text-amber-900">
            <Sparkles className="w-3.5 h-3.5 text-amber-600 shrink-0" />
            <span>Phát hiện các cặp cột có tên tương đồng hoặc đồng nghĩa giữa 2 file:</span>
            <span className="text-[11px] font-normal text-amber-700 italic">(Bấm vào để tự động chọn làm khóa đối chiếu)</span>
          </div>
          <div className="flex flex-wrap items-center gap-2 pt-0.5">
            {detectedColumnPairs.map((pair, idx) => {
              const isSelected = mainKeyCol === pair.col1 && incomingKeyCol === pair.col2;
              return (
                <button
                  key={idx}
                  type="button"
                  onClick={() => handleApplySuggestedPair(pair.col1, pair.col2)}
                  className={`text-xs px-2.5 py-1 rounded-none border transition-all flex items-center gap-1.5 cursor-pointer ${
                    isSelected
                      ? "bg-amber-600 text-white border-amber-700 shadow-2xs font-bold"
                      : "bg-white hover:bg-amber-100 text-slate-800 border-amber-300 font-medium"
                  }`}
                  title={`${pair.reason} - Bấm để áp dụng`}
                >
                  <span className="font-mono text-[11.5px]">{pair.col1}</span>
                  <ArrowRight className="w-3 h-3 opacity-60" />
                  <span className="font-mono text-[11.5px]">{pair.col2}</span>
                  <span className={`text-[10px] px-1 py-0.2 rounded-none ${isSelected ? "bg-amber-800 text-white" : "bg-amber-200 text-amber-900"}`}>
                    {(pair.similarity * 100).toFixed(0)}%
                  </span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* NỘI DUNG CẤU HÌNH VLOOKUP MỀM (JOIN) */}
      {incomingData.length > 0 && mergeMode === "join" && (
        <div className="space-y-4 bg-slate-50/70 p-3.5 sm:p-4 border border-slate-200 rounded-none">
          {/* 1. CHỌN CỘT KHÓA ĐỐI CHIẾU */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5 items-end">
            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-800 flex items-center gap-1">
                <span className="w-4 h-4 rounded-none bg-[#286e42] text-white inline-flex items-center justify-center text-[10px] font-bold">1</span>
                Cột khóa ở Tệp 1 (Bảng chính):
              </label>
              <select
                value={mainKeyCol}
                onChange={e => {
                  setMainKeyCol(e.target.value);
                  setPreviewResult(null);
                }}
                className="w-full bg-white border border-slate-300 rounded-none px-2.5 py-1.5 text-xs text-slate-800 font-medium focus:outline-none focus:border-indigo-500"
              >
                <option value="">-- Chọn cột khóa (VD: MST, Mã xã, Mã ĐB, Tên...) --</option>
                {mainColumns.map(c => (
                  <option key={c} value={c}>
                    {c} {mainData.length > 0 && mainData[0][c] ? `(Ví dụ: ${String(mainData[0][c]).slice(0, 20)})` : ""}
                  </option>
                ))}
              </select>
            </div>

            <div className="space-y-1">
              <label className="text-xs font-bold text-slate-800 flex items-center gap-1">
                <span className="w-4 h-4 rounded-none bg-[#0284c7] text-white inline-flex items-center justify-center text-[10px] font-bold">2</span>
                Cột khóa ở Tệp 2 (Bảng đối chiếu):
              </label>
              <select
                value={incomingKeyCol}
                onChange={e => {
                  setIncomingKeyCol(e.target.value);
                  setPreviewResult(null);
                }}
                className="w-full bg-white border border-slate-300 rounded-none px-2.5 py-1.5 text-xs text-slate-800 font-medium focus:outline-none focus:border-indigo-500"
              >
                <option value="">-- Chọn cột khóa tương ứng ở Tệp 2 --</option>
                {incomingColumns.map(c => (
                  <option key={c} value={c}>
                    {c} {incomingData.length > 0 && incomingData[0][c] ? `(Ví dụ: ${String(incomingData[0][c]).slice(0, 20)})` : ""}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* 2. CÁC TÙY CHỌN BỎ QUA KHOẢNG TRẮNG, BỎ DẤU, KÝ TỰ ĐẶC BIỆT */}
          <div className="p-3 bg-white border border-slate-200 rounded-none space-y-2.5">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Sliders className="w-3.5 h-3.5 text-indigo-600" />
                Cơ chế làm sạch dữ liệu trước khi so khớp (Loại bỏ những yếu tố không ảnh hưởng):
              </span>
              <span className="text-[11px] text-slate-500 italic">
                Cả 2 tệp đều tự động chuẩn hóa qua bộ lọc này trước khi đem ra đối chiếu
              </span>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
              <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 bg-slate-50 p-2 border border-slate-200 hover:bg-slate-100">
                <input
                  type="checkbox"
                  checked={ignoreDiacritics}
                  onChange={e => {
                    setIgnoreDiacritics(e.target.checked);
                    setPreviewResult(null);
                  }}
                  className="text-indigo-600 rounded-none"
                />
                <span className="font-medium">Bỏ dấu tiếng Việt <span className="text-[10px] text-slate-400 block">(Tân Lập = Tan Lap)</span></span>
              </label>

              <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 bg-slate-50 p-2 border border-slate-200 hover:bg-slate-100">
                <input
                  type="checkbox"
                  checked={ignoreSpaces}
                  onChange={e => {
                    setIgnoreSpaces(e.target.checked);
                    setPreviewResult(null);
                  }}
                  className="text-indigo-600 rounded-none"
                />
                <span className="font-medium">Bỏ mọi khoảng trắng <span className="text-[10px] text-slate-400 block">(" 01 02 " = "0102")</span></span>
              </label>

              <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 bg-slate-50 p-2 border border-slate-200 hover:bg-slate-100">
                <input
                  type="checkbox"
                  checked={ignorePunctuation}
                  onChange={e => {
                    setIgnorePunctuation(e.target.checked);
                    setPreviewResult(null);
                  }}
                  className="text-indigo-600 rounded-none"
                />
                <span className="font-medium">Bỏ dấu gạch & ký tự lạ <span className="text-[10px] text-slate-400 block">(_, -, ., /, ,, :, ;)</span></span>
              </label>

              <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 bg-slate-50 p-2 border border-slate-200 hover:bg-slate-100">
                <input
                  type="checkbox"
                  checked={ignoreCase}
                  onChange={e => {
                    setIgnoreCase(e.target.checked);
                    setPreviewResult(null);
                  }}
                  className="text-indigo-600 rounded-none"
                />
                <span className="font-medium">Bỏ phân biệt HOA/thường <span className="text-[10px] text-slate-400 block">(CÔNG TY = cong ty)</span></span>
              </label>
            </div>

            {/* Mức độ tương đồng Fuzzy Match (>= 90%) */}
            <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-bold text-slate-800">Độ tương đồng tối thiểu (Fuzzy Match):</span>
                {[
                  { val: 1.0, label: "100% (Tuyệt đối)" },
                  { val: 0.95, label: "95% (Rất sát)" },
                  { val: 0.90, label: "90% (Khuyên dùng)" },
                  { val: 0.85, label: "85% (Mở rộng)" },
                ].map(item => (
                  <button
                    key={item.val}
                    type="button"
                    onClick={() => {
                      setSimilarityThreshold(item.val);
                      setPreviewResult(null);
                    }}
                    className={`px-2 py-0.5 text-xs font-bold rounded-none border cursor-pointer transition-colors ${
                      similarityThreshold === item.val
                        ? "bg-indigo-600 text-white border-indigo-700 shadow-2xs"
                        : "bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300"
                    }`}
                  >
                    {item.label}
                  </button>
                ))}
              </div>

              <div className="flex items-center gap-2">
                <span className="text-slate-500 font-mono text-[11px]">Tùy chỉnh:</span>
                <input
                  type="range"
                  min="0.75"
                  max="1.0"
                  step="0.01"
                  value={similarityThreshold}
                  onChange={e => {
                    setSimilarityThreshold(parseFloat(e.target.value));
                    setPreviewResult(null);
                  }}
                  className="w-24 accent-indigo-600 cursor-pointer"
                />
                <span className="font-bold font-mono text-indigo-700 w-10 text-right text-xs">
                  {(similarityThreshold * 100).toFixed(0)}%
                </span>
              </div>
            </div>
          </div>

          {/* 3. CHỌN CỘT CẦN LẤY SANG TỆP 1 */}
          <div className="p-3 bg-white border border-slate-200 rounded-none space-y-2">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-800">
                  3. Chọn cột cần nhặt từ Tệp 2 sang Tệp 1 ({selectedColsToPull.length} cột đã chọn):
                </span>
                <button
                  type="button"
                  onClick={handleSelectAllCols}
                  className="text-[11px] text-[#0284c7] hover:underline font-bold cursor-pointer"
                >
                  [Chọn tất cả]
                </button>
                <button
                  type="button"
                  onClick={handleDeselectAllCols}
                  className="text-[11px] text-slate-500 hover:underline font-medium cursor-pointer"
                >
                  [Bỏ chọn hết]
                </button>
              </div>

              <div className="flex items-center gap-1.5 text-xs">
                <span className="text-slate-600 text-[11px]">Thêm tiền tố tên cột mới:</span>
                <input
                  type="text"
                  value={prefixColName}
                  onChange={e => {
                    setPrefixColName(e.target.value);
                    setPreviewResult(null);
                  }}
                  placeholder="VD: T2_, Phụ_..."
                  className="w-24 bg-white border border-slate-300 px-2 py-0.5 text-xs font-mono text-slate-800 rounded-none"
                />
              </div>
            </div>

            {/* Danh sách checkbox các cột */}
            <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto p-2 bg-slate-50 border border-slate-200">
              {incomingColumns.map(col => {
                const isKey = col === incomingKeyCol;
                const isChecked = selectedColsToPull.includes(col);
                return (
                  <label
                    key={col}
                    className={`inline-flex items-center gap-1 px-2 py-1 text-xs border rounded-none cursor-pointer transition-colors ${
                      isKey
                        ? "bg-slate-200 border-slate-300 text-slate-400 cursor-not-allowed"
                        : isChecked
                        ? "bg-sky-100 border-sky-300 text-sky-900 font-bold"
                        : "bg-white border-slate-300 text-slate-700 hover:bg-slate-100 font-normal"
                    }`}
                    title={isKey ? "Đây là cột khóa, không thể nhặt sang" : `Nhặt cột ${col}`}
                  >
                    <input
                      type="checkbox"
                      disabled={isKey}
                      checked={isChecked}
                      onChange={() => toggleColToPull(col)}
                      className="text-sky-600 rounded-none w-3 h-3"
                    />
                    <span className="truncate max-w-[200px]">{col}</span>
                    {isKey && <span className="text-[9px] text-slate-500 italic">(Khóa)</span>}
                  </label>
                );
              })}
            </div>
          </div>

          {/* 4. KHU VỰC THAO TÁC: QUÉT THỬ & CHẨN ĐOÁN LỖI HOẶC ÁP DỤNG LUÔN */}
          <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
            <button
              type="button"
              onClick={handleRunDiagnosticPreview}
              disabled={isAnalyzing || !mainKeyCol || !incomingKeyCol}
              className="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs rounded-none transition-colors flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
            >
              <Search className="w-3.5 h-3.5" />
              <span>{isAnalyzing ? "Đang quét đối chiếu..." : "🔍 Quét thử đối chiếu & Chẩn đoán sai sót"}</span>
            </button>

            <button
              type="button"
              onClick={handleExecuteMerge}
              disabled={!mainKeyCol || !incomingKeyCol || selectedColsToPull.length === 0}
              className="px-5 py-2 bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs rounded-none transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs border-0 disabled:opacity-50"
            >
              <GitMerge className="w-4 h-4" />
              <span>Thực Hiện Ghép Cột & Xem Kết Quả</span>
            </button>
          </div>

          {/* 5. BẢNG KẾT QUẢ QUÉT THỬ & CHẨN ĐOÁN (DIAGNOSTIC DASHBOARD) */}
          {previewResult && (
            <div className="bg-white border border-slate-300 rounded-none p-3 space-y-3 shadow-xs">
              {/* Thống kê tỷ lệ khớp */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-none">
                  <div className="text-[11px] text-slate-500">Tổng dòng Tệp chính:</div>
                  <div className="text-base font-bold font-mono text-slate-800">
                    {previewResult.stats.totalMain.toLocaleString("vi-VN")}
                  </div>
                </div>

                <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-none">
                  <div className="text-[11px] text-emerald-800 font-bold flex items-center gap-1">
                    <Check className="w-3 h-3" /> Khớp chính xác 100%:
                  </div>
                  <div className="text-base font-bold font-mono text-emerald-700">
                    {previewResult.stats.exactMatches.toLocaleString("vi-VN")} dòng
                  </div>
                </div>

                <div className="p-2.5 bg-indigo-50 border border-indigo-200 rounded-none">
                  <div className="text-[11px] text-indigo-800 font-bold flex items-center gap-1">
                    <Sparkles className="w-3 h-3" /> Khớp gần đúng (≥{(similarityThreshold * 100).toFixed(0)}%):
                  </div>
                  <div className="text-base font-bold font-mono text-indigo-700">
                    {previewResult.stats.fuzzyMatches.toLocaleString("vi-VN")} dòng
                  </div>
                </div>

                <div className="p-2.5 bg-rose-50 border border-rose-200 rounded-none">
                  <div className="text-[11px] text-rose-800 font-bold flex items-center gap-1">
                    <AlertCircle className="w-3 h-3" /> Không tìm thấy mã:
                  </div>
                  <div className="text-base font-bold font-mono text-rose-700">
                    {previewResult.stats.unmatched.toLocaleString("vi-VN")} dòng ({((previewResult.stats.unmatched / (previewResult.stats.totalMain || 1)) * 100).toFixed(1)}%)
                  </div>
                </div>
              </div>

              {/* Tabs chuyển đổi giữa Mẫu khớp và Chẩn đoán lỗi */}
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setPreviewTab("matched")}
                    className={`text-xs font-bold px-3 py-1 rounded-none border transition-colors cursor-pointer ${
                      previewTab === "matched"
                        ? "bg-emerald-600 text-white border-emerald-700"
                        : "bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200"
                    }`}
                  >
                    ✓ Xem mẫu các dòng khớp thành công ({previewResult.sampleMatches.length})
                  </button>

                  <button
                    type="button"
                    onClick={() => setPreviewTab("unmatched")}
                    className={`text-xs font-bold px-3 py-1 rounded-none border transition-colors cursor-pointer ${
                      previewTab === "unmatched"
                        ? "bg-rose-600 text-white border-rose-700"
                        : "bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200"
                    }`}
                  >
                    ⚠️ Chẩn đoán lỗi không khớp ({previewResult.stats.unmatched})
                  </button>
                </div>

                {previewResult.stats.unmatched > 0 && (
                  <button
                    type="button"
                    onClick={handleExportUnmatched}
                    className="text-xs font-bold text-rose-700 hover:text-rose-900 flex items-center gap-1 px-2 py-1 bg-rose-50 border border-rose-200 rounded-none cursor-pointer"
                    title="Tải danh sách các dòng không khớp về Excel để sửa"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Xuất file Excel các dòng không khớp</span>
                  </button>
                )}
              </div>

              {/* TAB 1: MẪU CÁC DÒNG KHỚP THÀNH CÔNG */}
              {previewTab === "matched" && (
                <div className="overflow-x-auto border border-slate-200 max-h-72">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-[#286e42] text-white text-[11px]">
                        <th className="px-2.5 py-1.5 border-r border-[#205835] w-12 text-center">Dòng</th>
                        <th className="px-2.5 py-1.5 border-r border-[#205835]">Khóa Tệp 1 [{mainKeyCol}]</th>
                        <th className="px-2.5 py-1.5 border-r border-[#205835]">Khóa Tệp 2 [{incomingKeyCol}]</th>
                        <th className="px-2.5 py-1.5 border-r border-[#205835] text-center w-28">Độ tương đồng</th>
                        <th className="px-2.5 py-1.5">Dữ liệu lấy sang (Cột đầu tiên: {selectedColsToPull[0] || "..."})</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 font-mono">
                      {previewResult.sampleMatches.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="p-4 text-center text-slate-400 italic font-sans text-xs">
                            Chưa tìm thấy dòng nào khớp với cấu hình hiện tại. Hãy thử giảm độ tương đồng hoặc kiểm tra lại cột khóa!
                          </td>
                        </tr>
                      ) : (
                        previewResult.sampleMatches.map((m, idx) => (
                          <tr key={idx} className="hover:bg-sky-50/40">
                            <td className="px-2.5 py-1 text-center text-slate-500 border-r border-slate-200">
                              #{m.rowIdx}
                            </td>
                            <td className="px-2.5 py-1 font-bold text-slate-800 border-r border-slate-200">
                              {m.mainVal}
                            </td>
                            <td className="px-2.5 py-1 text-slate-700 border-r border-slate-200">
                              {m.incomingVal}
                            </td>
                            <td className="px-2.5 py-1 text-center border-r border-slate-200">
                              {m.isExact ? (
                                <span className="bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded-none font-bold text-[10.5px]">
                                  100% (Chuẩn)
                                </span>
                              ) : (
                                <span className="bg-indigo-100 text-indigo-800 px-1.5 py-0.5 rounded-none font-bold text-[10.5px]">
                                  {(m.similarity * 100).toFixed(0)}% (Fuzzy)
                                </span>
                              )}
                            </td>
                            <td className="px-2.5 py-1 font-sans text-slate-700 truncate max-w-[320px]">
                              {Object.entries(m.pulledData).slice(0, 3).map(([col, val]) => (
                                <span key={col} className="mr-2 text-[11px] bg-slate-100 px-1 border border-slate-200">
                                  <strong>{col}:</strong> {String(val || "")}
                                </span>
                              ))}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              )}

              {/* TAB 2: CHẨN ĐOÁN LỖI KHÔNG KHỚP */}
              {previewTab === "unmatched" && (
                <div className="space-y-2">
                  <div className="p-2 bg-amber-50 border border-amber-200 text-amber-900 text-xs flex items-center gap-2">
                    <HelpCircle className="w-4 h-4 text-amber-600 shrink-0" />
                    <span>
                      <strong>Tại sao những dòng này không khớp?</strong> Bảng dưới đây phân tích từng dòng để bạn biết chính xác nguyên nhân (do ô rỗng, thiếu ký tự, hay không có mã tương đương bên Tệp 2).
                    </span>
                  </div>

                  <div className="overflow-x-auto border border-slate-200 max-h-72">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-rose-800 text-white text-[11px]">
                          <th className="px-2.5 py-1.5 border-r border-rose-900 w-12 text-center">Dòng</th>
                          <th className="px-2.5 py-1.5 border-r border-rose-900">Giá trị khóa Tệp 1 [{mainKeyCol}]</th>
                          <th className="px-2.5 py-1.5 border-r border-rose-900">Chẩn đoán lý do</th>
                          <th className="px-2.5 py-1.5">Mã gần giống nhất tìm thấy bên Tệp 2</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 font-sans">
                        {previewResult.sampleUnmatched.length === 0 ? (
                          <tr>
                            <td colSpan={4} className="p-4 text-center text-emerald-700 font-bold text-xs bg-emerald-50">
                              Tuyệt vời! 100% tất cả các dòng dữ liệu đều đã được khớp thành công!
                            </td>
                          </tr>
                        ) : (
                          previewResult.sampleUnmatched.map((un, idx) => (
                            <tr key={idx} className="hover:bg-rose-50/40">
                              <td className="px-2.5 py-1.5 text-center font-mono text-slate-500 border-r border-slate-200">
                                #{un.rowIdx}
                              </td>
                              <td className="px-2.5 py-1.5 font-bold font-mono text-rose-900 border-r border-slate-200 bg-rose-50/30">
                                "{un.mainVal}"
                              </td>
                              <td className="px-2.5 py-1.5 text-slate-700 border-r border-slate-200 text-[11.5px]">
                                {un.reason}
                              </td>
                              <td className="px-2.5 py-1.5 font-mono text-slate-600">
                                {un.bestCandidate ? (
                                  <div className="flex items-center gap-1.5">
                                    <span className="font-bold text-slate-800">"{un.bestCandidate.incomingVal}"</span>
                                    <span className="text-[10.5px] bg-amber-100 text-amber-900 px-1 py-0.2 border border-amber-200">
                                      {(un.bestCandidate.similarity * 100).toFixed(0)}% giống
                                    </span>
                                    <span className="text-[10px] text-slate-500 italic">
                                      (Nếu muốn nhận diện dòng này, hãy hạ ngưỡng xuống {(un.bestCandidate.similarity * 100 - 1).toFixed(0)}%)
                                    </span>
                                  </div>
                                ) : (
                                  <span className="text-slate-400 italic text-[11px]">(Không có mã nào tương tự)</span>
                                )}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* CHẾ ĐỘ UNION (GỘP HÀNG DỌC) */}
      {incomingData.length > 0 && mergeMode === "union" && (
        <div className="space-y-3 bg-slate-50 p-4 border border-slate-200 rounded-none">
          <div className="text-xs text-slate-700 space-y-1">
            <p className="font-bold text-slate-800">Gộp nối tiếp theo dòng (UNION):</p>
            <p className="text-slate-600">
              Tất cả {incomingData.length.toLocaleString("vi-VN")} dòng từ Tệp 2 sẽ được thêm trực tiếp vào cuối bảng Tệp 1.
              Các cột cùng tên sẽ được điền thẳng vào nhau, các cột mới từ Tệp 2 sẽ tự động được mở rộng thêm vào bảng.
            </p>
          </div>

          <div className="flex justify-end pt-2">
            <button
              type="button"
              onClick={handleExecuteMerge}
              className="px-5 py-2 bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs rounded-none transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs border-0"
            >
              <GitMerge className="w-4 h-4" />
              <span>Thực Hiện Gộp Nối Hàng Dọc ({incomingData.length.toLocaleString("vi-VN")} dòng) & Xem Kết Quả</span>
            </button>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 6. BẢNG DỮ LIỆU KẾT QUẢ SAU KHI THỰC HIỆN GHÉP (LIVE RESULT TABLE)       */}
      {/* ========================================================================= */}
      {mergedResultData && (
        <div
          ref={resultTableRef}
          className="mt-6 border-2 border-[#286e42] bg-white rounded-none shadow-sm space-y-3 p-3.5 sm:p-4"
        >
          {/* Header Bảng Kết Quả */}
          <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-200">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-3 h-3 bg-emerald-600 rounded-full animate-pulse"></span>
                <h4 className="text-sm sm:text-base font-bold text-slate-900 uppercase tracking-wide">
                  {mergedResultType === "join"
                    ? "BẢNG KẾT QUẢ ĐÃ GHÉP CỘT (VLOOKUP MỀM)"
                    : "BẢNG KẾT QUẢ ĐÃ GHÉP NỐI HÀNG DỌC (UNION)"}
                </h4>
                <span className="text-xs bg-[#286e42] text-white font-mono px-2 py-0.5 rounded-none font-bold">
                  {filteredResultData.length.toLocaleString("vi-VN")} dòng hiển thị
                </span>
                <span className="text-xs bg-slate-100 border border-slate-300 text-slate-700 font-mono px-2 py-0.5 rounded-none">
                  {mergedResultCols.length} cột
                </span>
              </div>
              <p className="text-xs text-slate-600 mt-1">
                {mergedResultType === "join" ? (
                  <span>
                    Đã bổ sung <strong>{mergedNewCols.length}</strong> cột mới từ Tệp 2 sang:{" "}
                    <span className="text-emerald-700 font-bold font-mono">
                      {mergedNewCols.map(c => `[${c}]`).join(", ")}
                    </span>
                    . Các cột này được tô màu xanh lá nổi bật trên bảng.
                  </span>
                ) : (
                  <span>
                    Bảng đã gộp bao gồm <strong>{originalMainRowCount.toLocaleString("vi-VN")}</strong> dòng từ Tệp 1 và{" "}
                    <strong className="text-sky-700">{mergedNewRowCount.toLocaleString("vi-VN")}</strong> dòng nối tiếp từ Tệp 2.
                  </span>
                )}
              </p>
            </div>

            {/* Các nút tác vụ nhanh */}
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleExportMergedExcel}
                className="px-3 py-1.5 bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs rounded-none transition-colors flex items-center gap-1.5 cursor-pointer shadow-2xs border-0"
                title="Tải bảng kết quả hoàn chỉnh về máy"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Tải tệp Excel kết quả (.xlsx)</span>
              </button>
            </div>
          </div>

          {/* Thanh công cụ: Tìm kiếm & Lọc dòng trên bảng kết quả */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-50 p-2.5 border border-slate-200 text-xs">
            {/* Bộ lọc theo trạng thái */}
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="font-bold text-slate-700 flex items-center gap-1">
                <Filter className="w-3.5 h-3.5 text-slate-500" /> Xem:
              </span>
              <button
                type="button"
                onClick={() => {
                  setResultFilterMode("all");
                  setResultPage(1);
                }}
                className={`px-2.5 py-1 rounded-none border cursor-pointer font-bold ${
                  resultFilterMode === "all"
                    ? "bg-[#0284c7] text-white border-sky-700"
                    : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                }`}
              >
                Tất cả ({mergedResultData.length.toLocaleString("vi-VN")})
              </button>

              {mergedResultType === "join" && (
                <>
                  <button
                    type="button"
                    onClick={() => {
                      setResultFilterMode("matched");
                      setResultPage(1);
                    }}
                    className={`px-2.5 py-1 rounded-none border cursor-pointer font-bold ${
                      resultFilterMode === "matched"
                        ? "bg-emerald-600 text-white border-emerald-700"
                        : "bg-white text-emerald-800 border-emerald-300 hover:bg-emerald-50"
                    }`}
                  >
                    ✓ Chỉ dòng ĐÃ KHỚP cột
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setResultFilterMode("unmatched");
                      setResultPage(1);
                    }}
                    className={`px-2.5 py-1 rounded-none border cursor-pointer font-bold ${
                      resultFilterMode === "unmatched"
                        ? "bg-rose-600 text-white border-rose-700"
                        : "bg-white text-rose-800 border-rose-300 hover:bg-rose-50"
                    }`}
                  >
                    ⚠️ Chỉ dòng CHƯA KHỚP (Ô trống)
                  </button>
                </>
              )}

              {mergedResultType === "union" && (
                <button
                  type="button"
                  onClick={() => {
                    setResultFilterMode("incoming_rows");
                    setResultPage(1);
                  }}
                  className={`px-2.5 py-1 rounded-none border cursor-pointer font-bold ${
                    resultFilterMode === "incoming_rows"
                      ? "bg-sky-600 text-white border-sky-700"
                      : "bg-white text-sky-800 border-sky-300 hover:bg-sky-50"
                  }`}
                >
                  + Chỉ xem dòng mới nối từ Tệp 2 ({mergedNewRowCount.toLocaleString("vi-VN")})
                </button>
              )}
            </div>

            {/* Ô tìm kiếm nhanh */}
            <div className="flex items-center gap-2 flex-1 max-w-xs ml-auto">
              <div className="relative w-full">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Tìm nhanh trong bảng..."
                  value={resultSearchTerm}
                  onChange={e => {
                    setResultSearchTerm(e.target.value);
                    setResultPage(1);
                  }}
                  className="w-full bg-white border border-slate-300 pl-8 pr-2.5 py-1 text-xs text-slate-800 rounded-none focus:outline-none focus:border-indigo-500"
                />
              </div>
              {resultSearchTerm && (
                <button
                  type="button"
                  onClick={() => setResultSearchTerm("")}
                  className="text-slate-400 hover:text-slate-600 text-xs px-1 cursor-pointer"
                >
                  ✕
                </button>
              )}
            </div>
          </div>

          {/* BẢNG DỮ LIỆU HIỂN THỊ KẾT QUẢ */}
          <div className="border border-slate-300 overflow-x-auto max-h-[500px] relative shadow-inner">
            <table className="w-full text-left text-xs border-collapse font-mono">
              <thead className="sticky top-0 z-10 shadow-xs">
                <tr className="bg-[#286e42] text-white text-[11px] font-sans">
                  <th className="px-2.5 py-2 border-r border-[#205835] w-14 text-center sticky left-0 bg-[#286e42] z-20">
                    STT
                  </th>
                  {mergedResultCols.map(col => {
                    const isNewCol = mergedNewCols.includes(col);
                    return (
                      <th
                        key={col}
                        className={`px-3 py-2 border-r border-[#205835] whitespace-nowrap ${
                          isNewCol
                            ? "bg-[#1e5433] text-emerald-200 font-bold ring-1 ring-emerald-400"
                            : ""
                        }`}
                        title={isNewCol ? `Cột mới được lấy từ Tệp 2: ${col}` : col}
                      >
                        <div className="flex items-center gap-1.5">
                          <span>{col}</span>
                          {isNewCol && (
                            <span className="bg-emerald-400 text-emerald-950 font-bold text-[9.5px] px-1 py-0.2 rounded-none">
                              MỚI
                            </span>
                          )}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-700 bg-white">
                {paginatedResultRows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={mergedResultCols.length + 1}
                      className="p-8 text-center text-slate-400 italic text-xs font-sans"
                    >
                      Không tìm thấy dòng nào phù hợp với bộ lọc hiện tại.
                    </td>
                  </tr>
                ) : (
                  paginatedResultRows.map((row, rowIdx) => {
                    const actualIdx = (resultPage - 1) * resultPageSize + rowIdx + 1;
                    const isIncomingRow = mergedResultType === "union" && actualIdx > originalMainRowCount;

                    return (
                      <tr
                        key={rowIdx}
                        className={`hover:bg-amber-50/50 transition-colors ${
                          isIncomingRow ? "bg-sky-50/40" : ""
                        }`}
                      >
                        <td className="px-2.5 py-1 text-center text-slate-500 border-r border-slate-200 text-[11px] sticky left-0 bg-white">
                          <div className="flex items-center justify-center gap-1">
                            <span>#{actualIdx}</span>
                            {isIncomingRow && (
                              <span className="text-[9px] bg-sky-200 text-sky-900 px-0.5 rounded-none font-sans font-bold">
                                Tệp 2
                              </span>
                            )}
                          </div>
                        </td>

                        {mergedResultCols.map(col => {
                          const val = row[col];
                          const isNewCol = mergedNewCols.includes(col);
                          const isFilled = val !== undefined && val !== null && String(val).trim() !== "";

                          return (
                            <td
                              key={col}
                              className={`px-3 py-1 border-r border-slate-200 text-[11.5px] truncate max-w-[280px] ${
                                isNewCol
                                  ? isFilled
                                    ? "bg-emerald-50/70 text-emerald-950 font-bold"
                                    : "bg-slate-50 text-slate-400 italic"
                                  : ""
                              }`}
                              title={String(val || "")}
                            >
                              {isFilled ? (
                                String(val)
                              ) : isNewCol ? (
                                <span className="text-[10px] text-slate-400 italic font-sans">(trống)</span>
                              ) : (
                                ""
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Phân trang kết quả */}
          <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600 pt-2 border-t border-slate-200">
            <div className="flex items-center gap-2">
              <span>
                Hiển thị dòng <strong>{(resultPage - 1) * resultPageSize + 1}</strong> đến{" "}
                <strong>{Math.min(resultPage * resultPageSize, filteredResultData.length)}</strong> trên tổng số{" "}
                <strong>{filteredResultData.length.toLocaleString("vi-VN")}</strong> dòng
              </span>

              <select
                value={resultPageSize}
                onChange={e => {
                  setResultPageSize(parseInt(e.target.value, 10));
                  setResultPage(1);
                }}
                className="bg-white border border-slate-300 px-1.5 py-0.5 text-xs rounded-none ml-2"
              >
                <option value={20}>20 dòng/trang</option>
                <option value={50}>50 dòng/trang</option>
                <option value={100}>100 dòng/trang</option>
                <option value={200}>200 dòng/trang</option>
              </select>
            </div>

            <div className="flex items-center gap-1">
              <button
                type="button"
                onClick={() => setResultPage(prev => Math.max(1, prev - 1))}
                disabled={resultPage <= 1}
                className="px-2 py-1 bg-white border border-slate-300 text-slate-700 hover:bg-slate-100 disabled:opacity-40 rounded-none cursor-pointer flex items-center gap-1 font-sans"
              >
                <ChevronLeft className="w-3.5 h-3.5" /> Trước
              </button>

              <span className="px-2.5 py-1 font-mono font-bold text-slate-800">
                {resultPage} / {totalResultPages}
              </span>

              <button
                type="button"
                onClick={() => setResultPage(prev => Math.min(totalResultPages, prev + 1))}
                disabled={resultPage >= totalResultPages}
                className="px-2 py-1 bg-white border border-slate-300 text-slate-700 hover:bg-slate-100 disabled:opacity-40 rounded-none cursor-pointer flex items-center gap-1 font-sans"
              >
                Sau <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default FileMerger;
