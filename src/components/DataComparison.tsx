import React, { useState, useMemo } from "react";
import {
  GitCompare,
  Upload,
  Download,
  CheckCircle2,
  AlertTriangle,
  Search,
  Sparkles,
  ArrowRight,
  RefreshCw,
  Copy,
  Check,
  Filter,
  Layers,
  ChevronRight,
  Database,
  Sliders,
  Wand2,
  FileSpreadsheet,
  CheckSquare,
  Square,
  BarChart3,
  TrendingUp,
  TrendingDown,
  Minus,
  Info,
  SlidersHorizontal,
  FolderOpen,
  Building2,
  GitMerge
} from "lucide-react";
import * as XLSX from "xlsx";
import { parse2DArrayWithSmartHeader } from "../utils/sharedHelpers";
import {
  autoDetectMatchingColumns,
  cleanForComparison,
  removeVietnameseTones
} from "../utils/fuzzyCompare";
import { AdminMergeResolver } from "./AdminMergeResolver";

interface DataComparisonProps {
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
  detectedSheets?: string[];
  sheetDataStore?: Record<string, { data: any[]; columns: string[] }>;
  currentActiveSheet?: string;
}

export const DataComparison: React.FC<DataComparisonProps> = ({
  mainData = [],
  fileName = "",
  setMainData,
  setRawImportedData,
  setColumns,
  setFileName,
  detectedSheets = [],
  sheetDataStore = {},
  currentActiveSheet = "",
}) => {
  // Dữ liệu nội bộ cho Tệp A nếu người dùng tải tệp trực tiếp trong tab này
  const [localDataA, setLocalDataA] = useState<any[] | null>(null);
  const [localFileNameA, setLocalFileNameA] = useState<string>("");

  const activeMainData = localDataA !== null ? localDataA : mainData;
  const activeFileName = localFileNameA || fileName;

  const [compareFile, setCompareFile] = useState<File | null>(null);
  const [compareData, setCompareData] = useState<any[]>([]);
  const [compareColumns, setCompareColumns] = useState<string[]>([]);
  const [compareSheetName, setCompareSheetName] = useState<string>("");

  // Khóa định danh 2 kỳ
  const [keyColA, setKeyColA] = useState<string>("");
  const [keyColB, setKeyColB] = useState<string>("");

  // Cột chỉ tiêu đối chiếu giá trị số liệu (VD: Doanh thu, Sản lượng, Số tiền...)
  const [compareValColA, setCompareValColA] = useState<string>("");
  const [compareValColB, setCompareValColB] = useState<string>("");

  // Tùy chọn làm sạch khóa khi so khớp (bỏ qua những thứ không ảnh hưởng)
  const [ignoreDiacritics, setIgnoreDiacritics] = useState<boolean>(true);
  const [ignoreSpaces, setIgnoreSpaces] = useState<boolean>(true);
  const [ignorePunctuation, setIgnorePunctuation] = useState<boolean>(true);
  const [ignoreCase, setIgnoreCase] = useState<boolean>(true);
  const [trimLeadingZeros, setTrimLeadingZeros] = useState<boolean>(false);

  // === 1. TÍNH NĂNG MỚI: LÀM SẠCH KHOẢNG CÁCH & CHỮ HOA / CHỮ THƯỜNG THEO CỘT ===
  const [showCleanPanel, setShowCleanPanel] = useState<boolean>(false);
  const [cleanTarget, setCleanTarget] = useState<"A" | "B" | "BOTH">("A");
  const [cleanSelectedCols, setCleanSelectedCols] = useState<string[]>([]);
  const [cleanTrimSpaces, setCleanTrimSpaces] = useState<boolean>(true);
  const [cleanCollapseSpaces, setCleanCollapseSpaces] = useState<boolean>(true);
  const [cleanRemoveAllSpaces, setCleanRemoveAllSpaces] = useState<boolean>(false);
  const [cleanCaseOption, setCleanCaseOption] = useState<"none" | "upper" | "lower" | "title">("upper");
  const [cleanRemoveAccents, setCleanRemoveAccents] = useState<boolean>(false);
  const [cleanSuccessMsg, setCleanSuccessMsg] = useState<string | null>(null);

  // === 2. TÍNH NĂNG MỚI: DANH SÁCH TRÙNG, DANH SÁCH LỆCH & BẢNG TỔNG HỢP ===
  const [activeResultTab, setActiveResultTab] = useState<"summary" | "same" | "diff" | "all">("summary");
  const [diffSubFilter, setDiffSubFilter] = useState<"all" | "diff_val" | "only_in_a" | "only_in_b">("all");
  const [searchQuery, setSearchQuery] = useState<string>("");

  // === 3. TÍNH NĂNG CHUYÊN BIỆT: GHÉP NỐI & ĐỐI SOÁT SÁP NHẬP XÃ / THÔN ===
  const [showAdminMergeMode, setShowAdminMergeMode] = useState<boolean>(false);

  // Tính năng chuyển/nhặt cột từ Tệp B sang Tệp A
  const [showTransferPanel, setShowTransferPanel] = useState<boolean>(false);
  const [selectedColToTransfer, setSelectedColToTransfer] = useState<string>("");
  const [targetColName, setTargetColName] = useState<string>("");
  const [transferMode, setTransferMode] = useState<"overwrite" | "empty_only" | "new_col">("new_col");
  const [transferSuccessMsg, setTransferSuccessMsg] = useState<string | null>(null);

  const mainColumns = useMemo(() => {
    if (activeMainData.length === 0) return [];
    return Object.keys(activeMainData[0]);
  }, [activeMainData]);

  // Cột có sẵn tùy theo đối tượng làm sạch
  const availableColumnsForClean = useMemo(() => {
    if (cleanTarget === "A") return mainColumns;
    if (cleanTarget === "B") return compareColumns;
    // Cả 2: lấy tất cả các cột duy nhất
    return Array.from(new Set([...mainColumns, ...compareColumns]));
  }, [cleanTarget, mainColumns, compareColumns]);

  // Tự động tìm kiếm các cột tương đồng / đồng nghĩa (>= 90% hoặc từ điển)
  const detectedMatchingPairs = useMemo(() => {
    if (mainColumns.length === 0 || compareColumns.length === 0) return [];
    return autoDetectMatchingColumns(mainColumns, compareColumns);
  }, [mainColumns, compareColumns]);

  // Đọc file chính A (Kỳ 1)
  const handleFileUploadA = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wsName], { header: 1, defval: "" });
        const parsed = parse2DArrayWithSmartHeader(rawRows);

        setLocalDataA(parsed.data);
        setLocalFileNameA(file.name);
        if (setMainData) setMainData(parsed.data);
        if (setColumns) setColumns(parsed.columns);
        if (setRawImportedData) setRawImportedData(parsed.data);
        if (setFileName) setFileName(file.name);

        // Tự động gán khóa nếu đã có dữ liệu file B
        if (compareColumns.length > 0) {
          const matches = autoDetectMatchingColumns(parsed.columns, compareColumns);
          if (matches.length > 0) {
            setKeyColA(matches[0].col1);
            setKeyColB(matches[0].col2);
            if (matches.length > 1) {
              setCompareValColA(matches[1].col1);
              setCompareValColB(matches[1].col2);
            }
          }
        } else if (parsed.columns.length > 0) {
          setKeyColA(parsed.columns[0]);
          setCompareValColA(parsed.columns[1] || parsed.columns[0]);
        }
      } catch (err: any) {
        alert("Lỗi đọc tệp A: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // Chọn nhanh Sheet làm File A
  const handleSelectSheetA = (sheetName: string) => {
    if (!sheetDataStore[sheetName]) return;
    const sData = sheetDataStore[sheetName].data || [];
    const sCols = sheetDataStore[sheetName].columns || (sData[0] ? Object.keys(sData[0]) : []);
    setLocalDataA(sData);
    setLocalFileNameA(`Sheet: ${sheetName}`);
    if (setMainData) setMainData(sData);
    if (setColumns) setColumns(sCols);
    if (setRawImportedData) setRawImportedData(sData);
    if (setFileName) setFileName(`Sheet: ${sheetName}`);

    if (compareColumns.length > 0) {
      const matches = autoDetectMatchingColumns(sCols, compareColumns);
      if (matches.length > 0) {
        setKeyColA(matches[0].col1);
        setKeyColB(matches[0].col2);
        if (matches.length > 1) {
          setCompareValColA(matches[1].col1);
          setCompareValColB(matches[1].col2);
        }
      }
    } else if (sCols.length > 0) {
      setKeyColA(sCols[0]);
      setCompareValColA(sCols[1] || sCols[0]);
    }
  };

  // Đọc file đối chiếu B
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setCompareFile(file);
    setCompareSheetName(file.name);
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wsName], { header: 1, defval: "" });
        const parsed = parse2DArrayWithSmartHeader(rawRows);
        setCompareData(parsed.data);
        setCompareColumns(parsed.columns);

        // Tự động gán khóa nếu phát hiện cột tương đồng
        const matches = autoDetectMatchingColumns(mainColumns, parsed.columns);
        if (matches.length > 0) {
          setKeyColA(matches[0].col1);
          setKeyColB(matches[0].col2);
          if (matches.length > 1) {
            setCompareValColA(matches[1].col1);
            setCompareValColB(matches[1].col2);
          }
        } else if (parsed.columns.length > 0) {
          setKeyColB(parsed.columns[0]);
          setCompareValColB(parsed.columns[1] || parsed.columns[0]);
        }
      } catch (err: any) {
        alert("Lỗi đọc tệp đối chiếu: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // Hoặc chọn nhanh từ Sheet khác trong cùng Workbook (nếu có)
  const handleSelectSheetFromWorkbook = (sheetName: string) => {
    if (!sheetDataStore[sheetName]) return;
    const sData = sheetDataStore[sheetName].data || [];
    const sCols = sheetDataStore[sheetName].columns || (sData[0] ? Object.keys(sData[0]) : []);
    setCompareData(sData);
    setCompareColumns(sCols);
    setCompareSheetName(`Sheet: ${sheetName}`);
    setCompareFile(null);

    // Tự động gán khóa nếu khớp
    const matches = autoDetectMatchingColumns(mainColumns, sCols);
    if (matches.length > 0) {
      setKeyColA(matches[0].col1);
      setKeyColB(matches[0].col2);
      if (matches.length > 1) {
        setCompareValColA(matches[1].col1);
        setCompareValColB(matches[1].col2);
      }
    } else if (sCols.length > 0) {
      setKeyColB(sCols[0]);
      setCompareValColB(sCols[1] || sCols[0]);
    }
  };

  // Hàm chuẩn hóa khóa
  const normalizeKey = (val: any): string => {
    let k = cleanForComparison(val, {
      ignoreCase,
      ignoreDiacritics,
      ignoreSpaces,
      ignorePunctuation,
    });
    if (trimLeadingZeros) {
      k = k.replace(/^0+/, "");
    }
    return k;
  };

  // =========================================================================
  // LOGIC LÀM SẠCH KHOẢNG CÁCH & ĐỊNH DẠNG CHỮ HOA/THƯỜNG THEO CỘT
  // =========================================================================
  const cleanStringHelper = (val: any): string => {
    if (val === undefined || val === null) return "";
    let s = String(val);

    // 1. Xóa dấu nếu chọn
    if (cleanRemoveAccents) {
      s = removeVietnameseTones(s);
    }

    // 2. Xử lý khoảng cách
    if (cleanRemoveAllSpaces) {
      s = s.replace(/\s+/g, "");
    } else {
      if (cleanCollapseSpaces) {
        s = s.replace(/\s+/g, " ");
      }
      if (cleanTrimSpaces) {
        s = s.trim();
      }
    }

    // 3. Xử lý chữ hoa / chữ thường
    if (cleanCaseOption === "upper") {
      s = s.toUpperCase();
    } else if (cleanCaseOption === "lower") {
      s = s.toLowerCase();
    } else if (cleanCaseOption === "title") {
      // Viết hoa chữ cái đầu mỗi từ
      s = s.toLowerCase().replace(/(^|\s)\S/g, char => char.toUpperCase());
    }

    return s;
  };

  const handleExecuteCleanColumns = () => {
    if (cleanSelectedCols.length === 0) {
      alert("Vui lòng tích chọn ít nhất một cột cần làm sạch khoảng cách và định dạng chữ hoa/thường!");
      return;
    }

    let modifiedCellsCount = 0;
    let modifiedRowsA = 0;
    let modifiedRowsB = 0;

    // Xử lý Kỳ 1 (Tệp A) nếu được chọn
    if ((cleanTarget === "A" || cleanTarget === "BOTH") && activeMainData.length > 0) {
      const updatedA = activeMainData.map(row => {
        let changed = false;
        const newRow = { ...row };
        cleanSelectedCols.forEach(col => {
          if (newRow[col] !== undefined && newRow[col] !== null) {
            const originalVal = String(newRow[col]);
            const cleanedVal = cleanStringHelper(originalVal);
            if (originalVal !== cleanedVal) {
              newRow[col] = cleanedVal;
              modifiedCellsCount++;
              changed = true;
            }
          }
        });
        if (changed) modifiedRowsA++;
        return newRow;
      });
      setLocalDataA(updatedA);
      if (setMainData) setMainData(updatedA);
      if (setRawImportedData) {
        setRawImportedData(updatedA);
      }
    }

    // Xử lý Kỳ 2 (Tệp B) nếu được chọn
    if ((cleanTarget === "B" || cleanTarget === "BOTH") && compareData.length > 0) {
      const updatedB = compareData.map(row => {
        let changed = false;
        const newRow = { ...row };
        cleanSelectedCols.forEach(col => {
          if (newRow[col] !== undefined && newRow[col] !== null) {
            const originalVal = String(newRow[col]);
            const cleanedVal = cleanStringHelper(originalVal);
            if (originalVal !== cleanedVal) {
              newRow[col] = cleanedVal;
              modifiedCellsCount++;
              changed = true;
            }
          }
        });
        if (changed) modifiedRowsB++;
        return newRow;
      });
      setCompareData(updatedB);
    }

    const targetDesc =
      cleanTarget === "BOTH"
        ? "Cả 2 kỳ (A & B)"
        : cleanTarget === "A"
        ? "Kỳ 1 (Tệp A)"
        : "Kỳ 2 (Tệp B)";

    const caseDesc =
      cleanCaseOption === "upper"
        ? "CHỮ HOA"
        : cleanCaseOption === "lower"
        ? "chữ thường"
        : cleanCaseOption === "title"
        ? "Viết Hoa Chữ Đầu"
        : "Giữ nguyên kiểu chữ";

    const msg = `Đã làm sạch khoảng cách và chuyển ${caseDesc} thành công cho ${cleanSelectedCols.length} cột trên ${targetDesc}! (Đã chuẩn hóa ${modifiedCellsCount.toLocaleString("vi-VN")} ô dữ liệu). Bảng so sánh đã tự động cập nhật.`;
    setCleanSuccessMsg(msg);
  };

  // Chọn nhanh các cột văn bản/mã thường gặp
  const handleQuickSelectTextCols = () => {
    const textCols = availableColumnsForClean.filter(c =>
      /mã|ma|tên|ten|họ|ho|mst|cccd|cmnd|id|địa chỉ|dia chi|xã|xa|huyện|huyen|tỉnh|tinh|phòng|ban|loại|loai|nhóm|nhom|số|so/i.test(c)
    );
    if (textCols.length > 0) {
      setCleanSelectedCols(textCols);
    } else {
      setCleanSelectedCols(availableColumnsForClean.slice(0, 5));
    }
  };

  // =========================================================================
  // TÍNH TOÁN KẾT QUẢ SO SÁNH 2 KỲ (CHI TIẾT & TỔNG HỢP)
  // =========================================================================
  const comparisonResults = useMemo(() => {
    if (!keyColA || !keyColB || compareData.length === 0) return null;

    // Index hóa Tệp B bằng Map
    const mapB = new Map<string, any>();
    compareData.forEach(row => {
      const k = normalizeKey(row[keyColB]);
      if (k) {
        if (!mapB.has(k)) {
          mapB.set(k, row);
        }
      }
    });

    const diffs: {
      key: string;
      originalKeyA: string;
      originalKeyB?: string;
      status: "diff" | "only_in_a" | "same";
      valA: any;
      valB: any;
      numA: number;
      numB: number;
      valDiff: number;
      percentChange?: number;
      rowA: any;
      rowB: any;
    }[] = [];

    const matchedKeysInB = new Set<string>();

    activeMainData.forEach(row => {
      const origKeyA = String(row[keyColA] || "");
      const k = normalizeKey(origKeyA);
      if (!k) return;

      const rowB = mapB.get(k);
      if (rowB) {
        matchedKeysInB.add(k);
        const vA = compareValColA ? row[compareValColA] : "";
        const vB = compareValColB ? rowB[compareValColB] : "";

        const numA = parseFloat(String(vA || "").replace(/,/g, "").replace(/\s/g, ""));
        const numB = parseFloat(String(vB || "").replace(/,/g, "").replace(/\s/g, ""));
        const hasNumbers = !isNaN(numA) && !isNaN(numB);

        const valDiff = hasNumbers ? numB - numA : 0;
        const percentChange = hasNumbers && numA !== 0 ? ((numB - numA) / Math.abs(numA)) * 100 : undefined;

        let isSame = false;
        if (!compareValColA && !compareValColB) {
          // Chỉ so khớp mã định danh
          isSame = true;
        } else if (hasNumbers) {
          isSame = Math.abs(numA - numB) < 0.0001;
        } else {
          isSame = String(vA || "").trim().toLowerCase() === String(vB || "").trim().toLowerCase();
        }

        diffs.push({
          key: k,
          originalKeyA: origKeyA,
          originalKeyB: String(rowB[keyColB] || ""),
          status: isSame ? "same" : "diff",
          valA: vA,
          valB: vB,
          numA: isNaN(numA) ? 0 : numA,
          numB: isNaN(numB) ? 0 : numB,
          valDiff,
          percentChange,
          rowA: row,
          rowB: rowB,
        });
      } else {
        const vA = compareValColA ? row[compareValColA] : "";
        const numA = parseFloat(String(vA || "").replace(/,/g, "").replace(/\s/g, ""));
        diffs.push({
          key: k,
          originalKeyA: origKeyA,
          status: "only_in_a",
          valA: vA,
          valB: "(Không có ở Kỳ 2)",
          numA: isNaN(numA) ? 0 : numA,
          numB: 0,
          valDiff: isNaN(numA) ? 0 : -numA,
          rowA: row,
          rowB: null,
        });
      }
    });

    // Các khóa chỉ có ở B (Kỳ 2)
    const onlyInB: {
      key: string;
      originalKeyB: string;
      valB: any;
      numB: number;
      rowB: any;
    }[] = [];

    compareData.forEach(rowB => {
      const origKeyB = String(rowB[keyColB] || "");
      const k = normalizeKey(origKeyB);
      if (k && !matchedKeysInB.has(k)) {
        const vB = compareValColB ? rowB[compareValColB] : "";
        const numB = parseFloat(String(vB || "").replace(/,/g, "").replace(/\s/g, ""));
        onlyInB.push({
          key: k,
          originalKeyB: origKeyB,
          valB: vB,
          numB: isNaN(numB) ? 0 : numB,
          rowB,
        });
      }
    });

    return { diffs, onlyInB };
  }, [
    activeMainData,
    compareData,
    keyColA,
    keyColB,
    compareValColA,
    compareValColB,
    ignoreCase,
    ignoreDiacritics,
    ignoreSpaces,
    ignorePunctuation,
    trimLeadingZeros,
  ]);

  // Danh sách Trùng (Khớp hoàn toàn)
  const matchingList = useMemo(() => {
    if (!comparisonResults) return [];
    return comparisonResults.diffs.filter(d => d.status === "same");
  }, [comparisonResults]);

  // Danh sách Lệch (Gồm: Lệch số liệu, Chỉ có ở Kỳ 1, Chỉ có ở Kỳ 2)
  const discrepancyList = useMemo(() => {
    if (!comparisonResults) return [];
    const diffsInA = comparisonResults.diffs.filter(d => d.status === "diff" || d.status === "only_in_a");
    const formattedOnlyInB = comparisonResults.onlyInB.map(b => ({
      key: b.key,
      originalKeyA: "(Không có ở Kỳ 1)",
      originalKeyB: b.originalKeyB,
      status: "only_in_b" as const,
      valA: "(Không có)",
      valB: b.valB,
      numA: 0,
      numB: b.numB,
      valDiff: b.numB,
      rowA: null,
      rowB: b.rowB,
    }));
    return [...diffsInA, ...formattedOnlyInB];
  }, [comparisonResults]);

  // Toàn bộ danh sách kết hợp
  const allCombinedList = useMemo(() => {
    if (!comparisonResults) return [];
    const formattedOnlyInB = comparisonResults.onlyInB.map(b => ({
      key: b.key,
      originalKeyA: "(Không có ở Kỳ 1)",
      originalKeyB: b.originalKeyB,
      status: "only_in_b" as const,
      valA: "(Không có)",
      valB: b.valB,
      numA: 0,
      numB: b.numB,
      valDiff: b.numB,
      rowA: null,
      rowB: b.rowB,
    }));
    return [...comparisonResults.diffs, ...formattedOnlyInB];
  }, [comparisonResults]);

  // Thống kê & Tổng hợp chuyên sâu
  const stats = useMemo(() => {
    if (!comparisonResults) return null;
    const same = comparisonResults.diffs.filter(d => d.status === "same").length;
    const diff = comparisonResults.diffs.filter(d => d.status === "diff").length;
    const onlyA = comparisonResults.diffs.filter(d => d.status === "only_in_a").length;
    const onlyB = comparisonResults.onlyInB.length;
    const totalA = activeMainData.length;
    const totalB = compareData.length;
    const totalDiff = diff + onlyA + onlyB;
    const totalAll = same + totalDiff;
    const matchRate = totalA > 0 ? (same / totalA) * 100 : 0;

    // Tổng hợp giá trị số liệu nếu có chọn cột
    const hasValCol = Boolean(compareValColA || compareValColB);

    // Tính tổng giá trị Kỳ 1 (trên các dòng chung + chỉ có ở A)
    const sumValA = comparisonResults.diffs.reduce((acc, d) => acc + (d.numA || 0), 0);
    // Tính tổng giá trị Kỳ 2 (trên các dòng chung + chỉ có ở B)
    const sumValB =
      comparisonResults.diffs
        .filter(d => d.status !== "only_in_a")
        .reduce((acc, d) => acc + (d.numB || 0), 0) +
      comparisonResults.onlyInB.reduce((acc, b) => acc + (b.numB || 0), 0);

    const sumDiffVal = sumValB - sumValA;

    // Tổng hợp theo từng nhóm trạng thái
    const sameSumA = comparisonResults.diffs
      .filter(d => d.status === "same")
      .reduce((acc, d) => acc + (d.numA || 0), 0);
    const sameSumB = comparisonResults.diffs
      .filter(d => d.status === "same")
      .reduce((acc, d) => acc + (d.numB || 0), 0);

    const diffSumA = comparisonResults.diffs
      .filter(d => d.status === "diff")
      .reduce((acc, d) => acc + (d.numA || 0), 0);
    const diffSumB = comparisonResults.diffs
      .filter(d => d.status === "diff")
      .reduce((acc, d) => acc + (d.numB || 0), 0);

    const onlyASumA = comparisonResults.diffs
      .filter(d => d.status === "only_in_a")
      .reduce((acc, d) => acc + (d.numA || 0), 0);

    const onlyBSumB = comparisonResults.onlyInB.reduce((acc, b) => acc + (b.numB || 0), 0);

    return {
      same,
      diff,
      onlyA,
      onlyB,
      totalA,
      totalB,
      totalDiff,
      totalAll,
      matchRate,
      hasValCol,
      sumValA,
      sumValB,
      sumDiffVal,
      sameSumA,
      sameSumB,
      diffSumA,
      diffSumB,
      onlyASumA,
      onlyBSumB,
    };
  }, [comparisonResults, activeMainData.length, compareData.length, compareValColA, compareValColB]);

  // Bộ lọc danh sách bảng chi tiết
  const currentDisplayList = useMemo(() => {
    let baseList: any[] = [];
    if (activeResultTab === "same") {
      baseList = matchingList;
    } else if (activeResultTab === "diff") {
      if (diffSubFilter === "diff_val") {
        baseList = discrepancyList.filter(d => d.status === "diff");
      } else if (diffSubFilter === "only_in_a") {
        baseList = discrepancyList.filter(d => d.status === "only_in_a");
      } else if (diffSubFilter === "only_in_b") {
        baseList = discrepancyList.filter(d => d.status === "only_in_b");
      } else {
        baseList = discrepancyList;
      }
    } else if (activeResultTab === "all") {
      baseList = allCombinedList;
    } else {
      // Summary: không hiển thị danh sách này
      return [];
    }

    if (!searchQuery.trim()) return baseList;
    const q = searchQuery.toLowerCase().trim();
    return baseList.filter(item => {
      if (item.key && item.key.toLowerCase().includes(q)) return true;
      if (item.originalKeyA && String(item.originalKeyA).toLowerCase().includes(q)) return true;
      if (item.originalKeyB && String(item.originalKeyB).toLowerCase().includes(q)) return true;
      if (item.valA !== undefined && String(item.valA).toLowerCase().includes(q)) return true;
      if (item.valB !== undefined && String(item.valB).toLowerCase().includes(q)) return true;
      if (item.rowA && Object.values(item.rowA).some(v => v !== null && v !== undefined && String(v).toLowerCase().includes(q))) return true;
      if (item.rowB && Object.values(item.rowB).some(v => v !== null && v !== undefined && String(v).toLowerCase().includes(q))) return true;
      return false;
    });
  }, [activeResultTab, matchingList, discrepancyList, allCombinedList, diffSubFilter, searchQuery]);

  // Tính tổng chân trang (Footer sum) cho danh sách đang hiển thị
  const tableFooterSummary = useMemo(() => {
    if (currentDisplayList.length === 0) return null;
    let sumA = 0;
    let sumB = 0;
    let sumDiff = 0;
    let countNumA = 0;
    let countNumB = 0;

    currentDisplayList.forEach(item => {
      if (item.numA !== undefined && !isNaN(item.numA)) {
        sumA += item.numA;
        countNumA++;
      }
      if (item.numB !== undefined && !isNaN(item.numB)) {
        sumB += item.numB;
        countNumB++;
      }
      if (item.valDiff !== undefined && !isNaN(item.valDiff)) {
        sumDiff += item.valDiff;
      }
    });

    return {
      sumA,
      sumB,
      sumDiff,
      hasNumbers: countNumA > 0 || countNumB > 0,
    };
  }, [currentDisplayList]);

  // =========================================================================
  // XUẤT EXCEL THÔNG MINH (HIỂN THỊ ĐỦ TẤT CẢ CÁC CỘT CỦA CẢ 2 FILE)
  // =========================================================================
  const handleExportMatchingList = () => {
    if (matchingList.length === 0) {
      alert("Không có bản ghi trùng khớp nào để xuất!");
      return;
    }
    const exportRows = matchingList.map((d, i) => {
      const rowA = d.rowA || {};
      const rowB = d.rowB || {};
      const rowObj: Record<string, any> = {
        STT: i + 1,
        "Trạng thái": "Trùng khớp hoàn toàn",
      };

      // Xuất trọn vẹn tất cả các cột của Tệp A (Kỳ 1)
      mainColumns.forEach(col => {
        rowObj[`[Tệp A] ${col}`] = rowA[col] !== undefined ? rowA[col] : "";
      });

      // Xuất trọn vẹn tất cả các cột của Tệp B (Kỳ 2)
      compareColumns.forEach(col => {
        rowObj[`[Tệp B] ${col}`] = rowB[col] !== undefined ? rowB[col] : "";
      });

      if (d.valDiff !== undefined && !isNaN(d.valDiff)) {
        rowObj["Chênh lệch (K2 - K1)"] = d.valDiff;
      }

      return rowObj;
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Danh_Sach_Trung");
    XLSX.writeFile(wb, `Danh_Sach_Trung_Day_Du_Cot_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleExportDiscrepancyList = () => {
    if (discrepancyList.length === 0) {
      alert("Không có bản ghi sai lệch nào để xuất!");
      return;
    }
    const exportRows = discrepancyList.map((d, i) => {
      const rowA = d.rowA || {};
      const rowB = d.rowB || {};
      let statusText = "Sai lệch số liệu";
      if (d.status === "only_in_a") statusText = "Chỉ có ở Kỳ 1 (Thiếu ở Kỳ 2)";
      else if (d.status === "only_in_b") statusText = "Chỉ có ở Kỳ 2 (Mới phát sinh)";

      const rowObj: Record<string, any> = {
        STT: i + 1,
        "Trạng thái lệch": statusText,
      };

      // Xuất trọn vẹn tất cả các cột của Tệp A (Kỳ 1)
      mainColumns.forEach(col => {
        rowObj[`[Tệp A] ${col}`] = rowA[col] !== undefined ? rowA[col] : "";
      });

      // Xuất trọn vẹn tất cả các cột của Tệp B (Kỳ 2)
      compareColumns.forEach(col => {
        rowObj[`[Tệp B] ${col}`] = rowB[col] !== undefined ? rowB[col] : "";
      });

      if (d.valDiff !== undefined && !isNaN(d.valDiff)) {
        rowObj["Chênh lệch (K2 - K1)"] = d.valDiff;
      }

      return rowObj;
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Danh_Sach_Lech");
    XLSX.writeFile(wb, `Danh_Sach_Lech_Day_Du_Cot_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleExportFullSummaryExcel = () => {
    if (!stats || !comparisonResults) return;

    const wb = XLSX.utils.book_new();

    // 1. Sheet Tổng hợp
    const summaryRows = [
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Tổng số bản ghi Kỳ 1 (Tệp A)", "GIÁ TRỊ": stats.totalA, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Tổng số bản ghi Kỳ 2 (Tệp B)", "GIÁ TRỊ": stats.totalB, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Số lượng bản ghi TRÙNG (Khớp cả 2 kỳ)", "GIÁ TRỊ": stats.same, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Tỷ lệ trùng khớp (%)", "GIÁ TRỊ": Number(stats.matchRate.toFixed(2)), "ĐƠN VỊ": "%" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Tổng số bản ghi SAI LỆCH", "GIÁ TRỊ": stats.totalDiff, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "- Trong đó: Trùng khóa nhưng lệch số liệu", "GIÁ TRỊ": stats.diff, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "- Trong đó: Chỉ có ở Kỳ 1 (Thiếu ở Kỳ 2)", "GIÁ TRỊ": stats.onlyA, "ĐƠN VỊ": "Dòng" },
      { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "- Trong đó: Chỉ có ở Kỳ 2 (Mới phát sinh)", "GIÁ TRỊ": stats.onlyB, "ĐƠN VỊ": "Dòng" },
    ];

    if (stats.hasValCol) {
      summaryRows.push(
        { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": `Tổng giá trị Kỳ 1 (${compareValColA})`, "GIÁ TRỊ": stats.sumValA, "ĐƠN VỊ": "Số liệu" },
        { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": `Tổng giá trị Kỳ 2 (${compareValColB})`, "GIÁ TRỊ": stats.sumValB, "ĐƠN VỊ": "Số liệu" },
        { "CHỈ TIÊU ĐỐI SOÁT 2 KỲ": "Tổng chênh lệch giá trị (Kỳ 2 - Kỳ 1)", "GIÁ TRỊ": stats.sumDiffVal, "ĐƠN VỊ": "Số liệu" }
      );
    }

    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Tong_Hop");

    // 2. Sheet Danh Sách Trùng (Đầy đủ tất cả các cột của cả 2 file)
    const wsMatching = XLSX.utils.json_to_sheet(
      matchingList.map((d, i) => {
        const rowA = d.rowA || {};
        const rowB = d.rowB || {};
        const rowObj: Record<string, any> = {
          STT: i + 1,
          "Trạng thái": "Trùng khớp hoàn toàn",
        };
        mainColumns.forEach(col => {
          rowObj[`[Tệp A] ${col}`] = rowA[col] !== undefined ? rowA[col] : "";
        });
        compareColumns.forEach(col => {
          rowObj[`[Tệp B] ${col}`] = rowB[col] !== undefined ? rowB[col] : "";
        });
        if (d.valDiff !== undefined && !isNaN(d.valDiff)) {
          rowObj["Chênh lệch (K2 - K1)"] = d.valDiff;
        }
        return rowObj;
      })
    );
    XLSX.utils.book_append_sheet(wb, wsMatching, "Danh_Sach_Trung");

    // 3. Sheet Danh Sách Lệch (Đầy đủ tất cả các cột của cả 2 file)
    const wsDiscrepancy = XLSX.utils.json_to_sheet(
      discrepancyList.map((d, i) => {
        const rowA = d.rowA || {};
        const rowB = d.rowB || {};
        let statusText = "Sai lệch số liệu";
        if (d.status === "only_in_a") statusText = "Chỉ có ở Kỳ 1 (Thiếu ở Kỳ 2)";
        else if (d.status === "only_in_b") statusText = "Chỉ có ở Kỳ 2 (Mới phát sinh)";

        const rowObj: Record<string, any> = {
          STT: i + 1,
          "Trạng thái lệch": statusText,
        };
        mainColumns.forEach(col => {
          rowObj[`[Tệp A] ${col}`] = rowA[col] !== undefined ? rowA[col] : "";
        });
        compareColumns.forEach(col => {
          rowObj[`[Tệp B] ${col}`] = rowB[col] !== undefined ? rowB[col] : "";
        });
        if (d.valDiff !== undefined && !isNaN(d.valDiff)) {
          rowObj["Chênh lệch (K2 - K1)"] = d.valDiff;
        }
        return rowObj;
      })
    );
    XLSX.utils.book_append_sheet(wb, wsDiscrepancy, "Danh_Sach_Lech");

    XLSX.writeFile(wb, `BaoCao_SoSanh_Day_Du_Cot_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Tính năng chuyển cột từ Tệp B sang Tệp A (VLOOKUP Transfer)
  const handleTransferColumnToA = () => {
    if (!keyColA || !keyColB || !selectedColToTransfer || !setMainData) {
      alert("Vui lòng chọn đủ khóa 2 tệp và cột cần chuyển từ Tệp B!");
      return;
    }

    const colNameFinal = targetColName.trim() || selectedColToTransfer;
    const mapB = new Map<string, any>();
    compareData.forEach(row => {
      const k = normalizeKey(row[keyColB]);
      if (k && !mapB.has(k)) {
        mapB.set(k, row);
      }
    });

    let fillCount = 0;
    const updatedMainData = activeMainData.map(row => {
      const origKeyA = String(row[keyColA] || "");
      const k = normalizeKey(origKeyA);
      const rowB = k ? mapB.get(k) : null;
      const valFromB = rowB ? rowB[selectedColToTransfer] : "";

      const updatedRow = { ...row };
      if (transferMode === "empty_only") {
        if (
          updatedRow[colNameFinal] === undefined ||
          updatedRow[colNameFinal] === null ||
          String(updatedRow[colNameFinal]).trim() === ""
        ) {
          if (valFromB !== undefined && valFromB !== "") {
            updatedRow[colNameFinal] = valFromB;
            fillCount++;
          }
        }
      } else {
        if (valFromB !== undefined && valFromB !== "") {
          updatedRow[colNameFinal] = valFromB;
          fillCount++;
        } else if (updatedRow[colNameFinal] === undefined) {
          updatedRow[colNameFinal] = "";
        }
      }

      return updatedRow;
    });

    if (setColumns && !mainColumns.includes(colNameFinal)) {
      setColumns([...mainColumns, colNameFinal]);
    }

    setLocalDataA(updatedMainData);
    if (setMainData) setMainData(updatedMainData);
    setTransferSuccessMsg(
      `Đã chuyển thành công cột "${colNameFinal}" từ Tệp B sang Tệp A! Đã điền dữ liệu cho ${fillCount.toLocaleString("vi-VN")} dòng khớp.`
    );
    setShowTransferPanel(false);
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <GitCompare className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                So Sánh 2 Kỳ Dữ Liệu &amp; Đối Soát Số Liệu
              </span>
              {activeMainData.length > 0 && (
                <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                  File A: {activeMainData.length.toLocaleString("vi-VN")} dòng • {mainColumns.length} cột
                  {activeFileName && ` (${activeFileName})`}
                </span>
              )}
              {compareData.length > 0 && (
                <span className="text-[11px] font-mono text-emerald-900 bg-emerald-200/80 px-2 py-0.5 rounded-none font-bold border border-emerald-300">
                  File B: {compareData.length.toLocaleString("vi-VN")} dòng • {compareColumns.length} cột
                  {compareSheetName && ` (${compareSheetName})`}
                </span>
              )}
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5 shrink-0">
          {/* CỤM NÚT TẢI TỆP: TẢI FILE A VÀ TẢI FILE B (VÀO 1 MỐI) */}
          <div className="inline-flex items-center bg-white border border-sky-400 p-0.5 shadow-2xs">
            {/* Nút Tải file A */}
            <label
              className="bg-[#0369a1] hover:bg-[#0284c7] text-white font-bold text-xs px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0"
              title={activeFileName ? `File A: ${activeFileName} (${activeMainData.length.toLocaleString("vi-VN")} dòng). Bấm để tải lại.` : "Tải lên tệp dữ liệu A (Kỳ 1)"}
            >
              <Upload className="w-3.5 h-3.5 text-sky-200" />
              <span>Tải file A</span>
              <input type="file" accept=".xlsx,.xls,.csv,.txt" onChange={handleFileUploadA} className="hidden" />
            </label>

            <span className="w-px h-4 bg-slate-300 mx-0.5" />

            {/* Nút Tải file B */}
            <label
              className="bg-[#047857] hover:bg-[#059669] text-white font-bold text-xs px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0"
              title={compareSheetName || compareFile?.name ? `File B: ${compareSheetName || compareFile?.name} (${compareData.length.toLocaleString("vi-VN")} dòng). Bấm để tải lại.` : "Tải lên tệp dữ liệu B (Kỳ 2)"}
            >
              <Upload className="w-3.5 h-3.5 text-emerald-200" />
              <span>Tải file B</span>
              <input type="file" accept=".xlsx,.xls,.csv,.txt" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>

          {/* Nếu có nhiều sheet trong tệp hiện tại */}
          {detectedSheets.length > 1 && (
            <div className="relative inline-flex items-center">
              <select
                onChange={e => {
                  if (e.target.value) handleSelectSheetFromWorkbook(e.target.value);
                }}
                defaultValue=""
                className="bg-white border border-emerald-500 text-emerald-900 font-bold text-xs px-2 py-1 rounded-none cursor-pointer focus:outline-none focus:ring-1 focus:ring-emerald-500"
                title="Chọn Sheet làm File B"
              >
                <option value="" disabled>
                  Sheet File B...
                </option>
                {detectedSheets.map(s => (
                  <option key={s} value={s}>
                    {s} ({sheetDataStore[s]?.data?.length || 0})
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Nút bật tắt chế độ chuyên biệt Sáp nhập Xã/Thôn */}
          <button
            type="button"
            onClick={() => setShowAdminMergeMode(!showAdminMergeMode)}
            className={`text-xs font-bold px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 shadow-2xs ${
              showAdminMergeMode
                ? "bg-indigo-900 text-white ring-2 ring-indigo-400"
                : "bg-indigo-700 hover:bg-indigo-800 text-white"
            }`}
            title="Giải quyết bài toán sáp nhập xã/thôn: Hộ NN ↔ Cá thể"
          >
            <Building2 className="w-3.5 h-3.5 text-amber-300" />
            <span>Sáp nhập</span>
            <span className="text-[10px]">{showAdminMergeMode ? "▲" : "▼"}</span>
          </button>

          {/* Nút bật tắt công cụ làm sạch */}
          <button
            type="button"
            onClick={() => setShowCleanPanel(!showCleanPanel)}
            className={`text-xs font-bold px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 shadow-2xs ${
              showCleanPanel ? "bg-amber-700 text-white ring-2 ring-amber-400" : "bg-amber-600 hover:bg-amber-700 text-white"
            }`}
            title="Làm sạch khoảng trắng, chữ hoa/thường"
          >
            <Wand2 className="w-3.5 h-3.5" />
            <span>Làm sạch</span>
            <span className="text-[10px]">{showCleanPanel ? "▲" : "▼"}</span>
          </button>

          {comparisonResults && (
            <>
              <button
                type="button"
                onClick={() => setShowTransferPanel(!showTransferPanel)}
                className={`text-xs font-bold px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 shadow-2xs ${
                  showTransferPanel ? "bg-[#4338ca] text-white" : "bg-[#4f46e5] hover:bg-[#4338ca] text-white"
                }`}
                title="Chuyển cột từ Tệp B sang Tệp A"
              >
                <Database className="w-3.5 h-3.5" />
                <span>Lấy cột sang A</span>
                <span className="text-[10px]">{showTransferPanel ? "▲" : "▼"}</span>
              </button>

              <button
                type="button"
                onClick={handleExportFullSummaryExcel}
                className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 shadow-2xs"
                title="Xuất toàn bộ báo cáo Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Xuất Excel</span>
              </button>
            </>
          )}
        </div>
      </div>

      {/* THÔNG BÁO THÀNH CÔNG LÀM SẠCH HOẶC CHUYỂN CỘT */}
      {cleanSuccessMsg && (
        <div className="mx-3.5 p-2.5 bg-emerald-50 border border-emerald-300 rounded-none text-emerald-900 text-xs font-semibold flex items-center justify-between gap-2 animate-fade-in">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>{cleanSuccessMsg}</span>
          </div>
          <button
            type="button"
            onClick={() => setCleanSuccessMsg(null)}
            className="text-emerald-700 hover:text-emerald-900 text-xs font-bold border-0 bg-transparent cursor-pointer"
          >
            ✕ Đóng
          </button>
        </div>
      )}

      {transferSuccessMsg && (
        <div className="mx-3.5 p-2.5 bg-sky-50 border border-sky-300 rounded-none text-sky-900 text-xs font-semibold flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-sky-600 shrink-0" />
            <span>{transferSuccessMsg}</span>
          </div>
          <button
            type="button"
            onClick={() => setTransferSuccessMsg(null)}
            className="text-sky-700 hover:text-sky-900 text-xs font-bold border-0 bg-transparent cursor-pointer"
          >
            ✕ Đóng
          </button>
        </div>
      )}

      {/* ======================================================================= */}
      {/* CÔNG CỤ CHUYÊN BIỆT: GHÉP NỐI & ĐỐI SOÁT SÁP NHẬP XÃ / THÔN */}
      {/* ======================================================================= */}
      {showAdminMergeMode && (
        <div className="mx-3.5 animate-slide-up">
          {compareData.length === 0 ? (
            <div className="p-4 bg-indigo-50 border-2 border-indigo-400 text-indigo-950 text-xs space-y-2.5 shadow-xs">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 font-bold text-sm">
                  <Building2 className="w-5 h-5 text-indigo-700" />
                  <span>CHẾ ĐỘ XỬ LÝ SÁP NHẬP XÃ / THÔN (HỘ ĐIỀU TRA NN ↔ ĐỊA BÀN CÁ THỂ)</span>
                </div>
                <button
                  type="button"
                  onClick={() => setShowAdminMergeMode(false)}
                  className="text-indigo-800 hover:text-black font-bold text-xs bg-indigo-100 hover:bg-indigo-200 px-2 py-0.5 border border-indigo-300 cursor-pointer"
                >
                  ✕ Đóng
                </button>
              </div>
              <p className="text-indigo-900 text-xs leading-relaxed">
                Giải quyết triệt để bài toán: Xã sáp nhập vào nhau sẽ bị mất mã hoặc đổi mã mới, thôn sáp nhập cũng đổi tên hoặc mã địa bàn.
              </p>
              <div className="p-3 bg-white border border-indigo-300 flex items-center justify-between gap-3">
                <div className="text-xs text-slate-700">
                  <b>Bước 1:</b> Vui lòng tải lên <b>Tệp Kỳ 2 (Tệp B)</b> - là danh sách địa bàn hoặc cơ sở kỳ Cá thể năm nay (đã sáp nhập).
                </div>
                <label className="bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs px-3 py-1.5 shadow-2xs flex items-center gap-1.5 cursor-pointer whitespace-nowrap border-0">
                  <Upload className="w-4 h-4" />
                  <span>Tải Lên Tệp Địa Bàn Kỳ 2 (B) Ngay</span>
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="hidden" />
                </label>
              </div>
            </div>
          ) : (
            <AdminMergeResolver
              mainData={activeMainData}
              mainColumns={mainColumns}
              compareData={compareData}
              compareColumns={compareColumns}
              compareFileName={compareFile?.name || compareSheetName || "Tệp Kỳ 2"}
              setMainData={updated => {
                setLocalDataA(updated);
                if (setMainData) setMainData(updated);
              }}
              setColumns={setColumns}
              onClose={() => setShowAdminMergeMode(false)}
            />
          )}
        </div>
      )}

      {/* ======================================================================= */}
      {/* 1. KHU VỰC LÀM SẠCH KHOẢNG CÁCH, CHỮ HOA/CHỮ THƯỜNG THEO CỘT HOẶC NHIỀU CỘT */}
      {/* ======================================================================= */}
      {showCleanPanel && (
        <div className="mx-3.5 bg-amber-50/95 border-2 border-amber-400 p-3 sm:p-4 space-y-3 shadow-xs animate-slide-up">
          <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-amber-300">
            <div className="flex items-center gap-2">
              <div className="w-6 h-6 bg-amber-600 text-white flex items-center justify-center font-bold text-xs">
                <Wand2 className="w-3.5 h-3.5" />
              </div>
              <div>
                <h4 className="text-xs font-bold text-slate-900 uppercase font-mono tracking-wide flex items-center gap-1.5">
                  Làm sạch khoảng cách &amp; Định dạng chữ hoa/chữ thường theo cột
                </h4>
                <p className="text-[11px] text-amber-900">
                  Chọn một hoặc nhiều cột cần chuẩn hóa, chọn định dạng và ấn nút <b>"Thực hiện làm sạch &gt;"</b> để xử lý trực tiếp.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-[11px] font-bold text-amber-950 bg-amber-200/80 px-2 py-0.5 border border-amber-300">
                Đã chọn: {cleanSelectedCols.length} cột
              </span>
              <button
                type="button"
                onClick={() => setShowCleanPanel(false)}
                className="text-xs text-amber-900 hover:text-black font-bold bg-amber-200 hover:bg-amber-300 px-2 py-0.5 border border-amber-400 cursor-pointer"
              >
                ✕ Thu gọn
              </button>
            </div>
          </div>

          {/* Bước 1: Chọn kỳ dữ liệu & Bước 2: Tùy chọn định dạng */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-start">
            {/* Cột trái: Chọn kỳ & Kiểu định dạng */}
            <div className="md:col-span-5 space-y-2.5 bg-white p-2.5 border border-amber-300">
              <div>
                <label className="block text-[11px] font-bold text-slate-800 mb-1">
                  1. Áp dụng cho kỳ dữ liệu nào:
                </label>
                <div className="grid grid-cols-3 gap-1.5 text-xs">
                  <label
                    className={`flex items-center justify-center gap-1 py-1 px-1.5 border cursor-pointer text-center font-bold ${
                      cleanTarget === "A"
                        ? "bg-sky-600 text-white border-sky-700"
                        : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
                    }`}
                  >
                    <input
                      type="radio"
                      name="cleanTarget"
                      checked={cleanTarget === "A"}
                      onChange={() => setCleanTarget("A")}
                      className="hidden"
                    />
                    <span>Kỳ 1 (Tệp A)</span>
                  </label>

                  <label
                    className={`flex items-center justify-center gap-1 py-1 px-1.5 border cursor-pointer text-center font-bold ${
                      cleanTarget === "B"
                        ? "bg-emerald-600 text-white border-emerald-700"
                        : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
                    }`}
                  >
                    <input
                      type="radio"
                      name="cleanTarget"
                      checked={cleanTarget === "B"}
                      onChange={() => setCleanTarget("B")}
                      className="hidden"
                    />
                    <span>Kỳ 2 (Tệp B)</span>
                  </label>

                  <label
                    className={`flex items-center justify-center gap-1 py-1 px-1.5 border cursor-pointer text-center font-bold ${
                      cleanTarget === "BOTH"
                        ? "bg-amber-600 text-white border-amber-700"
                        : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
                    }`}
                  >
                    <input
                      type="radio"
                      name="cleanTarget"
                      checked={cleanTarget === "BOTH"}
                      onChange={() => setCleanTarget("BOTH")}
                      className="hidden"
                    />
                    <span>Cả 2 Kỳ (A &amp; B)</span>
                  </label>
                </div>
              </div>

              {/* Tùy chọn Khoảng cách */}
              <div className="pt-2 border-t border-slate-200 space-y-1.5">
                <span className="block text-[11px] font-bold text-slate-800">
                  2. Làm sạch khoảng cách (Whitespace):
                </span>
                <div className="space-y-1 text-xs text-slate-800">
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={cleanTrimSpaces}
                      onChange={e => setCleanTrimSpaces(e.target.checked)}
                      className="rounded-none text-[#286e42]"
                    />
                    <span>Xóa khoảng trắng thừa ở đầu và cuối (Trim)</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={cleanCollapseSpaces}
                      disabled={cleanRemoveAllSpaces}
                      onChange={e => setCleanCollapseSpaces(e.target.checked)}
                      className="rounded-none text-[#286e42]"
                    />
                    <span>Gom nhiều khoảng trắng kép liên tiếp thành 1</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer text-amber-900 font-medium">
                    <input
                      type="checkbox"
                      checked={cleanRemoveAllSpaces}
                      onChange={e => setCleanRemoveAllSpaces(e.target.checked)}
                      className="rounded-none text-amber-700"
                    />
                    <span>Xóa sạch toàn bộ khoảng trắng (dành cho MST / ID)</span>
                  </label>
                </div>
              </div>

              {/* Tùy chọn Chữ hoa / chữ thường */}
              <div className="pt-2 border-t border-slate-200 space-y-1.5">
                <span className="block text-[11px] font-bold text-slate-800">
                  3. Định dạng kiểu chữ (Case):
                </span>
                <div className="grid grid-cols-2 gap-1 text-xs text-slate-800">
                  <label className="flex items-center gap-1.5 cursor-pointer p-1 bg-slate-50 border border-slate-200">
                    <input
                      type="radio"
                      name="cleanCaseOption"
                      checked={cleanCaseOption === "upper"}
                      onChange={() => setCleanCaseOption("upper")}
                      className="text-[#286e42]"
                    />
                    <span className="font-bold text-emerald-900">CHỮ HOA TOÀN BỘ</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer p-1 bg-slate-50 border border-slate-200">
                    <input
                      type="radio"
                      name="cleanCaseOption"
                      checked={cleanCaseOption === "lower"}
                      onChange={() => setCleanCaseOption("lower")}
                      className="text-[#286e42]"
                    />
                    <span>chữ thường toàn bộ</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer p-1 bg-slate-50 border border-slate-200">
                    <input
                      type="radio"
                      name="cleanCaseOption"
                      checked={cleanCaseOption === "title"}
                      onChange={() => setCleanCaseOption("title")}
                      className="text-[#286e42]"
                    />
                    <span>Viết Hoa Chữ Đầu Từ</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer p-1 bg-slate-50 border border-slate-200">
                    <input
                      type="radio"
                      name="cleanCaseOption"
                      checked={cleanCaseOption === "none"}
                      onChange={() => setCleanCaseOption("none")}
                      className="text-[#286e42]"
                    />
                    <span>Giữ nguyên kiểu chữ</span>
                  </label>
                </div>

                <label className="flex items-center gap-1.5 cursor-pointer text-xs text-slate-700 pt-1">
                  <input
                    type="checkbox"
                    checked={cleanRemoveAccents}
                    onChange={e => setCleanRemoveAccents(e.target.checked)}
                    className="rounded-none text-[#286e42]"
                  />
                  <span>Chuyển thành tiếng Việt KHÔNG DẤU</span>
                </label>
              </div>
            </div>

            {/* Cột phải: Danh sách các cột để chọn & Nút bấm thực hiện > */}
            <div className="md:col-span-7 bg-white p-2.5 border border-amber-300 space-y-2 flex flex-col justify-between h-full">
              <div className="space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-1.5">
                  <span className="text-[11px] font-bold text-slate-800 uppercase">
                    4. Tích chọn các cột cần xử lý:
                  </span>
                  <div className="flex items-center gap-1 text-xs">
                    <button
                      type="button"
                      onClick={() => setCleanSelectedCols([...availableColumnsForClean])}
                      className="px-2 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] border border-slate-300 cursor-pointer"
                    >
                      Chọn tất cả
                    </button>
                    <button
                      type="button"
                      onClick={handleQuickSelectTextCols}
                      className="px-2 py-0.5 bg-amber-100 hover:bg-amber-200 text-amber-900 text-[11px] border border-amber-300 cursor-pointer font-bold"
                      title="Chọn nhanh các cột Mã định danh, MST, Họ tên, Địa chỉ..."
                    >
                      ⚡ Chọn cột Mã/Tên
                    </button>
                    <button
                      type="button"
                      onClick={() => setCleanSelectedCols([])}
                      className="px-2 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] border border-slate-300 cursor-pointer"
                    >
                      Bỏ chọn
                    </button>
                  </div>
                </div>

                {/* Danh sách các cột dạng checkbox grid */}
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 max-h-[160px] overflow-y-auto p-1.5 border border-slate-200 bg-slate-50/60">
                  {availableColumnsForClean.map(col => {
                    const isChecked = cleanSelectedCols.includes(col);
                    return (
                      <label
                        key={col}
                        className={`flex items-center gap-1.5 text-xs px-2 py-1 border cursor-pointer transition-colors select-none ${
                          isChecked
                            ? "bg-emerald-50 border-emerald-500 font-bold text-emerald-950 shadow-2xs"
                            : "bg-white border-slate-200 text-slate-700 hover:bg-slate-100"
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => {
                            if (isChecked) {
                              setCleanSelectedCols(cleanSelectedCols.filter(c => c !== col));
                            } else {
                              setCleanSelectedCols([...cleanSelectedCols, col]);
                            }
                          }}
                          className="rounded-none text-[#286e42]"
                        />
                        <span className="truncate" title={col}>
                          {col}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

              {/* Nút to, rõ ràng ấn nút > để thực hiện */}
              <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
                <div className="text-[11px] text-slate-600">
                  Mẫu thử nghiệm:{" "}
                  <code className="bg-slate-100 px-1 py-0.5 text-slate-800 border">
                    {cleanStringHelper("   hộ   kinh   doanh   nguyễn văn a  ")}
                  </code>
                </div>

                <button
                  type="button"
                  onClick={handleExecuteCleanColumns}
                  disabled={cleanSelectedCols.length === 0}
                  className="px-4 py-2 bg-[#286e42] hover:bg-[#205835] disabled:bg-slate-300 disabled:text-slate-500 text-white font-bold text-xs rounded-none shadow-xs transition-all flex items-center gap-2 cursor-pointer border-0 ring-2 ring-emerald-400"
                >
                  <Wand2 className="w-4 h-4" />
                  <span>LÀM SẠCH KHOẢNG CÁCH &amp; CHỮ HOA/THƯỜNG</span>
                  <ArrowRight className="w-4 h-4 font-extrabold" />
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* GỢI Ý KHÓA TỰ ĐỘNG */}
      {compareData.length > 0 && detectedMatchingPairs.length > 0 && (
        <div className="mx-3.5 px-2.5 py-1.5 bg-amber-50/80 border border-amber-200 rounded-none flex flex-wrap items-center gap-2">
          <span className="text-[11px] font-bold text-amber-900 shrink-0 flex items-center gap-1">
            <Sparkles className="w-3.5 h-3.5 text-amber-600" />
            Gợi ý khóa:
          </span>
          <div className="flex flex-wrap items-center gap-1.5">
            {detectedMatchingPairs.map((pair, pIdx) => (
              <button
                key={pIdx}
                type="button"
                onClick={() => {
                  setKeyColA(pair.col1);
                  setKeyColB(pair.col2);
                }}
                className={`text-[11px] px-2 py-0.5 rounded-none border flex items-center gap-1 transition-colors cursor-pointer ${
                  keyColA === pair.col1 && keyColB === pair.col2
                    ? "bg-[#286e42] text-white border-[#1d4f2f] font-bold shadow-2xs"
                    : "bg-white text-slate-800 border-amber-300 hover:bg-amber-100/70"
                }`}
              >
                <span className="font-semibold text-indigo-700">{pair.col1}</span>
                <span className="text-slate-400">↔</span>
                <span className="font-semibold text-emerald-700">{pair.col2}</span>
                <span className="text-[10px] text-amber-700">({(pair.similarity * 100).toFixed(0)}%)</span>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* CẤU HÌNH KHÓA & BỎ QUA YẾU TỐ */}
      {compareData.length > 0 && (
        <div className="mx-3.5 bg-white border border-sky-200 rounded-none p-3 space-y-2.5">
          <div className="grid grid-cols-1 md:grid-cols-4 gap-2.5">
            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                1. Khóa 1 (Tệp A):
              </label>
              <select
                value={keyColA}
                onChange={e => setKeyColA(e.target.value)}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="">-- Chọn khóa 1 --</option>
                {mainColumns.map(c => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                2. Khóa 2 (Tệp B):
              </label>
              <select
                value={keyColB}
                onChange={e => setKeyColB(e.target.value)}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="">-- Chọn khóa 2 --</option>
                {compareColumns.map(c => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                3. Số liệu 1:
              </label>
              <select
                value={compareValColA}
                onChange={e => setCompareValColA(e.target.value)}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="">-- Tùy chọn (đối soát số liệu) --</option>
                {mainColumns.map(c => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                4. Số liệu 2:
              </label>
              <select
                value={compareValColB}
                onChange={e => setCompareValColB(e.target.value)}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="">-- Tùy chọn (đối soát số liệu) --</option>
                {compareColumns.map(c => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* CÁC TÙY CHỌN BỎ QUA */}
          <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center gap-3 text-xs text-slate-700">
            <span className="font-bold text-slate-900 text-[11px] flex items-center gap-1">
              <Sliders className="w-3.5 h-3.5 text-[#286e42]" /> Bỏ qua:
            </span>
            <label className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={ignoreSpaces}
                onChange={e => setIgnoreSpaces(e.target.checked)}
                className="rounded-none border-slate-300 text-[#286e42] focus:ring-[#286e42]"
              />
              <span>Khoảng trắng</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={ignoreDiacritics}
                onChange={e => setIgnoreDiacritics(e.target.checked)}
                className="rounded-none border-slate-300 text-[#286e42] focus:ring-[#286e42]"
              />
              <span>Dấu tiếng Việt</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={ignorePunctuation}
                onChange={e => setIgnorePunctuation(e.target.checked)}
                className="rounded-none border-slate-300 text-[#286e42] focus:ring-[#286e42]"
              />
              <span>Ký tự đặc biệt</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={ignoreCase}
                onChange={e => setIgnoreCase(e.target.checked)}
                className="rounded-none border-slate-300 text-[#286e42] focus:ring-[#286e42]"
              />
              <span>Hoa/thường</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={trimLeadingZeros}
                onChange={e => setTrimLeadingZeros(e.target.checked)}
                className="rounded-none border-slate-300 text-[#286e42] focus:ring-[#286e42]"
              />
              <span>Số 0 đầu</span>
            </label>
          </div>
        </div>
      )}

      {/* BẢNG ĐIỀU KHIỂN NHẶT / CHUYỂN CỘT TỪ TỆP B SANG TỆP A (SUB-PANEL) */}
      {showTransferPanel && compareData.length > 0 && (
        <div className="mx-3.5 bg-sky-50/70 border border-sky-300 rounded-none p-3.5 space-y-3 animate-slide-up">
          <div className="flex items-center justify-between pb-1.5 border-b border-sky-200">
            <div className="flex items-center gap-2">
              <Database className="w-4 h-4 text-[#0284c7]" />
              <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Lấy Cột Từ Tệp B Chuyển Sang Bảng Chính (Kỳ 1)
              </h4>
            </div>
            <button
              type="button"
              onClick={() => setShowTransferPanel(false)}
              className="text-slate-500 hover:text-slate-800 text-xs font-bold bg-transparent border-0 cursor-pointer"
            >
              ✕ Đóng
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 items-end">
            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                1. Chọn cột từ Tệp B cần lấy sang:
              </label>
              <select
                value={selectedColToTransfer}
                onChange={e => {
                  setSelectedColToTransfer(e.target.value);
                  if (!targetColName) setTargetColName(e.target.value);
                }}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="">-- Chọn cột cần lấy từ B --</option>
                {compareColumns.map(c => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                2. Tên cột mới hiển thị ở Kỳ 1 (Tệp A):
              </label>
              <input
                type="text"
                value={targetColName}
                onChange={e => setTargetColName(e.target.value)}
                placeholder="Nhập tên cột ở Tệp A..."
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
              />
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-700 mb-1">
                3. Quy tắc cập nhật:
              </label>
              <select
                value={transferMode}
                onChange={e => setTransferMode(e.target.value as any)}
                className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
              >
                <option value="new_col">Tạo cột mới / Ghi đè cột cũ</option>
                <option value="empty_only">Chỉ điền vào khi ô của Tệp A còn trống</option>
              </select>
            </div>
          </div>

          <div className="flex items-center justify-between pt-1">
            <span className="text-xs text-slate-600">
              Khóa ghép nối: <strong className="text-indigo-700">{keyColA || "Chưa chọn"}</strong> (A) ↔{" "}
              <strong className="text-emerald-700">{keyColB || "Chưa chọn"}</strong> (B)
            </span>
            <button
              type="button"
              onClick={handleTransferColumnToA}
              disabled={!selectedColToTransfer || !keyColA || !keyColB}
              className="px-3.5 py-1.5 bg-[#059669] hover:bg-[#047857] text-white text-xs font-bold rounded-none shadow-xs transition-colors flex items-center gap-1.5 cursor-pointer disabled:opacity-50 border-0"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Xác nhận chuyển cột này sang Tệp A</span>
            </button>
          </div>
        </div>
      )}

      {/* ======================================================================= */}
      {/* 2. KHU VỰC HIỂN THỊ KẾT QUẢ SO SÁNH: DANH SÁCH TRÙNG, LỆCH & BẢNG TỔNG HỢP */}
      {/* ======================================================================= */}
      {comparisonResults && stats && (
        <div className="mx-3.5 space-y-3">
          {/* HỆ THỐNG THẺ CHỈ SỐ TỔNG HỢP NHANH (CARDS) */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-2.5">
            {/* Card 1: Tổng số bản ghi 2 kỳ */}
            <div className="bg-white p-3 border border-sky-300 rounded-none shadow-2xs">
              <div className="flex items-center justify-between text-slate-500 text-[11px] font-bold uppercase mb-1">
                <span>Tổng bản ghi</span>
                <GitCompare className="w-4 h-4 text-sky-600" />
              </div>
              <div className="flex items-baseline gap-2">
                <span className="text-lg font-bold text-slate-900 font-mono">
                  {stats.totalA.toLocaleString("vi-VN")}
                </span>
                <span className="text-xs text-slate-500">↔</span>
                <span className="text-lg font-bold text-emerald-800 font-mono">
                  {stats.totalB.toLocaleString("vi-VN")}
                </span>
              </div>
              <div className="text-[11px] text-slate-600 mt-1 flex items-center justify-between">
                <span>Kỳ 1: {stats.totalA.toLocaleString("vi-VN")}</span>
                <span>Kỳ 2: {stats.totalB.toLocaleString("vi-VN")}</span>
              </div>
            </div>

            {/* Card 2: Danh sách TRÙNG */}
            <div
              onClick={() => setActiveResultTab("same")}
              className={`p-3 border rounded-none shadow-2xs cursor-pointer transition-all ${
                activeResultTab === "same"
                  ? "bg-emerald-50 border-emerald-600 ring-2 ring-emerald-500"
                  : "bg-white border-emerald-300 hover:bg-emerald-50/50"
              }`}
            >
              <div className="flex items-center justify-between text-emerald-800 text-[11px] font-bold uppercase mb-1">
                <span className="flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" /> Trùng khớp
                </span>
                <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1 py-0.2 font-mono">
                  {stats.matchRate.toFixed(1)}%
                </span>
              </div>
              <div className="text-xl font-bold text-emerald-900 font-mono">
                {stats.same.toLocaleString("vi-VN")}
                <span className="text-xs font-normal text-emerald-700 ml-1">dòng</span>
              </div>
              <div className="text-[11px] text-emerald-700 mt-1">
                Khớp cả 2 kỳ
              </div>
            </div>

            {/* Card 3: Danh sách LỆCH */}
            <div
              onClick={() => {
                setActiveResultTab("diff");
                setDiffSubFilter("all");
              }}
              className={`p-3 border rounded-none shadow-2xs cursor-pointer transition-all ${
                activeResultTab === "diff"
                  ? "bg-amber-50 border-amber-600 ring-2 ring-amber-500"
                  : "bg-white border-amber-300 hover:bg-amber-50/50"
              }`}
            >
              <div className="flex items-center justify-between text-amber-900 text-[11px] font-bold uppercase mb-1">
                <span className="flex items-center gap-1">
                  <AlertTriangle className="w-3.5 h-3.5 text-amber-600" /> Sai lệch
                </span>
                <span className="text-[10px] bg-amber-100 text-amber-900 px-1 py-0.2 font-mono">
                  {(100 - stats.matchRate).toFixed(1)}%
                </span>
              </div>
              <div className="text-xl font-bold text-amber-900 font-mono">
                {stats.totalDiff.toLocaleString("vi-VN")}
                <span className="text-xs font-normal text-amber-800 ml-1">dòng</span>
              </div>
              <div className="text-[11px] text-amber-800 mt-1 flex items-center justify-between">
                <span>Lệch số: {stats.diff}</span>
                <span>Chỉ ở A: {stats.onlyA}</span>
                <span>Chỉ ở B: {stats.onlyB}</span>
              </div>
            </div>

            {/* Card 4: Tổng giá trị số liệu 2 kỳ */}
            <div className="bg-white p-3 border border-indigo-300 rounded-none shadow-2xs">
              <div className="flex items-center justify-between text-indigo-900 text-[11px] font-bold uppercase mb-1">
                <span>Chênh lệch số liệu</span>
                {stats.sumDiffVal >= 0 ? (
                  <TrendingUp className="w-4 h-4 text-emerald-600" />
                ) : (
                  <TrendingDown className="w-4 h-4 text-rose-600" />
                )}
              </div>
              <div className="text-lg font-bold font-mono text-slate-900">
                {stats.hasValCol ? (
                  <span className={stats.sumDiffVal >= 0 ? "text-emerald-700" : "text-rose-700"}>
                    {stats.sumDiffVal >= 0 ? "+" : ""}
                    {stats.sumDiffVal.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                  </span>
                ) : (
                  <span className="text-xs text-slate-400 font-normal">Chưa chọn cột số liệu</span>
                )}
              </div>
              {stats.hasValCol && (
                <div className="text-[11px] text-slate-600 mt-1 flex items-center justify-between truncate font-mono">
                  <span>K1: {stats.sumValA.toLocaleString("vi-VN", { maximumFractionDigits: 1 })}</span>
                  <span>K2: {stats.sumValB.toLocaleString("vi-VN", { maximumFractionDigits: 1 })}</span>
                </div>
              )}
            </div>
          </div>

          {/* DẢI TABS ĐIỀU HƯỚNG XEM KẾT QUẢ & CÁC NÚT XUẤT EXCEL */}
          <div className="bg-white p-2 border border-sky-300 flex flex-wrap items-center justify-between gap-2.5">
            <div className="flex flex-wrap items-center gap-1.5">
              {/* Tab 1: Tổng hợp */}
              <button
                type="button"
                onClick={() => setActiveResultTab("summary")}
                className={`px-3 py-1.5 text-xs font-bold rounded-none border transition-colors flex items-center gap-1.5 cursor-pointer ${
                  activeResultTab === "summary"
                    ? "bg-[#286e42] text-white border-[#1d4f2f] shadow-2xs ring-1 ring-emerald-600"
                    : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
                }`}
              >
                <BarChart3 className="w-3.5 h-3.5" />
                <span>📊 TỔNG HỢP</span>
              </button>

              {/* Tab 2: Danh sách TRÙNG */}
              <button
                type="button"
                onClick={() => setActiveResultTab("same")}
                className={`px-3 py-1.5 text-xs font-bold rounded-none border transition-colors flex items-center gap-1.5 cursor-pointer ${
                  activeResultTab === "same"
                    ? "bg-[#059669] text-white border-[#047857] shadow-2xs ring-1 ring-emerald-600"
                    : "bg-emerald-50 text-emerald-900 border-emerald-300 hover:bg-emerald-100"
                }`}
              >
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-300" />
                <span>✅ TRÙNG ({stats.same.toLocaleString("vi-VN")})</span>
              </button>

              {/* Tab 3: Danh sách LỆCH */}
              <button
                type="button"
                onClick={() => {
                  setActiveResultTab("diff");
                  setDiffSubFilter("all");
                }}
                className={`px-3 py-1.5 text-xs font-bold rounded-none border transition-colors flex items-center gap-1.5 cursor-pointer ${
                  activeResultTab === "diff"
                    ? "bg-[#ea580c] text-white border-[#c2410c] shadow-2xs ring-1 ring-amber-600"
                    : "bg-amber-50 text-amber-900 border-amber-300 hover:bg-amber-100"
                }`}
              >
                <AlertTriangle className="w-3.5 h-3.5 text-amber-300" />
                <span>⚠️ LỆCH ({stats.totalDiff.toLocaleString("vi-VN")})</span>
              </button>

              {/* Tab 4: Toàn bộ danh sách */}
              <button
                type="button"
                onClick={() => setActiveResultTab("all")}
                className={`px-3 py-1.5 text-xs font-bold rounded-none border transition-colors flex items-center gap-1.5 cursor-pointer ${
                  activeResultTab === "all"
                    ? "bg-[#0284c7] text-white border-[#0369a1] shadow-2xs"
                    : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
                }`}
              >
                <span>📑 TOÀN BỘ ({stats.totalAll.toLocaleString("vi-VN")})</span>
              </button>
            </div>

            {/* Các nút xuất danh sách riêng biệt */}
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={handleExportMatchingList}
                className="px-2.5 py-1 text-xs font-bold bg-[#059669] hover:bg-[#047857] text-white rounded-none cursor-pointer border-0 shadow-2xs flex items-center gap-1"
                title="Tải riêng danh sách các bản ghi trùng khớp ra Excel"
              >
                <Download className="w-3 h-3" />
                <span>Xuất trùng</span>
              </button>

              <button
                type="button"
                onClick={handleExportDiscrepancyList}
                className="px-2.5 py-1 text-xs font-bold bg-[#ea580c] hover:bg-[#c2410c] text-white rounded-none cursor-pointer border-0 shadow-2xs flex items-center gap-1"
                title="Tải riêng danh sách các bản ghi sai lệch ra Excel"
              >
                <Download className="w-3 h-3" />
                <span>Xuất lệch</span>
              </button>
            </div>
          </div>

          {/* =================================================================== */}
          {/* NỘI DUNG TAB 1: BẢNG TỔNG HỢP CHI TIẾT ĐỐI SOÁT 2 KỲ */}
          {/* =================================================================== */}
          {activeResultTab === "summary" && (
            <div className="space-y-3 animate-fade-in">
              {/* Bảng phân tích chi tiết từng nhóm */}
              <div className="bg-white border border-sky-300 rounded-none overflow-hidden shadow-xs">
                <div className="p-3 bg-sky-100/70 border-b border-sky-200 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <BarChart3 className="w-4 h-4 text-[#286e42]" />
                    <h4 className="text-xs font-bold text-slate-900 uppercase">
                      Bảng Tổng Hợp Phân Tích Đối Soát Số Liệu Giữa 2 Kỳ
                    </h4>
                  </div>
                  <span className="text-[11px] text-slate-600 font-medium">
                    Khóa liên kết: <b className="font-mono text-indigo-700">{keyColA}</b> (Kỳ 1) ↔{" "}
                    <b className="font-mono text-emerald-700">{keyColB}</b> (Kỳ 2)
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-[#286e42] text-white font-bold border-b border-[#1d4f2f]">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] w-12 text-center">STT</th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f]">NHÓM ĐỐI SOÁT</th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] text-right">SỐ BẢN GHI</th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] text-right">TỶ LỆ (%)</th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] text-right">
                          TỔNG GIÁ TRỊ KỲ 1 ({compareValColA || "Kỳ 1"})
                        </th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] text-right">
                          TỔNG GIÁ TRỊ KỲ 2 ({compareValColB || "Kỳ 2"})
                        </th>
                        <th className="py-2.5 px-3 border-r border-[#1d4f2f] text-right">CHÊNH LỆCH (K2 - K1)</th>
                        <th className="py-2.5 px-3 text-center">XEM CHI TIẾT</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {/* Hàng 1: Trùng khớp */}
                      <tr className="hover:bg-emerald-50/60 bg-emerald-50/20">
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-emerald-900 border-r border-slate-100">
                          1
                        </td>
                        <td className="py-2.5 px-3 font-bold text-emerald-900 border-r border-slate-100 flex items-center gap-2">
                          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                          <div>
                            <div>🟢 Trùng khớp hoàn toàn (Khớp mã &amp; số liệu)</div>
                            <div className="text-[10px] font-normal text-slate-500">
                              Bản ghi có mặt ở cả 2 kỳ và chỉ tiêu khớp hoàn toàn
                            </div>
                          </div>
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-emerald-900 border-r border-slate-100">
                          {stats.same.toLocaleString("vi-VN")}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-semibold text-emerald-800 border-r border-slate-100">
                          {stats.matchRate.toFixed(1)}%
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.sameSumA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.sameSumB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-emerald-700 border-r border-slate-100">
                          0
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <button
                            type="button"
                            onClick={() => setActiveResultTab("same")}
                            className="text-[11px] px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-none cursor-pointer border-0 shadow-2xs"
                          >
                            Xem {stats.same} dòng &gt;
                          </button>
                        </td>
                      </tr>

                      {/* Hàng 2: Lệch số liệu */}
                      <tr className="hover:bg-amber-50/60 bg-amber-50/20">
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-amber-900 border-r border-slate-100">
                          2
                        </td>
                        <td className="py-2.5 px-3 font-bold text-amber-900 border-r border-slate-100 flex items-center gap-2">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                          <div>
                            <div>🟠 Trùng khóa nhưng sai lệch số liệu</div>
                            <div className="text-[10px] font-normal text-slate-500">
                              Cùng mã định danh nhưng giá trị chỉ tiêu so sánh khác nhau
                            </div>
                          </div>
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-amber-900 border-r border-slate-100">
                          {stats.diff.toLocaleString("vi-VN")}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-semibold text-amber-800 border-r border-slate-100">
                          {stats.totalAll > 0 ? ((stats.diff / stats.totalAll) * 100).toFixed(1) : 0}%
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.diffSumA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.diffSumB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold border-r border-slate-100">
                          <span
                            className={
                              stats.diffSumB - stats.diffSumA >= 0 ? "text-emerald-700" : "text-rose-700"
                            }
                          >
                            {stats.diffSumB - stats.diffSumA >= 0 ? "+" : ""}
                            {(stats.diffSumB - stats.diffSumA).toLocaleString("vi-VN", {
                              maximumFractionDigits: 2,
                            })}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <button
                            type="button"
                            onClick={() => {
                              setActiveResultTab("diff");
                              setDiffSubFilter("diff_val");
                            }}
                            className="text-[11px] px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-none cursor-pointer border-0 shadow-2xs"
                          >
                            Xem {stats.diff} dòng &gt;
                          </button>
                        </td>
                      </tr>

                      {/* Hàng 3: Chỉ có ở Kỳ 1 */}
                      <tr className="hover:bg-rose-50/60 bg-rose-50/20">
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-rose-900 border-r border-slate-100">
                          3
                        </td>
                        <td className="py-2.5 px-3 font-bold text-rose-900 border-r border-slate-100 flex items-center gap-2">
                          <Minus className="w-4 h-4 text-rose-600 shrink-0" />
                          <div>
                            <div>🔴 Chỉ có ở Kỳ 1 (Thiếu ở Kỳ 2)</div>
                            <div className="text-[10px] font-normal text-slate-500">
                              Bản ghi có ở Kỳ 1 nhưng không tìm thấy trong Kỳ 2
                            </div>
                          </div>
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-rose-900 border-r border-slate-100">
                          {stats.onlyA.toLocaleString("vi-VN")}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-semibold text-rose-800 border-r border-slate-100">
                          {stats.totalAll > 0 ? ((stats.onlyA / stats.totalAll) * 100).toFixed(1) : 0}%
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.onlyASumA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-400 border-r border-slate-100">
                          -
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-rose-700 border-r border-slate-100">
                          -{stats.onlyASumA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <button
                            type="button"
                            onClick={() => {
                              setActiveResultTab("diff");
                              setDiffSubFilter("only_in_a");
                            }}
                            className="text-[11px] px-2.5 py-1 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-none cursor-pointer border-0 shadow-2xs"
                          >
                            Xem {stats.onlyA} dòng &gt;
                          </button>
                        </td>
                      </tr>

                      {/* Hàng 4: Chỉ có ở Kỳ 2 */}
                      <tr className="hover:bg-indigo-50/60 bg-indigo-50/20">
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-indigo-900 border-r border-slate-100">
                          4
                        </td>
                        <td className="py-2.5 px-3 font-bold text-indigo-900 border-r border-slate-100 flex items-center gap-2">
                          <Sparkles className="w-4 h-4 text-indigo-600 shrink-0" />
                          <div>
                            <div>🔵 Chỉ có ở Kỳ 2 (Mới phát sinh)</div>
                            <div className="text-[10px] font-normal text-slate-500">
                              Bản ghi mới xuất hiện trong Kỳ 2 mà Kỳ 1 chưa có
                            </div>
                          </div>
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-indigo-900 border-r border-slate-100">
                          {stats.onlyB.toLocaleString("vi-VN")}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-semibold text-indigo-800 border-r border-slate-100">
                          {stats.totalAll > 0 ? ((stats.onlyB / stats.totalAll) * 100).toFixed(1) : 0}%
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-400 border-r border-slate-100">
                          -
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                          {stats.onlyBSumB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono font-bold text-emerald-700 border-r border-slate-100">
                          +{stats.onlyBSumB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <button
                            type="button"
                            onClick={() => {
                              setActiveResultTab("diff");
                              setDiffSubFilter("only_in_b");
                            }}
                            className="text-[11px] px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-none cursor-pointer border-0 shadow-2xs"
                          >
                            Xem {stats.onlyB} dòng &gt;
                          </button>
                        </td>
                      </tr>
                    </tbody>

                    {/* DÒNG TỔNG CỘNG FOOTER */}
                    <tfoot className="bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900">
                      <tr>
                        <td colSpan={2} className="py-3 px-3 text-center uppercase tracking-wide border-r border-slate-300">
                          TỔNG CỘNG TOÀN BỘ ĐỐI SOÁT
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-sm border-r border-slate-300">
                          {stats.totalAll.toLocaleString("vi-VN")}
                        </td>
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-300">100.0%</td>
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-300">
                          {stats.sumValA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-300">
                          {stats.sumValB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-sm border-r border-slate-300">
                          <span className={stats.sumDiffVal >= 0 ? "text-emerald-700" : "text-rose-700"}>
                            {stats.sumDiffVal >= 0 ? "+" : ""}
                            {stats.sumDiffVal.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                          </span>
                        </td>
                        <td className="py-3 px-3 text-center">
                          <button
                            type="button"
                            onClick={handleExportFullSummaryExcel}
                            className="text-[11px] px-2.5 py-1 bg-[#286e42] hover:bg-[#205835] text-white font-bold rounded-none cursor-pointer border-0 shadow-2xs flex items-center justify-center gap-1 mx-auto"
                          >
                            <Download className="w-3 h-3" />
                            <span>Xuất Excel</span>
                          </button>
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* =================================================================== */}
          {/* NỘI DUNG TAB 2, 3, 4: BẢNG CHI TIẾT DANH SÁCH TRÙNG / LỆCH / TOÀN BỘ */}
          {/* =================================================================== */}
          {activeResultTab !== "summary" && (
            <div className="space-y-2 animate-fade-in">
              {/* Bộ lọc con cho Danh Sách Lệch */}
              {activeResultTab === "diff" && (
                <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50/80 p-2 border border-amber-300">
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="text-xs font-bold text-amber-950 uppercase mr-1 flex items-center gap-1">
                      <Filter className="w-3.5 h-3.5" /> Lọc loại lệch:
                    </span>
                    <button
                      type="button"
                      onClick={() => setDiffSubFilter("all")}
                      className={`px-2.5 py-1 text-xs font-bold rounded-none border cursor-pointer ${
                        diffSubFilter === "all"
                          ? "bg-amber-600 text-white border-amber-700"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-amber-100/50"
                      }`}
                    >
                      Tất cả lệch ({stats.totalDiff})
                    </button>
                    <button
                      type="button"
                      onClick={() => setDiffSubFilter("diff_val")}
                      className={`px-2.5 py-1 text-xs font-bold rounded-none border cursor-pointer ${
                        diffSubFilter === "diff_val"
                          ? "bg-amber-600 text-white border-amber-700"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-amber-100/50"
                      }`}
                    >
                      Lệch số liệu ({stats.diff})
                    </button>
                    <button
                      type="button"
                      onClick={() => setDiffSubFilter("only_in_a")}
                      className={`px-2.5 py-1 text-xs font-bold rounded-none border cursor-pointer ${
                        diffSubFilter === "only_in_a"
                          ? "bg-rose-600 text-white border-rose-700"
                          : "bg-white text-rose-800 border-rose-200 hover:bg-rose-50"
                      }`}
                    >
                      Chỉ có ở Kỳ 1 ({stats.onlyA})
                    </button>
                    <button
                      type="button"
                      onClick={() => setDiffSubFilter("only_in_b")}
                      className={`px-2.5 py-1 text-xs font-bold rounded-none border cursor-pointer ${
                        diffSubFilter === "only_in_b"
                          ? "bg-indigo-600 text-white border-indigo-700"
                          : "bg-white text-indigo-800 border-indigo-200 hover:bg-indigo-50"
                      }`}
                    >
                      Chỉ có ở Kỳ 2 ({stats.onlyB})
                    </button>
                  </div>

                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={e => setSearchQuery(e.target.value)}
                      placeholder="Tìm mã hoặc số liệu..."
                      className="pl-7 pr-2.5 py-1 text-xs border border-slate-300 rounded-none bg-white w-48 font-medium focus:outline-none focus:border-amber-500"
                    />
                  </div>
                </div>
              )}

              {/* Ô tìm kiếm cho Tab Trùng hoặc Toàn Bộ */}
              {activeResultTab !== "diff" && (
                <div className="flex items-center justify-between gap-2 bg-slate-50 p-2 border border-slate-200">
                  <span className="text-xs font-bold text-slate-700">
                    {activeResultTab === "same" ? "DANH SÁCH BẢN GHI TRÙNG KHỚP" : "TOÀN BỘ DỮ LIỆU ĐỐI SOÁT"} (
                    {currentDisplayList.length.toLocaleString("vi-VN")} bản ghi)
                  </span>
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={e => setSearchQuery(e.target.value)}
                      placeholder="Tìm kiếm mã hoặc giá trị..."
                      className="pl-7 pr-2.5 py-1 text-xs border border-slate-300 rounded-none bg-white w-56 font-medium focus:outline-none focus:border-sky-500"
                    />
                  </div>
                </div>
              )}

              {/* BẢNG CHI TIẾT HIỂN THỊ TẤT CẢ CÁC CỘT CỦA CẢ 2 TỆP */}
              <div className="border border-sky-300 rounded-none overflow-x-auto max-h-[540px] overflow-y-auto bg-white shadow-inner">
                <table className="w-full text-left border-collapse text-xs min-w-max">
                  <thead className="sticky top-0 z-10 shadow-xs">
                    {/* HÀNG TIÊU ĐỀ 1: PHÂN NHÓM TỆP A & TỆP B */}
                    <tr className="text-white text-[11px] font-bold uppercase tracking-wider">
                      <th
                        rowSpan={2}
                        className="bg-[#286e42] py-2 px-2.5 w-10 text-center border-r border-[#1d4f2f] sticky left-0 z-20"
                      >
                        #
                      </th>
                      <th
                        rowSpan={2}
                        className="bg-[#286e42] py-2 px-3 border-r border-[#1d4f2f] whitespace-nowrap sticky left-10 z-20"
                      >
                        Trạng thái
                      </th>
                      {stats?.hasValCol && (
                        <th
                          rowSpan={2}
                          className="bg-[#1f5c35] py-2 px-3 border-r border-[#1d4f2f] text-right whitespace-nowrap"
                        >
                          Chênh lệch (K2 - K1)
                        </th>
                      )}
                      <th
                        colSpan={mainColumns.length}
                        className="bg-[#0369a1] py-1.5 px-3 text-center border-r border-[#0284c7] font-bold"
                      >
                        Tệp A (Kỳ 1) • {mainColumns.length} cột
                      </th>
                      <th
                        colSpan={compareColumns.length}
                        className="bg-[#047857] py-1.5 px-3 text-center font-bold"
                      >
                        Tệp B (Kỳ 2) • {compareColumns.length} cột
                      </th>
                    </tr>

                    {/* HÀNG TIÊU ĐỀ 2: TÊN CÁC CỘT CỦA TỆP A VÀ TỆP B */}
                    <tr className="text-[11px] font-semibold">
                      {mainColumns.map(colA => (
                        <th
                          key={`th-a-${colA}`}
                          className={`py-1.5 px-2.5 border-r border-sky-600 whitespace-nowrap ${
                            colA === keyColA
                              ? "bg-sky-900 text-amber-300 font-bold"
                              : "bg-[#0284c7] text-white"
                          }`}
                          title={`Tệp A: ${colA}${colA === keyColA ? " (Khóa 1)" : ""}`}
                        >
                          <span className="opacity-75 mr-1 font-mono text-[10px]">[A]</span>
                          {colA}
                          {colA === keyColA && <span className="ml-1 text-amber-300">🔑</span>}
                        </th>
                      ))}
                      {compareColumns.map(colB => (
                        <th
                          key={`th-b-${colB}`}
                          className={`py-1.5 px-2.5 border-r border-emerald-600 whitespace-nowrap ${
                            colB === keyColB
                              ? "bg-emerald-950 text-amber-300 font-bold"
                              : "bg-[#059669] text-white"
                          }`}
                          title={`Tệp B: ${colB}${colB === keyColB ? " (Khóa 2)" : ""}`}
                        >
                          <span className="opacity-75 mr-1 font-mono text-[10px]">[B]</span>
                          {colB}
                          {colB === keyColB && <span className="ml-1 text-amber-300">🔑</span>}
                        </th>
                      ))}
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-200">
                    {currentDisplayList.length === 0 ? (
                      <tr>
                        <td
                          colSpan={mainColumns.length + compareColumns.length + (stats?.hasValCol ? 3 : 2)}
                          className="py-8 text-center text-slate-400 italic"
                        >
                          Không có bản ghi nào khớp với điều kiện lọc hiện tại.
                        </td>
                      </tr>
                    ) : (
                      currentDisplayList.slice(0, 300).map((d: any, idx: number) => {
                        const rowA = d.rowA;
                        const rowB = d.rowB;
                        return (
                          <tr
                            key={idx}
                            className={`hover:bg-amber-50/50 transition-colors ${
                              d.status === "diff"
                                ? "bg-amber-50/40"
                                : d.status === "only_in_a"
                                ? "bg-rose-50/30"
                                : d.status === "only_in_b"
                                ? "bg-indigo-50/30"
                                : idx % 2 === 1
                                ? "bg-slate-50/50"
                                : "bg-white"
                            }`}
                          >
                            {/* Cột 1: STT */}
                            <td className="py-1 px-2.5 text-center text-slate-500 font-mono text-[11px] border-r border-slate-200 sticky left-0 bg-inherit z-5">
                              {idx + 1}
                            </td>

                            {/* Cột 2: Trạng thái */}
                            <td className="py-1 px-2.5 border-r border-slate-200 text-xs font-medium whitespace-nowrap sticky left-10 bg-inherit z-5">
                              {d.status === "same" ? (
                                <span className="text-emerald-700 font-bold flex items-center gap-1">
                                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" /> Trùng khớp
                                </span>
                              ) : d.status === "diff" ? (
                                <span className="text-amber-800 font-bold flex items-center gap-1">
                                  <AlertTriangle className="w-3.5 h-3.5 text-amber-600" /> Lệch số
                                </span>
                              ) : d.status === "only_in_a" ? (
                                <span className="text-rose-700 font-bold flex items-center gap-1">
                                  <Minus className="w-3 h-3 text-rose-600" /> Chỉ ở A
                                </span>
                              ) : (
                                <span className="text-indigo-700 font-bold flex items-center gap-1">
                                  <Sparkles className="w-3 h-3 text-indigo-600" /> Chỉ ở B
                                </span>
                              )}
                            </td>

                            {/* Cột Chênh lệch nếu có */}
                            {stats?.hasValCol && (
                              <td className="py-1 px-2.5 text-right font-mono font-bold border-r border-slate-200 whitespace-nowrap">
                                {d.valDiff !== undefined && !isNaN(d.valDiff) ? (
                                  <span
                                    className={
                                      d.valDiff > 0
                                        ? "text-emerald-700"
                                        : d.valDiff < 0
                                        ? "text-rose-700"
                                        : "text-slate-600"
                                    }
                                  >
                                    {d.valDiff > 0 ? "+" : ""}
                                    {d.valDiff.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                                  </span>
                                ) : (
                                  "-"
                                )}
                              </td>
                            )}

                            {/* TẤT CẢ CÁC CỘT CỦA TỆP A */}
                            {mainColumns.map(colA => {
                              const val = rowA ? rowA[colA] : undefined;
                              return (
                                <td
                                  key={`td-a-${colA}`}
                                  className={`py-1 px-2.5 border-r border-slate-200 font-mono text-[11px] whitespace-nowrap ${
                                    colA === keyColA ? "font-bold text-sky-950 bg-sky-100/40" : "text-slate-700"
                                  }`}
                                >
                                  {val !== undefined && val !== null && val !== "" ? String(val) : "-"}
                                </td>
                              );
                            })}

                            {/* TẤT CẢ CÁC CỘT CỦA TỆP B */}
                            {compareColumns.map(colB => {
                              const val = rowB ? rowB[colB] : undefined;
                              return (
                                <td
                                  key={`td-b-${colB}`}
                                  className={`py-1 px-2.5 border-r border-slate-200 font-mono text-[11px] whitespace-nowrap ${
                                    colB === keyColB ? "font-bold text-emerald-950 bg-emerald-100/40" : "text-slate-700"
                                  }`}
                                >
                                  {val !== undefined && val !== null && val !== "" ? String(val) : "-"}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })
                    )}
                  </tbody>

                  {/* DÒNG TỔNG CỘNG CHO DANH SÁCH CHI TIẾT */}
                  {tableFooterSummary && tableFooterSummary.hasNumbers && (
                    <tfoot className="bg-slate-100 font-bold border-t-2 border-slate-300 sticky bottom-0 text-slate-900 z-10">
                      <tr>
                        <td
                          colSpan={stats?.hasValCol ? 2 : 2}
                          className="py-2.5 px-3 text-center uppercase tracking-wide border-r border-slate-300"
                        >
                          TỔNG CỘNG ({currentDisplayList.length.toLocaleString("vi-VN")} DÒNG)
                        </td>
                        {stats?.hasValCol && (
                          <td className="py-2.5 px-3 text-right font-mono border-r border-slate-300">
                            <span className={tableFooterSummary.sumDiff >= 0 ? "text-emerald-700" : "text-rose-700"}>
                              {tableFooterSummary.sumDiff >= 0 ? "+" : ""}
                              {tableFooterSummary.sumDiff.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}
                            </span>
                          </td>
                        )}
                        <td
                          colSpan={mainColumns.length + compareColumns.length}
                          className="py-2.5 px-3 text-xs text-slate-500 italic"
                        >
                          {compareValColA && (
                            <span className="mr-3">
                              {compareValColA}: <strong>{tableFooterSummary.sumA.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}</strong>
                            </span>
                          )}
                          {compareValColB && (
                            <span>
                              {compareValColB}: <strong>{tableFooterSummary.sumB.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}</strong>
                            </span>
                          )}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>

              {currentDisplayList.length > 300 && (
                <div className="text-center text-[11px] text-slate-500 italic">
                  Đang hiển thị 300 / {currentDisplayList.length.toLocaleString("vi-VN")} bản ghi. Hãy dùng nút xuất Excel để tải toàn bộ.
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* THÔNG ĐIỆP KHI CHƯA TẢI ĐỦ TỆP HOẶC THIẾU TỆP (GOM VÀO 1 MỐI ĐỂ TẢI NHANH) */}
      {(activeMainData.length === 0 || compareData.length === 0) && (
        <div className="p-6 bg-white mx-3.5 my-2 border border-dashed border-sky-300 space-y-4 text-center">
          <div className="w-12 h-12 bg-sky-50 text-sky-600 rounded-full flex items-center justify-center mx-auto shadow-inner">
            <GitCompare className="w-6 h-6" />
          </div>
          <div className="space-y-1">
            <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wide">
              {activeMainData.length === 0 && compareData.length === 0
                ? "Nạp Dữ Liệu 2 Kỳ Để Đối Soát (Tải File A & Tải File B)"
                : activeMainData.length === 0
                ? "Chưa Nạp Dữ Liệu File A (Kỳ 1)"
                : "Chưa Nạp Dữ Liệu File B (Kỳ 2 Đối Chiếu)"}
            </h4>
            <p className="text-xs text-slate-600 max-w-lg mx-auto">
              Vui lòng tải lên cả 2 tệp dữ liệu vào bảng dưới đây để hệ thống tự động tìm khóa chung, phát hiện bản ghi trùng khớp và đối soát chênh lệch số liệu.
            </p>
          </div>

          {/* KHỐI 2 TỆP VÀO 1 MỐI ĐỂ TẢI NHANH */}
          <div className="max-w-xl mx-auto grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1 text-left">
            {/* Thẻ File A */}
            <div className={`p-3.5 border rounded-none shadow-2xs ${activeMainData.length > 0 ? "bg-sky-50/70 border-sky-400" : "bg-slate-50 border-slate-300"}`}>
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-xs font-bold text-sky-950 uppercase flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-sky-600 inline-block" /> File A (Kỳ 1)
                </span>
                {activeMainData.length > 0 ? (
                  <span className="text-[10px] bg-sky-200 text-sky-900 px-1.5 py-0.5 font-bold">
                    ✓ {activeMainData.length.toLocaleString("vi-VN")} dòng
                  </span>
                ) : (
                  <span className="text-[10px] bg-amber-100 text-amber-900 px-1.5 py-0.5 font-semibold">
                    Chưa tải
                  </span>
                )}
              </div>
              <p className="text-[11px] text-slate-600 mb-3 truncate" title={activeFileName || ""}>
                {activeFileName ? `Tệp: ${activeFileName}` : "Chưa chọn tệp dữ liệu A"}
              </p>
              <label className="w-full bg-[#0369a1] hover:bg-[#0284c7] text-white font-bold text-xs py-2 px-3 rounded-none transition-colors shadow-2xs flex items-center justify-center gap-1.5 cursor-pointer">
                <Upload className="w-3.5 h-3.5" />
                <span>Tải file A</span>
                <input type="file" accept=".xlsx,.xls,.csv,.txt" onChange={handleFileUploadA} className="hidden" />
              </label>
            </div>

            {/* Thẻ File B */}
            <div className={`p-3.5 border rounded-none shadow-2xs ${compareData.length > 0 ? "bg-emerald-50/70 border-emerald-400" : "bg-slate-50 border-slate-300"}`}>
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-xs font-bold text-emerald-950 uppercase flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-600 inline-block" /> File B (Kỳ 2)
                </span>
                {compareData.length > 0 ? (
                  <span className="text-[10px] bg-emerald-200 text-emerald-900 px-1.5 py-0.5 font-bold">
                    ✓ {compareData.length.toLocaleString("vi-VN")} dòng
                  </span>
                ) : (
                  <span className="text-[10px] bg-amber-100 text-amber-900 px-1.5 py-0.5 font-semibold">
                    Chưa tải
                  </span>
                )}
              </div>
              <p className="text-[11px] text-slate-600 mb-3 truncate" title={compareSheetName || compareFile?.name || ""}>
                {compareSheetName || compareFile?.name ? `Tệp: ${compareSheetName || compareFile?.name}` : "Chưa chọn tệp đối chiếu B"}
              </p>
              <label className="w-full bg-[#047857] hover:bg-[#059669] text-white font-bold text-xs py-2 px-3 rounded-none transition-colors shadow-2xs flex items-center justify-center gap-1.5 cursor-pointer">
                <Upload className="w-3.5 h-3.5" />
                <span>Tải file B</span>
                <input type="file" accept=".xlsx,.xls,.csv,.txt" onChange={handleFileUpload} className="hidden" />
              </label>
            </div>
          </div>

          {detectedSheets.length > 1 && (
            <div className="pt-2 flex flex-wrap items-center justify-center gap-2">
              <span className="text-xs text-slate-600 font-medium">Hoặc chọn nhanh Sheet từ cùng tệp Excel:</span>
              <select
                onChange={e => {
                  if (e.target.value) handleSelectSheetFromWorkbook(e.target.value);
                }}
                defaultValue=""
                className="bg-white border border-emerald-500 text-emerald-900 font-bold text-xs px-2.5 py-1.5 rounded-none cursor-pointer"
              >
                <option value="" disabled>
                  Chọn Sheet làm File B...
                </option>
                {detectedSheets.map(s => (
                  <option key={s} value={s}>
                    Sheet: {s} ({sheetDataStore[s]?.data?.length || 0} dòng)
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default DataComparison;
