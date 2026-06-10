import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
// --- INDEXEDDB STORAGE FOR LARGE FILES (40-50MB+) INTEGRATED DIRECTLY FOR RELIABLE PORTABILITY ---
const DB_NAME = "VTongDatabase";
const DB_VERSION = 1;
const STORE_NAME = "appState";

interface AppState {
  mainData: any[];
  rawImportedData: any[];
  columns: string[];
  fileName: string;
  mapping: {
    mota: string;
    manganh: string;
    xa: string;
    doanhthu: string;
    laodong: string;
    idCol: string;
  };
  customColConfigs: any[];
}

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => {
      reject(new Error("Không thể khởi tạo cơ sở dữ liệu IndexedDB"));
    };
    request.onsuccess = () => {
      resolve(request.result);
    };
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
  });
}

async function saveAppState(state: AppState): Promise<void> {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      // IndexedDB lưu trữ đối tượng dạng bản sao cấu trúc thô trực tiếp cực mạnh mẽ và chuẩn xác
      const request = store.put(state, "currentSession");
      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error("Lỗi khi lưu dữ liệu vào IndexedDB"));
    });
  } catch (error) {
    console.error("IndexedDB Save Error:", error);
  }
}

async function loadAppState(): Promise<AppState | null> {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get("currentSession");
      request.onsuccess = () => {
        const val = request.result;
        if (!val) {
          resolve(null);
          return;
        }
        if (typeof val === "string") {
          try {
            resolve(JSON.parse(val));
          } catch (e) {
            reject(new Error("Dữ liệu phiên làm việc bị hỏng, đang khởi tạo lại..."));
          }
        } else {
          resolve(val); // Tương thích ngược dạng Object cũ
        }
      };
      request.onerror = () => {
        reject(new Error("Lỗi khi đọc dữ liệu từ IndexedDB"));
      };
    });
  } catch (error) {
    console.error("IndexedDB Load Error:", error);
    return null;
  }
}

async function clearAppState(): Promise<void> {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.delete("currentSession");
      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error("Lỗi khi xóa dữ liệu IndexedDB"));
    });
  } catch (error) {
    console.error("IndexedDB Clear Error:", error);
  }
}
import { 
  Home, 
  FileSpreadsheet, 
  GitMerge, 
  Combine, 
  Scissors, 
  BarChart3, 
  PieChart,
  Activity, 
  CheckSquare, 
  Download, 
  Loader2, 
  FileUp, 
  AlertTriangle, 
  CheckCircle2, 
  Info, 
  Brain,
  Layers,
  ArrowRight,
  ArrowRightLeft,
  Database,
  RefreshCw,
  Search,
  Plus,
  Trash2,
  FileCheck,
  Compass,
  Lock,
  KeyRound,
  LogOut
} from "lucide-react";

import { 
  vsicRawData, 
  normalizeSectorCode, 
  getSectorHierarchy, 
  smartSuggestSectorByDescription,
  getSectorLevel,
  getParentSectorCode,
  lookupSectorNameWithFallback
} from "./data/vsic";

import SectorRevenueChart from "./components/sectorRevenueChart";
import VsicCatalogExplorer from "./components/vsicCatalogExplorer";
import DescriptorMatchScanner from "./components/descriptorMatchScanner";
import { BeautifulReportTable } from "./components/BeautifulReportTable";

// Interface define
interface ColumnMapping {
  mota: string;
  manganh: string;
  xa: string;
  doanhthu: string;
  laodong: string;
  idCol: string; 
}

interface LogicRule {
  col: string;
  op: string;
  val: string;
}

export default function App() {
  const [activeTab, setActiveTab] = useState<string>("trangchu");
  const [loading, setLoading] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string>("");

  // Trạng thái mật khẩu bảo vệ ứng dụng độc lập tránh bắt đăng nhập email phiền hà
  const [appPassword, setAppPassword] = useState<string>(() => {
    return localStorage.getItem("vsic_app_password") || "admin123";
  });
  const [isAuthorized, setIsAuthorized] = useState<boolean>(() => {
    return localStorage.getItem("vsic_app_authorized") === "true";
  });
  const [typedPassword, setTypedPassword] = useState<string>("");
  const [passwordError, setPasswordError] = useState<string>("");
  const [showPasswordChangeModal, setShowPasswordChangeModal] = useState<boolean>(false);
  const [newPasswordVal, setNewPasswordVal] = useState<string>("");

  const handleCheckPassword = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (typedPassword === appPassword) {
      localStorage.setItem("vsic_app_authorized", "true");
      setIsAuthorized(true);
      setPasswordError("");
    } else {
      setPasswordError("Mật khẩu truy cập chưa chính xác! Vui lòng kiểm tra lại.");
    }
  };

  const handleChangePassword = () => {
    if (!newPasswordVal.trim()) {
      alert("Vui lòng nhập mật khẩu mới!");
      return;
    }
    setAppPassword(newPasswordVal);
    localStorage.setItem("vsic_app_password", newPasswordVal);
    setShowPasswordChangeModal(false);
    setNewPasswordVal("");
    alert(`Đã đổi mật khẩu hành trị thành công sang: ${newPasswordVal}`);
  };

  const handleLogout = () => {
    localStorage.removeItem("vsic_app_authorized");
    setIsAuthorized(false);
    setTypedPassword("");
  };

  // Storage chính
  const [mainData, setMainData] = useState<any[]>([]);
  const [rawImportedData, setRawImportedData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>("");

  // Setup reporting states
  const [quickReportResultRows, setQuickReportResultRows] = useState<any[]>([]);
  const [quickReportResultCols, setQuickReportResultCols] = useState<string[]>([]);
  const [quickReportLevel, setQuickReportLevel] = useState<number>(1);

  // Column Mapping
  const [mapping, setMapping] = useState<ColumnMapping>({
    mota: "",
    manganh: "",
    xa: "",
    doanhthu: "",
    laodong: "",
    idCol: ""
  });

  const [customColConfigs, setCustomColConfigs] = useState<{
    originalName: string;
    use: boolean;
    newName: string;
    role: "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" | "";
  }[]>([]);

  // Trạng thái cho Dual-Pane Mapping và double click, cùng kiểu định dạng báo cáo xoay Pivot
  const [selectedTargetKey, setSelectedTargetKey] = useState<keyof ColumnMapping>("mota");
  const [reportType, setReportType] = useState<"flat" | "pivot">("pivot");
  const [isConfigExpanded, setIsConfigExpanded] = useState<boolean>(true);

  // Setup dữ liệu so sánh (Diff)
  const [oldData, setOldData] = useState<any[]>([]);
  const [oldFileName, setOldFileName] = useState<string>("");
  const [newData, setNewData] = useState<any[]>([]);
  const [newFileName, setNewFileName] = useState<string>("");
  const [diffKey, setDiffKey] = useState<string>("");

  // Setup dữ liệu ghép nối (Merge)
  const [leftData, setLeftData] = useState<any[]>([]);
  const [leftFileName, setLeftFileName] = useState<string>("");
  const [rightData, setRightData] = useState<any[]>([]);
  const [rightFileName, setRightFileName] = useState<string>("");
  const [leftKey, setLeftKey] = useState<string>("");
  const [rightKey, setRightKey] = useState<string>("");

  // Trang phân tách
  const [splitCol, setSplitCol] = useState<string>("");

  // Trạng thái ghép nhiều sheet từ cùng một file Excel tải lên
  const [detectedWorkbook, setDetectedWorkbook] = useState<any | null>(null);
  const [detectedSheets, setDetectedSheets] = useState<string[]>([]);
  const [selectedSheetsToMerge, setSelectedSheetsToMerge] = useState<string[]>([]);
  const [sheetMergeCommonKey, setSheetMergeCommonKey] = useState<string>("");

  // Các state hỗ trợ Báo cáo phối hợp hai chiều (Xã × Ngành VSIC)
  const [crossReportData, setCrossReportData] = useState<any[]>([]);
  const [crossReportCols, setCrossReportCols] = useState<string[]>([]);
  const [crossReportManganhCol, setCrossReportManganhCol] = useState<string>("");
  const [crossReportXaCol, setCrossReportXaCol] = useState<string>("");
  const [crossReportDoanhThuCol, setCrossReportDoanhThuCol] = useState<string>("");
  const [crossReportLaoDongCol, setCrossReportLaoDongCol] = useState<string>("");
  const [crossReportLevel, setCrossReportLevel] = useState<number>(2); // 1: Cấp 1, 2: Cấp 2, 5: Giữ nguyên

  // Quy tắc tổng hợp (Aggregate rules)
  const [groupByCols, setGroupByCols] = useState<string[]>([]);
  const [aggRules, setAggRules] = useState<{ col: string; op: string }[]>([]);
  const [newAggCol, setNewAggCol] = useState<string>("");
  const [newAggOp, setNewAggOp] = useState<string>("sum");

  // Quy tắc kiểm tra logic đa điều kiện
  const [ifRules, setIfRules] = useState<LogicRule[]>([]);
  const [thenRules, setThenRules] = useState<LogicRule[]>([]);
  const [ifCombine, setIfCombine] = useState<"AND" | "OR">("AND");
  const [thenCombine, setThenCombine] = useState<"AND" | "OR">("AND");

  // Quy tắc mới cho Logic
  const [newIfRule, setNewIfRule] = useState<LogicRule>({ col: "", op: "==", val: "" });
  const [newThenRule, setNewThenRule] = useState<LogicRule>({ col: "", op: "==", val: "" });

  // States cho Phân Hệ 1: Tổng hợp ngành cấp 2 (Mới độc lập)
  const [t2IndustryCol, setT2IndustryCol] = useState<string>("");
  const [t2MetricCols, setT2MetricCols] = useState<string[]>([]);
  const [t2AggMethod, setT2AggMethod] = useState<"sum" | "avg">("sum");
  const [t2ReportData, setT2ReportData] = useState<any[]>([]);
  const [t2ReportCols, setT2ReportCols] = useState<string[]>([]);
  const [t2ReportLevel, setT2ReportLevel] = useState<number>(2);

  // States riêng biệt cho người dùng chọn cột báo cáo nhanh và xoay đa năng trực tiếp không bị bó buộc
  const [quickReportManganhCol, setQuickReportManganhCol] = useState<string>("");
  const [quickReportXaCol, setQuickReportXaCol] = useState<string>("");
  const [quickReportDoanhThuCol, setQuickReportDoanhThuCol] = useState<string>("");
  const [quickReportLaoDongCol, setQuickReportLaoDongCol] = useState<string>("");
  const [pivotManganhCol, setPivotManganhCol] = useState<string>("");

  // States cho Phân Hệ 2: Chuẩn hóa khớp ngành VSIC cấp 5 (Mới độc lập)
  const [stdIndustryCol, setStdIndustryCol] = useState<string>("");
  const [stdDescriptionCol, setStdDescriptionCol] = useState<string>("");
  const [stdReportAnomalies, setStdReportAnomalies] = useState<any[]>([]);
  const [stdMatchStats, setStdMatchStats] = useState<{ total: number; valid: number; invalid: number; conflicts: number }>({ total: 0, valid: 0, invalid: 0, conflicts: 0 });

  // States cho Phân Hệ Đối chiếu 2 cột tự chọn (Theo yêu cầu người dùng)
  const [crossCompareColA, setCrossCompareColA] = useState<string>("");
  const [crossCompareColB, setCrossCompareColB] = useState<string>("");
  const [crossCompareRule, setCrossCompareRule] = useState<string>("normalize");
  const [crossCompareAnomalies, setCrossCompareAnomalies] = useState<any[]>([]);
  const [crossCompareStats, setCrossCompareStats] = useState<{ total: number; matchCount: number; mismatchCount: number }>({ total: 0, matchCount: 0, mismatchCount: 0 });

  // Báo cáo động độc lập khởi tạo trống để người dùng tự chọn, không đoán bừa bãi
  useEffect(() => {
    // Không tự động đoán gán cứng các cột phục vụ báo cáo động nữa.
  }, [columns]);



  // Phân trang cho viewer
  const [viewPage, setViewPage] = useState<number>(1);
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [inconsistenciesTab, setInconsistenciesTab] = useState<"desc" | "code">("desc");
  const pageSize = 50;

  // Khôi phục dữ liệu từ IndexedDB khi mở ứng dụng
  useEffect(() => {
    async function restoreSession() {
      setLoading(true);
      setStatusMessage("Đang kiểm tra dữ liệu phiên làm việc trước đó trong bộ nhớ...");
      try {
        const saved = await loadAppState();
        if (saved && saved.mainData && saved.mainData.length > 0) {
          setMainData(saved.mainData);
          setRawImportedData(saved.rawImportedData || saved.mainData);
          setColumns(saved.columns);
          setFileName(saved.fileName);
          if (saved.mapping) setMapping(saved.mapping);
          if (saved.customColConfigs) setCustomColConfigs(saved.customColConfigs);
          setStatusMessage(`Đã khôi phục thành công tệp "${saved.fileName}" (${saved.mainData.length} dòng) từ phiên trước!`);
        } else {
          setStatusMessage("");
        }
      } catch (err) {
        console.error("Lỗi khi đọc dữ liệu lưu trữ:", err);
      } finally {
        setLoading(false);
      }
    }
    restoreSession();
  }, []);

  // Tự động lưu trữ phiên làm việc
  const autoSaveSession = async (
    customMainData?: any[],
    customRawData?: any[],
    customCols?: string[],
    customFileName?: string,
    customMapping?: ColumnMapping,
    customConfigs?: any[]
  ) => {
    try {
      await saveAppState({
        mainData: customMainData !== undefined ? customMainData : mainData,
        rawImportedData: customRawData !== undefined ? customRawData : rawImportedData,
        columns: customCols !== undefined ? customCols : columns,
        fileName: customFileName !== undefined ? customFileName : fileName,
        mapping: customMapping !== undefined ? customMapping : mapping,
        customColConfigs: customConfigs !== undefined ? customConfigs : customColConfigs
      });
    } catch (err) {
      console.error("Không thể lưu trạng thái phiên:", err);
    }
  };

  // Tự động lưu trữ phiên làm việc có Trễ Debounce 1.5 giây để tối ưu hiệu năng ghi file lớn
  useEffect(() => {
    if (mainData.length === 0) return;

    const timer = setTimeout(() => {
      saveAppState({
        mainData,
        rawImportedData,
        columns,
        fileName,
        mapping,
        customColConfigs
      }).catch(err => console.error("Lỗi tự động lưu phiên làm việc:", err));
    }, 1500);

    return () => clearTimeout(timer);
  }, [mainData, rawImportedData, columns, fileName, mapping, customColConfigs]);

  // Thuật toán hiệu chỉnh dải ô (Range) trực tiếp trên Worksheet cực kỳ nhanh chóng và an toàn
  // Hỗ trợ cấu trúc Dense Worksheet (lưu trữ dạng mảng 2 chiều !data thay vì flat keys), tránh cực đọ việc lặp hàng triệu lần và không gọi Object.keys() loại bỏ hoàn toàn "Too many properties to enumerate"
  const optimizeAndCompactSheet = (wb: XLSX.WorkBook, sheetName: string): XLSX.WorkSheet => {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) return ws;
    
    let range;
    try {
      range = XLSX.utils.decode_range(ws['!ref']);
    } catch (e) {
      return ws;
    }

    const startRow = range.s.r;
    const endRow = range.e.r;
    const startCol = range.s.c;
    const endCol = range.e.c;

    if (endRow <= startRow) return ws;

    // Hàm kiểm tra nhanh xem dòng r có dữ liệu thực tế hay không
    const checkRowHasData = (r: number): boolean => {
      const maxColToSearch = Math.min(endCol, startCol + 50); // Giới hạn kiểm tra 50 cột đầu tiên để tối ưu tốc độ dò tìm
      if (ws['!data']) {
        const rowData = ws['!data'][r];
        if (!rowData) return false;
        for (let c = startCol; c <= maxColToSearch; c++) {
          const cellObj = rowData[c];
          if (cellObj && cellObj.v !== undefined && cellObj.v !== null && String(cellObj.v).trim() !== "") {
            return true;
          }
        }
      } else {
        for (let c = startCol; c <= maxColToSearch; c++) {
          const cellRef = XLSX.utils.encode_cell({ r, c });
          const cellObj = ws[cellRef];
          if (cellObj && cellObj.v !== undefined && cellObj.v !== null && String(cellObj.v).trim() !== "") {
            return true;
          }
        }
      }
      return false;
    };

    // Dò ngược từ endRow về startRow để tìm dòng thực tế cuối cùng chứa dữ liệu nhanh chóng và chính xác 100%
    let realLastRow = startRow;
    for (let r = endRow; r >= startRow; r--) {
      if (checkRowHasData(r)) {
        realLastRow = r;
        break;
      }
    }

    // Giải phóng bớt mảng dòng nếu sử dụng cấu hình Sheet dạng dense
    if (ws['!data'] && ws['!data'].length > realLastRow + 1) {
      ws['!data'].length = realLastRow + 1;
    }

    // Cập nhật dải ô !ref hiệu chuẩn trực tiếp cực kỳ nhẹ nhàng mà không nhân bản dữ liệu
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: startRow, c: startCol },
      e: { r: realLastRow, c: endCol }
    });

    return ws;
  };

  // Bộ phân giải CSV tối ưu hóa cao cho các tệp lớn (như 50MB+), loại bỏ lỗi "Too many properties to enumerate"
  const parseCSV = (rawText: string): any[] => {
    let text = rawText;
    if (text.charCodeAt(0) === 0xFEFF) {
      text = text.substring(1);
    }

    // Tự động phát hiện dấu phân tách (comma, semicolon, tab) từ dòng dữ liệu đầu tiên
    const firstLineEnd = text.indexOf('\n');
    const firstLine = firstLineEnd === -1 ? text : text.substring(0, firstLineEnd);
    
    let delimiter = ',';
    const commaCount = (firstLine.match(/,/g) || []).length;
    const semicolonCount = (firstLine.match(/;/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    
    if (semicolonCount > commaCount && semicolonCount > tabCount) {
      delimiter = ';';
    } else if (tabCount > commaCount && tabCount > semicolonCount) {
      delimiter = '\t';
    }

    const length = text.length;
    const rows: string[][] = [];
    let currentRow: string[] = [];
    
    let i = 0;
    let inQuotes = false;
    let start = 0;
    
    while (i < length) {
      const char = text[i];
      
      if (char === '"') {
        if (!inQuotes) {
          inQuotes = true;
          start = i + 1; // nhảy qua dấu ngoặc kép mở đầu
        } else {
          // Kiểm tra dấu ngoặc kép trốn (escaped quote "")
          if (i + 1 < length && text[i + 1] === '"') {
            i++; // bỏ qua dấu ngoặc kép thứ hai
          } else {
            inQuotes = false;
          }
        }
      } else if (!inQuotes) {
        if (char === delimiter) {
          let cell = text.substring(start, i);
          // Loại bỏ ngoặc kép bao quanh khi đọc chuỗi trường
          if (text[i - 1] === '"' && text[start - 1] === '"') {
            cell = cell.substring(0, cell.length - 1);
          }
          if (cell.includes('""')) {
            cell = cell.replace(/""/g, '"');
          }
          currentRow.push(cell.trim());
          start = i + 1;
        } else if (char === '\n' || char === '\r') {
          let cell = text.substring(start, i);
          if (text[i - 1] === '"' && text[start - 1] === '"') {
            cell = cell.substring(0, cell.length - 1);
          }
          if (cell.includes('""')) {
            cell = cell.replace(/""/g, '"');
          }
          currentRow.push(cell.trim());
          
          if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) {
            rows.push(currentRow);
          }
          currentRow = [];
          
          if (char === '\r' && i + 1 < length && text[i + 1] === '\n') {
            i++;
          }
          start = i + 1;
        }
      }
      i++;
    }
    
    if (start < length) {
      let cell = text.substring(start, length);
      if (text[length - 1] === '"' && text[start - 1] === '"') {
        cell = cell.substring(0, cell.length - 1);
      }
      if (cell.includes('""')) {
        cell = cell.replace(/""/g, '"');
      }
      currentRow.push(cell.trim());
    }
    if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) {
      rows.push(currentRow);
    }

    if (rows.length === 0) return [];

    const headers = rows[0];
    const data: any[] = [];
    for (let idx = 1; idx < rows.length; idx++) {
      const r = rows[idx];
      const obj: any = {};
      let hasData = false;
      for (let c = 0; c < headers.length; c++) {
        const headerName = headers[c] || `Cột ${c + 1}`;
        const val = r[c] !== undefined ? r[c] : "";
        obj[headerName] = val;
        if (val !== "") hasData = true;
      }
      if (hasData) {
        data.push(obj);
      }
    }

    return data;
  };

  // Đọc file CSV hoặc Excel bằng xlsx hoặc Bộ phân giải CSV tối ưu
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "main" | "old" | "new" | "left" | "right") => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang tải tệp: ${file.name}...`);

    const isCSV = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");

    if (isCSV) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const text = evt.target?.result as string;
          if (!text) {
            throw new Error("Không thể đọc nội dung tệp tin!");
          }

          setStatusMessage(`Đang xử lý nội dung CSV: ${file.name}...`);
          const data = parseCSV(text);

          if (data.length === 0) {
            alert("Tệp trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const cols = Object.keys(data[0] as any);

          if (type === "main") {
            setRawImportedData(data);
            setMainData(data);
            setColumns(cols);
            setFileName(file.name);

            // Giữ mọi cột trống hoàn toàn để người dùng tự do lựa chọn thủ công tại các chức năng tương ứng, không tự động đoán
            setQuickReportManganhCol("");
            setStdIndustryCol("");
            setCrossCompareColA("");
            setStdDescriptionCol("");
            setQuickReportXaCol("");
            setQuickReportDoanhThuCol("");
            setQuickReportLaoDongCol("");

            const autoMap: ColumnMapping = { 
              mota: "", 
              manganh: "", 
              xa: "", 
              doanhthu: "", 
              laodong: "", 
              idCol: "" 
            };
            setMapping(autoMap);

            // Khởi tạo danh sách cấu hình cột động từ tệp vừa nạp
            const initConfigs = cols.map(c => {
              return {
                originalName: c,
                use: true,
                newName: c, // giữ nguyên tên ban đầu, cho phép người dùng sửa đổi trực tiếp
                role: "" as any
              };
            });
            setCustomColConfigs(initConfigs);

            setActiveTab("xemdulieu");

            // Lưu vĩnh viễn vào IndexedDB để không bị mất khi F5 hoặc đóng tab
            autoSaveSession(data, data, cols, file.name, autoMap, initConfigs);
          } else if (type === "old") {
            setOldData(data);
            setOldFileName(file.name);
          } else if (type === "new") {
            setNewData(data);
            setNewFileName(file.name);
          } else if (type === "left") {
            setLeftData(data);
            setLeftFileName(file.name);
          } else if (type === "right") {
            setRightData(data);
            setRightFileName(file.name);
          }

          setStatusMessage(`Đã tải thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file CSV: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsText(file, "UTF-8");
    } else {
      // Đối với tệp Excel Binary (.xlsx, .xls, .ods,...)
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          if (!arrayBuffer) {
            throw new Error("Không thể đọc nội dung tệp tin!");
          }

          // Dùng dense: true để SheetJS sinh cấu trúc dải ô dạng 2D Array
          // Loại bỏ tuyệt đối việc lặp keys và không sinh hàng triệu keys phẳng trên WorkSheet lặp
          const wb = XLSX.read(arrayBuffer, { 
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false
          });

          if (type === "main") {
            setDetectedWorkbook(wb);
            setDetectedSheets(wb.SheetNames);
            if (wb.SheetNames.length > 1) {
              setSelectedSheetsToMerge(wb.SheetNames);
              setSheetMergeCommonKey("");
            } else {
              setSelectedSheetsToMerge([]);
              setSheetMergeCommonKey("");
            }
          }

          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];

          // sheet_to_json hỗ trợ 100% Dense worksheets sinh ra từ dense: true
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          if (data.length === 0) {
            alert("Tệp trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const cols = Object.keys(data[0] as any);

          if (type === "main") {
            setRawImportedData(data);
            setMainData(data);
            setColumns(cols);
            setFileName(file.name);

            // Giữ mọi cột trống hoàn toàn để người dùng tự do lựa chọn thủ công tại các chức năng tương ứng, không tự động đoán
            setQuickReportManganhCol("");
            setStdIndustryCol("");
            setCrossCompareColA("");
            setStdDescriptionCol("");
            setQuickReportXaCol("");
            setQuickReportDoanhThuCol("");
            setQuickReportLaoDongCol("");

            const autoMap: ColumnMapping = { 
              mota: "", 
              manganh: "", 
              xa: "", 
              doanhthu: "", 
              laodong: "", 
              idCol: "" 
            };
            setMapping(autoMap);

            // Khởi tạo danh sách cấu hình cột động từ tệp vừa nạp
            const initConfigs = cols.map(c => {
              return {
                originalName: c,
                use: true,
                newName: c, // giữ nguyên tên ban đầu, cho phép người dùng sửa đổi trực tiếp
                role: "" as any
              };
            });
            setCustomColConfigs(initConfigs);

            setActiveTab("xemdulieu");

            // Lưu vĩnh viễn vào IndexedDB để không bị mất khi F5 hoặc đóng tab
            autoSaveSession(data, data, cols, file.name, autoMap, initConfigs);

          } else if (type === "old") {
            setOldData(data);
            setOldFileName(file.name);
          } else if (type === "new") {
            setNewData(data);
            setNewFileName(file.name);
          } else if (type === "left") {
            setLeftData(data);
            setLeftFileName(file.name);
          } else if (type === "right") {
            setRightData(data);
            setRightFileName(file.name);
          }

          setStatusMessage(`Đã tải thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file Excel: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Reset toàn bộ dữ liệu
  const clearData = () => {
    setMainData([]);
    setRawImportedData([]);
    setColumns([]);
    setFileName("");
    setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
    setCustomColConfigs([]);
    clearAppState().catch(err => console.error("Lỗi khi xóa dữ liệu IndexedDB:", err));
  };

  // Hàm bắt đầu thực hiện ghép các sheet đã chọn dựa trên một cột chung
  const handleMergeWorkbookSheets = async () => {
    if (!detectedWorkbook) {
      alert("Không tìm thấy tệp Excel đang thao tác!");
      return;
    }
    if (selectedSheetsToMerge.length < 2) {
      alert("Vui lòng chọn ít nhất 2 sheet để thực hiện ghép/gộp dữ liệu!");
      return;
    }
    if (!sheetMergeCommonKey) {
      alert("Vui lòng chọn cột liên kết chung (Mã số thuế/Mã định danh) để liên kết các dòng!");
      return;
    }

    setLoading(true);
    setProgress(20);
    setStatusMessage("Đang quét nội dung các sheet và bóc tách dữ liệu...");
    await sleep(200);

    try {
      // Đọc toàn bộ dữ liệu của từng sheet được chọn
      const sheetDataMap = new Map<string, any[]>();
      selectedSheetsToMerge.forEach(sheetName => {
        const ws = detectedWorkbook.Sheets[sheetName];
        if (ws) {
          const data = XLSX.utils.sheet_to_json(ws) as any[];
          sheetDataMap.set(sheetName, data);
        }
      });

      setProgress(50);
      setStatusMessage("Đang khớp nối dữ liệu các dòng theo cột chung...");
      await sleep(200);

      // Tiến hành gộp dữ liệu sử dụng Map
      const mergedMap = new Map<string, any>();
      const allColsSet = new Set<string>();

      // Duyệt qua từng sheet, gộp dữ liệu
      selectedSheetsToMerge.forEach(sheetName => {
        const rows = sheetDataMap.get(sheetName) || [];
        rows.forEach(row => {
          const rawKeyVal = row[sheetMergeCommonKey];
          if (rawKeyVal === undefined || rawKeyVal === null) {
            // Dòng không có khóa chung, ta vẫn giữ nhưng gán một khóa tạm độc nhất để tránh bị đè mất dòng
            const idTemp = `_no_key_${Math.random().toString(36).substring(2, 11)}`;
            mergedMap.set(idTemp, { ...row });
            Object.keys(row).forEach(k => allColsSet.add(k));
          } else {
            const keyStr = String(rawKeyVal).trim();
            if (keyStr === "") {
              const idTemp = `_no_key_${Math.random().toString(36).substring(2, 11)}`;
              mergedMap.set(idTemp, { ...row });
              Object.keys(row).forEach(k => allColsSet.add(k));
            } else {
              if (mergedMap.has(keyStr)) {
                // Trộn dòng mới vào dòng đã tồn tại
                mergedMap.set(keyStr, { ...mergedMap.get(keyStr), ...row });
              } else {
                mergedMap.set(keyStr, { ...row });
              }
              Object.keys(row).forEach(k => allColsSet.add(k));
            }
          }
        });
      });

      const mergedList = Array.from(mergedMap.values());
      const updatedCols = Array.from(allColsSet);

      if (mergedList.length === 0) {
        alert("Kết quả ghép dữ liệu rỗng! Vui lòng kiểm tra lại cột chung.");
        setLoading(false);
        return;
      }

      setRawImportedData(mergedList);
      setMainData(mergedList);
      setColumns(updatedCols);

      // Cấu hình lại cột
      const initConfigs = updatedCols.map(c => {
        return {
          originalName: c,
          use: true,
          newName: c,
          role: "" as any
        };
      });
      setCustomColConfigs(initConfigs);

      // Không tự gán cột ID làm idCol chính dể người dùng tự chọn
      const autoMap: ColumnMapping = { 
        mota: "", 
        manganh: "", 
        xa: "", 
        doanhthu: "", 
        laodong: "", 
        idCol: "" 
      };
      setMapping(autoMap);

      setProgress(100);
      setStatusMessage(`Ghép thành công ${selectedSheetsToMerge.length} sheets thành ${mergedList.length} dòng dữ liệu thống nhất!`);
      await sleep(300);

      // Lưu IndexedDB
      autoSaveSession(mergedList, mergedList, updatedCols, fileName, autoMap, initConfigs);
      setActiveTab("xemdulieu");

    } catch (err: any) {
      alert("Lỗi trong quá trình ghép các Sheet: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // Hàm tính toán báo cáo chéo phối hợp hai chiều (Xã x Ngành VSIC)
  const handleCalcCrossReport = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng tải hoặc nạp dữ liệu chính trước.");
      return;
    }
    if (!crossReportManganhCol) {
      alert("Vui lòng chỉ định cột chứa Mã Ngành ở bộ chọn!");
      return;
    }
    if (!crossReportXaCol) {
      alert("Vui lòng chỉ định cột chứa Xã / Địa bàn ở bộ chọn!");
      return;
    }

    setLoading(true);
    setProgress(30);
    setStatusMessage("Đang quét phân cấp ngành VSIC và gom tích chỉ số...");
    await sleep(200);

    try {
      const groupedMap = new Map<string, {
        xa: string;
        nganhCode: string;
        nganhLabel: string;
        sumDoanhThu: number;
        sumLaoDong: number;
        countDN: number;
      }>();

      mainData.forEach(row => {
        const xaVal = String(row[crossReportXaCol] || "Khác").trim();
        const rawMng = String(row[crossReportManganhCol] || "").trim();
        const mngNormalized = normalizeSectorCode(rawMng);

        let nganhCode = mngNormalized;
        let nganhLabel = "";

        if (crossReportLevel === 1) {
          // Lấy Ngành Cấp 1 (Chữ cái A-U)
          let s1Code = "";
          if (mngNormalized) {
            if (/^[a-zA-Z]$/.test(mngNormalized)) {
              s1Code = mngNormalized.toUpperCase();
            } else {
              const sec2Code = mngNormalized.slice(0, 2);
              s1Code = getParentSectorCode(sec2Code) || "";
            }
          }
          nganhCode = s1Code || "CHUA_PHAN_LOAI";
          nganhLabel = vsicRawData[nganhCode] || "Ngành cấp 1 chưa định dạng chuẩn";
        } else if (crossReportLevel === 2) {
          // Lấy Ngành Cấp 2 (2 số đầu)
          const s2Code = mngNormalized ? mngNormalized.slice(0, 2) : "";
          nganhCode = s2Code || "CHUA_PHAN_LOAI";
          nganhLabel = vsicRawData[nganhCode] || "Ngành cấp 2 chưa định dạng chuẩn";
        } else {
          // Giữ nguyên (Cấp 5)
          nganhCode = mngNormalized || "CHUA_PHAN_LOAI";
          nganhLabel = vsicRawData[nganhCode] || "Mã kinh tế chưa ghi nhận chuẩn";
        }

        const groupKey = `${xaVal}|||${nganhCode}`;

        // Trích xuất doanh thu
        let dtVal = 0;
        if (crossReportDoanhThuCol) {
          const val = parseFloat(String(row[crossReportDoanhThuCol]).replace(/[^0-9.\-]/g, ""));
          if (!isNaN(val)) dtVal = val;
        }

        // Trích xuất lao động
        let ldVal = 0;
        if (crossReportLaoDongCol) {
          const val = parseFloat(String(row[crossReportLaoDongCol]).replace(/[^0-9.\-]/g, ""));
          if (!isNaN(val)) ldVal = val;
        }

        if (groupedMap.has(groupKey)) {
          const prev = groupedMap.get(groupKey)!;
          prev.sumDoanhThu += dtVal;
          prev.sumLaoDong += ldVal;
          prev.countDN += 1;
        } else {
          groupedMap.set(groupKey, {
            xa: xaVal,
            nganhCode: nganhCode,
            nganhLabel: nganhLabel,
            sumDoanhThu: dtVal,
            sumLaoDong: ldVal,
            countDN: 1
          });
        }
      });

      setProgress(70);
      setStatusMessage("Xây dựng dải tổng hợp phân cấp hai chiều...");
      await sleep(150);

      // Sắp xếp
      const listResults = Array.from(groupedMap.values()).sort((a, b) => {
        const cmpXa = a.xa.localeCompare(b.xa, "vi");
        if (cmpXa !== 0) return cmpXa;
        return a.nganhCode.localeCompare(b.nganhCode);
      });

      const reportRows: any[] = [];
      let totalDN = 0;
      let totalDoanhThu = 0;
      let totalLaoDong = 0;

      listResults.forEach((val, index) => {
        reportRows.push({
          "STT": index + 1,
          "Địa bàn (Xã)": val.xa,
          [`Mã ngành hạch toán`]: val.nganhCode,
          "Tên phân loại ngành kinh tế": val.nganhLabel,
          "Số lượng Doanh nghiệp (DN)": val.countDN,
          "Tổng Doanh Thu hạch toán": Math.round(val.sumDoanhThu * 100) / 100,
          "Tổng Lao Động hạch toán": Math.round(val.sumLaoDong * 100) / 100
        });

        totalDN += val.countDN;
        totalDoanhThu += val.sumDoanhThu;
        totalLaoDong += val.sumLaoDong;
      });

      // Tạo dòng sum toàn bảng
      reportRows.push({
        "STT": "LŨY KẾ",
        "Địa bàn (Xã)": "TỔNG CỘNG LŨY KẾ TOÀN BỘ BẢNG",
        [`Mã ngành hạch toán`]: "-",
        "Tên phân loại ngành kinh tế": "-",
        "Số lượng Doanh nghiệp (DN)": totalDN,
        "Tổng Doanh Thu hạch toán": Math.round(totalDoanhThu * 100) / 100,
        "Tổng Lao Động hạch toán": Math.round(totalLaoDong * 100) / 100
      });

      setCrossReportData(reportRows);
      if (reportRows.length > 0) {
        setCrossReportCols(Object.keys(reportRows[0]));
      }

      setProgress(100);
      setStatusMessage(`Tính toán chéo thành công ${reportRows.length - 1} dòng chi tiết theo Xã × Phân cấp ngành!`);
      await sleep(250);

    } catch (e: any) {
      alert("Xảy ra lỗi khi tính toán tổng hợp chéo: " + e.message);
    } finally {
      setLoading(false);
    }
  };

  // Xuất file báo cáo chéo
  const handleExportCrossReportExcel = () => {
    if (crossReportData.length === 0) {
      alert("Không có dữ liệu báo cáo chéo phối hợp để xuất!");
      return;
    }

    try {
      const ws = XLSX.utils.json_to_sheet(crossReportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Hop_Xa_Nganh");

      // Set width
      ws["!cols"] = [
        { wch: 8 },  // STT
        { wch: 22 }, // Xã
        { wch: 22 }, // Mã ngành
        { wch: 45 }, // Tên phân loại ngành
        { wch: 20 }, // Số lượng DN
        { wch: 22 }, // Tổng doanh thu
        { wch: 22 }  // Tổng lao động
      ];

      XLSX.writeFile(wb, `Bao_Cao_Tong_Hop_Phoi_Hop_Xa_Nganh_Cap_${crossReportLevel}.xlsx`);
    } catch (err: any) {
      alert("Lỗi xuất Excel: " + err.message);
    }
  };

  // Áp dụng định nghĩa lại tên cột & tái cấu trúc bảng dữ liệu mới
  const handleApplyColumnRedefinition = async () => {
    if (rawImportedData.length === 0) {
      alert("Không tìm thấy dữ liệu tệp gốc để tái cấu trúc! Hãy nạp tệp chính trước.");
      return;
    }

    const activeConfigs = customColConfigs.filter(cfg => cfg.use && cfg.newName.trim() !== "");
    if (activeConfigs.length === 0) {
      alert("Vui lòng chọn sử dụng ít nhất một cột và đặt tên dễ hiểu hợp lệ!");
      return;
    }

    setLoading(true);
    setProgress(10);
    setStatusMessage("Đang tiến hành lọc bỏ cột thừa và đổi tên cột theo định nghĩa của bạn...");
    await sleep(200);

    // Tạo bảng chứa dữ liệu mới gồm các cột được chọn và tên cột mới
    const restructuredRows = rawImportedData.map(row => {
      const newRow: any = {};
      activeConfigs.forEach(cfg => {
        const val = row[cfg.originalName];
        newRow[cfg.newName.trim()] = val !== undefined && val !== null ? val : "";
      });
      return newRow;
    });

    // Cập nhật cấu hình mapping bảo toàn theo chỉ định người dùng khi đổi tên cột
    const newMapping: ColumnMapping = { ...mapping };
    Object.keys(newMapping).forEach((roleKey) => {
      const currentMappedCol = newMapping[roleKey as keyof ColumnMapping];
      if (currentMappedCol) {
        const config = customColConfigs.find(cfg => cfg.originalName === currentMappedCol && cfg.use);
        if (config) {
          newMapping[roleKey as keyof ColumnMapping] = config.newName.trim();
        } else {
          newMapping[roleKey as keyof ColumnMapping] = "";
        }
      }
    });

    const newCols = Object.keys(restructuredRows[0] || {});

    setMainData(restructuredRows);
    setColumns(newCols);
    setMapping(newMapping);
    setViewPage(1);

    setProgress(100);
    setStatusMessage(`Tái cấu trúc bảng thành công! Đã giữ lại ${activeConfigs.length} cột tự định nghĩa.`);
    
    // Lưu vĩnh viễn trạng thái đã bổ sung/khớp cột mới vào IndexedDB
    autoSaveSession(restructuredRows, rawImportedData, newCols, fileName, newMapping, customColConfigs);

    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu"); // Di chuyển tới tab hiển thị bảng dữ liệu mới
  };

  // Bộ lọc dữ liệu viewer
  const filteredData = useMemo(() => {
    if (!searchTerm) return mainData;
    const term = searchTerm.toLowerCase();
    return mainData.filter(row => {
      return Object.values(row).some(val => String(val).toLowerCase().includes(term));
    });
  }, [mainData, searchTerm]);

  // Phân tích bất nhất mã ngành và mô tả
  const inconAnalysis = useMemo(() => {
    if (activeTab !== "chuanhoanganh") {
      return { descToCodes: [], codeToDescs: [] };
    }
    const targetMota = stdDescriptionCol || mapping.mota;
    const targetManganh = stdIndustryCol || mapping.manganh;

    if (!mainData.length || !targetMota || !targetManganh) {
      return { descToCodes: [], codeToDescs: [] };
    }

    // 1. CÙNG MÔ TẢ -> KHÁC MÃ LIÊN KẾT
    const descMap = new Map<string, Array<{ code: string; rowIdx: number; row: any }>>();
    mainData.forEach((row, idx) => {
      if (!row || typeof row !== 'object') return;
      const rawMota = String(row[targetMota] || "").trim();
      const cleanMota = rawMota.toLowerCase().replace(/\s+/g, " ");
      if (!cleanMota) return;
      const code = normalizeSectorCode(row[targetManganh]);
      
      if (!descMap.has(cleanMota)) {
        descMap.set(cleanMota, []);
      }
      descMap.get(cleanMota)!.push({ code, rowIdx: idx, row });
    });

    const descToCodes: Array<{
      motaText: string;
      occurrences: number;
      codes: Array<{ code: string; count: number; rows: number[] }>;
    }> = [];

    descMap.forEach((occurrences, cleanMota) => {
      if (occurrences.length <= 1) return; // TIẾT KIỆM TỐI ĐA BỘ NHỚ: Không thể mâu thuẫn nếu xuất hiện <= 1 lần

      const codeCounts = new Map<string, number[]>();
      occurrences.forEach(occ => {
        if (!codeCounts.has(occ.code)) {
          codeCounts.set(occ.code, []);
        }
        codeCounts.get(occ.code)!.push(occ.rowIdx);
      });

      if (codeCounts.size > 1) {
        if (descToCodes.length < 2000) {
          const originalText = occurrences[0].row[targetMota] || cleanMota;
          descToCodes.push({
            motaText: originalText,
            occurrences: occurrences.length,
            codes: Array.from(codeCounts.entries()).map(([code, rowIndices]) => ({
              code,
              count: rowIndices.length,
              rows: rowIndices
            }))
          });
        }
      }
    });

    // 2. CÙNG MÃ -> MÔ TẢ TRÁI QUY LUẬT KINH DOANH (Sản xuất vs Thương mại)
    const codeMap = new Map<string, Array<{ desc: string; rowIdx: number; row: any }>>();
    mainData.forEach((row, idx) => {
      if (!row || typeof row !== 'object') return;
      const code = normalizeSectorCode(row[targetManganh]);
      if (!code) return;
      const rawMota = String(row[targetMota] || "").trim();
      if (!rawMota) return;
      
      if (!codeMap.has(code)) {
        codeMap.set(code, []);
      }
      codeMap.get(code)!.push({ desc: rawMota, rowIdx: idx, row });
    });

    const codeToDescs: Array<{
      code: string;
      tenNganh: string;
      occurrences: number;
      conflicts: Array<{
        type: "SẢN XUẤT" | "THƯƠNG MẠI" | "DỊCH VỤ / KHÁC";
        descText: string;
        rowIdx: number;
      }>;
    }> = [];

    const tradeKeywords = ["bán", "mua", "thương mại", "đại lý", "cửa hàng", "phân phối", "wholesale", "retail", "shop", "siêu thị", "chợ"];
    const industrialKeywords = ["sản xuất", "gia công", "chế tạo", "làm mộc", "chế biến", "lắp ráp", "chế tác", "luyện kim", "nhà xưởng", "nhà máy"];

    codeMap.forEach((items, code) => {
      if (codeToDescs.length >= 2000) return; // Giới hạn hiển thị để tránh overload DOM render
      
      const typesSeen = new Set<string>();
      const conflictsList: any[] = [];

      items.forEach(item => {
        const lc = item.desc.toLowerCase();
        const hasTrade = tradeKeywords.some(kw => lc.includes(kw));
        const hasIndustrial = industrialKeywords.some(kw => lc.includes(kw));
        
        let type: "SẢN XUẤT" | "THƯƠNG MẠI" | "DỊCH VỤ / KHÁC" = "DỊCH VỤ / KHÁC";
        if (hasIndustrial) {
          type = "SẢN XUẤT";
        } else if (hasTrade) {
          type = "THƯƠNG MẠI";
        }
        typesSeen.add(type);
        conflictsList.push({
          type,
          descText: item.desc,
          rowIdx: item.rowIdx
        });
      });

      if (typesSeen.has("SẢN XUẤT") && typesSeen.has("THƯƠNG MẠI")) {
        const stdName = vsicRawData[code] || "Tên ngành chưa xác định";
        codeToDescs.push({
          code,
          tenNganh: stdName,
          occurrences: items.length,
          conflicts: conflictsList
        });
      }
    });

    return { descToCodes, codeToDescs };
  }, [mainData, mapping.mota, mapping.manganh, stdDescriptionCol, stdIndustryCol, activeTab]);

  // Phân trang dữ liệu hiển thị
  const paginatedData = useMemo(() => {
    const startIdx = (viewPage - 1) * pageSize;
    return filteredData.slice(startIdx, startIdx + pageSize);
  }, [filteredData, viewPage]);

  // Tổng số trang
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;

  // Lấy dữ liệu Excel dạng an toàn
  const getSafeExportData = (data: any[], selectedCols: string[]) => {
    return data.map(row => {
      const obj: any = {};
      selectedCols.forEach(col => {
        obj[col] = row[col] !== undefined ? row[col] : "";
      });
      return obj;
    });
  };

  // Tạm nghỉ bằng Promise cho việc Render mượt mà & hiển thị progress bar
  const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

  // 1. CHỨC NĂNG GHÉP NỐI DỮ LIỆU (Left Join)
  const handleMerge = async () => {
    if (leftData.length === 0 || rightData.length === 0) {
      alert("Vui lòng tải đủ cả 2 bảng dữ liệu Trái & Phải!");
      return;
    }
    if (!leftKey || !rightKey) {
      alert("Vui lòng chọn cột khóa liên kết cho cả 2 bảng!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu ghép nối dữ liệu...");

    await sleep(200);

    // Xử lý index phục vụ tra cứu nhanh
    const rightMap = new Map();
    const batchSize = Math.max(1, Math.floor(rightData.length / 10));
    
    for (let i = 0; i < rightData.length; i++) {
      const row = rightData[i];
      const kv = String(row[rightKey] || "").trim();
      if (kv) {
        rightMap.set(kv, row);
      }
      
      // Update %
      if (i % batchSize === 0 || i === rightData.length - 1) {
        const pct = Math.floor((i / rightData.length) * 40); // 40% tiến trình đầu
        setProgress(pct);
        setStatusMessage(`Đang lập chỉ mục bảng phải: ${i}/${rightData.length} dòng...`);
        await sleep(10);
      }
    }

    // Tiến hành ghép nối dữ liệu
    const mergedResults: any[] = [];
    const mainCols = Object.keys(leftData[0] || {});
    const rightCols = Object.keys(rightData[0] || {}).filter(c => c !== rightKey);

    const stepSize = Math.max(1, Math.floor(leftData.length / 10));

    for (let j = 0; j < leftData.length; j++) {
      const leftRow = leftData[j];
      const matchKey = String(leftRow[leftKey] || "").trim();
      const matchedRight = rightMap.get(matchKey);

      const mergedRow = { ...leftRow };
      rightCols.forEach(rc => {
        // Tránh trùng lặp tên cột, nếu trùng thì cộng đuôi _phải
        const finalColName = mainCols.includes(rc) ? `${rc}_Phai` : rc;
        mergedRow[finalColName] = matchedRight ? matchedRight[rc] : "";
      });

      mergedResults.push(mergedRow);

      if (j % stepSize === 0 || j === leftData.length - 1) {
        const pct = 40 + Math.floor((j / leftData.length) * 60); // 41% - 100% tiến trình tiếp theo
        setProgress(pct);
        setStatusMessage(`Đang ánh xạ & ghép dòng: ${j}/${leftData.length} dòng...`);
        await sleep(10);
      }
    }

    const mergedCols = Object.keys(mergedResults[0] || {});
    setMainData(mergedResults);
    setRawImportedData(mergedResults);
    setColumns(mergedCols);
    setFileName(`GhepNoi_${leftFileName}_vs_${rightFileName}.xlsx`);
    
    // Let the user select the unique ID column themselves
    setMapping({
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: ""
    });

    const initMergedConfigs = mergedCols.map(c => {
      return {
        originalName: c,
        use: true,
        newName: c,
        role: "" as any
      };
    });
    setCustomColConfigs(initMergedConfigs);

    setProgress(100);
    setStatusMessage(`Ghép nối thành công hoàn tất! Thu được ${mergedResults.length} dòng dữ liệu.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // 2. CHỨC NĂNG SO SÁNH CŨ - MỚI (DIFF)
  const handleCompare = async () => {
    if (oldData.length === 0 || newData.length === 0) {
      alert("Vui lòng tải đầy đủ tệp dữ liệu Cũ và Mới!");
      return;
    }
    if (!diffKey) {
      alert("Vui lòng chọn Cột Khóa định danh độc nhất (Unique Key)!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Khởi động so sánh dữ liệu...");
    await sleep(200);

    const oldMap = new Map();
    oldData.forEach(row => {
      if (!row || typeof row !== 'object') return;
      const k = String(row[diffKey] || "").trim();
      if (k) oldMap.set(k, row);
    });

    const newMap = new Map();
    newData.forEach(row => {
      if (!row || typeof row !== 'object') return;
      const k = String(row[diffKey] || "").trim();
      if (k) newMap.set(k, row);
    });

    const resultRows: any[] = [];
    const allKeys = Array.from(new Set([...oldMap.keys(), ...newMap.keys()]));
    const batchSize = Math.max(1, Math.floor(allKeys.length / 20));

    // Lấy tập hợp tất cả các cột của cả hai file
    const firstOldRow = oldData.find(r => r && typeof r === 'object') || {};
    const firstNewRow = newData.find(r => r && typeof r === 'object') || {};
    const oldCols = Object.keys(firstOldRow);
    const newCols = Object.keys(firstNewRow);
    const unionCols = Array.from(new Set([...oldCols, ...newCols])).filter(c => c !== diffKey);

    for (let i = 0; i < allKeys.length; i++) {
      const key = allKeys[i];
      const oldRow = oldMap.get(key);
      const newRow = newMap.get(key);

      const combined: any = { [diffKey]: key };

      if (oldRow && !newRow) {
        // Đã xóa
        unionCols.forEach(col => {
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = "";
        });
        combined["TrangThai_SoSanh"] = "❌ Đã xóa";
      } else if (!oldRow && newRow) {
        // Mới thêm
        unionCols.forEach(col => {
          combined[`${col}_Cu`] = "";
          combined[`${col}_Moi`] = newRow[col] || "";
        });
        combined["TrangThai_SoSanh"] = "✅ Mới thêm";
      } else {
        // Tồn tại ở cả 2 - cần kiểm tra thay đổi
        const changedCols: string[] = [];
        unionCols.forEach(col => {
          const valCu = String(oldRow[col] !== undefined ? oldRow[col] : "").trim();
          const valMoi = String(newRow[col] !== undefined ? newRow[col] : "").trim();
          
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = newRow[col] || "";

          if (valCu !== valMoi) {
            changedCols.push(col);
          }
        });

        if (changedCols.length > 0) {
          combined["TrangThai_SoSanh"] = `⚠️ Thay đổi: [${changedCols.join(", ")}]`;
        } else {
          combined["TrangThai_SoSanh"] = "💡 Không đổi";
        }
      }

      resultRows.push(combined);

      if (i % batchSize === 0 || i === allKeys.length - 1) {
        const pct = Math.floor((i / allKeys.length) * 100);
        setProgress(pct);
        setStatusMessage(`Đang so sánh đối chiếu: ${i}/${allKeys.length} dòng có khóa...`);
        await sleep(10);
      }
    }

    const compareCols = Object.keys(resultRows[0] || {});
    setMainData(resultRows);
    setRawImportedData(resultRows);
    setColumns(compareCols);
    setFileName(`SoSanhDiff_${oldFileName}_vs_${newFileName}.xlsx`);
    setMapping({
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: ""
    });

    const initCompareConfigs = compareCols.map(c => {
      return {
        originalName: c,
        use: true,
        newName: c,
        role: "" as any
      };
    });
    setCustomColConfigs(initCompareConfigs);

    setProgress(100);
    setStatusMessage(`So sánh thành công! Tìm thấy tổng cộng ${resultRows.length} khóa định danh.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // 3. CHỨC NĂNG TÁCH DỮ LIỆU THEO CỘT HOÀN TOÀN TỰ ĐỘNG (EXPORT ZIP)
  const handleSplitData = async () => {
    if (mainData.length === 0) {
      alert("Không có dữ liệu trong hệ thống! Vui lòng nạp tệp chính trước.");
      return;
    }
    if (!splitCol) {
      alert("Vui lòng chọn cột phân sách dữ liệu!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Khởi tạo tách dữ liệu...");
    await sleep(200);

    // Phân nhóm dòng theo giá trị cột tách
    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const val = String(row[splitCol] || "Rong").trim();
      const safeVal = val.replace(/[^a-zA-Z0-9_\-À-ỹ\s]/g, "");
      if (!groups.has(safeVal)) {
        groups.set(safeVal, []);
      }
      groups.get(safeVal)?.push(row);
    });

    const zip = new JSZip();
    const groupKeys = Array.from(groups.keys());
    const batchSize = Math.max(1, Math.floor(groupKeys.length / 10));

    for (let i = 0; i < groupKeys.length; i++) {
      const key = groupKeys[i];
      const rows = groups.get(key) || [];

      // Tạo một Worksheet Excel mới cho group cụ thể
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");

      // Build nhị phân của file Excel
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
      
      // Chuyển string binary thành ArrayBuffer
      const buf = new ArrayBuffer(wbout.length);
      const view = new Uint8Array(buf);
      for (let j = 0; j < wbout.length; j++) {
        view[j] = wbout.charCodeAt(j) & 0xFF;
      }

      // Nạp vào Zip
      zip.file(`Tach_File_${key}.xlsx`, buf);

      // Cập nhật % và status
      if (i % batchSize === 0 || i === groupKeys.length - 1) {
        const pct = Math.floor((i / groupKeys.length) * 90);
        setProgress(pct);
        setStatusMessage(`Đang nén dữ liệu cho khối '${key}': ${i}/${groupKeys.length} file...`);
        await sleep(10);
      }
    }

    setProgress(95);
    setStatusMessage("Đang đóng gói tệp nén ZIP tải xuống...");
    await sleep(300);

    // Hoàn thành xuất nén
    const content = await zip.generateAsync({ type: "blob" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(content);
    link.download = `TachFile_${splitCol}_${fileName.replace(/\.[^/.]+$/, "")}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setProgress(100);
    setStatusMessage(`Đã tách thành công thành ${groupKeys.length} tệp Excel riêng lẻ và tải về ZIP thành công.`);
    await sleep(400);
    setLoading(false);
  };

  // 4. CHỨC NĂNG TỔNG HỢP BÁO CÁO ĐỘNG (DYNAMIC PIVOT & AGGREGATE)
  const cleanNumberForSummary = (val: any): number => {
    if (val === null || val === undefined) return 0;
    if (typeof val === "number") return val;
    let str = String(val).trim();
    if (!str) return 0;
    
    // Khử nhanh các dấu phân cách hàng nghìn loại dấu phẩy hay chấm
    if (str.includes(",") && str.includes(".")) {
      const lastComma = str.lastIndexOf(",");
      const lastDot = str.lastIndexOf(".");
      if (lastComma > lastDot) {
        str = str.replace(/\./g, "").replace(/,/g, ".");
      } else {
        str = str.replace(/,/g, "");
      }
    } else if (str.includes(",")) {
      const parts = str.split(",");
      if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) {
        str = str.replace(/,/g, "");
      } else {
        str = str.replace(/,/g, ".");
      }
    } else if (str.includes(" ")) {
      str = str.replace(/\s/g, "");
    }
    const parsed = parseFloat(str.replace(/[^0-9.\-]/g, ""));
    return isNaN(parsed) ? 0 : parsed;
  };

  const addAggRule = () => {
    if (!newAggCol) {
      alert("Vui lòng chọn cột cần tính toán!");
      return;
    }
    // Tránh trùng lặp hoàn toàn
    if (aggRules.some(r => r.col === newAggCol && r.op === newAggOp)) {
      alert("Quy tắc này đã tồn tại!");
      return;
    }
    setAggRules([...aggRules, { col: newAggCol, op: newAggOp }]);
  };

  const removeAggRule = (idx: number) => {
    setAggRules(aggRules.filter((_, i) => i !== idx));
  };

  const handleRunSummary = async () => {
    if (mainData.length === 0) {
      alert("Yêu cầu nạp dữ liệu chính để tổng hợp.");
      return;
    }
    if (groupByCols.length === 0) {
      alert("Chọn ít nhất một cột để Gom Nhóm (Group By)!");
      return;
    }
    if (aggRules.length === 0) {
      alert("Cấu hình ít nhất một phép tính toán!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Đang tính toán tổng hợp dữ liệu...");
    await sleep(200);

    // Áp dụng thuật toán tổng hợp nhóm
    // Group dữ liệu dựa trên value tổ hợp của groupByCols, tự bẻ cấp 1, cấp 2 nếu chọn cột ảo
    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const compositeKeyObj: any = {};
      groupByCols.forEach(col => {
        const selectedManganhCol = pivotManganhCol || mapping.manganh;
        if (col === "_virtual_sector_cap2" && selectedManganhCol) {
          const mng = normalizeSectorCode(row[selectedManganhCol]);
          const sec2Code = mng ? mng.slice(0, 2) : "";
          const sec2Name = vsicRawData[sec2Code] || "Ngành cấp 2 chưa định nghĩa";
          compositeKeyObj["Ngành Cấp 2"] = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định";
        } else if (col === "_virtual_sector_cap1" && selectedManganhCol) {
          const mng = normalizeSectorCode(row[selectedManganhCol]);
          let sec1Code = "";
          if (mng) {
            if (/^[a-zA-Z]$/.test(mng)) {
              sec1Code = mng.toUpperCase();
            } else {
              const sec2Code = mng.slice(0, 2);
              sec1Code = getParentSectorCode(sec2Code) || "";
            }
          }
          const sec1Name = vsicRawData[sec1Code] || "Ngành cấp 1 chưa định nghĩa";
          compositeKeyObj["Ngành Cấp 1"] = sec1Code ? `${sec1Code} - ${sec1Name}` : "Chưa xác định";
        } else {
          compositeKeyObj[col] = row[col] !== undefined ? String(row[col]) : "[Rỗng]";
        }
      });
      const keyStr = JSON.stringify(compositeKeyObj);
      if (!groups.has(keyStr)) {
        groups.set(keyStr, []);
      }
      groups.get(keyStr)?.push(row);
    });

    const summaryRows: any[] = [];
    const keysArray = Array.from(groups.keys());
    const batchSize = Math.max(1, Math.floor(keysArray.length / 5));

    for (let k = 0; k < keysArray.length; k++) {
      const keyStr = keysArray[k];
      const rows = groups.get(keyStr) || [];
      const groupValueObj = JSON.parse(keyStr);

      const resultRow: any = { ...groupValueObj };

      // Chạy các quy tắc tính toán cho tổ hợp
      aggRules.forEach(rule => {
        const { col, op } = rule;
        const colValues = rows.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== "");
        const numValues = colValues.map(v => cleanNumberForSummary(v));

        let calcVal: number = 0;
        const colHeader = `${col}_${op}`;

        if (op === "count") {
          calcVal = colValues.length;
        } else if (op === "nunique") {
          calcVal = new Set(colValues).size;
        } else if (op === "sum") {
          calcVal = numValues.reduce((sum, v) => sum + v, 0);
        } else if (op === "mean") {
          calcVal = numValues.length > 0 ? numValues.reduce((sum, v) => sum + v, 0) / numValues.length : 0;
          calcVal = Math.round(calcVal * 100) / 100;
        } else if (op === "min") {
          calcVal = numValues.length > 0 ? Math.min(...numValues) : 0;
        } else if (op === "max") {
          calcVal = numValues.length > 0 ? Math.max(...numValues) : 0;
        }

        resultRow[colHeader] = calcVal;
      });

      resultRow["So_Luong_DN_Trong_Nhom"] = rows.length;
      summaryRows.push(resultRow);

      if (k % batchSize === 0 || k === keysArray.length - 1) {
        const pct = Math.floor((k / keysArray.length) * 100);
        setProgress(pct);
        setStatusMessage(`Tính toán chỉ số gom nhóm: ${k}/${keysArray.length} tổ hợp...`);
        await sleep(10);
      }
    }

    setMainData(summaryRows);
    setColumns(Object.keys(summaryRows[0] || {}));
    setFileName(`BaoCaoTongHop_${fileName}`);
    setProgress(100);
    setStatusMessage(`Báo cáo tổng hợp nhóm hoàn tất thành công! Tạo thành ${summaryRows.length} dòng báo cáo.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // PHÂN HỆ 1: TỔNG HỢP BÁO CÁO THEO ĐẦU NGÀNH CẤP 1 & CẤP 2 VSIC (MỚI ĐỘC LẬP)
  const handleExportT2Excel = () => {
    if (t2ReportData.length === 0) {
      alert(`Chưa có dữ liệu báo cáo ngành cấp ${t2ReportLevel} để xuất!`);
      return;
    }
    setLoading(true);
    setStatusMessage(`Đang chuẩn bị tệp Excel Báo cáo Ngành Cấp ${t2ReportLevel}...`);
    setTimeout(() => {
      try {
        const ws = XLSX.utils.json_to_sheet(t2ReportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `Báo cáo ngành cấp ${t2ReportLevel}`);
        XLSX.writeFile(wb, `Bao_Cao_Nganh_Cap${t2ReportLevel}_${t2AggMethod === "sum" ? "Tong" : "TB"}.xlsx`);
        setStatusMessage(`Tải xuống Báo cáo Ngành Cấp ${t2ReportLevel} thành công!`);
      } catch (e: any) {
        alert(`Lỗi xuất Excel ngành cấp ${t2ReportLevel}: ` + e.message);
      } finally {
        setLoading(false);
      }
    }, 200);
  };

  const handleCalcLevelSummary = async (level: 1 | 2) => {
    if (mainData.length === 0) {
      alert("Không tìm thấy dữ liệu nguồn chính! Vui lòng nạp tệp chính trước.");
      return;
    }
    if (!t2IndustryCol) {
      alert("Vui lòng chọn cột chứa Mã ngành ở bộ chọn!");
      return;
    }
    if (t2MetricCols.length === 0) {
      alert("Vui lòng tích chọn ít nhất một cột số liệu (Doanh thu, Lao động...) để tổng hợp!");
      return;
    }

    setT2ReportLevel(level);
    setLoading(true);
    setProgress(10);
    setStatusMessage(`Đang quét dữ liệu nguồn và tách lọc để lấy mã ngành cấp ${level}...`);
    await sleep(200);

    // Grouping map
    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const rawVal = row[t2IndustryCol];
      const normalized = normalizeSectorCode(rawVal);
      
      let finalCode = "";
      if (level === 2) {
        // Lấy 2 số đầu của mã ngành làm ngành cấp 2 (VSIC luôn gồm 2 chữ số đầu đại diện cấp 2)
        finalCode = normalized ? normalized.slice(0, 2) : "";
      } else {
        // level === 1 (Ngành cấp 1 dạng chữ cái A..U)
        if (normalized) {
          if (/^[a-zA-Z]$/.test(normalized)) {
            finalCode = normalized.toUpperCase();
          } else {
            const sec2Code = normalized.slice(0, 2);
            finalCode = getParentSectorCode(sec2Code) || "";
          }
        }
      }
      const grpKey = finalCode || "CHUA_XAC_DINH";

      if (!groups.has(grpKey)) {
        groups.set(grpKey, []);
      }
      groups.get(grpKey)!.push(row);
    });

    setProgress(50);
    setStatusMessage(`Đang đối chiếu ngành cấp ${level} nhúng trong code và cộng dồn lũy kế chỉ số dữ liệu...`);
    await sleep(150);

    const reportRows: any[] = [];
    const sortedKeys = Array.from(groups.keys()).sort((a, b) => {
      if (a === "CHUA_XAC_DINH") return 1;
      if (b === "CHUA_XAC_DINH") return -1;
      return a.localeCompare(b);
    });

    sortedKeys.forEach((key, idx) => {
      const rowsInGroup = groups.get(key) || [];
      const count = rowsInGroup.length;
      
      let levelName = `Mã ngành cấp ${level} trống hoặc không hợp lệ`;
      if (key !== "CHUA_XAC_DINH") {
        levelName = vsicRawData[key] || `Ngành cấp ${level} chưa định nghĩa chuẩn mực`;
      }

      const reportRow: any = {
        "STT": idx + 1,
        [`Mã ngành cấp ${level}`]: key === "CHUA_XAC_DINH" ? "Chưa xác định" : key,
        [`Tên ngành cấp ${level}`]: levelName,
        "Số đơn vị (DN)": count
      };

      // Aggregate each selected numerical column
      t2MetricCols.forEach(metricCol => {
        let sum = 0;
        let validCount = 0;

        rowsInGroup.forEach(r => {
          const rawNum = r[metricCol];
          if (rawNum !== undefined && rawNum !== null && rawNum !== "") {
            const num = parseFloat(String(rawNum).replace(/[^0-9.\-]/g, ""));
            if (!isNaN(num)) {
              sum += num;
              validCount++;
            }
          }
        });

        const fieldKey = `${t2AggMethod === "sum" ? "Tổng" : "Trung bình"} ${metricCol}`;
        if (t2AggMethod === "sum") {
          reportRow[fieldKey] = Math.round(sum * 100) / 100;
        } else {
          const avg = validCount > 0 ? (sum / validCount) : 0;
          reportRow[fieldKey] = Math.round(avg * 100) / 100;
        }
      });

      reportRows.push(reportRow);
    });

    // Tạo dòng Tổng Cộng (Grand Total Row)
    if (reportRows.length > 0) {
      const totalRow: any = {
        "STT": "",
        [`Mã ngành cấp ${level}`]: "TỔNG CỘNG LŨY KẾ",
        [`Tên ngành cấp ${level}`]: `Hệ thống tổng quy nạp thành ${sortedKeys.length} nhóm ngành cấp ${level}`,
        "Số đơn vị (DN)": mainData.length
      };

      t2MetricCols.forEach(metricCol => {
        let sumAll = 0;
        let rowValidCount = 0;

        mainData.forEach(r => {
          const rawVal = r[metricCol];
          if (rawVal !== undefined && rawVal !== null && rawVal !== "") {
            const num = parseFloat(String(rawVal).replace(/[^0-9.\-]/g, ""));
            if (!isNaN(num)) {
              sumAll += num;
              rowValidCount++;
            }
          }
        });

        const fieldKey = `${t2AggMethod === "sum" ? "Tổng" : "Trung bình"} ${metricCol}`;
        if (t2AggMethod === "sum") {
          totalRow[fieldKey] = Math.round(sumAll * 100) / 100;
        } else {
          const avgAll = rowValidCount > 0 ? (sumAll / rowValidCount) : 0;
          totalRow[fieldKey] = Math.round(avgAll * 100) / 100;
        }
      });

      reportRows.push(totalRow);
    }

    setT2ReportData(reportRows);
    if (reportRows.length > 0) {
      setT2ReportCols(Object.keys(reportRows[0]));
    }

    setProgress(100);
    setStatusMessage(`Tổng hợp báo cáo ngành cấp ${level} hoàn tất thành công!`);
    await sleep(300);
    setLoading(false);
  };

  // PHÂN HỆ 2: CHUẨN HÓA & KHỚP MÃ NGÀNH VSIC CẤP 5 (MỚI ĐỘC LẬP)
  const handleStandardizeSectorsAndMatch = async () => {
    if (mainData.length === 0) {
      alert("Không tìm thấy dữ liệu nguồn chính! Vui lòng nạp tệp chính trước.");
      return;
    }
    if (!stdIndustryCol) {
      alert("Vui lòng chọn cột chứa Mã ngành cấp 5 cần chuẩn hóa!");
      return;
    }

    setLoading(true);
    setProgress(5);
    setStatusMessage("Đang quét danh sách mã ngành và tiến hành chuẩn hóa mẫu tự liên hợp...");
    await sleep(200);

    try {
      let validCount = 0;
      let invalidCount = 0;
      let conflictCount = 0;
      const anomalies: any[] = [];

      // Tạo mảng bản ghi mới bổ sung cột của "Tên Ngành Chuẩn VSIC" và "Trạng Thái Đối Chiếu VSIC" sát cạnh cột gốc
      const updatedRows = mainData.map((row, idx) => {
        if (!row || typeof row !== 'object') return row;
        const rawCode = row[stdIndustryCol];
        const rawDesc = stdDescriptionCol ? String(row[stdDescriptionCol] || "") : "";
        
        const cleanCode = normalizeSectorCode(rawCode);
        const lookupResult = lookupSectorNameWithFallback(cleanCode);
        const isExistInVSIC = lookupResult.level > 0;
        const stdName = lookupResult.name;

        if (isExistInVSIC) {
          validCount++;
        } else {
          invalidCount++;
        }

        // Đối chiếu quy luật logic hoạt động mô tả & mã ngành để phát hiện mâu thuẫn lệch vai trò
        let auditStatus = lookupResult.exactMatched ? "✅ Đạt chuẩn VSIC quốc gia" : "✅ Khớp quy nạp cấp học";
        let anomalyReason = "";

        if (!isExistInVSIC) {
          auditStatus = "❌ Mã lỗi / Chưa thuộc VSIC";
          anomalyReason = `Mã ngành "${rawCode}" không tìm thấy trong danh mục hệ thống phân cấp VSIC quốc gia`;
        } else if (rawDesc.trim() !== "") {
          const descLow = rawDesc.toLowerCase();
          const stdLow = stdName.toLowerCase();

          // 1. Lệch hạch toán: Nông nghiệp vs Phục vụ thương mại
          const hasFeederWords = descLow.includes("trồng") || descLow.includes("nuôi") || descLow.includes("bắt") || descLow.includes("thu hoạch") || descLow.includes("đánh bắt");
          const hasTradeWords = descLow.includes("bán buôn") || descLow.includes("bán lẻ") || descLow.includes("môi giới") || descLow.includes("đại lý") || descLow.includes("thương mại");
          
          const stdIsFeeder = stdLow.includes("nông nghiệp") || stdLow.includes("lâm nghiệp") || stdIsFeederWord(stdLow);
          const stdIsTrade = stdLow.includes("bán buôn") || stdLow.includes("bán lẻ") || stdLow.includes("thương mại");

          if (hasTradeWords && stdIsFeeder) {
            auditStatus = "⚠️ Nghi ngờ lệch mã (Khai mâu thuẫn giữa Phân phối thương mại và sản xuất nông nghiệp)";
            anomalyReason = `Mô tả ghi thương mại (${rawDesc}) nhưng lại gán mã thuộc ngành trồng trọt/chăn nuôi sản xuất trực tiếp (${cleanCode} - ${stdName})`;
            conflictCount++;
          } else if (hasFeederWords && stdIsTrade) {
            auditStatus = "⚠️ Nghi ngờ lệch mã (Khai mâu thuẫn giữa Tự sản tự tiêu nông nghiệp và phân phối đại lý)";
            anomalyReason = `Mô tả ghi trồng trọt, khai mỏ (${rawDesc}) nhưng mã ngành lại gán đại lý bán buôn, dịch vụ phân phối (${cleanCode} - ${stdName})`;
            conflictCount++;
          }

          // 2. Chế biến sản xuất vs Dịch vụ ăn uống, xây dựng
          const hasManufacture = descLow.includes("sản xuất") || descLow.includes("chế tạo") || descLow.includes("gia công") || descLow.includes("lắp đặt");
          const hasService = descLow.includes("ăn uống") || descLow.includes("nhà hàng") || descLow.includes("quán") || descLow.includes("giáo dục") || descLow.includes("dịch vụ");

          const stdIsManufacture = stdLow.includes("sản xuất") || stdLow.includes("chế biến") || stdLow.includes("chế tạo");
          const stdIsService = stdLow.includes("ăn uống") || stdLow.includes("nhà hàng") || stdLow.includes("giáo dục") || stdLow.includes("dịch vụ");

          if (hasManufacture && stdIsService) {
            auditStatus = "⚠️ Nghi ngờ lệch mã (Sản xuất gia công vs Dịch vụ ăn uống hoặc đào tạo)";
            anomalyReason = `Mô tả ghi chế tạo gia công (${rawDesc}) nhưng mã ngành lại thuộc về cung ứng ăn uống hoặc dịch vụ dân sinh (${cleanCode} - ${stdName})`;
            conflictCount++;
          } else if (hasService && stdIsManufacture) {
            auditStatus = "⚠️ Nghi ngờ lệch mã (Cung ứng dịch vụ vs Chế biến sản xuất công nghiệp)";
            anomalyReason = `Mô tả ghi phục vụ ẩm thực, giáo dục (${rawDesc}) nhưng mã ngành lại hạch toán vào sản xuất công nghiệp nặng/nhẹ (${cleanCode} - ${stdName})`;
            conflictCount++;
          }
        }

        if (anomalyReason) {
          anomalies.push({
            dongSTT: idx + 1,
            maDN: row["Mã Số Thuế"] || row["MaST"] || `Bản ghi số ${idx + 1}`,
            maGoc: rawCode,
            motaGoc: rawDesc,
            nganhChuan: stdName || "(Thất bại khi tra cứu)",
            phanTichloi: anomalyReason
          });
        }

        // Xây dựng Bản ghi mới co cụm, bơm cột Tên Ngành Chuẩn VSIC và Trạng Thái Đối Chiếu VSIC nằm ngay bên cạnh cột Mô Tả Hoạt Động / Mã Ngành để dễ đối chiếu
        const flexRow: any = {};
        Object.keys(row).forEach(key => {
          flexRow[key] = row[key];
          if (key === stdDescriptionCol) {
            flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ KHÔNG TÌM THẤY MÃ TRONG VSIC";
            flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
          }
        });

        // Nếu không khớp được vị trí cột mô tả thì tự chêm cột mới vào kế cột Mã ngành
        if (flexRow["Tên Ngành Chuẩn VSIC"] === undefined) {
          Object.keys(row).forEach(key => {
            flexRow[key] = row[key];
            if (key === stdIndustryCol) {
              flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ KHÔNG TÌM THẤY MÃ TRONG VSIC";
              flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
            }
          });
        }

        // Nếu vẫn thiếu do trùng cấu hình dặc biệt
        if (flexRow["Tên Ngành Chuẩn VSIC"] === undefined) {
          flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ KHÔNG TÌM THẤY MÃ TRONG VSIC";
          flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
        }

        return flexRow;
      });

      const newCols = Object.keys(updatedRows[0] || {});
      setMainData(updatedRows);
      setColumns(newCols);
      setStdReportAnomalies(anomalies);
      setStdMatchStats({
        total: updatedRows.length,
        valid: validCount,
        invalid: invalidCount,
        conflicts: conflictCount
      });

      // Tự sao lưu vĩnh viễn vào hệ thống
      autoSaveSession(updatedRows, rawImportedData, newCols, fileName, mapping, customColConfigs);

      setProgress(100);
      setStatusMessage("Chuẩn hóa mã ngành thành công! Đã bổ sung cột trực tiếp bên cạnh cột của bạn.");
      await sleep(350);
      setLoading(false);
      
      alert(`Chuẩn hóa hoàn tất!\n- Tổng cộng: ${updatedRows.length} dòng\n- Khớp VSIC: ${validCount} dòng\n- Lệch chuẩn: ${invalidCount} dòng\n- Nghi ngờ bất nhất mã vs mô tả: ${conflictCount} dòng.`);
    } catch (err: any) {
      alert("Lỗi quá trình chuẩn hóa VSIC: " + err.message);
      setLoading(false);
    }
  };

  const stdIsFeederWord = (txt: string): boolean => {
    return txt.includes("trồng") || txt.includes("nuôi") || txt.includes("thủy sản") || txt.includes("mỏ") || txt.includes("lâm nghiệp");
  };

  // PHÂN HỆ ĐỐI CHIẾU CHÉO SONG SONG 2 CỘT TÙY CHỌN (Yêu cầu người dùng)
  const handleCrossColumnCompare = async () => {
    if (mainData.length === 0) {
      alert("Không tìm thấy dữ liệu nguồn chính! Vui lòng nạp tệp chính trước bản ghi.");
      return;
    }
    if (!crossCompareColA || !crossCompareColB) {
      alert("Vui lòng chọn đầy đủ cả 2 cột cần so khớp, đối chiếu!");
      return;
    }
    if (crossCompareColA === crossCompareColB) {
      alert("Vui lòng chọn 2 cột có tên khác nhau để so sánh đối chiếu!");
      return;
    }

    setLoading(true);
    setProgress(15);
    setStatusMessage(`Đang tiến hành đối chiếu song song hai cột: [${crossCompareColA}] và [${crossCompareColB}]...`);
    await sleep(250);

    try {
      let matchCount = 0;
      let mismatchCount = 0;
      const anomalies: any[] = [];

      const updatedRows = mainData.map((row, idx) => {
        if (!row || typeof row !== 'object') return row;
        const valA = row[crossCompareColA] !== undefined && row[crossCompareColA] !== null ? String(row[crossCompareColA]).trim() : "";
        const valB = row[crossCompareColB] !== undefined && row[crossCompareColB] !== null ? String(row[crossCompareColB]).trim() : "";

        let isMatch = false;
        let explanation = "";

        if (crossCompareRule === "exact") {
          isMatch = valA === valB;
          if (!isMatch) {
            explanation = `Ký tự khác hoàn toàn (so sánh chuẩn xác cả chữ hoa/thường, dấu cách)`;
          }
        } else if (crossCompareRule === "normalize") {
          const cleanA = valA.toLowerCase().replace(/\s+/g, " ");
          const cleanB = valB.toLowerCase().replace(/\s+/g, " ");
          isMatch = cleanA === cleanB;
          if (!isMatch) {
            explanation = `Chuỗi văn bản gốc không trùng nhau (sau khi đã chuẩn hóa khoảng trắng & bỏ viết hoa)`;
          }
        } else if (crossCompareRule === "sector_code") {
          const codeA = valA.replace(/\D/g, "");
          const codeB = valB.replace(/\D/g, "");
          if (codeA === codeB && codeA !== "") {
            isMatch = true;
          } else if (codeA !== "" && codeB !== "") {
            isMatch = codeA.startsWith(codeB) || codeB.startsWith(codeA);
            if (isMatch) {
              explanation = `Khấu chuẩn quy nạp phân cấp theo logic cha-con (VD: ${valA} so với ${valB})`;
            } else {
              explanation = `Mã ngành hoàn toàn khác biệt nhóm phân cấp (VD: ${valA} so với ${valB})`;
            }
          } else {
            isMatch = valA === valB;
            if (!isMatch) {
              explanation = `Mã bị trống hoặc không thể phân giải số ngành hơp chuẩn`;
            }
          }
        } else if (crossCompareRule === "substring") {
          const descLowA = valA.toLowerCase();
          const descLowB = valB.toLowerCase();
          isMatch = descLowA.includes(descLowB) || descLowB.includes(descLowA);
          if (isMatch) {
            explanation = `Thỏa mãn: Một giá trị chứa phụ đề / từ khóa của giá trị còn lại`;
          } else {
            explanation = `Không có bất kỳ cụm từ khóa liên đới chéo nhau`;
          }
        }

        if (isMatch) {
          matchCount++;
        } else {
          mismatchCount++;
          anomalies.push({
            dongSTT: idx + 1,
            maDN: row["Mã Số Thuế"] || row["MaST"] || row["Số GPKD"] || `Bản ghi số ${idx + 1}`,
            valA: valA || "(Không có dữ liệu)",
            valB: valB || "(Không có dữ liệu)",
            reason: explanation
          });
        }

        const flexRow: any = {};
        const colCompareResult = `Đối Chiếu [${crossCompareColA}] vs [${crossCompareColB}]`;
        const colCompareFlag = `Đánh Dấu Lệch [${crossCompareColA}] vs [${crossCompareColB}]`;

        Object.keys(row).forEach(key => {
          flexRow[key] = row[key];
          if (key === crossCompareColB) {
            flexRow[colCompareResult] = isMatch ? "✅ TRÙNG KHỚP" : "❌ LỆCH BẤT NHẤT";
            flexRow[colCompareFlag] = isMatch ? "" : "⚠️ SAI LỆCH CẦN SỬA";
          }
        });

        if (flexRow[colCompareResult] === undefined) {
          flexRow[colCompareResult] = isMatch ? "✅ TRÙNG KHỚP" : "❌ LỆCH BẤT NHẤT";
          flexRow[colCompareFlag] = isMatch ? "" : "⚠️ SAI LỆCH CẦN SỬA";
        }

        return flexRow;
      });

      const newCols = Object.keys(updatedRows[0] || {});
      setMainData(updatedRows);
      setColumns(newCols);
      setCrossCompareAnomalies(anomalies);
      setCrossCompareStats({
        total: updatedRows.length,
        matchCount: matchCount,
        mismatchCount: mismatchCount
      });

      autoSaveSession(updatedRows, rawImportedData, newCols, fileName, mapping, customColConfigs);

      setProgress(100);
      setStatusMessage(`Đối chiếu chéo hoàn tất! Phát hiện ${mismatchCount} lỗi lệch.`);
      await sleep(350);
      setLoading(false);

      alert(`Đối chiếu hoàn tất!\n- Tổng cộng: ${updatedRows.length} dòng\n- Khớp nhau: ${matchCount} dòng\n- Sai lệch/Mâu thuẫn: ${mismatchCount} dòng.\nCác cột báo cáo mới đã được tự động thêm vào bảng tính của bạn.`);
    } catch (err: any) {
      alert("Lỗi quá trình đối chiếu: " + err.message);
      setLoading(false);
    }
  };

  // Xuất file báo cáo tổng hợp ngành và xã ra Excel
  const handleExportQuickReport = () => {
    if (quickReportResultRows.length === 0) {
      alert("Chưa có dữ liệu báo cáo để xuất!");
      return;
    }
    try {
      const ws = XLSX.utils.json_to_sheet(quickReportResultRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, `Báo Cáo Ngành Cấp ${quickReportLevel}`);
      XLSX.writeFile(wb, `BaoCao_TongHop_NganhCap${quickReportLevel}_Va_Xa_${reportType}.xlsx`);
    } catch (e: any) {
      alert("Lỗi quá trình xuất Excel: " + e.message);
    }
  };

  // BÊ NGÀNH CẤP 1 & CẤP 2 SANG CỘT MỚI (Lấy trực tiếp từ Danh mục đã nạp trong bộ nhớ)
  const handleAppendSectorsToMainData = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện!");
      return;
    }
    const targetManganh = quickReportManganhCol || mapping.manganh;
    if (!targetManganh) {
      alert("Vui lòng chỉ định cột chứa Mã ngành ở bộ chọn!");
      return;
    }

    setLoading(true);
    setProgress(15);
    setStatusMessage("Đang đối chiếu và nạp thông tin Ngành cấp 1 & 2 trực tiếp từ bộ nhớ...");
    await sleep(250);

    try {
      let matchedCount = 0;
      const updatedRows = mainData.map((row) => {
        if (!row || typeof row !== "object") return row;
        const rawCode = row[targetManganh];
        const mng = normalizeSectorCode(rawCode);
        
        // Ngành Cấp 2 (2 chữ số đầu)
        const sec2Code = mng ? mng.slice(0, 2) : "";
        const sec2Name = vsicRawData[sec2Code] || "Chưa định nghĩa chuẩn";
        
        // Ngành Cấp 1 (Lĩnh vực lớn, chữ cái A-U quy thuộc)
        let sec1Code = "";
        if (mng) {
          if (/^[a-zA-Z]$/.test(mng)) {
            sec1Code = mng.toUpperCase();
          } else {
            sec1Code = getParentSectorCode(sec2Code) || "";
          }
        }
        const sec1Name = vsicRawData[sec1Code] || "Chưa định nghĩa chuẩn";

        if (sec2Code && vsicRawData[sec2Code]) {
          matchedCount++;
        }

        // Tạo bản ghi mới, chèn trực tiếp các cột ngành Cấp 1 & Cấp 2 ngay sát bên cạnh cột Mã Ngành để người dùng có thể dễ dàng kiểm soát
        const flexRow: any = {};
        Object.keys(row).forEach((key) => {
          flexRow[key] = row[key];
          if (key === targetManganh) {
            flexRow["Mã Ngành Cấp 1"] = sec1Code;
            flexRow["Ngành Cấp 1"] = sec1Name;
            flexRow["Mã Ngành Cấp 2"] = sec2Code;
            flexRow["Ngành Cấp 2"] = sec2Name;
          }
        });

        // Đảm bảo không bị thiếu trong trường hợp đặc sản
        if (flexRow["Mã Ngành Cấp 1"] === undefined) {
          flexRow["Mã Ngành Cấp 1"] = sec1Code;
          flexRow["Ngành Cấp 1"] = sec1Name;
          flexRow["Mã Ngành Cấp 2"] = sec2Code;
          flexRow["Ngành Cấp 2"] = sec2Name;
        }

        return flexRow;
      });

      const newCols = Object.keys(updatedRows[0] || {});
      setMainData(updatedRows);
      setColumns(newCols);
      
      // Tự sao lưu vĩnh viễn
      autoSaveSession(updatedRows, rawImportedData, newCols, fileName, mapping, customColConfigs);

      setProgress(100);
      setStatusMessage("Nạp và chêm cột phân cấp Ngành cấp 1 & 2 thành công!");
      await sleep(350);
      setLoading(false);

      alert(`Hoàn thành bê ngành cấp 1 & 2!\n- Tổng cộng: ${updatedRows.length} dòng dữ liệu\n- Đối khớp danh mục thành công: ${matchedCount} dòng\n- Các cột mới đã được tạo trực tiếp sát bên cạnh cột [${targetManganh}] ở bảng tính gốc của bạn.`);
    } catch (err: any) {
      alert("Lỗi quá trình bê cột ngành: " + err.message);
      setLoading(false);
    }
  };

  // 5. CHỨC NĂNG BÁO CÁO NHANH THEO PHÂN CẤP NGÀNH & XÃ CHUẨN XÁC
  const handleQuickReport = async (level: number) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi chạy báo cáo nhanh.");
      return;
    }
    const targetManganh = quickReportManganhCol || mapping.manganh;
    const targetXa = quickReportXaCol || mapping.xa;
    const targetDoanhThu = quickReportDoanhThuCol || mapping.doanhthu;
    const targetLaoDong = quickReportLaoDongCol || mapping.laodong;

    if (!targetManganh) {
      alert("Vui lòng chỉ định cột chứa Mã ngành ở bộ chọn!");
      return;
    }
    if (!targetXa) {
      alert("Vui lòng chỉ định cột chứa Xã / Địa bàn ở bộ chọn!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang tạo báo cáo nhanh Ngành Cấp ${level} kết hợp Xã...`);
    await sleep(200);

    try {
      // Gom dữ liệu mỏng và phân tách bằng bộ nhớ mã ngành chuẩn
      const processedData = mainData.map(row => {
        if (!row || typeof row !== 'object') {
          return {
            _temNganhCap: "Chưa xác định",
            _tempXa: "Khác"
          };
        }
        const mng = normalizeSectorCode(row[targetManganh]);
        
        let tenNganhLabel = "";
        if (level === 2) {
          // Tách 2 chữ số đầu của cột mã ngành do người dùng chỉ định
          const sec2Code = mng ? mng.slice(0, 2) : "";
          // Tra cứu trong bộ nhớ tên của mã ngành cấp 2
          const sec2Name = vsicRawData[sec2Code] || "Ngành cấp 2 chưa định nghĩa";
          tenNganhLabel = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định - Ngành cấp 2 chưa định nghĩa";
        } else {
          // level === 1 (Quá trình quy nạp cấp 1)
          let sec1Code = "";
          if (mng) {
            if (/^[a-zA-Z]$/.test(mng)) {
              sec1Code = mng.toUpperCase();
            } else {
              const sec2Code = mng.slice(0, 2);
              sec1Code = getParentSectorCode(sec2Code) || "";
            }
          }
          const sec1Name = vsicRawData[sec1Code] || "Ngành cấp 1 chưa định nghĩa";
          tenNganhLabel = sec1Code ? `${sec1Code} - ${sec1Name}` : "Chưa xác định - Ngành cấp 1 chưa định nghĩa";
        }

        return {
          ...row,
          _temNganhCap: tenNganhLabel,
          _tempXa: String(row[targetXa] || "Khác").trim()
        };
      });

      let finalReportRows: any[] = [];

      if (reportType === "pivot") {
        setStatusMessage("Đang tiến hành xoay (Pivot) gom nhóm theo từng Ngành Kinh Tế làm cột...");
        await sleep(150);

        const communes = Array.from(new Set(processedData.map(r => r._tempXa))).sort();
        const sectorLabels = Array.from(new Set(processedData.map(r => r._temNganhCap))).sort();

        communes.forEach((commune, cIdx) => {
          const communeObj: any = {
            "Địa_Bàn_Xã": commune
          };

          let totalCommuneDN = 0;
          let totalCommuneDoanhThu = 0;
          let totalCommuneLaoDong = 0;

          sectorLabels.forEach(sector => {
            const matchedRows = processedData.filter(r => r._tempXa === commune && r._temNganhCap === sector);
            let sumDoanhThu = 0;
            let sumLaoDong = 0;

            matchedRows.forEach(r => {
              if (targetDoanhThu) {
                const val = parseFloat(String(r[targetDoanhThu]).replace(/[^0-9.\-]/g, ""));
                if (!isNaN(val)) sumDoanhThu += val;
              }
              if (targetLaoDong) {
                const val = parseFloat(String(r[targetLaoDong]).replace(/[^0-9.\-]/g, ""));
                if (!isNaN(val)) sumLaoDong += val;
              }
            });

            // Hiển thị kề nhau 2 cột Tổng Doanh Thu và Tổng Lao Động cho đúng ngành
            communeObj[`${sector} - Tổng Doanh Thu`] = Math.round(sumDoanhThu * 100) / 100;
            communeObj[`${sector} - Tổng Lao Động`] = Math.round(sumLaoDong);

            totalCommuneDN += matchedRows.length;
            totalCommuneDoanhThu += sumDoanhThu;
            totalCommuneLaoDong += sumLaoDong;
          });

          communeObj["Số_DN_Địa_Phương"] = totalCommuneDN;
          communeObj["Tổng_Doanh_Thu_Địa_Phương"] = Math.round(totalCommuneDoanhThu * 100) / 100;
          communeObj["Tổng_Lao_Động_Địa_Phương"] = Math.round(totalCommuneLaoDong);

          finalReportRows.push(communeObj);
        });
      } else {
        // Gom nhóm phẳng truyền thống
        const groups = new Map<string, any[]>();
        processedData.forEach(row => {
          const key = JSON.stringify({ Ngành: row._temNganhCap, Xã: row._tempXa });
          if (!groups.has(key)) groups.set(key, []);
          groups.get(key)?.push(row);
        });

        const keys = Array.from(groups.keys());
        keys.forEach(keyStr => {
          const dims = JSON.parse(keyStr);
          const rowsObj = groups.get(keyStr) || [];

          let sumDoanhThu = 0;
          let sumLaoDong = 0;

          rowsObj.forEach(r => {
            if (targetDoanhThu) {
              const val = parseFloat(String(r[targetDoanhThu]).replace(/[^0-9.\-]/g, ""));
              if (!isNaN(val)) sumDoanhThu += val;
            }
            if (targetLaoDong) {
              const val = parseFloat(String(r[targetLaoDong]).replace(/[^0-9.\-]/g, ""));
              if (!isNaN(val)) sumLaoDong += val;
            }
          });

          finalReportRows.push({
            [`Ngành_Cấp_${level}`]: dims.Ngành,
            "Địa_Bàn_Xã": dims.Xã,
            "Số_Lượng_Doanh_Nghiệp": rowsObj.length,
            "Tổng_Doanh_Thu_Tích_Lũy": Math.round(sumDoanhThu * 100) / 100,
            "Tổng_Lao_Động_Hợp_Lực": Math.round(sumLaoDong)
          });
        });
      }

      setQuickReportResultRows(finalReportRows);
      setQuickReportResultCols(Object.keys(finalReportRows[0] || {}));
      setQuickReportLevel(level);

      setProgress(100);
      setStatusMessage(`Tạo báo cáo nhanh ${reportType === "pivot" ? "xoay cột Pivot" : "dạng phẳng"} Ngành Cấp ${level} thành công!`);
      await sleep(350);
      setLoading(false);
      
      alert("Tạo báo cáo nhanh hoàn tất! Dữ liệu đã được nạp gọn gàng và hiển thị bảng báo cáo kết xuất.");
    } catch (err: any) {
      alert("Lỗi quá trình tạo báo cáo nhanh: " + err.message);
      setLoading(false);
    }
  };

  // 6. CHUẨN HÓA & PHÂN TÍCH NGÀNH (KHẮP MÃ THÔNG MINH + GOI Y AI VỚI TIẾN TRÌNH THỰC TẾ)
  const handleStandardizeSectors = async (useAI: boolean) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện.");
      return;
    }
    if (!mapping.mota || !mapping.manganh) {
      alert("Vui lòng cấu hình cột 'Mô tả hoạt động' và 'Mã ngành DTV' ở trang nạp file!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu phân tích chuẩn hóa ngành...");
    await sleep(200);

    const standardizedResults: any[] = [];
    const batchSize = Math.max(1, Math.floor(mainData.length / 10));

    for (let index = 0; index < mainData.length; index++) {
      const row = mainData[index];
      const motaVal = String(row[mapping.mota] || "").trim();
      const maDtvVal = normalizeSectorCode(row[mapping.manganh]);

      // Phân cấp mã ngành ĐTV đăng ký
      const hier = getSectorHierarchy(maDtvVal);
      const cap1Info = hier["1"];
      const cap2Info = hier["2"];
      const cap3Info = hier["3"];
      const cap4Info = hier["4"];
      const cap5Info = hier["5"];

      // Tra cứu nhanh AI/Smart Matcher gợi ý
      let goiyMa = "";
      let goiyTen = "";
      let diemTuongDong = "0.00";
      let giaiThich = "";
      let linhvucSuggest = "";

      if (useAI) {
        // Thực hiện cuộc gọi Gemini API Server proxy
        setStatusMessage(`[Phân tích AI] Đang dịch nghĩa dòng ${index + 1}/${mainData.length}: "${motaVal.slice(0, 30)}..."`);
        try {
          const res = await fetch("/api/gemini/analyze", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ description: motaVal })
          });
          const data = await res.json();
          if (data && !data.error) {
            goiyMa = normalizeSectorCode(data.goiy_ma);
            goiyTen = data.goiy_ten || "";
            linhvucSuggest = data.cap_1_tin_cay || "";
            giaiThich = data.giai_thich || "";
            diemTuongDong = "0.95"; // Ước lượng độ tinh cậy AI
          } else {
            // Trục trặc hoặc chưa cấu hình API key -> fall back qua Smart Matcher cục bộ
            const local = smartSuggestSectorByDescription(motaVal);
            if (local) {
              goiyMa = local.ma;
              goiyTen = local.ten;
              diemTuongDong = local.diem.toFixed(2);
              giaiThich = `Mô hình cục bộ đề xuất phân nhóm (AI Server fallback)`;
            }
          }
        } catch (e) {
          // Fall back
          const local = smartSuggestSectorByDescription(motaVal);
          if (local) {
            goiyMa = local.ma;
            goiyTen = local.ten;
            diemTuongDong = local.diem.toFixed(2);
            giaiThich = `Mô hình cục bộ đề xuất phân nhóm (Network error fallback)`;
          }
        }
      } else {
        // Chạy hoàn toàn bằng thuật toán so khớp từ khóa tiếng Việt thông minh siêu tốc (Smart Matcher)
        const local = smartSuggestSectorByDescription(motaVal);
        if (local) {
          goiyMa = local.ma;
          goiyTen = local.ten;
          diemTuongDong = local.diem.toFixed(2);
          giaiThich = `Khớp từ khóa thông minh thành công dạt hiệu số tích hợp`;
          
          const sugHier = getSectorHierarchy(goiyMa);
          linhvucSuggest = sugHier["1"]?.ma || "";
        }
      }

      // Đọc trong danh mục bộ nhớ chuẩn mã ngành cấp 5 của DTV nhập (Có hỗ trợ quy nạp cấp học nếu nhập mã con)
      const lookupResultDtv = lookupSectorNameWithFallback(maDtvVal);
      const stdCap5Ten = lookupResultDtv.name;

      // Đánh giá Trạng thái logic khớp mã
      let trangThai = "✅ Hợp lệ";

      // Lấy cấp 1 thực tế của DN đăng ký
      const dtvLinhVuc = cap1Info?.ma || "";

      // Phân tích đối chiếu tương đồng và phát hiện sai lệch thực tế theo mẫu ĐTV ghi
      const lcMota = motaVal.toLowerCase();
      
      const hasTradeKeywords = ["bán", "mua", "thương mại", "đại lý", "cửa hàng", "phân phối", "wholesale", "retail", "shop"].some(kw => lcMota.includes(kw));
      const hasIndustrialKeywords = ["sản xuất", "gia công", "chế tạo", "làm mộc", "chế biến", "lắp ráp", "chế tác", "luyện kim", "nhà xưởng", "nhà máy"].some(kw => lcMota.includes(kw));

      // Nhận diện mã ngành Công nghiệp: Cấp hai từ 10 tới 33 hoặc là thuộc nhóm C (Công nghiệp chế biến chế tạo)
      const activeSub2 = maDtvVal ? maDtvVal.slice(0, 2) : "";
      const sub2Num = parseInt(activeSub2, 10);
      const isIndustrialCode = (!isNaN(sub2Num) && sub2Num >= 10 && sub2Num <= 33) || dtvLinhVuc === "C";

      if (lookupResultDtv.level === 0) {
        trangThai = "❌ Lỗi: Mã ngành ĐTV không tồn tại trên danh mục VSIC chuẩn";
      } else if (hasTradeKeywords && !hasIndustrialKeywords && isIndustrialCode) {
        trangThai = "❌ Lỗi: Ghi ngành thương mại/đại lý bán lẻ nhưng gán mã ngành Công nghiệp (Bắt đầu bằng 10-33, Nhóm C)";
      } else if (hasIndustrialKeywords && !isIndustrialCode) {
        trangThai = "❌ Lỗi: Ghi ngành Sản xuất/Gia công/Làm mộc nhưng lại gán mã ngành dịch vụ/thương mại (không phải công nghiệp)";
      } else if (linhvucSuggest && dtvLinhVuc && linhvucSuggest !== dtvLinhVuc) {
        trangThai = `❌ Lỗi (LỆCH LĨNH VỰC): Mô tả hoạt động kinh doanh thiên về Nhóm [${linhvucSuggest}] nhưng Mã đăng ký thuộc Nhóm [${dtvLinhVuc}]`;
      } else if (goiyMa && parseFloat(diemTuongDong) > 0.6) {
        // Kiểm tra xem phân nhóm cấp 2 đăng ký có lệch với gợi ý không
        const regCap2 = cap2Info?.ma || "";
        const sugHier = getSectorHierarchy(goiyMa);
        const sugCap2 = sugHier["2"]?.ma || "";

        if (regCap2 && sugCap2 && regCap2 !== sugCap2) {
          trangThai = `⚠️ Cảnh báo (LỆCH CHI TIẾT CẤP 2): Hệ thống gợi ý mã [${goiyMa}] (${sugHier["2"]?.ten}), đăng ký thực nhập mã [${maDtvVal}]`;
        }
      }

      standardizedResults.push({
        ...row,
        "Hiệu_Chỉnh_ĐTV_Ghi": motaVal,
        "Tên_Ngành_Cấp_5_Chuẩn_VSIC": stdCap5Ten, // Đặt cạnh tên ngành ĐTV ghi
        "TrangThai_KiemTra_VSIC": trangThai,
        "Goiy_MaNganh_GoiY": goiyMa,
        "Goiy_TenNganh_GoiY": goiyTen,
        "Do_Tin_Cay_Matcher": diemTuongDong,
        "Giai_Thich_Phan_Tich": giaiThich,
        "Nganh_Cap_1": cap1Info?.ma || "",
        "Ten_Nganh_Cap_1": cap1Info?.ten || "",
        "Nganh_Cap_2": cap2Info?.ma || "",
        "Ten_Nganh_Cap_2": cap2Info?.ten || "",
        "Nganh_Cap_3": cap3Info?.ma || "",
        "Ten_Nganh_Cap_3": cap3Info?.ten || "",
        "Nganh_Cap_4": cap4Info?.ma || "",
        "Ten_Nganh_Cap_4": cap4Info?.ten || "",
        "Nganh_Cap_5": maDtvVal,
        "Ten_Nganh_Cap_5": cap5Info?.ten || stdCap5Ten
      });

      // Chỉ nghỉ ngắn để UI giữ responsive và mượt mà
      if (index % batchSize === 0 || index === mainData.length - 1) {
        const pct = Math.floor((index / mainData.length) * 100);
        setProgress(pct);
        setStatusMessage(`Đang chạy chuẩn hóa nâng cao: Dòng ${index}/${mainData.length}...`);
        await sleep(15);
      }
    }

    setMainData(standardizedResults);
    setColumns(Object.keys(standardizedResults[0] || {}));
    setFileName(`ChuanHoaNganh_VSIC_${fileName}`);

    setProgress(100);
    setStatusMessage(`Phân tích & Chuẩn hóa hoàn tất! Đã rà soát và phân tách 5 cấp cho ${standardizedResults.length} dòng dữ liệu.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // 7. CỖ MÁY KIỂM TRA LOGIC ĐA ĐIỀU KIỆN (NẾU ... THÌ PHẢI...)
  const handleLogicRuleAdd = (type: "if" | "then") => {
    if (type === "if") {
      if (!newIfRule.col) {
        alert("Vui lòng chọn cột điều kiện NẾU!");
        return;
      }
      setIfRules([...ifRules, newIfRule]);
      setNewIfRule({ col: "", op: "==", val: "" });
    } else {
      if (!newThenRule.col) {
        alert("Vui lòng chọn cột điều kiện THÌ PHẢI!");
        return;
      }
      setThenRules([...thenRules, newThenRule]);
      setNewThenRule({ col: "", op: "==", val: "" });
    }
  };

  const handleLogicCheck = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi kiểm tra logic.");
      return;
    }
    if (ifRules.length === 0 || thenRules.length === 0) {
      alert("Hãy định cấu hình ít nhất 1 quy tắc 'NẾU' và 1 quy tắc 'THÌ PHẢI'!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu kiểm tra logic đa điều kiện...");
    await sleep(200);

    // Chạy phép tính toán logic
    const checkValue = (rowVal: any, op: string, compareVal: string) => {
      const v1 = String(rowVal !== undefined && rowVal !== null ? rowVal : "").trim().toLowerCase();
      const v2 = String(compareVal).trim().toLowerCase();

      // Kiểm tra dạng rỗng
      if (op === "trống") return v1 === "";
      if (op === "không trống") return v1 !== "";

      // Kiểm tra dạng số
      const num1 = parseFloat(v1.replace(/[^0-9.\-]/g, ""));
      const num2 = parseFloat(v2.replace(/[^0-9.\-]/g, ""));

      if (!isNaN(num1) && !isNaN(num2)) {
        if (op === "==") return num1 === num2;
        if (op === "!=") return num1 !== num2;
        if (op === ">") return num1 > num2;
        if (op === "<") return num1 < num2;
        if (op === ">=") return num1 >= num2;
        if (op === "<=") return num1 <= num2;
      }

      // Kiểm tra dạng chuỗi
      if (op === "==") return v1 === v2;
      if (op === "!=") return v1 !== v2;
      if (op === "chứa") return v1.includes(v2);
      if (op === "không chứa") return !v1.includes(v2);

      return false;
    };

    const results = mainData.map((row, index) => {
      if (!row || typeof row !== 'object') return row;
      // 1. Phép toán NẾU
      const ifMatches = ifRules.map(r => checkValue(row[r.col], r.op, r.val));
      const satisfiesIf = ifCombine === "AND" 
        ? ifMatches.every(v => v === true) 
        : ifMatches.some(v => v === true);

      // 2. Phép toán THÌ PHẢI
      const thenMatches = thenRules.map(r => checkValue(row[r.col], r.op, r.val));
      const satisfiesThen = thenCombine === "AND"
        ? thenMatches.every(v => v === true)
        : thenMatches.some(v => v === true);

      let biViPham = false;
      let noteLoi = "";

      // Nếu dòng thỏa mãn điều kiện NẾU, thì bắt buộc PHẢI thỏa mãn điều kiện THÌ
      if (satisfiesIf && !satisfiesThen) {
        biViPham = true;
        const descriptIf = ifRules.map(r => `(${r.col} ${r.op} '${r.val}')`).join(` ${ifCombine} `);
        const descriptThen = thenRules.map(r => `(${r.col} ${r.op} '${r.val}')`).join(` ${thenCombine} `);
        noteLoi = `[VI PHẠM LOGIC] NẾU thỏa mãn: { ${descriptIf} } THÌ PHẢI: { ${descriptThen} }; `;
      }

      const existingLoi = String(row["Loi_Logic"] || "");

      return {
        ...row,
        "Loi_Logic": biViPham 
          ? (existingLoi ? existingLoi + noteLoi : noteLoi) 
          : (existingLoi || "✅ Đạt")
      };
    });

    setMainData(results);
    setColumns(Object.keys(results[0] || {}));
    setFileName(`KiemTraLogic_${fileName}`);

    setProgress(100);
    setStatusMessage(`Kiểm tra hoàn tất! Đã phân tích kiểm tra và phát hiện các dòng lỗi.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // 8. XUẤT FILE EXCEL CUỐI CÙNG
  const handleExportExcel = () => {
    const exportData = searchTerm ? filteredData : mainData;
    if (exportData.length === 0) {
      alert("Không có dữ liệu để xuất file!");
      return;
    }
    setLoading(true);
    setStatusMessage("Đang tạo tệp Excel phục vụ tải xuống (Gồm tệp tính toán & Danh mục ngành VSIC chuẩn)...");

    setTimeout(() => {
      try {
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Tinh_Toan");

        // Thêm sheet Danh mục ngành VSIC chuẩn quốc gia làm tài liệu tham khảo trong bản lưu
        const vsicRows = Object.entries(vsicRawData).map(([code, name]) => {
          let cap = 5;
          if (/^[A-Z]$/.test(code)) cap = 1;
          else if (code.length === 2) cap = 2;
          else if (code.length === 3) cap = 3;
          else if (code.length === 4) cap = 4;
          return {
            "Mã VSIC": code,
            "Tên Phân Cấp Ngành": name,
            "Phân Cấp": `Cấp ${cap}`
          };
        });
        const wsVsic = XLSX.utils.json_to_sheet(vsicRows);
        XLSX.utils.book_append_sheet(wb, wsVsic, "Danh_Muc_Nganh_VSIC_Chuan");

        // Đổi tên file đầu ra có hậu tố lọc nếu người dùng đang tìm kiếm/lọc
        let outName = fileName || "Ket_Qua_Bao_Cao.xlsx";
        if (searchTerm) {
          const safeSuffix = `_Loc_${searchTerm.trim().slice(0, 15).replace(/[^a-zA-Z0-9À-ỹ]/g, "_")}`;
          const lastDot = outName.lastIndexOf(".");
          if (lastDot !== -1) {
            outName = outName.slice(0, lastDot) + safeSuffix + outName.slice(lastDot);
          } else {
            outName = outName + safeSuffix + ".xlsx";
          }
        }

        XLSX.writeFile(wb, outName);
        setStatusMessage(`Đã tải xuống file Excel thành công! (Dữ liệu gồm ${exportData.length} dòng và Danh mục ngành)`);
      } catch (e: any) {
        alert("Lỗi khi kết xuất Excel: " + e.message);
      } finally {
        setLoading(false);
      }
    }, 200);
  };

  if (!isAuthorized) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-[#111827] text-gray-100 font-sans px-4 selection:bg-purple-600 selection:text-white">
        {/* Khóa bảo mật phi hành trạm VSIC */}
        <div className="w-full max-w-md bg-[#1f2937]/90 border border-purple-500/20 rounded-2xl p-8 shadow-2xl space-y-6 backdrop-blur-md relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-1.5 bg-gradient-to-r from-purple-600 to-indigo-500"></div>
          
          <div className="text-center space-y-2">
            <div className="mx-auto w-14 h-14 bg-gradient-to-tr from-purple-600 to-indigo-500 rounded-2xl flex items-center justify-center shadow-lg shadow-purple-900/30">
              <Lock className="w-7 h-7 text-white animate-pulse" />
            </div>
            <h2 className="text-xl font-bold tracking-tight text-white pt-2">CỔNG BẢO MẬT TRUY CẬP</h2>
            <p className="text-xs text-gray-400">Vui lòng nhập mật khẩu nội bộ để sử dụng hệ thống VSIC</p>
          </div>

          <form onSubmit={handleCheckPassword} className="space-y-4">
            <div className="space-y-1.5">
              <label className="text-xs text-gray-400 font-semibold font-mono">MẬT KHẨU TRUY CẬP:</label>
              <input
                type="password"
                value={typedPassword}
                onChange={(e) => setTypedPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
                className="w-full bg-[#111827] border border-[#374151] rounded-xl px-4 py-3 text-sm text-white focus:outline-none focus:border-purple-500 focus:ring-1 focus:ring-purple-500 transition-all placeholder:text-gray-600 font-mono"
                autoFocus
              />
              {passwordError && (
                <p className="text-red-400 text-[11px] font-medium flex items-center gap-1 mt-1 font-mono">
                  ⚠️ {passwordError}
                </p>
              )}
            </div>

            <button
              type="submit"
              className="w-full bg-gradient-to-r from-purple-600 to-indigo-500 hover:from-purple-700 hover:to-indigo-650 text-white font-bold text-sm py-3 px-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer"
            >
              🔐 Xác Nhận Trạm Làm Việc
            </button>
          </form>

          <div className="border-t border-gray-800/60 pt-4 text-center space-y-1">
            <p className="text-[11px] text-amber-400/85 italic">
              💡 Gợi ý: Mật khẩu mặc định là <strong className="font-mono bg-amber-950 px-1.5 py-0.5 rounded border border-amber-900/40 text-amber-300">admin123</strong>
            </p>
            <p className="text-[10px] text-gray-500 font-mono">Hệ thống bảo lưu mã khóa cục bộ an toàn trong trình duyệt của bạn</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-[#111827] text-gray-100 font-sans selection:bg-purple-600 selection:text-white overflow-hidden">
      
      {/* Header chính mang phong cách Cosmic Space Station sang trọng */}
      <header className="border-b border-[#374151] bg-[#1f2937]/90 backdrop-blur-md sticky top-0 z-40 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-tr from-purple-600 to-indigo-500 p-2.5 rounded-xl shadow-lg shadow-purple-900/30 ring-1 ring-purple-400/50">
            <Layers className="w-6 h-6 text-white animate-pulse" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-white flex items-center gap-2">
              HỆ THỐNG <span className="bg-purple-600 text-[10px] uppercase font-bold px-2 py-0.5 rounded-full tracking-widest">VSIC V38.5</span>
            </h1>
            <p className="text-xs text-gray-400 font-mono">CÔNG CỤ HỖ TRỢ SO SÁNH TỔNG HỢP DỮ LIỆU</p>
          </div>
        </div>

        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 mr-1">
            <button
              onClick={() => {
                setNewPasswordVal("");
                setShowPasswordChangeModal(true);
              }}
              className="px-3 py-1.5 bg-[#111827] hover:bg-gray-800 text-gray-300 rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition-colors border border-[#374151]"
              title="Thiết lập/Đổi mật khẩu bảo vệ riêng tư"
            >
              <KeyRound className="w-3.5 h-3.5 text-purple-400" />
              Đổi MK
            </button>
            <button
              onClick={handleLogout}
              className="px-3 py-1.5 bg-red-950/40 hover:bg-red-900/40 text-red-300 border border-red-900/50 rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition-colors"
              title="Khóa trạm làm việc ngay"
            >
              <LogOut className="w-3.5 h-3.5" />
              Khóa
            </button>
          </div>

          {fileName ? (
            <div className="bg-[#111827] border border-[#374151] rounded-lg px-4 py-1.5 flex items-center gap-2 text-xs">
              <Database className="w-4 h-4 text-emerald-400" />
              <span className="text-gray-300 font-medium">Hiện tại: </span>
              <span className="text-emerald-400 font-mono max-w-[200px] truncate" title={fileName}>{fileName}</span>
              <span className="bg-gray-800 text-gray-400 px-1.5 py-0.5 rounded font-mono font-semibold-">{mainData.length} dòng</span>
              <button 
                onClick={clearData}
                className="text-red-400 hover:text-red-300 ml-2 font-bold cursor-pointer transition-colors"
                title="Xóa dữ liệu nạp lại"
              >
                Xóa
              </button>
            </div>
          ) : (
            <span className="text-xs text-amber-400/90 bg-amber-950/40 border border-amber-900/50 rounded-lg px-4 py-1.5 flex items-center gap-1.5">
              <AlertTriangle className="w-3.5 h-3.5" /> Chưa có dữ liệu nguồn
            </span>
          )}
        </div>
      </header>

      {/* Main Layout split: Sidebar + Content */}
      <div className="flex flex-1 overflow-hidden">
        
        {/* Navigation Sidebar */}
        <aside className="w-72 bg-[#1f2937]/60 border-r border-[#374151] p-5 space-y-2 flex flex-col justify-between">
          <div className="space-y-1.5">
            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 mb-2">Thao tác dữ liệu</div>
            
            <button 
              onClick={() => setActiveTab("trangchu")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "trangchu" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-sm" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Home className="w-4 h-4" /> 🏠 Trang Chủ Tổng Quan
            </button>

            <button 
              onClick={() => setActiveTab("xemdulieu")}
              className={`w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "xemdulieu" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-sm" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <span className="flex items-center gap-3">
                <FileSpreadsheet className="w-4 h-4" /> 📂 Xem & Định Nghĩa Cột
              </span>
              <span className="text-[10px] font-mono bg-[#111827] text-gray-400 px-1.5 py-0.5 rounded-md">{mainData.length}</span>
            </button>

            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Công cụ liên hợp</div>

            <button 
              onClick={() => setActiveTab("ghepnoi")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "ghepnoi" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <GitMerge className="w-4 h-4 text-blue-400" /> 🌿 Ghép Nối Dữ Liệu
            </button>

            <button 
              onClick={() => setActiveTab("sosanh")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "sosanh" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Combine className="w-4 h-4 text-cyan-400" /> 🔍 So Sánh Đối Chiếu
            </button>

            <button 
              onClick={() => setActiveTab("tachfile")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "tachfile" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Scissors className="w-4 h-4 text-pink-400" /> ✂️ Tách File Hàng Loạt
            </button>

            <button 
              onClick={() => setActiveTab("tonghop")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "tonghop" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <BarChart3 className="w-4 h-4 text-amber-400" /> 📊 Tổng Hợp Báo Cáo
            </button>

            <button 
              onClick={() => setActiveTab("bieudotrucquan")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "bieudotrucquan" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-purple-500/10 shadow-sm animate-pulse" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <PieChart className="w-4 h-4 text-cyan-400" /> 📈 Biểu Đồ Trực Quan
            </button>

            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Thông minh & Rà soát</div>

            <button 
              onClick={() => setActiveTab("chuanhoanganh")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "chuanhoanganh" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 animate-pulse" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Brain className="w-4 h-4 text-indigo-400" /> 🧠 Chuẩn Hóa VSIC & AI
            </button>

            <button 
              onClick={() => setActiveTab("kiemtralogic")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "kiemtralogic" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <CheckSquare className="w-4 h-4 text-emerald-400" /> 🛂 Cỗ Máy Kiểm Tra Logic
            </button>

            <button 
              onClick={() => setActiveTab("doichieumota")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "doichieumota" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-purple-500/10 shadow-sm" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <ArrowRightLeft className="w-4 h-4 text-purple-400 animate-pulse" /> 🔄 Đối Chiếu Mô Tả Ngành
            </button>

            <button 
              onClick={() => setActiveTab("danhmucvsic")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "danhmucvsic" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Database className="w-4 h-4 text-amber-400" /> 🗂️ Danh Mục Ngành VSIC
            </button>
          </div>

          {/* Footer Sidebar */}
          <div className="bg-[#111827]/80 rounded-xl p-3.5 border border-purple-950/40 text-[10px] text-gray-400 font-mono leading-relaxed space-y-1.5 shadow-inner">
            <div className="flex items-center gap-1.5 text-emerald-400 font-semibold">
              <span className="relative flex h-2 w-2">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
              </span>
              💾 BỘ NHỚ LOCAL WORKSPACE
            </div>
            <div className="text-gray-500 text-[9px] leading-normal font-sans">
              Dữ liệu của bạn được lưu an toàn trực tiếp trong cơ sở dữ liệu trình duyệt (IndexedDB). Bạn có thể tắt máy, đóng tab thoải mái và khi mở lại chương trình, dữ liệu sẽ tự động khôi phục 100%!
            </div>
          </div>
        </aside>

        {/* Content Area */}
        <main className="flex-1 bg-[#111827] overflow-y-auto p-6 md:p-8">
          
          {/* Lớp hiển thị nạp dữ liệu/ tiến trình hệ thống khi chạy */}
          {loading && (
            <div className="fixed inset-0 z-50 bg-[#111827]/80 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl max-w-md w-full p-6 text-center space-y-4 shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 h-1 bg-gradient-to-r from-purple-600 via-indigo-500 to-cyan-400 transition-all duration-300" style={{ width: `${progress}%` }}></div>
                <Loader2 className="w-12 h-12 text-purple-500 mx-auto animate-spin" />
                <h3 className="text-lg font-bold text-white">Đang xử lý dữ liệu</h3>
                <p className="text-sm text-gray-400 font-mono leading-relaxed min-h-[40px]">{statusMessage}</p>
                
                <div className="w-full bg-[#111827] rounded-full h-2.5 overflow-hidden border border-gray-800">
                  <div className="bg-purple-600 h-2.5 rounded-full transition-all duration-300" style={{ width: `${progress}%` }}></div>
                </div>
                <div className="text-xs font-bold text-purple-400 tracking-wider font-mono">{progress}% Hoàn Thành</div>
              </div>
            </div>
          )}

          {/* Lớp hiển thị đổi mật khẩu truy cập */}
          {showPasswordChangeModal && (
            <div className="fixed inset-0 z-50 bg-[#111827]/80 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl max-w-sm w-full p-6 space-y-4 shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-purple-600 to-indigo-500"></div>
                <div className="text-center space-y-1">
                  <div className="mx-auto w-10.5 h-10.5 bg-purple-950/40 border border-purple-500/20 rounded-xl flex items-center justify-center">
                    <KeyRound className="w-5.5 h-5.5 text-purple-400" />
                  </div>
                  <h3 className="text-base font-bold text-white pt-1">ĐỔI MẬT KHẨU BẢO VỆ</h3>
                  <p className="text-xs text-gray-400 text-center font-sans">Thiết lập mật khẩu riêng tư cho trình quản lý</p>
                </div>

                <div className="space-y-4">
                  <div className="space-y-1.5 font-sans">
                    <label className="text-[10.5px] font-bold text-gray-300 font-mono block">MẬT KHẨU MỚI TIN CẬY:</label>
                    <input
                      type="text"
                      className="w-full bg-[#111827] border border-[#374151] rounded-lg px-3 py-2 text-xs text-white focus:outline-none focus:border-purple-500 font-mono"
                      placeholder="Nhập mật khẩu mới..."
                      value={newPasswordVal}
                      onChange={(e) => setNewPasswordVal(e.target.value)}
                      autoFocus
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-2 pt-2">
                    <button
                      onClick={() => setShowPasswordChangeModal(false)}
                      className="w-full bg-[#1e293b] hover:bg-gray-800 text-gray-400 font-bold text-xs py-2 px-3 rounded-lg border border-gray-800 transition-all cursor-pointer font-sans"
                    >
                      Hủy Bỏ
                    </button>
                    <button
                      onClick={handleChangePassword}
                      className="w-full bg-gradient-to-r from-purple-600 to-indigo-500 hover:from-purple-700 hover:to-indigo-600 text-white font-bold text-xs py-2 px-3 rounded-lg shadow transition-all cursor-pointer font-sans"
                    >
                      Xác Nhận Đổi
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 1. TAB TRANG CHỦ */}
          {activeTab === "trangchu" && (
            <div className="space-y-8 animate-fade-in">
              <div className="bg-gradient-to-r from-purple-900/40 via-[#1f2937] to-[#1f2937] border border-purple-500/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-center gap-6 justify-between">
                <div className="space-y-3 max-w-2xl">
                  <span className="bg-purple-900/50 border border-purple-500/30 text-purple-400 text-xs font-mono font-bold px-3 py-1 rounded-full uppercase tracking-wider">
                    Hệ Thống DM ngành kinh tế VISC2025 
                  </span>
                  <h2 className="text-2xl md:text-3xl font-extrabold text-white tracking-tight">
                    HỆ THỐNG XỬ LÝ DỮ LIỆU TỔNG HỢP SO SÁNH 
                  </h2>
                  <p className="text-gray-300 text-sm leading-relaxed">
                    Công cụ chuyên sâu hỗ trợ các cuộc điều tra có thể đọc và xử lý tất cá các loại bảng biểu không phân biệt ví trí cột hay tên cột có chức năng ghép, tách tệp lớn, so khớp, rà soát logic đa chỉ tiêu, tổng hợp và xử lý liên kết ngành kinh tế VISC2025.
                  </p>
                  <div className="pt-2 flex items-center gap-4">
                    <button 
                      onClick={() => setActiveTab("xemdulieu")}
                      className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-sm px-6 py-2.5 rounded-xl transition-all shadow-md shadow-purple-900/30 flex items-center gap-2 cursor-pointer"
                    >
                      📂 Nạp file dữ liệu của bạn để bắt đầu <ArrowRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
                <div className="w-full md:w-auto flex justify-center">
                  <div className="bg-gradient-to-tr from-[#374151] to-purple-800/20 border border-[#4b5563] p-6 rounded-2xl text-center space-y-2 min-w-[200px] shadow-sm">
                    <div className="text-4xl font-extrabold text-white font-mono">197</div>
                    <div className="text-[11px] font-bold text-gray-400 tracking-wider uppercase font-mono">Mã ngành VSIC nhúng</div>
                    <div className="text-[10px] text-green-400 font-mono">Đầy đủ 5 cấp phân chiêu</div>
                  </div>
                </div>
              </div>

              {/* PHẦN DANH SÁCH CHỨC NĂNG CHÍNH */}
              <div className="space-y-6">
                <div className="border-b border-[#374151] pb-4">
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <Layers className="w-5 h-5 text-purple-400 animate-pulse" /> NỘI DUNG & KÍCH HOẠT CHỨC NĂNG NHANH
                  </h3>
                  <p className="text-xs text-gray-400 mt-1">
                    Hệ thống tích hợp đầy đủ 7 phân hệ cốt lõi chuyên sâu. Bạn có thể xem nhanh hướng dẫn và nhấp trực tiếp vào bất kỳ thẻ nào dưới đây để bắt đầu ngay:
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {/* CHỨC NĂNG 1: Xem & Định nghĩa cột */}
                  <div className="bg-[#1f2937]/50 border border-purple-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-purple-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-purple-950/50 border border-purple-500/30 p-2.5 rounded-xl text-purple-400">
                          <FileSpreadsheet className="w-5 h-5" />
                        </div>
                        <span className="bg-purple-900/40 text-purple-300 border border-purple-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          NẠP ĐẦU VÀO
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-purple-300 transition-colors">
                        📂 Xem &amp; Định Nghĩa Cột
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Nơi tải lên file Excel/CSV gốc. Hỗ trợ <strong>Việt hóa / đặt tên lại</strong> cho các cột viết tắt khó nhớ, lọc bỏ cột thừa và gán vai trò kinh doanh (Mô tả ngành, Mã ngành, Xã, Doanh thu, Lao động,...) để toàn bộ hệ thống nhận diện tự động.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("xemdulieu")}
                      className="w-full bg-[#111827] hover:bg-purple-900/30 text-purple-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-purple-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Xem & Cấu Hình Cột <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 2: Ghép nối dữ liệu */}
                  <div className="bg-[#1f2937]/50 border border-blue-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-blue-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-blue-950/50 border border-blue-500/30 p-2.5 rounded-xl text-blue-400">
                          <GitMerge className="w-5 h-5" />
                        </div>
                        <span className="bg-blue-900/40 text-blue-300 border border-blue-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          TÍCH HỢP FILE
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-blue-300 transition-colors">
                        🌿 Ghép Nối Dữ Liệu
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tự động tích hợp thông tin từ tệp phụ lục vào tệp tin chính dựa trên cột liên quan chung (như Mã số thuế, Mã doanh nghiệp,...). Giải quyết nỗi lo ghép thủ công dễ gây xô lệch hoặc gõ nhầm dòng dữ liệu.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("ghepnoi")}
                      className="w-full bg-[#111827] hover:bg-blue-900/30 text-blue-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-blue-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Ghép Nối Dữ Liệu <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 3: So sánh chéo */}
                  <div className="bg-[#1f2937]/50 border border-cyan-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-cyan-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-cyan-950/50 border border-cyan-500/30 p-2.5 rounded-xl text-cyan-400">
                          <Combine className="w-5 h-5" />
                        </div>
                        <span className="bg-cyan-900/40 text-cyan-300 border border-cyan-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          THẨM ĐỊNH CHÉO
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-cyan-300 transition-colors">
                        🔍 So Sánh Đối Chiếu
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        So sánh đối chuẩn chi tiết giữa 2 thời điểm hoặc 2 nguồn danh sách khác biệt. Tìm kiếm siêu tốc doanh nghiệp bị thiếu, cơ sở kinh doanh mới phát sinh hoặc có dao động dữ liệu bất thường về thu nhập và lao động.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("sosanh")}
                      className="w-full bg-[#111827] hover:bg-cyan-900/30 text-cyan-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-cyan-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở So Sánh Đối Chiếu <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 4: Tách file hàng loạt */}
                  <div className="bg-[#1f2937]/50 border border-pink-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-pink-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-pink-950/50 border border-pink-500/30 p-2.5 rounded-xl text-pink-400">
                          <Scissors className="w-5 h-5" />
                        </div>
                        <span className="bg-pink-900/40 text-pink-300 border border-pink-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          BẺ TÁCH HOÀN LOẠT
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-pink-300 transition-colors">
                        ✂️ Tách File Hàng Loạt
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Bẻ chia nhanh một danh sách tổng thể khổng lồ thành vô vàn file Excel lẻ tương ứng với từng <strong>Xã/Phường</strong> hoặc <strong>Nhóm ngành cụ thể</strong>. Xuất nén gọn thành thư mục .ZIP để gửi đi nhanh chóng.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("tachfile")}
                      className="w-full bg-[#111827] hover:bg-pink-900/30 text-pink-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-pink-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Tách File Hàng Loạt <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 5: Tổng hợp báo cáo */}
                  <div className="bg-[#1f2937]/50 border border-amber-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-amber-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-amber-950/50 border border-amber-500/30 p-2.5 rounded-xl text-amber-400">
                          <BarChart3 className="w-5 h-5" />
                        </div>
                        <span className="bg-amber-900/40 text-amber-300 border border-amber-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          ĐA CHIỀU PIVOT
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-amber-300 transition-colors">
                        📊 Tổng Hợp Báo Cáo
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tự động nhóm tổng số chỉ tiêu Doanh thu, Lao động lũy kế từ mức Ngành Cấp 5 chi tiết lên đến Ngành tổng quát cấp 1. Hỗ trợ biểu kết xuất dạng phẳng thông thường hoặc dạng bảng xoay <strong>Pivot Table</strong> chiều ngang.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("tonghop")}
                      className="w-full bg-[#111827] hover:bg-amber-900/30 text-amber-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-amber-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Tổng Hợp Báo Cáo <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 6: Chuẩn hóa khớp ngành VSIC & so sánh tra lỗi*/}
                  <div className="bg-[#1f2937]/50 border border-indigo-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-indigo-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-indigo-950/50 border border-indigo-500/30 p-2.5 rounded-xl text-indigo-400">
                          <Brain className="w-5 h-5" />
                        </div>
                        <span className="bg-indigo-900/40 text-indigo-300 border border-indigo-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          TRA CỨU MÃ NGÀNH VÀ MÔ TẢ NGÀNH
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-indigo-300 transition-colors">
                        🧠 Khớp ngành VSIC & so sánh mô tả với mã để phát hiện sai lệch
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tự động rã tách mã ngành 1-5 cấp đầy đủ không đứt gãy phả hệ liên kết. Sử dụng thuật toán NLP tinh lọc so sánh ngữ nghĩa văn bản của "Mô tả hoạt động thực tế" so với "Mã ngành" đăng ký để phát hiện và chỉnh lỗi mã sai lệch.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("chuanhoanganh")}
                      className="w-full bg-[#111827] hover:bg-indigo-900/30 text-indigo-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-indigo-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Tra cứu mô tả& mã  <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 7: Công cụ kiểm tra logic */}
                  <div className="bg-[#1f2937]/50 border border-emerald-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-emerald-500/40 transition-all col-span-1 md:col-span-2 lg:col-span-3 group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-emerald-950/50 border border-emerald-500/30 p-2.5 rounded-xl text-emerald-400">
                          <CheckSquare className="w-5 h-5" />
                        </div>
                        <span className="bg-emerald-900/40 text-emerald-300 border border-emerald-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          RÀ SOÁT QUY TẮC ĐỘNG
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-emerald-300 transition-colors">
                        Control Panel 🛂 Công cụ Kiểm Tra Logic
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Thiết lập các quy định ràng buộc điều kiện thông minh dạng <code>NẾU (điều kiện) THÌ PHẢI (điều kiện kia)</code> tùy ý. Bộ máy sẽ lập tức rà soát toàn bộ tệp Excel thu về danh sách hồ sơ bất tuần tự (ví dụ: mô tả có chứa xi-măng nhưng mã ngành lại đăng ký nông nghiệp, hoặc lao động siêu lớn nhưng doanh thu rỗng).
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("kiemtralogic")}
                      className="w-full bg-[#111827] hover:bg-emerald-900/30 text-emerald-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-emerald-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Công cụ Rà Quét Logic <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              </div>

              {/* Gợi ý quy trình xử lý dữ liệu chuẩn */}
              <div className="bg-[#1e1b4b]/20 border border-purple-500/20 rounded-2xl p-6 space-y-4">
                <h4 className="text-sm font-bold text-purple-300 flex items-center gap-2">
                  <Info className="w-4 h-4 text-purple-400 animate-pulse" /> ĐỀ XUẤT 3 BƯỚC VẬN HÀNH CHUẨN TRÊN HỆ THỐNG
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-xs text-gray-300 leading-relaxed">
                  <div className="space-y-1">
                    <span className="font-bold text-white block">01. Nạp & Tiền xử lý dữ liệu</span>
                    <p className="text-gray-400">
                      Truy cập bảng <strong>📂 Xem & Định nghĩa cột</strong>. Tải tệp Excel gốc lên, thực hiện đổi tên Việt hóa cho các cột và lọc bỏ những cột không sửi dụng đến.
                    </p>
                  </div>
                  <div className="space-y-1">
                    <span className="font-bold text-white block">02. Chuẩn hóa & Rà rà logic</span>
                    <p className="text-gray-400">
                      Sử dụng <strong>🧠 Khớp ngành VSIC & so sánh mô tả với mã </strong> để hoàn thiện liên kết 5 cấp ngành nghề; tiếp theo sử dụng <strong>🛂 Công cụ Logic</strong> thiết lập các quy chuẩn kiểm tra để lọc sạch các bản ghi lỗi hoặc dị thường.
                    </p>
                  </div>
                  <div className="space-y-1">
                    <span className="font-bold text-white block">03. Tổng hợp báo cáo & xuất file </span>
                    <p className="text-gray-400">
                      Qua trang <strong>📊 Tổng Hợp Báo Cáo</strong> để quy nập các chỉ thị hoặc chọn <strong>✂️ Tách File</strong> tạo tệp zip con của các xã gửi cho từng địa bàn. Bấm tải tệp Excel thành phẩm là xong!
                    </p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 2. TAB FILE VIEWER & COLUMN MAPPING */}
          {activeTab === "xemdulieu" && (
            <div className="space-y-6 animate-fade-in">
              
              {/* Box Upload chính */}
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <h3 className="text-lg font-bold text-white flex items-center gap-2">
                      <FileSpreadsheet className="w-5 h-5 text-purple-400" /> FILE DỮ LIỆU NGUỒN CHÍNH
                    </h3>
                    <p className="text-xs text-gray-400">Tải lên tệp dữ liệu chính (Excel/CSV) của bạn hoặc định nghĩa nhanh các cột chỉ định bên dưới.</p>
                  </div>

                  <label className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-purple-900/20 flex items-center gap-2 cursor-pointer self-start w-full md:w-auto justify-center">
                    <FileUp className="w-4 h-4" /> TẢI FILE DỮ LIỆU CHÍNH (EXCEL, CSV)
                    <input 
                      type="file" 
                      accept=".xlsx, .xls, .csv, .txt" 
                      onChange={(e) => handleFileUpload(e, "main")} 
                      className="hidden" 
                    />
                  </label>
                </div>

                {/* HIỂN THỊ CẤU HÌNH GHÉP CÁC SHEET KHI PHÁT HIỆN TỆP NHIỀU SHEET */}
                {detectedSheets.length > 1 && (
                  <div className="bg-[#111827]/90 rounded-2xl p-5 border border-amber-500/35 space-y-4 animate-slide-up mt-4">
                    <div className="flex items-center gap-2 border-b border-gray-800 pb-3">
                      <div className="p-1.5 bg-amber-950/50 rounded-lg border border-amber-500/25">
                        <FileSpreadsheet className="w-5 h-5 text-amber-400" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-white uppercase tracking-wider font-mono">
                          ⚡ Phát hiện File có nhiều Sheet ({detectedSheets.length} Sheets)
                        </h4>
                        <p className="text-[11px] text-amber-200/80">
                          Bạn có thể ghép (gộp) dữ liệu của nhiều Sheet này lại với nhau dựa trên một cột chung (ví dụ: Mã số thuế, Mã định danh, ID,...).
                        </p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#1f2937]/40 p-4 rounded-xl border border-gray-800">
                      <div>
                        <label className="text-[11px] font-bold text-gray-300 block mb-1.5 font-mono">
                          1. CHỌN CÁC SHEET MUỐN GHÉP:
                        </label>
                        <div className="max-h-[140px] overflow-y-auto space-y-1.5 p-2 bg-[#111827] rounded-lg border border-gray-800">
                          {detectedSheets.map(sheet => {
                            const isSelected = selectedSheetsToMerge.includes(sheet);
                            return (
                              <label key={sheet} className="flex items-center gap-2 text-xs text-gray-300 hover:text-white cursor-pointer select-none">
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={() => {
                                    if (isSelected) {
                                      setSelectedSheetsToMerge(selectedSheetsToMerge.filter(s => s !== sheet));
                                    } else {
                                      setSelectedSheetsToMerge([...selectedSheetsToMerge, sheet]);
                                    }
                                  }}
                                  className="rounded border-gray-700 bg-gray-950 text-amber-500 focus:ring-amber-500"
                                />
                                <span className={isSelected ? "text-amber-300 font-semibold" : ""}>{sheet}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>

                      <div className="flex flex-col justify-between space-y-3">
                        <div className="space-y-1.5">
                          <label className="text-[11px] font-bold text-gray-200 block font-mono">
                            2. CHỌN CỘT LIÊN KẾT CHUNG (KEY COLUMN):
                          </label>
                          <select
                            value={sheetMergeCommonKey}
                            onChange={(e) => setSheetMergeCommonKey(e.target.value)}
                            className="w-full bg-[#111827] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:border-amber-500 focus:ring-1 focus:ring-amber-500 font-mono"
                          >
                            <option value="">-- Chọn cột định danh dùng để gộp dòng --</option>
                            {columns.map(c => (
                              <option key={c} value={c}>
                                🔑 Cột: {c}
                              </option>
                            ))}
                          </select>
                          <p className="text-[10px] text-gray-500 italic">
                            Hệ thống sẽ đồng nhất, phối hợp các thông tin cột của dòng từ các Sheet dựa theo giá trị trùng khớp tại cột này.
                          </p>
                        </div>

                        <button
                          onClick={handleMergeWorkbookSheets}
                          disabled={selectedSheetsToMerge.length < 2 || !sheetMergeCommonKey}
                          className="w-full bg-gradient-to-r from-amber-500 to-yellow-600 hover:from-amber-600 hover:to-yellow-700 disabled:from-gray-700 disabled:to-gray-800 disabled:text-gray-500 disabled:cursor-not-allowed text-white font-bold text-xs py-2.5 px-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer"
                        >
                          ⚡ GHÉP CÁC SHEET THÀNH 1 BẢNG CHUNG
                        </button>
                      </div>
                    </div>
                  </div>
                )}

              {/* Phần cấu hình định nghĩa lại tên cột theo phong cách của người dùng (CUSTOM RE-DEFINITION GRID) */}
              {rawImportedData.length > 0 && (
                  <div className="bg-[#111827]/90 rounded-2xl p-5 border border-purple-500/20 space-y-5 animate-slide-up">
                    
                    <div className="border-b border-gray-800 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <div className="text-xs font-bold text-purple-400 tracking-wider uppercase font-mono flex items-center gap-1.5 cursor-pointer select-none" onClick={() => setIsConfigExpanded(!isConfigExpanded)}>
                          <Database className="w-5 h-5 text-purple-400 animate-pulse" /> ĐỊNH NGHĨA LẠI TÊN CỘT DỄ NHỚ & LỌC CỘT THỪA {isConfigExpanded ? "▼" : "▲"}
                        </div>
                        <p className="text-xs text-gray-400 mt-1">
                          Sửa đổi các từ viết tắt khó nhớ thành tiếng Việt rõ ràng. Cột nào chưa chọn sẽ bị loại khỏi bảng để giữ bộ dữ liệu sạch nhất.
                        </p>
                      </div>
                      
                      <div className="flex gap-2 items-center">
                        <button
                          onClick={() => setIsConfigExpanded(!isConfigExpanded)}
                          className="bg-purple-950/60 hover:bg-purple-900/60 text-purple-300 font-bold text-[11px] px-3.5 py-1.5 rounded-lg border border-purple-800/30 cursor-pointer transition-all flex items-center gap-1"
                        >
                          {isConfigExpanded ? "👁️ Thu gọn bảng" : "⚙️ Mở rộng định nghĩa cột"}
                        </button>
                        {isConfigExpanded && (
                          <>
                            <button 
                              onClick={() => {
                                // Reset everything back to original state
                                const resetConfigs = customColConfigs.map(c => ({
                                  ...c,
                                  use: true,
                                  newName: c.originalName
                                }));
                                setCustomColConfigs(resetConfigs);
                              }}
                              className="bg-gray-800 hover:bg-gray-700 text-gray-300 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-gray-700 cursor-pointer"
                              title="Hoàn tác tất cả tên cột về tên gốc"
                            >
                              Khôi Phục Tên Gốc
                            </button>
                            <button 
                              onClick={() => {
                                // Quick auto prefill nice names
                                const prefilled = customColConfigs.map(cfg => {
                                  const c = cfg.originalName;
                                  let newN = c;
                                  if (c === "MoTa") newN = "Mô Tả Hoạt Động";
                                  else if (c === "MaNganhDTV") newN = "Mã Ngành Đăng Ký";
                                  else if (c === "Xa") newN = "Địa bàn (Xã)";
                                  else if (c === "DoanhThu") newN = "Doanh Thu";
                                  else if (c === "LaoDong") newN = "Tổng số Lao Động";
                                  else if (c === "MaST") newN = "Mã Số Thuế";
                                  return { ...cfg, newName: newN };
                                });
                                setCustomColConfigs(prefilled);
                              }}
                              className="bg-purple-950/40 hover:bg-purple-900/40 text-purple-300 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-purple-800/30 cursor-pointer"
                            >
                              Tự Động Đề Xuất Tên Việt Hóa
                            </button>
                          </>
                        )}
                      </div>
                    </div>

                    {isConfigExpanded ? (
                      <>
                        {/* Hướng dẫn chi tiết */}
                        <div className="bg-[#1e1b4b]/30 border border-purple-900/30 rounded-xl p-4 text-xs text-gray-300 space-y-1.5 leading-relaxed">
                          <div className="font-bold text-purple-300 flex items-center gap-1.5">
                            ⚙️ Cách thức vận hành (Định nghĩa trực quan):
                          </div>
                          <ul className="list-disc list-inside space-y-1 text-[11px] text-gray-400 pl-1">
                            <li><strong>Đặt tên cột dễ nhớ:</strong> Viết trực tiếp vào ô nhập bên dưới để thay đổi tên cột hiển thị theo từ ngữ dễ thuộc của riêng bạn.</li>
                            <li><strong>Lọc cột thừa:</strong> Bạn có thể bỏ tích ở cột không cần thiết, khi bấm áp dụng hệ thống sẽ sinh ra một <strong>Bảng dữ liệu mới hoàn hảo</strong> chỉ chứa các cột thích hợp.</li>
                            <li><strong>Gán vai trò (Mục tiêu):</strong> Gán vai trò cho cột giúp các thuật toán (Báo cáo xã, nhóm ngành, xử lý lỗi logic bằng AI) tự động tìm đúng dữ liệu mà không bị đứt gãy.</li>
                          </ul>
                        </div>

                        {/* Bảng Danh sách Cấu hình Cột */}
                        <div className="overflow-x-auto border border-gray-800 rounded-xl bg-[#0f172a]/40">
                          <table className="w-full text-left text-xs min-w-[700px]">
                            <thead>
                              <tr className="bg-[#1f2937]/50 border-b border-gray-800 text-gray-400 font-mono text-[11px]">
                                <th className="p-3 text-center w-[70px]">SỬ DỤNG</th>
                                <th className="p-3 text-center w-[50px]">STT</th>
                                <th className="p-3">TÊN CỘT GỐC TRONG FILE (NHẤP ĐÚP ĐỂ CHỌN NHANH ⚡)</th>
                                <th className="p-3">TÊN MỚI ĐỊNH NGHĨA (ĐỂ TRỐNG = LOẠI BỎ KHỎI FILE)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-800/60 font-sans">
                              {customColConfigs.map((cfg, idx) => {
                                const isIncluded = cfg.newName.trim() !== "";
                                return (
                                  <tr 
                                    key={cfg.originalName} 
                                    className={`transition-colors hover:bg-gray-800/15 ${
                                      isIncluded ? "bg-purple-950/5" : "bg-gray-950/30 opacity-60"
                                    }`}
                                    onDoubleClick={() => {
                                      // Nháy đúp vào cột gốc để điền nhanh tên mới
                                      const updated = [...customColConfigs];
                                      updated[idx].newName = cfg.originalName;
                                      updated[idx].use = true;
                                      setCustomColConfigs(updated);
                                    }}
                                    title="Nhấp đúp vào dòng này để tự động điền nhanh Tên cột gốc thành Tên mới!"
                                  >
                                    {/* Cột Checkbox Sử Dụng */}
                                    <td className="p-3 text-center">
                                      <input 
                                        type="checkbox"
                                        checked={cfg.use && isIncluded}
                                        onChange={(e) => {
                                          const updated = [...customColConfigs];
                                          updated[idx].use = e.target.checked;
                                          if (e.target.checked && updated[idx].newName.trim() === "") {
                                            updated[idx].newName = cfg.originalName;
                                          } else if (!e.target.checked) {
                                            updated[idx].newName = "";
                                          }
                                          setCustomColConfigs(updated);
                                        }}
                                        className="w-4 h-4 rounded border-gray-700 bg-gray-900 text-purple-600 focus:ring-purple-500 accent-purple-600 cursor-pointer"
                                      />
                                    </td>

                                    {/* STT */}
                                    <td className="p-3 text-center text-gray-500 font-mono text-[11px]">
                                      {idx + 1}
                                    </td>

                                    {/* Tên Gốc */}
                                    <td 
                                      className="p-3 font-semibold text-gray-300 font-mono cursor-pointer hover:text-purple-400 transition-all"
                                      title="Nhấn đúp vào đây để chọn nhanh giữ tên cột gốc làm định nghĩa!"
                                    >
                                      <div className="flex items-center gap-2">
                                        <span className="bg-[#1f2937] px-2.5 py-1 rounded text-gray-200 border border-gray-700 max-w-[250px] truncate block">
                                          {cfg.originalName}
                                        </span>
                                        <span className="text-[10px] text-gray-500 hover:text-purple-400 select-none">
                                          {isIncluded ? "⚡ Đã gán" : "🖱️ Nháy đúp để lấy"}
                                        </span>
                                      </div>
                                    </td>

                                    {/* Input Tên Mới */}
                                    <td className="p-3">
                                      <input 
                                        type="text"
                                        value={cfg.newName}
                                        onChange={(e) => {
                                          const updated = [...customColConfigs];
                                          updated[idx].newName = e.target.value;
                                          updated[idx].use = e.target.value.trim() !== "";
                                          setCustomColConfigs(updated);
                                        }}
                                        className="w-full bg-gray-950 border border-purple-950/50 rounded-lg px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-medium placeholder-gray-700 font-mono"
                                        placeholder="Điền tên mới hoặc để trống dể loại bỏ..."
                                      />
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>

                        {/* Nút Kích hoạt Tái cấu trúc bảng */}
                        <div className="flex justify-between items-center pt-2">
                          <button 
                            onClick={() => {
                              // Clear all defined names to let user select only what they want
                              const cleared = customColConfigs.map(c => ({
                                ...c,
                                use: false,
                                newName: ""
                              }));
                              setCustomColConfigs(cleared);
                            }}
                            className="bg-red-950/30 hover:bg-red-900/40 text-red-400 border border-red-900/30 font-bold text-[11px] px-3.5 py-2 rounded-xl transition-all cursor-pointer"
                          >
                            ❌ Xóa hết định nghĩa (Để trống tất cả)
                          </button>

                          <button
                            onClick={handleApplyColumnRedefinition}
                            className="bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-md shadow-purple-950/40 flex items-center gap-2 cursor-pointer border border-purple-500/20 hover:scale-[1.02] active:scale-[0.98]"
                          >
                            <FileCheck className="w-4 h-4" />⚡ XÁC NHẬN ĐỊNH NGHĨA & LỌC GỌN NHẸ TỔ TẠO FILE MỚI
                          </button>
                        </div>
                      </>
                    ) : (
                      <div className="flex items-center justify-between text-xs text-gray-400 bg-purple-950/10 p-3 rounded-xl border border-purple-500/10">
                        <span>💡 Bảng cấu hình định nghĩa tên cột đang được thu gọn để nhường lại không gian biểu diễn danh sách dữ liệu.</span>
                        <button
                          onClick={() => setIsConfigExpanded(true)}
                          className="bg-purple-950/85 hover:bg-purple-900/85 text-purple-300 font-bold text-xs px-3.5 py-1.5 rounded-lg border border-purple-800/30 cursor-pointer transition-all"
                        >
                          ⚙️ Hiện bảng cấu hình
                        </button>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Danh sách dữ liệu chính */}
              {mainData.length > 0 ? (
                <div className="bg-[#1f2937] border border-[#374151] rounded-2xl overflow-hidden shadow-sm space-y-4 p-4">
                  
                  {/* Thanh công cụ lọc */}
                  <div className="flex flex-col sm:flex-row items-center justify-between gap-4 border-b border-[#374151] pb-4">
                    <div className="relative w-full sm:max-w-xs">
                      <Search className="absolute left-3 top-2.5 h-4 w-4 text-gray-400" />
                      <input 
                        type="text" 
                        placeholder="Tìm nhanh mọi vùng..." 
                        value={searchTerm}
                        onChange={(e) => { setSearchTerm(e.target.value); setViewPage(1); }}
                        className="w-full bg-[#111827] border border-[#374151] rounded-xl pl-9 pr-4 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                      />
                    </div>

                    <div className="flex items-center gap-3">
                      <div className="text-xs text-gray-400">
                        Hiển thị {paginatedData.length}/{filteredData.length} dòng
                      </div>
                      <button 
                        onClick={handleExportExcel}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer"
                      >
                        <Download className="w-4 h-4" /> Xuất File báo cáo Excel
                      </button>
                    </div>
                  </div>

                  {/* Bảng dữ liệu bảng tính preview */}
                  <div className="overflow-x-auto max-h-[500px] relative">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#111827] text-gray-400 border-b border-gray-800 font-mono sticky top-0 z-10 shadow-sm">
                          {columns.map(col => (
                            <th key={col} className="p-3 font-semibold text-center whitespace-nowrap min-w-[120px]">
                              {col === mapping.mota && "📝 "}{col === mapping.manganh && "🏷️ "}{col === mapping.xa && "🗺️ "}{col === mapping.idCol && "🔑 "}{col}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {paginatedData.map((row, rIdx) => (
                          <tr key={rIdx} className="border-b border-gray-800/40 hover:bg-gray-800/50 transition-colors">
                            {columns.map(col => {
                              const cellValue = row[col];
                              return (
                                <td key={col} className={`p-3 truncate max-w-[220px] text-center font-sans ${col === mapping.mota ? "text-slate-200 text-left" : "text-gray-300"}`} title={String(cellValue)}>
                                  {cellValue === null || cellValue === undefined ? "" : String(cellValue)}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {/* Thanh phân trang Pagination */}
                  <div className="flex items-center justify-between border-t border-[#374151] pt-4 text-xs">
                    <span className="text-gray-400">
                      Trang <strong className="text-white">{viewPage}</strong> / {totalPages}
                    </span>
                    <div className="flex gap-2">
                      <button 
                        disabled={viewPage === 1}
                        onClick={() => setViewPage(prev => Math.max(1, prev - 1))}
                        className={`px-3 py-1.5 rounded-lg border border-gray-700 font-semibold ${viewPage === 1 ? "bg-[#111827] text-gray-600 cursor-not-allowed" : "bg-[#111827] hover:bg-[#374151] text-gray-300 cursor-pointer"}`}
                      >
                        Trước
                      </button>
                      <button 
                        disabled={viewPage === totalPages}
                        onClick={() => setViewPage(prev => Math.min(totalPages, prev + 1))}
                        className={`px-3 py-1.5 rounded-lg border border-gray-700 font-semibold ${viewPage === totalPages ? "bg-[#111827] text-gray-600 cursor-not-allowed" : "bg-[#111827] hover:bg-[#374151] text-gray-300 cursor-pointer"}`}
                      >
                        Sau
                      </button>
                    </div>
                  </div>

                </div>
              ) : (
                <div className="bg-[#1f2937]/40 border-2 border-dashed border-[#374151] p-12 text-center rounded-2xl space-y-4">
                  <Database className="w-12 h-12 text-[#4b5563] mx-auto animate-pulse" />
                  <div>
                    <h4 className="text-base font-bold text-white">Chưa có cơ sở dữ liệu nạp vào</h4>
                    <p className="text-xs text-gray-400 max-w-md mx-auto pt-1 leading-relaxed">
                      Hãy chọn "Tải tệp dữ liệu chính" ở ô phía trên để nạp bảng tài liệu và kích hoạt toàn bộ cơ cấu.
                    </p>
                  </div>
                </div>
              )}

            </div>
          )}

          {/* 3. TAB GHÉP NỐI DỮ LIỆU */}
          {activeTab === "ghepnoi" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <GitMerge className="w-5 h-5 text-blue-400" /> GHÉP NỐI HAI BẢNG TẬP DỮ LIỆU
                </h3>
                <p className="text-xs text-gray-400">Kết hợp hai tệp dữ liệu dựa theo trường khóa liên kết tương ứng (left outer join).</p>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  
                  {/* BẢNG TRÁI */}
                  <div className="bg-[#111827]/60 rounded-xl p-5 border border-blue-500/10 space-y-4 text-center">
                    <h4 className="text-sm font-bold text-blue-400">📊 BẢNG TRÁI (DỮ LIỆU CHÍNH)</h4>
                    <label className="inline-block bg-[#1f2937] hover:bg-[#374151] border border-blue-500/30 text-xs text-blue-300 font-semibold px-4 py-2.5 rounded-lg cursor-pointer transition-all">
                      Chọn File Trái
                      <input type="file" onChange={(e) => handleFileUpload(e, "left")} className="hidden" />
                    </label>
                    <div className="text-xs text-gray-400 font-mono select-none">{leftFileName ? `📂 ${leftFileName} (${leftData.length} dòng)` : "Chưa tải bảng trái"}</div>
                    
                    {leftData.length > 0 && (
                      <div className="text-left space-y-1">
                        <label className="text-[11px] font-bold text-gray-500 block">Chọn cột khóa chính bên Trái</label>
                        <select 
                          value={leftKey} 
                          onChange={(e) => setLeftKey(e.target.value)}
                          className="w-full bg-[#111827] border border-[#374151] rounded-lg px-2.5 py-1.5 text-xs text-white"
                        >
                          <option value="">-- Chọn Khóa --</option>
                          {Object.keys(leftData[0] || {}).map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                    )}
                  </div>

                  {/* BẢNG PHẢI */}
                  <div className="bg-[#111827]/60 rounded-xl p-5 border border-teal-500/10 space-y-4 text-center">
                    <h4 className="text-sm font-bold text-teal-400">📊 BẢNG PHẢI (THÔNG TIN GHÉP THÊM)</h4>
                    <label className="inline-block bg-[#1f2937] hover:bg-[#374151] border border-teal-500/30 text-xs text-teal-300 font-semibold px-4 py-2.5 rounded-lg cursor-pointer transition-all">
                      Chọn File Phải
                      <input type="file" onChange={(e) => handleFileUpload(e, "right")} className="hidden" />
                    </label>
                    <div className="text-xs text-gray-400 font-mono select-none">{rightFileName ? `📂 ${rightFileName} (${rightData.length} dòng)` : "Chưa tải bảng phải"}</div>
                    
                    {rightData.length > 0 && (
                      <div className="text-left space-y-1">
                        <label className="text-[11px] font-bold text-gray-500 block">Chọn cột khóa liên kết bên Phải</label>
                        <select 
                          value={rightKey} 
                          onChange={(e) => setRightKey(e.target.value)}
                          className="w-full bg-[#111827] border border-[#374151] rounded-lg px-2.5 py-1.5 text-xs text-white"
                        >
                          <option value="">-- Chọn Khóa --</option>
                          {Object.keys(rightData[0] || {}).map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                    )}
                  </div>

                </div>

                <div className="pt-4 border-t border-gray-800 flex justify-end">
                  <button 
                    onClick={handleMerge}
                    className="bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-blue-900/30 font-sans cursor-pointer flex items-center gap-1.5"
                  >
                    <GitMerge className="w-4 h-4" /> THỰC THI GHÉP NỐI (LEFT OUTER JOIN)
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* 4. TAB SO SÁNH CŨ MỚI (DIFF) */}
          {activeTab === "sosanh" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <Combine className="w-5 h-5 text-cyan-400" /> SO SÁNH HAI FILE DỮ LIỆU CŨ & MỚI
                </h3>
                <p className="text-xs text-gray-400">Rà soát và đánh dấu trạng thái thay đổi ("Mới thêm", "Đã xóa", "Lệch thay đổi") dựa vào cột mã định danh chung.</p>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  
                  {/* FILE CŨ */}
                  <div className="bg-[#111827]/60 rounded-xl p-5 border border-gray-800 space-y-4 text-center">
                    <h4 className="text-sm font-bold text-gray-400">📁 FILE DỮ LIỆU BẢN CŨ</h4>
                    <label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-white border border-[#4b5563] font-semibold px-4 py-2.5 rounded-lg cursor-pointer transition-all">
                      Tải File Cũ
                      <input type="file" onChange={(e) => handleFileUpload(e, "old")} className="hidden" />
                    </label>
                    <div className="text-xs text-gray-400 font-mono select-none">{oldFileName ? `📂 ${oldFileName} (${oldData.length} dòng)` : "Chưa tải file cũ"}</div>
                  </div>

                  {/* FILE MỚI */}
                  <div className="bg-[#111827]/60 rounded-xl p-5 border border-cyan-500/10 space-y-4 text-center">
                    <h4 className="text-sm font-bold text-cyan-400">📁 FILE DỮ LIỆU BẢN MỚI</h4>
                    <label className="inline-block bg-[#1f2937] hover:bg-[#374151] border border-cyan-500/30 text-xs text-cyan-300 font-semibold px-4 py-2.5 rounded-lg cursor-pointer transition-all">
                      Tải File Mới
                      <input type="file" onChange={(e) => handleFileUpload(e, "new")} className="hidden" />
                    </label>
                    <div className="text-xs text-gray-400 font-mono select-none">{newFileName ? `📂 ${newFileName} (${newData.length} dòng)` : "Chưa tải file mới"}</div>
                  </div>

                </div>

                {oldData.length > 0 && newData.length > 0 && (
                  <div className="max-w-md space-y-1 bg-[#111827]/80 rounded-xl p-4 border border-[#374151] mx-auto">
                    <label className="text-xs font-bold text-gray-400 block">Chọn Cột Khóa chính đối chiếu độc nhất</label>
                    <select 
                      value={diffKey} 
                      onChange={(e) => setDiffKey(e.target.value)}
                      className="w-full bg-[#1f2937] border border-[#374151] rounded-lg px-2.5 py-1.5 text-xs text-white"
                    >
                      <option value="">-- Chọn cột khóa --</option>
                      {/* Lấy các cột chung của cả 2 bảng */}
                      {Object.keys(oldData[0] || {}).filter(c => Object.keys(newData[0] || {}).includes(c)).map(c => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>
                )}

                <div className="pt-4 border-t border-gray-800 flex justify-end">
                  <button 
                    onClick={handleCompare}
                    className="bg-cyan-600 hover:bg-cyan-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-cyan-900/30 font-sans cursor-pointer flex items-center gap-1.5"
                  >
                    <Combine className="w-4 h-4" /> BẮT ĐẦU SO SÁNH & DIFF
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* 5. TAB TÁCH DỮ LIỆU THEO CỘT */}
          {activeTab === "tachfile" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <Scissors className="w-5 h-5 text-pink-400" /> TÁCH FILE HÀNG LOẠT THEO CỘT CHỈ ĐỊNH
                </h3>
                <p className="text-xs text-gray-400">Chia nhỏ bảng tính lớn của bạn thành nhiều file Excel riêng biệt dựa trên giá trị cột đã chọn (ví dụ: tách theo từng Địa Phương Xã) và đóng gói tải xuống ZIP.</p>

                {mainData.length > 0 ? (
                  <div className="max-w-md space-y-4 bg-[#111827] rounded-xl p-5 border border-[#374151]">
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-gray-400 block">Chọn cột để định nghĩa tách file</label>
                      <select 
                        value={splitCol} 
                        onChange={(e) => setSplitCol(e.target.value)}
                        className="w-full bg-[#1f2937] border border-[#374151] rounded-lg px-2.5 py-1.5 text-xs text-white"
                      >
                        <option value="">-- Chọn cột --</option>
                        {columns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <button 
                      onClick={handleSplitData}
                      className="w-full bg-pink-600 hover:bg-pink-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-pink-900/30 flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      <Scissors className="w-4 h-4" /> KHỞI CHẠY BẮT ĐẦU TÁCH HÀNG LOẠT & ZIP DOWNLOAD
                    </button>
                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950 font-sans">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước ở Tab "Xem & Định Nghĩa Cột"!
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 6. TAB TỔNG HỢP BÁO CÁO ĐỘNG */}
          {activeTab === "tonghop" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-amber-400" /> TỔNG HỢP BÁO CÁO THEO NGÀNH CẤP 1 & CẤP 2 (KẾT HỢP ĐỊA BÀN XÃ)
                  </h3>
                  <p className="text-xs text-gray-400 font-sans">
                    Phân hệ hạch toán tổng hợp chuyên sâu cho phép quy thuộc ngành từ mã ngành bất kỳ lên cấp 1 (lĩnh vực lớn A-U) hoặc tách thành ngành cấp 2 (2 số đầu), sau đó cộng gom doanh thu, quy mô lao động theo từng xã địa phương dựa trên danh mục nạp vào.
                  </p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6">
                    {/* BỘ LỰA CHỌN CỘT THỦ CÔNG */}
                    <div className="bg-[#111827]/80 p-5 rounded-xl border border-gray-850 space-y-4 shadow-lg">
                      <span className="text-xs font-bold text-cyan-400 tracking-wider uppercase font-mono block">
                        Cấu hình các cột đầu vào (Chỉ định rõ cột, không đoán bừa bãi)
                      </span>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Mã Ngành VSIC:</label>
                          <select 
                            value={quickReportManganhCol} 
                            onChange={(e) => setQuickReportManganhCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột chứa mã ngành --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Xã / Địa Bàn:</label>
                          <select 
                            value={quickReportXaCol} 
                            onChange={(e) => setQuickReportXaCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột xã/phường/địa bàn --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Doanh Thu (Tùy chọn):</label>
                          <select 
                            value={quickReportDoanhThuCol} 
                            onChange={(e) => setQuickReportDoanhThuCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-sans"
                          >
                            <option value="">-- Chọn cột doanh thu (bỏ trống nếu không tính) --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Lao Động (Tùy chọn):</label>
                          <select 
                            value={quickReportLaoDongCol} 
                            onChange={(e) => setQuickReportLaoDongCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-sans"
                          >
                            <option value="">-- Chọn cột lao động (bỏ trống nếu không tính) --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                      </div>
                    </div>

                    {/* HÌNH THỨC TRÌNH BÀY BÁO CÁO */}
                    <div className="bg-[#111827]/80 p-5 rounded-xl border border-gray-850 space-y-3">
                      <span className="text-xs font-bold text-amber-400 tracking-wider uppercase font-mono block">
                        Cấu hình định dạng hạch toán đầu ra
                      </span>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <label className="flex items-start gap-3 text-xs text-gray-300 hover:text-white cursor-pointer select-none bg-[#111827]/50 p-3 rounded-lg border border-gray-800">
                          <input 
                            type="radio" 
                            name="quickReportFormatPivot"
                            checked={reportType === "pivot"} 
                            onChange={() => setReportType("pivot")}
                            className="mt-1 text-amber-500 focus:ring-amber-500 bg-gray-900 border-gray-750"
                          />
                          <div>
                            <div className="font-bold text-gray-100 font-sans">Bảng xoay ngang Pivot (Khuyên dùng)</div>
                            <div className="text-[10.5px] text-gray-400 mt-1 font-sans">
                              Mỗi xã địa bàn hiển thị thành một hàng ngang. Các ngành và chỉ tiêu doanh thu, lao động xếp kề nhau làm cột song song để dễ nhìn và hạch toán so sánh.
                            </div>
                          </div>
                        </label>

                        <label className="flex items-start gap-3 text-xs text-gray-300 hover:text-white cursor-pointer select-none bg-[#111827]/50 p-3 rounded-lg border border-gray-800">
                          <input 
                            type="radio" 
                            name="quickReportFormatPivot"
                            checked={reportType === "flat"} 
                            onChange={() => setReportType("flat")}
                            className="mt-1 text-amber-500 focus:ring-amber-500 bg-gray-900 border-gray-750"
                          />
                          <div>
                            <div className="font-bold text-gray-100 font-sans">Bảng phẳng danh sách truyền thống</div>
                            <div className="text-[10.5px] text-gray-400 mt-1 font-sans">
                              Mỗi dòng là một cặp địa bàn xã và ngành cụ thể cùng số DN, tổng doanh thu và tổng lao động tương ứng.
                            </div>
                          </div>
                        </label>
                      </div>
                    </div>

                    {/* NÚT THỰC THI CHẠY TỔNG HỢP & BÊ CỘT */}
                    <div className="space-y-3">
                      <div className="flex flex-col sm:flex-row gap-3">
                        <button 
                          onClick={() => handleQuickReport(1)}
                          className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3 rounded-xl transition-all shadow-md hover:shadow-emerald-900/30 cursor-pointer flex items-center justify-center gap-2 font-sans"
                        >
                          📈 Chạy Tổng Hợp Ngành Cấp 1 &amp; Xã
                        </button>

                        <button 
                          onClick={() => handleQuickReport(2)}
                          className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3 rounded-xl transition-all shadow-md hover:shadow-emerald-900/30 cursor-pointer flex items-center justify-center gap-2 font-sans"
                        >
                          📈 Chạy Tổng Hợp Ngành Cấp 2 &amp; Xã
                        </button>
                      </div>

                      <button
                        onClick={handleAppendSectorsToMainData}
                        className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold text-xs py-3.5 rounded-xl transition-all shadow-lg hover:shadow-indigo-900/30 flex items-center justify-center gap-2 cursor-pointer font-sans border border-indigo-500/30"
                        title="Đối khớp trực tiếp mã ngành đã chỉ định sang Danh mục đã nạp trong bộ nhớ, sau đó thêm mới các cột Ngành cấp 1 và Cấp 2 vào bảng tính gốc"
                      >
                        ⚡ BÊ THÔNG TIN NGÀNH CẤP 1 &amp; CẤP 2 SANG CỘT MỚI (TRÊN BẢNG TÍNH GỐC)
                      </button>
                    </div>

                    {/* BẢNG HIỂN THỊ KẾT QUẢ TRỰC TIẾP */}
                    {quickReportResultRows.length > 0 && (
                      <BeautifulReportTable
                        rows={quickReportResultRows}
                        cols={quickReportResultCols}
                        level={quickReportLevel}
                        reportType={reportType}
                        onExport={handleExportQuickReport}
                      />
                    )}
                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950 font-sans">
                    ⚠️ Vui lòng nạp dữ liệu chính ở trang đầu tiên trước khi chạy hạch toán tổng hợp.
                  </div>
                )}
              </div>
            </div>
          )}


          {/* 7. TAB CHUẨN HÓA & PHÂN TÍCH NGÀNH TRẬT TỰ AI */}
          {activeTab === "chuanhoanganh" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <Brain className="w-5 h-5 text-indigo-400 animate-pulse" /> CHUẨN HÓA & KHỚP MÃ NGÀNH THÔNG MINH
                  </h3>
                  <p className="text-xs text-gray-400">So khớp mã ngành đăng ký, rà soát mô tả so với nhóm mã gợi ý, tự động bẻ tách mã ngành ra thành 5 cấp chi tiết (từ Cấp 1-Chương rộng đến Cấp 5-Nhóm con) giải quyết triệt để lỗi thắt nút phân cấp đăng nghiệp.</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6 border-t border-gray-800 pt-6">
                    {/* BỘ ĐIỀU KHIỂN CHUẨN HÓA ĐỘC LẬP TÙY CHỌN CỘT */}
                    <div className="bg-[#111827]/85 p-6 rounded-2xl border border-indigo-500/20 shadow-xl space-y-5">
                      <div className="flex items-center gap-2 border-b border-gray-800 pb-3">
                        <CheckCircle2 className="w-5 h-5 text-indigo-400" />
                        <h4 className="text-sm font-bold text-white uppercase tracking-wider font-mono">1. Cấu hình & Kích hoạt chuẩn hóa mã ngành VSIC</h4>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                        {/* Chọn cột mã ngành */}
                        <div className="space-y-1.5">
                          <label className="text-xs font-bold text-gray-300 block font-mono">CHỌN CỘT CHỨA MÃ NGÀNH KINH TẾ (CẤP 5):</label>
                          <select
                            value={stdIndustryCol}
                            onChange={(e) => setStdIndustryCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-indigo-500 font-mono"
                          >
                            <option value="">-- Chọn cột mã ngành --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <span className="text-[10px] text-gray-500 block italic leading-normal">
                            Hệ thống sẽ chuẩn hóa mã ngành này sang cấp 5, thêm mô tả tự động đặt sát cạnh cột mô tả doanh nghiệp của bạn.
                          </span>
                        </div>

                        {/* Chọn cột mô tả */}
                        <div className="space-y-1.5">
                          <label className="text-xs font-bold text-gray-300 block font-mono">CHỌN CỘT CHI TIẾT MÔ TẢ HOẠT ĐỘNG KD (ĐỐI CHIẾU LOGIC LỆCH):</label>
                          <select
                            value={stdDescriptionCol}
                            onChange={(e) => setStdDescriptionCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-indigo-500"
                          >
                            <option value="">-- Chọn cột mô tả hoạt động --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <span className="text-[10px] text-gray-500 block italic leading-normal">
                            Hệ thống tự đối chiếu từ khóa trong mô tả với mã ngành chuẩn để cảnh báo các dòng nghi ngờ hạch toán lệch vai trò.
                          </span>
                        </div>
                      </div>

                      <div className="flex justify-end pt-2">
                        <button
                          onClick={handleStandardizeSectorsAndMatch}
                          className="bg-gradient-to-r from-indigo-600 to-[#4f46e5] hover:from-indigo-700 hover:to-[#4338ca] text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-md shadow-indigo-950/40 flex items-center gap-2 cursor-pointer border border-indigo-500/20 active:scale-95"
                        >
                          <Brain className="w-4 h-4 text-indigo-200" /> ⚡ THỰC HIỆN SO KHỚP & CHUẨN HÓA TOÀN BỘ NGÀNH
                        </button>
                      </div>
                    </div>

                    {/* BỘ ĐIỀU KHIỂN ĐỐI CHIẾU SONG SONG HAI CỘT TÙY CHỌN (Khớp & Đánh Dấu Sai Lệch) */}
                    <div className="bg-[#111827]/85 p-6 rounded-2xl border border-cyan-500/20 shadow-xl space-y-5">
                      <div className="flex items-center gap-2 border-b border-gray-800 pb-3">
                        <Combine className="w-5 h-5 text-cyan-400" />
                        <h4 className="text-sm font-bold text-white uppercase tracking-wider font-mono">2. ĐỐI CHIẾU SONG SONG HAI CỘT TÙY CHỌN & ĐÁNH DẤU SAI LỆCH</h4>
                      </div>

                      <p className="text-xs text-gray-400">
                        Chỉ định bất kỳ 2 cột nào trong tệp dữ liệu đã nạp (ví dụ: so sánh giữa <em>"Mã ngành tự gõ"</em> với <em>"Mã ngành do AI gợi ý"</em> hoặc <em>"Tên ngành doanh nghiệp khai"</em> với <em>"Tên ngành chuẩn VSIC"</em>). Chương trình sẽ rà lỗi song song, đánh dấu trạng thái và ghi chú ngay lập tức vào bảng tính.
                      </p>

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                        {/* Chọn cột 1 */}
                        <div className="space-y-1.5">
                          <label className="text-xs font-bold text-cyan-400 block font-mono">CHỌN CỘT THỨ NHẤT (CỘT A):</label>
                          <select
                            value={crossCompareColA}
                            onChange={(e) => setCrossCompareColA(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-mono"
                          >
                            <option value="">-- Chọn Cột A --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <span className="text-[10px] text-gray-500 block italic leading-normal">
                            Giá trị gốc đem so sánh.
                          </span>
                        </div>

                        {/* Chọn cột 2 */}
                        <div className="space-y-1.5">
                          <label className="text-xs font-bold text-cyan-400 block font-mono">CHỌN CỘT THỨ HAI (CỘT B):</label>
                          <select
                            value={crossCompareColB}
                            onChange={(e) => setCrossCompareColB(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-mono"
                          >
                            <option value="">-- Chọn Cột B --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <span className="text-[10px] text-gray-500 block italic leading-normal">
                            Giá trị đối chứng tin cậy.
                          </span>
                        </div>

                        {/* Quy luật so khớp */}
                        <div className="space-y-1.5">
                          <label className="text-xs font-bold text-gray-300 block font-mono">CHỌN LUẬT ĐỐI CHIẾU LỆCH:</label>
                          <select
                            value={crossCompareRule}
                            onChange={(e) => setCrossCompareRule(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-700 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-sans"
                          >
                            <option value="normalize">Chuẩn hóa chữ thường + Bỏ dấu cách dư thừa (Khuyên dùng cho tên ngành)</option>
                            <option value="exact">Trùng khớp chính xác tuyệt đối (Phân biệt hoa thường)</option>
                            <option value="sector_code">So khớp mã ngành VSIC (Lọc chỉ lấy số, chấp nhận quy nạp cha-con)</option>
                            <option value="substring">Chứa cụm từ của nhau (Substring match)</option>
                          </select>
                          <span className="text-[10px] text-gray-500 block italic leading-normal">
                            Phương thức logic để lọc và đánh dấu lỗi lệch.
                          </span>
                        </div>
                      </div>

                      <div className="flex justify-end pt-2">
                        <button
                          onClick={handleCrossColumnCompare}
                          className="bg-gradient-to-r from-cyan-600 to-teal-600 hover:from-cyan-700 hover:to-teal-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-md shadow-cyan-950/40 flex items-center gap-2 cursor-pointer border border-cyan-500/20 active:scale-95"
                        >
                          <Combine className="w-4 h-4 text-cyan-200" /> ⚡ THỰC THI KIỂM TRA ĐỔI CHIẾU & ĐÁNH DẤU SAI LỆCH
                        </button>
                      </div>
                    </div>

                    {/* HIỂN THỊ THÔNG SỐ ĐỐI CHIẾU CHÉO */}
                    {crossCompareStats.total > 0 && (
                      <div className="bg-[#1e293b]/25 border border-cyan-500/10 p-5 rounded-2xl grid grid-cols-3 gap-4 text-center">
                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-gray-400 font-bold uppercase font-mono mb-1">Tổng bản ghi đối chiếu</div>
                          <div className="text-lg font-bold text-white font-mono">{crossCompareStats.total}</div>
                        </div>

                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-emerald-400 font-bold uppercase font-mono mb-1">Khớp trùng nhau</div>
                          <div className="text-lg font-bold text-emerald-400 font-mono">
                            {crossCompareStats.matchCount} <span className="text-[10px] font-normal text-gray-400">({Math.round(crossCompareStats.matchCount / crossCompareStats.total * 100)}%)</span>
                          </div>
                        </div>

                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-red-500/10">
                          <div className="text-[10px] text-rose-400 font-bold uppercase font-mono mb-1">Phát hiện Lệch sai</div>
                          <div className="text-lg font-bold text-rose-400 font-mono">
                            {crossCompareStats.mismatchCount} <span className="text-[10px] font-normal text-gray-400">({Math.round(crossCompareStats.mismatchCount / crossCompareStats.total * 100)}%)</span>
                          </div>
                        </div>
                      </div>
                    )}

                    {/* DANH SÁCH LỖI LỆCH SONG SONG PHÁT HIỆN ĐƯỢC */}
                    {crossCompareAnomalies.length > 0 && (
                      <div className="bg-[#1e293b]/40 border border-rose-500/20 rounded-2xl p-5 space-y-4">
                        <div className="flex items-center justify-between border-b border-gray-800 pb-3">
                          <div className="flex items-center gap-2">
                            <span className="bg-rose-950 text-rose-400 font-mono text-[10px] px-2.5 py-0.5 rounded-full border border-rose-800 animate-pulse">Lệch sai</span>
                            <h4 className="text-sm font-bold text-white uppercase">Danh sách phát hiện lệch giữa 2 cột ({crossCompareAnomalies.length} dòng lỗi)</h4>
                          </div>
                        </div>

                        <div className="overflow-x-auto border border-gray-850 rounded-xl bg-gray-950/40 max-h-[300px]">
                          <table className="w-full text-left text-xs min-w-[700px]">
                            <thead>
                              <tr className="bg-[#1f2937] border-b border-gray-850 text-gray-300 font-mono text-[11px] sticky top-0 z-10">
                                <th className="p-3 text-center w-[60px]">DÒNG</th>
                                <th className="p-3 w-[150px]">MÃ DOANH NGHIỆP</th>
                                <th className="p-3">GIÁ TRỊ CỘT A ({crossCompareColA})</th>
                                <th className="p-3">GIÁ TRỊ CỘT B ({crossCompareColB})</th>
                                <th className="p-3 text-amber-400">LUẬT BÁO LỆCH</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-800/45 text-gray-200">
                              {crossCompareAnomalies.map((item, idx) => (
                                <tr key={idx} className="hover:bg-red-950/5 transition-colors">
                                  <td className="p-3 text-center font-mono text-gray-400">{item.dongSTT}</td>
                                  <td className="p-3 font-semibold text-gray-300">{item.maDN}</td>
                                  <td className="p-3 font-sans text-red-300 bg-red-950/10 font-medium">{item.valA}</td>
                                  <td className="p-3 font-sans text-emerald-300 bg-emerald-950/5">{item.valB}</td>
                                  <td className="p-3 font-sans text-amber-300 italic">{item.reason}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        <div className="text-[11px] text-gray-400">
                          💡 <strong>Mẹo nhỏ:</strong> Tệp tin hiện tại đã được tự động thêm 2 cột: <strong>"Đối Chiếu [{crossCompareColA}] vs [{crossCompareColB}]"</strong> và <strong>"Đánh Dấu Lệch [{crossCompareColA}] vs [{crossCompareColB}]"</strong>. Bạn có thể sang tab <em>"Xem Dữ Liệu"</em> để kiểm tra bảng hoặc tải trực tiếp File Excel về máy để có dòng cảnh báo này trong bảng tính!
                        </div>
                      </div>
                    )}

                    {/* HIỂN THỊ KẾT QUẢ THỐNG KÊ SO KHỚP CHUẨN VSIC */}
                    {stdMatchStats.total > 0 && (
                      <div className="bg-[#1e293b]/35 border border-indigo-500/10 p-5 rounded-2xl grid grid-cols-2 md:grid-cols-4 gap-4 text-center">
                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-gray-400 font-bold uppercase font-mono mb-1">Tổng bản ghi rà soát</div>
                          <div className="text-xl font-bold text-white font-mono">{stdMatchStats.total}</div>
                        </div>

                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-emerald-400 font-bold uppercase font-mono mb-1">Khớp chuẩn VSIC</div>
                          <div className="text-xl font-bold text-emerald-400 font-mono">{stdMatchStats.valid} <span className="text-[10px] font-normal text-gray-400">({Math.round(stdMatchStats.valid / stdMatchStats.total * 100)}%)</span></div>
                        </div>

                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-red-400 font-bold uppercase font-mono mb-1">Lỗi mã VSIC</div>
                          <div className="text-xl font-bold text-red-400 font-mono">{stdMatchStats.invalid}</div>
                        </div>

                        <div className="bg-[#111827]/60 p-3 rounded-xl border border-gray-800">
                          <div className="text-[10px] text-amber-400 font-bold uppercase font-mono mb-1">Nghi ngờ bất nhất logic</div>
                          <div className="text-xl font-bold text-amber-400 font-mono">{stdMatchStats.conflicts}</div>
                        </div>
                      </div>
                    )}

                    {/* DANH SÁCH BẢN GHI PHÁT HIỆN LỆCH LOGIC CHÉO */}
                    {stdReportAnomalies.length > 0 && (
                      <div className="bg-[#1e293b]/40 border border-amber-500/20 rounded-2xl p-5 space-y-4">
                        <div className="flex items-center justify-between border-b border-gray-800 pb-3">
                          <div className="flex items-center gap-2">
                            <span className="bg-amber-950 text-amber-300 font-mono text-[10px] px-2.5 py-0.5 rounded-full border border-amber-800">CẢNH BÁO LỆCH MÃ</span>
                            <h4 className="text-sm font-bold text-white">DANH SÁCH PHÁT HIỆN XUNG ĐỘT QUY LUẬT LOGIC ({stdReportAnomalies.length} DÀN Ý)</h4>
                          </div>
                        </div>

                        <div className="overflow-x-auto border border-gray-850 rounded-xl bg-gray-950/40 max-h-[300px]">
                          <table className="w-full text-left text-xs min-w-[750px]">
                            <thead>
                              <tr className="bg-[#1f2937] border-b border-gray-850 text-gray-300 font-mono text-[11px] sticky top-0 z-10">
                                <th className="p-3 text-center w-[60px]">DÒNG</th>
                                <th className="p-3 w-[120px]">MÃ DOANH NGHIỆP</th>
                                <th className="p-3 w-[100px] text-center">MÃ NGÀNH NHẬP</th>
                                <th className="p-3">MÔ TẢ HOẠT ĐỘNG THỰC TẾ</th>
                                <th className="p-3">TÊN NGÀNH CHUẨN (VSIC)</th>
                                <th className="p-3 text-amber-400">PHÁT HIỆN MÂU THUẪN LOGIC</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-800/45 text-gray-200">
                              {stdReportAnomalies.map((item, idx) => (
                                <tr key={idx} className="hover:bg-gray-800/10 transition-colors">
                                  <td className="p-3 text-center font-mono text-gray-400">{item.dongSTT}</td>
                                  <td className="p-3 font-semibold text-gray-300">{item.maDN}</td>
                                  <td className="p-3 text-center font-mono font-bold text-red-400 bg-red-950/5">{item.maGoc}</td>
                                  <td className="p-3 font-sans line-clamp-2 max-w-[250px] truncate" title={item.motaGoc}>{item.motaGoc}</td>
                                  <td className="p-3 font-sans text-gray-300">{item.nganhChuan}</td>
                                  <td className="p-3 font-sans text-amber-300 bg-amber-950/10 italic font-medium">{item.phanTichloi}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {/* CÔNG CỤ TỰ ĐỘNG PHÂN TÍCH BẤT NHẤT QUÁN ĐA CHIỀU (MÃ ⇄ MÔ TẢ) */}
                    <div className="bg-[#111827]/80 rounded-2xl p-6 border border-purple-500/20 space-y-6">
                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-gray-800/40 pb-4">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="bg-pink-950/50 border border-pink-500/30 text-pink-400 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md animate-pulse">Giám sát chéo</span>
                            <h3 className="text-base font-bold text-white flex items-center gap-2">
                              <Activity className="w-5 h-5 text-pink-400" /> CÔNG CỤ TỰ ĐỘNG PHÂN TÍCH BẤT NHẤT QUÁN QUY LUẬT (MÃ VSIC ⇄ MÔ TẢ)
                            </h3>
                          </div>
                          <p className="text-xs text-gray-400 mt-1">
                            Tính năng tự động dò quét và tính phân nhám chéo đối xử tức thì để định danh 2 lỗi mâu thuẫn khét tiếng nhất khi gõ ngành nghề cho nhân viên thống kê.
                          </p>
                        </div>
                      </div>

                      {/* STATS BENTO ROW */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div 
                          onClick={() => setInconsistenciesTab("desc")}
                          className={`p-4 rounded-xl border transition-all cursor-pointer flex items-start gap-3 ${
                            inconsistenciesTab === "desc" 
                              ? "bg-purple-950/20 border-purple-500/40 shadow-sm" 
                              : "bg-[#1f2937]/30 border-gray-800 hover:border-gray-700 hover:bg-[#1f2937]/50"
                          }`}
                        >
                          <div className="bg-purple-950/50 border border-purple-500/30 p-2 rounded-lg text-purple-400">
                            <AlertTriangle className="w-4 h-4 text-purple-400" />
                          </div>
                          <div className="space-y-1">
                            <div className="text-xs text-gray-400 font-medium">Lỗi Loại A: Cùng Mô tả ➔ Khác Mã ngành</div>
                            <div className="text-lg font-bold text-white flex items-center gap-2">
                              {inconAnalysis.descToCodes.length} <span className="text-[10px] font-normal text-purple-400 bg-purple-950/60 px-2 py-0.5 rounded-full font-mono">bản ghi mâu thuẫn</span>
                            </div>
                            <p className="text-[10px] text-gray-500 leading-relaxed font-sans">Một mô tả hoạt động viết giống hệt nhau nhưng lại bị phân gán sang các mã ngành nghề khác nhau ở từng dòng.</p>
                          </div>
                        </div>

                        <div 
                          onClick={() => setInconsistenciesTab("code")}
                          className={`p-4 rounded-xl border transition-all cursor-pointer flex items-start gap-3 ${
                            inconsistenciesTab === "code" 
                              ? "bg-pink-950/20 border-pink-500/40 shadow-sm" 
                              : "bg-[#1f2937]/30 border-gray-800 hover:border-gray-700 hover:bg-[#1f2937]/50"
                          }`}
                        >
                          <div className="bg-pink-950/50 border border-pink-500/30 p-2 rounded-lg text-pink-400">
                            <Combine className="w-4 h-4 text-pink-400" />
                          </div>
                          <div className="space-y-1">
                            <div className="text-xs text-gray-400 font-medium">Lỗi Loại B: Cùng Mã ngành ➔ Trái Loại hình</div>
                            <div className="text-lg font-bold text-white flex items-center gap-2">
                              {inconAnalysis.codeToDescs.length} <span className="text-[10px] font-normal text-pink-400 bg-pink-950/60 px-2 py-0.5 rounded-full font-mono">nhóm ngành xung đột</span>
                            </div>
                            <p className="text-[10px] text-gray-500 leading-relaxed font-sans">Trong cùng một mã ngành hạch toán, dòng thì mô tả "Sản xuất/Gia công" nhưng dòng khác lại mô tả "Bán buôn/Bán lẻ".</p>
                          </div>
                        </div>
                      </div>

                      {/* MAIN SCROLLER RESULTS DISPLAY */}
                      <div className="bg-[#111827] border border-gray-800 rounded-xl p-4 min-h-[220px] max-h-[350px] overflow-y-auto space-y-3">
                        {inconsistenciesTab === "desc" ? (
                          <>
                            <div className="flex items-center justify-between border-b border-gray-800 pb-2">
                              <span className="text-xs font-bold text-purple-300">DANH SÁCH MÔ TẢ TRÙNG NHAU - KHÁC MÃ NGÀNH GÁN ({inconAnalysis.descToCodes.length})</span>
                              <span className="text-[10px] font-mono text-gray-500">Tìm thấy mâu thuẫn nhập liệu</span>
                            </div>

                            {inconAnalysis.descToCodes.length === 0 ? (
                              <div className="text-xs text-emerald-400 flex items-center justify-center h-44 gap-1.5 font-mono">
                                <CheckCircle2 className="w-4 h-4 text-emerald-400 animate-pulse" /> Tuyệt vời! Không phát hiện mâu thuẫn "Cùng một mô tả gán khác mã".
                              </div>
                            ) : (
                              <div className="space-y-3 divide-y divide-gray-900">
                                {inconAnalysis.descToCodes.map((item, idx) => (
                                  <div key={idx} className="pt-3 first:pt-0 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
                                    <div className="space-y-1.5 flex-1">
                                      <div className="font-semibold text-white bg-[#1f2937]/40 px-2.5 py-1.5 rounded-lg border border-gray-800 leading-relaxed">
                                        📝 "{item.motaText}" <span className="text-gray-500 text-[10px] ml-1">({item.occurrences} lần khai giống nhau)</span>
                                      </div>
                                      
                                      <div className="flex flex-wrap gap-2 items-center text-[11px]">
                                        <span className="text-gray-500">Được gán mã chép chéo chập chĩnh:</span>
                                        {item.codes.map((c, cidx) => (
                                          <span key={cidx} className="bg-purple-950/40 text-purple-300 px-2 py-0.5 rounded border border-purple-500/20 font-mono flex items-center gap-1">
                                            <strong>{c.code}</strong> 
                                            <span className="text-[9px] text-gray-500">({c.count} dòng)</span>
                                          </span>
                                        ))}
                                      </div>
                                    </div>

                                    <button
                                      onClick={() => {
                                        setSearchTerm(item.motaText);
                                        setViewPage(1);
                                        setIsConfigExpanded(false);
                                        setActiveTab("xemdulieu");
                                      }}
                                      className="bg-purple-600 hover:bg-purple-700 text-white px-3 py-1.5 rounded-lg text-[10px] font-bold self-start md:self-center cursor-pointer transition-all flex items-center gap-1 active:scale-95 whitespace-nowrap"
                                    >
                                      <Search className="w-3 h-3" /> Lọc dòng gốc
                                    </button>
                                  </div>
                                ))}
                              </div>
                            )}
                          </>
                        ) : (
                          <>
                            <div className="flex items-center justify-between border-b border-gray-800 pb-2">
                              <span className="text-xs font-bold text-pink-300 font-sans">DANH SÁCH MÃ VSIC MÂU THUẪN LOẠI HÌNH HOẠT ĐỘNG ({inconAnalysis.codeToDescs.length})</span>
                              <span className="text-[10px] font-mono text-gray-500">Xung đột đặc trưng kinh doanh [Sản xuất vs Thương mại]</span>
                            </div>

                            {inconAnalysis.codeToDescs.length === 0 ? (
                              <div className="text-xs text-emerald-400 flex items-center justify-center h-44 gap-1.5 font-mono">
                                <CheckCircle2 className="w-4 h-4 text-emerald-400 animate-pulse" /> Tuyệt vời! Không có mâu thuẫn "Sản xuất vs Thương mại" chập chéo chung một mã.
                              </div>
                            ) : (
                              <div className="space-y-4 divide-y divide-gray-900">
                                {inconAnalysis.codeToDescs.map((item, idx) => (
                                  <div key={idx} className="pt-4 first:pt-0 space-y-2.5">
                                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-2 border-l-2 border-pink-500 pl-2">
                                      <div>
                                        <span className="bg-pink-900/40 text-pink-300 text-[10px] font-mono font-bold px-1.5 py-0.5 rounded border border-pink-500/20 mr-2">
                                          {item.code}
                                        </span>
                                        <span className="text-xs font-medium text-white">{item.tenNganh}</span>
                                      </div>
                                      <button
                                        onClick={() => {
                                          setSearchTerm(item.code);
                                          setViewPage(1);
                                          setIsConfigExpanded(false);
                                          setActiveTab("xemdulieu");
                                        }}
                                        className="bg-pink-600 hover:bg-pink-700 text-white px-2.5 py-1 rounded-[6px] text-[10px] font-bold cursor-pointer transition-all flex items-center gap-1 self-start whitespace-nowrap"
                                      >
                                        <Search className="w-2.5 h-2.5" /> Lọc dòng gốc của mã
                                      </button>
                                    </div>

                                    {/* PHẤN CHIA DANH SÁCH CHI TIẾT */}
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pl-2.5">
                                      <div className="bg-[#1f2937]/30 border border-gray-900 rounded-lg p-2.5 space-y-1.5">
                                        <div className="text-[10px] font-bold text-amber-400 uppercase tracking-wider font-mono flex items-center gap-1">
                                          ⚙️ SẢN XUẤT / GIA CÔNG
                                        </div>
                                        <div className="space-y-1">
                                          {item.conflicts.filter(c => c.type === "SẢN XUẤT").slice(0, 3).map((c, cidx) => (
                                            <div key={cidx} className="text-[11px] text-gray-300 bg-[#111827] px-2 py-1 rounded text-ellipsis overflow-hidden whitespace-nowrap">
                                              Dòng {c.rowIdx + 1}: "{c.descText}"
                                            </div>
                                          ))}
                                          {item.conflicts.filter(c => c.type === "SẢN XUẤT").length > 3 && (
                                            <div className="text-[9px] text-gray-500 pl-1 font-sans">Và thêm {item.conflicts.filter(c => c.type === "SẢN XUẤT").length - 3} dòng khác...</div>
                                          )}
                                        </div>
                                      </div>

                                      <div className="bg-[#1f2937]/30 border border-gray-900 rounded-lg p-2.5 space-y-1.5">
                                        <div className="text-[10px] font-bold text-blue-400 uppercase tracking-wider font-mono flex items-center gap-1">
                                          🏪 THƯƠNG MẠI / CỬA HÀNG BÁN
                                        </div>
                                        <div className="space-y-1">
                                          {item.conflicts.filter(c => c.type === "THƯƠNG MẠI").slice(0, 3).map((c, cidx) => (
                                            <div key={cidx} className="text-[11px] text-gray-300 bg-[#111827] px-2 py-1 rounded text-ellipsis overflow-hidden whitespace-nowrap">
                                              Dòng {c.rowIdx + 1}: "{c.descText}"
                                            </div>
                                          ))}
                                          {item.conflicts.filter(c => c.type === "THƯƠNG MẠI").length > 3 && (
                                            <div className="text-[9px] text-gray-500 pl-1 font-sans">Và thêm {item.conflicts.filter(c => c.type === "THƯƠNG MẠI").length - 3} dòng khác...</div>
                                          )}
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            )}
                          </>
                        )}
                      </div>
                    </div>

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 8. TAB CỖ MÁY KIỂM TRA LOGIC ĐA ĐIỀU KIỆN */}
          {activeTab === "kiemtralogic" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <CheckSquare className="w-5 h-5 text-emerald-400" /> CỖ MÁY KIỂM TRA LOGIC ĐA CHỈ TIÊU (RULE ENGINE)
                  </h3>
                  <p className="text-xs text-gray-400">Tạo ra các quy tắc ràng buộc rà quét dữ liệu dạng: NẾU thỏa mãn (Điều kiện bước 1) THÌ PHẢI bắt buộc thỏa mãn (Điều kiện bước 2). Hệ thống tự rà rà soát và ghi chú dòng sai phạm vào cột "Loi_Logic".</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6 border-t border-gray-800 pt-6">
                    
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                      
                      {/* ĐIỀU KIỆN 1: BƯỚC NẾU */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-indigo-500/10 space-y-4">
                        <div className="text-xs font-bold text-indigo-400 uppercase tracking-wider font-mono">BƯỚC 1: ĐIỀU KIỆN NẾU (IF CONDITIONS)</div>
                        <p className="text-[11px] text-gray-400">Thiết lập các điều kiện để tìm mục tiêu cần kiểm tra rà soát:</p>
                        
                        <div className="grid grid-cols-3 gap-2">
                          <select 
                            value={newIfRule.col} 
                            onChange={(e) => setNewIfRule({ ...newIfRule, col: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="">Cột</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          
                          <select 
                            value={newIfRule.op} 
                            onChange={(e) => setNewIfRule({ ...newIfRule, op: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="==">bằng (==)</option>
                            <option value="!=">khác (!=)</option>
                            <option value=">">lớn hơn (&gt;)</option>
                            <option value="<">nhỏ hơn (&lt;)</option>
                            <option value=">=">lớn hơn bằng (&gt;=)</option>
                            <option value="<=">nhỏ hơn bằng (&lt;=)</option>
                            <option value="chứa">chứa (string)</option>
                            <option value="không chứa">không chứa (string)</option>
                            <option value="trống">để rỗng (empty)</option>
                            <option value="không trống">có dữ liệu</option>
                          </select>

                          {newIfRule.op !== "trống" && newIfRule.op !== "không trống" && (
                            <input 
                              type="text"
                              placeholder="Giá trị..."
                              value={newIfRule.val}
                              onChange={(e) => setNewIfRule({ ...newIfRule, val: e.target.value })}
                              className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                            />
                          )}
                        </div>

                        <button 
                          onClick={() => handleLogicRuleAdd("if")}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-4 py-2 rounded-lg transition-all w-full cursor-pointer"
                        >
                          Thêm dòng NẾU
                        </button>

                        <div className="space-y-1 max-h-[140px] overflow-y-auto border-t border-gray-800 pt-3">
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-bold text-gray-400">QUY TẮC "NẾU" HIỆN TẠI:</span>
                            <div className="flex gap-2 text-[10px]">
                              <label className="flex items-center gap-1 text-gray-300">
                                <input type="radio" checked={ifCombine === "AND"} onChange={() => setIfCombine("AND")} className="scale-75" /> VÀ (AND)
                              </label>
                              <label className="flex items-center gap-1 text-gray-300">
                                <input type="radio" checked={ifCombine === "OR"} onChange={() => setIfCombine("OR")} className="scale-75" /> HOẶC (OR)
                              </label>
                            </div>
                          </div>
                          {ifRules.length === 0 ? (
                            <div className="text-[11px] text-gray-500 italic">Chưa dựng quy tắc...</div>
                          ) : (
                            ifRules.map((rule, idx) => (
                              <div key={idx} className="flex justify-between items-center bg-[#181d29] px-3 py-1.5 rounded-lg border border-gray-800 text-xs text-gray-300">
                                <span>{rule.col} {rule.op} {rule.op !== "trống" && rule.op !== "không trống" ? `'${rule.val}'` : ""}</span>
                                <button onClick={() => setIfRules(ifRules.filter((_, i) => i !== idx))} className="text-red-400 hover:text-red-300 cursor-pointer">X</button>
                              </div>
                            ))
                          )}
                        </div>
                      </div>

                      {/* ĐIỀU KIỆN 2: BƯỚC THÌ PHẢI */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-emerald-500/10 space-y-4">
                        <div className="text-xs font-bold text-emerald-400 uppercase tracking-wider font-mono">BƯỚC 2: THÌ BẮT BUỘC PHẢI (THEN MUST CONDITIONS)</div>
                        <p className="text-[11px] text-gray-400">Nếu thỏa mãn Bước 1, thì dữ liệu bắt buộc PHẢI đạt tất cả ràng buộc sau:</p>
                        
                        <div className="grid grid-cols-3 gap-2">
                          <select 
                            value={newThenRule.col} 
                            onChange={(e) => setNewThenRule({ ...newThenRule, col: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="">Cột</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          
                          <select 
                            value={newThenRule.op} 
                            onChange={(e) => setNewThenRule({ ...newThenRule, op: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="==">bằng (==)</option>
                            <option value="!=">khác (!=)</option>
                            <option value=">">lớn hơn (&gt;)</option>
                            <option value="<">nhỏ hơn (&lt;)</option>
                            <option value=">=">lớn hơn bằng (&gt;=)</option>
                            <option value="<=">nhỏ hơn bằng (&lt;=)</option>
                            <option value="chứa">chứa (string)</option>
                            <option value="không chứa">không chứa (string)</option>
                            <option value="trống">để rỗng (empty)</option>
                            <option value="không trống">có dữ liệu</option>
                          </select>

                          {newThenRule.op !== "trống" && newThenRule.op !== "không trống" && (
                            <input 
                              type="text"
                              placeholder="Giá trị..."
                              value={newThenRule.val}
                              onChange={(e) => setNewThenRule({ ...newThenRule, val: e.target.value })}
                              className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                            />
                          )}
                        </div>

                        <button 
                          onClick={() => handleLogicRuleAdd("then")}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-lg transition-all w-full cursor-pointer"
                        >
                          Thêm dòng THÌ PHẢI
                        </button>

                        <div className="space-y-1 max-h-[140px] overflow-y-auto border-t border-gray-800 pt-3">
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-bold text-gray-400">QUY TẮC "THÌ PHẢI" HIỆN TẠI:</span>
                            <div className="flex gap-2 text-[10px]">
                              <label className="flex items-center gap-1 text-gray-300">
                                <input type="radio" checked={thenCombine === "AND"} onChange={() => setThenCombine("AND")} className="scale-75" /> VÀ (AND)
                              </label>
                              <label className="flex items-center gap-1 text-gray-300">
                                <input type="radio" checked={thenCombine === "OR"} onChange={() => setThenCombine("OR")} className="scale-75" /> HOẶC (OR)
                              </label>
                            </div>
                          </div>
                          {thenRules.length === 0 ? (
                            <div className="text-[11px] text-gray-500 italic">Chưa dựng quy tắc...</div>
                          ) : (
                            thenRules.map((rule, idx) => (
                              <div key={idx} className="flex justify-between items-center bg-[#15241e] px-3 py-1.5 rounded-lg border border-gray-800 text-xs text-gray-300">
                                <span>{rule.col} {rule.op} {rule.op !== "trống" && rule.op !== "không trống" ? `'${rule.val}'` : ""}</span>
                                <button onClick={() => setThenRules(thenRules.filter((_, i) => i !== idx))} className="text-red-400 hover:text-red-300 cursor-pointer">X</button>
                              </div>
                            ))
                          )}
                        </div>
                      </div>

                    </div> 

                    <button 
                      onClick={handleLogicCheck}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all w-full flex items-center justify-center gap-2 cursor-pointer shadow-md"
                    >
                      <CheckSquare className="w-5 h-5 text-purple-300" /> BẮT ĐẦU CHẠY KIỂM TRA LỌC LOGIC ĐA QUY TẮC
                    </button>

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 9. TAB BIỂU ĐỒ TRỰC QUAN DOANH THU THEO NGÀNH CẤP 1 */}
          {activeTab === "bieudotrucquan" && (
            <div className="space-y-6 animate-fade-in">
              <SectorRevenueChart 
                mainData={mainData} 
                columns={columns} 
                mapping={mapping} 
              />
            </div>
          )}

          {/* 10. TAB TRA CỨU DANH MỤC NGÀNH VSIC CHUẨN */}
          {activeTab === "danhmucvsic" && (
            <div className="space-y-6 animate-fade-in">
              <VsicCatalogExplorer />
            </div>
          )}

          {/* 11. TAB ĐỐI CHIẾU MÔ TẢ ĐTV VÀ TÊN NGÀNH CHUẨN */}
          {activeTab === "doichieumota" && (
            <div className="space-y-6 animate-fade-in">
              {mainData.length > 0 ? (
                <DescriptorMatchScanner mainData={mainData} columns={columns} mapping={mapping} />
              ) : (
                <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-8 text-center space-y-3">
                  <div className="bg-amber-500/10 border border-amber-500/20 text-amber-400 p-4 rounded-xl inline-block">
                    ⚠️ Chưa có dữ liệu nguồn chính!
                  </div>
                  <p className="text-xs text-gray-400 max-w-md mx-auto">
                    Vui lòng quay lại tab <strong>Trang Chủ / Nạp Dữ Liệu</strong> để tải tệp Excel dữ liệu của bạn lên trước khi thực hiện quy trình so sánh đối chiếu ngữ nghĩa mô tả ĐTV và tên ngành chuẩn.
                  </p>
                </div>
              )}
            </div>
          )}

        </main>
      </div>

    </div>
  );
}
