import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

// --- INDEXEDDB STORAGE FOR LARGE FILES (40-50MB+) ---
const DB_NAME = "HÊ THỐNG XỬ LÝ SO SÁNH TỔNG HỢP DATA";
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
  userSectorMap?: [string, string][];
  userSectorFileName?: string;
}

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => reject(new Error("Không thể khởi tạo IndexedDB"));
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
  });
}

const CHUNK_SIZE = 5000;

async function clearOldChunks(db: IDBDatabase): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.openKeyCursor();
    request.onsuccess = (event: any) => {
      const cursor = event.target.result;
      if (cursor) {
        const key = String(cursor.key);
        if (key.startsWith("mainData_chunk_") || key.startsWith("rawImportedData_chunk_")) {
          store.delete(key);
        }
        cursor.continue();
      } else {
        resolve();
      }
    };
    request.onerror = () => reject(request.error);
  });
}

async function saveArrayInChunks(db: IDBDatabase, prefix: string, array: any[]): Promise<void> {
  const numChunks = Math.ceil(array.length / CHUNK_SIZE);
  for (let i = 0; i < numChunks; i++) {
    const start = i * CHUNK_SIZE;
    const chunk = array.slice(start, start + CHUNK_SIZE);
    await new Promise<void>((resolve) => setTimeout(resolve, 5));
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.put(chunk, `${prefix}_chunk_${i}`);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error(`Lỗi lưu mảnh ${prefix} ${i}`));
    });
  }
}

async function loadArrayInChunks(db: IDBDatabase, prefix: string, totalLength: number): Promise<any[]> {
  const numChunks = Math.ceil(totalLength / CHUNK_SIZE);
  const result: any[] = [];
  for (let i = 0; i < numChunks; i++) {
    await new Promise<void>((resolve) => setTimeout(resolve, 2));
    const chunk: any[] = await new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(`${prefix}_chunk_${i}`);
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error || new Error(`Lỗi tải mảnh ${prefix} ${i}`));
    });
    result.push(...chunk);
  }
  return result;
}

async function saveAppState(state: AppState, forceSaveData: boolean = false): Promise<void> {
  try {
    const db = await openDB();
    const { mainData, rawImportedData, ...meta } = state;
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const metaData = {
        ...meta,
        mainDataLength: mainData.length,
        rawImportedDataLength: rawImportedData.length,
        isChunked: true
      };
      const request = store.put(metaData, "sessionMeta");
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error("Lỗi lưu metadata"));
    });
    if (forceSaveData) {
      await clearOldChunks(db);
      await saveArrayInChunks(db, "mainData", mainData);
      await saveArrayInChunks(db, "rawImportedData", rawImportedData);
    }
    // Lưu danh mục ngành
    const sectorMapArray = Array.from(state.userSectorMap || []);
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const req1 = store.put(sectorMapArray, "userSectorMap");
      const req2 = store.put(state.userSectorFileName || "", "userSectorFileName");
      Promise.all([req1, req2]).then(() => resolve()).catch(reject);
    });
    await new Promise<void>((resolve) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      store.delete("currentSession");
      resolve();
    });
  } catch (error) {
    console.error("IndexedDB Save Error:", error);
  }
}

async function loadAppState(): Promise<AppState | null> {
  try {
    const db = await openDB();
    const meta: any = await new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get("sessionMeta");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    if (meta && meta.isChunked) {
      const mainData = await loadArrayInChunks(db, "mainData", meta.mainDataLength || 0);
      const rawImportedData = await loadArrayInChunks(db, "rawImportedData", meta.rawImportedDataLength || 0);
      return { ...meta, mainData, rawImportedData };
    }
    const legacy = await new Promise<any>((resolve) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get("currentSession");
      request.onsuccess = () => resolve(request.result);
    });
    return legacy || null;
  } catch (error) {
    console.error("IndexedDB Load Error:", error);
    return null;
  }
}

async function clearAppState(): Promise<void> {
  try {
    const db = await openDB();
    await clearOldChunks(db);
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      store.delete("sessionMeta");
      store.delete("userSectorMap");
      store.delete("userSectorFileName");
      resolve();
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
  Search,
  Plus,
  Trash2,
  FileCheck,
  Lock,
  KeyRound,
  LogOut
} from "lucide-react";

// Import các hàm xử lý từ vsic.ts (đã được sửa động)
import { 
  normalizeSectorCode, 
  getSectorHierarchy, 
  smartSuggestSectorByDescription,
  getSectorLevel,
  getParentSectorCode,
  lookupSectorNameWithFallback,
  isSummaryRow,
  clearAllSectorsInVSIC,
  loadSectorsIntoVSIC
} from "./data/vsic";

import SectorRevenueChart from "./components/sectorRevenueChart";
import VsicCatalogExplorer from "./components/vsicCatalogExplorer";
import DescriptorMatchScanner from "./components/descriptorMatchScanner";
import { BeautifulReportTable } from "./components/BeautifulReportTable";

// Component xem trước dữ liệu (dùng chung cho các tab)
const DataPreviewTable: React.FC<{ data: any[]; columns: string[] }> = ({ data, columns }) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [viewPage, setViewPage] = useState(1);
  const pageSize = 50;
  const filteredData = useMemo(() => {
    if (!searchTerm) return data;
    const term = searchTerm.toLowerCase();
    return data.filter(row => Object.values(row).some(val => String(val).toLowerCase().includes(term)));
  }, [data, searchTerm]);
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  const paginatedData = filteredData.slice((viewPage - 1) * pageSize, viewPage * pageSize);
  const handleExport = () => {
    if (data.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "KetQua");
    XLSX.writeFile(wb, `KetQua_XuLy.xlsx`);
  };
  if (data.length === 0) return null;
  return (
    <div className="mt-6 bg-[#1f2937] border border-[#374151] rounded-2xl overflow-hidden shadow-sm space-y-4 p-4">
      <div className="flex flex-col sm:flex-row items-center justify-between gap-4 border-b border-[#374151] pb-4">
        <div className="relative w-full sm:max-w-xs">
          <Search className="absolute left-3 top-2.5 h-4 w-4 text-gray-400" />
          <input type="text" placeholder="Tìm kiếm trong kết quả..." value={searchTerm} onChange={(e) => { setSearchTerm(e.target.value); setViewPage(1); }} className="w-full bg-[#111827] border border-[#374151] rounded-xl pl-9 pr-4 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500" />
        </div>
        <button onClick={handleExport} className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-xl flex items-center gap-1.5"><Download className="w-4 h-4" /> Xuất Excel</button>
      </div>
      <div className="overflow-x-auto max-h-[500px] relative">
        <table className="w-full text-left text-xs border-collapse">
          <thead className="bg-[#111827] text-gray-400 border-b border-gray-800 font-mono sticky top-0 z-10">
            <tr>{columns.map(col => <th key={col} className="p-3 font-semibold text-center whitespace-nowrap min-w-[120px]">{col}</th>)}</tr>
          </thead>
          <tbody>
            {paginatedData.map((row, idx) => (
              <tr key={idx} className="border-b border-gray-800/40 hover:bg-gray-800/50">
                {columns.map(col => <td key={col} className="p-3 truncate max-w-[220px] text-center text-gray-300" title={String(row[col])}>{row[col] === undefined ? "" : String(row[col])}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex items-center justify-between pt-2 text-xs">
        <span className="text-gray-400">Trang <strong className="text-white">{viewPage}</strong> / {totalPages}</span>
        <div className="flex gap-2">
          <button disabled={viewPage === 1} onClick={() => setViewPage(prev => Math.max(1, prev - 1))} className="px-3 py-1.5 rounded-lg border border-gray-700 bg-[#111827] text-gray-300 disabled:text-gray-600 disabled:cursor-not-allowed">Trước</button>
          <button disabled={viewPage === totalPages} onClick={() => setViewPage(prev => Math.min(totalPages, prev + 1))} className="px-3 py-1.5 rounded-lg border border-gray-700 bg-[#111827] text-gray-300 disabled:text-gray-600 disabled:cursor-not-allowed">Sau</button>
        </div>
      </div>
    </div>
  );
};

interface ColumnMapping {
  mota: string;
  manganh: string;
  xa: string;
  doanhthu: string;
  laodong: string;
  idCol: string;
}
interface LogicRule { col: string; op: string; val: string; }

export default function App() {
  const [activeTab, setActiveTab] = useState<string>("trangchu");
  const [loading, setLoading] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string>("");

  // Mật khẩu
  const [appPassword, setAppPassword] = useState<string>(() => localStorage.getItem("vsic_app_password") || "admin123");
  const [isAuthorized, setIsAuthorized] = useState<boolean>(() => localStorage.getItem("vsic_app_authorized") === "true");
  const [typedPassword, setTypedPassword] = useState<string>("");
  const [passwordError, setPasswordError] = useState<string>("");
  const [showPasswordChangeModal, setShowPasswordChangeModal] = useState<boolean>(false);
  const [newPasswordVal, setNewPasswordVal] = useState<string>("");

  // State dữ liệu chính
  const [mainData, setMainData] = useState<any[]>([]);
  const [rawImportedData, setRawImportedData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [quickReportResultRows, setQuickReportResultRows] = useState<any[]>([]);
  const [quickReportResultCols, setQuickReportResultCols] = useState<string[]>([]);
  const [quickReportLevel, setQuickReportLevel] = useState<number>(1);
  const [mapping, setMapping] = useState<ColumnMapping>({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
  const [customColConfigs, setCustomColConfigs] = useState<{ originalName: string; use: boolean; newName: string; role: "" | "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" }[]>([]);

  // State danh mục người dùng
  const [userSectorMap, setUserSectorMap] = useState<Map<string, string>>(new Map());
  const [userSectorFileName, setUserSectorFileName] = useState<string>("");

  // Các state khác
  const [calcColName, setCalcColName] = useState<string>("");
  const [calcCol1, setCalcCol1] = useState<string>("");
  const [calcCol2, setCalcCol2] = useState<string>("");
  const [calcOperator, setCalcOperator] = useState<"+" | "-" | "*" | "/" | "concat">("+");
  const [calcType, setCalcType] = useState<"column" | "constant">("column");
  const [calcConstant, setCalcConstant] = useState<string>("");
  const [calcRounding, setCalcRounding] = useState<"none" | "int" | "1dec" | "2dec">("none");
  const [selectedTargetKey, setSelectedTargetKey] = useState<keyof ColumnMapping>("mota");
  const [reportType, setReportType] = useState<"flat" | "pivot">("pivot");
  const [isConfigExpanded, setIsConfigExpanded] = useState<boolean>(true);
  const [oldData, setOldData] = useState<any[]>([]);
  const [oldFileName, setOldFileName] = useState<string>("");
  const [newData, setNewData] = useState<any[]>([]);
  const [newFileName, setNewFileName] = useState<string>("");
  const [diffKey, setDiffKey] = useState<string>("");
  const [leftData, setLeftData] = useState<any[]>([]);
  const [leftFileName, setLeftFileName] = useState<string>("");
  const [rightData, setRightData] = useState<any[]>([]);
  const [rightFileName, setRightFileName] = useState<string>("");
  const [leftKey, setLeftKey] = useState<string>("");
  const [rightKey, setRightKey] = useState<string>("");
  const [splitCol, setSplitCol] = useState<string>("");
  const [detectedWorkbook, setDetectedWorkbook] = useState<any | null>(null);
  const [detectedSheets, setDetectedSheets] = useState<string[]>([]);
  const [selectedSheetsToMerge, setSelectedSheetsToMerge] = useState<string[]>([]);
  const [sheetMergeCommonKey, setSheetMergeCommonKey] = useState<string>("");
  const [crossReportData, setCrossReportData] = useState<any[]>([]);
  const [crossReportCols, setCrossReportCols] = useState<string[]>([]);
  const [crossReportManganhCol, setCrossReportManganhCol] = useState<string>("");
  const [crossReportXaCol, setCrossReportXaCol] = useState<string>("");
  const [crossReportDoanhThuCol, setCrossReportDoanhThuCol] = useState<string>("");
  const [crossReportLaoDongCol, setCrossReportLaoDongCol] = useState<string>("");
  const [crossReportLevel, setCrossReportLevel] = useState<number>(2);
  const [groupByCols, setGroupByCols] = useState<string[]>([]);
  const [aggRules, setAggRules] = useState<{ col: string; op: string }[]>([]);
  const [newAggCol, setNewAggCol] = useState<string>("");
  const [newAggOp, setNewAggOp] = useState<string>("sum");
  const [ifRules, setIfRules] = useState<LogicRule[]>([]);
  const [thenRules, setThenRules] = useState<LogicRule[]>([]);
  const [ifCombine, setIfCombine] = useState<"AND" | "OR">("AND");
  const [thenCombine, setThenCombine] = useState<"AND" | "OR">("AND");
  const [newIfRule, setNewIfRule] = useState<LogicRule>({ col: "", op: "==", val: "" });
  const [newThenRule, setNewThenRule] = useState<LogicRule>({ col: "", op: "==", val: "" });
  const [t2IndustryCol, setT2IndustryCol] = useState<string>("");
  const [t2MetricCols, setT2MetricCols] = useState<string[]>([]);
  const [t2AggMethod, setT2AggMethod] = useState<"sum" | "avg">("sum");
  const [t2ReportData, setT2ReportData] = useState<any[]>([]);
  const [t2ReportCols, setT2ReportCols] = useState<string[]>([]);
  const [t2ReportLevel, setT2ReportLevel] = useState<number>(2);
  const [quickReportManganhCol, setQuickReportManganhCol] = useState<string>("");
  const [quickReportXaCol, setQuickReportXaCol] = useState<string>("");
  const [quickReportDoanhThuCol, setQuickReportDoanhThuCol] = useState<string>("");
  const [quickReportLaoDongCol, setQuickReportLaoDongCol] = useState<string>("");
  const [pivotManganhCol, setPivotManganhCol] = useState<string>("");
  const [stdIndustryCol, setStdIndustryCol] = useState<string>("");
  const [stdDescriptionCol, setStdDescriptionCol] = useState<string>("");
  const [stdReportAnomalies, setStdReportAnomalies] = useState<any[]>([]);
  const [stdMatchStats, setStdMatchStats] = useState<{ total: number; valid: number; invalid: number; conflicts: number }>({ total: 0, valid: 0, invalid: 0, conflicts: 0 });
  const [crossCompareColA, setCrossCompareColA] = useState<string>("");
  const [crossCompareColB, setCrossCompareColB] = useState<string>("");
  const [crossCompareRule, setCrossCompareRule] = useState<string>("normalize");
  const [crossCompareAnomalies, setCrossCompareAnomalies] = useState<any[]>([]);
  const [crossCompareStats, setCrossCompareStats] = useState<{ total: number; matchCount: number; mismatchCount: number }>({ total: 0, matchCount: 0, mismatchCount: 0 });
  const [viewPage, setViewPage] = useState<number>(1);
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [inconsistenciesTab, setInconsistenciesTab] = useState<"desc" | "code">("desc");
  const pageSize = 50;

  // Khôi phục dữ liệu từ IndexedDB và localStorage
  useEffect(() => {
    async function restoreSession() {
      setLoading(true);
      setStatusMessage("Đang khôi phục phiên làm việc...");
      try {
        const saved = await loadAppState();
        if (saved && saved.mainData && saved.mainData.length > 0) {
          setMainData(saved.mainData);
          setRawImportedData(saved.rawImportedData || saved.mainData);
          setColumns(saved.columns);
          setFileName(saved.fileName);
          if (saved.mapping) setMapping(saved.mapping);
          if (saved.customColConfigs) setCustomColConfigs(saved.customColConfigs);
          if (saved.userSectorMap) {
            const recovered = new Map(saved.userSectorMap);
            setUserSectorMap(recovered);
            clearAllSectorsInVSIC();
            loadSectorsIntoVSIC(Object.fromEntries(recovered));
          }
          if (saved.userSectorFileName) setUserSectorFileName(saved.userSectorFileName);
        }
        // Khôi phục danh mục từ localStorage nếu chưa có
        const localCatalog = localStorage.getItem("custom_vsic_data");
        if (localCatalog && !userSectorMap.size) {
          const obj = JSON.parse(localCatalog);
          const recovered = new Map(Object.entries(obj));
          setUserSectorMap(recovered);
          clearAllSectorsInVSIC();
          loadSectorsIntoVSIC(obj);
          const localName = localStorage.getItem("custom_vsic_filename");
          if (localName) setUserSectorFileName(localName);
        }
      } catch (err) { console.error(err); }
      finally { setLoading(false); }
    }
    restoreSession();
  }, []);

  // Hàm upload danh mục ngành
  const handleUploadUserSectorCatalog = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setStatusMessage(`Đang nạp danh mục ngành: ${file.name}...`);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        let rows: any[] = [];
        const data = evt.target?.result;
        const isCSV = file.name.toLowerCase().endsWith('.csv');
        if (isCSV) {
          const text = data as string;
          rows = parseCSV(text);
        } else {
          const wb = XLSX.read(data, { type: 'array', dense: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          rows = XLSX.utils.sheet_to_json(ws);
        }
        if (rows.length === 0) throw new Error("File danh mục trống");
        const firstRow = rows[0];
        let codeCol = Object.keys(firstRow).find(k => /mã|ma|code|Mã/i.test(k)) || Object.keys(firstRow)[0];
        let nameCol = Object.keys(firstRow).find(k => /tên|ten|name|Tên/i.test(k)) || Object.keys(firstRow)[1] || codeCol;
        const newMap = new Map<string, string>();
        rows.forEach(row => {
          const code = String(row[codeCol] || "").trim();
          const name = String(row[nameCol] || "").trim();
          if (code && name) newMap.set(code, name);
        });
        if (newMap.size === 0) throw new Error("Không tìm thấy cặp mã - tên hợp lệ");
        clearAllSectorsInVSIC();
        const plainObj = Object.fromEntries(newMap);
        loadSectorsIntoVSIC(plainObj);
        localStorage.setItem("custom_vsic_data", JSON.stringify(plainObj));
        localStorage.setItem("custom_vsic_is_pure", "false");
        localStorage.setItem("custom_vsic_filename", file.name);
        setUserSectorMap(newMap);
        setUserSectorFileName(file.name);
        setStatusMessage(`✅ Đã nạp thành công ${newMap.size} mã ngành từ ${file.name}`);
        // Tự động lưu vào IndexedDB
        saveAppState({
          mainData, rawImportedData, columns, fileName, mapping, customColConfigs,
          userSectorMap: Array.from(newMap), userSectorFileName: file.name
        }, false);
      } catch (err: any) {
        alert("Lỗi khi đọc file danh mục: " + err.message);
      } finally {
        setLoading(false);
      }
    };
    if (file.name.toLowerCase().endsWith('.csv')) reader.readAsText(file, "UTF-8");
    else reader.readAsArrayBuffer(file);
  };

  const handleClearUserSectors = () => {
    clearAllSectorsInVSIC();
    localStorage.removeItem("custom_vsic_data");
    localStorage.removeItem("custom_vsic_is_pure");
    localStorage.removeItem("custom_vsic_filename");
    setUserSectorMap(new Map());
    setUserSectorFileName("");
    setStatusMessage("🗑️ Đã xóa toàn bộ danh mục ngành tùy chỉnh.");
    alert("Đã xóa sạch danh mục ngành đã nạp.");
  };

  // ==================== CÁC HÀM XỬ LÝ CHÍNH ====================
  
  // Hàm parse CSV
  const parseCSV = (rawText: string): any[] => {
    let text = rawText;
    if (text.charCodeAt(0) === 0xFEFF) text = text.substring(1);
    const firstLineEnd = text.indexOf('\n');
    const firstLine = firstLineEnd === -1 ? text : text.substring(0, firstLineEnd);
    let delimiter = ',';
    const commaCount = (firstLine.match(/,/g) || []).length;
    const semicolonCount = (firstLine.match(/;/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    if (semicolonCount > commaCount && semicolonCount > tabCount) delimiter = ';';
    else if (tabCount > commaCount && tabCount > semicolonCount) delimiter = '\t';
    const length = text.length;
    const rows: string[][] = [];
    let currentRow: string[] = [];
    let i = 0, inQuotes = false, start = 0;
    while (i < length) {
      const char = text[i];
      if (char === '"') {
        if (!inQuotes) { inQuotes = true; start = i + 1; }
        else if (i + 1 < length && text[i + 1] === '"') { i++; }
        else { inQuotes = false; }
      } else if (!inQuotes) {
        if (char === delimiter) {
          let cell = text.substring(start, i);
          if (text[i-1] === '"' && text[start-1] === '"') cell = cell.substring(0, cell.length-1);
          if (cell.includes('""')) cell = cell.replace(/""/g, '"');
          currentRow.push(cell.trim());
          start = i+1;
        } else if (char === '\n' || char === '\r') {
          let cell = text.substring(start, i);
          if (text[i-1] === '"' && text[start-1] === '"') cell = cell.substring(0, cell.length-1);
          if (cell.includes('""')) cell = cell.replace(/""/g, '"');
          currentRow.push(cell.trim());
          if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) rows.push(currentRow);
          currentRow = [];
          if (char === '\r' && i+1 < length && text[i+1] === '\n') i++;
          start = i+1;
        }
      }
      i++;
    }
    if (start < length) {
      let cell = text.substring(start, length);
      if (text[length-1] === '"' && text[start-1] === '"') cell = cell.substring(0, cell.length-1);
      if (cell.includes('""')) cell = cell.replace(/""/g, '"');
      currentRow.push(cell.trim());
    }
    if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) rows.push(currentRow);
    if (rows.length === 0) return [];
    const headers = rows[0];
    const data: any[] = [];
    for (let idx = 1; idx < rows.length; idx++) {
      const r = rows[idx];
      const obj: any = {};
      let hasData = false;
      for (let c = 0; c < headers.length; c++) {
        const headerName = headers[c] || `Cột ${c+1}`;
        const val = r[c] !== undefined ? r[c] : "";
        obj[headerName] = val;
        if (val !== "") hasData = true;
      }
      if (hasData) data.push(obj);
    }
    return data;
  };

  const chunkProcess = async <T, R>(array: T[], size: number, processFn: (item: T, index: number) => R, onProgress?: (percent: number) => void): Promise<R[]> => {
    const result: R[] = [];
    const len = array.length;
    for (let i = 0; i < len; i += size) {
      const chunk = array.slice(i, i+size);
      for (let j = 0; j < chunk.length; j++) result.push(processFn(chunk[j], i+j));
      if (onProgress) onProgress(Math.min(100, Math.round((i/len)*100)));
      await new Promise(r => setTimeout(r, 0));
    }
    return result;
  };

  const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
  const cleanNumberForSummary = (val: any): number => {
    if (val === null || val === undefined) return 0;
    if (typeof val === "number") return val;
    let str = String(val).trim();
    if (!str) return 0;
    if (str.includes(",") && str.includes(".")) {
      const lastComma = str.lastIndexOf(",");
      const lastDot = str.lastIndexOf(".");
      if (lastComma > lastDot) str = str.replace(/\./g, "").replace(/,/g, ".");
      else str = str.replace(/,/g, "");
    } else if (str.includes(",")) {
      const parts = str.split(",");
      if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) str = str.replace(/,/g, "");
      else str = str.replace(/,/g, ".");
    } else if (str.includes(" ")) str = str.replace(/\s/g, "");
    const parsed = parseFloat(str.replace(/[^0-9.\-]/g, ""));
    return isNaN(parsed) ? 0 : parsed;
  };

  const autoSaveSession = async (customMainData?: any[], customRawData?: any[], customCols?: string[], customFileName?: string, customMapping?: ColumnMapping, customConfigs?: any[]) => {
    try {
      await saveAppState({
        mainData: customMainData !== undefined ? customMainData : mainData,
        rawImportedData: customRawData !== undefined ? customRawData : rawImportedData,
        columns: customCols !== undefined ? customCols : columns,
        fileName: customFileName !== undefined ? customFileName : fileName,
        mapping: customMapping !== undefined ? customMapping : mapping,
        customColConfigs: customConfigs !== undefined ? customConfigs : customColConfigs,
        userSectorMap: Array.from(userSectorMap),
        userSectorFileName
      }, true);
    } catch (err) { console.error(err); }
  };

  // Hàm tải file chính
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
          const data = parseCSV(text);
          if (data.length === 0) throw new Error("Tệp trống");
          const cols = Object.keys(data[0]);
          if (type === "main") {
            setRawImportedData(data); setMainData(data); setColumns(cols); setFileName(file.name);
            setQuickReportManganhCol(""); setStdIndustryCol(""); setCrossCompareColA(""); setStdDescriptionCol("");
            setQuickReportXaCol(""); setQuickReportDoanhThuCol(""); setQuickReportLaoDongCol("");
            const autoMap: ColumnMapping = { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" };
            setMapping(autoMap);
            const initConfigs = cols.map(c => ({ originalName: c, use: true, newName: c, role: "" as any }));
            setCustomColConfigs(initConfigs);
            setActiveTab("xemdulieu");
            autoSaveSession(data, data, cols, file.name, autoMap, initConfigs);
          } else if (type === "old") { setOldData(data); setOldFileName(file.name); }
          else if (type === "new") { setNewData(data); setNewFileName(file.name); }
          else if (type === "left") { setLeftData(data); setLeftFileName(file.name); }
          else if (type === "right") { setRightData(data); setRightFileName(file.name); }
          setStatusMessage(`Đã tải ${data.length} dòng.`);
        } catch(err: any) { alert("Lỗi CSV: "+err.message); }
        finally { setLoading(false); }
      };
      reader.readAsText(file, "UTF-8");
    } else {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          const wb = XLSX.read(arrayBuffer, { type: "array", dense: true, cellFormula: false, cellHTML: false, cellStyles: false });
          if (type === "main") {
            setDetectedWorkbook(wb);
            setDetectedSheets(wb.SheetNames);
            if (wb.SheetNames.length > 1) { setSelectedSheetsToMerge(wb.SheetNames); setSheetMergeCommonKey(""); }
            else { setSelectedSheetsToMerge([]); setSheetMergeCommonKey(""); }
          }
          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];
          if (data.length === 0) throw new Error("Tệp trống");
          const cols = Object.keys(data[0]);
          if (type === "main") {
            setRawImportedData(data); setMainData(data); setColumns(cols); setFileName(file.name);
            setQuickReportManganhCol(""); setStdIndustryCol(""); setCrossCompareColA(""); setStdDescriptionCol("");
            setQuickReportXaCol(""); setQuickReportDoanhThuCol(""); setQuickReportLaoDongCol("");
            const autoMap: ColumnMapping = { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" };
            setMapping(autoMap);
            const initConfigs = cols.map(c => ({ originalName: c, use: true, newName: c, role: "" as any }));
            setCustomColConfigs(initConfigs);
            setActiveTab("xemdulieu");
            autoSaveSession(data, data, cols, file.name, autoMap, initConfigs);
          } else if (type === "old") { setOldData(data); setOldFileName(file.name); }
          else if (type === "new") { setNewData(data); setNewFileName(file.name); }
          else if (type === "left") { setLeftData(data); setLeftFileName(file.name); }
          else if (type === "right") { setRightData(data); setRightFileName(file.name); }
          setStatusMessage(`Đã tải ${data.length} dòng.`);
        } catch(err: any) { alert("Lỗi Excel: "+err.message); }
        finally { setLoading(false); }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Ghép nối dữ liệu
  const handleMerge = async () => {
    if (leftData.length === 0 || rightData.length === 0) { alert("Vui lòng tải đủ cả 2 bảng!"); return; }
    if (!leftKey || !rightKey) { alert("Vui lòng chọn cột khóa!"); return; }
    setLoading(true); setProgress(0); setStatusMessage("Bắt đầu ghép nối...");
    await sleep(200);
    const rightMap = new Map();
    for (const row of rightData) { const kv = String(row[rightKey] || "").trim(); if (kv) rightMap.set(kv, row); }
    const mergedResults: any[] = [];
    const rightCols = Object.keys(rightData[0] || {}).filter(c => c !== rightKey);
    for (const leftRow of leftData) {
      const matchKey = String(leftRow[leftKey] || "").trim();
      const matchedRight = rightMap.get(matchKey);
      const mergedRow = { ...leftRow };
      rightCols.forEach(rc => { const finalColName = Object.keys(leftRow).includes(rc) ? `${rc}_Phai` : rc; mergedRow[finalColName] = matchedRight ? matchedRight[rc] : ""; });
      mergedResults.push(mergedRow);
    }
    const mergedCols = Object.keys(mergedResults[0] || {});
    setMainData(mergedResults);
    setRawImportedData(mergedResults);
    setColumns(mergedCols);
    setFileName(`GhepNoi_${leftFileName}_vs_${rightFileName}.xlsx`);
    setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
    const initMergedConfigs = mergedCols.map(c => ({ originalName: c, use: true, newName: c, role: "" as any }));
    setCustomColConfigs(initMergedConfigs);
    setProgress(100); setStatusMessage(`Ghép nối thành công! ${mergedResults.length} dòng.`);
    await sleep(400);
    setLoading(false);
  };

  // So sánh Cũ - Mới
  const handleCompare = async () => {
    if (oldData.length === 0 || newData.length === 0) { alert("Vui lòng tải đầy đủ tệp Cũ và Mới!"); return; }
    if (!diffKey) { alert("Vui lòng chọn Cột Khóa!"); return; }
    setLoading(true); setProgress(0); setStatusMessage("Khởi động so sánh...");
    await sleep(200);
    const oldMap = new Map(); oldData.forEach(row => { const k = String(row[diffKey] || "").trim(); if(k) oldMap.set(k, row); });
    const newMap = new Map(); newData.forEach(row => { const k = String(row[diffKey] || "").trim(); if(k) newMap.set(k, row); });
    const allKeys = Array.from(new Set([...oldMap.keys(), ...newMap.keys()]));
    const firstOldRow = oldData.find(r => r && typeof r === 'object') || {};
    const firstNewRow = newData.find(r => r && typeof r === 'object') || {};
    const oldCols = Object.keys(firstOldRow);
    const newCols = Object.keys(firstNewRow);
    const unionCols = Array.from(new Set([...oldCols, ...newCols])).filter(c => c !== diffKey);
    const resultRows: any[] = [];
    for (const key of allKeys) {
      const oldRow = oldMap.get(key);
      const newRow = newMap.get(key);
      const combined: any = { [diffKey]: key };
      if (oldRow && !newRow) {
        unionCols.forEach(col => { combined[`${col}_Cu`] = oldRow[col] || ""; combined[`${col}_Moi`] = ""; });
        combined["TrangThai_SoSanh"] = "❌ Đã xóa";
      } else if (!oldRow && newRow) {
        unionCols.forEach(col => { combined[`${col}_Cu`] = ""; combined[`${col}_Moi`] = newRow[col] || ""; });
        combined["TrangThai_SoSanh"] = "✅ Mới thêm";
      } else {
        const changedCols: string[] = [];
        unionCols.forEach(col => {
          const valCu = String(oldRow[col] !== undefined ? oldRow[col] : "").trim();
          const valMoi = String(newRow[col] !== undefined ? newRow[col] : "").trim();
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = newRow[col] || "";
          if (valCu !== valMoi) changedCols.push(col);
        });
        combined["TrangThai_SoSanh"] = changedCols.length ? `⚠️ Thay đổi: [${changedCols.join(", ")}]` : "💡 Không đổi";
      }
      resultRows.push(combined);
    }
    const compareCols = Object.keys(resultRows[0] || {});
    setMainData(resultRows);
    setRawImportedData(resultRows);
    setColumns(compareCols);
    setFileName(`SoSanhDiff_${oldFileName}_vs_${newFileName}.xlsx`);
    setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
    const initCompareConfigs = compareCols.map(c => ({ originalName: c, use: true, newName: c, role: "" as any }));
    setCustomColConfigs(initCompareConfigs);
    setProgress(100); setStatusMessage(`So sánh thành công! ${resultRows.length} khóa.`);
    await sleep(400);
    setLoading(false);
  };

  // Báo cáo nhanh theo cấp ngành và xã
  const handleQuickReport = async (level: number) => {
    if (mainData.length === 0) { alert("Vui lòng nạp dữ liệu chính trước."); return; }
    const targetManganh = quickReportManganhCol || mapping.manganh;
    const targetXa = quickReportXaCol || mapping.xa;
    const targetDoanhThu = quickReportDoanhThuCol || mapping.doanhthu;
    const targetLaoDong = quickReportLaoDongCol || mapping.laodong;
    if (!targetManganh || !targetXa) { alert("Vui lòng chỉ định cột Mã ngành và Xã."); return; }
    if (userSectorMap.size === 0) { alert("Vui lòng tải danh mục ngành trước khi chạy báo cáo!"); return; }
    setLoading(true); setProgress(0); setStatusMessage(`Đang tạo báo cáo nhanh cấp ${level}...`);
    await sleep(200);
    try {
      const processedData = await chunkProcess(mainData, 5000, (row) => {
        const mng = normalizeSectorCode(row[targetManganh]);
        let tenNganhLabel = "";
        if (level === 2) {
          const sec2Code = mng ? mng.slice(0,2) : "";
          const sec2Name = userSectorMap.get(sec2Code) || "";
          tenNganhLabel = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định";
        } else {
          let sec1Code = "";
          if (mng) { if (/^[A-Z]$/.test(mng)) sec1Code = mng.toUpperCase(); else sec1Code = getParentSectorCode(mng) || ""; }
          const sec1Name = userSectorMap.get(sec1Code) || "";
          tenNganhLabel = sec1Code ? `${sec1Code} - ${sec1Name}` : "Chưa xác định";
        }
        return { ...row, _temNganhCap: tenNganhLabel, _tempXa: String(row[targetXa] || "Khác").trim() };
      });
      let finalReportRows: any[] = [];
      if (reportType === "pivot") {
        const communes = Array.from(new Set(processedData.map(r => r._tempXa))).sort();
        const sectorLabels = Array.from(new Set(processedData.map(r => r._temNganhCap))).sort();
        communes.forEach(commune => {
          const communeObj: any = { "Địa_Bàn_Xã": commune };
          sectorLabels.forEach(sector => {
            const matchedRows = processedData.filter(r => r._tempXa === commune && r._temNganhCap === sector);
            let sumDoanhThu = 0, sumLaoDong = 0;
            matchedRows.forEach(r => {
              if (targetDoanhThu) { const v = parseFloat(String(r[targetDoanhThu]).replace(/[^0-9.\-]/g, "")); if (!isNaN(v)) sumDoanhThu += v; }
              if (targetLaoDong) { const v = parseFloat(String(r[targetLaoDong]).replace(/[^0-9.\-]/g, "")); if (!isNaN(v)) sumLaoDong += v; }
            });
            communeObj[`${sector} - Tổng Doanh Thu`] = Math.round(sumDoanhThu * 100) / 100;
            communeObj[`${sector} - Tổng Lao Động`] = Math.round(sumLaoDong);
          });
          finalReportRows.push(communeObj);
        });
      } else {
        const groups = new Map<string, any[]>();
        processedData.forEach(row => { const key = JSON.stringify({ Ngành: row._temNganhCap, Xã: row._tempXa }); if (!groups.has(key)) groups.set(key, []); groups.get(key)!.push(row); });
        groups.forEach((rowsObj, keyStr) => {
          const dims = JSON.parse(keyStr);
          let sumDoanhThu = 0, sumLaoDong = 0;
          rowsObj.forEach(r => {
            if (targetDoanhThu) { const v = parseFloat(String(r[targetDoanhThu]).replace(/[^0-9.\-]/g, "")); if (!isNaN(v)) sumDoanhThu += v; }
            if (targetLaoDong) { const v = parseFloat(String(r[targetLaoDong]).replace(/[^0-9.\-]/g, "")); if (!isNaN(v)) sumLaoDong += v; }
          });
          finalReportRows.push({ [`Ngành_Cấp_${level}`]: dims.Ngành, "Địa_Bàn_Xã": dims.Xã, "Số_Lượng_Doanh_Nghiệp": rowsObj.length, "Tổng_Doanh_Thu_Tích_Lũy": Math.round(sumDoanhThu*100)/100, "Tổng_Lao_Động_Hợp_Lực": Math.round(sumLaoDong) });
        });
      }
      setQuickReportResultRows(finalReportRows);
      setQuickReportResultCols(Object.keys(finalReportRows[0] || {}));
      setQuickReportLevel(level);
      setMapping(prev => ({ ...prev, manganh: targetManganh, xa: targetXa, doanhthu: targetDoanhThu || prev.doanhthu, laodong: targetLaoDong || prev.laodong }));
      setProgress(100); setStatusMessage(`Báo cáo nhanh cấp ${level} thành công!`);
      await sleep(350); setLoading(false);
    } catch(err: any) { alert("Lỗi: "+err.message); setLoading(false); }
  };

  const handleExportQuickReport = () => {
    if (quickReportResultRows.length === 0) { alert("Chưa có dữ liệu báo cáo!"); return; }
    const ws = XLSX.utils.json_to_sheet(quickReportResultRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `BaoCao_NganhCap${quickReportLevel}`);
    XLSX.writeFile(wb, `BaoCao_TongHop_NganhCap${quickReportLevel}_Va_Xa_${reportType}.xlsx`);
  };

  // Xuất Excel tổng hợp
  const handleExportExcel = () => {
    const exportData = searchTerm ? filteredData : mainData;
    if (exportData.length === 0) { alert("Không có dữ liệu!"); return; }
    setLoading(true); setStatusMessage("Đang tạo file Excel...");
    setTimeout(() => {
      try {
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Tinh_Toan");
        if (userSectorMap.size > 0) {
          const vsicRows = Array.from(userSectorMap.entries()).map(([code, name]) => ({ "Mã VSIC": code, "Tên Phân Cấp Ngành": name, "Phân Cấp": `Cấp ${getSectorLevel(code)}` }));
          const wsVsic = XLSX.utils.json_to_sheet(vsicRows);
          XLSX.utils.book_append_sheet(wb, wsVsic, "Danh_Muc_Nganh_Chuan");
        }
        let outName = fileName || "Ket_Qua_Bao_Cao.xlsx";
        if (searchTerm) { const safeSuffix = `_Loc_${searchTerm.trim().slice(0,15).replace(/[^a-zA-Z0-9À-ỹ]/g,"_")}`; outName = outName.replace(/\.xlsx$/i, safeSuffix+".xlsx"); }
        XLSX.writeFile(wb, outName);
        setStatusMessage("Xuất Excel thành công!");
      } catch(e:any) { alert("Lỗi xuất Excel: "+e.message); }
      finally { setLoading(false); }
    }, 200);
  };

  // Tách file hàng loạt
  const handleSplitData = async () => {
    if (mainData.length === 0) { alert("Không có dữ liệu!"); return; }
    if (!splitCol) { alert("Chọn cột để tách!"); return; }
    setLoading(true);
    setStatusMessage("Đang tách file...");
    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const val = String(row[splitCol] || "Rong").trim();
      const safeVal = val.replace(/[^a-zA-Z0-9_\-À-ỹ\s]/g, "");
      if (!groups.has(safeVal)) groups.set(safeVal, []);
      groups.get(safeVal)!.push(row);
    });
    const zip = new JSZip();
    for (let [key, rows] of groups.entries()) {
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
      const buf = new ArrayBuffer(wbout.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < wbout.length; i++) view[i] = wbout.charCodeAt(i) & 0xFF;
      zip.file(`Tach_${key}.xlsx`, buf);
    }
    const content = await zip.generateAsync({ type: "blob" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(content);
    link.download = `TachFile_${splitCol}.zip`;
    link.click();
    setLoading(false);
    setStatusMessage("Đã tách và tải về ZIP.");
  };

  // Chuẩn hóa VSIC (thêm cột tên ngành chuẩn)
  const handleStandardizeSectorsAndMatch = async () => {
    if (mainData.length === 0) { alert("Không có dữ liệu!"); return; }
    if (!stdIndustryCol) { alert("Vui lòng chọn cột Mã ngành!"); return; }
    if (userSectorMap.size === 0) { alert("Vui lòng tải danh mục ngành trước!"); return; }
    setLoading(true); setProgress(0); setStatusMessage("Đang chuẩn hóa mã ngành...");
    try {
      let validCount = 0, invalidCount = 0, conflictCount = 0;
      const updatedRows = await chunkProcess(mainData, 5000, (row, idx) => {
        const rawCode = row[stdIndustryCol];
        const clean = normalizeSectorCode(rawCode);
        const lookup = lookupSectorNameWithFallback(clean);
        const stdName = lookup.name;
        const isValid = lookup.exactMatched;
        if (isValid) validCount++; else invalidCount++;
        let auditStatus = isValid ? "✅ Đạt chuẩn" : "❌ Không có trong danh mục";
        const newRow: any = {};
        Object.keys(row).forEach(k => {
          newRow[k] = row[k];
          if (k === stdIndustryCol) {
            newRow["Tên Ngành Chuẩn VSIC"] = stdName || "(Không tìm thấy)";
            newRow["Trạng Thái Đối Chiếu"] = auditStatus;
          }
        });
        if (!newRow["Tên Ngành Chuẩn VSIC"]) {
          newRow["Tên Ngành Chuẩn VSIC"] = stdName || "(Không tìm thấy)";
          newRow["Trạng Thái Đối Chiếu"] = auditStatus;
        }
        return newRow;
      });
      const newCols = Object.keys(updatedRows[0] || {});
      setMainData(updatedRows);
      setColumns(newCols);
      setStdMatchStats({ total: updatedRows.length, valid: validCount, invalid: invalidCount, conflicts: 0 });
      autoSaveSession(updatedRows, rawImportedData, newCols, fileName, mapping, customColConfigs);
      setStatusMessage("Chuẩn hóa thành công!");
    } catch(e: any) { alert(e.message); }
    finally { setLoading(false); }
  };

  // Đối chiếu song song 2 cột
  const handleCrossColumnCompare = async () => {
    if (mainData.length === 0) { alert("Không có dữ liệu!"); return; }
    if (!crossCompareColA || !crossCompareColB) { alert("Chọn đủ 2 cột!"); return; }
    setLoading(true); setProgress(0); setStatusMessage("Đang đối chiếu...");
    try {
      let matchCount = 0, mismatchCount = 0;
      const anomalies: any[] = [];
      const updatedRows = await chunkProcess(mainData, 5000, (row, idx) => {
        const valA = String(row[crossCompareColA] ?? "").trim();
        const valB = String(row[crossCompareColB] ?? "").trim();
        let isMatch = false;
        let reason = "";
        if (crossCompareRule === "exact") isMatch = valA === valB;
        else if (crossCompareRule === "normalize") isMatch = valA.toLowerCase().replace(/\s+/g, " ") === valB.toLowerCase().replace(/\s+/g, " ");
        else if (crossCompareRule === "sector_code") {
          const codeA = valA.replace(/\D/g, ""), codeB = valB.replace(/\D/g, "");
          isMatch = codeA === codeB && codeA !== "";
        } else if (crossCompareRule === "substring") isMatch = valA.toLowerCase().includes(valB.toLowerCase()) || valB.toLowerCase().includes(valA.toLowerCase());
        if (isMatch) matchCount++;
        else {
          mismatchCount++;
          if (anomalies.length < 2000) anomalies.push({ dongSTT: idx+1, valA, valB, reason: "Không khớp theo quy tắc" });
        }
        const newRow: any = {};
        Object.keys(row).forEach(k => { newRow[k] = row[k]; if (k === crossCompareColB) { newRow[`SoSanh_${crossCompareColA}_vs_${crossCompareColB}`] = isMatch ? "✅ Trùng" : "❌ Lệch"; } });
        if (!newRow[`SoSanh_${crossCompareColA}_vs_${crossCompareColB}`]) newRow[`SoSanh_${crossCompareColA}_vs_${crossCompareColB}`] = isMatch ? "✅ Trùng" : "❌ Lệch";
        return newRow;
      });
      setMainData(updatedRows);
      setColumns(Object.keys(updatedRows[0] || {}));
      setCrossCompareStats({ total: updatedRows.length, matchCount, mismatchCount });
      setCrossCompareAnomalies(anomalies);
      autoSaveSession(updatedRows, rawImportedData, Object.keys(updatedRows[0] || {}), fileName, mapping, customColConfigs);
      setStatusMessage("Đối chiếu hoàn tất!");
    } catch(e: any) { alert(e.message); }
    finally { setLoading(false); }
  };

  // Bổ sung cột ngành cấp 1 và cấp 2
  const handleAppendSectorsToMainData = async () => {
    if (mainData.length === 0) { alert("Không có dữ liệu!"); return; }
    const targetManganh = quickReportManganhCol || mapping.manganh;
    if (!targetManganh) { alert("Chọn cột mã ngành!"); return; }
    if (userSectorMap.size === 0) { alert("Tải danh mục ngành trước!"); return; }
    setLoading(true);
    setStatusMessage("Đang thêm cột cấp 1 và cấp 2...");
    try {
      const updatedRows = await chunkProcess(mainData, 5000, (row) => {
        const rawCode = row[targetManganh];
        const mng = normalizeSectorCode(rawCode);
        const sec2Code = mng ? mng.slice(0,2) : "";
        const sec2Name = userSectorMap.get(sec2Code) || "";
        let sec1Code = "";
        if (mng) { if (/^[A-Z]$/.test(mng)) sec1Code = mng.toUpperCase(); else sec1Code = getParentSectorCode(mng) || ""; }
        const sec1Name = userSectorMap.get(sec1Code) || "";
        const newRow: any = {};
        Object.keys(row).forEach(k => { newRow[k] = row[k]; if (k === targetManganh) { newRow["Mã Ngành Cấp 1"] = sec1Code; newRow["Ngành Cấp 1"] = sec1Name; newRow["Mã Ngành Cấp 2"] = sec2Code; newRow["Ngành Cấp 2"] = sec2Name; } });
        if (!newRow["Mã Ngành Cấp 1"]) { newRow["Mã Ngành Cấp 1"] = sec1Code; newRow["Ngành Cấp 1"] = sec1Name; newRow["Mã Ngành Cấp 2"] = sec2Code; newRow["Ngành Cấp 2"] = sec2Name; }
        return newRow;
      });
      setMainData(updatedRows);
      setColumns(Object.keys(updatedRows[0] || {}));
      autoSaveSession(updatedRows, rawImportedData, Object.keys(updatedRows[0] || {}), fileName, mapping, customColConfigs);
      setStatusMessage("Đã thêm cột ngành cấp 1 và 2!");
    } catch(e: any) { alert(e.message); }
    finally { setLoading(false); }
  };

  // Kiểm tra logic đa điều kiện (stub - cần phát triển thêm)
  const handleLogicCheck = async () => {
    alert("Chức năng này đang được phát triển. Vui lòng cấu hình quy tắc trong code.");
  };

  // Filter và pagination cho tab xem dữ liệu
  const filteredData = useMemo(() => {
    if (!searchTerm) return mainData;
    const term = searchTerm.toLowerCase();
    return mainData.filter(row => Object.values(row).some(val => String(val).toLowerCase().includes(term)));
  }, [mainData, searchTerm]);
  const paginatedData = filteredData.slice((viewPage-1)*pageSize, viewPage*pageSize);
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;

  // ==================== RENDER ====================
  if (!isAuthorized) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-[#111827] text-gray-100 font-sans px-4">
        <div className="w-full max-w-md bg-[#1f2937]/90 border border-purple-500/20 rounded-2xl p-8 shadow-2xl space-y-6 backdrop-blur-md">
          <div className="text-center space-y-2">
            <div className="mx-auto w-14 h-14 bg-gradient-to-tr from-purple-600 to-indigo-500 rounded-2xl flex items-center justify-center"><Lock className="w-7 h-7 text-white animate-pulse" /></div>
            <h2 className="text-xl font-bold tracking-tight text-white">CỔNG BẢO MẬT TRUY CẬP</h2>
            <p className="text-xs text-gray-400">Vui lòng nhập mật khẩu nội bộ</p>
          </div>
          <form onSubmit={(e) => { e.preventDefault(); if(typedPassword === appPassword) { localStorage.setItem("vsic_app_authorized","true"); setIsAuthorized(true); setPasswordError(""); } else setPasswordError("Mật khẩu không chính xác!"); }} className="space-y-4">
            <input type="password" value={typedPassword} onChange={e=>setTypedPassword(e.target.value)} placeholder="Nhập mật khẩu..." className="w-full bg-[#111827] border border-[#374151] rounded-xl px-4 py-3 text-sm text-white focus:border-purple-500 focus:ring-1 focus:ring-purple-500" autoFocus />
            {passwordError && <p className="text-red-400 text-xs">{passwordError}</p>}
            <button type="submit" className="w-full bg-gradient-to-r from-purple-600 to-indigo-500 hover:from-purple-700 text-white font-bold text-sm py-3 rounded-xl">🔐 Xác Nhận</button>
          </form>
          <div className="border-t border-gray-800/60 pt-4 text-center"><p className="text-[11px] text-amber-400/85">💡 Mật khẩu mặc định: <strong className="font-mono bg-amber-950 px-1.5 py-0.5 rounded">admin123</strong></p></div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-[#111827] text-gray-100 font-sans overflow-hidden">
      <header className="border-b border-[#374151] bg-[#1f2937]/90 backdrop-blur-md sticky top-0 z-40 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-tr from-purple-600 to-indigo-500 p-2.5 rounded-xl shadow-lg shadow-purple-900/30"><Layers className="w-6 h-6 text-white animate-pulse" /></div>
          <div><h1 className="text-xl font-bold tracking-tight text-white">HỆ THỐNG <span className="bg-purple-600 text-[10px] uppercase font-bold px-2 py-0.5 rounded-full">VSIC V38.5</span></h1><p className="text-xs text-gray-400 font-mono">CÔNG CỤ HỖ TRỢ SO SÁNH TỔNG HỢP DỮ LIỆU</p></div>
        </div>
        <div className="flex items-center gap-4">
          <button onClick={() => { setNewPasswordVal(""); setShowPasswordChangeModal(true); }} className="px-3 py-1.5 bg-[#111827] hover:bg-gray-800 text-gray-300 rounded-lg text-xs font-semibold flex items-center gap-1.5"><KeyRound className="w-3.5 h-3.5 text-purple-400" /> Đổi MK</button>
          <button onClick={() => { localStorage.removeItem("vsic_app_authorized"); setIsAuthorized(false); }} className="px-3 py-1.5 bg-red-950/40 hover:bg-red-900/40 text-red-300 border border-red-900/50 rounded-lg text-xs font-semibold flex items-center gap-1.5"><LogOut className="w-3.5 h-3.5" /> Khóa</button>
          {fileName ? <div className="bg-[#111827] border border-[#374151] rounded-lg px-4 py-1.5 flex items-center gap-2 text-xs"><Database className="w-4 h-4 text-emerald-400" /><span className="text-gray-300 font-medium">Hiện tại: </span><span className="text-emerald-400 font-mono max-w-[200px] truncate">{fileName}</span><span className="bg-gray-800 text-gray-400 px-1.5 py-0.5 rounded font-mono">{mainData.length} dòng</span><button onClick={() => { setMainData([]); clearAppState(); }} className="text-red-400 hover:text-red-300 ml-2">Xóa</button></div> : <span className="text-xs text-amber-400/90 bg-amber-950/40 border border-amber-900/50 rounded-lg px-4 py-1.5"><AlertTriangle className="w-3.5 h-3.5 inline mr-1" /> Chưa có dữ liệu nguồn</span>}
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden">
        <aside className="w-72 bg-[#1f2937]/60 border-r border-[#374151] p-5 space-y-2 flex flex-col justify-between">
          <div className="space-y-1.5">
            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 mb-2">Thao tác dữ liệu</div>
            <button onClick={() => setActiveTab("trangchu")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "trangchu" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><Home className="w-4 h-4" /> Trang Chủ</button>
            <button onClick={() => setActiveTab("xemdulieu")} className={`w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "xemdulieu" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><span className="flex items-center gap-3"><FileSpreadsheet className="w-4 h-4" /> Xem & Định Nghĩa Cột</span><span className="text-[10px] font-mono bg-[#111827] text-gray-400 px-1.5 py-0.5 rounded-md">{mainData.length}</span></button>
            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Công cụ liên hợp</div>
            <button onClick={() => setActiveTab("ghepnoi")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "ghepnoi" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><GitMerge className="w-4 h-4 text-blue-400" /> Ghép Nối Dữ Liệu</button>
            <button onClick={() => setActiveTab("sosanh")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "sosanh" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><Combine className="w-4 h-4 text-cyan-400" /> So Sánh Đối Chiếu</button>
            <button onClick={() => setActiveTab("tachfile")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "tachfile" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><Scissors className="w-4 h-4 text-pink-400" /> Tách File Hàng Loạt</button>
            <button onClick={() => setActiveTab("tonghop")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "tonghop" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><BarChart3 className="w-4 h-4 text-amber-400" /> Tổng Hợp Báo Cáo</button>
            <button onClick={() => setActiveTab("bieudotrucquan")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "bieudotrucquan" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><PieChart className="w-4 h-4 text-cyan-400" /> Biểu Đồ Trực Quan</button>
            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Thông minh & Rà soát</div>
            <button onClick={() => setActiveTab("chuanhoanganh")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "chuanhoanganh" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><Brain className="w-4 h-4 text-indigo-400" /> Chuẩn Hóa VSIC & AI</button>
            <button onClick={() => setActiveTab("kiemtralogic")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "kiemtralogic" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><CheckSquare className="w-4 h-4 text-emerald-400" /> Cỗ Máy Kiểm Tra Logic</button>
            <button onClick={() => setActiveTab("doichieumota")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "doichieumota" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><ArrowRightLeft className="w-4 h-4 text-purple-400" /> Đối Chiếu Mô Tả Ngành</button>
            <button onClick={() => setActiveTab("danhmucvsic")} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold ${activeTab === "danhmucvsic" ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" : "text-gray-300 hover:bg-[#374151]/50"}`}><Database className="w-4 h-4 text-amber-400" /> Danh Mục Ngành VSIC</button>
          </div>
          <div className="bg-[#111827]/80 rounded-xl p-3.5 border border-purple-950/40 text-[10px] text-gray-400 font-mono leading-relaxed"><div className="flex items-center gap-1.5 text-emerald-400 font-semibold"><span className="relative flex h-2 w-2"><span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span><span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span></span>💾 BỘ NHỚ LOCAL WORKSPACE</div><div>Dữ liệu được lưu an toàn trong IndexedDB. Tắt máy, đóng tab vẫn khôi phục 100%!</div></div>
        </aside>

        <main className="flex-1 bg-[#111827] overflow-y-auto p-6 md:p-8">
          {loading && (
            <div className="fixed inset-0 z-50 bg-[#111827]/80 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl max-w-md w-full p-6 text-center space-y-4 shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 h-1 bg-gradient-to-r from-purple-600 to-cyan-400 transition-all duration-300" style={{ width: `${progress}%` }}></div>
                <Loader2 className="w-12 h-12 text-purple-500 mx-auto animate-spin" />
                <h3 className="text-lg font-bold text-white">Đang xử lý dữ liệu</h3>
                <p className="text-sm text-gray-400 font-mono min-h-[40px]">{statusMessage}</p>
                <div className="w-full bg-[#111827] rounded-full h-2.5 overflow-hidden border border-gray-800"><div className="bg-purple-600 h-2.5 rounded-full transition-all duration-300" style={{ width: `${progress}%` }}></div></div>
                <div className="text-xs font-bold text-purple-400 tracking-wider font-mono">{progress}% Hoàn Thành</div>
              </div>
            </div>
          )}

          {showPasswordChangeModal && (
            <div className="fixed inset-0 z-50 bg-[#111827]/80 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl max-w-sm w-full p-6 space-y-4">
                <div className="text-center space-y-1"><h3 className="text-base font-bold text-white">ĐỔI MẬT KHẨU</h3></div>
                <input type="text" className="w-full bg-[#111827] border border-[#374151] rounded-lg px-3 py-2 text-xs text-white" placeholder="Mật khẩu mới" value={newPasswordVal} onChange={e=>setNewPasswordVal(e.target.value)} autoFocus />
                <div className="grid grid-cols-2 gap-2">
                  <button onClick={()=>setShowPasswordChangeModal(false)} className="bg-[#1e293b] hover:bg-gray-800 text-gray-400 font-bold text-xs py-2 rounded-lg">Hủy</button>
                  <button onClick={()=>{ if(newPasswordVal.trim()){ setAppPassword(newPasswordVal); localStorage.setItem("vsic_app_password", newPasswordVal); setShowPasswordChangeModal(false); setNewPasswordVal(""); alert(`Đã đổi mật khẩu thành ${newPasswordVal}`); } else alert("Vui lòng nhập mật khẩu mới!"); }} className="bg-gradient-to-r from-purple-600 to-indigo-500 text-white font-bold text-xs py-2 rounded-lg">Xác nhận</button>
                </div>
              </div>
            </div>
          )}

          {/* Tab Trang chủ */}
          {activeTab === "trangchu" && (
            <div className="space-y-8 animate-fade-in">
              <div className="bg-gradient-to-r from-purple-900/40 via-[#1f2937] to-[#1f2937] border border-purple-500/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-center gap-6 justify-between">
                <div className="space-y-3 max-w-2xl">
                  <span className="bg-purple-900/50 border border-purple-500/30 text-purple-400 text-xs font-mono font-bold px-3 py-1 rounded-full uppercase tracking-wider">Phiên bản V38.5</span>
                  <h2 className="text-2xl md:text-3xl font-extrabold text-white tracking-tight">Hệ Thống Phân Tích & Chuẩn Hóa Dữ Liệu Ngành</h2>
                  <p className="text-gray-300 text-sm leading-relaxed">Công cụ chuyên sâu hỗ trợ thống kê dữ liệu doanh nghiệp, ghép tách tệp lớn, so khớp, rà soát logic đa chỉ tiêu.</p>
                  <div className="pt-2 flex items-center gap-4"><button onClick={()=>setActiveTab("xemdulieu")} className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-sm px-6 py-2.5 rounded-xl shadow-md shadow-purple-900/30 flex items-center gap-2">📂 Nạp file dữ liệu <ArrowRight className="w-4 h-4" /></button></div>
                </div>
                <div className="w-full md:w-auto flex justify-center"><div className="bg-gradient-to-tr from-[#374151] to-purple-800/20 border border-[#4b5563] p-6 rounded-2xl text-center space-y-2 min-w-[200px]"><div className="text-4xl font-extrabold text-white font-mono">{userSectorMap.size}</div><div className="text-[11px] font-bold text-gray-400 tracking-wider uppercase">Mã ngành đã nạp</div><div className="text-[10px] text-green-400 font-mono">Từ danh mục của bạn</div></div></div>
              </div>
              <div className="space-y-6"><div className="border-b border-[#374151] pb-4"><h3 className="text-lg font-bold text-white flex items-center gap-2"><Layers className="w-5 h-5 text-purple-400" /> HƯỚNG DẪN NHANH</h3></div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-xs text-gray-300 leading-relaxed">
                <div className="bg-[#1f2937]/50 border border-purple-500/20 rounded-2xl p-4"><span className="font-bold text-white">01. Nạp & Tiền xử lý</span><p className="text-gray-400 mt-1">Tải file Excel/CSV chính, đặt tên cột dễ nhớ, gán vai trò (mã ngành, xã, doanh thu...).</p></div>
                <div className="bg-[#1f2937]/50 border border-indigo-500/20 rounded-2xl p-4"><span className="font-bold text-white">02. Chuẩn hóa VSIC</span><p className="text-gray-400 mt-1">Tải danh mục ngành (Excel/CSV), sau đó dùng AI để chuẩn hóa và rà lỗi logic.</p></div>
                <div className="bg-[#1f2937]/50 border border-emerald-500/20 rounded-2xl p-4"><span className="font-bold text-white">03. Tổng hợp & Xuất báo cáo</span><p className="text-gray-400 mt-1">Chọn cấp độ ngành, chạy báo cáo, xem biểu đồ, tải file Excel kết quả.</p></div>
              </div></div>
            </div>
          )}

          {/* Tab Xem dữ liệu */}
          {activeTab === "xemdulieu" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div><h3 className="text-lg font-bold text-white">FILE DỮ LIỆU NGUỒN CHÍNH</h3><p className="text-xs text-gray-400">Tải lên tệp dữ liệu chính (Excel/CSV)</p></div>
                  <label className="bg-purple-600 hover:bg-purple-700 text-white text-xs px-6 py-2.5 rounded-xl flex items-center gap-2 cursor-pointer"><FileUp className="w-4 h-4" /> TẢI FILE <input type="file" accept=".xlsx,.xls,.csv,.txt" onChange={(e)=>handleFileUpload(e,"main")} className="hidden" /></label>
                </div>
                {/* Phần cấu hình cột có thể thêm vào đây nếu cần */}
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab Ghép nối */}
          {activeTab === "ghepnoi" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white">GHÉP NỐI HAI BẢNG</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="bg-[#111827]/60 rounded-xl p-5 text-center"><h4 className="text-sm font-bold text-blue-400">BẢNG TRÁI</h4><label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-blue-300 px-4 py-2 rounded-lg cursor-pointer">Chọn File <input type="file" onChange={(e)=>handleFileUpload(e,"left")} className="hidden" /></label><div className="text-xs text-gray-400 mt-2">{leftFileName ? `📂 ${leftFileName} (${leftData.length} dòng)` : "Chưa tải"}</div>{leftData.length>0 && <select value={leftKey} onChange={e=>setLeftKey(e.target.value)} className="w-full mt-3 bg-[#111827] border border-gray-700 rounded-lg px-2 py-1 text-xs"><option value="">-- Chọn khóa --</option>{Object.keys(leftData[0]||{}).map(c=><option key={c}>{c}</option>)}</select>}</div>
                  <div className="bg-[#111827]/60 rounded-xl p-5 text-center"><h4 className="text-sm font-bold text-teal-400">BẢNG PHẢI</h4><label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-teal-300 px-4 py-2 rounded-lg cursor-pointer">Chọn File <input type="file" onChange={(e)=>handleFileUpload(e,"right")} className="hidden" /></label><div className="text-xs text-gray-400 mt-2">{rightFileName ? `📂 ${rightFileName} (${rightData.length} dòng)` : "Chưa tải"}</div>{rightData.length>0 && <select value={rightKey} onChange={e=>setRightKey(e.target.value)} className="w-full mt-3 bg-[#111827] border border-gray-700 rounded-lg px-2 py-1 text-xs"><option value="">-- Chọn khóa --</option>{Object.keys(rightData[0]||{}).map(c=><option key={c}>{c}</option>)}</select>}</div>
                </div>
                <div className="flex justify-end"><button onClick={handleMerge} className="bg-blue-600 hover:bg-blue-700 text-white text-xs px-6 py-2.5 rounded-xl">THỰC THI GHÉP NỐI</button></div>
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab So sánh */}
          {activeTab === "sosanh" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white">SO SÁNH HAI FILE CŨ & MỚI</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="bg-[#111827]/60 rounded-xl p-5 text-center"><h4 className="text-sm font-bold text-gray-400">FILE CŨ</h4><label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-white px-4 py-2 rounded-lg cursor-pointer">Tải File <input type="file" onChange={(e)=>handleFileUpload(e,"old")} className="hidden" /></label><div className="text-xs text-gray-400 mt-2">{oldFileName ? `📂 ${oldFileName} (${oldData.length} dòng)` : "Chưa tải"}</div></div>
                  <div className="bg-[#111827]/60 rounded-xl p-5 text-center"><h4 className="text-sm font-bold text-cyan-400">FILE MỚI</h4><label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-cyan-300 px-4 py-2 rounded-lg cursor-pointer">Tải File <input type="file" onChange={(e)=>handleFileUpload(e,"new")} className="hidden" /></label><div className="text-xs text-gray-400 mt-2">{newFileName ? `📂 ${newFileName} (${newData.length} dòng)` : "Chưa tải"}</div></div>
                </div>
                {oldData.length>0 && newData.length>0 && <div className="max-w-md mx-auto"><select value={diffKey} onChange={e=>setDiffKey(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-xs"><option value="">-- Chọn cột khóa chung --</option>{Object.keys(oldData[0]||{}).filter(c=>Object.keys(newData[0]||{}).includes(c)).map(c=><option key={c}>{c}</option>)}</select></div>}
                <div className="flex justify-end"><button onClick={handleCompare} className="bg-cyan-600 hover:bg-cyan-700 text-white text-xs px-6 py-2.5 rounded-xl">BẮT ĐẦU SO SÁNH</button></div>
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab Tách file */}
          {activeTab === "tachfile" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white">TÁCH FILE HÀNG LOẠT THEO CỘT</h3>
                {mainData.length>0 ? <div className="max-w-md space-y-4"><select value={splitCol} onChange={e=>setSplitCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-xs"><option value="">-- Chọn cột để tách --</option>{columns.map(c=><option key={c}>{c}</option>)}</select><button onClick={handleSplitData} className="bg-pink-600 hover:bg-pink-700 text-white text-xs px-6 py-2.5 rounded-xl w-full">TÁCH & ZIP</button></div> : <div className="text-amber-400 text-xs">Chưa có dữ liệu nguồn</div>}
              </div>
            </div>
          )}

          {/* Tab Tổng hợp báo cáo */}
          {activeTab === "tonghop" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <h3 className="text-lg font-bold text-white">TỔNG HỢP BÁO CÁO THEO NGÀNH & XÃ</h3>
                {mainData.length>0 && (
                  <>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div><label className="text-xs text-gray-400">Cột Mã ngành</label><select value={quickReportManganhCol} onChange={e=>setQuickReportManganhCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                      <div><label className="text-xs text-gray-400">Cột Xã</label><select value={quickReportXaCol} onChange={e=>setQuickReportXaCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                      <div><label className="text-xs text-gray-400">Cột Doanh thu</label><select value={quickReportDoanhThuCol} onChange={e=>setQuickReportDoanhThuCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Tùy chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                      <div><label className="text-xs text-gray-400">Cột Lao động</label><select value={quickReportLaoDongCol} onChange={e=>setQuickReportLaoDongCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Tùy chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                    </div>
                    <div className="flex gap-3">
                      <button onClick={()=>handleQuickReport(1)} className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-4 py-2 rounded-xl">Báo cáo cấp 1 (Lĩnh vực)</button>
                      <button onClick={()=>handleQuickReport(2)} className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-4 py-2 rounded-xl">Báo cáo cấp 2 (Ngành)</button>
                    </div>
                  </>
                )}
                {quickReportResultRows.length>0 && <BeautifulReportTable rows={quickReportResultRows} cols={quickReportResultCols} level={quickReportLevel} reportType={reportType} onExport={handleExportQuickReport} />}
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab Biểu đồ trực quan */}
          {activeTab === "bieudotrucquan" && (
            <SectorRevenueChart mainData={mainData} columns={columns} mapping={mapping} />
          )}

          {/* Tab Chuẩn hóa VSIC & AI */}
          {activeTab === "chuanhoanganh" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <h3 className="text-lg font-bold text-white">CHUẨN HÓA & KHỚP MÃ NGÀNH THÔNG MINH</h3>
                {mainData.length>0 && (
                  <>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div><label className="text-xs text-gray-400">Cột Mã ngành (cấp 5)</label><select value={stdIndustryCol} onChange={e=>setStdIndustryCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                      <div><label className="text-xs text-gray-400">Cột Mô tả hoạt động</label><select value={stdDescriptionCol} onChange={e=>setStdDescriptionCol(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded-lg px-2 py-1 text-sm"><option value="">-- Chọn --</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div>
                    </div>
                    <button onClick={handleStandardizeSectorsAndMatch} className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs px-4 py-2 rounded-xl">CHUẨN HÓA MÃ NGÀNH</button>
                  </>
                )}
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab Kiểm tra logic */}
          {activeTab === "kiemtralogic" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white">KIỂM TRA LOGIC ĐA ĐIỀU KIỆN</h3>
                <div className="text-xs text-gray-400">Chức năng đang phát triển. Vui lòng cấu hình quy tắc trong code.</div>
                <button onClick={handleLogicCheck} className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-4 py-2 rounded-xl">KIỂM TRA LOGIC</button>
              </div>
              {mainData.length > 0 && <DataPreviewTable data={mainData} columns={columns} />}
            </div>
          )}

          {/* Tab Đối chiếu mô tả ngành */}
          {activeTab === "doichieumota" && (
            <DescriptorMatchScanner mainData={mainData} columns={columns} mapping={mapping} />
          )}

          {/* Tab Danh mục ngành VSIC */}
          {activeTab === "danhmucvsic" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div><h3 className="text-lg font-bold text-white flex items-center gap-2"><Database className="w-5 h-5 text-cyan-400" /> NẠP BỔ SUNG DANH MỤC NGÀNH CỦA BẠN (EXCEL / CSV)</h3><p className="text-xs text-gray-400">Tải lên file danh mục (mã, tên) để hệ thống sử dụng làm chuẩn.</p></div>
                  <div className="flex gap-3"><label className="bg-cyan-600 hover:bg-cyan-700 text-white font-bold text-xs px-4 py-2 rounded-xl flex items-center gap-2 cursor-pointer"><FileUp className="w-4 h-4" /> TẢI FILE <input type="file" accept=".xlsx,.xls,.csv" onChange={handleUploadUserSectorCatalog} className="hidden" /></label><button onClick={handleClearUserSectors} className="bg-red-800/60 hover:bg-red-700/80 text-red-200 font-bold text-xs px-4 py-2 rounded-xl">Xóa sạch</button></div>
                </div>
                {userSectorFileName && <div className="bg-[#111827] rounded-xl p-3 border border-cyan-500/20 flex items-center gap-2 text-xs"><CheckCircle2 className="w-4 h-4 text-cyan-400" /><span className="text-gray-300">Đã nạp: <strong>{userSectorFileName}</strong> ( {userSectorMap.size} mã ngành )</span><span className="text-green-400 ml-auto">✅ Sẵn sàng</span></div>}
                <div className="text-xs text-gray-500 border-t border-gray-800 pt-3">💡 Hướng dẫn: File cần có ít nhất 2 cột (Mã ngành, Tên ngành). Hệ thống tự động phát hiện tên cột.</div>
              </div>
              <VsicCatalogExplorer />
            </div>
          )}
        </main>
      </div>
    </div>
  );
}