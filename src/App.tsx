import React, { useState, useMemo, useEffect } from "react";
import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import SampleSelector from "./components/SampleSelector";
import { Analytics } from '@vercel/analytics/react';
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

export interface AiMacro {
  id: string;
  name: string;
  prompt: string;
  module: "tonghop" | "chuanhoanganh";
  columns: {
    xa?: string;
    manganh?: string;
    doanhthu?: string;
    laodong?: string;
    mota?: string;
  };
  level?: number;
}


function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    try {
      if (typeof window === "undefined" || !window.indexedDB) {
        reject(new Error("IndexedDB is not supported or is disabled in this environment."));
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onerror = () => {
        reject(new Error("IndexedDB opening failed due to security/sandbox limits."));
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
    } catch (e: any) {
      reject(e);
    }
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
    
    // Yield to the main thread to remain fully responsive
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
    // Tránh giữ luồng thao tác liên tục
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
    
    // 1. Luôn ghi nhận Metadata (rất nhẹ, < 1ms)
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

    // 2. Chỉ thực hiện lưu chuỗi khối dữ liệu khổng lồ nếu được bật cờ bắt buộc (nhập file mới, tính toán lại, gộp sheet)
    if (forceSaveData) {
      await clearOldChunks(db);
      await saveArrayInChunks(db, "mainData", mainData);
      await saveArrayInChunks(db, "rawImportedData", rawImportedData);
    }

    // Xóa định dạng session cũ nếu có
    await new Promise<void>((resolve) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      store.delete("currentSession");
      resolve();
    });
  } catch (error) {
    console.info("IndexedDB Save Info: IndexedDB storage is inactive or sandboxed.", error);
  }
}

async function loadAppState(): Promise<AppState | null> {
  try {
    const db = await openDB();
    
    // 1. Tải Metadata trước
    const meta: any = await new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get("sessionMeta");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Lỗi đọc sessionMeta"));
    });

    if (meta && meta.isChunked) {
      const mainDataLength = meta.mainDataLength || 0;
      const rawImportedDataLength = meta.rawImportedDataLength || 0;
      
      const mainData = await loadArrayInChunks(db, "mainData", mainDataLength);
      const rawImportedData = await loadArrayInChunks(db, "rawImportedData", rawImportedDataLength);
      
      return {
        ...meta,
        mainData,
        rawImportedData
      };
    }

    // Tương thích ngược: Đọc session thô cũ
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
          resolve(val);
        }
      };
      request.onerror = () => reject(request.error || new Error("Lỗi đọc currentSession"));
    });
  } catch (error) {
    console.info("IndexedDB Load Info: IndexedDB storage is inactive or sandboxed.", error);
    return null;
  }
}

async function clearAppState(): Promise<void> {
  try {
    const db = await openDB();
    await clearOldChunks(db);
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.delete("sessionMeta");
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error("Lỗi xóa sessionMeta"));
    });
  } catch (error) {
    console.info("IndexedDB Clear Info: IndexedDB storage is inactive or sandboxed.", error);
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
  Upload,
  Sparkles,
  Loader2, 
  FileUp, 
  AlertTriangle, 
  CheckCircle2, 
  Info, 
  Brain,
  Layers,
  ArrowRight,
  ArrowRightLeft,
  ArrowLeftRight,
  BrainCircuit,
  Database,
  RefreshCw,
  Search,
  Plus,
  Trash2,
  FileCheck,
  Compass,
  Lock,
  KeyRound,
  LogOut,
  Sliders,
  FileText,
  Cpu,
  Zap,
  Save
} from "lucide-react";

import { 
  vsicRawData, 
  normalizeSectorCode, 
  getSectorHierarchy, 
  smartSuggestSectorByDescription,
  getSectorLevel,
  getParentSectorCode,
  lookupSectorNameWithFallback,
  isSummaryRow,
  clearAllSectorsInVSIC,
  clearAllParentsInVSIC
} from "./data/vsic";

import VsicCatalogExplorer from "./components/vsicCatalogExplorer";
import { BeautifulReportTable } from "./components/BeautifulReportTable";
import { MainDataInlinePreview } from "./components/MainDataInlinePreview";
import SectorRevenueChart, { parseRobustNumber } from "./components/sectorRevenueChart";
import PdfToWord from "./components/PdfToWord";

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
  isFieldCompare?: boolean;
}

// Hàm hỗ trợ so khớp đàn hồi mềm dẻo và lấy dữ liệu cột từ hàng (row) để chống sai phông chữ bừa bãi hoặc hoa thường lệch lạc từ AI
export function getFlexibleValue(row: any, keyName: string): any {
  if (!row || !keyName) return "";
  if (keyName in row) return row[keyName];

  const cleanKey = keyName.toLowerCase().replace(/\s+/g, "").trim();
  const actualKeys = Object.keys(row);

  // Thử khớp viết thường không khoảng trắng
  for (const k of actualKeys) {
    const cleanK = k.toLowerCase().replace(/\s+/g, "").trim();
    if (cleanK === cleanKey) {
      return row[k];
    }
  }

  // Thử khớp không dấu tiếng Việt
  const stripDiacritics = (str: string) => 
    str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, "");

  const searchNormalized = stripDiacritics(keyName);
  for (const k of actualKeys) {
    if (stripDiacritics(k) === searchNormalized) {
      return row[k];
    }
  }

  return row[keyName] || "";
}

// Hàm chuẩn hóa biểu thức AI sang biểu diễn an toàn qua getFlexibleValue
export function normalizeAiExpression(expr: string): string {
  if (!expr) return "";
  let clean = expr;
  
  // 1. Chuyển đổi các cặp row['Cột'] hoặc row["Cột"] thành getFlexibleValue(row, 'Cột')
  clean = clean.replace(/(?:row|Row)\s*\[\s*['"]([^'"]+)['"]\s*\]/g, "getFlexibleValue(row, '$1')");
  
  // 2. Chuyển đổi các dạng row.TenCol thành getFlexibleValue(row, 'TenCol')
  // Chỉ khớp thuộc tính alpha-numeric bắt đầu bằng chữ cái
  clean = clean.replace(/(?:row|Row)\.([a-zA-Z_][a-zA-Z0-9_]*)/g, "getFlexibleValue(row, '$1')");
  
  return clean;
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
    // Cho phép tự động vào ứng dụng ngay lần đầu (mặc định là true trừ phi chủ động Log Out)
    return localStorage.getItem("vsic_app_authorized") !== "false";
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
    localStorage.setItem("vsic_app_authorized", "false");
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

  // Trạng thái cho Cột tính toán mới (Calculated Columns)
  const [calcColName, setCalcColName] = useState<string>("");
  const [calcCol1, setCalcCol1] = useState<string>("");
  const [calcCol2, setCalcCol2] = useState<string>("");
  const [calcOperator, setCalcOperator] = useState<"+" | "-" | "*" | "/" | "concat">("+");
  const [calcType, setCalcType] = useState<"column" | "constant">("column");
  const [calcConstant, setCalcConstant] = useState<string>("");
  const [calcRounding, setCalcRounding] = useState<"none" | "int" | "1dec" | "2dec">("none");

  // Trạng thái cho Dual-Pane Mapping và double click, cùng kiểu định dạng báo cáo xoay Pivot
  const [selectedTargetKey, setSelectedTargetKey] = useState<keyof ColumnMapping>("mota");
  const [reportType, setReportType] = useState<"flat" | "pivot">("pivot");
  const [isConfigExpanded, setIsConfigExpanded] = useState<boolean>(true);

  // AI Học Đặt Tên Cột (Column Rule Learning)
  const [aiColLearnPrompt, setAiColLearnPrompt] = useState<string>("");
  const [isLearningColAi, setIsLearningColAi] = useState<boolean>(false);
  const [learningColLogs, setLearningColLogs] = useState<string[]>([]);
  const [newColCommandName, setNewColCommandName] = useState<string>("");
  const [colLearnedCommands, setColLearnedCommands] = useState<{
    id: string;
    name: string;
    description: string;
    createdAt: string;
    mappings: {
      originalMatch: string;
      newName: string;
      role: "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" | "";
      use: boolean;
    }[];
  }[]>(() => {
    try {
      const saved = localStorage.getItem("colLearnedCommands");
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.warn("Lỗi đọc colLearnedCommands từ localStorage:", e);
    }
    // Danh mục lệnh học mặc định siêu thông minh
    return [
      {
        id: "default-thue",
        name: "🏷️ Quy chuẩn về dạng cột Tiếng Việt",
        description: "Nhận diện & Việt hóa: Mã Số Thuế -> 🔑 MST, Xa/Phuong -> 🗺️ Địa bàn Xã, DoanhThu -> 💰 Doanh Thu, LaoDong -> 👥 Số Lao Động.",
        createdAt: new Date().toISOString(),
        mappings: [
          { originalMatch: "MST", newName: "Mã Số Thuế", role: "idCol", use: true },
          { originalMatch: "MaST", newName: "Mã Số Thuế", role: "idCol", use: true },
          { originalMatch: "Mã Số Thuế", newName: "Mã Số Thuế", role: "idCol", use: true },
          { originalMatch: "Xa", newName: "Địa bàn (Xã)", role: "xa", use: true },
          { originalMatch: "Phuong", newName: "Địa bàn (Xã)", role: "xa", use: true },
          { originalMatch: "DiaBan", newName: "Địa bàn (Xã)", role: "xa", use: true },
          { originalMatch: "DoanhThu", newName: "Doanh Thu", role: "doanhthu", use: true },
          { originalMatch: "DoanhSo", newName: "Doanh Thu", role: "doanhthu", use: true },
          { originalMatch: "LaoDong", newName: "Số Lao Động", role: "laodong", use: true },
          { originalMatch: "NhanSu", newName: "Số Lao Động", role: "laodong", use: true },
          { originalMatch: "MaNganh", newName: "Mã Ngành Đăng Ký", role: "manganh", use: true },
          { originalMatch: "MoTa", newName: "Mô Tả Hoạt Động", role: "mota", use: true },
          { originalMatch: "NganhNghe", newName: "Mô Tả Hoạt Động", role: "mota", use: true }
        ]
      },
      {
        id: "default-rutgon",
        name: "🧹 Tối Giản Hóa (Chỉ giữ ID và Mô Tả Ngành)",
        description: "Loại bỏ mọi cột thừa ngoại trừ 🔑 Mã định danh (MST) và 📝 Mô tả hoạt động kinh doanh để tối ưu hóa hiệu năng rà soát.",
        createdAt: new Date().toISOString(),
        mappings: [
          { originalMatch: "MST", newName: "Mã Số Thuế", role: "idCol", use: true },
          { originalMatch: "MaST", newName: "Mã Số Thuế", role: "idCol", use: true },
          { originalMatch: "ID", newName: "Mã Định Danh", role: "idCol", use: true },
          { originalMatch: "MoTa", newName: "Mô Tả Hoạt Động", role: "mota", use: true },
          { originalMatch: "NganhNghe", newName: "Mô Tả Hoạt Động", role: "mota", use: true }
        ]
      }
    ];
  });

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
  const [logicRuleMode, setLogicRuleMode] = useState<"conflict" | "must_satisfy">("conflict");
  const [logicFilterMode, setLogicFilterMode] = useState<"all" | "if_satisfied" | "violated">("if_satisfied");

  // Trí tuệ Nhân tạo - Học và lưu lệnh thông minh qua AI
  const [aiRulePrompt, setAiRulePrompt] = useState<string>("");
  const [aiTranslatedExpression, setAiTranslatedExpression] = useState<string>("");
  const [customRuleName, setCustomRuleName] = useState<string>("");
  const [aiScanMetrics, setAiScanMetrics] = useState<{
    total: number;
    violated: number;
    passed: number;
    violatedPercent: string;
    passedPercent: string;
    expression: string;
    prompt: string;
  } | null>(null);
  const [savedAiRules, setSavedAiRules] = useState<{ id: string; name: string; prompt: string; expression: string }[]>(() => {
    try {
      const saved = localStorage.getItem("vsic_saved_ai_rules");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  // Quy tắc mới cho Logic
  const [newIfRule, setNewIfRule] = useState<LogicRule>({ col: "", op: "==", val: "", isFieldCompare: false });
  const [newThenRule, setNewThenRule] = useState<LogicRule>({ col: "", op: "==", val: "", isFieldCompare: false });

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
  const [quickReportSumCols, setQuickReportSumCols] = useState<string[]>([]);
  const [reportAiPrompt, setReportAiPrompt] = useState<string>("");
  const [reportAiLogs, setReportAiLogs] = useState<string[]>([]);
  const [isReportAiRunning, setIsReportAiRunning] = useState<boolean>(false);
  const [schemaAiPrompt, setSchemaAiPrompt] = useState<string>("");
  const [isSchemaAiRunning, setIsSchemaAiRunning] = useState<boolean>(false);
  const [schemaAiLogs, setSchemaAiLogs] = useState<string[]>([]);
  const [pivotManganhCol, setPivotManganhCol] = useState<string>("");

  // States cho Phân Hệ 2: Chuẩn hóa khớp ngành VSIC cấp 5 (Mới độc lập)
  const [stdIndustryCol, setStdIndustryCol] = useState<string>("");
  const [stdDescriptionCol, setStdDescriptionCol] = useState<string>("");
  const [stdReportAnomalies, setStdReportAnomalies] = useState<any[]>([]);
  const [stdMatchStats, setStdMatchStats] = useState<{ total: number; valid: number; invalid: number; conflicts: number }>({ total: 0, valid: 0, invalid: 0, conflicts: 0 });

  // States cho Phân Hệ Đối chiếu 2 cột tự chọn (Theo yêu cầu người dùng)
  const [crossCompareColA, setCrossCompareColA] = useState<string>("");
  const [crossCompareColB, setCrossCompareColB] = useState<string>("");
  const [crossCompareRule, setCrossCompareRule] = useState<string>("semantic");
  const [crossCompareAnomalies, setCrossCompareAnomalies] = useState<any[]>([]);
  const [crossCompareStats, setCrossCompareStats] = useState<{ total: number; matchCount: number; mismatchCount: number }>({ total: 0, matchCount: 0, mismatchCount: 0 });

  // --- TRẠNG THÁI & PHƯƠNG THỨC CHO HỆ THỐNG TỰ HỌC LỆNH THÔNG MINH (AI MACRO STORAGE) ---
  const [aiMacros, setAiMacros] = useState<AiMacro[]>([]);
  const [macroPrompt, setMacroPrompt] = useState<string>("");
  const [isLearning, setIsLearning] = useState<boolean>(false);
  const [learningResult, setLearningResult] = useState<any | null>(null);
  const [customMacroName, setCustomMacroName] = useState<string>("");

  useEffect(() => {
    const saved = localStorage.getItem("ai_macros_vsic");
    if (saved) {
      try {
        setAiMacros(JSON.parse(saved));
      } catch (e) {
        setAiMacros([]);
      }
    } else {
      const sampleMacros: AiMacro[] = [
        {
          id: "m_sample_1",
          name: "📊 Báo cáo Doanh thu & Lao động theo Xã (Ngành Cấp 2)",
          prompt: "Tổng hợp các số liệu theo xã và mã ngành cấp 2 tập trung vào doanh thu và tổng lao động thực tế.",
          module: "tonghop",
          columns: {
            xa: "Địa_Bàn_Xã",
            manganh: "Mã_Ngành_VSIC",
            doanhthu: "Doanh_Thu_Tích_Lũy",
            laodong: "Tổng Lao động",
            mota: ""
          },
          level: 2
        },
        {
          id: "m_sample_2",
          name: "🧠 Kiểm tra VSIC: Mô tả thực tế vs Mã Ngành Cấp 5",
          prompt: "Kiểm tra mã ngành và rà quét sự không đồng nhất giữa mô tả chi tiết với mã ngành.",
          module: "chuanhoanganh",
          columns: {
            xa: "",
            manganh: "Mã_Ngành_VSIC",
            doanhthu: "",
            laodong: "",
            mota: "Mô_Tả_Hoạt_Động"
          }
        }
      ];
      setAiMacros(sampleMacros);
      localStorage.setItem("ai_macros_vsic", JSON.stringify(sampleMacros));
    }
  }, []);

  const handleLearnMacro = async () => {
    if (!macroPrompt.trim()) {
      alert("Vui lòng nhập câu lệnh tiếng Việt để AI học phương án chọn cột!");
      return;
    }
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu nguồn chính trước khi yêu cầu AI học lệnh chọn cột.");
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần Cài đặt > Khóa bí mật (Secrets) của AI Studio hoặc trong file .env!");
      return;
    }

    setIsLearning(true);
    setStatusMessage("Hệ thống Macro AI đang phân tích khẩu lệnh tiếng Việt và ánh xạ các cột dữ liệu...");
    setLoading(true);
    setProgress(20);

    try {
      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build"
          }
        }
      });

      const systemPrompt = `Bạn là chuyên gia phân tích khẩu lệnh tiếng Việt và tự động ánh xạ (mapping) các cột dữ liệu tương ứng cho hai phân hệ (module) phân tích:
1. Phân hệ "tonghop" (Tổng hợp báo cáo nhanh): Yêu cầu tìm các cột Địa bàn Xã, Mã Ngành, Doanh thu, Lao động và xác định cấp ngành (1 hoặc 2) để hạch toán.
2. Phân hệ "chuanhoanganh" (Kiểm tra & Chuẩn hoá ngành): Yêu cầu tìm các cột chứa Mô tả hoạt động kinh doanh (hoạt động kinh tế) và Mã Ngành VSIC để chuẩn hóa.

Dưới đây là danh sách các cột đang có thực tế trong bảng dữ liệu Excel/CSV của người dùng:
[${columns.filter(c => !c.startsWith("_")).map(c => `'${c}'`).join(", ")}]

Nhiệm vụ của bạn là phân tích câu lệnh tiếng Việt của người dùng và chọn ra những cột khớp nhất từ danh sách trên. Không tự ý chế tên cột không tồn tại trong danh sách. 
Nếu không khớp được cột nào phù hợp hoặc câu lệnh không đề cập đến trường đó, hãy trả về giá trị chuỗi rỗng "" cho cột đó.

Hãy phân tích kỹ các cột:
- Với cột "xa": tìm các cột có tên chứa từ "xã", "xa", "phường", "địa bàn", "dia ban", "địa phương", "dia phuong".
- Với cột "manganh": tìm các cột có tên chứa "mã ngành", "manganh", "vsic", "mã ngành cấp 5", "mã hoạt động".
- Với cột "doanhthu": tìm các cột chứa "doanh thu", "doanhthu", "thu nhập", "thunhap", "tiền", "tien", "doanh số".
- Với cột "laodong": tìm các cột chứa "lao động", "laodong", "số người", "quy mô lao động", "nhân sự".
- Với cột "mota": tìm các cột chứa từ "mô tả", "mota", "hoạt động kinh doanh", "tên ngành", "ngành nghề thực tế", "noi dung".

Hãy xác định trường "module" dựa vào ý định:
- Nếu người dùng muốn "tổng hợp", "báo cáo", "hạch toán", "tính tỉ trọng", "tỷ trọng", "thống kê theo xã/ngành": "module" là "tonghop".
- Nếu người dùng muốn "kiểm tra", "chuẩn hóa", "so khớp mô tả", "rà quét lệch mã", "khớp mã ngành vsic": "module" là "chuanhoanganh".

Hãy trả về một định dạng JSON duy nhất, KHÔNG giải thích dông dài, KHÔNG bọc trong khối markdown \`\`\`, định dạng chính xác tuyệt đối như sau:
{
  "module": "tonghop" hoặc "chuanhoanganh",
  "name": "Tên gợi ý ngắn gọn cho lệnh đã học, ví dụ 'Báo cáo Doanh thu Xã'",
  "columns": {
    "xa": "Tên cột khớp với địa bàn xã (nếu có)",
    "manganh": "Tên cột khớp với mã ngành/vsic (nếu có)",
    "doanhthu": "Tên cột khớp với doanh thu (nếu có)",
    "laodong": "Tên cột khớp với lao động/quy mô (nếu có)",
    "mota": "Tên cột khớp với mô tả hoạt động thực tế (nếu có)"
  },
  "level": 1 hoặc 2 (nếu module là "tonghop", xác định cấp ngành là 1 hay 2. Nếu không đề cập thì mặc định trả về 2)
}`;

      setProgress(50);
      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: `Câu lệnh tiếng Việt cần học: "${macroPrompt}"\nHãy phân tích và trả về JSON ánh xạ khớp đúng cột.`,
        config: {
          systemInstruction: systemPrompt,
          temperature: 0.1,
          responseMimeType: "application/json"
        }
      });

      setProgress(85);
      const textResult = response.text || "";
      let parsed = JSON.parse(textResult.trim());
      
      setLearningResult(parsed);
      setCustomMacroName(parsed.name || `Học lệnh: ${macroPrompt.substring(0, 30)}`);
      
      setProgress(100);
      setStatusMessage("AI học lệnh thành công! Bạn có thể đặt lại tên và bấm Lưu.");
      await sleep(400);

    } catch (err: any) {
      alert("AI không thể dịch và học lệnh này. Lỗi: " + err.message);
    } finally {
      setIsLearning(false);
      setLoading(false);
    }
  };

  const handleSaveLearnMacro = () => {
    if (!learningResult) {
      alert("Chưa có kết quả dịch từ AI để lưu!");
      return;
    }
    const macroName = customMacroName.trim() || learningResult.name || "Lệnh tự học chưa đặt tên";
    
    const newMacro: AiMacro = {
      id: "macro_" + Date.now(),
      name: macroName,
      prompt: macroPrompt,
      module: learningResult.module,
      columns: learningResult.columns,
      level: learningResult.level
    };

    const updated = [newMacro, ...aiMacros];
    setAiMacros(updated);
    localStorage.setItem("ai_macros_vsic", JSON.stringify(updated));
    setLearningResult(null);
    setMacroPrompt("");
    alert(`Đã lưu thành công lệnh học "${macroName}" vào bộ nhớ Workspace vạn năng!`);
  };

  const handleExecuteMacro = async (macro: AiMacro) => {
    if (!macro) return;
    setActiveTab(macro.module);

    if (macro.module === "tonghop") {
      const colManganh = macro.columns?.manganh || "";
      const colXa = macro.columns?.xa || "";
      const colDoanhThu = macro.columns?.doanhthu || "";
      const colLaoDong = macro.columns?.laodong || "";
      const targetLevel = macro.level || 2;

      setQuickReportManganhCol(colManganh);
      setQuickReportXaCol(colXa);
      setQuickReportDoanhThuCol(colDoanhThu);
      setQuickReportLaoDongCol(colLaoDong);

      setLoading(true);
      setProgress(20);
      setStatusMessage(`[Tái sử dụng 0s] Đang chạy báo cáo nhanh: ${macro.name}...`);
      await sleep(300);
      setProgress(60);

      try {
        await handleQuickReport(targetLevel, colManganh, colXa, colDoanhThu, colLaoDong);
        setStatusMessage(`Tổng hợp thành công lệnh: ${macro.name}`);
        setProgress(100);
      } catch (err: any) {
        alert("Lỗi khi chạy báo cáo nhanh: " + err.message);
      } finally {
        setLoading(false);
      }
    } else if (macro.module === "chuanhoanganh") {
      const colIndustry = macro.columns?.manganh || "";
      const colDesc = macro.columns?.mota || "";

      setStdIndustryCol(colIndustry);
      setStdDescriptionCol(colDesc);

      setLoading(true);
      setProgress(20);
      setStatusMessage(`[Tái sử dụng 0s] Đang chuẩn hóa - đối chiếu khớp ngành VSIC...`);
      await sleep(300);
      setProgress(60);

      try {
        await handleStandardizeSectorsAndMatch(colIndustry, colDesc);
        setStatusMessage(`Đối sánh thành công lệnh: ${macro.name}`);
        setProgress(100);
      } catch (err: any) {
        alert("Lỗi khi chạy chuẩn hóa đối chiếu: " + err.message);
      } finally {
        setLoading(false);
      }
    }
  };

  const handleDeleteMacro = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm("Bạn có chắc chắn muốn xóa lệnh học này?")) return;
    const updated = aiMacros.filter(m => m.id !== id);
    setAiMacros(updated);
    localStorage.setItem("ai_macros_vsic", JSON.stringify(updated));
  };

  const handleExportMacros = () => {
    if (aiMacros.length === 0) {
      alert("Không có lệnh học nào để xuất!");
      return;
    }
    const blob = new Blob([JSON.stringify(aiMacros, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `bo_nho_hoc_lenh_ai_macro_${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleImportMacros = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const imported = JSON.parse(event.target?.result as string);
        if (Array.isArray(imported)) {
          const validated = imported.filter(m => m && m.name && m.module && m.columns);
          if (validated.length === 0) {
            alert("File không chứa dữ liệu cấu hình học lệnh hợp lệ!");
            return;
          }
          const merged = [...validated, ...aiMacros];
          const unique = merged.filter((item, index, self) =>
            self.findIndex(t => t.prompt === item.prompt || t.id === item.id) === index
          );
          setAiMacros(unique);
          localStorage.setItem("ai_macros_vsic", JSON.stringify(unique));
          alert(`Đã nạp thành công và đồng bộ ${validated.length} lệnh học vạn năng!`);
        } else {
          alert("Tệp cấu hình không đúng định dạng JSON mảng lệnh!");
        }
      } catch (err: any) {
        alert("Có lỗi xảy ra khi nạp tệp: " + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  // Báo cáo động độc lập khởi tạo trống để người dùng tự chọn
  useEffect(() => {
    // Không tự động đoán gán cứng các cột phục vụ báo cáo động nữa.
  }, [columns]);


  // Tự động gán lựa chọn cột báo cáo nhanh dựa theo cấu hình gán cột có sẵn của dữ liệu nguồn
  useEffect(() => {
    if (mapping.manganh && !quickReportManganhCol) {
      setQuickReportManganhCol(mapping.manganh);
    }
    if (mapping.xa && !quickReportXaCol) {
      setQuickReportXaCol(mapping.xa);
    }
    if (mapping.doanhthu && !quickReportDoanhThuCol) {
      setQuickReportDoanhThuCol(mapping.doanhthu);
    }
    if (mapping.laodong && !quickReportLaoDongCol) {
      setQuickReportLaoDongCol(mapping.laodong);
    }
  }, [mapping]);



  // Phân trang cho viewer
  const [viewPage, setViewPage] = useState<number>(1);
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [inconsistenciesTab, setInconsistenciesTab] = useState<"desc" | "code">("desc");
  const pageSize = 50;

  // Giới hạn hiển thị cho danh sách bất nhất tránh làm đơ trình duyệt khi nạp hàng ngàn dòng
  const [visibleDescInconCount, setVisibleDescInconCount] = useState<number>(50);
  const [visibleCodeInconCount, setVisibleCodeInconCount] = useState<number>(50);

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

  // Reset giới hạn xem mâu thuẫn để tránh lag trình duyệt khi chuyển tab
  useEffect(() => {
    setVisibleDescInconCount(50);
    setVisibleCodeInconCount(50);
  }, [mainData, activeTab]);

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
      }, true); // Bắt buộc lưu toàn bộ chuỗi khối dữ liệu thô
    } catch (err) {
      console.warn("Không thể lưu trạng thái phiên (sử dụng cache bộ nhớ):", err);
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
      }, false).catch(err => console.warn("Lỗi tự động lưu phiên làm việc:", err)); // Chỉ lưu metadata siêu nhanh
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

  // Helper xử lý mảng lớn theo từng cụm (chunks) để tránh treo luồng chính (Main Thread)
  const chunkProcess = async <T, R>(
    array: T[],
    size: number,
    processFn: (item: T, index: number) => R,
    onProgress?: (percent: number) => void
  ): Promise<R[]> => {
    const result: R[] = [];
    const len = array.length;
    if (len === 0) return [];
    
    for (let i = 0; i < len; i += size) {
      const chunk = array.slice(i, i + size);
      for (let j = 0; j < chunk.length; j++) {
        result.push(processFn(chunk[j], i + j));
      }
      if (onProgress) {
        onProgress(Math.min(100, Math.round((i / len) * 100)));
      }
      // Trả lại quyền hoạt động cho trình duyệt vẽ giao diện & xử lý sự kiện
      await new Promise<void>(resolve => setTimeout(resolve, 0));
    }
    return result;
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

            // Giữ mọi cột trống hoàn toàn để người dùng lựa chọn thủ công tại các chức năng tương ứng
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

            // Giữ mọi cột trống hoàn toàn để người dùng tự do lựa chọn thủ công tại các chức năng tương ứng
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
    if (window.confirm("Bạn có chắc chắn muốn dọn sạch toàn bộ dữ liệu nguồn và thông số cấu hình cột hiện tại không?")) {
      setMainData([]);
      setRawImportedData([]);
      setColumns([]);
      setFileName("");
      setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
      setCustomColConfigs([]);
      clearAppState().catch(err => console.warn("Lỗi khi xóa dữ liệu IndexedDB:", err));

      const deleteCatalog = window.confirm("Bạn có muốn XÓA SẠCH DỨT ĐIỂM cả Danh mục mã ngành VSIC tùy chỉnh đã lưu trong trình duyệt không (để nạp lại tệp mới từ đầu sạch 100%)?");
      if (deleteCatalog) {
        localStorage.removeItem("custom_vsic_data");
        localStorage.removeItem("custom_vsic_parents");
        localStorage.setItem("custom_vsic_is_pure", "true");
        clearAllSectorsInVSIC();
        clearAllParentsInVSIC();
        alert("Đã xóa dứt điểm toàn bộ dữ liệu giao dịch và Danh mục mã ngành chuẩn tùy chỉnh!");
        window.location.reload();
      } else {
        alert("Đã xóa dữ liệu giao dịch thành công! (Giữ lại danh mục mã ngành chuẩn đã nạp)");
      }
    }
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
        if (isSummaryRow(row)) return;
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
              s1Code = getParentSectorCode(mngNormalized) || "";
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
          dtVal = parseRobustNumber(row[crossReportDoanhThuCol]);
        }

        // Trích xuất lao động
        let ldVal = 0;
        if (crossReportLaoDongCol) {
          ldVal = parseRobustNumber(row[crossReportLaoDongCol]);
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
          "Mã ngành tổng hợp": val.nganhCode,                  // Đã đồng nhất tên
          "Tên phân loại ngành kinh tế": val.nganhLabel,
          "Số lượng Đơn vị (DV)": val.countDN,                 // Chữ Đ viết hoa
          "Tổng Doanh Thu": Math.round(val.sumDoanhThu * 100) / 100,
          "Tổng Lao Động": Math.round(val.sumLaoDong * 100) / 100
        });

        totalDN += val.countDN;
        totalDoanhThu += val.sumDoanhThu;
        totalLaoDong += val.sumLaoDong;
      });

      // Tạo dòng sum toàn bảng
      reportRows.push({
        "STT": "LŨY KẾ",
        "Địa bàn (Xã)": "TỔNG CỘNG LŨY KẾ TOÀN BỘ BẢNG",
        "Mã ngành tổng hợp": "-",                              // Đã đồng nhất giống hệt bên trên
        "Tên phân loại ngành kinh tế": "-",
        "Số lượng Đơn vị (DV)": totalDN,                       // Đã đồng nhất chữ Đ viết hoa
        "Tổng Doanh Thu": Math.round(totalDoanhThu * 100) / 100,
        "Tổng Lao Động": Math.round(totalLaoDong * 100) / 100
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
        { wch: 20 }, // Số lượng DV
        { wch: 22 }, // Tổng doanh thu
        { wch: 22 }  // Tổng lao động
      ];

      XLSX.writeFile(wb, `Bao_Cao_Tong_Hop_Phoi_Hop_Xa_Nganh_Cap_${crossReportLevel}.xlsx`);
    } catch (err: any) {
      alert("Lỗi xuất Excel: " + err.message);
    }
  };

  // Tự động gán nhãn hàng loạt theo mẫu điều tra A1-A7, giữ các cột định danh MaTKCS, Maxa, MaDiaBan, IDcoso, chuẩn bị tính Tổng lao động
  const applyA1A7Template = () => {
    if (customColConfigs.length === 0) {
      alert("Hãy tải tệp tin Excel lên trước để có khung cột áp dụng mẫu điều tra A1-A7!");
      return;
    }

    const dict: { [key: string]: string } = {
      "a1_1": "Địa điểm",
      "a1_2": "Địa điểm này cơ sở đi thuê/mượn",
      "a1_3": "Họ tên",
      "a1_3_1": "Giới tính",
      "a1_3_2": "Năm sinh",
      "a1_3_3": "Dân tộc",
      "a1_3_4": "Quốc tịch",
      "a1_3_5": "Trình độ",
      "a1_4": "Tình trạng (ĐKKD)",
      "a1_5_1": "Mã số thuế",
      "a2_1": "LĐ Thuê",
      "a2_2": "LĐ gia đình",
      "a2_3_1": "lao động nữ",
      "a3_1_1_1": "TSCĐ nhà",
      "a3_1_2_1": "TSCĐ Phương tiện",
      "a3_1_3_1": "TSCĐ máy Móc",
      "a3_1_4_1": "Dụng cụ",
      "a3_1_5_1": "TSCD khác",
      "a3_1t": "Tổng TSCĐ",
      "a3_2": "Vốn bỏ ra SXKD",
      "a3_3": "Vay Nợ",
      "a5_1_1": "Mô tả sản phẩm",
      "a5_1_2": "MÃ SẢN PHẨM",
      "a5_2": "Chi lao động thuê",
      "a5_3": "Tiền thuê địa điểm",
      "a5_4": "Tiền điện, nước, nhiên liệu",
      "a5_5": "Chi phí nguyên liệu, vật liệu",
      "a5_6": "Tổng số tiền vốn bình quân một tháng",
      "a5_7": "Chi phí khác",
      "a5_8t": "Tổng chi phí",
      "a5_9": "Tiền Lãi",
      "a5_10": "Doanh thu",
      "a5_10t": "Tổng Doanh Thu",
      "a6_1_1_1": "1. Điện",
      "a6_1_2_1": "2. Than",
      "a6_1_3_1": "3. Xăng",
      "a6_1_4_1": "4. Dầu mazut (FO)",
      "a6_1_5_1": "5. Dầu diezel (DO)",
      "a6_1_6_1": "6. Dầu hỏa",
      "a6_1_7_1": "7. Dầu nhờn",
      "a6_1_8_1": "8. Dầu khác",
      "a6_1_9_1": "9. LPG (Gas, ...)",
      "a6_1_10_1": "10. Khí sinh học (Biogas,..)",
      "a6_1_11_1": "11. Khác: rác thải, trấu, bã mía,…",
      "a7_1": "sử dụng internet",
      "a7_2": "cơ sở có bán hàng qua Internet",
      "a7_3": "tỷ trọng doanh thu qua Internet",
      "a7_4_1": "Mua, thuê phần cứng",
      "a7_4_1_2": "Số tiền đã chi phần cứng",
      "a7_4_2_2": "Số tiền đã chi phần mềm",
      "a7_4_3_2": "Số tiền đã chi khác CNTT"
    };

    const idCols = ["matkcs", "maxa", "madiaban", "idcoso"];

    const updated = customColConfigs.map(cfg => {
      const origLower = cfg.originalName.toLowerCase().trim();
      
      // Khớp từ điển mẫu dán nhãn
      if (dict[origLower]) {
        return {
          ...cfg,
          use: true,
          newName: dict[origLower]
        };
      }
      
      // Giữ nguyên các cột định danh cốt lõi
      if (idCols.includes(origLower)) {
        return {
          ...cfg,
          use: true,
          newName: cfg.originalName // Giữ nguyên chữ hoa thường gốc của cột định danh
        };
      }

      // Còn lại loại bỏ để sạch bảng theo đúng yêu cầu người dùng
      return {
        ...cfg,
        use: false,
        newName: ""
      };
    });

    setCustomColConfigs(updated);
    alert("Đã tự động điền dán nhãn tiếng Việt chuẩn A1-A7, giữ lại các định danh gốc. 'Tổng lao động' sẽ được tự động tính toán từ LĐ Thuê + LĐ Gia Đình khi bạn bấm áp dụng tái cấu trúc!");
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

    // Tạo bảng chứa dữ liệu mới gồm các cột được chọn và tên cột mới bằng chu trình chunkProcess an toàn
    const restructuredRows = await chunkProcess(
      rawImportedData,
      10000,
      row => {
        const newRow: any = {};
        activeConfigs.forEach(cfg => {
          const val = row[cfg.originalName];
          newRow[cfg.newName.trim()] = val !== undefined && val !== null ? val : "";
        });

        // Tự động tính bổ sung 'Tổng lao động' = Số lao động thuê (A2_1 / LĐ Thuê) + Số lao động gia đình (A2_2 / LĐ gia đình)
        let val1: any = undefined;
        let val2: any = undefined;

        // Quét tìm giá trị của A2_1 và A2_2 một cách thông minh (qua cả tên cũ và tên mới)
        activeConfigs.forEach(cfg => {
          const origL = cfg.originalName.toLowerCase().trim();
          const newTL = cfg.newName.toLowerCase().trim();
          if (origL === "a2_1" || newTL === "ld thuê" || newTL === "ld thue" || newTL === "lao động thuê" || newTL === "lao dong thue") {
            val1 = row[cfg.originalName];
          }
          if (origL === "a2_2" || newTL === "ld gia đình" || newTL === "ld gia dinh" || newTL === "lao động gia đình" || newTL === "lao dong gia dinh") {
            val2 = row[cfg.originalName];
          }
        });

        if (val1 !== undefined || val2 !== undefined) {
          const num1 = parseRobustNumber(val1);
          const num2 = parseRobustNumber(val2);
          newRow["Tổng lao động"] = num1 + num2;
        }

        return newRow;
      },
      pct => {
        setProgress(10 + Math.round(pct * 0.8));
        setStatusMessage(`Đang tái cấu trúc bảng: đã hoàn thành ${pct}%...`);
      }
    );

    // Cập nhật cấu hình mapping bảo toàn theo chỉ định người dùng khi đổi tên cột
    const newMapping: ColumnMapping = {
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: ""
    };

    // 1. Đồng bộ từ các role trực tiếp được khai báo trên các cột đang sử dụng
    activeConfigs.forEach(cfg => {
      if (cfg.role && cfg.role in newMapping) {
        newMapping[cfg.role as keyof ColumnMapping] = cfg.newName.trim();
      }
    });

    // 2. Dự phòng: Nếu vai trò nào chưa được gán bằng cột mới, ta kiểm tra vai trò đó trong mapping cũ có liên kết cột gốc nào không
    Object.keys(mapping).forEach((roleKey) => {
      const key = roleKey as keyof ColumnMapping;
      if (!newMapping[key]) {
        const oldMappedCol = mapping[key];
        if (oldMappedCol) {
          const config = activeConfigs.find(cfg => cfg.originalName === oldMappedCol);
          if (config) {
            newMapping[key] = config.newName.trim();
          }
        }
      }
    });

    // 3. Dự phòng các cột chính theo độ trùng khớp tương đối nếu còn sót vai trò cốt lõi
    const currentNewCols = activeConfigs.map(c => c.newName.trim());
    if (!newMapping.mota) {
      const found = currentNewCols.find(c => /mô tả|mota|nội dung|hoạt động|tên ngành/i.test(c));
      if (found) newMapping.mota = found;
    }
    if (!newMapping.manganh) {
      const found = currentNewCols.find(c => /mã ngành|manganh|vsic|mã nghe|mã nghề/i.test(c));
      if (found) newMapping.manganh = found;
    }
    if (!newMapping.xa) {
      const found = currentNewCols.find(c => /địa bàn|xã|phường|thị trấn|diaban|xa/i.test(c));
      if (found) newMapping.xa = found;
    }
    if (!newMapping.doanhthu) {
      const found = currentNewCols.find(c => /doanh thu|doanhthu|doanh số|thu nhập/i.test(c));
      if (found) newMapping.doanhthu = found;
    }
    if (!newMapping.laodong) {
      const found = currentNewCols.find(c => /lao động|laodong|số người|nhân sự/i.test(c));
      if (found) newMapping.laodong = found;
    }

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

  // Thực hiện phép tính toán học hoặc ghép chữ giữa các cột (Column to Column calculations)
  const handleCalculateColumn = async () => {
    if (rawImportedData.length === 0) {
      alert("Chưa có dữ liệu nguồn chính! Hãy tải tệp Excel lên trước.");
      return;
    }
    const safeNewName = calcColName.trim();
    if (!safeNewName) {
      alert("Vui lòng nhập tên cho cột kết quả mới!");
      return;
    }
    if (!calcCol1) {
      alert("Vui lòng chọn cột thành phần thứ nhất!");
      return;
    }
    if (calcType === "column" && !calcCol2) {
      alert("Vui lòng chọn cột thành phần thứ hai!");
      return;
    }
    if (calcType === "constant" && calcConstant.trim() === "") {
      alert("Vui lòng nhập giá trị hằng số để thực hiện phép tính!");
      return;
    }

    setLoading(true);
    setProgress(20);
    setStatusMessage(`Đang tính toán tạo cột mới [${safeNewName}]...`);
    await sleep(200);

    try {
      const computeRowValue = (row: any) => {
        const val1 = row[calcCol1];
        let val2: any;
        if (calcType === "column") {
          val2 = row[calcCol2];
        } else {
          const rawConst = calcConstant.trim();
          val2 = isNaN(Number(rawConst)) ? rawConst : Number(rawConst);
        }

        if (calcOperator === "concat") {
          const str1 = val1 !== undefined && val1 !== null ? String(val1) : "";
          const str2 = val2 !== undefined && val2 !== null ? String(val2) : "";
          return `${str1} ${str2}`.trim();
        } else {
          const num1 = parseRobustNumber(val1);
          const num2 = parseRobustNumber(val2);
          
          let resultNum = 0;
          if (calcOperator === "+") resultNum = num1 + num2;
          else if (calcOperator === "-") resultNum = num1 - num2;
          else if (calcOperator === "*") resultNum = num1 * num2;
          else if (calcOperator === "/") {
            resultNum = num2 !== 0 ? num1 / num2 : 0;
          }

          if (calcRounding === "int") {
            return Math.round(resultNum);
          } else if (calcRounding === "1dec") {
            return Math.round(resultNum * 10) / 10;
          } else if (calcRounding === "2dec") {
            return Math.round(resultNum * 100) / 100;
          }
          return resultNum;
        }
      };

      const computedRaw = await chunkProcess<any, any>(
        rawImportedData,
        10000,
        (row: any) => ({
          ...row,
          [safeNewName]: computeRowValue(row)
        }),
        pct => {
          setProgress(20 + Math.round(pct * 0.4));
          setStatusMessage(`Đang tính định mức cột gốc: ${pct}%...`);
        }
      );

      const computedMain = await chunkProcess<any, any>(
        mainData,
        10000,
        (row: any) => ({
          ...row,
          [safeNewName]: computeRowValue(row)
        }),
        pct => {
          setProgress(60 + Math.round(pct * 0.4));
          setStatusMessage(`Đang tính định mức cột hiển thị: ${pct}%...`);
        }
      );

      const newCols = [...columns];
      if (!newCols.includes(safeNewName)) {
        newCols.push(safeNewName);
      }

      let updatedColConfigs = [...customColConfigs];
      const configExists = updatedColConfigs.some(cfg => cfg.originalName === safeNewName);
      if (!configExists) {
        updatedColConfigs.push({
          originalName: safeNewName,
          use: true,
          newName: safeNewName,
          role: ""
        });
      } else {
        updatedColConfigs = updatedColConfigs.map(cfg => {
          if (cfg.originalName === safeNewName) {
            return { ...cfg, use: true, newName: safeNewName };
          }
          return cfg;
        });
      }

      setRawImportedData(computedRaw);
      setMainData(computedMain);
      setColumns(newCols);
      setCustomColConfigs(updatedColConfigs);
      
      setCalcColName("");
      setProgress(100);
      setStatusMessage(`Đã tính toán thành công và bổ sung cột [${safeNewName}]!`);
      
      autoSaveSession(computedMain, computedRaw, newCols, fileName, mapping, updatedColConfigs);
    } catch (err: any) {
      alert("Lỗi tính toán cột: " + err.message);
    } finally {
      await sleep(300);
      setLoading(false);
    }
  };

  // Bộ lọc dữ liệu viewer
  const filteredData = useMemo(() => {
    if (!searchTerm) return mainData;
    const term = searchTerm.toLowerCase();
    return mainData.filter(row => {
      return Object.values(row).some(val => String(val).toLowerCase().includes(term));
    });
  }, [mainData, searchTerm]);

  // Bộ lọc dữ liệu logic cho Tab Kiểm tra Logic
  const filteredLogicData = useMemo(() => {
    if (!mainData || mainData.length === 0) return [];
    // Chỉ hiển thị những hộ có thuộc tính kiểm chứng logic (_satisfiesIf hoặc _violated)
    const hasBeenScanned = mainData.some(row => "_satisfiesIf" in row || "_violated" in row);
    if (!hasBeenScanned) return [];

    if (logicFilterMode === "if_satisfied") {
      return mainData.filter(row => row._satisfiesIf === true);
    }
    if (logicFilterMode === "violated") {
      return mainData.filter(row => row._violated === true);
    }
    // Chế độ "all" - Chỉ hiện tất cả các hộ có liên quan đã quét lọt vào quy tắc (thỏa mãn NẾU hoặc bị lỗi) để tránh hiện cả bảng tính gốc khổng lồ
    return mainData.filter(row => row._satisfiesIf === true || row._violated === true);
  }, [mainData, logicFilterMode]);

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

    // 2. CÙNG MÃ -> KHÁC MÔ TẢ CHI TIẾT (Yêu cầu người dùng tích hợp gom mã ngành giống nhau nhưng mô tả khác nhau)
    const codeMap = new Map<string, Array<{ desc: string; rowIdx: number; row: any }>>();
    mainData.forEach((row, idx) => {
      if (!row || typeof row !== 'object') return;
      const code = normalizeSectorCode(row[targetManganh]);
      if (!code) return;
      const rawMota = String(row[targetMota] || "").trim();
      const cleanMota = rawMota.toLowerCase().replace(/\s+/g, " ");
      if (!cleanMota) return;

      if (!codeMap.has(code)) {
        codeMap.set(code, []);
      }
      codeMap.get(code)!.push({ desc: rawMota, rowIdx: idx, row });
    });

    const codeToDescs: Array<{
      codeValue: string;
      occurrences: number;
      descriptions: Array<{ desc: string; count: number; rows: number[] }>;
    }> = [];

    codeMap.forEach((occurrences, code) => {
      if (occurrences.length <= 1) return;

      const descCounts = new Map<string, { originalText: string; rowIdxSelection: number[] }>();
      occurrences.forEach(occ => {
        const key = occ.desc.toLowerCase().replace(/\s+/g, " ");
        if (!descCounts.has(key)) {
          descCounts.set(key, { originalText: occ.desc, rowIdxSelection: [] });
        }
        descCounts.get(key)!.rowIdxSelection.push(occ.rowIdx);
      });

      if (descCounts.size > 1) {
        if (codeToDescs.length < 2000) {
          codeToDescs.push({
            codeValue: code,
            occurrences: occurrences.length,
            descriptions: Array.from(descCounts.entries()).map(([key, item]) => ({
              desc: item.originalText,
              count: item.rowIdxSelection.length,
              rows: item.rowIdxSelection
            }))
          });
        }
      }
    });

    // Sắp xếp giảm dần theo mức độ phổ biến / mâu thuẫn để rà quét các lỗi nghiêm trọng nhất lên đầu
    descToCodes.sort((a, b) => b.occurrences - a.occurrences);
    codeToDescs.sort((a, b) => b.occurrences - a.occurrences);

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
    // Don't force redirect, render table inline directly!
    // setActiveTab("xemdulieu");
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
    // Don't force redirect, render table inline directly!
    // setActiveTab("xemdulieu");
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
    return parseRobustNumber(val);
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
      if (isSummaryRow(row)) return;
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
              sec1Code = getParentSectorCode(mng) || "";
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
    // Don't force redirect, render table inline directly!
    // setActiveTab("xemdulieu");
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
      if (isSummaryRow(row)) return;
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
            finalCode = getParentSectorCode(normalized) || "";
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
            sum += parseRobustNumber(rawNum);
            validCount++;
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
            sumAll += parseRobustNumber(rawVal);
            rowValidCount++;
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
  const handleStandardizeSectorsAndMatch = async (optIndustryCol?: string, optDescriptionCol?: string) => {
    if (mainData.length === 0) {
      alert("Không tìm thấy dữ liệu nguồn chính! Vui lòng nạp tệp chính trước.");
      return;
    }
    const targetIndustryCol = optIndustryCol || stdIndustryCol;
    const targetDescriptionCol = optDescriptionCol || stdDescriptionCol;

    if (!targetIndustryCol) {
      alert("Vui lòng chọn cột chứa Mã ngành cấp 5 cần chuẩn hóa!");
      return;
    }

    if (optIndustryCol) setStdIndustryCol(optIndustryCol);
    if (optDescriptionCol) setStdDescriptionCol(optDescriptionCol);

    setLoading(true);
    setProgress(5);
    setStatusMessage("Đang quét danh sách mã ngành và tiến hành chuẩn hóa mẫu tự liên hợp...");
    await sleep(200);

    try {
      let validCount = 0;
      let invalidCount = 0;
      let conflictCount = 0;
      const anomalies: any[] = [];

      // Tạo mảng bản ghi mới bổ sung cột của "Tên Ngành Chuẩn VSIC" và "Trạng Thái Đối Chiếu VSIC" bằng chunkProcess
      const updatedRows = await chunkProcess(
        mainData,
        5000,
        (row, idx) => {
          if (!row || typeof row !== 'object') return row;
          const rawCode = row[targetIndustryCol];
          const rawDesc = targetDescriptionCol ? String(row[targetDescriptionCol] || "") : "";
          
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
          }

          if (anomalyReason && anomalies.length < 5000) {
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
            if (key === targetDescriptionCol) {
              flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ KHÔNG TÌM THẤY MÃ TRONG VSIC";
              flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
            }
          });

          // Nếu không khớp được vị trí cột mô tả thì tự chêm cột mới vào kế cột Mã ngành
          if (flexRow["Tên Ngành Chuẩn VSIC"] === undefined) {
            Object.keys(row).forEach(key => {
              flexRow[key] = row[key];
              if (key === targetIndustryCol) {
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
        },
        pct => {
          setProgress(5 + Math.round(pct * 0.9));
          setStatusMessage(`Đang chuẩn hóa & phân tích mã ngành VSIC: ${pct}%...`);
        }
      );

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
      
      alert(`Chuẩn hóa hoàn tất!\n- Tổng cộng: ${updatedRows.length} dòng\n- Khớp VSIC: ${validCount} dòng\n- Lệch chuẩn: ${invalidCount} dòng.`);
    } catch (err: any) {
      alert("Lỗi quá trình chuẩn hóa VSIC: " + err.message);
      setLoading(false);
    }
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

      const updatedRows = await chunkProcess(
        mainData,
        5000,
        (row, idx) => {
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
            const cleanText = (s: string) => {
              if (!s) return "";
              let clean = s.toString().toLowerCase().trim();
              
              // Chuẩn hóa ký tự có dấu tiếng Việt về không dấu
              clean = clean.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
              clean = clean.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
              clean = clean.replace(/ì|í|ị|ỉ|ĩ/g, "i");
              clean = clean.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
              clean = clean.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
              clean = clean.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
              clean = clean.replace(/đ/g, "d");
              
              // Chuyển đổi các từ viết tắt phổ biến thường gặp của điều tra viên
              clean = clean.replace(/\bsx\b/g, "san xuat");
              clean = clean.replace(/\bkd\b/g, "kinh doanh");
              clean = clean.replace(/\btm\b/g, "thuong mai");
              clean = clean.replace(/\bdv\b/g, "dich vu");
              clean = clean.replace(/\bbl\b/g, "ban le");
              clean = clean.replace(/\bbb\b/g, "ban buon");
              
              // Loại bỏ ký tự đặc biệt, chỉ giữ lại chữ cái và số
              clean = clean.replace(/[^a-z0-9\s]/g, " ");
              return clean.replace(/\s+/g, " ").trim();
            };

            const normA = cleanText(valA);
            const normB = cleanText(valB);

            if (normA && normB && (normA.includes(normB) || normB.includes(normA))) {
              isMatch = true;
              explanation = `Thỏa mãn: Chứa chuỗi ký tự của nhau (sau khi chuẩn hóa không dấu)`;
            } else {
              // Phân tách thành tập hợp từ khóa (bỏ qua liên từ & từ đính kèm vô hại)
              const stopWords = ["va", "cac", "cua", "hoac", "cho", "doi", "voi", "nhu", "nhung", "mot", "bi", "o", "tai"];
              const wordsA = normA.split(" ").filter(w => w.length > 1 && !stopWords.includes(w));
              const wordsB = normB.split(" ").filter(w => w.length > 1 && !stopWords.includes(w));

              if (wordsA.length > 0 && wordsB.length > 0) {
                const setA = new Set(wordsA);
                const setB = new Set(wordsB);
                
                // Đếm số từ trùng khớp song phương
                const overlapAInB = wordsA.filter(w => setB.has(w)).length;
                const overlapBInA = wordsB.filter(w => setA.has(w)).length;

                const minLen = Math.min(wordsA.length, wordsB.length);
                const maxOverlap = Math.max(overlapAInB, overlapBInA);

                // Điều kiện khớp từ khóa thông minh: Trùng tuột tất cả từ ở chuỗi ngắn, hoặc đạt tỷ lệ cao >= 75%
                if (maxOverlap >= minLen || (minLen > 2 && maxOverlap >= minLen - 1) || (maxOverlap / minLen >= 0.75)) {
                  isMatch = true;
                  explanation = `Thỏa mãn: Khớp từ khóa cốt lõi thông minh (Trùng ${maxOverlap}/${minLen} từ chính)`;
                } else {
                  explanation = `Thực sự lệch: Không tìm thấy cụm từ khóa liên khớp (Chỉ trùng ${maxOverlap}/${minLen} từ chính)`;
                }
              } else {
                explanation = `Thực sự lệch: Trống hoặc không phân giải được từ khóa chính để đối so`;
              }
            }
          } else if (crossCompareRule === "semantic") {
            const cleanStr = (s: string) => {
              if (!s) return "";
              let clean = s.toString().toLowerCase().trim();
              
              clean = clean.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
              clean = clean.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
              clean = clean.replace(/ì|í|ị|ỉ|ĩ/g, "i");
              clean = clean.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
              clean = clean.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
              clean = clean.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
              clean = clean.replace(/đ/g, "d");
              
              clean = clean.replace(/\bsx\b/g, "san xuat");
              clean = clean.replace(/\bkd\b/g, "kinh doanh");
              clean = clean.replace(/\btm\b/g, "thuong mai");
              clean = clean.replace(/\bdv\b/g, "dich vu");
              clean = clean.replace(/\bbl\b/g, "ban le");
              clean = clean.replace(/\bbb\b/g, "ban buon");
              
              clean = clean.replace(/[^a-z0-9\s]/g, " ");
              return clean.replace(/\s+/g, " ").trim();
            };

            const normA = cleanStr(valA);
            const normB = cleanStr(valB);

            if (normA && normB && (normA === normB || normA.includes(normB) || normB.includes(normA))) {
              isMatch = true;
              explanation = "Khớp chuỗi lý tưởng: Chứa cụm từ của nhau sau khi lược dấu";
            } else {
              // 1. Phân nhóm Hoạt động (Động từ)
              const isMfgA = /\b(may|mac|theu|khau|det|san xuat|gia cong|che bien|lam|che tao|che banh|lap rap)\b/.test(normA);
              const isMfgB = /\b(may|mac|theu|khau|det|san xuat|gia cong|che bien|lam|che tao|che banh|lap rap)\b/.test(normB);

              const isCommA = /\b(ban|mua|ban le|ban buon|kinh doanh|thuong mai|mua ban|phan phoi|cung cap|dai ly|cua hang|ky gui)\b/.test(normA);
              const isCommB = /\b(ban|mua|ban le|ban buon|kinh doanh|thuong mai|mua ban|phan phoi|cung cap|dai ly|cua hang|ky gui)\b/.test(normB);

              const isFoodA = /\b(an|uong|phuc vu|quan an|nha hang|giai khat|nuoc giai khat|ca phe|cafe|com|pho|bun|lau|nuong|banh cuon|banh mi)\b/.test(normA);
              const isFoodB = /\b(an|uong|phuc vu|quan an|nha hang|giai khat|nuoc giai khat|ca phe|cafe|com|pho|bun|lau|nuong|banh cuon|banh mi)\b/.test(normB);

              const isConstA = /\b(xay dung|thi cong|lap dat|hoan thien|lam|trung tu|lap|op|lat|ba|son|cau thang|canh cong|khung nhom|tran nhua)\b/.test(normA);
              const isConstB = /\b(xay dung|thi cong|lap dat|hoan thien|lam|trung tu|lap|op|lat|ba|son|cau thang|canh cong|khung nhom|tran nhua)\b/.test(normB);

              const isRepairA = /\b(sua|sua chua|bao duong|bao tri|trung tu)\b/.test(normA);
              const isRepairB = /\b(sua|sua chua|bao duong|bao tri|trung tu)\b/.test(normB);

              // 2. Định nghĩa các nhóm danh từ chuyên đề đồng nghĩa (Thematic Nouns)
              const nounGroups = [
                {
                  id: "garment",
                  synonyms: ["ao", "quan", "vay", "dam", "mac", "may", "trang phuc", "ao dai"]
                },
                {
                  id: "food",
                  synonyms: ["banh", "banh day", "banh cuon", "banh mi", "bun", "pho", "com", "xoi", "gio", "cha", "thuc pham", "bot", "gao", "luong thuc", "an uong"]
                },
                {
                  id: "vehicle",
                  synonyms: ["xe dap", "xe may", "o to", "xe con", "phu tung", "lop xe", "sam xe", "ruot xe", "ve xe", "yen xe"]
                },
                {
                  id: "electronics",
                  synonyms: ["tivi", "ti vi", "dien thoai", "may tinh", "nghe nhin", "dien tu", "gia dung", "do dien", "am thanh", "loa", "dai", "am ly", "tu lanh", "may giat", "dieu hoa"]
                },
                {
                  id: "construction",
                  synonyms: ["nhom", "cua", "kinh", "cua kinh", "khung nhom", "canh cong", "cau thang", "sat", "thep", "kim loai", "tran nhua", "nhua", "thiet bi lap dat", "ton", "ngoi", "vua", "xi mang"]
                },
                {
                  id: "agriculture",
                  synonyms: ["rau", "qua", "trai cay", "nong san", "heo", "lon", "ga", "vit", "bo", "trau", "gia suc", "gia cam", "giet mo", "mo", "thit"]
                },
              ];

              let matchedId = null;
              for (const group of nounGroups) {
                const hasA = group.synonyms.some(s => normA.includes(s) || s.includes(normA) || normA.split(" ").includes(s));
                const hasB = group.synonyms.some(s => normB.includes(s) || s.includes(normB) || normB.split(" ").includes(s));
                if (hasA && hasB) {
                  matchedId = group.id;
                  break;
                }
              }

              if (matchedId) {
                if (matchedId === "garment" && (isMfgA || isMfgB) && (isMfgA || isMfgB || normB.includes("trang phuc"))) {
                  isMatch = true;
                  const customSecName = vsicRawData["14"] || vsicRawData["141"] || "Hoạt động May mặc / Sản xuất trang phục";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                } else if (matchedId === "food" && (isMfgA || isFoodA || isMfgB || isFoodB)) {
                  isMatch = true;
                  const customSecName = vsicRawData["10"] || vsicRawData["56"] || "Chế biến thực phẩm, bánh từ bột hoặc dịch vụ ăn uống";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                } else if (matchedId === "vehicle" && (isRepairA || isCommA) && (isRepairB || isCommB)) {
                  isMatch = true;
                  const customSecName = vsicRawData["4540"] || vsicRawData["454"] || vsicRawData["45"] || "Sửa chữa, bảo bảo dưỡng hoặc mua bán xe đạp, xe máy";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                } else if (matchedId === "electronics" && (isRepairA || isCommA || normA.includes("do dien")) && (isRepairB || isCommB || normB.includes("thiet bi"))) {
                  isMatch = true;
                  const customSecName = vsicRawData["95210"] || vsicRawData["952"] || vsicRawData["95"] || "Sửa chữa hoặc dịch vụ thương mại thiết bị điện tử gia dụng";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                } else if (matchedId === "construction" && (isMfgA || isConstA || isCommA) && (isMfgB || isConstB || isCommB)) {
                  isMatch = true;
                  const customSecName = vsicRawData["41"] || vsicRawData["43"] || "Thi công, lắp đặt vật tư xây dựng hoặc sản xuất cấu kiện kim loại";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                } else if (matchedId === "agriculture" && (normA.includes("giet") || normA.includes("mo") || isCommA) && (normB.includes("giet") || normB.includes("mo") || isCommB)) {
                  isMatch = true;
                  const customSecName = vsicRawData["1010"] || vsicRawData["01"] || "Giết mổ gia súc, gia cầm hoặc bán lẻ rau quả nông sản";
                  explanation = `Khớp nghĩa VSIC: ${customSecName}`;
                }
              }

              if (!isMatch) {
                // Fuzzy fallback if high similarity of non-stopwords
                const stopWords = ["va", "cac", "cua", "hoac", "cho", "doi", "voi", "nhu", "nhung", "mot", "bi", "o", "tai"];
                const wordsA = normA.split(" ").filter(w => w.length > 1 && !stopWords.includes(w));
                const wordsB = normB.split(" ").filter(w => w.length > 1 && !stopWords.includes(w));
                
                if (wordsA.length > 0 && wordsB.length > 0) {
                  const setA = new Set(wordsA);
                  const setB = new Set(wordsB);
                  const overlap = wordsA.filter(w => setB.has(w)).length;
                  const minLen = Math.min(wordsA.length, wordsB.length);
                  
                  if (overlap / minLen >= 0.4) {
                    isMatch = true;
                    explanation = `Đồng nhất mức cao: Trùng khớp từ khóa chủ đạo (${overlap}/${minLen} từ chính)`;
                  } else {
                    explanation = "Thực sự lệch: Khái niệm khác biệt hoàn toàn hoặc không liên đới từ đồng nghĩa";
                  }
                } else {
                  explanation = "Thực sự lệch: Không thể phân tích cấu trúc từ ngữ";
                }
              }
            }
          }

          if (isMatch) {
            matchCount++;
          } else {
            mismatchCount++;
            anomalies.push({
              _rowIdx: idx + 1,
              valA,
              valB,
              explanation
            });
          }

          return {
            ...row,
            _crossCompareMatch: isMatch ? "Trùng khớp" : "Lệch biệt",
            _crossCompareExplanation: explanation
          };
        },
        pct => {
          setProgress(15 + Math.round(pct * 0.8));
          setStatusMessage(`Đang đối chiếu song song: ${pct}%...`);
        }
      );

      const newCols = Object.keys(updatedRows[0] || {});
      setMainData(updatedRows);
      setColumns(newCols);
      setCrossCompareAnomalies(anomalies);
      setCrossCompareStats({
        total: updatedRows.length,
        matchCount,
        mismatchCount
      });

      // Tự sao lưu vĩnh viễn vào hệ thống
      autoSaveSession(updatedRows, rawImportedData, newCols, fileName, mapping, customColConfigs);

      setProgress(100);
      setStatusMessage("Đối chiếu song song hoàn tất!");
      await sleep(350);
      setLoading(false);

      alert(`Đối chiếu hoàn tất!\n- Trùng khớp: ${matchCount} dòng\n- Lệch biệt: ${mismatchCount} dòng.`);
    } catch (err: any) {
      alert("Lỗi đối chiếu song song: " + err.message);
      setLoading(false);
    }
  };

  // 5. CHỨC NĂNG BÁO CÁO NHANH THEO PHÂN CẤP NGÀNH & XÃ CHUẨN XÁC
  const handleQuickReport = async (level: number, optManganh?: string, optXa?: string, optDoanhThu?: string, optLaoDong?: string) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi chạy báo cáo nhanh.");
      return;
    }
    const targetManganh = optManganh || quickReportManganhCol || mapping.manganh;
    const targetXa = optXa || quickReportXaCol || mapping.xa;
    const targetDoanhThu = optDoanhThu || quickReportDoanhThuCol || mapping.doanhthu;
    const targetLaoDong = optLaoDong || quickReportLaoDongCol || mapping.laodong;

    if (!targetManganh) {
      alert("Vui lòng chỉ định cột chứa Mã ngành hoặc tiêu chí phân loại gộp ở bộ chọn!");
      return;
    }
    if (!targetXa) {
      alert("Vui lòng chỉ định cột chứa Xã / Địa bàn ở bộ chọn!");
      return;
    }

    // Xây dựng danh sách chỉ tiêu cộng dồn động (không khoá cứng cột)
    const sumCols: string[] = [];
    quickReportSumCols.forEach(col => {
      if (col && columns.includes(col)) {
        sumCols.push(col);
      }
    });

    // Nếu không cấu hình chỉ tiêu phụ động, tự chuyển về tương thích ngược dựa vào lựa chọn Doanh Thu và Lao Động
    if (sumCols.length === 0) {
      if (targetDoanhThu && columns.includes(targetDoanhThu)) sumCols.push(targetDoanhThu);
      if (targetLaoDong && columns.includes(targetLaoDong)) sumCols.push(targetLaoDong);
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang tạo báo cáo tổng hợp gộp nhóm...`);
    await sleep(200);

    try {
      const processedData = await chunkProcess(
        mainData,
        10000,
        row => {
          if (!row || typeof row !== 'object') {
            return {
              _temNganhCap: "Chưa xác định",
              _tempXa: "Khác"
            };
          }
          
          let tenNganhLabel = "";
          if (level === 0) {
            // Gom nhóm trực tiếp bằng nội dung chuỗi gốc trong cột, không tra cứu bảng VSIC (Phù hợp mọi cuộc điều tra dấn số/nông nghiệp/địa bàn bất kỳ)
            tenNganhLabel = String(row[targetManganh] || "Chưa xác định / Bỏ trống").trim();
          } else {
            const mng = normalizeSectorCode(row[targetManganh]);
            if (level === 2) {
              const sec2Code = mng ? mng.slice(0, 2) : "";
              const sec2Name = vsicRawData[sec2Code] || "Ngành cấp 2 chưa định nghĩa";
              tenNganhLabel = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định - Ngành cấp 2 chưa định nghĩa";
            } else {
              let sec1Code = "";
              if (mng) {
                if (/^[a-zA-Z]$/.test(mng)) {
                  sec1Code = mng.toUpperCase();
                } else {
                  sec1Code = getParentSectorCode(mng) || "";
                }
              }
              const sec1Name = vsicRawData[sec1Code] || "Ngành cấp 1 chưa định nghĩa";
              tenNganhLabel = sec1Code ? `${sec1Code} - ${sec1Name}` : "Chưa xác định - Ngành cấp 1 chưa định nghĩa";
            }
          }

          return {
            ...row,
            _temNganhCap: tenNganhLabel,
            _tempXa: String(row[targetXa] || "Khác").trim()
          };
        },
        pct => {
          setProgress(Math.round(pct * 0.4));
          setStatusMessage(`Đang chuẩn bị phân tích dữ liệu gộp địa bàn/chỉ tiêu: ${pct}%...`);
        }
      );

      let finalReportRows: any[] = [];

      if (reportType === "pivot") {
        setStatusMessage("Đang xoay gom nhóm Pivot theo cột...");
        await sleep(150);

        const communes = Array.from(new Set(processedData.map(r => r._tempXa))).sort();
        const sectorLabels = Array.from(new Set(processedData.map(r => r._temNganhCap))).sort();

        // Gom trước các dòng theo cặp Xã và Phân nhóm
        const groupedMap = new Map<string, any[]>();
        processedData.forEach(r => {
          const key = `${r._tempXa || ""}||${r._temNganhCap || ""}`;
          let list = groupedMap.get(key);
          if (!list) {
            list = [];
            groupedMap.set(key, list);
          }
          list.push(r);
        });

        communes.forEach((commune) => {
          const communeObj: any = {
            "Địa_Bàn_Xã": commune
          };

          let totalCommuneDN = 0;
          const totalAccumulate: { [col: string]: number } = {};
          sumCols.forEach(col => {
            totalAccumulate[col] = 0;
          });

          sectorLabels.forEach(sector => {
            const matchedRows = groupedMap.get(`${commune}||${sector}`) || [];
            
            // Tính tổng từng chỉ tiêu được chọn
            const columnSums: { [col: string]: number } = {};
            sumCols.forEach(col => {
              columnSums[col] = 0;
            });

            matchedRows.forEach(r => {
              sumCols.forEach(col => {
                const val = parseRobustNumber(r[col]);
                columnSums[col] += val;
              });
            });

            // Tạo các cột xoay động
            sumCols.forEach(col => {
              communeObj[`${sector} - Tổng ${col}`] = Math.round(columnSums[col] * 100) / 100;
              totalAccumulate[col] += columnSums[col];
            });

            totalCommuneDN += matchedRows.length;
          });

          communeObj["Số_Dòng_Tổng_Hợp"] = totalCommuneDN;
          sumCols.forEach(col => {
            communeObj[`Tổng_Cộng_${col}_Toàn_Xã`] = Math.round(totalAccumulate[col] * 100) / 100;
          });

          finalReportRows.push(communeObj);
        });
      } else {
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

          const rowObj: any = {};
          if (level === 0) {
            rowObj["Nhóm_Phân_Loại"] = dims.Ngành;
          } else {
            rowObj[`Ngành_Cấp_${level}`] = dims.Ngành;
          }
          rowObj["Địa_Bàn_Xã"] = dims.Xã;
          rowObj["Số_Lượng_Bản_Ghi"] = rowsObj.length;

          sumCols.forEach(col => {
            let sumCol = 0;
            rowsObj.forEach(r => {
              sumCol += parseRobustNumber(r[col]);
            });
            rowObj[`Tổng_${col}`] = Math.round(sumCol * 100) / 100;
          });

          finalReportRows.push(rowObj);
        });
      }

      setQuickReportResultRows(finalReportRows);
      setQuickReportResultCols(Object.keys(finalReportRows[0] || {}));
      setQuickReportLevel(level);

      setProgress(100);
      setStatusMessage(`Tạo báo cáo tổng hợp gộp thành công!`);
      await sleep(350);
      setLoading(false);
      
      alert("Tạo báo cáo tổng hợp hoàn tất! Dữ liệu đã được nạp gọn gàng và hiển thị bảng báo cáo kết xuất.");
    } catch (err: any) {
      alert("Lỗi quá trình tạo báo cáo nhanh: " + err.message);
      setLoading(false);
    }
  };

  const handleSchemaAiAutopilot = async (customCommand?: string) => {
    if (rawImportedData.length === 0) {
      alert("Vui lòng tải tệp Excel hoặc dữ liệu gốc chính lên trước khi thực hiện!");
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần cài đặt của AI Studio hoặc trong tệp .env để sử dụng Trợ lý AI Định nghĩa!");
      return;
    }

    setIsSchemaAiRunning(true);
    setSchemaAiLogs(["🔍 Đang khởi tạo Trợ lý AI định cấu hình bảng vạn năng...", "📋 Đọc dữ liệu mô hình các trường hiện có..."]);
    await sleep(400);

    try {
      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: { headers: { "User-Agent": "aistudio-build" } }
      });

      const userCmd = customCommand || schemaAiPrompt || "Hãy tự động Việt hóa các tên cột và chỉ định vai trò";
      
      setSchemaAiLogs(prev => [
        ...prev,
        `💬 Khẩu lệnh yêu cầu: "${userCmd}"`,
        `🧠 Đang gửi thông số cấu hình và nạp dữ liệu phân tích tới Gemini...`
      ]);

      const systemInstruction = `Bạn là một chuyên gia khoa học dữ liệu và kỹ sư phân tích nghiệp vụ (BA) tối tân.
Nhiệm vụ của bạn là phân tích danh sách các cột hiện tại của bảng tính dữ liệu:
[${columns.filter(c => !c.startsWith("_")).map(c => `'${c}'`).join(", ")}]

And yêu cầu thao tác của người dùng: "${userCmd}"

Hãy phân loại yêu cầu này thuộc về 1 trong 2 hành vi:
- "redefine": Định nghĩa lại tên cột, Việt hóa, dán nhãn vai trò tương thích giúp hệ thống dễ xử lý.
- "calculate": Tạo cột mới bằng phép tính toán học hoặc ghép chuỗi văn bản giữa các cột với nhau hoặc hằng số.

1. Nếu hành vi là "redefine" (hoặc mặc định nếu không có từ khóa phép tính toán rõ ràng):
   Hãy đề xuất Việt hóa tên cột thân thiện hơn (Ví dụ: MaNganh -> Mã Ngành ĐK, MoTa -> Mô Tả Hoạt Động, Xa -> Địa Bàn Xã Phường), chọn sử dụng hay không (mặc định luôn sử dụng true nếu thấy liên quan), và khớp vai trò tương thích giúp hệ thống (mota, manganh, xa, doanhthu, laodong, idCol):
   - "mota": cột liên quan đến mô tả ngành nghề, hoạt động kinh doanh (VD: MoTa, NganhNghe, Tên ngành).
   - "manganh": cột gồm mã số ngành kinh tế cấp 1-5 (VD: MaNganh, VSIC, MaNganhDTV).
   - "xa": cột địa bàn xã phường (VD: Xa, Phuong, DiaBan).
   - "doanhthu": cột số liệu doanh thu, doanh số, trị giá, thu nhập (VD: DoanhThu, DoanhSo, SanLuong, TriGia).
   - "laodong": cột quy mô lao động, số lượng người, nhân sự (VD: LaoDong, NhanSu, SoNguoi).
   - "idCol": cột mã số thuế, số thứ tự định danh doanh nghiệp duy nhất (VD: MaST, MST, ID).

   Trả về danh sách redefinitions dạng JSON. Các tên cột mới PHẢI viết bằng tiếng Việt có dấu đẹp đẽ, viết hoa chữ cái đầu tiên mỗi từ, sạch sẽ và ngắn gọn thích hợp làm tiêu đề bảng biểu.

2. Nếu hành vi là "calculate" (Người dùng yêu cầu tính toán như cộng, trừ, nhân, chia, ghép nối, phần trăm, VAT, trung bình, năng suất, thuế):
   Hãy khớp các cột cần tính toán từ danh sách thực tế của người dùng:
   - calcColName: Tên cột kết quả mới viết liền không dấu hoặc có dấu tiếng Việt thích hợp (Ví dụ: DoanhThuBinhQuan, ThueVAT, NangSuatLD, DiaBanGop).
   - calcCol1: Cột thích hợp thứ nhất (A) có trong danh sách gốc của người dùng.
   - calcOperator: Một trong các ký tự phép toán: "+", "-", "*", "/", "concat".
   - calcType: "column" (nếu đối tượng thứ hai là một cột khác) hoặc "constant" (nếu đối tượng thứ hai là một hằng số/chuỗi cố định).
   - calcCol2: Tên cột thứ hai (B) nếu calcType là "column".
   - calcConstant: Giá trị số hoặc chuỗi hằng số cố định (Ví dụ: "0.1", "1000000", "VND") nếu calcType là "constant".
   - calcRounding: Cách làm tròn phù hợp: "none", "int" (làm tròn số nguyên), "1dec" (1 số thập phân), "2dec" (2 số thập phân).

Trả về cấu trúc JSON duy nhất như sau, tuyệt đối không được thêm bất cứ bình luận, giải thích dông dài hay mã markdown bên ngoài khối JSON:
{
  "action": "redefine" | "calculate",
  "redefinitions": [
    {
      "originalName": "Tên gốc trong cột thực tế",
      "newName": "Tên tiếng Việt mới sạch sẽ",
      "use": true,
      "role": "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" | ""
    }
  ],
  "calculation": {
    "calcColName": "Tên cột kết quả mới",
    "calcCol1": "Cột thứ nhất A",
    "calcOperator": "+" | "-" | "*" | "/" | "concat",
    "calcType": "column" | "constant",
    "calcCol2": "Cột thứ hai B",
    "calcConstant": "giá trị hằng số",
    "calcRounding": "none" | "int" | "1dec" | "2dec"
  },
  "explanation": "Giải thích ngắn gọn súc tích bằng tiếng Việt vì sao chọn cách giải quyết này"
}`;

      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: `Danh sách các cột thực tế hiện tại của bảng tính: [${columns.join(", ")}]\nCâu lệnh: "${userCmd}"`,
        config: {
          systemInstruction,
          temperature: 0.1,
          responseMimeType: "application/json"
        }
      });

      const responseText = response.text || "";
      const parsed = JSON.parse(responseText.trim());

      setSchemaAiLogs(prev => [...prev, `⚡ Phản hồi từ Bộ não AI đã nạp! Bắt đầu phân cấu trúc...`]);
      await sleep(450);

      const action = parsed.action || "redefine";
      const explanation = parsed.explanation || "AI đã khớp nối hoàn tất dữ liệu cấu hình.";

      if (action === "redefine") {
        setSchemaAiLogs(prev => [...prev, `📝 Nhận diện thao tác: Việt hóa / Định nhãn vai trò cấu trúc cột...`]);
        await sleep(350);

        const newConfigs = (parsed.redefinitions || []).map((item: any) => ({
          originalName: item.originalName,
          newName: item.newName || item.originalName,
          use: item.use !== false,
          role: item.role || ""
        }));

        if (newConfigs.length > 0) {
          setCustomColConfigs(newConfigs);
          
          const foundMota = newConfigs.find((c: any) => c.role === "mota")?.originalName;
          const foundManganh = newConfigs.find((c: any) => c.role === "manganh")?.originalName;
          const foundXa = newConfigs.find((c: any) => c.role === "xa")?.originalName;
          const foundDoanhthu = newConfigs.find((c: any) => c.role === "doanhthu")?.originalName;
          const foundLaodong = newConfigs.find((c: any) => c.role === "laodong")?.originalName;

          setMapping(prev => ({
            ...prev,
            mota: foundMota || prev.mota,
            manganh: foundManganh || prev.manganh,
            xa: foundXa || prev.xa,
            doanhthu: foundDoanhthu || prev.doanhthu,
            laodong: foundLaodong || prev.laodong
          }));

          setSchemaAiLogs(prev => [
            ...prev,
            `✅ Đã tự động cấu bản ánh xạ vai trò:`,
            foundMota ? ` - Mô tả: ${foundMota}` : "",
            foundManganh ? ` - Mã ngành: ${foundManganh}` : "",
            foundXa ? ` - Địa bàn xã: ${foundXa}` : "",
            foundDoanhthu ? ` - Doanh thu: ${foundDoanhthu}` : "",
            foundLaodong ? ` - Lao động: ${foundLaodong}` : ""
          ].filter(Boolean));
        }

        setSchemaAiLogs(prev => [
          ...prev,
          `⚡ Thành công! Đã tự động cập nhật cấu hình hệ thống.`,
          `💬 Đánh giá từ AI: ${explanation}`
        ]);
        
        setProgress(100);
        setStatusMessage("Định nghĩa bảng dữ liệu thành công!");
      } else {
        setSchemaAiLogs(prev => [...prev, `🧮 Nhận diện thao tác: Tính toán / Ghép nối kết xuất cột mới...`]);
        await sleep(350);

        const calc = parsed.calculation || {};
        const calcColName = calc.calcColName || "CotMoi";
        const calcCol1 = calc.calcCol1;
        const calcOperator = calc.calcOperator || "+";
        const calcType = calc.calcType || "constant";
        const calcCol2 = calc.calcCol2;
        const calcConstant = calc.calcConstant;
        const calcRounding = calc.calcRounding || "none";

        if (!calcCol1) {
          throw new Error("Không tìm thấy cột thứ nhất dùng cho phép tính.");
        }

        const safeNewName = calcColName.trim();
        setSchemaAiLogs(prev => [
          ...prev,
          `📉 Đang thực thi phép tính: ${calcCol1} ${calcOperator} ${calcType === "column" ? calcCol2 : calcConstant} -> tạo cột [${safeNewName}]`,
          `⏳ Đang xử lý tính toán đồng bộ trên tệp dữ liệu lớn...`
        ]);

        const computeRowValue = (row: any) => {
          if (!row) return "";
          const val1Raw = String(row[calcCol1] || "");
          const val1Str = val1Raw.replace(/[^0-9.\-]/g, "");
          const val1 = parseFloat(val1Str);

          let val2 = 0;
          let val2Str = "";
          if (calcType === "column" && calcCol2) {
            val2Str = String(row[calcCol2] || "");
            val2 = parseFloat(val2Str.replace(/[^0-9.\-]/g, ""));
          } else if (calcConstant) {
            val2Str = String(calcConstant);
            val2 = parseFloat(val2Str);
          }

          if (calcOperator === "concat") {
            return (val1Raw + " " + (calcType === "column" && calcCol2 ? String(row[calcCol2] || "") : String(calcConstant || ""))).trim();
          }

          if (isNaN(val1)) {
            return "";
          }

          let resultNum = 0;
          switch (calcOperator) {
            case "+": resultNum = val1 + (isNaN(val2) ? 0 : val2); break;
            case "-": resultNum = val1 - (isNaN(val2) ? 0 : val2); break;
            case "*": resultNum = val1 * (isNaN(val2) ? 1 : val2); break;
            case "/": 
              if (isNaN(val2) || val2 === 0) return "";
              resultNum = val1 / val2; 
              break;
            default: resultNum = val1;
          }

          if (calcRounding === "int") {
            return Math.round(resultNum);
          } else if (calcRounding === "1dec") {
            return Math.round(resultNum * 10) / 10;
          } else if (calcRounding === "2dec") {
            return Math.round(resultNum * 100) / 100;
          } else {
            return resultNum;
          }
        };

        const computedRaw = await chunkProcess<any, any>(
          rawImportedData,
          10000,
          (row: any) => ({
            ...row,
            [safeNewName]: computeRowValue(row)
          }),
          pct => {
            setProgress(20 + Math.round(pct * 0.4));
            setStatusMessage(`Tính dòng tệp gốc: ${pct}%...`);
          }
        );

        const computedMain = await chunkProcess<any, any>(
          mainData,
          10000,
          (row: any) => ({
            ...row,
            [safeNewName]: computeRowValue(row)
          }),
          pct => {
            setProgress(60 + Math.round(pct * 0.4));
            setStatusMessage(`Tính dòng hiển thị: ${pct}%...`);
          }
        );

        const newCols = [...columns];
        if (!newCols.includes(safeNewName)) {
          newCols.push(safeNewName);
        }

        let updatedColConfigs = [...customColConfigs];
        const configExists = updatedColConfigs.some(cfg => cfg.originalName === safeNewName);
        if (!configExists) {
          updatedColConfigs.push({
            originalName: safeNewName,
            use: true,
            newName: safeNewName,
            role: ""
          });
        } else {
          updatedColConfigs = updatedColConfigs.map(cfg => {
            if (cfg.originalName === safeNewName) {
              return { ...cfg, use: true, newName: safeNewName };
            }
            return cfg;
          });
        }

        setRawImportedData(computedRaw);
        setMainData(computedMain);
        setColumns(newCols);
        setCustomColConfigs(updatedColConfigs);

        autoSaveSession(computedMain, computedRaw, newCols, fileName, mapping, updatedColConfigs);

        setSchemaAiLogs(prev => [
          ...prev,
          `⚡ Thành công! Cột mới [${safeNewName}] đã được tạo và nạp hoàn tất vào hệ thống.`,
          `💬 Đánh giá từ AI: ${explanation}`
        ]);
        
        setProgress(100);
        setStatusMessage("Thực thi phép tính kết xuất thành công!");
      }

    } catch (err: any) {
      setSchemaAiLogs(prev => [...prev, `❌ Thất bại: ${err.message}`]);
      alert("Trợ lý AI Autopilot thất bại: " + err.message);
    } finally {
      setIsSchemaAiRunning(false);
    }
  };

  const handleReportAiAutopilot = async (customCommand?: string) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện.");
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần cài đặt của AI Studio hoặc trong file .env để sử dụng Trợ lý AI Autopilot!");
      return;
    }

    setIsReportAiRunning(true);
    setReportAiLogs(["🔍 Bắt đầu khởi động Trợ lý AI Lập báo cáo tự động...", "📁 Nạp danh bạ cột dữ liệu từ file thực tế..."]);
    
    try {
      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: { headers: { "User-Agent": "aistudio-build" } }
      });

      const userCmd = customCommand || reportAiPrompt || "Phát hiện các cột và hạch toán báo cáo nhanh";
      
      const systemInstruction = `Bạn là một trợ lý AI phân tích dữ liệu chuyên nghiệp phụ trách lập báo cáo tổng hợp Kinh tế - Xã hội cho địa bàn Việt Nam.
      Có nhiệm vụ đọc hiểu yêu cầu phân tích của người dùng và nhận diện các cột chính xác trong file để lập báo cáo.
      
      Các vai trò của cột chính và từ khóa nhận diện phù hợp:
      - Cột "manganh": Các từ khóa như 'mã ngành', 'ma nganh', 'manganh', 'vsic', 'ngành', 'nganh'.
      - Cột "xa": Các từ khóa như 'xã', 'xa', 'phường', 'phuong', 'thị trấn', 'thi tran', 'địa bàn', 'dia ban', 'mã xã', 'ma xa'.
      - Cột "doanhthu": Các từ khóa như 'doanh thu', 'doanhthu', 'doanh số', 'doanh so', 'thu nhập', 'thu nhap', 'tiền', 'so tien', 'trị giá', 'tri gia'.
      - Cột "laodong": Các từ khóa như 'lao động', 'laodong', 'nhân viên', 'số người', 'so nguoi', 'quy mô', 'quy mo', 'nhân sự'. Nếu không thấy, ưu tiên để giá trị rỗng "".

      Đồng thời chỉ định "level" đại diện cho cấp ngành hạch toán:
      - 1: Nếu người dùng muốn phân tích nhóm ngành cấp 1 lớn (A, B, C...).
      - 2: (Mặc định) Nếu người dùng muốn phân phối lọc chi tiết ngành cấp 2 (2 chữ số) hoặc không nói rõ.

      Hãy trả về định dạng JSON duy nhất dưới đây, TUYỆT ĐỐI không viết giải thích gì ngoài JSON:
      {
        "manganh": "Tên cột mã ngành VSIC chính xác tìm thấy",
        "xa": "Tên cột địa bàn xã chính xác tìm thấy",
        "doanhthu": "Tên cột doanh thu chính xác tìm thấy (hoặc \"\" nếu không thấy)",
        "laodong": "Tên cột lao động chính xác tìm thấy (hoặc \"\" nếu không thấy)",
        "level": 1 hoặc 2,
        "explanation": "Lời giải thích ngắn gọn bằng tiếng Việt vì sao chọn các cột này và đề xuất"
      }`;

      // Thêm log hoạt động
      setReportAiLogs(prev => [...prev, `🧠 Đang phân tích so khớp bằng Generative AI...`, `💬 Khẩu lệnh phân tích: "${userCmd}"`]);

      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: `Danh sách cột thực tế: [${columns.join(", ")}]\nCâu lệnh: "${userCmd}"`,
        config: {
          systemInstruction,
          temperature: 0.1,
          responseMimeType: "application/json"
        }
      });

      const responseText = response.text || "";
      const parsed = JSON.parse(responseText.trim());

      const detectedMng = parsed.manganh || "";
      const detectedXa = parsed.xa || "";
      const detectedDt = parsed.doanhthu || "";
      const detectedLd = parsed.laodong || "";
      const level = parsed.level || 2;
      const explanation = parsed.explanation || "AI đã ghép nối thành công.";

      setReportAiLogs(prev => [
        ...prev,
        `✨ AI khớp cột thành công!`,
        `📍 Cột mã ngành: "${detectedMng || "Không tìm thấy"}"`,
        `📍 Cột địa bàn xã: "${detectedXa || "Không tìm thấy"}"`,
        `📍 Cột doanh thu: "${detectedDt || "Không tìm thấy (Bỏ qua)"}"`,
        `📍 Cột lao động: "${detectedLd || "Không tìm thấy (Bỏ qua)"}"`,
        `⚙️ Phân loại ngành: Cấp ${level}`,
        `💬 Giải thích: ${explanation}`,
        `🚀 Khởi tạo tiến trình hạch toán & Render biểu đồ...`
      ]);

      if (detectedMng) setQuickReportManganhCol(detectedMng);
      if (detectedXa) setQuickReportXaCol(detectedXa);
      if (detectedDt) setQuickReportDoanhThuCol(detectedDt);
      if (detectedLd) setQuickReportLaoDongCol(detectedLd);

      setMapping(prev => ({
        ...prev,
        manganh: detectedMng || prev.manganh,
        xa: detectedXa || prev.xa,
        doanhthu: detectedDt || prev.doanhthu,
        laodong: detectedLd || prev.laodong
      }));

      await sleep(1000);

      if (!detectedMng || !detectedXa) {
        throw new Error("Không tự động định danh được cột Mã Ngành hoặc Xã Địa Bàn từ file của bạn. Hãy chọn thủ công các cột ở phía dưới hoặc viết khẩu lệnh chỉ rõ hơn!");
      }

      await handleQuickReport(level, detectedMng, detectedXa, detectedDt, detectedLd);
      
      setReportAiLogs(prev => [...prev, `🎉 Lập báo cáo autopilot thành công tuyệt vời!`]);
    } catch (err: any) {
      setReportAiLogs(prev => [...prev, `❌ Thất bại: ${err.message}`]);
      alert("Autopilot thất bại: " + err.message);
    } finally {
      setIsReportAiRunning(false);
    }
  };

  const handleAppendSectorsToMainData = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước.");
      return;
    }
    const manganhCol = quickReportManganhCol || mapping.manganh;
    if (!manganhCol) {
      alert("Chưa cấu hình hoặc nhận diện được cột Mã Ngành!");
      return;
    }

    setLoading(true);
    setStatusMessage("Đang đối chiếu & trích xuất Ngành cấp 1, cấp 2...");
    setProgress(15);
    await sleep(300);

    try {
      const updatedRows = mainData.map(row => {
        const rawMng = String(row[manganhCol] || "").trim();
        const mngNormalized = normalizeSectorCode(rawMng);

        let parent1Code = "";
        if (mngNormalized) {
          if (/^[a-zA-Z]$/.test(mngNormalized)) {
            parent1Code = mngNormalized.toUpperCase();
          } else {
            parent1Code = getParentSectorCode(mngNormalized) || "";
          }
        }
        const s1Name = vsicRawData[parent1Code] || "Ngành cấp 1 chưa định nghĩa";
        const label1 = parent1Code ? `${parent1Code} - ${s1Name}` : "Chưa xác định";

        const parent2Code = mngNormalized ? mngNormalized.slice(0, 2) : "";
        const s2Name = vsicRawData[parent2Code] || "Ngành cấp 2 chưa định nghĩa";
        const label2 = parent2Code ? `${parent2Code} - ${s2Name}` : "Chưa xác định";

        return {
          ...row,
          "Ngành Cấp 1": label1,
          "Ngành Cấp 2": label2
        };
      });

      const newCols = [...columns];
      if (!newCols.includes("Ngành Cấp 1")) newCols.push("Ngành Cấp 1");
      if (!newCols.includes("Ngành Cấp 2")) newCols.push("Ngành Cấp 2");

      setMainData(updatedRows);
      setColumns(newCols);
      alert("Đã thêm thành công 2 cột 'Ngành Cấp 1' và 'Ngành Cấp 2' vào bảng tính chính!");
    } catch (err: any) {
      alert("Lỗi khi thêm cột ngành: " + err.message);
    } finally {
      setLoading(false);
      setProgress(100);
    }
  };

  const handleExportQuickReport = () => {
    if (quickReportResultRows.length === 0) {
      alert("Không có số liệu báo cáo để xuất!");
      return;
    }
    try {
      const ws = XLSX.utils.json_to_sheet(quickReportResultRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, `Báo cáo cấp ${quickReportLevel}`);
      XLSX.writeFile(wb, `Bao_Cao_Tong_Hop_Nganh_Cap_${quickReportLevel}_Va_Xa.xlsx`);
    } catch (err: any) {
      alert("Lỗi xuất Excel: " + err.message);
    }
  };

  // ==================== HỆ THỐNG AI HỌC LỆNH ĐỊNH NGHĨA CỘT ====================
  const handleAiColLearn = async (useGemini = true) => {
    if (customColConfigs.length === 0) {
      alert("Vui lòng nạp file dữ liệu chính hoặc dữ liệu nguồn trước để AI nhận diện các cột thực tế!");
      return;
    }

    const currentCols = customColConfigs.map(c => c.originalName);
    setIsLearningColAi(true);
    setLearningColLogs([
      "🔮 Khởi tạo sách để AI học lệnh ánh xạ..."
    ]);

    await sleep(350);
    setLearningColLogs(prev => [...prev, `📂 Đọc cấu trúc cột thực tế (${currentCols.length} cột): [${currentCols.join(", ")}]`]);
    await sleep(400);

    const promptText = aiColLearnPrompt.trim() || "Việt hóa gọn gàng tất cả các cột, tự gán vai trò tương thích cho Mã số thuế, doanh thu, địa bàn xã, mã ngành và lao động";

    // Mẫu fallback thông minh (Local Rule Heuristics Engine)
    const runSmartHeuristics = () => {
      const pl = promptText.toLowerCase();
      const mappings = customColConfigs.map(cfg => {
        const c = cfg.originalName;
        const cl = c.toLowerCase();
        let newName = cfg.newName || c;
        let role = cfg.role;
        let use = cfg.use;

        // Nếu người dùng yêu cầu bỏ/bỏ qua/loại trừ cột
        if (pl.includes("loại bỏ") || pl.includes("bỏ bớt") || pl.includes("chỉ giữ")) {
          // Nếu chỉ giữ một số cột nhất định
          if (pl.includes("chỉ giữ")) {
            const keepList = ["mst", "mast", "doanhthu", "laodong", "xa", "mota", "nganh", "id"];
            const shouldKeep = keepList.some(k => cl.includes(k));
            if (!shouldKeep) {
              use = false;
              newName = "";
            }
          }
        }

        // Nhận diện theo chỉ thị cụ thể trong prompt hoặc từ khóa mặc định
        if (cl.includes("st") || cl.includes("mst") || cl.includes("tax") || cl.includes("id") || cl.includes("taxcode")) {
          newName = "Mã Số Thuế";
          role = "idCol";
          use = true;
          if (pl.includes("mst thành") || pl.includes("tax thành")) {
            const match = promptText.match(/(?:mst|tax|mã số thuế)\s+(?:thành|là)\s+["'‘“]?([^"'‘”,\s]+)/i);
            if (match && match[1]) newName = match[1];
          }
        } else if (cl.includes("revenue") || cl.includes("thu") || cl.includes("so") || cl.includes("sale") || cl.includes("tien")) {
          newName = "Doanh Thu";
          role = "doanhthu";
          use = true;
          if (pl.includes("doanh thu thành") || pl.includes("doanh số thành")) {
            const match = promptText.match(/(?:doanh thu|doanh số)\s+(?:thành|là)\s+["'‘“]?([^"'‘”,\s]+)/i);
            if (match && match[1]) newName = match[1];
          }
        } else if (cl.includes("laodong") || cl.includes("ld") || cl.includes("labor") || cl.includes("nhansu") || cl.includes("nguoi")) {
          newName = "Số Lao Động";
          role = "laodong";
          use = true;
        } else if (cl.includes("xa") || cl.includes("phuong") || cl.includes("diaban") || cl.includes("town") || cl.includes("district")) {
          newName = "Địa bàn (Xã)";
          role = "xa";
          use = true;
        } else if (cl.includes("nganh") || cl.includes("vsic") || cl.includes("code") || cl.includes("dtv")) {
          newName = "Mã Ngành ĐK";
          role = "manganh";
          use = true;
        } else if (cl.includes("mota") || cl.includes("nganhnghe") || cl.includes("desc") || cl.includes("act") || cl.includes("nội dung")) {
          newName = "Mô Tả Hoạt Động";
          role = "mota";
          use = true;
        }

        return { originalMatch: c, newName, role, use };
      });
      return mappings;
    };

    try {
      let finalMappings: any[] = [];
      const apiKey = useGemini ? ((import.meta as any).env?.VITE_GEMINI_API_KEY || "") : "";

      if (useGemini && !apiKey) {
        setIsLearningColAi(false);
        alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong cấu hình Secrets của AI Studio hoặc tệp .env để huấn luyện bằng trí tuệ nhân tạo Gemini!\n\nNếu chưa có khoá API, bạn có thể sử dụng nút [HỌC ĐỊNH DẠNG TRỰC TIẾP] bên cạnh để kích hoạt thuật toán nhận diện local offline tự động.");
        return;
      }

      if (useGemini && apiKey) {
        setLearningColLogs(prev => [...prev, `🧠 Đang thiết lập kết nối mô hình Gemini 3.5 để phân tích lệnh của bạn...`]);
        await sleep(400);
        setLearningColLogs(prev => [...prev, `💬 Đang phân tích cú pháp tự nhiên: "${promptText}"`]);
        await sleep(350);
        setLearningColLogs(prev => [...prev, "📡 Gửi yêu cầu ánh xạ thông minh lên máy chủ mô hình bảo mật..."]);
        
        const ai = new GoogleGenAI({
          apiKey: apiKey,
          httpOptions: { headers: { "User-Agent": "aistudio-build" } }
        });

        const systemInstruction = `Bạn là Trí tuệ Nhân tạo học và huấn luyện ánh xạ cột dữ liệu Excel/CSV cho phần mềm Kế hoạch Phát triển Kinh tế Việt Nam.
Dựa trên danh sách cột thực tế: [${currentCols.join(", ")}]
Và chỉ thị tự nhiên học lệnh của người dùng: "${promptText}"

Hãy phân tích và ánh xạ tên mới, gán vai trò thích hợp.
Các vai trò hệ thống quy chuẩn chỉ gồm các nhãn sau hoặc để rỗng "":
- "idCol": Mã số thuế, mã định danh, số thứ tự chính (VD: MaST, MST, ID).
- "mota": Mô tả ngành nghề, diễn giải hoạt động (VD: MoTa, NganhNghe).
- "manganh": Mã ngành thực tế từ 1 đến 5 số (VD: MaNganh, VSIC).
- "xa": Địa bàn xã/phường (VD: Xa, Phuong, DiaBan).
- "doanhthu": Số liệu kinh doanh, doanh thu, lợi nhuận (VD: DoanhThu, DoanhSo).
- "laodong": Số lượng nhân sự, lao động (VD: LaoDong, NhanSu).

Hãy trả về một mảng JSON trực tiếp đại diện cho các trường được ánh xạ, tuyệt đối không viết thêm lời bình luận, không bọc thẻ markdown ngoài cục diện JSON. Định dạng bắt buộc:
[
  { "originalMatch": "tên_cột_gốc_chính_xác", "newName": "tên_mới_việt_hóa", "role": "vai_trò", "use": true/false }
]`;

        const response = await ai.models.generateContent({
          model: "gemini-3.5-flash",
          contents: `Cấu trúc cụ thể các cột thực tế: [${currentCols.join(", ")}]\nKhẩu lệnh người dùng muốn học: "${promptText}"`,
          config: {
            systemInstruction,
            temperature: 0.1,
            responseMimeType: "application/json"
          }
        });

        const text = response.text || "";
        finalMappings = JSON.parse(text.trim());
        setLearningColLogs(prev => [...prev, "🧬 Hệ thống AI Gemini đã phân tích thông tin & phản hồi cấu trúc tối ưu thành công!"]);
        await sleep(300);
      } else {
        setLearningColLogs(prev => [...prev, "⚡ Khởi động Thuật toán Nhận diện trực tiếp (Local Heuristics Engine)..."]);
        await sleep(400);
        setLearningColLogs(prev => [...prev, "🔍 Đang rà soát và so khớp từ khóa tiếng Việt râu ria, không dấu & đối tỷ lệ từ viết tắt..."]);
        await sleep(500);
        setLearningColLogs(prev => [...prev, "⚙️ Đang tiến hành gán vai trò (Mã số thuế, doanh thu, xã, lao động, mô tả, ngành) tự động..."]);
        await sleep(400);
        finalMappings = runSmartHeuristics();
      }

      if (!Array.isArray(finalMappings) || finalMappings.length === 0) {
        throw new Error("Dữ liệu phản hồi hoặc phân tích không hợp lệ.");
      }

      // Cập nhật lên cấu hình hiển thị hiện thời
      const updated = customColConfigs.map(cfg => {
        const match = finalMappings.find(m => m.originalMatch.toLowerCase() === cfg.originalName.toLowerCase());
        if (match) {
          return {
            ...cfg,
            newName: match.newName || cfg.newName,
            role: match.role || cfg.role,
            use: match.use !== undefined ? match.use : cfg.use
          };
        }
        return cfg;
      });

      setCustomColConfigs(updated);

      // Đơn giản hóa đồng bộ sang Mapping
      const nextMapping = { ...mapping };
      updated.forEach(u => {
        if (u.role && u.role in nextMapping) {
          nextMapping[u.role as keyof typeof mapping] = u.originalName;
        }
      });
      setMapping(nextMapping);

      setLearningColLogs(prev => [
        ...prev,
        "✨ Hoàn tất việc gán nhãn cho các cột thực tế theo lệnh học thành công!",
        `📊 Đã xử lý ${updated.filter(u => u.newName).length} cột hoạt động.`,
        "💡 Bạn có thể lưu lại thiết lập của bảng này thành một mẫu Lệch học riêng biệt để tái sử dụng ở ô bên dưới."
      ]);

      alert("AI đã học lệnh và cấu hình cột của bạn thành công! Hãy duyệt lại bảng cấu hình.");
    } catch (err: any) {
      setLearningColLogs(prev => [...prev, `❌ Lỗi: ${err.message}. Đang tự động xử lý bằng Heuristics local...`]);
      const mappings = runSmartHeuristics();
      const updated = customColConfigs.map(cfg => {
        const match = mappings.find(m => m.originalMatch.toLowerCase() === cfg.originalName.toLowerCase());
        if (match) {
          return {
            ...cfg,
            newName: match.newName,
            role: match.role,
            use: match.use
          };
        }
        return cfg;
      });
      setCustomColConfigs(updated);
      setLearningColLogs(prev => [...prev, "✅ Đã nạp thành công thiết lập gán bằng danh sách từ khóa dự phòng local!"]);
    } finally {
      setIsLearningColAi(false);
    }
  };

  const applyLearnedCommand = (cmd: typeof colLearnedCommands[0]) => {
    if (customColConfigs.length === 0) {
      alert("Hãy tải tệp tin Excel lên trước để có khung cột áp dụng lệnh học!");
      return;
    }

    const updated = customColConfigs.map(cfg => {
      // Tìm khớp chính xác
      let bestMatch = cmd.mappings.find(m => m.originalMatch.toLowerCase() === cfg.originalName.toLowerCase());
      
      if (!bestMatch) {
        // Khớp loãng (substring)
        bestMatch = cmd.mappings.find(m => {
          const pat = m.originalMatch.toLowerCase();
          const orig = cfg.originalName.toLowerCase();
          return orig.includes(pat) || pat.includes(orig);
        });
      }

      if (bestMatch) {
        return {
          ...cfg,
          newName: bestMatch.newName,
          role: bestMatch.role,
          use: bestMatch.use
        };
      }
      return cfg;
    });

    setCustomColConfigs(updated);

    // Đồng bộ sang vai trò toàn cục
    const nextMapping = { ...mapping };
    updated.forEach(u => {
      if (u.role && u.role in nextMapping) {
        nextMapping[u.role as keyof typeof mapping] = u.originalName;
      }
    });
    setMapping(nextMapping);

    alert(`🎉 Đã áp dụng thành công lệnh học: "${cmd.name}". Các cột trùng khớp đã tự động đổi tên và gán vai trò.`);
  };

  const handleSaveCurrentAsCommand = () => {
    const cmdName = newColCommandName.trim();
    if (!cmdName) {
      alert("Vui lòng điền tên cho Lệnh học mới của bạn!");
      return;
    }

    if (customColConfigs.length === 0) {
      alert("Không có cấu hình cột nào hiện tại để biến thành lệnh học. Vui lòng nạp và đặt tên cột trước!");
      return;
    }

    const activeMappings = customColConfigs.map(c => ({
      originalMatch: c.originalName,
      newName: c.newName,
      role: c.role,
      use: c.use
    }));

    const newCommand = {
      id: "cmd-" + Date.now(),
      name: `🎓 ${cmdName}`,
      description: `Nhận dạng & phân loại tự chế cho ${activeMappings.filter(m => m.newName).length} cột hoạt động.`,
      createdAt: new Date().toISOString(),
      mappings: activeMappings
    };

    const nextCommands = [newCommand, ...colLearnedCommands];
    setColLearnedCommands(nextCommands);
    localStorage.setItem("colLearnedCommands", JSON.stringify(nextCommands));
    setNewColCommandName("");
    alert(`💾 Đã lưu Lệnh học "${cmdName}" vào cơ sở dữ liệu của bạn thành công!`);
  };

  const handleDeleteCommand = (id: string) => {
    if (id.startsWith("default-")) {
      alert("Đây là các Lệnh học hệ thống mặc định, không thể xóa!");
      return;
    }
    if (confirm("Bạn có chắc chắn muốn xóa bỏ Lệnh học định nghĩa cột này? Hành động này không thể hoàn tác.")) {
      const nextCommands = colLearnedCommands.filter(c => c.id !== id);
      setColLearnedCommands(nextCommands);
      localStorage.setItem("colLearnedCommands", JSON.stringify(nextCommands));
    }
  };

  const handleExportCommands = () => {
    try {
      const dataStr = JSON.stringify(colLearnedCommands, null, 2);
      const dataUri = 'data:application/json;charset=utf-8,'+ encodeURIComponent(dataStr);
      
      const exportFileDefaultName = 'ai-column-learned-commands.json';
      
      const linkElement = document.createElement('a');
      linkElement.setAttribute('href', dataUri);
      linkElement.setAttribute('download', exportFileDefaultName);
      linkElement.click();
    } catch (e: any) {
      alert("Lỗi xuất file: " + e.message);
    }
  };

  const handleImportCommands = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], "UTF-8");
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            // Validate basic structure
            const isValid = parsed.every(item => item.name && Array.isArray(item.mappings));
            if (!isValid) {
              alert("File lệnh học không đúng định dạng chuẩn. Vui lòng kiểm tra lại!");
              return;
            }
            
            // Merge with existing
            const merged = [...parsed, ...colLearnedCommands.filter(c => !parsed.some(p => p.id === c.id))];
            setColLearnedCommands(merged);
            localStorage.setItem("colLearnedCommands", JSON.stringify(merged));
            alert(`📥 Đã nhập thành công ${parsed.length} Lệnh học mới vào thư viện dữ liệu!`);
          } else {
            alert("Định dạng dữ liệu tệp lệnh học phải là một mảng JSON các rules mẫu.");
          }
        } catch (error: any) {
          alert("Lỗi phân tích cú pháp tệp JSON: " + error.message);
        }
      };
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
    // Don't force redirect, render table inline directly!
    // setActiveTab("xemdulieu");
  };

  // 7. TRÌNH KIỂM TRA LOGIC ĐA ĐIỀU KIỆN (NẾU ... THÌ PHẢI...)
  const handleLogicRuleAdd = (type: "if" | "then") => {
    if (type === "if") {
      if (!newIfRule.col) {
        alert("Vui lòng chọn cột điều kiện NẾU!");
        return;
      }
      setIfRules([...ifRules, newIfRule]);
      setNewIfRule({ col: "", op: "==", val: "", isFieldCompare: false });
    } else {
      if (!newThenRule.col) {
        alert("Vui lòng chọn cột điều kiện THÌ PHẢI!");
        return;
      }
      setThenRules([...thenRules, newThenRule]);
      setNewThenRule({ col: "", op: "==", val: "", isFieldCompare: false });
    }
  };

  const handleLogicCheck = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi kiểm tra logic.");
      return;
    }
    if (ifRules.length === 0) {
      alert("Hãy định cấu hình ít nhất 1 quy tắc rà soát 'NẾU' ở Bước 1!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu kiểm tra logic đa điều kiện...");
    await sleep(200);

    // Chạy phép tính toán logic cực kỳ ổn định, hỗ trợ so sánh số, ô rỗng mặc định coi như 0 khi so sánh số
    const checkValue = (rowVal: any, op: string, compareVal: string) => {
      const v1 = String(rowVal !== undefined && rowVal !== null ? rowVal : "").trim();
      const v2 = String(compareVal).trim();

      const v1LC = v1.toLowerCase();
      const v2LC = v2.toLowerCase();

      // 1. Kiểm tra dạng rỗng
      if (op === "trống") return v1 === "";
      if (op === "không trống") return v1 !== "";

      // 2. Kiểm tra số
      // Trích xuất phần số từ v1 và v2 (bằng cách giữ lại chỉ chữ số, dấu chấm, dấu trừ)
      const cleanV1 = v1.replace(/,/g, "").replace(/[^0-9.\-]/g, "");
      const cleanV2 = v2.replace(/,/g, "").replace(/[^0-9.\-]/g, "");

      const num1 = v1 === "" ? 0 : parseFloat(cleanV1);
      const num2 = parseFloat(cleanV2);

      // Nếu cả 2 đều parse được ra số hợp lệ (hoặc v1 trống và v2 là số)
      const isNum2 = !isNaN(num2) && cleanV2 !== "";
      const isNum1 = !isNaN(num1) && (v1 === "" || cleanV1 !== "");

      if (isNum2 && isNum1) {
        if (op === "==") return num1 === num2;
        if (op === "!=") return num1 !== num2;
        if (op === ">") return num1 > num2;
        if (op === "<") return num1 < num2;
        if (op === ">=") return num1 >= num2;
        if (op === "<=") return num1 <= num2;
      }

      // 3. Nếu không phải số hoặc parse lỗi, so khớp chuỗi
      if (op === "==") return v1LC === v2LC;
      if (op === "!=") return v1LC !== v2LC;
      if (op === "chứa") return v1LC.includes(v2LC);
      if (op === "không chứa") return !v1LC.includes(v2LC);

      // Dự phòng toán tử so sánh cho chuỗi
      if (op === ">") return v1LC > v2LC;
      if (op === "<") return v1LC < v2LC;
      if (op === ">=") return v1LC >= v2LC;
      if (op === "<=") return v1LC <= v2LC;

      return false;
    };

    const results = await chunkProcess(
      mainData,
      10000,
      (row, index) => {
        if (!row || typeof row !== 'object') return row;
        // 1. Phép toán NẾU
        const ifMatches = ifRules.map(r => {
          const compVal = r.isFieldCompare ? String(row[r.val] !== undefined && row[r.val] !== null ? row[r.val] : "") : r.val;
          return checkValue(row[r.col], r.op, compVal);
        });
        const satisfiesIf = ifCombine === "AND" 
          ? ifMatches.every(v => v === true) 
          : ifMatches.some(v => v === true);

        let biViPham = false;
        let noteLoi = "";

        const getRuleDescription = (r: LogicRule) => {
          const rightSide = r.isFieldCompare ? `Cột [${r.val}]` : `'${r.val}'`;
          return `(${r.col} ${r.op} ${rightSide})`;
        };

        if (thenRules.length === 0) {
          // Chỉ rà soát và lọc theo điều kiện NẾU độc lập
          if (satisfiesIf) {
            biViPham = true;
            const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
            noteLoi = `[ĐÃ TÌM THẤY] Thỏa mãn điều kiện lọc kiểm tra: { ${descriptIf} }; `;
          }
        } else {
          // 2. Phép toán THÌ PHẢI
          const thenMatches = thenRules.map(r => {
            const compVal = r.isFieldCompare ? String(row[r.val] !== undefined && row[r.val] !== null ? row[r.val] : "") : r.val;
            return checkValue(row[r.col], r.op, compVal);
          });
          const satisfiesThen = thenCombine === "AND"
            ? thenMatches.every(v => v === true)
            : thenMatches.some(v => v === true);

          // Nếu logicRuleMode là conflict: Báo lỗi nếu thỏa mãn CẢ HAI (NẾU và ĐỒNG THỜI CÓ)
          // Nếu logicRuleMode là must_satisfy: Báo lỗi nếu thỏa mãn NẾU nhưng KHÔNG đạt THÌ PHẢI
          if (logicRuleMode === "conflict") {
            if (satisfiesIf && satisfiesThen) {
              biViPham = true;
              const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
              const descriptThen = thenRules.map(getRuleDescription).join(` ${thenCombine} `);
              noteLoi = `[MÂU THUẪN LOGIC] Thỏa mãn đồng thời: { NẾU: ${descriptIf} } và { CÓ THÊM: ${descriptThen} }; `;
            }
          } else {
            if (satisfiesIf && !satisfiesThen) {
              biViPham = true;
              const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
              const descriptThen = thenRules.map(getRuleDescription).join(` ${thenCombine} `);
              noteLoi = `[VI PHẠM LOGIC] NẾU thỏa mãn: { ${descriptIf} } THÌ BẮT BUỘC PHẢI THỎA MÃN: { ${descriptThen} }; `;
            }
          }
        }

        return {
          ...row,
          "Loi_Logic": biViPham ? noteLoi : "✅ Đạt",
          "_satisfiesIf": satisfiesIf,
          "_violated": biViPham
        };
      },
      pct => {
        setProgress(pct);
        setStatusMessage(`Đang chạy kiểm tra logic đa điều kiện: ${pct}%...`);
      }
    );

    setMainData(results);
    setColumns(Object.keys(results[0] || {}));
    setFileName(`KiemTraLogic_${fileName}`);

    const totalRows = results.length;
    const violatedCount = results.filter((r: any) => r && r._violated === true).length;
    const passedCount = totalRows - violatedCount;

    setAiScanMetrics({
      total: totalRows,
      violated: violatedCount,
      passed: passedCount,
      violatedPercent: totalRows > 0 ? ((violatedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
      passedPercent: totalRows > 0 ? ((passedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
      expression: "Quy tắc rà soát logic đa điều kiện",
      prompt: "Kiểm tra logic thủ công"
    });

    if (violatedCount > 0) {
      setLogicFilterMode("violated");
    } else {
      setLogicFilterMode("all");
    }

    setProgress(100);
    setStatusMessage(`Kiểm tra hoàn tất! Đã phân tích kiểm tra và phát hiện các dòng lỗi.`);
    await sleep(400);
    setLoading(false);
  };

  // --- TRÍ TUỆ NHÂN TẠO - RÀ QUÉT LOGIC TỰ HỌC HỎI QUA AI GEMINI ---
  const handleAiLogicScan = async (overridePrompt?: string) => {
    const activePrompt = overridePrompt || aiRulePrompt;
    if (!activePrompt.trim()) {
      alert("Vui lòng nhập khẩu lệnh rà quét bằng tiếng Việt!");
      return;
    }
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện quét!");
      return;
    }

    const activePromptLC = activePrompt.toLowerCase();
    const isReportRequest = (
      activePromptLC.includes("tổng hợp") || 
      activePromptLC.includes("tong hop") || 
      activePromptLC.includes("báo cáo") || 
      activePromptLC.includes("bao cao") ||
      activePromptLC.includes("phân tích doanh thu")
    ) && (
      activePromptLC.includes("xã") || 
      activePromptLC.includes("xa") || 
      activePromptLC.includes("ngành") || 
      activePromptLC.includes("nganh") ||
      activePromptLC.includes("cơ cấu") ||
      activePromptLC.includes("tỉ trọng") ||
      activePromptLC.includes("tỷ trọng")
    );

    if (isReportRequest) {
      // 1. Tự dò tìm các cột phù hợp cho Báo cáo
      const colManganh = columns.find(c => /mã\s*ngành|manganh|vsic|mã\s*nghe|manghe|ngành/i.test(c)) || quickReportManganhCol || mapping.manganh || "";
      const colXa = columns.find(c => /xã|xa|địa\s*bàn|dia\s*ban|phường|phuong/i.test(c)) || quickReportXaCol || mapping.xa || "";
      const colDoanhThu = columns.find(c => /doanh\s*thu|doanhthu|thu\s*nhập|thunhap|tiền|tien/i.test(c)) || quickReportDoanhThuCol || mapping.doanhthu || "";
      const colLaoDong = columns.find(c => /lao\s*động|laodong|người|nguoi|nhân\s*sự|nhansu/i.test(c)) || quickReportLaoDongCol || mapping.laodong || "";

      if (colManganh) setQuickReportManganhCol(colManganh);
      if (colXa) setQuickReportXaCol(colXa);
      if (colDoanhThu) setQuickReportDoanhThuCol(colDoanhThu);
      if (colLaoDong) setQuickReportLaoDongCol(colLaoDong);

      setLoading(true);
      setProgress(40);
      setStatusMessage("Hệ thống phát hiện lệnh Tổng Hợp Báo Cáo! Đang tự động chuyển sang Tab 'Tổng Hợp Báo Cáo'...");
      
      await sleep(1000);
      setActiveTab("tonghop");
      
      const targetLevel = activePromptLC.includes("cấp 1") || activePromptLC.includes("cap 1") ? 1 : 2;
      setProgress(75);
      setStatusMessage(`Đang chạy hạch toán tổng hợp: Ngành Cấp ${targetLevel} và Địa bàn Xã / Phường...`);
      await sleep(600);

      try {
        await handleQuickReport(targetLevel, colManganh, colXa, colDoanhThu, colLaoDong);
      } catch (err: any) {
        alert("Lỗi rẽ hướng tổng hợp liên kết: " + err.message);
      }
      setLoading(false);
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần Cài đặt > Khóa bí mật (Secrets) của AI Studio hoặc trong file .env!");
      return;
    }

    setLoading(true);
    setProgress(10);
    setStatusMessage("Trí tuệ nhân tạo đang phân tích và dịch khẩu lệnh...");

    try {
      let expression = "";

      // Nếu chạy nhanh bằng luật đã phân tích trước đó, bỏ qua gọi AI
      if (overridePrompt) {
        const matched = savedAiRules.find(r => r.prompt === overridePrompt || r.name === overridePrompt);
        if (matched) {
          expression = matched.expression;
        }
      }

      if (!expression) {
        const ai = new GoogleGenAI({
          apiKey: apiKey,
          httpOptions: {
            headers: {
              "User-Agent": "aistudio-build"
            }
          }
        });

        const systemPrompt = `Bạn là chuyên gia chuyển dịch khẩu lệnh tiếng Việt thành biểu thức điều kiện JavaScript chính xác cho bảng dữ liệu.
Mục tiêu là: Dịch yêu cầu tìm kiếm lỗi dữ liệu của người dùng thành một biểu thức logic JavaScript trả về true khi dòng đó bị lỗi hoặc vi phạm điều kiện.
Bạn PHẢI sử dụng biến đối tượng là 'row' để truy cập các cột của dòng.

Các cột dữ liệu hiện tại trong file của người dùng gồm: [${columns.filter(c => !c.startsWith("_")).map(c => `'${c}'`).join(", ")}].
Hãy phân tích ngôn từ của người dùng và khớp chính xác các cột trên. Nếu cột có tiếng Việt, hãy truy cập theo dạng row['Tên Cột'].
Luôn chú ý kiểu dữ liệu (nếu so sánh số, hãy dùng parseFloat(row['Tên Cột']) hoặc so sánh trực tiếp, loại bỏ dấu phẩy ngăn cách hàng nghìn nếu cần).

Ví dụ:
1. "DonGia < 0" -> parseFloat(String(row['DonGia'] || '0').replace(/,/g, '')) < 0
2. "Nhà thuê mượn = 1 nhưng tài sản lớn hơn 0" -> row['Thuê mượn'] == 1 && parseFloat(String(row['Tài sản'] || '0').replace(/,/g, '')) > 0
3. "Số điện thoại bị trống" -> !row['Số điện thoại'] || String(row['Số điện thoại']).trim() === ''

Quy tắc cực kỳ quan trọng:
CHỈ TRẢ VỀ DUY NHẤT CHUỖI BIỂU THỨC LOGIC TRONG JAVASCRIPT ĐỂ ĐƯA VÀO HÀM EVAL/NEW FUNCTION.
KHÔNG giải thích, KHÔNG bọc trong khối mã markdown (\`\`\`), KHÔNG chứa bất cứ từ ngữ thừa thãi hay ký tự nào ngoài biểu thức.`;

        const response = await ai.models.generateContent({
          model: "gemini-3.5-flash",
          contents: `Khẩu lệnh người dùng: "${activePrompt}"\nHãy chuyển dịch thành biểu thức Javascript viết dạng row['Cột']...`,
          config: {
            systemInstruction: systemPrompt,
            temperature: 0.1
          }
        });

        const rawResult = response.text || "";
        expression = rawResult
          .replace(/```javascript/gi, "")
          .replace(/```js/gi, "")
          .replace(/```/g, "")
          .trim();

        if (!expression) {
          throw new Error("Mô hình AI phản hồi trống hoặc không hợp lệ.");
        }
      }

      setAiTranslatedExpression(expression);
      setCustomRuleName(overridePrompt ? "" : `Quy tắc: ${activePrompt.substring(0, 25)}`);
      setStatusMessage("Dịch thuật thành công! Bắt đầu rà quét dữ liệu bằng bộ quét hiệu năng...");
      setProgress(30);
      await sleep(200);

      let violatedCount = 0;
      let passedCount = 0;
      const normalizedExpr = normalizeAiExpression(expression);

      const results = await chunkProcess(
        mainData,
        10000,
        (row, index) => {
          if (!row || typeof row !== 'object') return row;
          
          let biViPham = false;
          try {
            const evaluator = new Function("row", "getFlexibleValue", `
              try {
                return !!(${normalizedExpr});
              } catch(e) {
                return false;
              }
            `);
            biViPham = evaluator(row, getFlexibleValue);
          } catch (err) {
            biViPham = false;
          }

          if (biViPham) {
            violatedCount++;
          } else {
            passedCount++;
          }

          return {
            ...row,
            "Loi_Logic": biViPham ? `[LỖI AI-LỆNH]: thỏa mãn quy tắc "${activePrompt}"` : "✅ Đạt",
            "_satisfiesIf": true,
            "_violated": biViPham
          };
        },
        pct => {
          setProgress(Math.round(30 + (pct * 0.7)));
          setStatusMessage(`Trí tuệ nhân tạo đang quét dữ liệu: ${pct}%...`);
        }
      );

      setMainData(results);
      setColumns(Object.keys(results[0] || {}));
      setFileName(`QuetAI_${fileName}`);

      const totalRows = results.length;
      setAiScanMetrics({
        total: totalRows,
        violated: violatedCount,
        passed: passedCount,
        violatedPercent: totalRows > 0 ? ((violatedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
        passedPercent: totalRows > 0 ? ((passedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
        expression: expression,
        prompt: activePrompt
      });

      if (violatedCount > 0) {
        setLogicFilterMode("violated");
      } else {
        setLogicFilterMode("all");
      }

      setProgress(100);
      setStatusMessage(`Đã rà quét hoàn tất bằng AI dựa trên biểu thức: "${expression}"`);
      await sleep(400);
      setLoading(false);

      // Tự động cuộn mượt xuống khu vực Bảng tổng hợp kết quả và bảng xem nhanh
      setTimeout(() => {
        const docEl = document.getElementById("ai-scan-summary-section");
        if (docEl) {
          docEl.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }, 250);

    } catch (err: any) {
      console.error(err);
      alert(`Đã xảy ra lỗi khi quét bằng AI: ${err.message || err}`);
      setLoading(false);
    }
  };

  const handleSaveAiRule = () => {
    if (!aiRulePrompt.trim() || !aiTranslatedExpression.trim()) {
      alert("Chưa có biểu thức nào được AI dịch thành công để lưu cả!");
      return;
    }
    const ruleName = customRuleName.trim() || `Luật rà quét ${aiRulePrompt.substring(0, 20)}...`;
    
    // Đảm bảo không trùng lặp ID
    const newRule = {
      id: String(Date.now()),
      name: ruleName,
      prompt: aiRulePrompt.trim(),
      expression: aiTranslatedExpression.trim()
    };

    const updated = [newRule, ...savedAiRules];
    setSavedAiRules(updated);
    localStorage.setItem("vsic_saved_ai_rules", JSON.stringify(updated));
    alert(`Đã lưu thành công quy tắc "${ruleName}" vào bộ nhớ nhanh của trình duyệt!`);
  };

  const handleExportAiRules = () => {
    if (savedAiRules.length === 0) {
      alert("Chưa có danh sách quy tắc học lệnh nào được lưu để xuất!");
      return;
    }
    try {
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(savedAiRules, null, 2));
      const downloadAnchor = document.createElement("a");
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", "AI_Logic_Rules_Backup.json");
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
    } catch (err: any) {
      alert(`Lỗi xuất file: ${err.message || err}`);
    }
  };

  const handleImportAiRules = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], "UTF-8");
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            const valid = parsed.every(item => item.id && item.name && item.prompt && item.expression);
            if (!valid) {
              alert("Lỗi: Các quy tắc trong file chứa định dạng không hợp lệ!");
              return;
            }
            // Hợp nhất tránh trùng lặp
            const merged = [...parsed, ...savedAiRules.filter(existing => !parsed.some(p => p.id === existing.id))];
            setSavedAiRules(merged);
            localStorage.setItem("vsic_saved_ai_rules", JSON.stringify(merged));
            alert(`Đã nạp và đồng bộ thành công ${parsed.length} quy tắc học lệnh thông minh bằng AI!`);
          } else {
            alert("Tệp tin JSON tải lên không hợp lệ (phải là một danh sách các quy tắc)!");
          }
        } catch (err: any) {
          alert(`Nạp lỗi! Đọc nội dung bị lỗi: ${err.message || err}`);
        }
      };
    }
  };

  const handleDeleteAiRule = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = savedAiRules.filter(r => r.id !== id);
    setSavedAiRules(updated);
    localStorage.setItem("vsic_saved_ai_rules", JSON.stringify(updated));
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

  const renderAiMacroCognitiveCenter = () => {
    // Chỉ hiển thị các lệnh học tương ứng với tab phân hệ đang thiết lập để tránh làm loãng hoặc xung đột giao diện
    const displayedMacros = aiMacros.filter(macro => macro.module === activeTab);

    return (
      <div className="bg-gradient-to-br from-[#1e1b4b]/90 via-[#111827] to-[#0f172a] border border-[#4338ca]/30 rounded-2xl p-6 mb-6 space-y-6 shadow-2xl relative overflow-hidden backdrop-blur-md">
        <div className="absolute top-0 right-0 w-80 h-80 bg-indigo-500/10 rounded-full blur-3xl -z-10 pointer-events-none"></div>
        <div className="absolute bottom-0 left-0 w-80 h-80 bg-purple-500/10 rounded-full blur-3xl -z-10 pointer-events-none"></div>

        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-gray-850 pb-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-indigo-600/20 border border-indigo-500/30 rounded-xl shadow-inner">
              <Brain className="w-6 h-6 text-indigo-400 animate-pulse" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-white uppercase tracking-wide flex items-center gap-2 font-mono">
                BỘ NÃO HỌC LỆNH THÔNG MIÊN (AI MACRO STORAGE ENGINE)
                <span className="text-[10px] bg-indigo-500/20 text-indigo-300 font-mono font-medium px-2 py-0.5 rounded-full border border-indigo-500/20">
                  DUAL-PHÂN HỆ v3.2
                </span>
              </h2>
              <p className="text-xs text-gray-400 mt-1 font-sans">
                Huấn luyện tự động nhận diện và ghép nối cột qua khẩu lệnh tiếng Việt. Lưu bộ nhớ vạn năng để kích hoạt 0 giây.
              </p>
            </div>
          </div>

          {/* Export/Import Buttons */}
          <div className="flex items-center gap-2">
            <button
              onClick={handleExportMacros}
              className="px-3 py-1.5 bg-[#1e293b] hover:bg-gray-800 text-gray-300 font-semibold text-xs rounded-xl border border-gray-800 transition-all cursor-pointer flex items-center gap-1.5 hover:text-white font-sans"
              title="Tải tệp lưu các quy tắc đã học về máy tính"
            >
              <Download className="w-3.5 h-3.5 text-indigo-400" /> Xuất Bộ Nhớ (.json)
            </button>
            
            <label className="px-3 py-1.5 bg-[#1e293b] hover:bg-gray-800 text-gray-300 font-semibold text-xs rounded-xl border border-gray-800 transition-all cursor-pointer flex items-center gap-1.5 hover:text-white font-sans">
              <Upload className="w-3.5 h-3.5 text-teal-400" /> Nạp Cấu Hình
              <input
                type="file"
                accept=".json"
                onChange={handleImportMacros}
                className="hidden"
              />
            </label>
          </div>
        </div>

        {/* Input box to teach AI */}
        <div className="bg-[#1e1b4b]/20 border border-[#312e81]/40 p-5 rounded-xl space-y-4">
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-indigo-300 flex items-center gap-1.5 font-mono">
              <Sparkles className="w-3.5 h-3.5 text-indigo-400" /> NHẬP KHẨU LỆNH ĐỂ AI TỰ HỌC CHỌN CỘT:
            </label>
            <div className="flex flex-col sm:flex-row gap-2">
              <input
                type="text"
                value={macroPrompt}
                onChange={(e) => setMacroPrompt(e.target.value)}
                placeholder="Ví dụ: 'Tổng hợp theo xã ngành cấp 2 cho doanh thu và lao động' hoặc 'Kiểm tra mã ngành và mô tả'"
                className="flex-1 bg-[#0f172a] border border-[#312e81]/60 rounded-xl px-4 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-500 font-sans"
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !isLearning) {
                    handleLearnMacro();
                  }
                }}
              />
              <button
                onClick={handleLearnMacro}
                disabled={isLearning}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-xl text-xs font-bold transition-all disabled:opacity-50 cursor-pointer flex items-center justify-center gap-1.5 border border-indigo-500/20 font-sans"
              >
                {isLearning ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 animate-spin" /> Đang Học Lệnh...
                  </>
                ) : (
                  <>
                    <Zap className="w-3.5 h-3.5 text-amber-300" /> AI Học Chọn Cột
                  </>
                )}
              </button>
            </div>
            <p className="text-[10.5px] text-gray-500 font-sans leading-normal">
              * Khuyên dùng: Vui lòng nạp dữ liệu chính trước ở Trang Chủ hoặc Tab "Xem &amp; Định Nghĩa Cột" sau đó gửi yêu cầu học. Trực tiếp so khớp các cột có sẵn trong file của bạn.
            </p>
          </div>

          {/* AI learning feedback */}
          {learningResult && (
            <div className="bg-[#0f172a]/90 border border-emerald-500/25 p-4.5 rounded-xl space-y-4.5 animate-fade-in">
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-gray-800 pb-2">
                <span className="text-xs font-bold text-emerald-400 flex items-center gap-1.5 font-mono">
                  <CheckCircle2 className="w-4 h-4" /> KẾT QUẢ PHÂN TÍCH CHỌN CỘT THÀNH CÔNG:
                </span>
                <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-emerald-500/10 text-emerald-300 border border-emerald-500/20">
                  PHÂN HỆ MỤC TIÊU: {learningResult.module === "tonghop" ? "TỔNG HỢP BÁO CÁO ĐỘNG" : "KIỂM TRA CHUẨN VSIC"}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 bg-[#111827]/40 p-3 rounded-lg border border-gray-800">
                {learningResult.module === "tonghop" ? (
                  <>
                    <div>
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Địa Bàn Xã</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.xa || "⚠️ (Trống/Bỏ qua)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Mã Ngành</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.manganh || "⚠️ (Trống/Bỏ qua)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Doanh Thu</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.doanhthu || "⚠️ (Không tính)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Lao Động</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.laodong || "⚠️ (Không tính)"}
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="sm:col-span-2">
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Mô Tả Thực Tế</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.mota || "⚠️ (Không thấy)"}
                      </div>
                    </div>
                    <div className="sm:col-span-2">
                      <div className="text-[9.5px] text-gray-400 uppercase font-mono font-bold">Cột Mã Ngành VSIC</div>
                      <div className="text-xs text-white font-semibold font-mono mt-0.5 truncate bg-[#0f172a] px-1.5 py-1 rounded border border-gray-800/50">
                        {learningResult.columns.manganh || "⚠️ (Không thấy)"}
                      </div>
                    </div>
                  </>
                )}
              </div>

              <div className="flex flex-col sm:flex-row items-end gap-3 pt-1">
                <div className="flex-1 space-y-1 w-full">
                  <label className="text-[10px] font-bold text-gray-400 font-mono block uppercase">ĐẶT TÊN TUỲ CHỈNH CHO LỆNH ĐỂ LƯU THƯ VIỆN:</label>
                  <input
                    type="text"
                    value={customMacroName}
                    onChange={(e) => setCustomMacroName(e.target.value)}
                    placeholder="Ví dụ: Báo cáo Doanh thu Xã, Kiểm tra VSIC,..."
                    className="w-full bg-[#111827] border border-gray-800 rounded-lg px-3 py-2 text-xs text-white focus:outline-none focus:border-indigo-500 font-sans font-bold"
                  />
                </div>
                <button
                  onClick={handleSaveLearnMacro}
                  className="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-5 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow border border-emerald-500/20 shrink-0 font-sans h-9.5"
                >
                  <Save className="w-4 h-4" /> [💾 Lưu học lệnh này]
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Learned Commands Library (Tái sử dụng 0 giây) */}
        <div className="space-y-3 pt-1">
          <div className="text-xs font-bold text-indigo-300 tracking-wider uppercase font-mono flex items-center gap-1.5">
            <Cpu className="w-4 h-4 text-purple-400" /> THƯ VIỆN LỆNH ĐÃ HỌC (TÁI SỬ DỤNG 0 GIÂY - KHÔNG GỌI LẠI GEMINI):
          </div>

          {displayedMacros.length === 0 ? (
            <div className="bg-[#111827]/40 rounded-xl p-5 text-center text-xs text-gray-500 border border-gray-850">
              Chưa có lệnh học nào phù hợp cho phân hệ này được thiết lập. Hãy nhập câu lệnh ở trên để AI tạo mới!
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {displayedMacros.map((macro) => {
                const isTongHop = macro.module === "tonghop";
                return (
                  <div
                    key={macro.id}
                    onClick={() => handleExecuteMacro(macro)}
                    className="group bg-[#0f172a] hover:bg-[#1e1b4b]/20 border border-gray-850 hover:border-indigo-500/50 rounded-xl p-4 transition-all cursor-pointer flex flex-col justify-between gap-3 shadow hover:shadow-indigo-950/20 relative"
                  >
                    <div className="space-y-2">
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-sans font-bold text-xs text-[#f1f5f9] group-hover:text-indigo-350 transition-colors line-clamp-1">
                          {macro.name}
                        </span>
                        <div className="flex items-center gap-1 shrink-0">
                          <span className={`text-[9px] font-mono px-1.5 py-0.5 rounded font-bold ${
                            isTongHop 
                              ? "bg-amber-500/10 text-amber-300 border border-amber-500/25" 
                              : "bg-indigo-500/10 text-indigo-300 border border-indigo-500/25"
                          }`}>
                            {isTongHop ? "Báo cáo" : "VSIC"}
                          </span>
                          <button
                            onClick={(e) => handleDeleteMacro(macro.id, e)}
                            className="text-gray-500 hover:text-red-400 p-0.5 rounded hover:bg-red-500/10 transition-all cursor-pointer"
                            title="Xóa lệnh này"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                      <p className="text-[10.5px] text-gray-400 line-clamp-2 leading-relaxed font-sans">
                        "{macro.prompt}"
                      </p>
                    </div>

                    <div className="flex flex-wrap gap-1.5 pt-2 border-t border-gray-850/60 text-[9px] font-mono text-gray-500">
                      {isTongHop ? (
                        <>
                          {macro.columns?.xa && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-gray-800/40">Xã: {macro.columns.xa}</span>}
                          {macro.columns?.manganh && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-gray-800/40">Mã: {macro.columns.manganh}</span>}
                          {macro.columns?.doanhthu && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-gray-800/40">DT: {macro.columns.doanhthu}</span>}
                          {macro.columns?.laodong && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-gray-800/40">LĐ: {macro.columns.laodong}</span>}
                        </>
                      ) : (
                        <>
                          {macro.columns?.mota && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[120px] border border-gray-800/40 text-[#a5f3fc]">Mô tả: {macro.columns.mota}</span>}
                          {macro.columns?.manganh && <span className="bg-[#1e293b]/70 px-1.5 py-0.5 rounded truncate max-w-[120px] border border-gray-800/40">Mã: {macro.columns.manganh}</span>}
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    );
  };

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
              HỆ THỐNG <span className="bg-purple-600 text-[10px] uppercase font-bold px-2 py-0.5 rounded-full tracking-widest">VSIC 2025</span>
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

            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Thông minh & Rà soát</div>

            <button 
              onClick={() => setActiveTab("chuanhoanganh")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "chuanhoanganh" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 animate-pulse" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Brain className="w-4 h-4 text-indigo-400" /> 🧠 Kiểm tra ngành VSIC
            </button>

            <button 
              onClick={() => setActiveTab("kiemtralogic")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "kiemtralogic" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <CheckSquare className="w-4 h-4 text-emerald-400" /> 🛡️ Kiểm Tra Logic
            </button>
            <button 
              onClick={() => setActiveTab("chonmau")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "chonmau" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Layers className="w-4 h-4 text-orange-400" /> 🎯 Chọn mẫu điều tra
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

            <button 
              onClick={() => setActiveTab("pdf2word")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "pdf2word" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-sm" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <FileText className="w-4 h-4 text-violet-400" /> 📄 Đọc PDF & Sang Word
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
          <div className={activeTab === "trangchu" ? "block" : "hidden"}>
            <div className="space-y-8 animate-fade-in">
              <div className="bg-gradient-to-r from-purple-900/40 via-[#1f2937] to-[#1f2937] border border-purple-500/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-center gap-6 justify-between">
                <div className="space-y-3 max-w-2xl">
                  <span className="bg-purple-900/50 border border-purple-500/30 text-purple-400 text-xs font-mono font-bold px-3 py-1 rounded-full uppercase tracking-wider">
                    HỆ THỐNG VISC 2025
                  </span>
                  <h2 className="text-2xl md:text-3xl font-extrabold text-white tracking-tight">
                    Công cụ xử lý so sánh kiểm tra tổng hợp dữ liệu.
                  </h2>
                  <p className="text-gray-300 text-sm leading-relaxed">
                    Là công cụ hỗ trợ người dùng xử lý so sánh kiểm tra và tổng hợp. Cho phép người dùng nạp tất cả các loai file excel/csv với mọi cấu trúc mà không phải chỉnh sửa thủ công. Hệ thống tự động nhận diện thông minh, dù tên cột hoặc thứ tự cột trong file của bạn khác nhau.
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
                    <div className="text-4xl font-extrabold text-white font-mono">
                      {Object.keys(vsicRawData).length}
                    </div>
                    <div className="text-[11px] font-bold text-gray-400 tracking-wider uppercase font-mono">Mã ngành VSIC hiện có</div>
                    <div className="text-[10px] text-green-400 font-mono">.
                      {localStorage.getItem("custom_vsic_is_pure") === "true" 
                        ? "Danh mục nạp riêng (Sạch 100%)" 
                        : "Bao gồm mã mẫu tích hợp"}
                    </div>
                  </div>
                </div>
              </div>

              {/* PHẦN DANH SÁCH CHỨC NĂNG CHÍNH */}
              <div className="space-y-6">
                <div className="border-b border-[#374151] pb-4">
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <Layers className="w-5 h-5 text-purple-400 animate-pulse" /> HƯỚNG DẪN & KÍCH HOẠT CHỨC NĂNG NHANH
                  </h3>
                  <p className="text-xs text-gray-400 mt-1">
                    
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
                        Tải file Excel/CSV gốc. Hỗ trợ <strong>Việt hóa / đặt tên lại</strong> cho các cột viết tắt khó nhớ. Đối với các cột không cần thiết để xuất báo cáo thì xóa trắng không đặt tên hệ thống sẽ tự động loại bỏ.
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
                        🔗 Ghép Nối Liên Kết Dữ Liệu
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Nhập hai tệp dữ liệu riêng biệt và kết hợp chúng thành một bảng duy nhất thông qua cột định danh chung (Left Join). Rất phù hợp nếu muốn bổ sung địa bàn cho các hồ sơ từ bảng bên ngoài.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("ghepnoi")}
                      className="w-full bg-[#111827] hover:bg-blue-900/30 text-blue-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-blue-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Bắt Đầu Ghép Nối <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 3: So Sánh Đối Chiếu */}
                  <div className="bg-[#1f2937]/50 border border-cyan-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-cyan-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-cyan-950/50 border border-cyan-500/30 p-2.5 rounded-xl text-cyan-400">
                          <ArrowLeftRight className="w-5 h-5" />
                        </div>
                        <span className="bg-cyan-900/40 text-cyan-300 border border-cyan-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          ĐỐI CHIẾU CŨ - MỚI
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-cyan-300 transition-colors">
                        🔄 So Sánh Đấu Đầu Đối Chiếu
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Dễ dàng so sánh hai thời điểm hoặc hai danh sách khác nhau (CŨ vs MỚI) để truy tìm các thay đổi: thêm mới, bị xóa hoặc biến động doanh thu lao động lớn. Hệ thống tự động đối chiếu thông minh.
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
                          PHÂN RÃ BẢNG
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-pink-300 transition-colors">
                        ✂️ Chia Tách File Theo Địa Bàn
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tách nhanh gọn file tổng thành nhiều file Excel con theo Xã / Phường, tự động loại bỏ các dòng trống không hợp lệ và nén thành tệp ZIP tải về tiện dụng.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("tachfile")}
                      className="w-full bg-[#111827] hover:bg-pink-900/30 text-pink-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-pink-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Chia Tách File <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 5: Tổng hợp báo cáo */}
                  <div className="bg-[#1f2937]/50 border border-indigo-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-indigo-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-indigo-950/50 border border-indigo-500/30 p-2.5 rounded-xl text-indigo-400">
                          <BarChart3 className="w-5 h-5" />
                        </div>
                        <span className="bg-indigo-900/40 text-indigo-300 border border-indigo-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          XỬ LÝ SỐ LIỆU ĐỘNG
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-indigo-300 transition-colors">
                        📊 Tổng Hợp Báo Cáo Đa Chiều
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Đếm bản ghi, sum tổng số lao động, tính toán doanh thu của các công ty dựa theo đơn vị Xã / Phường hoặc theo Mã Ngành Kinh Tế (VSIC) tự động, kết xuất báo cáo nhanh.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("tonghop")}
                      className="w-full bg-[#111827] hover:bg-indigo-900/30 text-indigo-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-indigo-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Tổng Hợp Báo Cáo <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 6: Chuẩn hóa khớp ngành */}
                  <div className="bg-[#1f2937]/50 border border-amber-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-amber-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-amber-950/50 border border-amber-500/30 p-2.5 rounded-xl text-amber-400">
                          <BrainCircuit className="w-5 h-5" />
                        </div>
                        <span className="bg-amber-900/40 text-amber-300 border border-amber-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          CHUẨN HÓA AI & MÃ VSIC
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-amber-300 transition-colors">
                        🧠 Khớp &amp; Sắp Mã Ngành VSIC
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Hệ thống tự động tra cứu, khớp mã ngành dựa trên mô tả văn bản hoạt động thực tế với danh mục 5 cấp Hệ thống ngành kinh tế Việt Nam (VSIC) của Tổng cục Thống kê.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("chuanhoanganh")}
                      className="w-full bg-[#111827] hover:bg-amber-900/30 text-amber-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-amber-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Khớp Mã Ngành <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 7: Kiểm Quy Tắc Logic */}
                  <div className="bg-[#1f2937]/50 border border-emerald-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-emerald-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-emerald-950/50 border border-emerald-500/30 p-2.5 rounded-xl text-emerald-400">
                          <CheckCircle2 className="w-5 h-5" />
                        </div>
                        <span className="bg-emerald-900/40 text-emerald-300 border border-emerald-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          RÀ SOÁT QUY TẮC ĐỘNG
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-emerald-300 transition-colors">
                        🛂 Kiểm Tra Quy Tắc Logic
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Thiết lập quy tắc ràng buộc động "NẾU - THÌ" (ví dụ: NẾU doanh thu = 0 THÌ lao động phải = 0) để quét toàn bộ dữ liệu chỉ trong vài giây và báo cáo vi phạm.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("kiemtralogic")}
                      className="w-full bg-[#111827] hover:bg-[#10b981]/10 text-emerald-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-emerald-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Kiểm Tra Logic <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 8: PDF & Word */}
                  <div className="bg-[#1f2937]/50 border border-violet-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-violet-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-violet-950/50 border border-violet-500/30 p-2.5 rounded-xl text-violet-400">
                          <FileText className="w-5 h-5" />
                        </div>
                        <span className="bg-violet-900/40 text-violet-300 border border-violet-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          TIỆN ÍCH CHUYỂN ĐỔI
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-violet-300 transition-colors">
                        📄 Đọc PDF Sang Word / Excel
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tải lên tệp PDF scan hoặc kết xuất từ phần mềm để hệ thống tự động bóc tách số liệu bảng biểu, cho phép trích xuất tải về dạng DOCX hoặc XLSX hoàn toàn bảo mật.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("pdf2word")}
                      className="w-full bg-[#111827] hover:bg-violet-900/30 text-violet-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-violet-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Chuyển PDF Sang Word <ArrowRight className="w-3.5 h-3.5" />
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
                      Truy cập bảng <strong>📂 Xem & Định nghĩa cột</strong>. Tải tệp Excel gốc lên, thực hiện đổi tên Việt hóa dễ thương cho các cột và dán nhãn vai trò tương thích giúp hệ thống dễ chỉ huy dữ liệu chuẩn xác.
                    </p>
                  </div>
                  <div className="space-y-1">
                    <span className="font-bold text-white block">02. Kiểm tra mã ngành & Rà rà logic</span>
                    <p className="text-gray-400">
                      Sử dụng <strong>🧠 Kiểm tra khớp ngành </strong> để hoàn thiện liên kết 5 cấp ngành nghề; tiếp theo sử dụng <strong>🛂  Kiểm tra Logic</strong> thiết lập các quy chuẩn kiểm tra để lọc sạch các bản ghi lỗi hoặc dị thường.
                    </p>
                  </div>
                  <div className="space-y-1">
                    <span className="font-bold text-white block">03. Tổng hợp báo cáo & xuất file</span>
                    <p className="text-gray-400">
                      Qua trang <strong>📊 Tổng Hợp Báo Cáo</strong> để quy nập các chỉ thị hoặc chọn <strong>✂️ Tách File</strong> tạo tệp zip con của các xã gửi cho từng địa bàn. Bấm tải tệp Excel thành phẩm là xong!
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* 2. TAB FILE VIEWER & COLUMN MAPPING */}
          <div className={activeTab === "xemdulieu" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in font-sans">
              
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
                        <div className="flex items-center justify-between mb-1.5">
                          <label className="text-[11px] font-bold text-gray-300 block font-mono">
                            1. CHỌN CÁC SHEET MUỐN GHÉP:
                          </label>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={() => setSelectedSheetsToMerge([...detectedSheets])}
                              className="text-[10px] text-amber-400 hover:text-amber-300 hover:underline cursor-pointer bg-transparent border-0"
                            >
                              Chọn tất cả
                            </button>
                            <span className="text-gray-600 text-[10px]">|</span>
                            <button
                              type="button"
                              onClick={() => setSelectedSheetsToMerge([])}
                              className="text-[10px] text-gray-400 hover:text-gray-300 hover:underline cursor-pointer bg-transparent border-0"
                            >
                              Bỏ chọn cả
                            </button>
                          </div>
                        </div>
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

                      <div className="space-y-4 flex flex-col justify-between">
                        <div className="space-y-2">
                          <label className="text-[11px] font-bold text-gray-300 block font-mono">
                            2. CHỌN CỘT CHUNG (ID/MST) ĐỘNG:
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
                          type="button"
                          onClick={handleMergeWorkbookSheets}
                          disabled={selectedSheetsToMerge.length < 2 || !sheetMergeCommonKey}
                          className="w-full bg-gradient-to-r from-amber-500 to-yellow-600 hover:from-amber-600 hover:to-yellow-700 disabled:from-gray-700 disabled:to-gray-800 disabled:text-gray-500 disabled:cursor-not-allowed text-white font-bold text-xs py-2.5 px-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer border-0"
                        >
                          ⚡ GHÉP CÁC SHEET THÀNH 1 BẢNG CHUNG
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* === HỆ THỐNG AI HỌC LỆNH ĐỊNH NGHĨA CỘT === */}
              {rawImportedData.length > 0 && (
                <div className="bg-[#1e1b4b]/40 border border-purple-500/30 rounded-2xl p-6 space-y-6 animate-slide-up">
                  <div className="flex flex-col lg:flex-row lg:items-center justify-between border-b border-purple-900/40 pb-4 gap-4">
                    <div>
                      <h4 className="text-base font-bold text-white flex items-center gap-2">
                        <BrainCircuit className="w-5.5 h-5.5 text-purple-400 animate-pulse" /> 🧠 TRỢ LÝ AI HỌC LỆNH & QUẢN LÝ ÁNH XẠ CỘT THÔNG MINH
                      </h4>
                      <p className="text-xs text-purple-200/60 mt-1">
                        Dạy AI nhận dạng mẫu tiêu đề cột bằng giọng nói/văn bản tự nhiên, hoặc kích hoạt thư viện lệnh học định cấu hình mẫu một lần cho mọi file sau.
                      </p>
                    </div>

                    <div className="flex flex-wrap gap-2 items-center">
                      <label className="bg-purple-950/40 hover:bg-purple-900/50 text-purple-300 font-bold text-xs px-3.5 py-2 rounded-xl transition-all border border-purple-800/40 flex items-center gap-1.5 cursor-pointer">
                        <Upload className="w-3.5 h-3.5" /> Nhập lệnh học (.json)
                        <input
                          type="file"
                          accept=".json"
                          onChange={handleImportCommands}
                          className="hidden"
                        />
                      </label>
                      <button
                        onClick={handleExportCommands}
                        className="bg-purple-950/40 hover:bg-purple-900/50 text-purple-300 font-bold text-xs px-3.5 py-2 rounded-xl transition-all border border-purple-800/40 flex items-center gap-1.5 cursor-pointer"
                        title="Tải tệp JSON chứa toàn bộ lệnh học của bạn để backup hoặc chia sẻ"
                      >
                        <Download className="w-3.5 h-3.5" /> Xuất thư viện lệnh (.json)
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
                    {/* Cột 1: Huấn luyện AI */}
                    <div className="xl:col-span-7 space-y-4 bg-[#111827]/60 p-4 rounded-xl border border-purple-900/25">
                      <div>
                        <label className="text-xs font-bold text-gray-200 block mb-1.5 uppercase font-mono">
                          🗣️ Nhập khẩu lệnh của bạn hoặc chọn các mẫu gợi ý bên dưới:
                        </label>
                        <textarea
                          rows={3}
                          value={aiColLearnPrompt}
                          onChange={(e) => setAiColLearnPrompt(e.target.value)}
                          placeholder="Ví dụ: Đặt tên cho MST thành 'Mã Số Thuế' và gán vai trò idCol, cột DoanhThu mới tên là 'Doanh Thu 2024' vai trò doanhthu, loại bỏ các cột không dùng khác..."
                          className="w-full bg-gray-950 border border-purple-950 hover:border-purple-500/50 rounded-xl px-3 py-2.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 placeholder-gray-650 leading-relaxed font-sans"
                        />
                      </div>

                      {/* Gợi ý Lệnh nhanh */}
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Đổi tên cột MaST thành Mã Số Thuế gán vai trò idCol, cột Xa thành Địa bàn Xã vai trò xa, MoTa thành Mô Tả Hoạt Động vai trò mota.")}
                          className="text-[10px] bg-purple-950/30 hover:bg-purple-950/60 text-purple-300 px-2.5 py-1 rounded-lg border border-purple-900/30 transition-all cursor-pointer"
                        >
                          📌 Cú pháp thuế chuẩn
                        </button>
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Chỉ giữ lại cột Mã Số Thuế và Mô tả hoạt động kinh doanh, loại bỏ tất cả các cột dư thừa khác ra khỏi file mới.")}
                          className="text-[10px] bg-purple-950/30 hover:bg-purple-950/60 text-purple-300 px-2.5 py-1 rounded-lg border border-purple-900/30 transition-all cursor-pointer"
                        >
                          📌 Rút gọn giữ MST & Mô tả
                        </button>
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Việt hóa có dấu thật gọn cho mọi tiêu đề cột, gán đúng vai trò số cho DoanhThu và LaoDong.")}
                          className="text-[10px] bg-purple-950/30 hover:bg-purple-950/60 text-purple-300 px-2.5 py-1 rounded-lg border border-purple-900/30 transition-all cursor-pointer"
                        >
                          📌 Việt hóa gọn gàng tự động
                        </button>
                      </div>

                      <div className="flex gap-2">
                        <button
                          onClick={() => handleAiColLearn(true)}
                          disabled={isLearningColAi}
                          className={`flex-1 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 disabled:opacity-80 disabled:from-purple-950 disabled:to-indigo-950 disabled:text-purple-300 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-md shadow-purple-950/40 relative overflow-hidden ${
                            isLearningColAi ? "animate-pulse border border-purple-500/50" : ""
                          }`}
                        >
                          {isLearningColAi ? (
                            <>
                              <Loader2 className="w-4 h-4 text-purple-400 animate-spin" />
                              <span>ĐANG HUẤN LUYỆN GEMINI...</span>
                            </>
                          ) : (
                            <>
                              <BrainCircuit className="w-4 h-4 text-purple-200" />
                              <span>🧠 HUẤN LUYỆN QUA AI (GEMINI)</span>
                            </>
                          )}
                        </button>
                        
                        <button
                          onClick={() => handleAiColLearn(false)}
                          disabled={isLearningColAi}
                          className={`bg-[#374151] hover:bg-[#4b5563] disabled:opacity-50 disabled:bg-gray-800 disabled:text-gray-500 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer border border-[#4b5563] ${
                            isLearningColAi ? "opacity-60" : ""
                          }`}
                          title="Học lệnh tức thì bằng bộ phân tích từ khóa tiếng Việt không cần API key"
                        >
                          {isLearningColAi ? (
                            <Loader2 className="w-4 h-4 text-gray-500 animate-spin" />
                          ) : (
                            <Zap className="w-4 h-4 text-amber-400" />
                          )}
                          <span>HỌC ĐỊNH DẠNG TRỰC TIẾP</span>
                        </button>
                      </div>

                      {/* Log học lệnh */}
                      {learningColLogs.length > 0 && (
                        <div className="bg-gray-950 rounded-lg p-3 border border-purple-950/55 max-h-[120px] overflow-y-auto space-y-1 font-mono text-[10px] text-purple-200">
                          {learningColLogs.map((log, lidx) => (
                            <div key={lidx} className="leading-relaxed whitespace-pre-wrap">{log}</div>
                          ))}
                        </div>
                      )}

                      {/* Nút chạy áp dụng trực tiếp sau khi học */}
                      <div className="pt-3 border-t border-purple-950/40 space-y-2">
                        <button
                          onClick={handleApplyColumnRedefinition}
                          className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs px-4 py-3 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-emerald-950/40 border border-emerald-500/20 hover:scale-[1.01] active:scale-[0.99]"
                          title="Áp dụng ngay các cột đã định nghĩa để đổi tên và lọc dữ liệu chính"
                        >
                          <FileCheck className="w-4 h-4 text-emerald-200" />
                          ⚡ CHẠY ÁP DỤNG LỆNH & TẠO FILE SẠCH NGAY LẬP TỨC
                        </button>
                        <p className="text-[10px] text-gray-400 text-center leading-relaxed">
                          (Nhấn nút này để thực thi việc đổi tên, khớp nối lọc cột và chuyển kết quả sang tab <span className="text-purple-300 font-semibold">Xem Dữ Liệu</span>)
                        </p>
                      </div>
                    </div>

                    {/* Cột 2: Thư viện Lệnh học đã tích lũy */}
                    <div className="xl:col-span-5 space-y-4 bg-[#111827]/60 p-4 rounded-xl border border-purple-900/25 flex flex-col justify-between">
                      <div className="space-y-3">
                        <div className="flex items-center justify-between border-b border-purple-900/20 pb-2">
                          <label className="text-xs font-bold text-purple-300 block uppercase font-mono">
                            🎓 THƯ VIỆN LỆNH HỌC TÍCH LŨY ({colLearnedCommands.length}):
                          </label>
                        </div>

                        <div className="space-y-2 max-h-[160px] overflow-y-auto pr-1">
                          {colLearnedCommands.map((cmd) => (
                            <div key={cmd.id} className="bg-purple-950/10 border border-purple-900/20 p-2.5 rounded-lg flex items-start justify-between gap-3 text-xs">
                              <div className="space-y-0.5">
                                <div className="font-bold text-purple-200">{cmd.name}</div>
                                <div className="text-[10px] text-gray-400 leading-relaxed text-wrap">{cmd.description}</div>
                              </div>
                              <div className="flex gap-1.5 flex-shrink-0">
                                <button
                                  type="button"
                                  onClick={() => applyLearnedCommand(cmd)}
                                  className="bg-purple-900/60 hover:bg-purple-800 text-[10px] text-purple-200 font-bold px-2 py-1 rounded cursor-pointer transition-all"
                                  title="Áp dụng mẫu gán nhãn cột này lên bảng tính hiện thời"
                                >
                                  Áp dụng
                                </button>
                                {!cmd.id.startsWith("default-") && (
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteCommand(cmd.id)}
                                    className="text-[10px] hover:text-red-400 text-gray-500 font-bold px-1.5 py-1 rounded cursor-pointer transition-all"
                                    title="Xóa lệnh học này khỏi máy tính"
                                  >
                                    Xóa
                                  </button>
                                )}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Lưu lệnh học mới */}
                      <div className="bg-[#1f2937]/30 p-2.5 rounded-lg border border-purple-900/10 space-y-2 mt-2">
                        <label className="text-[11px] font-bold text-gray-300 block uppercase font-mono">
                          💾 Lưu cấu hình bảng hiện tại thành lệnh học mới:
                        </label>
                        <div className="flex gap-2 font-sans">
                          <input
                            type="text"
                            value={newColCommandName}
                            onChange={(e) => setNewColCommandName(e.target.value)}
                            placeholder="Tên lệnh học, vd: Cấu hình bảng xã Tân Bình"
                            className="flex-1 bg-gray-950 border border-purple-950/55 rounded-lg px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono"
                          />
                          <button
                            type="button"
                            onClick={handleSaveCurrentAsCommand}
                            className="bg-purple-700 hover:bg-purple-600 text-white font-bold text-xs px-3.5 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1"
                          >
                            <Save className="w-3.5 h-3.5" /> Lưu Lệnh
                          </button>
                        </div>
                      </div>
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
                                <th className="p-3 w-[260px]">VAI TRÒ HỆ THỐNG (MỤC TIÊU)</th>
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

                                    {/* Vai trò */}
                                    <td className="p-3">
                                      <select
                                        value={cfg.role || ""}
                                        onChange={(e) => {
                                          const selectedRole = e.target.value as any;
                                          const index = idx;
                                          const updated = [...customColConfigs];
                                          const oldRole = updated[index].role;
                                          
                                          if (selectedRole !== "") {
                                            updated.forEach((c, i) => {
                                              if (i !== index && c.role === selectedRole) {
                                                c.role = "";
                                              }
                                            });
                                          }
                                          
                                          updated[index].role = selectedRole;
                                          setCustomColConfigs(updated);

                                          setMapping(prev => {
                                            const next = { ...prev };
                                            if (oldRole && next[oldRole] === updated[index].originalName) {
                                              next[oldRole] = "";
                                            }
                                            if (selectedRole !== "") {
                                              next[selectedRole] = updated[index].originalName;
                                            }
                                            return next;
                                          });
                                        }}
                                        className="w-full bg-gray-950 border border-gray-850 hover:border-purple-500/50 rounded-lg px-2 py-1.5 text-[11px] text-gray-300 focus:outline-none focus:ring-1 focus:ring-purple-500 font-medium font-sans cursor-pointer"
                                      >
                                        <option value="">-- Để trống / Không gán --</option>
                                        <option value="idCol">🔑 Mã định danh độc nhất (ID/MST)</option>
                                        <option value="mota">📝 Mô tả hoạt động kinh doanh</option>
                                        <option value="manganh">🏷️ Mã ngành kinh tế (VSIC)</option>
                                        <option value="xa">🗺️ Địa bàn Xã / Phường</option>
                                        <option value="doanhthu">💰 Doanh thu / Doanh số</option>
                                        <option value="laodong">👥 Quy mô lao động / Nhân sự</option>
                                      </select>
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

              {/* PHÉP TÍNH CỘT VỚI CỘT (COLUMN FORMULA CALCULATOR) */}
              {rawImportedData.length > 0 && (
                <div className="bg-[#111827]/95 rounded-2xl p-5 border border-purple-500/20 space-y-4 animate-slide-up">
                  <div className="border-b border-gray-800 pb-3">
                    <div className="text-xs font-bold text-indigo-400 tracking-wider uppercase font-mono flex items-center gap-1.5 font-medium">
                      <Database className="w-5 h-5 text-indigo-400 animate-pulse" /> 🧮 CÔNG CỤ TÌNH PHÉP TÍNH CỘT VỚI CỘT & TẠO CỘT MỚI
                    </div>
                    <p className="text-xs text-gray-400 mt-1">
                      Tính toán số liệu nâng cao trực tiếp trên bảng tính của bạn. Bạn có thể cộng, trừ, nhân, chia 2 cột với nhau hoặc tính toán với một hằng số cố định, hoặc ghép nội dung cột văn bản.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end bg-[#0f172a]/50 p-4 rounded-xl border border-gray-800">
                    
                    {/* Tên cột kết quả */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-gray-300 block font-mono">
                        ✍️ 1. TÊN CỘT KẾT QUẢ MỚI:
                      </label>
                      <input
                        type="text"
                        value={calcColName}
                        onChange={(e) => setCalcColName(e.target.value)}
                        placeholder="VD: DoanhThuBinhQuan, Cong_X_Y"
                        className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono font-medium"
                      />
                    </div>

                    {/* Cột 1 */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-gray-300 block font-mono">
                        📂 2. CỘT THỨ NHẤT (A):
                      </label>
                      <select
                        value={calcCol1}
                        onChange={(e) => setCalcCol1(e.target.value)}
                        className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono"
                      >
                        <option value="">-- Chọn Cột A --</option>
                        {columns.map(c => (
                          <option key={c} value={c}>📊 {c}</option>
                        ))}
                      </select>
                    </div>

                    {/* Phép toán */}
                    <div className="md:col-span-1.5 space-y-1.5">
                      <label className="text-[11px] font-bold text-gray-300 block font-mono text-center">
                        ➕ PHÉP TOÁN:
                      </label>
                      <select
                        value={calcOperator}
                        onChange={(e) => setCalcOperator(e.target.value as any)}
                        className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white text-center focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono font-bold"
                      >
                        <option value="+">➕ Cộng (+)</option>
                        <option value="-">➖ Trừ (-)</option>
                        <option value="*">✖️ Nhân (*)</option>
                        <option value="/">➗ Chia (/)</option>
                        <option value="concat">🔗 Ghép chữ</option>
                      </select>
                    </div>

                    {/* Loại cột 2: Cột hay Hằng số */}
                    <div className="md:col-span-1.5 space-y-1.5">
                      <label className="text-[11px] font-bold text-gray-300 block font-mono text-center">
                        🎯 ĐỐI TƯỢNG B:
                      </label>
                      <select
                        value={calcType}
                        onChange={(e) => setCalcType(e.target.value as any)}
                        className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono"
                      >
                        <option value="column">Cột khác</option>
                        <option value="constant">Hằng số</option>
                      </select>
                    </div>

                    {/* Cột 2 hoặc Hằng số */}
                    <div className="md:col-span-3 space-y-1.5">
                      {calcType === "column" ? (
                        <>
                          <label className="text-[11px] font-bold text-gray-300 block font-mono">
                            📂 3. CỘT THỨ HAI (B):
                          </label>
                          <select
                            value={calcCol2}
                            onChange={(e) => setCalcCol2(e.target.value)}
                            className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono"
                          >
                            <option value="">-- Chọn Cột B --</option>
                            {columns.map(c => (
                              <option key={c} value={c}>📊 {c}</option>
                            ))}
                          </select>
                        </>
                      ) : (
                        <>
                          <label className="text-[11px] font-bold text-gray-300 block font-mono">
                            🔢 3. NHẬP GIÁ TRỊ HẰNG SỐ (B):
                          </label>
                          <input
                            type="text"
                            value={calcConstant}
                            onChange={(e) => setCalcConstant(e.target.value)}
                            placeholder="VD: 1000000, 1.2, Chuỗi chữ"
                            className="w-full bg-[#111827] border border-gray-700 hover:border-purple-500/50 rounded-lg px-2.5 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-mono"
                          />
                        </>
                      )}
                    </div>
                  </div>

                  {/* Lựa chọn làm tròn & nút xử lý */}
                  <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-1">
                    <div className="flex flex-wrap items-center gap-4 text-xs">
                      <span className="font-bold text-gray-400 font-mono">🎯 CHẾ ĐỘ LÀM TRÒN:</span>
                      <label className="flex items-center gap-1.5 cursor-pointer text-gray-300 hover:text-white">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "none"}
                          onChange={() => setCalcRounding("none")}
                          className="text-indigo-600 focus:ring-indigo-500 border-gray-700 bg-gray-900"
                        />
                        Không làm tròn
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-gray-300 hover:text-white">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "int"}
                          onChange={() => setCalcRounding("int")}
                          className="text-indigo-600 focus:ring-indigo-500 border-gray-700 bg-gray-900"
                        />
                        Số nguyên
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-gray-300 hover:text-white">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "1dec"}
                          onChange={() => setCalcRounding("1dec")}
                          className="text-indigo-600 focus:ring-indigo-500 border-gray-700 bg-gray-900"
                        />
                        1 chữ số thập phân
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-gray-300 hover:text-white">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "2dec"}
                          onChange={() => setCalcRounding("2dec")}
                          className="text-indigo-600 focus:ring-indigo-500 border-gray-700 bg-gray-900"
                        />
                        2 chữ số thập phân
                      </label>
                    </div>

                    <button
                      onClick={handleCalculateColumn}
                      className="bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-700 hover:to-purple-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-indigo-950/40 flex items-center gap-2 cursor-pointer border border-indigo-500/20"
                    >
                      ⚡ THỰC HIỆN PHÉP TÍNH & THÊM CỘT
                    </button>
                  </div>
                </div>
              )}

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
          </div>

          {/* 3. TAB GHÉP NỐI DỮ LIỆU */}
          <div className={activeTab === "ghepnoi" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in font-sans">
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

              {activeTab === "ghepnoi" && mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={mainData} 
                  columns={columns} 
                  title="BẢNG DỮ LIỆU GHÉP NỐI HIỆN TẠI" 
                  subtitle="Xem nhanh thông tin tích hợp của hai bảng dữ liệu sau khi gộp."
                  mapping={mapping}
                />
              )}
            </div>
          </div>

          {/* 4. TAB SO SÁNH CŨ MỚI (DIFF) */}
          <div className={activeTab === "sosanh" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in font-sans">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <Combine className="w-5 h-5 text-cyan-400" /> SO SÁNH HAI FILE DỮ LIỆU CŨ &amp; MỚI
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
                    <Combine className="w-4 h-4" /> BẮT ĐẦU SO SÁNH &amp; DIFF
                  </button>
                </div>
              </div>

              {activeTab === "sosanh" && mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={mainData} 
                  columns={columns} 
                  title="BẢNG DỮ LIỆU ĐỐI CHIẾU SỢI CHỈ ĐỎ (DIFF)" 
                  subtitle="Xem nhanh danh sách kết quả rà soát chênh lệch cũ mới, trạng thái cập nhật được điền tự động."
                  mapping={mapping}
                />
              )}
            </div>
          </div>

          {/* 5. TAB TÁCH DỮ LIỆU THEO CỘT */}
          <div className={activeTab === "tachfile" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in font-sans">
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

              {activeTab === "tachfile" && mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={mainData} 
                  columns={columns} 
                  title="DỮ LIỆU NGUỒN CHUẨB BỊ TÁCH FILE" 
                  subtitle="Xem nhanh danh sách dữ liệu chính sẽ được phân chia hệ thống."
                  mapping={mapping}
                />
              )}
            </div>
          </div>

          {/* 6. TAB TỔNG HỢP BÁO CÁO ĐỘNG */}
          <div className={activeTab === "tonghop" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-amber-400" /> TỔNG HỢP GỘP NHÓM &amp; PHÁT TRIỂN BÁO CÁO LINH HOẠT
                  </h3>
                  <p className="text-xs text-gray-400 font-sans mt-0.5">
                    Hệ thống hạch toán đa năng không khóa cứng cột. Cho phép bạn gộp nhóm dữ liệu gốc theo địa bàn xã, phân cấp mã ngành VSIC (Cấp 1 &amp; Cấp 2) hoặc phân nhóm trực tiếp từ bất kỳ tiêu chí dữ liệu nông nghiệp, dân số, công nghiệp nào để làm các loại điều tra thống kê khác nhau.
                  </p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6">
                    {/* BỘ LỰA CHỌN CỘT THỦ CÔNG */}
                    <div className="bg-[#111827]/80 p-5 rounded-xl border border-gray-850 space-y-5 shadow-lg">
                      <span className="text-xs font-bold text-cyan-400 tracking-wider uppercase font-mono block">
                        ⚙️ Cấu hình các cột phân nhóm dữ liệu gốc
                      </span>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Mã Ngành Hoặc Phân Nhóm Chính:</label>
                          <select 
                            value={quickReportManganhCol} 
                            onChange={(e) => setQuickReportManganhCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột chính phân nhóm --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <p className="text-[10px] text-gray-500 mt-1 font-sans">
                            Chọn cột mã ngành để quy nạp lên cấp 1, cấp 2 (VSIC) hoặc cột đặc tính gốc để phân tích trực tiếp.
                          </p>
                        </div>

                        <div>
                          <label className="text-xs font-semibold text-gray-300 block mb-1">Cột Xã / Địa Bàn / Đơn vị:</label>
                          <select 
                            value={quickReportXaCol} 
                            onChange={(e) => setQuickReportXaCol(e.target.value)}
                            className="w-full bg-[#1e293b] border border-gray-750 rounded-lg px-2.5 py-2 text-xs text-white focus:ring-1 focus:ring-cyan-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột xã/phường/địa bàn --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <p className="text-[10px] text-gray-500 mt-1 font-sans">
                            Dữ liệu sẽ được gộp và hiển thị lũy kế chi tiết theo từng giá trị địa bàn này.
                          </p>
                        </div>
                      </div>

                      {/* KHU VỰC THIẾT LẬP CHỈ TIÊU CỘNG DỒN ĐỘNG - KHÔNG KHÓA CỨNG */}
                      <div className="border-t border-gray-800 pt-4 space-y-3">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                          <div>
                            <span className="text-xs font-bold text-amber-400 tracking-wider uppercase font-mono block">
                              📊 Chọn cột chỉ tiêu số để cộng dồn (Không hạn chế số lượng)
                            </span>
                            <span className="text-[10.5px] text-gray-400">
                              Chọn nhiều chỉ tiêu tùy thích để phần mềm thực hiện cộng tổng cho từng nhóm (Ví dụ: Doanh thu, Lao động, Sản lượng, Vốn...).
                            </span>
                          </div>
                          
                          <div className="flex items-center gap-1.5 self-start">
                            <button
                              type="button"
                              onClick={() => {
                                // Tự động chọn các cột có khả năng là số dựa vào hàng mẫu thứ nhất
                                const firstRow = mainData[0] || {};
                                const numericCols = columns.filter(col => {
                                  const val = String(firstRow[col] || "");
                                  return val && !isNaN(parseFloat(val.replace(/[^0-9.\-]/g, "")));
                                });
                                // Bỏ qua cột xã và mã ngành
                                const cleanNumerics = numericCols.filter(col => col !== quickReportManganhCol && col !== quickReportXaCol);
                                setQuickReportSumCols(cleanNumerics.length > 0 ? cleanNumerics : columns.slice(0, 5));
                              }}
                              className="bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/10 rounded px-2 py-1 text-[10px] font-bold cursor-pointer transition-all active:scale-95"
                            >
                              ⚙️ Tự động chọn cột số
                            </button>
                            <button
                              type="button"
                              onClick={() => setQuickReportSumCols([])}
                              className="bg-gray-800 hover:bg-gray-750 text-gray-400 rounded px-2 py-1 text-[10px] font-bold cursor-pointer transition-all active:scale-95"
                            >
                              ❌ Xóa tất cả lựa chọn
                            </button>
                          </div>
                        </div>

                        {/* LIST CHECKBOX ĐỘNG ĐỂ LỰA CHỌN CỘT CHỈ TIÊU */}
                        <div className="bg-[#111827] border border-gray-850 rounded-xl p-3 max-h-[160px] overflow-y-auto grid grid-cols-2 md:grid-cols-4 gap-2">
                          {columns.map(col => {
                            const isChecked = quickReportSumCols.includes(col);
                            return (
                              <label 
                                key={col} 
                                className={`flex items-center gap-2 p-1.5 rounded-lg border text-xs cursor-pointer select-none transition-all ${
                                  isChecked 
                                    ? "bg-amber-950/20 border-amber-500/40 text-amber-300" 
                                    : "bg-gray-900/40 border-gray-800/40 text-gray-400 hover:bg-gray-800 hover:text-white"
                                }`}
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => {
                                    if (isChecked) {
                                      setQuickReportSumCols(prev => prev.filter(c => c !== col));
                                    } else {
                                      setQuickReportSumCols(prev => [...prev, col]);
                                    }
                                  }}
                                  className="rounded text-amber-550 focus:ring-amber-500 bg-gray-950 border-gray-800 w-3.5 h-3.5"
                                />
                                <span className="truncate" title={col}>{col}</span>
                              </label>
                            );
                          })}
                        </div>

                        {quickReportSumCols.length > 0 && (
                          <div className="flex flex-wrap gap-1.5 items-center bg-gray-900/30 p-2 rounded-lg border border-gray-850/40">
                            <span className="text-[10px] text-gray-500 uppercase font-bold font-mono">Đang chọn ({quickReportSumCols.length}):</span>
                            {quickReportSumCols.map(col => (
                              <span key={col} className="bg-amber-950/40 text-amber-400 border border-amber-500/20 px-2 py-0.5 rounded text-[10px] font-mono flex items-center gap-1">
                                {col}
                                <span 
                                  onClick={() => setQuickReportSumCols(prev => prev.filter(c => c !== col))}
                                  className="hover:text-red-400 cursor-pointer text-xs leading-none font-bold ml-1"
                                >
                                  ×
                                </span>
                              </span>
                            ))}
                          </div>
                        )}
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
                            className="mt-1 text-amber-500 focus:ring-amber-550 bg-gray-900 border-gray-750"
                          />
                          <div>
                            <div className="font-bold text-gray-100 font-sans">Bảng xoay ngang Pivot (Khuyên dùng)</div>
                            <div className="text-[10.5px] text-gray-400 mt-1 font-sans">
                              Mỗi xã địa bàn hiển thị thành một hàng ngang. Các nhóm phân loại và các chỉ tiêu được xoay thành các cột mở rộng liền kề song song hỗ trợ rà soát nhanh chóng.
                            </div>
                          </div>
                        </label>

                        <label className="flex items-start gap-3 text-xs text-gray-300 hover:text-white cursor-pointer select-none bg-[#111827]/50 p-3 rounded-lg border border-gray-800">
                          <input 
                            type="radio" 
                            name="quickReportFormatPivot"
                            checked={reportType === "flat"} 
                            onChange={() => setReportType("flat")}
                            className="mt-1 text-amber-500 focus:ring-amber-550 bg-gray-900 border-gray-750"
                          />
                          <div>
                            <div className="font-bold text-gray-100 font-sans">Bảng phẳng danh sách truyền thống</div>
                            <div className="text-[10.5px] text-gray-400 mt-1 font-sans">
                              Dạng danh mục phẳng chuẩn hóa. Mỗi dòng tương ứng một cặp địa bàn xã và nhóm phân loại với các chỉ tiêu cộng dồn xếp dọc.
                            </div>
                          </div>
                        </label>
                      </div>
                    </div>

                    {/* NÚT THỰC THI CHẠY TỔNG HỢP & BÊ CỘT */}
                    <div className="space-y-3">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                        <button 
                          onClick={() => handleQuickReport(0)}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-3.5 rounded-xl transition-all shadow-md hover:shadow-indigo-900/30 cursor-pointer flex items-center justify-center gap-2 font-sans border border-indigo-500/20 active:scale-95"
                          title="Hỗ trợ mọi loại điều tra: gom nhóm trực tiếp dựa trên tất cả giá trị có sẵn trong cột mà không cần tra cứu danh mục mã ngành VSIC"
                        >
                          📈 Chạy Gộp Nhóm Trực Tiếp &amp; Xã (Cho mọi loại điều tra)
                        </button>

                        <button 
                          onClick={() => handleQuickReport(1)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3.5 rounded-xl transition-all shadow-md hover:shadow-emerald-900/30 cursor-pointer flex items-center justify-center gap-2 font-sans active:scale-95"
                        >
                          📈 Tổng Hợp Ngành Cấp 1 &amp; Xã (Tra cứu VSIC)
                        </button>

                        <button 
                          onClick={() => handleQuickReport(2)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3.5 rounded-xl transition-all shadow-md hover:shadow-emerald-900/30 cursor-pointer flex items-center justify-center gap-2 font-sans active:scale-95"
                        >
                          📈 Tổng Hợp Ngành Cấp 2 &amp; Xã (Tra cứu VSIC)
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

                    {/* BẢNG HIỂN THỊ KẾT QUẢ ĐẦU RA */}
                    {activeTab === "tonghop" && quickReportResultRows.length > 0 && (
                      <div className="space-y-8 pt-4 animate-fade-in">
                        {/* Đồ thị doanh thu theo ngành cấp 1 quy nạp */}
                        <SectorRevenueChart 
                          mainData={mainData} 
                          columns={columns} 
                          reportLevel={quickReportLevel}
                          mapping={{
                            mota: mapping.mota || "",
                            manganh: quickReportManganhCol || mapping.manganh || "",
                            xa: quickReportXaCol || mapping.xa || "",
                            doanhthu: quickReportSumCols[0] || mapping.doanhthu || "",
                            laodong: quickReportLaoDongCol || mapping.laodong || "",
                            idCol: mapping.idCol || ""
                          }}
                        />

                        <BeautifulReportTable
                          rows={quickReportResultRows}
                          cols={quickReportResultCols}
                          level={quickReportLevel}
                          reportType={reportType}
                          onExport={handleExportQuickReport}
                        />
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950 font-sans">
                    ⚠️ Vui lòng nạp dữ liệu chính ở trang đầu tiên trước khi chạy hạch toán tổng hợp.
                  </div>
                )}
              </div>
            </div>
          </div>


          {/* 7. TAB KIỂM TRA & PHÂN TÍCH NGÀNH */}
          <div className={activeTab === "chuanhoanganh" ? "block font-sans" : "hidden"}>
            <div className="space-y-6 animate-fade-in font-sans">
              {renderAiMacroCognitiveCenter()}
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6 animate-fade-in">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <Brain className="w-5 h-5 text-indigo-400 animate-pulse" /> CÔNG CỤ TỰ ĐỘNG PHÂN TÍCH BẤT NHẤT QUÁN QUY LUẬT VSIC
                  </h3>
                  <p className="text-xs text-gray-400">Rà soát chéo thông minh: Quét toàn bộ tệp tin dữ liệu chính, phát hiện mâu thuẫn phân gán lỗi giữa Mô tả hoạt động thực tế và Bản mã ngành VSIC.</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6 border-t border-gray-800 pt-6">
                    
                    {/* KHU VỰC THIẾT LẬP CHỌN 2 CỘT ĐỀ PHÂN TÍCH */}
                    <div className="bg-[#111827] rounded-xl p-5 border border-purple-950/30 space-y-4">
                      <h4 className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5 pb-2 border-b border-gray-800">
                        <Sliders className="w-4 h-4 text-purple-400" /> THIẾT LẬP 2 CỘT RÀ SÁT CHÉO
                      </h4>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="space-y-1.5">
                          <label className="text-xs text-gray-300 font-semibold flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
                            Chọn Cột chứa Mô tả/Tên ngành thực tế:
                          </label>
                          <select
                            value={stdDescriptionCol || mapping.mota || ""}
                            onChange={(e) => setStdDescriptionCol(e.target.value)}
                            className="w-full bg-[#1f2937] border border-[#374151] rounded-xl px-3 py-2.5 text-xs text-white outline-none focus:ring-1 focus:ring-purple-500 cursor-pointer"
                          >
                            <option value="">-- Chọn cột mô tả --</option>
                            {columns.map(col => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                          <p className="text-[10px] text-gray-500">Chọn cột text chứa nội dung chi tiết hoạt động kinh tế.</p>
                        </div>

                        <div className="space-y-1.5">
                          <label className="text-xs text-gray-300 font-semibold flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-amber-400"></span>
                            Chọn Cột chứa Mã Ngành VSIC cấp 5:
                          </label>
                          <select
                            value={stdIndustryCol || mapping.manganh || ""}
                            onChange={(e) => setStdIndustryCol(e.target.value)}
                            className="w-full bg-[#1f2937] border border-[#374151] rounded-xl px-3 py-2.5 text-xs text-white outline-none focus:ring-1 focus:ring-purple-500 cursor-pointer"
                          >
                            <option value="">-- Chọn cột mã ngành --</option>
                            {columns.map(col => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                          <p className="text-[10px] text-gray-500">Cột chứa chuỗi mã định dạng cấp 5 (hoặc các cấp tự liên hợp).</p>
                        </div>
                      </div>
                    </div>

                    {/* BỘ NÚT CHẠY CHUẨN HÓA SANG CỘT MỚI */}
                    <div className="bg-[#111827]/80 rounded-2xl p-5 border border-purple-500/10 flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div className="space-y-1">
                        <h4 className="text-xs font-bold text-gray-300 uppercase tracking-wider font-mono">Bổ sung danh mục đối chiếu VSIC quốc gia</h4>
                        <p className="text-[11px] text-gray-400">Tự động đối sánh mã ngành của từng dòng với Danh mục chuẩn để sinh cột <strong>Tên Ngành Chuẩn VSIC</strong> bên cạnh cột dữ liệu gốc của bạn.</p>
                      </div>
                      <button
                        onClick={handleStandardizeSectorsAndMatch}
                        className="bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs py-2.5 px-5 rounded-xl cursor-pointer transition-all shadow-md active:scale-95 whitespace-nowrap flex items-center gap-1.5 self-start md:self-auto"
                      >
                        ⚡ CHẠY KHỚP MÃ &amp; BỔ SUNG CỘT CHUẨN VSIC
                      </button>
                    </div>

                    {/* HAI PANEL KẾT QUẢ ĐỐI SÁNH SONG SONG */}
                    <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 pt-2">
                      
                      {/* BẢNG 1: CÙNG MÔ TẢ -> KHÁC MÃ NGÀNH */}
                      <div className="bg-[#111827]/80 rounded-2xl p-5 border border-purple-500/20 space-y-4">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="bg-purple-950/50 border border-purple-500/30 text-purple-300 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Cùng Mô tả / Khác Mã</span>
                            <h4 className="text-sm font-bold text-white flex items-center gap-2">
                              📋 CÙNG MÔ TẢ ➔ LỆCH KHÁC MÃ VSIC ({inconAnalysis.descToCodes.length})
                            </h4>
                          </div>
                          <p className="text-[11px] text-gray-400 mt-1">
                            Tìm thấy các hộ khai nội dung kinh doanh giống hệt nhau nhưng bị cán bộ nhập / phân gán lệch sang các mã nghề khác nhau.
                          </p>
                        </div>

                        <div className="bg-[#111827] border border-gray-800 rounded-xl p-4 min-h-[250px] max-h-[450px] overflow-y-auto space-y-3">
                          {inconAnalysis.descToCodes.length === 0 ? (
                            <div className="text-xs text-emerald-400 flex items-center justify-center h-44 gap-1.5 font-mono">
                              <CheckCircle2 className="w-4 h-4 text-emerald-400 animate-pulse" /> Tuyệt vời! Không phát hiện mâu thuẫn "Cùng một mô tả gán khác mã".
                            </div>
                          ) : (
                            <div className="space-y-3 divide-y divide-gray-900">
                              {inconAnalysis.descToCodes.slice(0, visibleDescInconCount).map((item, idx) => (
                                <div key={idx} className="pt-3 first:pt-0 flex flex-col justify-between gap-2.5 text-xs">
                                  <div className="space-y-1.5 flex-1">
                                    <div className="font-semibold text-white bg-[#1f2937]/40 px-2.5 py-1.5 rounded-lg border border-gray-800 leading-relaxed font-mono">
                                      📝 "{item.motaText}" <span className="text-gray-500 text-[10px] ml-1">({item.occurrences} dòng giống nhau)</span>
                                    </div>
                                    <div className="flex flex-wrap gap-1.5 items-center text-[11px]">
                                      <span className="text-gray-500">Được phân gán xen kẽ:</span>
                                      {item.codes.map((c, cidx) => (
                                        <span key={cidx} className="bg-purple-950/40 text-purple-300 px-2 py-0.5 rounded border border-purple-500/15 font-mono flex items-center gap-1">
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
                                    className="bg-purple-600/20 hover:bg-purple-600 text-purple-300 hover:text-white px-2.5 py-1 rounded text-[10px] font-bold self-start cursor-pointer transition-all flex items-center gap-1 active:scale-95 shrink-0"
                                  >
                                    <Search className="w-3 h-3" /> Lọc {item.occurrences} dòng gốc
                                  </button>
                                </div>
                              ))}

                              {inconAnalysis.descToCodes.length > visibleDescInconCount && (
                                <div className="pt-3 pb-1 text-center">
                                  <button
                                    onClick={() => setVisibleDescInconCount(prev => prev + 100)}
                                    className="w-full bg-[#1b2330] hover:bg-purple-950/40 text-purple-300 border border-purple-500/20 rounded-lg py-2 text-xs font-bold font-sans cursor-pointer transition-all"
                                  >
                                    ➕ Xem tiếp {inconAnalysis.descToCodes.length - visibleDescInconCount < 100 ? inconAnalysis.descToCodes.length - visibleDescInconCount : 100} mâu thuẫn ẩn (Tổng {inconAnalysis.descToCodes.length})
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </div>

                      {/* BẢNG 2: GOM CÙNG MÃ NGÀNH -> KHÁC MÔ TẢ (Yêu cầu mới) */}
                      <div className="bg-[#111827]/80 rounded-2xl p-5 border border-indigo-500/20 space-y-4">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="bg-indigo-950/50 border border-indigo-500/30 text-indigo-300 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Cùng Mã / Khác Mô tả</span>
                            <h4 className="text-sm font-bold text-white flex items-center gap-2">
                              🏷️ GOM CÙNG MÃ VSIC ➔ KHÁC BIỆT MÔ TẢ CHỮ ({inconAnalysis.codeToDescs.length})
                            </h4>
                          </div>
                          <p className="text-[11px] text-gray-400 mt-1">
                            Gom nhóm các dòng ghi nhận cùng mã VSIC nhưng nội dung văn bản mô tả hoạt động thực tế lại khác nhau.
                          </p>
                        </div>

                        <div className="bg-[#111827] border border-gray-800 rounded-xl p-4 min-h-[250px] max-h-[450px] overflow-y-auto space-y-3">
                          {inconAnalysis.codeToDescs.length === 0 ? (
                            <div className="text-xs text-indigo-400 flex items-center justify-center h-44 gap-1.5 font-mono">
                              <CheckCircle2 className="w-4 h-4 text-indigo-400 animate-pulse" /> Tuyệt vời! Không phát hiện trường hợp đồng mã lệch chữ.
                            </div>
                          ) : (
                            <div className="space-y-3 divide-y divide-gray-900">
                              {inconAnalysis.codeToDescs.slice(0, visibleCodeInconCount).map((item, idx) => (
                                <div key={idx} className="pt-3 first:pt-0 flex flex-col justify-between gap-2.5 text-xs">
                                  <div className="space-y-1.5 flex-1">
                                    <div className="font-semibold text-amber-400 bg-[#1f2937]/40 px-2.5 py-1.5 rounded-lg border border-gray-800 leading-relaxed font-mono flex items-center justify-between">
                                      <span>🏷️ Mã VSIC: <strong className="text-white text-sm shrink-0 underline decoration-indigo-500/50">{item.codeValue}</strong></span>
                                      <span className="text-gray-500 text-[10px] font-normal font-sans">({item.occurrences} dòng)</span>
                                    </div>
                                    <div className="space-y-1">
                                      <div className="text-[10px] text-gray-500 uppercase font-semibold">Văn bản chi tiết lệch biệt:</div>
                                      <div className="max-h-[140px] overflow-y-auto pl-2 space-y-1 border-l-2 border-indigo-500/30">
                                        {item.descriptions.map((descObj, descIdx) => (
                                          <div key={descIdx} className="text-gray-300 bg-[#111827]/70 rounded p-1.5 flex items-start justify-between gap-2 border border-blue-950/20 text-[11px]">
                                            <span className="italic">"{descObj.desc}"</span>
                                            <span className="text-[9px] shrink-0 text-indigo-300 font-mono bg-indigo-950/70 px-1 py-0.5 rounded">
                                              ({descObj.count} dòng)
                                            </span>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                  <button
                                    onClick={() => {
                                      setSearchTerm(item.codeValue);
                                      setViewPage(1);
                                      setIsConfigExpanded(false);
                                      setActiveTab("xemdulieu");
                                    }}
                                    className="bg-indigo-600/20 hover:bg-indigo-600 text-indigo-300 hover:text-white px-2.5 py-1 rounded text-[10px] font-bold self-start cursor-pointer transition-all flex items-center gap-1 active:scale-95 shrink-0"
                                  >
                                    <Search className="w-3 h-3" /> Lọc {item.occurrences} dòng gốc
                                  </button>
                                </div>
                              ))}

                              {inconAnalysis.codeToDescs.length > visibleCodeInconCount && (
                                <div className="pt-3 pb-1 text-center">
                                  <button
                                    onClick={() => setVisibleCodeInconCount(prev => prev + 100)}
                                    className="w-full bg-[#1b2330] hover:bg-[#374151] text-indigo-300 border border-indigo-500/20 rounded-lg py-2 text-xs font-bold font-sans cursor-pointer transition-all"
                                  >
                                    ➕ Xem tiếp {inconAnalysis.codeToDescs.length - visibleCodeInconCount < 100 ? inconAnalysis.codeToDescs.length - visibleCodeInconCount : 100} sự khác biệt ẩn (Tổng {inconAnalysis.codeToDescs.length})
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </div>

                    </div>

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>

              {mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={mainData} 
                  columns={columns} 
                  title="BẢNG DỮ LIỆU CHUẨN HÓA VSIC" 
                  subtitle="Xem nhanh danh mục dữ liệu chính sau khi đã chuẩn hóa và khai bổ sung các cấp ngành VSIC."
                  mapping={mapping}
                />
              )}
            </div>
          </div>

          {/* 8. TAB KIỂM TRA LOGIC ĐA ĐIỀU KIỆN */}
          <div className={activeTab === "kiemtralogic" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in">
              <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
                <div>
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <CheckSquare className="w-5 h-5 text-emerald-400" /> TRÌNH THIẾT LẬP QUY TẮC KIỂM TRA DỮ LIỆU
                  </h3>
                  <p className="text-xs text-gray-400">Thiết lập quy tắc kiểm tra thông minh theo logic: Nếu (Điều kiện 1) xảy ra, thì (Điều kiện 2) bắt buộc phải đúng.</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6 border-t border-gray-800 pt-6">
                    
                    {/* PHÂN HỆ AI: HỌC LỆNH VÀ RÀ SOÁT THÔNG MINH QUA AI GEMINI */}
                    <div className="bg-gradient-to-r from-purple-950/20 to-indigo-950/20 border border-purple-500/20 rounded-2xl p-5 space-y-4">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-purple-500/10">
                        <div>
                          <h4 className="text-sm font-extrabold text-purple-300 flex items-center gap-2">
                            <Sparkles className="w-4 h-4 text-purple-400 animate-pulse" /> 🚀 TỰ HỌC LỆNH PHÂN TÍCH VÀ RÀ QUÉT BẰNG AI
                          </h4>
                          <p className="text-[11px] text-gray-400 mt-0.5">
                            Gõ điều kiện lỗi bằng tiếng Việt tự nhiên. Trợ lý AI sẽ tự động biên dịch sang biểu thức máy tính để quét toàn bộ file gốc.
                          </p>
                        </div>
                        
                        {/* Cặp nút Xuất/Nạp */}
                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            onClick={handleExportAiRules}
                            className="bg-indigo-900/30 hover:bg-indigo-900/40 text-indigo-300 border border-indigo-500/15 px-2.5 py-1 rounded-lg text-[11px] font-semibold cursor-pointer transition-all flex items-center gap-1 active:scale-95"
                            title="Tải bộ quy tắc logic đã lưu về máy dưới dạng JSON"
                          >
                            <Download className="w-3" /> Xuất bộ nhớ luật (.json)
                          </button>
                          
                          <label
                            className="bg-purple-900/30 hover:bg-purple-900/40 text-purple-300 border border-purple-500/15 px-2.5 py-1 rounded-lg text-[11px] font-semibold cursor-pointer transition-all flex items-center gap-1 active:scale-95"
                            title="Nạp bộ quy tắc logic (.json) từ máy tính của bạn"
                          >
                            <Upload className="w-3" /> Nạp tệp cấu hình
                            <input
                              type="file"
                              accept=".json"
                              onChange={handleImportAiRules}
                              className="hidden"
                            />
                          </label>
                        </div>
                      </div>

                      {/* KHU VỰC NHẬP LỆNH AI */}
                      <div className="grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
                        <div className="md:col-span-3 space-y-1.5">
                          <label className="text-xs font-bold text-gray-300 flex items-center gap-1">
                            ✍️ Nhập điều kiện lỗi tiếng Việt (Ví dụ: 'Tìm dòng có DonGia &lt; 0'):
                          </label>
                          <div className="relative">
                            <input
                              type="text"
                              value={aiRulePrompt}
                              onChange={(e) => setAiRulePrompt(e.target.value)}
                              placeholder="Nhập khẩu lệnh bằng tiếng Việt tự nhiên tại đây..."
                              className="w-full bg-[#111827] border border-purple-500/30 rounded-xl pl-3 pr-10 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 placeholder-gray-500"
                            />
                            {aiRulePrompt && (
                              <button
                                onClick={() => setAiRulePrompt("")}
                                className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-white text-xs cursor-pointer"
                              >
                                ✕
                              </button>
                            )}
                          </div>
                        </div>

                        <div>
                          <button
                            onClick={() => handleAiLogicScan()}
                            className="w-full bg-purple-600 hover:bg-purple-700 active:scale-[0.98] text-white font-bold text-xs py-2 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-md"
                          >
                            <Sparkles className="w-3.5 h-3.5 text-amber-200" /> QUÉT BẰNG AI GEMINI
                          </button>
                        </div>
                      </div>

                      {/* KẾT QUẢ DỊCH CỦA AI & KHU VỰC LƯU TRỮ */}
                      {aiTranslatedExpression && (
                        <div className="bg-[#111827]/80 rounded-xl p-4 border border-purple-500/20 space-y-3 animate-fade-in text-xs">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div className="font-mono text-[11px] text-gray-300">
                              <span className="text-purple-400 font-bold">🤖 Biểu thức máy hiểu (JS):</span> <code className="bg-[#1e1b4b] px-2 py-0.5 rounded text-purple-300 border border-purple-900/30 break-all">{aiTranslatedExpression}</code>
                            </div>
                            <span className="bg-emerald-950/50 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 rounded text-[10px] font-mono shrink-0">Dịch thành công!</span>
                          </div>

                          <div className="bg-[#1e1b4b]/20 p-3 rounded-lg border border-purple-950 flex flex-col sm:flex-row sm:items-center gap-3">
                            <div className="flex-1 space-y-1">
                              <span className="text-[10px] uppercase font-bold text-gray-400">🏷️ Đặt tên quy tắc này để lưu nhanh:</span>
                              <input
                                type="text"
                                value={customRuleName}
                                onChange={(e) => setCustomRuleName(e.target.value)}
                                placeholder="Ví dụ: Kiểm tra Đơn Giá âm, Mã ngành rỗng..."
                                className="w-full bg-[#111827] border border-gray-700/80 rounded-lg px-2 py-1 text-xs text-white"
                              />
                            </div>
                            <button
                              onClick={handleSaveAiRule}
                              className="bg-emerald-600 hover:bg-emerald-700 active:scale-[0.98] text-white font-bold text-xs px-4 py-2 rounded-lg transition-all cursor-pointer flex items-center gap-1 shrink-0 justify-center"
                            >
                              💾 Lưu học lệnh thông minh
                            </button>
                          </div>
                        </div>
                      )}

                      {/* BẢNG TỔNG HỢP KẾT QUẢ RÀ QUÉT AI */}
                      {aiScanMetrics && (
                        <div id="ai-scan-summary-section" className="bg-[#0f172a]/90 rounded-2xl p-5 border border-purple-500/35 space-y-4 animate-fade-in text-xs shadow-lg">
                          <div className="flex items-center justify-between border-b border-purple-500/10 pb-2">
                            <h5 className="text-[12px] font-extrabold text-purple-300 flex items-center gap-1.5 font-mono">
                              📊 BẢNG TỔNG HỢP KẾT QUẢ RÀ QUÉT AI
                            </h5>
                            <span className="text-[10px] bg-purple-950 text-purple-300 border border-purple-500/20 px-2 py-0.5 rounded font-mono">
                              Khớp tự học lệnh của AI
                            </span>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            <div className="bg-[#111827] border border-gray-800 p-3.5 rounded-xl text-center space-y-1">
                              <span className="text-[10px] text-gray-400 font-bold uppercase tracking-wider block">🔍 Tổng số đã quét</span>
                              <strong className="text-xl text-white font-mono">{aiScanMetrics.total.toLocaleString()}</strong>
                              <span className="text-[10px] text-gray-500 block">bản ghi dữ liệu gốc</span>
                            </div>

                            <div className={`border p-3.5 rounded-xl text-center space-y-1 transition-all ${
                              aiScanMetrics.violated > 0 
                                ? "bg-red-950/20 border-red-500/30 text-red-300" 
                                : "bg-gray-900/40 border-gray-800 text-gray-450"
                            }`}>
                              <span className="text-[10px] font-bold uppercase tracking-wider block">❌ Số dòng bị lỗi</span>
                              <strong className={`text-xl font-mono ${aiScanMetrics.violated > 0 ? "text-red-400 font-extrabold" : "text-gray-400 font-normal"}`}>
                                {aiScanMetrics.violated.toLocaleString()}
                              </strong>
                              <span className="text-[10px] block opacity-80">Tỷ lệ: {aiScanMetrics.violatedPercent}</span>
                            </div>

                            <div className="bg-emerald-950/10 border border-emerald-500/15 p-3.5 rounded-xl text-center space-y-1 text-emerald-300">
                              <span className="text-[10px] text-emerald-400 font-bold uppercase tracking-wider block">✅ Số dòng đạt chuẩn</span>
                              <strong className="text-xl text-emerald-400 font-mono">{aiScanMetrics.passed.toLocaleString()}</strong>
                              <span className="text-[10px] text-emerald-500 block">Tỷ lệ: {aiScanMetrics.passedPercent}</span>
                            </div>
                          </div>

                          {/* Bộ lọc hiển thị nhanh cho kết quả quét AI */}
                          <div className="bg-[#1e1b4b]/30 px-4 py-2.5 rounded-xl border border-purple-500/10 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-gray-300">
                            <div>
                              <p className="font-bold text-purple-200">ℹ️ Đang hiển thị trực quan dưới bảng:</p>
                              <p className="text-[11px] text-gray-400 mt-0.5 font-sans">
                                {logicFilterMode === "violated" 
                                  ? `Chỉ hiển thị ${aiScanMetrics.violated} dòng bị phát hiện vi phạm (quy quét được tô nền lỗi đỏ).`
                                  : `Đang hiển thị toàn bộ các dòng được rà quét (Bao gồm cả Đạt ✅ và Vi Phạm ❌).`}
                              </p>
                            </div>
                            <div className="flex gap-2 shrink-0">
                              <button
                                onClick={() => setLogicFilterMode("violated")}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all active:scale-95 cursor-pointer border ${
                                  logicFilterMode === "violated" 
                                    ? "bg-red-600 text-white border-red-500 shadow-sm" 
                                    : "bg-[#111827] text-gray-400 hover:text-white border-gray-700"
                                }`}
                              >
                                ❌ Xem dòng Lỗi ({aiScanMetrics.violated})
                              </button>
                              <button
                                onClick={() => setLogicFilterMode("all")}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all active:scale-95 cursor-pointer border ${
                                  logicFilterMode === "all" 
                                    ? "bg-purple-600 text-white border-purple-500 shadow-sm" 
                                    : "bg-[#111827] text-gray-400 hover:text-white border-gray-700"
                                }`}
                              >
                                🌐 Xem tất cả ({aiScanMetrics.total})
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* DANH SÁCH CÁC QUY TẮC ĐÃ LƯU (HỌC LỆNH CHẠY NHANH) */}
                      {savedAiRules.length > 0 && (
                        <div className="space-y-2 pt-2 border-t border-purple-500/10">
                          <div className="text-[10px] font-bold text-purple-300 tracking-wider font-mono">⚡ BỘ NHỚ HỌC LỆNH THÔNG MINH (BẤM NÚT CHẠY LUÔN KHÔNG CẦN CHỜ DỊCH AI):</div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                            {savedAiRules.map((rule) => (
                              <div
                                key={rule.id}
                                onClick={() => {
                                  setAiRulePrompt(rule.prompt);
                                  setAiTranslatedExpression(rule.expression);
                                  handleAiLogicScan(rule.prompt);
                                }}
                                className="group bg-[#111827]/60 hover:bg-purple-950/20 border border-gray-800 hover:border-purple-500/30 p-2.5 rounded-xl cursor-pointer transition-all flex items-center justify-between gap-2 active:scale-[0.99] text-xs text-left"
                                title={`Khẩu lệnh: "${rule.prompt}"\nBiểu thức: ${rule.expression}`}
                              >
                                <div className="min-w-0 flex-1 space-y-0.5">
                                  <div className="font-bold text-gray-200 group-hover:text-purple-300 transition-colors truncate">{rule.name}</div>
                                  <div className="text-[10px] text-gray-500 truncate italic">"{rule.prompt}"</div>
                                </div>
                                <button
                                  onClick={(e) => handleDeleteAiRule(rule.id, e)}
                                  className="text-gray-500 hover:text-red-400 p-1 rounded hover:bg-red-500/10 transition-colors cursor-pointer"
                                  title="Xoá quy tắc này"
                                >
                                  🗑️
                                </button>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* HOẶC THIẾT LẬP THỦ CÔNG */}
                    <div className="relative flex py-2 items-center">
                      <div className="flex-grow border-t border-gray-800"></div>
                      <span className="flex-shrink mx-4 text-[10px] text-gray-500 font-bold uppercase tracking-wider">HOẶC THIẾT LẬP THỦ CÔNG BẰNG BIỂU MẪU</span>
                      <div className="flex-grow border-t border-gray-800"></div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                      
                      {/* ĐIỀU KIỆN 1: BƯỚC NẾU */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-indigo-500/10 space-y-4">
                        <div className="text-xs font-bold text-indigo-400 uppercase tracking-wider font-mono">BƯỚC 1: XÁC ĐỊNH ĐỐI TƯỢNG CẦN KIỂM TRA</div>
                        <p className="text-[11px] text-gray-400">Thiết lập các điều kiện rà soát :</p>
                        
                        <div className="grid grid-cols-3 gap-2">
                          <select 
                            value={newIfRule.col} 
                            onChange={(e) => setNewIfRule({ ...newIfRule, col: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="">Cột</option>
                            {columns.filter(c => !c.startsWith("_")).map(c => <option key={c} value={c}>{c}</option>)}
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
                            newIfRule.isFieldCompare ? (
                              <select
                                value={newIfRule.val}
                                onChange={(e) => setNewIfRule({ ...newIfRule, val: e.target.value })}
                                className="bg-[#1f2937] border border-purple-500/40 rounded-lg px-2 py-1.5 text-xs text-white"
                              >
                                <option value="">Chọn cột so sánh</option>
                                {columns.filter(c => !c.startsWith("_") && c !== newIfRule.col).map(c => <option key={c} value={c}>{c}</option>)}
                              </select>
                            ) : (
                              <input 
                                type="text"
                                placeholder="Giá trị..."
                                value={newIfRule.val}
                                onChange={(e) => setNewIfRule({ ...newIfRule, val: e.target.value })}
                                className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                              />
                            )
                          )}
                        </div>

                        {newIfRule.op !== "trống" && newIfRule.op !== "không trống" && (
                          <div className="flex items-center gap-2 text-[10px] text-gray-400 mt-1">
                            <span>Kiểu so sánh:</span>
                            <button 
                              type="button"
                              onClick={() => setNewIfRule({ ...newIfRule, isFieldCompare: false, val: "" })}
                              className={`px-1.5 py-0.5 rounded transition-all cursor-pointer ${!newIfRule.isFieldCompare ? "bg-indigo-600/20 text-indigo-300 font-bold border border-indigo-500/20" : "bg-gray-800 text-gray-400 hover:text-white"}`}
                            >
                              📝 So sánh với Giá trị tĩnh
                            </button>
                            <button 
                              type="button"
                              onClick={() => setNewIfRule({ ...newIfRule, isFieldCompare: true, val: "" })}
                              className={`px-1.5 py-0.5 rounded transition-all cursor-pointer ${newIfRule.isFieldCompare ? "bg-indigo-600/20 text-indigo-300 font-bold border border-indigo-500/20" : "bg-gray-800 text-gray-400 hover:text-white"}`}
                            >
                              🔗 So sánh chéo giữa 2 Cột
                            </button>
                          </div>
                        )}

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
                                <span>
                                  {rule.col} {rule.op} {rule.op !== "trống" && rule.op !== "không trống" ? (rule.isFieldCompare ? `[${rule.val}]` : `'${rule.val}'`) : ""}
                                </span>
                                <button onClick={() => setIfRules(ifRules.filter((_, i) => i !== idx))} className="text-red-400 hover:text-red-300 cursor-pointer">X</button>
                              </div>
                            ))
                          )}
                        </div>
                      </div>

                      {/* ĐIỀU KIỆN 2: BƯỚC THÌ PHẢI */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-emerald-500/10 space-y-4">
                        <div className="text-xs font-bold text-emerald-400 uppercase tracking-wider font-mono">BƯỚC 2: YÊU CẦU DỮ LIỆU PHẢI ĐẠT ĐƯỢC</div>
                        <p className="text-[11px] text-gray-400">Nếu thỏa mãn Bước 1, dữ liệu bắt buộc PHẢI đạt ràng buộc này (<strong>Nhấn để trống Bước 2</strong> để lọc quét trực tiếp theo điều kiện Bước 1):</p>
                        
                        <div className="grid grid-cols-3 gap-2">
                          <select 
                            value={newThenRule.col} 
                            onChange={(e) => setNewThenRule({ ...newThenRule, col: e.target.value })}
                            className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                          >
                            <option value="">Cột</option>
                            {columns.filter(c => !c.startsWith("_")).map(c => <option key={c} value={c}>{c}</option>)}
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
                            newThenRule.isFieldCompare ? (
                              <select
                                value={newThenRule.val}
                                onChange={(e) => setNewThenRule({ ...newThenRule, val: e.target.value })}
                                className="bg-[#1f2937] border border-emerald-500/40 rounded-lg px-2 py-1.5 text-xs text-white"
                              >
                                <option value="">Chọn cột so sánh</option>
                                {columns.filter(c => !c.startsWith("_") && c !== newThenRule.col).map(c => <option key={c} value={c}>{c}</option>)}
                              </select>
                            ) : (
                              <input 
                                type="text"
                                placeholder="Giá trị..."
                                value={newThenRule.val}
                                onChange={(e) => setNewThenRule({ ...newThenRule, val: e.target.value })}
                                className="bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1.5 text-xs text-white"
                              />
                            )
                          )}
                        </div>

                        {newThenRule.op !== "trống" && newThenRule.op !== "không trống" && (
                          <div className="flex items-center gap-2 text-[10px] text-gray-400 mt-1">
                            <span>Kiểu so sánh:</span>
                            <button 
                              type="button"
                              onClick={() => setNewThenRule({ ...newThenRule, isFieldCompare: false, val: "" })}
                              className={`px-1.5 py-0.5 rounded transition-all cursor-pointer ${!newThenRule.isFieldCompare ? "bg-emerald-600/20 text-emerald-300 font-bold border border-emerald-500/20" : "bg-gray-800 text-gray-400 hover:text-white"}`}
                            >
                              📝 So sánh với Giá trị tĩnh
                            </button>
                            <button 
                              type="button"
                              onClick={() => setNewThenRule({ ...newThenRule, isFieldCompare: true, val: "" })}
                              className={`px-1.5 py-0.5 rounded transition-all cursor-pointer ${newThenRule.isFieldCompare ? "bg-emerald-600/20 text-emerald-300 font-bold border border-emerald-500/20" : "bg-gray-800 text-gray-400 hover:text-white"}`}
                            >
                              🔗 So sánh chéo giữa 2 Cột
                            </button>
                          </div>
                        )}

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
                                <span>
                                  {rule.col} {rule.op} {rule.op !== "trống" && rule.op !== "không trống" ? (rule.isFieldCompare ? `[${rule.val}]` : `'${rule.val}'`) : ""}
                                </span>
                                <button onClick={() => setThenRules(thenRules.filter((_, i) => i !== idx))} className="text-red-400 hover:text-red-300 cursor-pointer">X</button>
                              </div>
                            ))
                          )}
                        </div>
                      </div>

                    </div> 

                    {/* BỘ LỰA CHỌN THỂ THỨC VÀ BỘ LỌC HIỂN THỊ KẾT QUẢ TRỰC QUAN */}
                    <div className="bg-[#1e293b]/50 p-4 rounded-xl border border-gray-700/80 grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="text-xs font-bold text-orange-400 block mb-1">🛠️ CHẾ ĐỘ KIỂM TRA:</label>
                        <select
                          value={logicRuleMode}
                          onChange={(e) => setLogicRuleMode(e.target.value as "conflict" | "must_satisfy")}
                          className="w-full bg-[#111827] border border-gray-600 rounded-lg px-2.5 py-1.5 text-xs text-white focus:ring-1 focus:ring-amber-500 font-sans"
                        >
                          <option value="must_satisfy">🔹 "Quy tắc ràng buộc: NẾU [Bước 1] xảy ra ➔ [Bước 2] phải đúng"</option>
                          <option value="conflict">🔸 "Cảnh báo khi dữ liệu vi phạm (Dữ liệu không được phép tồn tại đồng thời)"</option>
                        </select>
                        <p className="text-[10px] text-gray-400 mt-1 font-sans">
                          {logicRuleMode === "must_satisfy" 
                            ? "Hệ thống báo vi phạm khi bản ghi thoả mãn NẾU nhưng KHÔNG thỏa THÌ PHẢI." 
                            : "Hệ thống báo vi phạm khi bản ghi thỏa mãn NẾU và thỏa mãn cả CÓ THÊM (mâu thuẫn hai thuộc tính xảy ra đồng thời)."}
                        </p>
                      </div>

                      <div>
                        <label className="text-xs font-bold text-teal-400 block mb-1">👁️ BỘ LỌC HIỂN THỊ BẢNG KẾT QUẢ:</label>
                        <select
                          value={logicFilterMode}
                          onChange={(e) => setLogicFilterMode(e.target.value as "all" | "if_satisfied" | "violated")}
                          className="w-full bg-[#111827] border border-gray-600 rounded-lg px-2.5 py-1.5 text-xs text-white focus:ring-1 focus:ring-teal-500 font-sans"
                        >
                          <option value="if_satisfied">🎯 "Chỉ lọc những hộ thỏa mãn điều kiện NẾU ở Bước 1"</option>
                          <option value="violated">❌ "Chỉ hiển thị những hộ có lỗi/vi phạm quy tắc logic"</option>
                          <option value="all">🌐 "Hiển thị tất cả mọi hộ đã khớp rà soát" (Có thỏa NẾU hoặc bị lỗi)</option>
                        </select>
                        <p className="text-[10px] text-gray-400 mt-1 font-sans">
                          {logicFilterMode === "if_satisfied" 
                            ? "Chỉ hiển thị các dòng thỏa mãn điều kiện NẾU ở Bước 1 (Ví dụ: kiểm tra đi thuê, chỉ hiện tất cả các hộ đi thuê)." 
                            : logicFilterMode === "violated" 
                            ? "Chỉ hiển thị các dòng bị vi phạm quy tắc logic (dòng có đánh dấu lỗi đỏ)." 
                            : "Hiện toàn bộ danh sách các hộ đã được quét lọt vào quy tắc (hoặc thỏa mãn điều kiện NẾU, hoặc bị vi phạm)."}
                        </p>
                      </div>
                    </div>

                    <button 
                      onClick={handleLogicCheck}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all w-full flex items-center justify-center gap-2 cursor-pointer shadow-md"
                    >
                      <CheckSquare className="w-5 h-5 text-purple-300" /> BẤM ĐỂ BẮT DẦU RÀ SOÁT SÔ LIỆU
                    </button>

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>

              {activeTab === "kiemtralogic" && mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={filteredLogicData} 
                  columns={columns} 
                  title="BẢNG DỮ LIỆU ĐÁNH GIÁ LOGIC" 
                  subtitle={`Bảng xem kết quả rà soát chéo chênh lệch theo bộ lọc đã chọn (${filteredLogicData.length} dòng).`}
                  mapping={mapping}
                />
              )}
            </div>
          </div>

          {/* 11. TAB ĐỌC PDF & CHUYỂN SANG WORD */}
          <div className={activeTab === "pdf2word" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in">
              {activeTab === "pdf2word" && <PdfToWord />}
            </div>
          </div>
          {/* 11. TAB CHỌN MẪU ĐIỀU TRA */}
          {activeTab === "chonmau" && <SampleSelector />}

          {/* 10. TAB TRA CỨU DANH MỤC NGÀNH VSIC CHUẨN */}
          <div className={activeTab === "danhmucvsic" ? "block" : "hidden"}>
            <div className="space-y-6 animate-fade-in">
              {activeTab === "danhmucvsic" && <VsicCatalogExplorer />}
            </div>
          </div>

        </main>
      </div>

      <Analytics />
    </div>
  );
}
