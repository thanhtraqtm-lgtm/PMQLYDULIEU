import React, { useState, useMemo, useEffect } from "react";
import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { getFlexibleValue, normalizeAiExpression, parseCSV, beautifyColumnName, scoreColumnForRole, getUniqueRoleAssignments, parse2DArrayWithSmartHeader } from "./utils/sharedHelpers";
import { AuthProvider, useAuth } from "./context/AuthContext";
import { DataEntry } from "./components/DataEntry";
import { VideoRoom } from "./components/VideoRoom";
import { AdminDashboard } from "./components/AdminDashboard";
import { LogIn, Key, HelpCircle, ShieldAlert, Radio, Users, Shield, CheckCircle } from "lucide-react";
// @ts-ignore
import logoImg from "./image/logo.jpg";
// @ts-ignore
import bannerImg from "./image/banner.jpg";
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



// Helper to calculate normal cumulative distribution function (CDF)
function normalCDF(x: number): number {
  const t = 1 / (1 + 0.2316419 * Math.abs(x));
  const d = 0.3989422804; // 1 / Math.sqrt(2 * Math.PI)
  const p = 1 - d * Math.exp(-x * x / 2) * (
    0.31938153 * t -
    0.356563782 * t * t +
    1.781477937 * Math.pow(t, 3) -
    1.821255978 * Math.pow(t, 4) +
    1.330274429 * Math.pow(t, 5)
  );
  return x >= 0 ? p : 1 - p;
}

// Helper to calculate the p-value of a Chi-Square statistic
function chiSquarePValue(chiSq: number, df: number): number {
  if (df <= 0) return 1.0;
  if (chiSq <= 0) return 1.0;

  if (df === 1) {
    return 2 * (1 - normalCDF(Math.sqrt(chiSq)));
  } else if (df === 2) {
    return Math.exp(-chiSq / 2);
  } else {
    // Wilson-Hilferty transformation of Chi-Square to normal distribution
    const z = (Math.pow(chiSq / df, 1/3) - (1 - 2/(9*df))) / Math.sqrt(2/(9*df));
    return 1 - normalCDF(z);
  }
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

const CHUNK_SIZE = 50000;

async function clearOldChunks(db: IDBDatabase, prefix?: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.openKeyCursor();
    request.onsuccess = (event: any) => {
      const cursor = event.target.result;
      if (cursor) {
        const key = String(cursor.key);
        if (prefix) {
          if (key.startsWith(`${prefix}_chunk_`)) {
            store.delete(key);
          }
        } else {
          if (key.startsWith("mainData_chunk_") || key.startsWith("rawImportedData_chunk_") ||
              key.startsWith("mainData_corp_chunk_") || key.startsWith("rawImportedData_corp_chunk_") ||
              key.startsWith("mainData_individual_chunk_") || key.startsWith("rawImportedData_individual_chunk_")) {
            store.delete(key);
          }
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
  if (numChunks === 0) return;

  const transaction = db.transaction(STORE_NAME, "readwrite");
  const store = transaction.objectStore(STORE_NAME);

  const promises: Promise<void>[] = [];
  for (let i = 0; i < numChunks; i++) {
    const start = i * CHUNK_SIZE;
    const chunk = array.slice(start, start + CHUNK_SIZE);

    promises.push(new Promise<void>((resolve, reject) => {
      const request = store.put(chunk, `${prefix}_chunk_${i}`);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error(`Lỗi lưu mảnh ${prefix} ${i}`));
    }));
  }

  await Promise.all(promises);

  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("Transaction failed"));
  });
}

async function loadArrayInChunks(db: IDBDatabase, prefix: string, totalLength: number): Promise<any[]> {
  const numChunks = Math.ceil(totalLength / CHUNK_SIZE);
  if (numChunks === 0) return [];

  const transaction = db.transaction(STORE_NAME, "readonly");
  const store = transaction.objectStore(STORE_NAME);

  const promises: Promise<any[]>[] = [];
  for (let i = 0; i < numChunks; i++) {
    promises.push(new Promise<any[]>((resolve, reject) => {
      const request = store.get(`${prefix}_chunk_${i}`);
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error || new Error(`Lỗi tải mảnh ${prefix} ${i}`));
    }));
  }

  const chunks = await Promise.all(promises);
  let result: any[] = [];
  for (let i = 0; i < chunks.length; i++) {
    result = result.concat(chunks[i]);
  }
  return result;
}

async function saveAppState(state: AppState, forceSaveData: boolean = false, mode?: "corp" | "individual"): Promise<void> {
  try {
    const db = await openDB();
    const { mainData, rawImportedData, ...meta } = state;
    const activeMode = mode || (localStorage.getItem("vsic_current_data_mode") as "corp" | "individual") || "corp";
    
    // 1. Luôn ghi nhận Metadata (rất nhẹ, < 1ms)
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const metaData = {
        ...meta,
        mainDataLength: mainData.length,
        rawImportedDataLength: rawImportedData.length,
        isChunked: true,
        dataMode: activeMode
      };
      const request = store.put(metaData, `sessionMeta_${activeMode}`);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error("Lỗi lưu metadata"));
    });

    // 2. Chỉ thực hiện lưu chuỗi khối dữ liệu khổng lồ nếu được bật cờ bắt buộc (nhập file mới, tính toán lại, gộp sheet)
    if (forceSaveData) {
      await clearOldChunks(db, `mainData_${activeMode}`);
      await clearOldChunks(db, `rawImportedData_${activeMode}`);
      await saveArrayInChunks(db, `mainData_${activeMode}`, mainData);
      await saveArrayInChunks(db, `rawImportedData_${activeMode}`, rawImportedData);
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

async function loadAppState(mode?: "corp" | "individual"): Promise<AppState | null> {
  try {
    const db = await openDB();
    const activeMode = mode || (localStorage.getItem("vsic_current_data_mode") as "corp" | "individual") || "corp";
    
    // 1. Tải Metadata trước
    const meta: any = await new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(`sessionMeta_${activeMode}`);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Lỗi đọc sessionMeta"));
    });

    if (meta && meta.isChunked) {
      const mainDataLength = meta.mainDataLength || 0;
      const rawImportedDataLength = meta.rawImportedDataLength || 0;
      
      const mainData = await loadArrayInChunks(db, `mainData_${activeMode}`, mainDataLength);
      const rawImportedData = await loadArrayInChunks(db, `rawImportedData_${activeMode}`, rawImportedDataLength);
      
      return {
        ...meta,
        mainData,
        rawImportedData
      };
    }

    // Tương thích ngược: Đọc session thô cũ nếu đang tìm kiếm chế độ Doanh nghiệp
    if (activeMode === "corp") {
      const legacyMeta: any = await new Promise((resolve) => {
        const transaction = db.transaction(STORE_NAME, "readonly");
        const store = transaction.objectStore(STORE_NAME);
        const request = store.get("sessionMeta");
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => resolve(null);
      });
      if (legacyMeta && legacyMeta.isChunked) {
        const mainData = await loadArrayInChunks(db, "mainData", legacyMeta.mainDataLength || 0);
        const rawImportedData = await loadArrayInChunks(db, "rawImportedData", legacyMeta.rawImportedDataLength || 0);
        return {
          ...legacyMeta,
          mainData,
          rawImportedData
        };
      }
    }

    return null;

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
  Save,
  Video,
  ChevronDown,
  ChevronUp,
  Mic,
  MicOff,
  Globe
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
import ComplexCalculations from "./components/ComplexCalculations";
import PdfToWord from "./components/PdfToWord";
import StatisticalOutliers from "./components/StatisticalOutliers";
import SmartCatalogMatcher from "./components/SmartCatalogMatcher";
import RulesStudio from "./components/RulesStudio";
import { FrequencyAnalysis } from "./components/FrequencyAnalysis";
import { CorrelationAnalysis } from "./components/CorrelationAnalysis";
import LogicChecking from "./components/LogicChecking";
import SamplingSelection from "./components/SamplingSelection";
import FileMerger from "./components/FileMerger";
import DataComparison from "./components/DataComparison";
import { PIPELINE_STEPS, GUIDE_SCENARIOS } from "./data/userGuides";

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

export function AuthScreen() {
  const { login, register, loading: authLoading, setForceOfflineMode, forceOffline } = useAuth();
  
  const [isRegisterMode, setIsRegisterMode] = useState(false);
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authUnitID, setAuthUnitID] = useState("");
  const [authDisplayName, setAuthDisplayName] = useState("");
  const [authRole, setAuthRole] = useState<"admin" | "user">("user");
  const [authError, setAuthError] = useState<string | null>(null);
  const [authSuccess, setAuthSuccess] = useState<string | null>(null);
  const [isOperationNotAllowed, setIsOperationNotAllowed] = useState(false);

  const handleAuthSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError(null);
    setAuthSuccess(null);
    setIsOperationNotAllowed(false);

    if (!authEmail.trim() || !authPassword.trim()) {
      setAuthError("Vui lòng điền đầy đủ tài khoản và mật khẩu!");
      return;
    }

    try {
      if (isRegisterMode) {
        if (!authUnitID.trim() || !authDisplayName.trim()) {
          setAuthError("Đăng ký yêu cầu nhập thêm Mã đơn vị và Tên hiển thị!");
          return;
        }
        await register(
          authEmail, 
          authPassword, 
          authUnitID, 
          authDisplayName, 
          authRole
        );
        setAuthSuccess("Đăng ký tài khoản thành công! Đã tự động kết nối.");
      } else {
        await login(authEmail, authPassword);
        setAuthSuccess("Đăng nhập thành công!");
      }
    } catch (err: any) {
      console.error(err);
      const errMsg = err.message || String(err);
      if (errMsg.includes("operation-not-allowed") || err.code === "auth/operation-not-allowed") {
        setIsOperationNotAllowed(true);
        setAuthError(
          "Lỗi Firebase (auth/operation-not-allowed): Tính năng Đăng nhập Email/Mật khẩu chưa được kích hoạt trong trang quản trị Firebase Console."
        );
      } else {
        setAuthError(errMsg || "Xử lý xác thực thất bại. Vui lòng thử lại!");
      }
    }
  };

  const handleQuickMockLogin = async (presetType: "hanoi" | "hcm" | "admin") => {
    setAuthError(null);
    setAuthSuccess(null);
    setIsOperationNotAllowed(false);
    try {
      // Vì đây là các nút đăng nhập nhanh mô phỏng, tự động bật offline mode để người dùng trải nghiệm ngay lập tức
      setForceOfflineMode(true);
      
      if (presetType === "admin") {
        await login("admin@chinhphu.gov.vn", "123456");
      } else if (presetType === "hanoi") {
        await login("donvi_hanoi@tphcm.gov.vn", "123456", "hanoi");
      } else if (presetType === "hcm") {
        await login("donvi_saigon@tphcm.gov.vn", "123456", "saigon");
      }
    } catch (err: any) {
      console.error("Lỗi đăng nhập nhanh:", err);
      const errMsg = err.message || String(err);
      if (errMsg.includes("operation-not-allowed") || err.code === "auth/operation-not-allowed") {
        setIsOperationNotAllowed(true);
        setAuthError(
          "Lỗi Firebase (auth/operation-not-allowed): Tính năng Đăng nhập Email/Mật khẩu chưa được kích hoạt trong trang quản trị Firebase Console."
        );
      } else {
        setAuthError(errMsg || "Lỗi đăng nhập nhanh mô phỏng.");
      }
    }
  };

  return (
    <div className="min-h-screen bg-slate-900 flex flex-col justify-center items-center p-4 font-sans text-slate-100">
      
      {/* Card xác thực trung tâm */}
      <div className="w-full max-w-md bg-slate-950 border border-slate-800 rounded-2xl p-8 space-y-6 shadow-2xl relative overflow-hidden">
        
        <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 via-sky-500 to-emerald-500" />

        {/* Tiêu đề */}
        <div className="text-center space-y-2">
          <div className="inline-flex p-3 bg-indigo-600/10 border border-indigo-500/20 text-indigo-400 rounded-2xl mb-2">
            <Radio className="w-6 h-6 animate-pulse" />
          </div>
          <h1 className="text-lg font-extrabold tracking-tight text-white uppercase">
            Cổng Kết Nối Liên Ngành
          </h1>
          <p className="text-xs text-slate-400">
            Nhập liệu &amp; Hội họp video trực tuyến qua Cloud Firebase &amp; Agora
          </p>
        </div>

        {/* 🔓 DIRECT BYPASS BUTTON FOR FULL OPEN PUBLIC ACCESS */}
        <div className="bg-emerald-500/10 border border-emerald-500/20 rounded-xl p-4 text-center space-y-3">
          <p className="text-xs text-emerald-300 leading-relaxed font-semibold">
            🔓 Cổng thông tin liên ngành đã được cấu hình <strong className="text-white">MỞ TỰ DO</strong> theo yêu cầu của bạn. Không bắt buộc phải đăng nhập!
          </p>
          <button
            type="button"
            onClick={async () => {
              const defaultAdmin = {
                uid: "admin_open_default_uid",
                email: "thanhtraqtm@gmail.com",
                unitID: "admin_central",
                role: "admin",
                displayName: "Quản trị viên Trung ương (Mở)",
                isMock: true
              };
              localStorage.setItem("system_auth_user", JSON.stringify(defaultAdmin));
              window.location.reload();
            }}
            className="w-full py-2.5 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 active:scale-95 transition-all font-black text-white rounded-xl text-xs shadow-lg shadow-emerald-900/20 cursor-pointer uppercase tracking-wider flex items-center justify-center gap-1.5"
          >
            <span>👉 BẤM VÀO ĐÂY ĐỂ VÀO THẲNG ỨNG DỤNG</span>
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleAuthSubmit} className="space-y-4">
          
          {authError && (
            <div className="p-4 bg-rose-500/10 border border-rose-500/20 text-rose-300 rounded-xl text-xs space-y-2">
              <div className="flex items-center gap-2 font-bold">
                <ShieldAlert className="w-4 h-4 shrink-0" />
                <span>{authError}</span>
              </div>
              {isOperationNotAllowed && (
                <div className="text-[11px] text-slate-350 leading-relaxed pt-2 border-t border-rose-500/10 space-y-1.5">
                  <p className="font-bold text-amber-400">💡 CÁCH KHẮC PHỤC:</p>
                  <ol className="list-decimal pl-4 space-y-1">
                    <li>Mở trang quản trị <b>Firebase Console</b> của dự án này.</li>
                    <li>Vào mục <b>Authentication</b> &gt; chọn tab <b>Sign-in method</b>.</li>
                    <li>Tìm <b>Email/Password</b> và nhấn <b>Enable</b> (Bật) rồi nhấn Lưu.</li>
                  </ol>
                  <div className="pt-2">
                    <button
                      type="button"
                      onClick={() => {
                        setForceOfflineMode(true);
                        setAuthSuccess("Đã kích hoạt chế độ Mô phỏng Offline thành công!");
                        setAuthError(null);
                        setIsOperationNotAllowed(false);
                      }}
                      className="w-full py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-lg text-xs transition cursor-pointer shadow-md animate-bounce mt-1"
                    >
                      🚀 Bật chế độ Mô Phỏng Offline để chạy ngay
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {authSuccess && (
            <div className="p-3 bg-emerald-500/10 border border-emerald-500/20 text-emerald-300 rounded-xl text-xs flex items-center gap-2">
              <CheckCircle className="w-4 h-4 shrink-0" />
              <span>{authSuccess}</span>
            </div>
          )}

          <div className="space-y-3">
            <div>
              <label className="text-[10px] font-bold text-slate-400 block mb-1 uppercase tracking-wider">Email truy cập</label>
              <input
                type="email"
                required
                placeholder="donvi_hanoi@tphcm.gov.vn hoặc admin@..."
                value={authEmail}
                onChange={e => setAuthEmail(e.target.value)}
                className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition"
              />
            </div>

            <div>
              <label className="text-[10px] font-bold text-slate-400 block mb-1 uppercase tracking-wider">Mật khẩu</label>
              <input
                type="password"
                required
                placeholder="••••••••"
                value={authPassword}
                onChange={e => setAuthPassword(e.target.value)}
                className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition"
              />
            </div>

            {isRegisterMode && (
              <>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1 uppercase tracking-wider">Mã đơn vị</label>
                    <input
                      type="text"
                      required
                      placeholder="hanoi, saigon,..."
                      value={authUnitID}
                      onChange={e => setAuthUnitID(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1 uppercase tracking-wider">Quyền hạn</label>
                    <select
                      value={authRole}
                      onChange={e => setAuthRole(e.target.value as any)}
                      className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition"
                    >
                      <option value="user">Đơn vị nhập liệu</option>
                      <option value="admin">Quản trị cấp cao</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-400 block mb-1 uppercase tracking-wider">Tên hiển thị</label>
                  <input
                    type="text"
                    required
                    placeholder="Sở KH & ĐT Hà Nội..."
                    value={authDisplayName}
                    onChange={e => setAuthDisplayName(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition"
                  />
                </div>
              </>
            )}
          </div>

          <button
            type="submit"
            disabled={authLoading}
            className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-500 active:translate-y-0.5 transition font-bold text-white rounded-xl text-sm cursor-pointer shadow-lg shadow-indigo-600/10 flex items-center justify-center gap-1.5"
          >
            {isRegisterMode ? <Key className="w-4 h-4" /> : <LogIn className="w-4 h-4" />}
            {isRegisterMode ? "Tạo tài khoản mới" : "Xác thực danh tính"}
          </button>
        </form>

        {/* Toggle chế độ */}
        <div className="text-center">
          <button
            onClick={() => setIsRegisterMode(!isRegisterMode)}
            className="text-xs text-indigo-400 hover:text-indigo-300 hover:underline transition"
          >
            {isRegisterMode ? "Đã có tài khoản? Quay về Đăng nhập" : "Chưa có tài khoản? Đăng ký tại đây"}
          </button>
        </div>

        {/* Khối Đăng Nhập Nhanh Mô Phỏng (Phòng trường hợp không có kết nối Firebase của người dùng) */}
        <div className="border-t border-slate-800/80 pt-4 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Trải nghiệm nhanh (Demo Presets)</span>
            <span className="px-2 py-0.5 bg-amber-500/10 text-amber-400 border border-amber-500/20 rounded text-[9px] font-bold">Mô Phỏng Offline</span>
          </div>
          
          <div className="grid grid-cols-3 gap-2">
            <button
              onClick={() => handleQuickMockLogin("hanoi")}
              className="py-1.5 bg-slate-900 hover:bg-slate-850 border border-slate-850 rounded-lg text-[10px] font-bold text-slate-300 hover:text-white transition cursor-pointer"
            >
              Đơn vị Hà Nội
            </button>
            <button
              onClick={() => handleQuickMockLogin("hcm")}
              className="py-1.5 bg-slate-900 hover:bg-slate-850 border border-slate-850 rounded-lg text-[10px] font-bold text-slate-300 hover:text-white transition cursor-pointer"
            >
              Đơn vị Sài Gòn
            </button>
            <button
              onClick={() => handleQuickMockLogin("admin")}
              className="py-1.5 bg-indigo-950/40 hover:bg-indigo-950/60 border border-indigo-900/30 rounded-lg text-[10px] font-bold text-indigo-300 hover:text-indigo-200 transition cursor-pointer"
            >
              Cấp Admin
            </button>
          </div>
          <p className="text-[9px] text-slate-500 leading-relaxed text-center">
            Nhấn các nút trên để trải nghiệm mượt mà lập tức mà không cần nạp Firebase Config thật.
          </p>
        </div>

      </div>

    </div>
  );
}

const getStepIcon = (id: number) => {
  switch (id) {
    case 1: return <FileSpreadsheet className="w-5 h-5" />;
    case 2: return <Brain className="w-5 h-5" />;
    case 3: return <Sliders className="w-5 h-5" />;
    case 4: return <BarChart3 className="w-5 h-5" />;
    default: return <Compass className="w-5 h-5" />;
  }
};

const getStepColors = (id: number, isSelected: boolean) => {
  if (isSelected) {
    switch (id) {
      case 1: return 'bg-indigo-50/70 border-indigo-500 shadow-[0_4px_12px_rgba(79,70,229,0.08)] ring-1 ring-indigo-500';
      case 2: return 'bg-emerald-50/70 border-emerald-500 shadow-[0_4px_12px_rgba(16,185,129,0.08)] ring-1 ring-emerald-500';
      case 3: return 'bg-teal-50/70 border-teal-500 shadow-[0_4px_12px_rgba(20,184,166,0.08)] ring-1 ring-teal-500';
      case 4: return 'bg-violet-50/70 border-violet-500 shadow-[0_4px_12px_rgba(139,92,246,0.08)] ring-1 ring-violet-500';
      default: return 'bg-indigo-50/70 border-indigo-500';
    }
  } else {
    switch (id) {
      case 1: return 'bg-slate-50/50 border-slate-200 hover:border-indigo-300 hover:bg-slate-50/80';
      case 2: return 'bg-slate-50/50 border-slate-200 hover:border-emerald-300 hover:bg-slate-50/80';
      case 3: return 'bg-slate-50/50 border-slate-200 hover:border-teal-300 hover:bg-slate-50/80';
      case 4: return 'bg-slate-50/50 border-slate-200 hover:border-violet-300 hover:bg-slate-50/80';
      default: return 'bg-slate-50/50 border-slate-200 hover:border-indigo-300';
    }
  }
};

const getIconBadgeStyles = (id: number, isSelected: boolean) => {
  if (isSelected) {
    switch (id) {
      case 1: return 'bg-indigo-600 text-white border-indigo-600';
      case 2: return 'bg-emerald-600 text-white border-emerald-600';
      case 3: return 'bg-teal-600 text-white border-teal-600';
      case 4: return 'bg-violet-600 text-white border-violet-600';
      default: return 'bg-indigo-600 text-white';
    }
  } else {
    switch (id) {
      case 1: return 'bg-white text-indigo-600 border-slate-200';
      case 2: return 'bg-white text-emerald-600 border-slate-200';
      case 3: return 'bg-white text-teal-600 border-slate-200';
      case 4: return 'bg-white text-violet-600 border-slate-200';
      default: return 'bg-white text-indigo-600';
    }
  }
};

const getActiveDotColor = (id: number) => {
  switch (id) {
    case 1: return 'bg-indigo-600';
    case 2: return 'bg-emerald-600';
    case 3: return 'bg-teal-600';
    case 4: return 'bg-violet-600';
    default: return 'bg-indigo-600';
  }
};

const getDetailStepBgColor = (id: number) => {
  switch (id) {
    case 1: return 'bg-indigo-50/30 border-indigo-100';
    case 2: return 'bg-emerald-50/30 border-emerald-100';
    case 3: return 'bg-teal-50/30 border-teal-100';
    case 4: return 'bg-violet-50/30 border-violet-100';
    default: return 'bg-indigo-50/30 border-indigo-100';
  }
};

const getDetailStepBadgeColor = (id: number) => {
  switch (id) {
    case 1: return 'bg-indigo-100 text-indigo-750';
    case 2: return 'bg-emerald-100 text-emerald-750';
    case 3: return 'bg-teal-100 text-teal-750';
    case 4: return 'bg-violet-100 text-violet-750';
    default: return 'bg-indigo-100 text-indigo-750';
  }
};

const getDetailActionBtnColor = (id: number) => {
  switch (id) {
    case 1: return 'bg-indigo-600 hover:bg-indigo-700 text-white border-0';
    case 2: return 'bg-emerald-600 hover:bg-emerald-700 text-white border-0';
    case 3: return 'bg-teal-600 hover:bg-teal-700 text-white border-0';
    case 4: return 'bg-violet-600 hover:bg-violet-700 text-white border-0';
    default: return 'bg-indigo-600 hover:bg-indigo-700 text-white border-0';
  }
};

export function MainAppContent() {
  const { user, logout } = useAuth();

  const [activeTab, setActiveTab] = useState<string>("trangchu");
  const [totalQueries, setTotalQueries] = useState<number>(() => Math.floor(Math.random() * 150) + 2480);

  // Tăng số lượt thao tác khi người dùng chuyển tab hoặc tương tác
  useEffect(() => {
    setTotalQueries(prev => prev + Math.floor(Math.random() * 3) + 1);
  }, [activeTab]);
  const [selectedPipelineStep, setSelectedPipelineStep] = useState<number | null>(1);
  const [selectedWizardScenario, setSelectedWizardScenario] = useState<string | null>("naptiep");
  const [tongHopSubTab, setTongHopSubTab] = useState<"goc" | "phu" | "phep_tinh" | "so_sanh">("goc");
  const [loading, setLoading] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string>("");

  const [collapsedBlocks, setCollapsedBlocks] = useState<{[key: string]: boolean}>({
    thaotac: false,
    rasoat: false,
    vsic: false,
    phantich: false,
    utilities: false,
    congtac: false, // Thêm block cộng tác liên ngành
  });

  // Trạng thái modal xác nhận xóa custom (để tránh lỗi window.confirm trong iframe)
  const [customConfirmModal, setCustomConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    note?: string;
    confirmText?: string;
    onConfirm: () => void;
  } | null>(null);

  const toggleBlock = (blockKey: string) => {
    setCollapsedBlocks(prev => ({
      ...prev,
      [blockKey]: !prev[blockKey]
    }));
  };

  const [openDropdown, setOpenDropdown] = useState<string | null>(null);

  // Close dropdown on click outside
  useEffect(() => {
    const handleOutsideClick = () => {
      setOpenDropdown(null);
    };
    window.addEventListener("click", handleOutsideClick);
    return () => {
      window.removeEventListener("click", handleOutsideClick);
    };
  }, []);

  // TỰ ĐỘNG ĐIỀU HƯỚNG THEO UNIT ID / ROLE KHI ĐĂNG NHẬP - Giữ Trang Chủ làm mặc định hiển thị đầu tiên khi mở web để đọc trợ lý
  useEffect(() => {
    setActiveTab("trangchu");
  }, [user]);

  useEffect(() => {
    if (["trangchu", "xemdulieu", "ghepnoi", "tachfile", "chonmau"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, thaotac: false }));
    } else if (["kiemtralogic", "sosanh"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, rasoat: false }));
    } else if (["chuanhoanganh", "danhmucvsic"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, vsic: false }));
    } else if (["tonghop", "tansuat", "tuongquan"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, phantich: false }));
    } else if (["pdf2word"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, utilities: false }));
    } else if (["dataentry", "videoroom", "admindashboard"].includes(activeTab)) {
      setCollapsedBlocks(prev => ({ ...prev, congtac: false }));
    }
  }, [activeTab]);

  // Trạng thái mật khẩu bảo vệ ứng dụng độc lập tránh bắt đăng nhập email phiền hà
  const [appPassword, setAppPassword] = useState<string>(() => {
    return localStorage.getItem("vsic_app_password") || "admin123";
  });
  
  // Trạng thái phân quyền tài khoản: "admin" (toàn quyền) hoặc "shared" (dùng chung, chỉ xem, không đổi được mật khẩu)
  const [userRole, setUserRole] = useState<"admin" | "shared">(() => {
    return (localStorage.getItem("vsic_app_user_role") as "admin" | "shared") || "admin";
  });

  const [isAuthorized, setIsAuthorized] = useState<boolean>(true);
  const [typedPassword, setTypedPassword] = useState<string>("");
  const [passwordError, setPasswordError] = useState<string>("");
  const [showPasswordChangeModal, setShowPasswordChangeModal] = useState<boolean>(false);
  const [newPasswordVal, setNewPasswordVal] = useState<string>("");

  const handleCheckPassword = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const cleanTyped = typedPassword.trim();
    if (cleanTyped === appPassword) {
      localStorage.setItem("vsic_app_authorized", "true");
      localStorage.setItem("vsic_app_user_role", "admin");
      setUserRole("admin");
      setIsAuthorized(true);
      setPasswordError("");
    } else if (cleanTyped === "user123") { // Mật khẩu tài khoản dùng chung cố định, không thể đổi từ UI
      localStorage.setItem("vsic_app_authorized", "true");
      localStorage.setItem("vsic_app_user_role", "shared");
      setUserRole("shared");
      setIsAuthorized(true);
      setPasswordError("");
    } else {
      setPasswordError("Mật khẩu truy cập chưa chính xác! Vui lòng kiểm tra lại.");
    }
  };

  const handleChangePassword = () => {
    if (userRole === "shared") {
      alert("Tài khoản dùng chung không có quyền thay đổi mật khẩu!");
      return;
    }
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
  
  // Trạng thái Ghi âm Microphone (Web Speech API) cho Cột định nghĩa
  const [isRecordingColMic, setIsRecordingColMic] = useState(false);

  // Chế độ dữ liệu làm việc độc lập (Doanh nghiệp - "corp" / Cá thể - "individual")
  const [dataMode, setDataMode] = useState<"corp" | "individual">("corp");

  const toggleMicCol = () => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("Trình duyệt không hỗ trợ nhận diện giọng nói (Web Speech API). Hãy dùng Google Chrome hoặc Microsoft Edge.");
      return;
    }

    if (isRecordingColMic) {
      setIsRecordingColMic(false);
      return;
    }

    const rec = new SpeechRecognition();
    rec.lang = "vi-VN";
    rec.interimResults = false;
    rec.maxAlternatives = 1;

    rec.onstart = () => {
      setIsRecordingColMic(true);
    };

    rec.onerror = (e: any) => {
      console.error(e);
      setIsRecordingColMic(false);
    };

    rec.onend = () => {
      setIsRecordingColMic(false);
    };

    rec.onresult = (event: any) => {
      const transcript = event.results[0][0].transcript;
      if (transcript) {
        setAiColLearnPrompt(prev => {
          const trimmed = prev.trim();
          return trimmed ? `${trimmed} ${transcript}` : transcript;
        });
      }
    };

    rec.start();
  };

  // Sampling states lifted from SamplingSelection.tsx
  const [sampCorpData, setSampCorpData] = useState<any[]>([]);
  const [sampCorpFileName, setSampCorpFileName] = useState<string>("");
  const [sampIndData, setSampIndData] = useState<any[]>([]);
  const [sampIndFileName, setSampIndFileName] = useState<string>("");

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

  // Trạng thái quản lý biểu tổng hợp song song, tính toán liên cột & đối sánh liên năm (YoY)
  interface CompiledDataset {
    id: string;
    name: string;
    rows: any[];
    cols: string[];
    level: number;
    reportType: "pivot" | "flat";
  }

  const [compiledDatasets, setCompiledDatasets] = useState<CompiledDataset[]>([]);

  const [secondaryFile, setSecondaryFile] = useState<{
    name: string;
    data: any[];
    columns: string[];
  } | null>(null);

  const [secManganhCol, setSecManganhCol] = useState<string>("");
  const [secXaCol, setSecXaCol] = useState<string>("");
  const [secSumCols, setSecSumCols] = useState<string[]>([]);
  const [secReportType, setSecReportType] = useState<"flat" | "pivot">("pivot");

  // Cấu hình phép tính cộng trừ nhân chia cột giữa 2 biểu tổng hợp
  const [mathDsA, setMathDsA] = useState<string>("");
  const [mathDsB, setMathDsB] = useState<string>("");
  const [mathColA, setMathColA] = useState<string>("");
  const [mathColB, setMathColB] = useState<string>("");
  const [mathOp, setMathOp] = useState<"+" | "-" | "*" | "/">("+");
  const [mathNewColName, setMathNewColName] = useState<string>("Hiệu_Chênh_Lệch_1");

  const [mathColA2, setMathColA2] = useState<string>("");
  const [mathColB2, setMathColB2] = useState<string>("");
  const [mathOp2, setMathOp2] = useState<"+" | "-" | "*" | "/">("-");
  const [mathNewColName2, setMathNewColName2] = useState<string>("Hiệu_Chênh_Lệch_2");

  const [mathColA3, setMathColA3] = useState<string>("");
  const [mathColB3, setMathColB3] = useState<string>("");
  const [mathOp3, setMathOp3] = useState<"+" | "-" | "*" | "/">("-");
  const [mathNewColName3, setMathNewColName3] = useState<string>("Hiệu_Chênh_Lệch_3");

  const [mathTreatMissingAsZero, setMathTreatMissingAsZero] = useState<boolean>(true);

  // States mới cho nạp đa tệp tin, tính toán liên file và lưu lệnh
  const [aggregateFiles, setAggregateFiles] = useState<{
    id: string;
    name: string;
    data: any[];
    columns: string[];
  }[]>([]);
  const [selectedFileIdToAggregate, setSelectedFileIdToAggregate] = useState<string>("main_data_file");
  const [mathFileAId, setMathFileAId] = useState<string>("");
  const [mathFileBId, setMathFileBId] = useState<string>("");
  const [mathKeyA, setMathKeyA] = useState<string>("");
  const [mathKeyB, setMathKeyB] = useState<string>("");
  const [mathKeyA2, setMathKeyA2] = useState<string>("");
  const [mathKeyB2, setMathKeyB2] = useState<string>("");
  const [mathFilterA, setMathFilterA] = useState<string>("");
  const [mathFilterB, setMathFilterB] = useState<string>("");

  interface SavedTongHopCommand {
    id: string;
    name: string;
    selectedFileIdToAggregate?: string;
    quickReportManganhCol?: string;
    quickReportXaCol?: string;
    quickReportSumCols?: string[];
    reportType?: "pivot" | "flat";
    mathFileAId?: string;
    mathFileBId?: string;
    mathColA?: string;
    mathColB?: string;
    mathOp?: "+" | "-" | "*" | "/";
    mathNewColName?: string;
    mathKeyA?: string;
    mathKeyB?: string;
  }

  const [savedTongHopCommands, setSavedTongHopCommands] = useState<SavedTongHopCommand[]>(() => {
    try {
      const saved = localStorage.getItem("savedTongHopCommands");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      console.warn("Lỗi đọc savedTongHopCommands từ localStorage:", e);
      return [];
    }
  });

  const allAvailableFiles = useMemo(() => {
    const list = [...aggregateFiles];
    if (mainData && mainData.length > 0) {
      list.unshift({
        id: "main_data_file",
        name: `📂 [TỆP CHÍNH] ${fileName || "Dữ liệu nguồn chính"}`,
        data: mainData,
        columns: columns
      });
    }
    return list;
  }, [mainData, fileName, columns, aggregateFiles]);

  // Cấu hình đối sánh đa niên độ (YoY)
  const [compareDsIds, setCompareDsIds] = useState<string[]>([]);
  const [compareKeyCol, setCompareKeyCol] = useState<string>("Địa_Bàn_Xã");
  const [compareColMapping, setCompareColMapping] = useState<{ [dsId: string]: string }>({});
  const [selectedCompareRowKey, setSelectedCompareRowKey] = useState<string>("");

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
    // Danh mục lệnh học mặc định
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

  // Trí tuệ Nhân tạo - Học và lưu lệnh qua AI
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

  // State người dùng chọn cột báo cáo nhanh và tổng hợp
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

  // Kết quả so sánh và ghép nối dữ liệu hiển thị tức thì dưới tab tương ứng
  const [mergedResultData, setMergedResultData] = useState<any[] | null>(null);
  const [compareResultData, setCompareResultData] = useState<any[] | null>(null);

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

  // --- STATES CHO CÔNG CỤ ĐA NĂNG ĐỈNH CAO ---
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [selectedRows, setSelectedRows] = useState<number[]>([]);
  const [aiCommandText, setAiCommandText] = useState<string>("");
  const [aiCommandResult, setAiCommandResult] = useState<any | null>(null);
  const [showAnalyticsDropdown, setShowAnalyticsDropdown] = useState<boolean>(true);
  
  // Local macro store (localStorage)
  const [savedMacros, setSavedMacros] = useState<{ id: string; name: string; columns: string[]; command: string; createdAt: string }[]>(() => {
    try {
      const saved = localStorage.getItem("vsic_use_macro_store");
      return saved ? JSON.parse(saved) : [
        {
          id: "macro_sample_1",
          name: "Mẫu: Tổng doanh thu toàn bộ",
          columns: ["DoanhThu", "Doanh_Thu_Tích_Lũy"],
          command: "tính tổng doanh thu",
          createdAt: "29/06/2026"
        },
        {
          id: "macro_sample_2",
          name: "Mẫu: Tần suất địa bàn xã",
          columns: ["Địa_Bàn_Xã"],
          command: "tần suất xã",
          createdAt: "29/06/2026"
        }
      ];
    } catch (e) {
      return [];
    }
  });

  useEffect(() => {
    localStorage.setItem("vsic_use_macro_store", JSON.stringify(savedMacros));
  }, [savedMacros]);

  // States cho phân tích thống kê
  const [tsSelectedCol, setTsSelectedCol] = useState<string>("");
  const [isTsCalculated, setIsTsCalculated] = useState<boolean>(false);
  const [isTqCalculated, setIsTqCalculated] = useState<boolean>(false);
  const [tqSelectedCol1, setTqSelectedCol1] = useState<string>("");
  const [tqSelectedCol2, setTqSelectedCol2] = useState<string>("");

  // States cho Bảng Chéo Tương Quan Hộ
  const [tqSubTab, setTqSubTab] = useState<"bang_cheo" | "tuyen_tinh">("bang_cheo");
  const [tqHangCol, setTqHangCol] = useState<string>("");
  const [tqCotCol, setTqCotCol] = useState<string>("");
  const [tqSearchTerm, setTqSearchTerm] = useState<string>("");
  const [tqShowResults, setTqShowResults] = useState<boolean>(true);
  
  // States cho modal hiển thị danh sách hộ của ô tương quan
  const [tqSelectedCell, setTqSelectedCell] = useState<{ hang: string | null; cot: string | null } | null>(null);
  const [tqModalSearchTerm, setTqModalSearchTerm] = useState<string>("");
  const [tqModalPage, setTqModalPage] = useState<number>(1);

  useEffect(() => {
    if (columns.length > 0) {
      if (!tqHangCol) {
        const defaultHang = columns.find(c => c.toLowerCase().includes("đất") || c.toLowerCase().includes("hoạt động")) || columns[0];
        setTqHangCol(defaultHang);
      }
      if (!tqCotCol) {
        const defaultCot = columns.find(c => (c.toLowerCase().includes("cây") || c.toLowerCase().includes("trồng") || c.toLowerCase().includes("thu")) && c !== (tqHangCol || columns[0])) || (columns[1] || columns[0]);
        setTqCotCol(defaultCot);
      }
    }
  }, [columns, tqHangCol, tqCotCol]);

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

  // --- HÀM XỬ LÝ LỆNH PHÂN TÍCH (AI COMMAND BAR) ---
  const processCommand = (cmdText: string, data: any[], cols: string[]) => {
    const cleanCmd = cmdText.trim().toLowerCase();
    
    // Try to find columns mentioned in the command text
    const matchedCols: string[] = [];
    cols.forEach(col => {
      const colNorm = col.toLowerCase().replace(/_/g, " ");
      const cmdNorm = cleanCmd.replace(/_/g, " ");
      if (cmdNorm.includes(colNorm) || cmdNorm.includes(col.toLowerCase())) {
        matchedCols.push(col);
      }
    });

    if (matchedCols.length === 0) {
      cols.forEach(col => {
        const colSub = col.toLowerCase();
        if (colSub.length >= 3 && cleanCmd.includes(colSub)) {
          matchedCols.push(col);
        }
      });
    }

    // 1. FILLING NULL / MISSING VALUES (Điền khuyết)
    if (cleanCmd.includes("điền") || cleanCmd.includes("fill") || cleanCmd.includes("khuyết") || cleanCmd.includes("trống")) {
      const targetCol = matchedCols[0] || cols[0];
      if (!targetCol) {
        return {
          success: false,
          message: "Không xác định được cột đích để thực hiện điền khuyết."
        };
      }
      
      let fillValue: any = 0;
      const numberMatch = cleanCmd.match(/bằng\s+([0-9.-]+)/) || cleanCmd.match(/fill\s+([0-9.-]+)/) || cleanCmd.match(/([0-9.-]+)$/);
      if (numberMatch) {
        fillValue = parseFloat(numberMatch[1]);
      } else {
        const quoteMatch = cleanCmd.match(/"([^"]+)"/) || cleanCmd.match(/'([^']+)'/);
        if (quoteMatch) {
          fillValue = quoteMatch[1];
        }
      }

      let modifiedCount = 0;
      const modifiedData = data.map(row => {
        const val = row[targetCol];
        if (val === null || val === undefined || String(val).trim() === "") {
          modifiedCount++;
          return { ...row, [targetCol]: fillValue };
        }
        return row;
      });

      return {
        success: true,
        commandType: "fill_null",
        message: `Đã hoàn thành điền khuyết cột [${targetCol}]!`,
        details: `Đã quét ${data.length} dòng dữ liệu, tìm thấy ${modifiedCount} ô trống và điền bằng giá trị [${fillValue}].`,
        modifiedData,
        summary: {
          col: targetCol,
          fillValue,
          modifiedCount,
          total: data.length
        }
      };
    }

    // 2. SUM / AVG / MIN / MAX / COUNT (Tính toán thống kê)
    const isSum = cleanCmd.includes("tổng") || cleanCmd.includes("sum");
    const isAvg = cleanCmd.includes("trung bình") || cleanCmd.includes("avg") || cleanCmd.includes("average");
    const isMax = cleanCmd.includes("lớn nhất") || cleanCmd.includes("max") || cleanCmd.includes("cao nhất");
    const isMin = cleanCmd.includes("nhỏ nhất") || cleanCmd.includes("min") || cleanCmd.includes("thấp nhất");
    
    if (isSum || isAvg || isMax || isMin) {
      const targetCol = matchedCols[0] || cols.find(c => {
        const lower = c.toLowerCase();
        return lower.includes("doanh") || lower.includes("thu") || lower.includes("lao") || lower.includes("dong") || lower.includes("so") || lower.includes("luong");
      }) || cols[0];

      if (!targetCol) {
        return {
          success: false,
          message: "Không tìm thấy cột dữ liệu số thích hợp để thực hiện phép tính."
        };
      }

      let sum = 0;
      let count = 0;
      let nonNumericCount = 0;
      let blankCount = 0;
      let minVal = Infinity;
      let maxVal = -Infinity;

      data.forEach(row => {
        const rawVal = row[targetCol];
        if (rawVal === null || rawVal === undefined || String(rawVal).trim() === "") {
          blankCount++;
          return;
        }
        const val = parseFloat(String(rawVal).replace(/,/g, ""));
        if (isNaN(val)) {
          nonNumericCount++;
        } else {
          sum += val;
          count++;
          if (val < minVal) minVal = val;
          if (val > maxVal) maxVal = val;
        }
      });

      if (count === 0) {
        return {
          success: false,
          message: `Cột [${targetCol}] không chứa dữ liệu số hợp lệ để tính toán.`,
          details: `Số dòng trống: ${blankCount}, số dòng chứa text không thể tính: ${nonNumericCount}.`
        };
      }

      const avg = sum / count;
      let finalVal = 0;
      let title = "";
      if (isSum) { finalVal = sum; title = "Tổng cộng (Sum)"; }
      else if (isAvg) { finalVal = avg; title = "Trung bình cộng (Average)"; }
      else if (isMax) { finalVal = maxVal; title = "Giá trị lớn nhất (Maximum)"; }
      else if (isMin) { finalVal = minVal; title = "Giá trị nhỏ nhất (Minimum)"; }

      return {
        success: true,
        commandType: "calculate",
        message: `Đã tính toán thành công cho cột [${targetCol}]`,
        details: `Đã rà soát ${data.length} dòng dữ liệu, lấy mẫu ${count} số dòng hợp lệ.`,
        summary: {
          col: targetCol,
          title,
          value: finalVal.toLocaleString("vi-VN", { maximumFractionDigits: 3 }),
          avg: avg.toLocaleString("vi-VN", { maximumFractionDigits: 3 }),
          sum: sum.toLocaleString("vi-VN", { maximumFractionDigits: 3 }),
          min: minVal.toLocaleString("vi-VN", { maximumFractionDigits: 3 }),
          max: maxVal.toLocaleString("vi-VN", { maximumFractionDigits: 3 }),
          count,
          blankCount,
          nonNumericCount
        }
      };
    }

    // 3. FILTERING DỮ LIỆU (Lọc dòng)
    if (cleanCmd.includes("lọc") || cleanCmd.includes("filter")) {
      const targetCol = matchedCols[0] || cols[0];
      if (!targetCol) {
        return {
          success: false,
          message: "Không tìm thấy cột tương ứng để thực hiện bộ lọc."
        };
      }

      let op: ">" | "<" | "=" | "contains" = "contains";
      let compValText = "";

      const cleanNoAccents = cleanCmd.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (cleanCmd.includes(">") || cleanNoAccents.includes("lon hon")) {
        op = ">";
      } else if (cleanCmd.includes("<") || cleanNoAccents.includes("nho hon")) {
        op = "<";
      } else if (cleanCmd.includes("=") || cleanCmd.includes("bằng") || cleanNoAccents.includes("bang")) {
        op = "=";
      }

      const valMatch = cleanCmd.match(/(?:>|<|=|\bbằng\b|bằng|is)\s*([0-9.-]+)/) || cleanCmd.match(/(?:>|<|=|\bbằng\b|bằng|is)\s*"([^"]+)"/) || cleanCmd.match(/\s+([0-9.-]+)$/);
      if (valMatch) {
        compValText = valMatch[1];
      } else {
        const words = cleanCmd.split(/\s+/);
        compValText = words[words.length - 1];
      }

      const filteredIndices: number[] = [];
      data.forEach((row, idx) => {
        const rawVal = row[targetCol];
        if (rawVal === null || rawVal === undefined) return;
        
        const rowValStr = String(rawVal).toLowerCase();
        const compValStr = compValText.toLowerCase();

        if (op === ">") {
          const rNum = parseFloat(rowValStr.replace(/,/g, ""));
          const cNum = parseFloat(compValStr);
          if (!isNaN(rNum) && !isNaN(cNum) && rNum > cNum) {
            filteredIndices.push(idx);
          }
        } else if (op === "<") {
          const rNum = parseFloat(rowValStr.replace(/,/g, ""));
          const cNum = parseFloat(compValStr);
          if (!isNaN(rNum) && !isNaN(cNum) && rNum < cNum) {
            filteredIndices.push(idx);
          }
        } else if (op === "=") {
          if (rowValStr === compValStr) {
            filteredIndices.push(idx);
          }
        } else {
          if (rowValStr.includes(compValStr)) {
            filteredIndices.push(idx);
          }
        }
      });

      return {
        success: true,
        commandType: "filter",
        message: `Đã lọc dữ liệu theo [${targetCol}]`,
        details: `Đã tìm thấy ${filteredIndices.length} dòng thỏa mãn điều kiện lọc: [${targetCol} ${op} ${compValText}].`,
        filteredIndices,
        summary: {
          col: targetCol,
          op,
          val: compValText,
          matchedCount: filteredIndices.length
        }
      };
    }

    // 4. FREQUENCY / COUNT UNIQUE (Tần suất)
    if (cleanCmd.includes("tần suất") || cleanCmd.includes("đếm") || cleanCmd.includes("thống kê") || cleanCmd.includes("frequency") || cleanCmd.includes("tỉ lệ")) {
      const targetCol = matchedCols[0] || cols[0];
      if (!targetCol) {
        return {
          success: false,
          message: "Không xác định được cột đích để thống kê tần suất."
        };
      }

      const counts: { [key: string]: number } = {};
      let totalCount = 0;
      data.forEach(row => {
        const val = row[targetCol];
        const key = val === null || val === undefined || String(val).trim() === "" ? "[Trống / Blank]" : String(val).trim();
        counts[key] = (counts[key] || 0) + 1;
        totalCount++;
      });

      const frequencyList = Object.entries(counts)
        .map(([value, count]) => ({
          value,
          count,
          percent: ((count / totalCount) * 100).toFixed(2) + "%"
        }))
        .sort((a, b) => b.count - a.count);

      return {
        success: true,
        commandType: "frequency",
        message: `Phân tích tần suất cho cột [${targetCol}]`,
        details: `Tìm thấy ${frequencyList.length} giá trị phân biệt trong tổng số ${totalCount} bản ghi.`,
        frequencyList,
        summary: {
          col: targetCol,
          uniqueCount: frequencyList.length,
          total: totalCount
        }
      };
    }

    return {
      success: false,
      message: "Lệnh chưa rõ hoặc chưa được hỗ trợ động cơ.",
      details: "Bạn hãy nhập các cú pháp mẫu như: 'tổng doanh thu', 'tần suất Địa_Bàn_Xã', 'điền khuyết DoanhThu bằng 0', 'lọc doanh thu > 5000000'."
    };
  };

  const handleExecuteCommand = () => {
    if (!aiCommandText.trim()) return;
    if (mainData.length === 0) {
      alert("Vui lòng tải tệp dữ liệu nguồn chính trước khi thực thi lệnh!");
      return;
    }
    const result = processCommand(aiCommandText, mainData, columns);
    setAiCommandResult(result);

    if (result.success) {
      if (result.modifiedData) {
        setMainData(result.modifiedData);
        saveAppState({
          mainData: result.modifiedData,
          rawImportedData,
          columns,
          fileName,
          mapping,
          customColConfigs
        }, true);
      }
      if (result.filteredIndices) {
        setRowIndicesFilter(result.filteredIndices);
        setRowFilterLabel(`Lệnh AI: "${aiCommandText}"`);
      }
    }
  };

  const handleSaveMacroFromCommand = () => {
    if (!aiCommandText.trim()) return;
    const name = prompt("Nhập tên bộ quy tắc Macro mới để lưu lại:", `Macro: ${aiCommandText}`);
    if (!name) return;
    
    const newMacro = {
      id: "macro_" + Date.now(),
      name,
      columns: selectedColumns.length > 0 ? [...selectedColumns] : (aiCommandResult?.summary?.col ? [aiCommandResult.summary.col] : []),
      command: aiCommandText,
      createdAt: new Date().toLocaleDateString("vi-VN")
    };

    setSavedMacros([newMacro, ...savedMacros]);
    alert("Đã lưu bộ quy tắc Macro thành công vào tủ lưu trữ local!");
  };

  const handleExportCommandReport = () => {
    if (!aiCommandResult || !aiCommandResult.success) return;
    
    let exportData: any[] = [];
    if (aiCommandResult.commandType === "frequency" && aiCommandResult.frequencyList) {
      exportData = aiCommandResult.frequencyList.map((item: any) => ({
        "Giá trị": item.value,
        "Số lượng xuất hiện (Tần suất)": item.count,
        "Tỷ lệ phần trăm": item.percent
      }));
    } else if (aiCommandResult.commandType === "calculate" && aiCommandResult.summary) {
      const s = aiCommandResult.summary;
      exportData = [
        { "Chỉ số": "Cột phân tích", "Giá trị": s.col },
        { "Chỉ số": s.title || "Phép tính", "Giá trị": s.value },
        { "Chỉ số": "Trung bình cộng (Avg)", "Giá trị": s.avg },
        { "Chỉ số": "Tổng cộng (Sum)", "Giá trị": s.sum },
        { "Chỉ số": "Nhỏ nhất (Min)", "Giá trị": s.min },
        { "Chỉ số": "Lớn nhất (Max)", "Giá trị": s.max },
        { "Chỉ số": "Số dòng hợp lệ tính toán", "Giá trị": s.count },
        { "Chỉ số": "Số ô khuyết (Trống)", "Giá trị": s.blankCount },
        { "Chỉ số": "Số ô chứa văn bản (Lỗi)", "Giá trị": s.nonNumericCount }
      ];
    } else {
      exportData = [
        { "Thông tin": "Kết quả thực hiện", "Nội dung": aiCommandResult.message },
        { "Thông tin": "Chi tiết kiểm tra", "Nội dung": aiCommandResult.details }
      ];
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    XLSX.utils.book_append_sheet(wb, ws, "BaoCaoPhanTich");
    XLSX.writeFile(wb, `Bao_Cao_Phan_Tich_${Date.now()}.xlsx`);
  };

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
        model: "gemini-2.5-flash",
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
  const [rowIndicesFilter, setRowIndicesFilter] = useState<number[] | null>(null);
  const [rowFilterLabel, setRowFilterLabel] = useState<string | null>(null);
  const [inconsistenciesTab, setInconsistenciesTab] = useState<"desc" | "code">("desc");
  const pageSize = 50;

  // Giới hạn hiển thị cho danh sách bất nhất tránh làm đơ trình duyệt khi nạp hàng ngàn dòng
  const [visibleDescInconCount, setVisibleDescInconCount] = useState<number>(50);
  const [visibleCodeInconCount, setVisibleCodeInconCount] = useState<number>(50);

  const switchDataMode = async (newMode: "corp" | "individual") => {
    if (newMode === dataMode) return;
    
    setLoading(true);
    setStatusMessage(`Đang lưu trữ dữ liệu Chế độ ${dataMode === "corp" ? "Doanh nghiệp" : "Cá thể"}...`);
    
    try {
      // 1. Lưu phiên hiện tại trước khi đổi sang chế độ mới
      await saveAppState({
        mainData,
        rawImportedData,
        columns,
        fileName,
        mapping,
        customColConfigs
      }, true, dataMode);
      
      setStatusMessage(`Đang tải dữ liệu Chế độ ${newMode === "corp" ? "Doanh nghiệp" : "Cá thể"}...`);
      await new Promise(resolve => setTimeout(resolve, 300));
      
      // 2. Cập nhật localStorage và tải dữ liệu chế độ mới
      localStorage.setItem("vsic_current_data_mode", newMode);
      setDataMode(newMode);
      
      const saved = await loadAppState(newMode);
      if (saved) {
        setMainData(saved.mainData || []);
        setRawImportedData(saved.rawImportedData || saved.mainData || []);
        setColumns(saved.columns || []);
        setFileName(saved.fileName || "");
        setMapping(saved.mapping || { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
        setCustomColConfigs(saved.customColConfigs || []);
      } else {
        // Khởi tạo trống nếu chưa có dữ liệu riêng
        setMainData([]);
        setRawImportedData([]);
        setColumns([]);
        setFileName("");
        setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
        setCustomColConfigs([]);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
      setStatusMessage("");
    }
  };

  // Khôi phục dữ liệu từ IndexedDB khi mở ứng dụng
  useEffect(() => {
    async function restoreSession() {
      setLoading(true);
      setStatusMessage("Đang kiểm tra dữ liệu phiên làm việc trước đó trong bộ nhớ...");
      try {
        const initialMode = (localStorage.getItem("vsic_current_data_mode") as "corp" | "individual") || "corp";
        setDataMode(initialMode);
        
        const saved = await loadAppState(initialMode);
        if (saved && saved.mainData && saved.mainData.length > 0) {
          setMainData(saved.mainData);
          setRawImportedData(saved.rawImportedData || saved.mainData);
          setColumns(saved.columns);
          setFileName(saved.fileName);
          if (saved.mapping) setMapping(saved.mapping);
          if (saved.customColConfigs) setCustomColConfigs(saved.customColConfigs);
          setStatusMessage(`Đã khôi phục thành công tệp chế độ ${initialMode === "corp" ? "Doanh nghiệp" : "Cá thể"} "${saved.fileName}" (${saved.mainData.length} dòng) từ phiên trước!`);
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
      }, true, dataMode); // Bắt buộc lưu toàn bộ chuỗi khối dữ liệu thô
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
      }, false, dataMode).catch(err => console.warn("Lỗi tự động lưu phiên làm việc:", err)); // Chỉ lưu metadata siêu nhanh
    }, 1500);

    return () => clearTimeout(timer);
  }, [mainData, rawImportedData, columns, fileName, mapping, customColConfigs, dataMode]);

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



  // Đọc file CSV hoặc Excel bằng xlsx hoặc Bộ phân giải CSV tối ưu
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "main") => {
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
            setRowIndicesFilter(null);
            setRowFilterLabel(null);
            setSearchTerm("");

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

          // Đọc mảng 2D và áp dụng chẩn đoán dòng tiêu đề thông minh tự động
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: "" });
          const parsedResult = parse2DArrayWithSmartHeader(rawRows);
          const data = parsedResult.data;

          if (data.length === 0) {
            alert(`Không tìm thấy dữ liệu hợp lệ trong tệp "${file.name}". Vui lòng kiểm tra lại xem tệp tin có chứa dữ liệu bảng hay không, hoặc thử cấu hình lại bảng dữ liệu.`);
            setLoading(false);
            return;
          }

          const cols = parsedResult.columns;

          if (type === "main") {
            setRawImportedData(data);
            setMainData(data);
            setColumns(cols);
            setFileName(file.name);
            setRowIndicesFilter(null);
            setRowFilterLabel(null);
            setSearchTerm("");

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


  // Reset toàn bộ dữ liệu, chỉ giữ lại DM ngành (VSIC)
  const clearData = () => {
    setCustomConfirmModal({
      isOpen: true,
      title: "Xác nhận xóa tệp nạp vào",
      message: "Hành động này sẽ xóa sạch tất cả dữ liệu tệp chính, tệp nạp thêm, tệp phụ, tệp chọn mẫu và thiết lập cấu hình đã nạp.",
      note: "Toàn bộ Danh mục ngành VSIC chuẩn và các Quy tắc học AI của bạn sẽ được giữ nguyên vẹn, không bị ảnh hưởng.",
      confirmText: "XÁC NHẬN XÓA SẠCH",
      onConfirm: () => {
        setMainData([]);
        setRawImportedData([]);
        setColumns([]);
        setFileName("");
        setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
        setCustomColConfigs([]);
        setRowIndicesFilter(null);
        setRowFilterLabel(null);
        setSearchTerm("");
        setAggregateFiles([]);
        setSecondaryFile(null);
        setSampCorpData([]);
        setSampCorpFileName("");
        setSampIndData([]);
        setSampIndFileName("");
        clearAppState().catch(err => console.warn("Lỗi khi xóa dữ liệu IndexedDB:", err));
        setStatusMessage("Đã xóa toàn bộ các tệp dữ liệu đã nạp thành công! Danh mục ngành VSIC và quy tắc AI vẫn được giữ nguyên vẹn.");
      }
    });
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



  // Hàm chuẩn hóa mô tả hoạt động theo yêu cầu của người dùng
  const cleanAndStandardizeDescription = (rawMota: string): string => {
    if (!rawMota) return "";
    
    // 1. Chuyển toàn bộ về chữ thường
    let text = rawMota.toLowerCase();

    // 2. Loại bỏ sạch các ký tự đặc biệt, dấu phẩy, và các khoảng trắng thừa
    // Giữ lại các chữ cái tiếng Việt có dấu, chữ thường a-z và số 0-9. Tất cả ký tự khác thay bằng khoảng trắng.
    text = text.replace(/[^a-z0-9àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/g, " ");
    text = text.replace(/\s+/g, " ").trim();

    // 3. Gom các biến thể về một từ khóa gốc trước khi làm sạch để bảo toàn ngữ nghĩa nhóm
    // Nhóm 1: Cafe, nước giải khát, trà sữa, sinh tố -> "bán cafe"
    const cafeRegex = /\b(bán cafe|bán nước|quán nước|bán quán nước|quán cafe|quán cà phê|bán cà phê|cửa hàng cafe|cà phê|cafe|nước giải khát|bán trà sữa|quán trà sữa|trà sữa|sinh tố|bán sinh tố|nước mía|bán nước mía)\b/gi;
    text = text.replace(cafeRegex, "bán cafe");

    // Nhóm 2: Tạp hóa -> "bán tạp hóa"
    const taphoaRegex = /\b(bán tạp hóa|tiệm tạp hóa|tạp hóa|tạp hoá|cửa hàng tạp hóa|bán lẻ tạp hóa|tạp hoá tổng hợp|bán tạp hoá)\b/gi;
    text = text.replace(taphoaRegex, "bán tạp hóa");

    // Nhóm 3: Ăn uống -> "dịch vụ ăn uống"
    const anuongRegex = /\b(bán đồ ăn sáng|bán đồ ăn|quán ăn sáng|quán ăn|tiệm ăn|nhà hàng ăn uống|dịch vụ ăn uống|bán cơm|quán cơm|bán phở|quán phở|bán bún|quán bún|ăn uống)\b/gi;
    text = text.replace(anuongRegex, "dịch vụ ăn uống");

    // Nhóm 4: Cắt tóc uốn tóc gội đầu -> "cắt tóc uốn tóc"
    const toctaiRegex = /\b(cắt tóc nam|cắt tóc nữ|uốn tóc|làm tóc|cắt tóc|làm đầu|salon tóc|tiệm tóc|hớt tóc|gội đầu|làm móng|làm nail|tiệm uốn tóc)\b/gi;
    text = text.replace(toctaiRegex, "cắt tóc uốn tóc");

    // Nhóm 5: Quần áo, thời trang -> "bán quần áo"
    const quanaoRegex = /\b(bán quần áo|shop quần áo|cửa hàng quần áo|bán lẻ quần áo|bán quần áo thời trang|thời trang nam nữ|bán váy|shop thời trang)\b/gi;
    text = text.replace(quanaoRegex, "bán quần áo");

    // Nhóm 6: Sửa xe -> "sửa chữa xe máy"
    const suaxeRegex = /\b(sửa chữa xe máy|sửa xe máy|tiệm sửa xe|tiệm sửa xe máy|sửa xe|vá vỏ|vá xe|sửa chữa ô tô|sửa ô tô)\b/gi;
    text = text.replace(suaxeRegex, "sửa chữa xe máy");

    // Nhóm 7: Rau quả trái cây -> "bán rau quả trái cây"
    const rauquaRegex = /\b(bán rau|bán rau củ|bán rau quả|bán hoa quả|bán trái cây|trái cây|rau củ quả)\b/gi;
    text = text.replace(rauquaRegex, "bán rau quả trái cây");

    // Nhóm 8: Xây dựng -> "xây dựng"
    const xaydungRegex = /\b(thầu xây dựng|thợ xây|phụ hồ|xây nhà|làm nề|xây trát|xây dựng nhà)\b/gi;
    text = text.replace(xaydungRegex, "xây dựng");

    // Nhóm 9: Thợ mộc đồ gỗ -> "làm mộc đồ gỗ"
    const mocRegex = /\b(làm mộc|đồ gỗ|xưởng mộc|mộc|sản xuất đồ gỗ|gia công đồ gỗ)\b/gi;
    text = text.replace(mocRegex, "làm mộc đồ gỗ");

    // Nhóm 10: Vé số -> "bán vé số"
    const vesoRegex = /\b(bán vé số|đại lý vé số|xổ số|kiến thiết|vé số dạo)\b/gi;
    text = text.replace(vesoRegex, "bán vé số");

    // 4. Loại bỏ các từ phụ không quan trọng ('tại', 'của', 'các', 'đồ', 'và', 'cho', 'ở'...) đứng độc lập
    // BẮT BUỘC GIỮ LẠI các từ định danh ngành cốt lõi: 'bán buôn', 'bán lẻ', 'sản xuất', 'sửa chữa', 'cho thuê', 'gia công', 'lắp đặt', 'thi công', 'vận tải', 'dịch vụ'...
    const stopWordsRegex = /\b(tại|của|các|đồ|và|cho|ở|nhà|phố|bằng|theo|với|trong|về|tự|hộ kinh doanh|hộ kd|hộ|cơ sở|cửa hàng|tiệm|chuyên|nhận|làm|kinh doanh|kd)\b/gi;
    text = text.replace(stopWordsRegex, "");

    // Chuẩn hóa khoảng trắng một lần nữa sau khi đã xử lý từ phụ
    text = text.replace(/\s+/g, " ").trim();

    return text;
  };

  // Phân tích bất nhất mã ngành và mô tả
  const inconAnalysis = useMemo(() => {
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
      const cleanMota = cleanAndStandardizeDescription(rawMota);
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
          descToCodes.push({
            motaText: cleanMota, // Hiển thị luôn từ khóa gốc đã gom chuẩn hóa cho trực quan
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
      const cleanMota = cleanAndStandardizeDescription(rawMota);
      if (!cleanMota) return;

      if (!codeMap.has(code)) {
        codeMap.set(code, []);
      }
      codeMap.get(code)!.push({ desc: cleanMota, rowIdx: idx, row });
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
        const key = occ.desc;
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
  }, [mainData, mapping.mota, mapping.manganh, stdDescriptionCol, stdIndustryCol]);

  // Ánh xạ trạng thái mâu thuẫn của từng dòng (theo index dòng gốc) để hỗ trợ hiển thị và xuất Excel có đánh dấu
  const rowInconStatusMap = useMemo(() => {
    const map = new Map<number, {
      isMinority: boolean;
      majorityCode: string;
      countOfThisCode: number;
      totalOccurrences: number;
      motaText: string;
    }>();

    if (!inconAnalysis || !inconAnalysis.descToCodes) return map;

    inconAnalysis.descToCodes.forEach(item => {
      const maxCount = Math.max(...item.codes.map((c: any) => c.count));
      const majorityCodeObj = item.codes.find((c: any) => c.count === maxCount);
      const majorityCode = majorityCodeObj ? majorityCodeObj.code : "";

      item.codes.forEach((c: any) => {
        const isMinority = c.count < maxCount;
        c.rows.forEach((rIdx: number) => {
          map.set(rIdx, {
            isMinority,
            majorityCode,
            countOfThisCode: c.count,
            totalOccurrences: item.occurrences,
            motaText: item.motaText
          });
        });
      });
    });

    return map;
  }, [inconAnalysis]);

  // Bộ lọc dữ liệu viewer
  const filteredData = useMemo(() => {
    let data = mainData;
    if (rowIndicesFilter !== null) {
      data = mainData.filter((_, idx) => rowIndicesFilter.includes(idx));
    }
    if (!searchTerm) return data;
    const term = searchTerm.toLowerCase();
    return data.filter(row => {
      return Object.values(row).some(val => String(val).toLowerCase().includes(term));
    });
  }, [mainData, searchTerm, rowIndicesFilter]);

  // Dữ liệu viewer tích hợp các cột đánh dấu phân tích khi có mâu thuẫn
  const augmentedFilteredData = useMemo(() => {
    return filteredData.map((row) => {
      const originalIdx = mainData.indexOf(row);
      if (originalIdx !== -1 && rowInconStatusMap.has(originalIdx)) {
        const info = rowInconStatusMap.get(originalIdx)!;
        return {
          ...row,
          "ĐÁNH DẤU SAI LỆCH (MÃ THIỂU SỐ)": info.isMinority ? "⚠️ THIỂU SỐ (Chọn ít hơn - Nghi ngờ gán sai)" : "✓ ĐA SỐ (Phổ biến nhất)",
          "GỢI Ý MÃ VSIC ĐÚNG (Mã đa số)": info.majorityCode,
          "SỐ DÒNG CÙNG MÃ NÀY": info.countOfThisCode,
          "TỔNG SỐ DÒNG CÙNG MÔ TẢ": info.totalOccurrences,
          "DÒNG SỐ TRONG FILE GỐC": originalIdx + 1
        };
      }
      return row;
    });
  }, [filteredData, mainData, rowInconStatusMap]);

  // Danh sách cột hiển thị tương ứng trên viewer
  const viewerColumns = useMemo(() => {
    if (rowIndicesFilter && rowIndicesFilter.length > 0) {
      return [
        "ĐÁNH DẤU SAI LỆCH (MÃ THIỂU SỐ)",
        "GỢI Ý MÃ VSIC ĐÚNG (Mã đa số)",
        "SỐ DÒNG CÙNG MÃ NÀY",
        "TỔNG SỐ DÒNG CÙNG MÔ TẢ",
        "DÒNG SỐ TRONG FILE GỐC",
        ...columns
      ];
    }
    return columns;
  }, [columns, rowIndicesFilter]);

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

  // Phân trang dữ liệu hiển thị
  const paginatedData = useMemo(() => {
    const startIdx = (viewPage - 1) * pageSize;
    return filteredData.slice(startIdx, startIdx + pageSize);
  }, [filteredData, viewPage]);

  // Tổng số trang
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;

  // Optimized Crosstab and Chi-Square Calculation Memo
  const crosstabResult = useMemo(() => {
    if (!mainData || mainData.length === 0) return null;
    const col1 = tqHangCol || (columns.length > 0 ? columns[0] : "");
    const col2 = tqCotCol || (columns.length > 1 ? columns[1] : (columns[0] || ""));

    if (!col1 || !col2) return null;

    const rawRowsSet = new Set<string>();
    const rawColsSet = new Set<string>();
    const matrix: Record<string, Record<string, number>> = {};
    const rowTotals: Record<string, number> = {};
    const colTotals: Record<string, number> = {};
    let grandTotal = 0;

    mainData.forEach(row => {
      let rVal = String(row[col1] ?? "").trim();
      let cVal = String(row[col2] ?? "").trim();
      if (!rVal) rVal = "(Trống)";
      if (!cVal) cVal = "(Trống)";

      rawRowsSet.add(rVal);
      rawColsSet.add(cVal);

      if (!matrix[rVal]) matrix[rVal] = {};
      matrix[rVal][cVal] = (matrix[rVal][cVal] || 0) + 1;

      rowTotals[rVal] = (rowTotals[rVal] || 0) + 1;
      colTotals[cVal] = (colTotals[cVal] || 0) + 1;
      grandTotal++;
    });

    const smartSort = (arr: string[]) => {
      return [...arr].sort((a, b) => {
        const aLow = a.toLowerCase().trim();
        const bLow = b.toLowerCase().trim();
        if (aLow === "có") return -1;
        if (bLow === "có") return 1;
        if (aLow === "không") {
          if (bLow === "có") return 1;
          return -1;
        }
        if (bLow === "không") {
          if (aLow === "có") return -1;
          return 1;
        }
        return a.localeCompare(b, "vi");
      });
    };

    const sortedRows = smartSort(Array.from(rawRowsSet));
    const sortedCols = smartSort(Array.from(rawColsSet));

    // Calculate Chi-Square test of Independence
    let chiSquare = 0;
    let df = 0;
    let pValue = 1.0;
    let chiSquareValid = false;

    const rCount = sortedRows.length;
    const cCount = sortedCols.length;

    if (rCount > 1 && cCount > 1 && grandTotal > 0) {
      chiSquareValid = true;
      df = (rCount - 1) * (cCount - 1);

      sortedRows.forEach(rVal => {
        const rTot = rowTotals[rVal] || 0;
        sortedCols.forEach(cVal => {
          const cTot = colTotals[cVal] || 0;
          const expected = (rTot * cTot) / grandTotal;
          if (expected > 0) {
            const observed = matrix[rVal]?.[cVal] || 0;
            chiSquare += Math.pow(observed - expected, 2) / expected;
          }
        });
      });

      pValue = chiSquarePValue(chiSquare, df);
    }

    return {
      col1,
      col2,
      matrix,
      rowTotals,
      colTotals,
      grandTotal,
      sortedRows,
      sortedCols,
      chiSquare,
      df,
      pValue,
      chiSquareValid
    };
  }, [mainData, tqHangCol, tqCotCol, columns]);

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
    if (!fileName.startsWith("BaoCaoTongHop_")) {
      setFileName(`BaoCaoTongHop_${fileName}`);
    }
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
          const stdName = lookupResult.name || (rawCode ? `Ngành/Sản phẩm CAPI (${rawCode})` : "");

          if (rawCode) {
            validCount++;
          } else {
            invalidCount++;
          }

          // Đối chiếu quy luật logic hoạt động mô tả & mã ngành để phát hiện mâu thuẫn lệch vai trò
          let auditStatus = "✅ Đạt chuẩn VSIC quốc gia";
          if (isExistInVSIC) {
            auditStatus = lookupResult.exactMatched ? "✅ Đạt chuẩn VSIC quốc gia" : "✅ Khớp quy nạp cấp học";
          } else if (rawCode) {
            auditStatus = "✅ Mã ngành chuẩn CAPI (Đã khớp)";
          } else {
            auditStatus = "⚠️ Thiếu thông tin mã ngành";
          }
          
          let anomalyReason = "";
          if (!rawCode) {
            anomalyReason = "Thiếu thông tin mã ngành / mã hoạt động trong dòng dữ liệu";
          }

          if (anomalyReason && anomalies.length < 5000) {
            anomalies.push({
              dongSTT: idx + 1,
              maDN: row["Mã Số Thuế"] || row["MaST"] || `Bản ghi số ${idx + 1}`,
              maGoc: rawCode,
              motaGoc: rawDesc,
              nganhChuan: stdName || "(Trống)",
              phanTichloi: anomalyReason
            });
          }

          // Xây dựng Bản ghi mới co cụm, bơm cột Tên Ngành Chuẩn VSIC và Trạng Thái Đối Chiếu VSIC nằm ngay bên cạnh cột Mô Tả Hoạt Động / Mã Ngành để dễ đối chiếu
          const flexRow: any = {};
          Object.keys(row).forEach(key => {
            flexRow[key] = row[key];
            if (key === targetDescriptionCol) {
              flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ Thiếu mã ngành";
              flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
            }
          });

          // Nếu không khớp được vị trí cột mô tả thì tự chêm cột mới vào kế cột Mã ngành
          if (flexRow["Tên Ngành Chuẩn VSIC"] === undefined) {
            Object.keys(row).forEach(key => {
              flexRow[key] = row[key];
              if (key === targetIndustryCol) {
                flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ Thiếu mã ngành";
                flexRow["Trạng Thái Đối Chiếu VSIC"] = auditStatus;
              }
            });
          }

          // Nếu vẫn thiếu do trùng cấu hình dặc biệt
          if (flexRow["Tên Ngành Chuẩn VSIC"] === undefined) {
            flexRow["Tên Ngành Chuẩn VSIC"] = stdName || "⚠️ Thiếu mã ngành";
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
    try {
      const activeFile = allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || (mainData && mainData.length > 0 ? { id: "main_data_file", name: "Dữ liệu chính", data: mainData, columns: columns } : null);
      if (!activeFile || !activeFile.data || activeFile.data.length === 0) {
        alert("Vui lòng nạp hoặc chọn dữ liệu trước khi chạy báo cáo.");
        return;
      }
      const targetData = activeFile.data || [];
      const targetColumns = activeFile.columns || [];

      let targetManganh = optManganh || quickReportManganhCol || mapping.manganh;
      let targetXa = optXa || quickReportXaCol || mapping.xa;
      let targetDoanhThu = optDoanhThu || quickReportDoanhThuCol || mapping.doanhthu;
      let targetLaoDong = optLaoDong || quickReportLaoDongCol || mapping.laodong;

      // Tự động dò tìm cột Mã ngành nếu bị trống
      if (!targetManganh) {
        const foundMng = targetColumns.find(c => /mã\s*ngành|manganh|vsic|mã\s*nghe|manghe|ngành|ma_nganh/i.test(c));
        if (foundMng) {
          targetManganh = foundMng;
          setQuickReportManganhCol(foundMng);
        }
      }

      // Tự động dò tìm cột Xã / Địa bàn nếu bị trống
      if (!targetXa) {
        const foundXa = targetColumns.find(c => /xã|xa|địa\s*bàn|dia\s*ban|phường|phuong|ma_xa|ten_xa/i.test(c));
        if (foundXa) {
          targetXa = foundXa;
          setQuickReportXaCol(foundXa);
        }
      }

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

      const sumColsToCheck = Array.isArray(quickReportSumCols) ? quickReportSumCols : [];
      sumColsToCheck.forEach(col => {
        if (col && (targetColumns.includes(col) || col === "Số lượng dòng" || col === "Số cơ sở")) {
          sumCols.push(col);
        }
      });

      // Nếu không cấu hình chỉ tiêu phụ động, tự chuyển về tương thích ngược dựa vào lựa chọn Doanh Thu và Lao Động
      if (sumCols.length === 0) {
        if (targetDoanhThu && targetColumns.includes(targetDoanhThu)) sumCols.push(targetDoanhThu);
        if (targetLaoDong && targetColumns.includes(targetLaoDong)) sumCols.push(targetLaoDong);
      }

      // Tự động tìm kiếm các cột số trong tệp dữ liệu nếu chưa chọn hoặc không tìm thấy cột doanh thu/lao động
      if (sumCols.length === 0) {
        const firstRow = targetData[0] || {};
        const detectedNumericCols = targetColumns.filter(col => {
          if (col === targetManganh || col === targetXa) return false;
          const isIdOrCode = /mã|mst|code|id|phone|đt|điện\s*thoại|tel|fax|stt|index|key|serial|no\./i.test(col);
          if (isIdOrCode) return false;
          const val = String(firstRow[col] || "");
          return val && !isNaN(parseFloat(val.replace(/[^0-9.\-]/g, "")));
        });
        if (detectedNumericCols.length > 0) {
          sumCols.push(...detectedNumericCols.slice(0, 3)); // Lấy tối đa 3 cột số đầu tiên làm mẫu
        }
      }

      // Nếu vẫn trống, lấy đại diện 1 cột số lượng dòng ảo làm chỉ tiêu cộng dồn
      if (sumCols.length === 0) {
        sumCols.push("Số lượng dòng");
      }

      setLoading(true);
      setProgress(0);
      setStatusMessage(`Đang tạo báo cáo tổng hợp gộp nhóm...`);
      await sleep(200);

      const processedData = await chunkProcess(
        targetData,
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
            // Gom nhóm trực tiếp bằng nội dung chuỗi gốc trong cột, không tra cứu bảng VSIC (Phù hợp mọi cuộc điều tra dân số/nông nghiệp/địa bàn bất kỳ)
            tenNganhLabel = String(row[targetManganh] || "Chưa xác định / Bỏ trống").trim();
          } else {
            const mng = normalizeSectorCode(row[targetManganh]);
            if (level === 1) {
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
            } else if (level === 2) {
              const sec2Code = mng ? mng.slice(0, 2) : "";
              const sec2Name = vsicRawData[sec2Code] || "Ngành cấp 2 chưa định nghĩa";
              tenNganhLabel = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định - Ngành cấp 2 chưa định nghĩa";
            } else if (level === 3) {
              const sec3Code = mng ? mng.slice(0, 3) : "";
              const sec3Name = vsicRawData[sec3Code] || "Ngành cấp 3 chưa định nghĩa";
              tenNganhLabel = sec3Code ? `${sec3Code} - ${sec3Name}` : "Chưa xác định - Ngành cấp 3 chưa định nghĩa";
            } else if (level === 4) {
              const sec4Code = mng ? mng.slice(0, 4) : "";
              const sec4Name = vsicRawData[sec4Code] || "Ngành cấp 4 chưa định nghĩa";
              tenNganhLabel = sec4Code ? `${sec4Code} - ${sec4Name}` : "Chưa xác định - Ngành cấp 4 chưa định nghĩa";
            } else if (level === 5) {
              const sec5Code = mng ? mng.slice(0, 5) : "";
              const sec5Name = vsicRawData[sec5Code] || "Ngành cấp 5 chưa định nghĩa";
              tenNganhLabel = sec5Code ? `${sec5Code} - ${sec5Name}` : "Chưa xác định - Ngành cấp 5 chưa định nghĩa";
            } else if (level === 6) {
              const sec2Code = mng ? mng.replace(/\D/g, "").slice(0, 2) : "";
              const sec2Num = parseInt(sec2Code, 10);
              if (!isNaN(sec2Num)) {
                if (sec2Num >= 5 && sec2Num <= 39) {
                  tenNganhLabel = "Công nghiệp (05-39)";
                } else if (sec2Num >= 41 && sec2Num <= 43) {
                  tenNganhLabel = "Xây dựng (41-43)";
                } else if (sec2Num >= 45 && sec2Num <= 47) {
                  tenNganhLabel = "Thương mại (45-47)";
                } else if (sec2Num >= 49 && sec2Num <= 53) {
                  tenNganhLabel = "Vận tải (49-53)";
                } else if (sec2Num >= 55 && sec2Num <= 99) {
                  tenNganhLabel = "Dịch vụ (55-99)";
                } else {
                  tenNganhLabel = `Khác - Ngoài danh mục (${sec2Code})`;
                }
              } else {
                tenNganhLabel = "Chưa xác định / Bỏ trống";
              }
            }
          }

          return {
            ...row,
            _temNganhCap: tenNganhLabel,
            _tempXa: String(row[targetXa] || "Khác").trim(),
            "Số lượng dòng": 1,
            "Số cơ sở": 1
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
            [targetXa]: commune
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

          communeObj["Số lượng dòng"] = totalCommuneDN;
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
            rowObj[targetManganh] = dims.Ngành;
          } else if (level === 6) {
            rowObj["Nhóm_Ngành_Chính"] = dims.Ngành;
          } else {
            rowObj[`Ngành_Cấp_${level}`] = dims.Ngành;
          }
          rowObj[targetXa] = dims.Xã;
          rowObj["Số lượng dòng"] = rowsObj.length;

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

      const newDataset: CompiledDataset = {
        id: "main_" + Date.now(),
        name: `Biểu gộp: ${fileName || "Dữ liệu chính"} (${level === 6 ? "Nhóm ngành chính" : "Cấp " + (level === 0 ? "gốc" : "ngành " + level)}, ${reportType === "pivot" ? "Pivot" : "Mẫu dọc"})`,
        rows: finalReportRows,
        cols: Object.keys(finalReportRows[0] || {}),
        level: level,
        reportType: reportType
      };
      setCompiledDatasets(prev => [newDataset, ...prev]);

      // Bổ sung: Lưu biểu gộp vào danh sách tệp tin bộ nhớ tạm để có thể chọn làm tệp nguồn trong phép tính & so sánh
      setAggregateFiles(prev => [
        ...prev,
        {
          id: newDataset.id,
          name: newDataset.name,
          data: finalReportRows,
          columns: newDataset.cols
        }
      ]);

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
- "redefine": Định nghĩa lại tên cột, Việt hóa, xóa bỏ các cột không cần thiết.
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

2. Nếu hành vi là "calculate" (Người dùng yêu cầu tính toán như cộng, trừ, nhân, chia, ghép nối, phần trăm, VAT, trung bình, năng suấ):
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
        model: "gemini-2.5-flash",
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
        model: "gemini-2.5-flash",
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

  const handleExportQuickReport = (customRows?: any[], customCols?: string[]) => {
    const rowsToExport = customRows || quickReportResultRows;
    if (rowsToExport.length === 0) {
      alert("Không có số liệu báo cáo để xuất!");
      return;
    }
    try {
      const ws = XLSX.utils.json_to_sheet(rowsToExport);
      const wb = XLSX.utils.book_new();
      const sheetName = quickReportLevel === 6 ? "Nhóm ngành chính" : `Báo cáo cấp ${quickReportLevel}`;
      const fileNameSuffix = quickReportLevel === 6 ? "Nhom_Nganh_Chinh" : `Nganh_Cap_${quickReportLevel}`;
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, `Bao_Cao_Tong_Hop_${fileNameSuffix}_Va_Xa_Da_Chinh_Sua.xlsx`);
    } catch (err: any) {
      alert("Lỗi xuất Excel: " + err.message);
    }
  };

  // Nạp thêm tệp tin dữ liệu phụ/khác cho phần Tổng Hợp
  const handleAggregateFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setLoading(true);
    setStatusMessage("Đang nạp và phân tích tệp dữ liệu...");
    
    let filesProcessed = 0;
    const totalFiles = files.length;

    Array.from(files).forEach((file: any) => {
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          if (!arrayBuffer) throw new Error("Không thể đọc nội dung tệp!");

          const wb = XLSX.read(arrayBuffer, { 
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false
          });

          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          
          // Đọc mảng 2D và áp dụng chẩn đoán dòng tiêu đề thông minh tự động
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: "" });
          const parsedResult = parse2DArrayWithSmartHeader(rawRows);
          const data = parsedResult.data;

          if (data.length === 0) {
            alert(`Tệp "${file.name}" trống hoặc không chứa dữ liệu hợp lệ!`);
            return;
          }

          const cols = parsedResult.columns;
          
          setAggregateFiles(prev => {
            const exists = prev.some(f => f.name === file.name);
            if (exists) {
              return prev.map(f => f.name === file.name ? { ...f, data, columns: cols } : f);
            }
            return [...prev, {
              id: "agg_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
              name: file.name,
              data,
              columns: cols
            }];
          });
          
        } catch (err: any) {
          alert(`Lỗi đọc tệp "${file.name}": ` + err.message);
        } finally {
          filesProcessed++;
          if (filesProcessed === totalFiles) {
            setLoading(false);
            alert(`Đã nạp thành công các tệp tin được chọn!`);
          }
        }
      };
    });
  };

  // Thực hiện cộng trừ nhân chia cột giữa các tệp tin khác nhau (hỗ trợ tối đa 3 cặp cột cùng lúc)
  const handlePerformCrossFileMath = () => {
    const fileA = allAvailableFiles.find(f => f.id === mathFileAId);
    const fileB = allAvailableFiles.find(f => f.id === mathFileBId);
    
    if (!fileA || !fileB) {
      alert("Vui lòng chọn đầy đủ 2 tệp tin nguồn để thực hiện tính toán!");
      return;
    }
    if (!mathColA || !mathColB) {
      alert("Vui lòng chọn cột số cần tính toán từ cả 2 tệp ở Cặp 1!");
      return;
    }
    if (!mathKeyA || !mathKeyB) {
      alert("Vui lòng chọn cột khóa liên kết để so khớp dòng giữa 2 tệp!");
      return;
    }
    if (!mathNewColName.trim()) {
      alert("Vui lòng đặt tên cho cột kết quả mới ở Cặp 1!");
      return;
    }

    setLoading(true);
    setStatusMessage("Đang thực hiện tính toán & so sánh các tệp gộp...");

    try {
      // Index file B by its key
      const mapB = new Map<string, any>();
      fileB.data.forEach(row => {
        let keyVal = String(row[mathKeyB] ?? "").trim().toLowerCase();
        if (mathKeyB2) {
          keyVal += "||" + String(row[mathKeyB2] ?? "").trim().toLowerCase();
        }
        if (keyVal !== "") {
          mapB.set(keyVal, row);
        }
      });

      const resultRows: any[] = [];
      const suffixB = `_tu_${fileB.name.substring(0, 10).replace(/[^a-zA-Z0-9]/g, "")}`;
      
      fileA.data.forEach(rowA => {
        let keyAVal = String(rowA[mathKeyA] ?? "").trim().toLowerCase();
        if (mathKeyA2) {
          keyAVal += "||" + String(rowA[mathKeyA2] ?? "").trim().toLowerCase();
        }
        const rowB = mapB.get(keyAVal);

        const newRow = { ...rowA };

        // CẶP 1: Bắt buộc
        if (mathColA && mathColB) {
          const valA = parseRobustNumber(rowA[mathColA]);
          const valB = rowB ? parseRobustNumber(rowB[mathColB]) : 0;

          let resultVal = 0;
          if (mathOp === "+") resultVal = valA + valB;
          else if (mathOp === "-") resultVal = valA - valB;
          else if (mathOp === "*") resultVal = valA * valB;
          else if (mathOp === "/") resultVal = valB !== 0 ? (valA / valB) : 0;

          resultVal = Math.round(resultVal * 100) / 100;
          newRow[`${mathNewColName}`] = resultVal;

          if (rowB && mathColB !== mathColA) {
            newRow[`${mathColB}${suffixB}`] = rowB[mathColB];
          }
        }

        // CẶP 2: Tùy chọn (Nếu được thiết lập)
        if (mathColA2 && mathColB2 && mathNewColName2.trim()) {
          const valA = parseRobustNumber(rowA[mathColA2]);
          const valB = rowB ? parseRobustNumber(rowB[mathColB2]) : 0;

          let resultVal = 0;
          if (mathOp2 === "+") resultVal = valA + valB;
          else if (mathOp2 === "-") resultVal = valA - valB;
          else if (mathOp2 === "*") resultVal = valA * valB;
          else if (mathOp2 === "/") resultVal = valB !== 0 ? (valA / valB) : 0;

          resultVal = Math.round(resultVal * 100) / 100;
          newRow[`${mathNewColName2}`] = resultVal;

          if (rowB && mathColB2 !== mathColA2) {
            newRow[`${mathColB2}${suffixB}`] = rowB[mathColB2];
          }
        }

        // CẶP 3: Tùy chọn (Nếu được thiết lập)
        if (mathColA3 && mathColB3 && mathNewColName3.trim()) {
          const valA = parseRobustNumber(rowA[mathColA3]);
          const valB = rowB ? parseRobustNumber(rowB[mathColB3]) : 0;

          let resultVal = 0;
          if (mathOp3 === "+") resultVal = valA + valB;
          else if (mathOp3 === "-") resultVal = valA - valB;
          else if (mathOp3 === "*") resultVal = valA * valB;
          else if (mathOp3 === "/") resultVal = valB !== 0 ? (valA / valB) : 0;

          resultVal = Math.round(resultVal * 100) / 100;
          newRow[`${mathNewColName3}`] = resultVal;

          if (rowB && mathColB3 !== mathColA3) {
            newRow[`${mathColB3}${suffixB}`] = rowB[mathColB3];
          }
        }

        resultRows.push(newRow);
      });

      const cleanNameA = fileA.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "").substring(0, 15);
      const cleanNameB = fileB.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "").substring(0, 15);
      const calculatedFileName = `SoSanh_TinhToan_${cleanNameA}_vs_${cleanNameB}`;
      const newFileId = "calc_" + Date.now();
      const newCols = Object.keys(resultRows[0] || {});

      // Add to aggregateFiles so it can be selected as a source file itself!
      setAggregateFiles(prev => [
        ...prev,
        {
          id: newFileId,
          name: calculatedFileName,
          data: resultRows,
          columns: newCols
        }
      ]);

      // Set it as the current active dataset / report preview
      setQuickReportResultRows(resultRows);
      setQuickReportResultCols(newCols);
      alert(`Phép tính & so sánh hoàn tất! Đã tạo tệp mới: "${calculatedFileName}" gồm ${resultRows.length} dòng và tự động nạp vào bảng xem.`);
    } catch (err: any) {
      alert("Lỗi khi thực hiện tính toán liên file: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleSaveTongHopCommand = () => {
    const name = prompt("Nhập tên cho lệnh tổng hợp này để lưu lại:");
    if (!name || !name.trim()) return;

    const newCommand: SavedTongHopCommand = {
      id: "cmd_" + Date.now(),
      name: name.trim(),
      selectedFileIdToAggregate,
      quickReportManganhCol,
      quickReportXaCol,
      quickReportSumCols,
      reportType,
      mathFileAId,
      mathFileBId,
      mathColA,
      mathColB,
      mathOp,
      mathNewColName,
      mathKeyA,
      mathKeyB
    };

    const updated = [newCommand, ...savedTongHopCommands];
    setSavedTongHopCommands(updated);
    localStorage.setItem("savedTongHopCommands", JSON.stringify(updated));
    alert(`Lưu lệnh "${name.trim()}" thành công!`);
  };

  const handleApplyTongHopCommand = (cmd: SavedTongHopCommand) => {
    if (cmd.selectedFileIdToAggregate !== undefined) setSelectedFileIdToAggregate(cmd.selectedFileIdToAggregate);
    if (cmd.quickReportManganhCol !== undefined) setQuickReportManganhCol(cmd.quickReportManganhCol);
    if (cmd.quickReportXaCol !== undefined) setQuickReportXaCol(cmd.quickReportXaCol);
    if (cmd.quickReportSumCols !== undefined) setQuickReportSumCols(cmd.quickReportSumCols);
    if (cmd.reportType !== undefined) setReportType(cmd.reportType);
    if (cmd.mathFileAId !== undefined) setMathFileAId(cmd.mathFileAId);
    if (cmd.mathFileBId !== undefined) setMathFileBId(cmd.mathFileBId);
    if (cmd.mathColA !== undefined) setMathColA(cmd.mathColA);
    if (cmd.mathColB !== undefined) setMathColB(cmd.mathColB);
    if (cmd.mathOp !== undefined) setMathOp(cmd.mathOp);
    if (cmd.mathNewColName !== undefined) setMathNewColName(cmd.mathNewColName);
    if (cmd.mathKeyA !== undefined) setMathKeyA(cmd.mathKeyA);
    if (cmd.mathKeyB !== undefined) setMathKeyB(cmd.mathKeyB);
    alert(`Đã áp dụng các thiết lập từ lệnh: "${cmd.name}"`);
  };

  const handleDeleteTongHopCommand = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm("Bạn có chắc chắn muốn xóa lệnh này?")) return;
    const updated = savedTongHopCommands.filter(c => c.id !== id);
    setSavedTongHopCommands(updated);
    localStorage.setItem("savedTongHopCommands", JSON.stringify(updated));
  };

  const handleExportTongHopCommands = () => {
    if (savedTongHopCommands.length === 0) {
      alert("Chưa có lệnh nào được lưu để xuất!");
      return;
    }
    try {
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(savedTongHopCommands, null, 2));
      const downloadAnchor = document.createElement("a");
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", "Lenh_Tong_Hop_Cau_Hinh.json");
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
    } catch (err: any) {
      alert("Lỗi xuất lệnh: " + err.message);
    }
  };

  const handleImportTongHopCommands = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const content = evt.target?.result as string;
        const parsed = JSON.parse(content);
        if (!Array.isArray(parsed)) {
          throw new Error("Định dạng tệp không hợp lệ, phải là một danh sách lệnh.");
        }
        
        const updated = [...parsed, ...savedTongHopCommands];
        const unique = updated.filter((v, i, a) => a.findIndex(t => t.id === v.id) === i);
        
        setSavedTongHopCommands(unique);
        localStorage.setItem("savedTongHopCommands", JSON.stringify(unique));
        alert(`Đã nhập thành công ${parsed.length} lệnh tổng hợp cấu hình!`);
      } catch (err: any) {
        alert("Lỗi khi nhập tệp lệnh: " + err.message);
      }
    };
    reader.readAsText(file);
  };

  // Hỗ trợ nạp file thứ hai / các năm cũ để tổng hợp song song
  const handleSecondaryFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage("Đang đọc tệp tin thứ hai / dữ liệu cũ...");
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        if (!arrayBuffer) throw new Error("Không thể đọc nội dung tệp tin!");

        const wb = XLSX.read(arrayBuffer, { 
          type: "array",
          dense: true,
          cellFormula: false,
          cellHTML: false,
          cellStyles: false
        });

        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        
        // Đọc mảng 2D và áp dụng chẩn đoán dòng tiêu đề thông minh tự động
        const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: "" });
        const parsedResult = parse2DArrayWithSmartHeader(rawRows);
        const data = parsedResult.data;

        if (data.length === 0) {
          alert("Tệp trống hoặc không chứa dữ liệu hợp lệ!");
          setLoading(false);
          return;
        }

        const cols = parsedResult.columns;
        setSecondaryFile({
          name: file.name,
          data: data,
          columns: cols
        });

        // Tự động phát hiện cột
        const autoMng = cols.find(c => c.toLowerCase().includes("mã ngành") || c.toLowerCase().includes("manganh") || c.toLowerCase().includes("ngành") || c.toLowerCase().includes("ma_nganh")) || "";
        const autoXa = cols.find(c => c.toLowerCase().includes("xã") || c.toLowerCase().includes("phường") || c.toLowerCase().includes("địa bàn") || c.toLowerCase().includes("dia_ban")) || "";
        setSecManganhCol(autoMng);
        setSecXaCol(autoXa);
        setSecSumCols([]);
        
        setLoading(false);
        alert(`Đọc tệp "${file.name}" thành công! Vui lòng cấu hình các cột phân nhóm bên dưới để tiến hành tổng hợp.`);
      } catch (err: any) {
        alert("Lỗi đọc tệp tin thứ hai: " + err.message);
        setLoading(false);
      }
    };
  };

  const handleQuickReportSecondary = async (level: number) => {
    if (!secondaryFile) {
      alert("Vui lòng tải tệp thứ hai lên trước!");
      return;
    }
    const targetManganh = secManganhCol;
    const targetXa = secXaCol;

    if (!targetManganh) {
      alert("Vui lòng chỉ định cột chứa Mã ngành / Phân nhóm ở tệp thứ hai!");
      return;
    }
    if (!targetXa) {
      alert("Vui lòng chỉ định cột chứa Xã / Địa bàn ở tệp thứ hai!");
      return;
    }

    const sumCols = secSumCols.filter(col => secondaryFile.columns.includes(col) || col === "Số lượng dòng" || col === "Số cơ sở");
    if (sumCols.length === 0) {
      alert("Vui lòng tích chọn ít nhất 1 chỉ tiêu cột số để cộng tổng!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang chạy tổng hợp tệp thứ hai: ${secondaryFile.name}...`);
    await sleep(200);

    try {
      const processedData = await chunkProcess(
        secondaryFile.data,
        10000,
        row => {
          if (!row || typeof row !== 'object') {
            return { _temNganhCap: "Chưa xác định", _tempXa: "Khác" };
          }
          let tenNganhLabel = "";
          if (level === 0) {
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
            _tempXa: String(row[targetXa] || "Khác").trim(),
            "Số lượng dòng": 1,
            "Số cơ sở": 1
          };
        },
        pct => {
          setProgress(Math.round(pct * 0.5));
          setStatusMessage(`Đang xử lý dữ liệu: ${pct}%...`);
        }
      );

      let finalReportRows: any[] = [];

      if (secReportType === "pivot") {
        const communes = Array.from(new Set(processedData.map(r => r._tempXa))).sort();
        const sectorLabels = Array.from(new Set(processedData.map(r => r._temNganhCap))).sort();

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
          const communeObj: any = { [targetXa]: commune };
          let totalCommuneDN = 0;
          const totalAccumulate: { [col: string]: number } = {};
          sumCols.forEach(col => { totalAccumulate[col] = 0; });

          sectorLabels.forEach(sector => {
            const matchedRows = groupedMap.get(`${commune}||${sector}`) || [];
            const columnSums: { [col: string]: number } = {};
            sumCols.forEach(col => { columnSums[col] = 0; });

            matchedRows.forEach(r => {
              sumCols.forEach(col => {
                columnSums[col] += parseRobustNumber(r[col]);
              });
            });

            sumCols.forEach(col => {
              communeObj[`${sector} - Tổng ${col}`] = Math.round(columnSums[col] * 100) / 100;
              totalAccumulate[col] += columnSums[col];
            });
            totalCommuneDN += matchedRows.length;
          });

          communeObj["Số lượng dòng"] = totalCommuneDN;
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

        groups.forEach((rowsObj, keyStr) => {
          const dims = JSON.parse(keyStr);
          const rowObj: any = {};
          if (level === 0) {
            rowObj[targetManganh] = dims.Ngành;
          } else {
            rowObj[`Ngành_Cấp_${level}`] = dims.Ngành;
          }
          rowObj[targetXa] = dims.Xã;
          rowObj["Số lượng dòng"] = rowsObj.length;

          sumCols.forEach(col => {
            let sumCol = 0;
            rowsObj.forEach(r => { sumCol += parseRobustNumber(r[col]); });
            rowObj[`Tổng_${col}`] = Math.round(sumCol * 100) / 100;
          });

          finalReportRows.push(rowObj);
        });
      }

      const newDataset: CompiledDataset = {
        id: "sec_" + Date.now(),
        name: `Biểu gộp: ${secondaryFile.name} (Cấp ${level === 0 ? "gốc" : "ngành " + level}, ${secReportType === "pivot" ? "Pivot" : "Mẫu dọc"})`,
        rows: finalReportRows,
        cols: Object.keys(finalReportRows[0] || {}),
        level: level,
        reportType: secReportType
      };

      setCompiledDatasets(prev => [newDataset, ...prev]);
      
      // Bổ sung: Lưu biểu gộp vào danh sách tệp tin bộ nhớ tạm để có thể chọn làm tệp nguồn trong phép tính & so sánh
      setAggregateFiles(prev => [
        ...prev,
        {
          id: newDataset.id,
          name: newDataset.name,
          data: finalReportRows,
          columns: newDataset.cols
        }
      ]);

      setQuickReportResultRows(finalReportRows);
      setQuickReportResultCols(newDataset.cols);
      setQuickReportLevel(level);

      setProgress(100);
      setStatusMessage(`Đã tổng hợp thành công tệp phụ!`);
      setSecondaryFile(null); // Clear form after compilation
      await sleep(350);
      setLoading(false);
      alert(`Tổng hợp thành công! Biểu "${newDataset.name}" đã được nạp gọn gàng vào hệ thống.`);
    } catch (err: any) {
      alert("Lỗi tổng hợp tệp thứ hai: " + err.message);
      setLoading(false);
    }
  };

  const getRowKey = (row: any, cols: string[]) => {
    if (!row) return "";
    const parts: string[] = [];
    
    // Tìm các cột không phải số, không chứa tổng cộng và không phải cột chỉ số dòng để làm khoá ghép nối dòng
    const nonNumericCols = cols.filter(col => {
      if (col.includes(" - Tổng ") || col.startsWith("Tổng_") || col.startsWith("Tổng_Cộng_") || col.includes("_Toàn_Xã")) return false;
      if (/số.*dòng|số.*mẫu|số.*cơ.*sở|bản.*ghi|record|count|số_dòng|số_dn|số_lượng_bản_ghi|Số_Dòng_Tổng_Hợp|Số_Lượng_Bản_Ghi/i.test(col)) return false;
      const val = row[col];
      return typeof val === "string" && isNaN(Number(val));
    });

    if (nonNumericCols.length > 0) {
      nonNumericCols.forEach(col => {
        parts.push(String(row[col] || ""));
      });
    } else {
      // Fallback
      const standardKeys = ["Địa_Bàn_Xã", "Địa_bàn_Xã", "Xã", "Ngành_Cấp_1", "Ngành_Cấp_2", "Nhóm_Phân_Loại", "Nhóm_Ngành_Chính"];
      standardKeys.forEach(key => {
        if (key in row) parts.push(String(row[key]));
      });
    }
    
    return parts.join("||");
  };

  // Thực thi cộng, trừ, nhân, chia cột giữa hai biểu tổng hợp
  const handlePerformMath = () => {
    const dsA = compiledDatasets.find(d => d.id === mathDsA);
    const dsB = compiledDatasets.find(d => d.id === mathDsB);
    if (!dsA || !dsB) {
      alert("Vui lòng chọn đủ 2 biểu tổng hợp để thực hiện phép tính!");
      return;
    }
    if (!mathColA || !mathColB) {
      alert("Vui lòng chọn cột cần tính toán từ cả 2 biểu!");
      return;
    }
    if (!mathNewColName.trim()) {
      alert("Vui lòng đặt tên cho cột kết quả mới!");
      return;
    }

    // Ánh xạ biểu B theo khoá dòng
    const mapB = new Map<string, any>();
    dsB.rows.forEach(r => {
      const key = getRowKey(r, dsB.cols);
      mapB.set(key, r);
    });

    const mathRows: any[] = [];
    const matchedKeysInB = new Set<string>();

    dsA.rows.forEach(rowA => {
      const key = getRowKey(rowA, dsA.cols);
      const rowB = mapB.get(key);
      
      const valA = parseRobustNumber(rowA[mathColA]);
      let valB = 0;
      if (rowB) {
        valB = parseRobustNumber(rowB[mathColB]);
        matchedKeysInB.add(key);
      } else {
        if (!mathTreatMissingAsZero) {
          // Keep as zero
        }
      }

      let resultVal = 0;
      if (mathOp === "+") resultVal = valA + valB;
      else if (mathOp === "-") resultVal = valA - valB;
      else if (mathOp === "*") resultVal = valA * valB;
      else if (mathOp === "/") resultVal = valB !== 0 ? (valA / valB) : 0;

      resultVal = Math.round(resultVal * 100) / 100;

      const newRow = {
        ...rowA,
        [`${mathNewColName}`]: resultVal
      };
      
      mathRows.push(newRow);
    });

    // Outer join - nạp nốt các dòng chỉ có ở biểu B
    dsB.rows.forEach(rowB => {
      const key = getRowKey(rowB, dsB.cols);
      if (!matchedKeysInB.has(key)) {
        const valA = 0;
        const valB = parseRobustNumber(rowB[mathColB]);

        let resultVal = 0;
        if (mathOp === "+") resultVal = valA + valB;
        else if (mathOp === "-") resultVal = valA - valB;
        else if (mathOp === "*") resultVal = valA * valB;
        else if (mathOp === "/") resultVal = valB !== 0 ? (valA / valB) : 0;

        resultVal = Math.round(resultVal * 100) / 100;

        const newRow: any = { ...rowB };
        newRow[`${mathNewColName}`] = resultVal;
        mathRows.push(newRow);
      }
    });

    const resultDataset: CompiledDataset = {
      id: "math_" + Date.now(),
      name: `Phép tính: [${dsA.name.slice(0, 18)}...] ${mathOp} [${dsB.name.slice(0, 18)}...]`,
      rows: mathRows,
      cols: Object.keys(mathRows[0] || {}),
      level: dsA.level,
      reportType: dsA.reportType
    };

    setCompiledDatasets(prev => [resultDataset, ...prev]);
    setQuickReportResultRows(mathRows);
    setQuickReportResultCols(resultDataset.cols);
    alert(`Thực hiện thành công! Đã tạo biểu kết quả mới "${resultDataset.name}" chứa cột tính toán "${mathNewColName}".`);
  };

  // Đối chiếu đa niên độ / So sánh các năm (YoY)
  const handleGenerateYearComparison = () => {
    if (compareDsIds.length < 2) {
      alert("Vui lòng chọn ít nhất 2 biểu tổng hợp để đối chiếu so sánh niên độ!");
      return;
    }
    if (!compareKeyCol) {
      alert("Vui lòng chọn cột định danh phân nhóm (ví dụ: Địa_Bàn_Xã hoặc Nhóm_Phân_Loại)!");
      return;
    }

    const missingMaps = compareDsIds.filter(id => !compareColMapping[id]);
    if (missingMaps.length > 0) {
      alert("Vui lòng chọn đầy đủ cột chỉ tiêu tương ứng cho tất cả các biểu cần so sánh!");
      return;
    }

    const datasets = compareDsIds.map(id => compiledDatasets.find(d => d.id === id)).filter(Boolean) as CompiledDataset[];
    if (datasets.length < 2) return;

    // Lấy tập hợp tất cả các khoá phân nhóm duy nhất
    const allKeys = new Set<string>();
    datasets.forEach(ds => {
      ds.rows.forEach(r => {
        const val = r[compareKeyCol];
        if (val !== undefined && val !== null) {
          allKeys.add(String(val));
        }
      });
    });

    const keyList = Array.from(allKeys).sort();
    const comparisonRows: any[] = [];

    keyList.forEach(keyVal => {
      const rowObj: any = {};
      rowObj[compareKeyCol] = keyVal;

      for (let i = 0; i < datasets.length; i++) {
        const ds = datasets[i];
        const chosenCol = compareColMapping[ds.id];
        const matchedRow = ds.rows.find(r => String(r[compareKeyCol]) === keyVal);
        const val = matchedRow ? parseRobustNumber(matchedRow[chosenCol]) : 0;
        
        rowObj[ds.name] = val;
      }

      // Nếu có đúng 2 năm/tệp so sánh, tính chênh lệch tăng trưởng trực quan
      if (datasets.length === 2) {
        const valA = rowObj[datasets[0].name]; // Năm mới / tệp thứ nhất
        const valB = rowObj[datasets[1].name]; // Năm cũ / tệp thứ hai
        const diff = Math.round((valA - valB) * 100) / 100;
        let pct = 0;
        if (valB !== 0) {
          pct = Math.round(((valA - valB) / valB) * 10000) / 100;
        } else if (valA > 0) {
          pct = 100;
        }
        rowObj["Chênh_Lệch_Tuyệt_Đối"] = diff;
        rowObj["Tăng_Trưởng_Phần_Trăm_YoY"] = pct;
      }

      comparisonRows.push(rowObj);
    });

    const resultDataset: CompiledDataset = {
      id: "compare_" + Date.now(),
      name: `Đối chiếu năm: [${datasets.map(d => d.name.slice(0, 15) + "..").join(" vs ")}]`,
      rows: comparisonRows,
      cols: Object.keys(comparisonRows[0] || {}),
      level: datasets[0].level,
      reportType: datasets[0].reportType
    };

    setCompiledDatasets(prev => [resultDataset, ...prev]);
    setQuickReportResultRows(comparisonRows);
    setQuickReportResultCols(resultDataset.cols);
    if (keyList.length > 0) {
      setSelectedCompareRowKey(keyList[0]);
    }
    alert(`Tạo bảng so sánh các năm thành công! Biểu "${resultDataset.name}" đã được hiển thị bảng số liệu bên dưới.`);
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
      const origNames = customColConfigs.map(c => c.originalName);
      const uniqueRoles = getUniqueRoleAssignments(origNames);

      const mappings = customColConfigs.map(cfg => {
        const c = cfg.originalName;
        const cl = c.toLowerCase();
        let newName = beautifyColumnName(c);
        let role = uniqueRoles[c] || "";
        let use = cfg.use;

        // Nếu có vai trò cụ thể được phân tích tự động, chuẩn hóa tên tiếng Việt tương ứng
        if (role === "idCol") {
          newName = "Mã Số Thuế";
          if (pl.includes("mst thành") || pl.includes("tax thành")) {
            const match = promptText.match(/(?:mst|tax|mã số thuế)\s+(?:thành|là)\s+["'‘“]?([^"'‘”,\s]+)/i);
            if (match && match[1]) newName = match[1];
          }
        } else if (role === "doanhthu") {
          newName = "Doanh Thu";
          if (pl.includes("doanh thu thành") || pl.includes("doanh số thành")) {
            const match = promptText.match(/(?:doanh thu|doanh số)\s+(?:thành|là)\s+["'‘“]?([^"'‘”,\s]+)/i);
            if (match && match[1]) newName = match[1];
          }
        } else if (role === "laodong") {
          newName = "Số Lao Động";
        } else if (role === "xa") {
          newName = "Địa bàn (Xã)";
        } else if (role === "manganh") {
          newName = "Mã Ngành ĐK";
        } else if (role === "mota") {
          newName = "Mô Tả Hoạt Động";
        }

        // Nếu người dùng yêu cầu loại bỏ bớt cột ngoài các vai trò chính hoặc loại bỏ cột thừa
        if (pl.includes("loại bỏ") || pl.includes("bỏ bớt") || pl.includes("chỉ giữ") || pl.includes("bỏ cột thừa") || pl.includes("bỏ chọn") || pl.includes("loại bỏ cột thừa")) {
          if (!role) {
            use = false;
            newName = "";
          }
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

ĐẶC BIỆT LƯU Ý VỀ CỘT THỪA (COLUMN REDUCTION):
Nếu người dùng yêu cầu "loại bỏ cột thừa", "loại bỏ các cột thừa", "bỏ bớt cột thừa", "bỏ chọn các cột", "bỏ cột", "chỉ giữ các cột chính", "loại bỏ cột phụ", hoặc diễn đạt tương đương, bạn BẮT BUỘC phải đặt "use": false cho tất cả những cột không được gán bất kỳ vai trò chính nào (vai trò để rỗng "").

Hãy trả về một mảng JSON trực tiếp đại diện cho các trường được ánh xạ, tuyệt đối không viết thêm lời bình luận, không bọc thẻ markdown ngoài cục diện JSON. Định dạng bắt buộc:
[
  { "originalMatch": "tên_cột_gốc_chính_xác", "newName": "tên_mới_việt_hóa", "role": "vai_trò", "use": true/false }
]`;

        const response = await ai.models.generateContent({
          model: "gemini-2.5-flash",
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
        trangThai = "❌ Lỗi: Lệch mâu thuẫn - Nội dung mô tả thực tế không ăn khớp với mã ngành điều tra viên chọn trên máy";
      } else if (hasTradeKeywords && !hasIndustrialKeywords && isIndustrialCode) {
        trangThai = "❌ Lỗi ĐTV: Mô tả ngành thương mại/bán lẻ nhưng chọn mã ngành Công nghiệp chế biến chế tạo";
      } else if (hasIndustrialKeywords && !isIndustrialCode) {
        trangThai = "❌ Lỗi ĐTV: Mô tả ngành Sản xuất/Gia công nhưng chọn mã ngành dịch vụ/thương mại";
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
    if (!fileName.startsWith("ChuanHoaNganh_VSIC_")) {
      setFileName(`ChuanHoaNganh_VSIC_${fileName}`);
    }

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
    if (!fileName.startsWith("KiemTraLogic_")) {
      setFileName(`KiemTraLogic_${fileName}`);
    }

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
      // 1. Tự dò tìm các cột phù hợp cho Báo cáo (ưu tiên cấu hình chọn thủ công trước)
      const colManganh = quickReportManganhCol || mapping.manganh || columns.find(c => /mã\s*ngành|manganh|vsic|mã\s*nghe|manghe|ngành/i.test(c)) || "";
      const colXa = quickReportXaCol || mapping.xa || columns.find(c => /xã|xa|địa\s*bàn|dia\s*ban|phường|phuong/i.test(c)) || "";
      const colDoanhThu = quickReportDoanhThuCol || mapping.doanhthu || columns.find(c => /doanh\s*thu|doanhthu|thu\s*nhập|thunhap|tiền|tien/i.test(c)) || "";
      const colLaoDong = quickReportLaoDongCol || mapping.laodong || columns.find(c => /lao\s*động|laodong|người|nguoi|nhân\s*sự|nhansu/i.test(c)) || "";

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
          model: "gemini-2.5-flash",
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
      if (!fileName.startsWith("QuetAI_")) {
        setFileName(`QuetAI_${fileName}`);
      }

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
    const exportRawData = (searchTerm || rowIndicesFilter) ? filteredData : mainData;
    if (exportRawData.length === 0) {
      alert("Không có dữ liệu để xuất file!");
      return;
    }
    setLoading(true);
    setStatusMessage("Đang tạo tệp Excel phục vụ tải xuống (Gồm tệp tính toán & Danh mục ngành VSIC chuẩn)...");

    setTimeout(() => {
      try {
        const exportData = exportRawData.map((row) => {
          const originalIdx = mainData.indexOf(row);
          if (originalIdx !== -1 && rowInconStatusMap.has(originalIdx)) {
            const info = rowInconStatusMap.get(originalIdx)!;
            const newRow: any = {};
            // Đưa các cột đánh dấu lên đầu để nhìn thấy ngay
            newRow["ĐÁNH DẤU SAI LỆCH (MÃ THIỂU SỐ)"] = info.isMinority ? "⚠️ THIỂU SỐ (Chọn ít hơn - Nghi ngờ gán sai)" : "✓ ĐA SỐ (Phổ biến nhất)";
            newRow["GỢI Ý MÃ VSIC ĐÚNG (Mã đa số)"] = info.majorityCode;
            newRow["SỐ DÒNG CÙNG MÃ NÀY"] = info.countOfThisCode;
            newRow["TỔNG SỐ DÒNG CÙNG MÔ TẢ"] = info.totalOccurrences;
            newRow["DÒNG SỐ TRONG FILE GỐC"] = originalIdx + 1;
            
            Object.keys(row).forEach(key => {
              if (key !== "_satisfiesIf" && key !== "_violated") {
                newRow[key] = row[key];
              }
            });
            return newRow;
          }
          return row;
        });

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

  // Hàm xuất Excel phân tích các hộ mâu thuẫn cùng mô tả nhưng khác mã ngành
  const handleExportInconsistentExcel = (descToCodes: any[]) => {
    if (!descToCodes || descToCodes.length === 0) {
      alert("Không có dữ liệu mâu thuẫn để xuất!");
      return;
    }
    setLoading(true);
    setStatusMessage("Đang chuẩn bị danh sách mâu thuẫn để xuất Excel...");

    setTimeout(() => {
      try {
        const targetMota = stdDescriptionCol || mapping.mota;
        const targetManganh = stdIndustryCol || mapping.manganh;

        // Xây dựng danh sách các dòng kèm phân tích
        const resultRows: any[] = [];

        descToCodes.forEach(item => {
          // Tìm maxCount
          const maxCount = Math.max(...item.codes.map((c: any) => c.count));
          // Tìm mã có maxCount (mã đa số)
          const majorityCodeObj = item.codes.find((c: any) => c.count === maxCount);
          const majorityCode = majorityCodeObj ? majorityCodeObj.code : "";

          item.codes.forEach((c: any) => {
            const isMinority = c.count < maxCount;
            c.rows.forEach((rIdx: number) => {
              const originalRow = mainData[rIdx];
              if (!originalRow) return;

              // Tạo dòng mới kết hợp thông tin phân tích
              const analysisObj: any = {};
              
              // Đưa các thông tin phân tích lên đầu
              analysisObj["Mô tả hoạt động"] = item.motaText;
              analysisObj["Mã VSIC hiện tại"] = originalRow[targetManganh] || "";
              analysisObj["Mã đa số (Gợi ý)"] = majorityCode;
              analysisObj["Trạng thái phân gán"] = isMinority ? "⚠️ THIỂU SỐ (Nghi ngờ gán sai)" : "✓ ĐA SỐ (Phổ biến nhất)";
              analysisObj["Số dòng cùng mã"] = c.count;
              analysisObj["Tổng số dòng cùng mô tả"] = item.occurrences;
              analysisObj["Dòng số (Trong file gốc)"] = rIdx + 1;

              // Copy các thuộc tính còn lại từ dòng gốc
              Object.keys(originalRow).forEach(key => {
                // Tránh trùng các cột phân tích đã thêm
                if (
                  key !== "_satisfiesIf" && 
                  key !== "_violated" &&
                  key !== "Mô tả hoạt động" &&
                  key !== "Mã VSIC hiện tại"
                ) {
                  analysisObj[key] = originalRow[key];
                }
              });

              resultRows.push(analysisObj);
            });
          });
        });

        // Sắp xếp resultRows theo "Mô tả hoạt động" để các dòng cùng mô tả nằm cạnh nhau
        resultRows.sort((a, b) => String(a["Mô tả hoạt động"]).localeCompare(String(b["Mô tả hoạt động"])));

        const ws = XLSX.utils.json_to_sheet(resultRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Mau_Thuan_Mota_Nganh");

        // Thêm danh sách tóm tắt các nhóm mâu thuẫn để họ nhìn tổng quan
        const summaryRows = descToCodes.map((item, idx) => {
          const maxCount = Math.max(...item.codes.map((c: any) => c.count));
          const majorityCodeObj = item.codes.find((c: any) => c.count === maxCount);
          const minorityCodes = item.codes.filter((c: any) => c.count < maxCount).map((c: any) => `${c.code} (${c.count} dòng)`).join(", ");
          
          return {
            "STT": idx + 1,
            "Mô tả hoạt động kinh tế": item.motaText,
            "Tổng số dòng": item.occurrences,
            "Mã đa số phổ biến nhất": majorityCodeObj ? `${majorityCodeObj.code} (${majorityCodeObj.count} dòng)` : "",
            "Các mã thiểu số lệch biệt": minorityCodes
          };
        });
        const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
        XLSX.utils.book_append_sheet(wb, wsSummary, "Tom_Tat_Mau_Thuan");

        const outName = "Bao_Cao_Mau_Thuan_Cung_Mota_Khac_Ma_VSIC.xlsx";
        XLSX.writeFile(wb, outName);
        setStatusMessage(`Đã xuất Excel danh sách mâu thuẫn thành công! (${resultRows.length} dòng)`);
      } catch (err: any) {
        alert("Lỗi khi kết xuất Excel mâu thuẫn: " + err.message);
      } finally {
        setLoading(false);
      }
    }, 200);
  };

  if (!isAuthorized) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-[#f1f5f9] text-slate-800 font-sans px-4 selection:bg-purple-600 selection:text-white">
        {/* Khóa bảo mật phi hành trạm VSIC */}
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-2xl p-8 shadow-2xl space-y-6 relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-1.5 bg-gradient-to-r from-purple-600 to-indigo-500"></div>
          
          <div className="text-center space-y-2">
            <div className="mx-auto w-14 h-14 bg-gradient-to-tr from-purple-600 to-indigo-500 rounded-2xl flex items-center justify-center shadow-lg shadow-purple-900/10">
              <Lock className="w-7 h-7 text-white animate-pulse" />
            </div>
            <h2 className="text-xl font-bold tracking-tight text-slate-900 pt-2">CỔNG BẢO MẬT TRUY CẬP</h2>
            <p className="text-xs text-slate-500">Vui lòng nhập mật khẩu nội bộ để sử dụng hệ thống VSIC</p>
          </div>

          <form onSubmit={handleCheckPassword} className="space-y-4">
            <div className="space-y-1.5">
              <label className="text-xs text-slate-600 font-bold font-mono">MẬT KHẨU TRUY CẬP:</label>
              <input
                type="password"
                value={typedPassword}
                onChange={(e) => setTypedPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm text-slate-900 focus:outline-none focus:border-purple-500 focus:ring-1 focus:ring-purple-500 transition-all placeholder:text-slate-400 font-mono"
                autoFocus
              />
              {passwordError && (
                <p className="text-red-600 text-[11px] font-semibold flex items-center gap-1 mt-1 font-mono">
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

          <div className="border-t border-slate-200 pt-4 text-center space-y-2">
            <p className="text-[11px] text-amber-600 italic font-semibold">
              💡 Gợi ý mật khẩu truy cập:
            </p>
            <div className="flex flex-col gap-1 items-center justify-center font-mono text-[11px] text-slate-600">
              <div>• Quản trị viên: <strong className="bg-amber-50 px-1.5 py-0.5 rounded border border-amber-200 text-amber-700">admin123</strong> (Được quyền đổi MK)</div>
              <div>• Dùng chung: <strong className="bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200 text-blue-700 font-bold">user123</strong> (Chỉ xem, Không thể đổi MK)</div>
            </div>
            <p className="text-[10px] text-slate-400 font-mono pt-1">Hệ thống bảo lưu mã khóa cục bộ an toàn trong trình duyệt của bạn</p>
          </div>
        </div>
      </div>
    );
  }

  const renderAiMacroCognitiveCenter = () => {
    // Chỉ hiển thị các lệnh học tương ứng với tab phân hệ đang thiết lập để tránh làm loãng hoặc xung đột giao diện
    const displayedMacros = aiMacros.filter(macro => macro.module === activeTab);

    return (
      <div className="bg-white border border-slate-200 rounded-2xl p-6 mb-6 space-y-6 shadow-sm relative overflow-hidden">
        <div className="absolute top-0 right-0 w-80 h-80 bg-indigo-50/20 rounded-full blur-3xl -z-10 pointer-events-none"></div>
        <div className="absolute bottom-0 left-0 w-80 h-80 bg-purple-50/20 rounded-full blur-3xl -z-10 pointer-events-none"></div>

        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-indigo-50 border border-indigo-150 rounded-xl shadow-inner">
              <Brain className="w-6 h-6 text-indigo-600 animate-pulse" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-800 uppercase tracking-wide flex items-center gap-2 font-mono">
                BỘ NÃO HỌC LỆNH THÔNG MIÊN (AI MACRO STORAGE ENGINE)
                <span className="text-[10px] bg-indigo-50 text-indigo-700 font-mono font-medium px-2 py-0.5 rounded-full border border-indigo-150">
                  DUAL-PHÂN HỆ v3.2
                </span>
              </h2>
              <p className="text-xs text-slate-500 mt-1 font-sans">
                Huấn luyện tự động nhận diện và ghép nối cột qua khẩu lệnh tiếng Việt. Lưu bộ nhớ vạn năng để kích hoạt 0 giây.
              </p>
            </div>
          </div>

          {/* Export/Import Buttons */}
          <div className="flex items-center gap-2">
            <button
              onClick={handleExportMacros}
              className="px-3 py-1.5 bg-white hover:bg-indigo-50 text-indigo-700 font-bold text-xs rounded-xl border border-slate-200 transition-all cursor-pointer flex items-center gap-1.5"
              title="Tải tệp lưu các quy tắc đã học về máy tính"
            >
              <Download className="w-3.5 h-3.5 text-indigo-600" /> Xuất Bộ Nhớ (.json)
            </button>
            
            <label className="px-3 py-1.5 bg-white hover:bg-indigo-50 text-indigo-700 font-bold text-xs rounded-xl border border-slate-200 transition-all cursor-pointer flex items-center gap-1.5">
              <Upload className="w-3.5 h-3.5 text-teal-650" /> Nạp Cấu Hình
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
        <div className="bg-slate-50 border border-slate-200 p-5 rounded-xl space-y-4">
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-indigo-700 flex items-center gap-1.5 font-mono">
              <Sparkles className="w-3.5 h-3.5 text-indigo-500" /> NHẬP KHẨU LỆNH ĐỂ AI TỰ HỌC CHỌN CỘT:
            </label>
            <div className="flex flex-col sm:flex-row gap-2">
              <input
                type="text"
                value={macroPrompt}
                onChange={(e) => setMacroPrompt(e.target.value)}
                placeholder="Ví dụ: 'Tổng hợp theo xã ngành cấp 2 cho doanh thu và lao động' hoặc 'Kiểm tra mã ngành và mô tả'"
                className="flex-1 bg-white border border-slate-300 rounded-xl px-4 py-2.5 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 font-sans shadow-sm"
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
            <p className="text-[10.5px] text-slate-500 font-sans leading-normal">
              * Khuyên dùng: Vui lòng nạp dữ liệu chính trước ở Trang Chủ hoặc Tab "Xem &amp; Định Nghĩa Cột" sau đó gửi yêu cầu học. Trực tiếp so khớp các cột có sẵn trong file của bạn.
            </p>
          </div>

          {/* AI learning feedback */}
          {learningResult && (
            <div className="bg-white border border-emerald-200 p-4.5 rounded-xl space-y-3 animate-fade-in shadow-sm">
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-200 pb-2">
                <span className="text-xs font-bold text-emerald-700 flex items-center gap-1.5 font-mono">
                  <CheckCircle2 className="w-4 h-4" /> KẾT QUẢ PHÂN TÍCH CHỌN CỘT THÀNH CÔNG:
                </span>
                <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200">
                  PHÂN HỆ MỤC TIÊU: {learningResult.module === "tonghop" ? "TỔNG HỢP BÁO CÁO ĐỘNG" : "KIỂM TRA CHUẨN VSIC"}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 bg-slate-50 p-3 rounded-lg border border-slate-200">
                {learningResult.module === "tonghop" ? (
                  <>
                    <div>
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Địa Bàn Xã</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.xa || "⚠️ (Trống/Bỏ qua)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Mã Ngành</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.manganh || "⚠️ (Trống/Bỏ qua)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Doanh Thu</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.doanhthu || "⚠️ (Không tính)"}
                      </div>
                    </div>
                    <div>
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Lao Động</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.laodong || "⚠️ (Không tính)"}
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="sm:col-span-2">
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Mô Tả Thực Tế</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.mota || "⚠️ (Không thấy)"}
                      </div>
                    </div>
                    <div className="sm:col-span-2">
                      <div className="text-[9.5px] text-slate-500 uppercase font-mono font-bold">Cột Mã Ngành VSIC</div>
                      <div className="text-xs text-slate-800 font-semibold font-mono mt-0.5 truncate bg-white px-1.5 py-1 rounded border border-slate-200">
                        {learningResult.columns.manganh || "⚠️ (Không thấy)"}
                      </div>
                    </div>
                  </>
                )}
              </div>

              <div className="flex flex-col sm:flex-row items-end gap-3 pt-1">
                <div className="flex-1 space-y-1 w-full">
                  <label className="text-[10px] font-bold text-slate-500 font-mono block uppercase">ĐẶT TÊN TUỲ CHỈNH CHO LỆNH ĐỂ LƯU THƯ VIỆN:</label>
                  <input
                    type="text"
                    value={customMacroName}
                    onChange={(e) => setCustomMacroName(e.target.value)}
                    placeholder="Ví dụ: Báo cáo Doanh thu Xã, Kiểm tra VSIC,..."
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 font-sans font-bold shadow-sm"
                  />
                </div>
                <button
                  onClick={handleSaveLearnMacro}
                  className="w-full sm:w-auto bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-xs px-5 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow border border-emerald-500/20 shrink-0 font-sans h-9.5"
                >
                  <Save className="w-4 h-4" /> [Lưu học lệnh này]
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Learned Commands Library (Tái sử dụng 0 giây) */}
        <div className="space-y-3 pt-1">
          <div className="text-xs font-bold text-indigo-700 tracking-wider uppercase font-mono flex items-center gap-1.5">
            <Cpu className="w-4 h-4 text-purple-650" /> THƯ VIỆN LỆNH ĐÃ HỌC (TÁI SỬ DỤNG 0 GIÂY - KHÔNG GỌI LẠI GEMINI):
          </div>

          {displayedMacros.length === 0 ? (
            <div className="bg-slate-50 rounded-xl p-5 text-center text-xs text-slate-500 border border-slate-200">
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
                    className="group bg-white hover:bg-indigo-50/50 border border-slate-200 hover:border-indigo-300 rounded-xl p-4 transition-all cursor-pointer flex flex-col justify-between gap-3 shadow-sm hover:shadow-md relative"
                  >
                    <div className="space-y-2">
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-sans font-extrabold text-xs text-slate-800 group-hover:text-indigo-700 transition-colors line-clamp-1">
                          {macro.name}
                        </span>
                        <div className="flex items-center gap-1 shrink-0">
                          <span className={`text-[9px] font-mono px-1.5 py-0.5 rounded font-bold ${
                            isTongHop 
                              ? "bg-amber-50 text-amber-750 border border-amber-200" 
                              : "bg-indigo-50 text-indigo-750 border border-indigo-200"
                          }`}>
                            {isTongHop ? "Báo cáo" : "VSIC"}
                          </span>
                          <button
                            onClick={(e) => handleDeleteMacro(macro.id, e)}
                            className="text-slate-400 hover:text-red-500 p-0.5 rounded hover:bg-red-50 transition-all cursor-pointer"
                            title="Xóa lệnh này"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                      <p className="text-[10.5px] text-slate-500 line-clamp-2 leading-relaxed font-sans">
                        "{macro.prompt}"
                      </p>
                    </div>

                    <div className="flex flex-wrap gap-1.5 pt-2 border-t border-slate-150 text-[9px] font-mono text-slate-500">
                      {isTongHop ? (
                        <>
                          {macro.columns?.xa && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-slate-200">Xã: {macro.columns.xa}</span>}
                          {macro.columns?.manganh && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-slate-200">Mã: {macro.columns.manganh}</span>}
                          {macro.columns?.doanhthu && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-slate-200">DT: {macro.columns.doanhthu}</span>}
                          {macro.columns?.laodong && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[85px] border border-slate-200">LĐ: {macro.columns.laodong}</span>}
                        </>
                      ) : (
                        <>
                          {macro.columns?.mota && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[120px] border border-slate-200 text-indigo-650">Mô tả: {macro.columns.mota}</span>}
                          {macro.columns?.manganh && <span className="bg-slate-50 px-1.5 py-0.5 rounded truncate max-w-[120px] border border-slate-200">Mã: {macro.columns.manganh}</span>}
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

  {/* Panel phía trên cùng thanh menu hiển thị Logo & Banner */}
  const renderHeaderPanel = () => {
    return (
      <div className="bg-gradient-to-r from-slate-50 via-white to-indigo-50/30 border-b border-indigo-100 shadow-sm relative z-40 px-6 py-3 flex flex-wrap items-center justify-between gap-4 select-none animate-fade-in">
        {/* Left: Logo & Title */}
        <div className="flex items-center gap-3.5">
          <div className="relative group/logo">
            <img 
              src={logoImg} 
              alt="VSIC Logo" 
              className="h-12 w-12 rounded-full object-cover shadow-md border-2 border-indigo-100 group-hover/logo:scale-105 transition-transform"
              onError={(e) => {
                e.currentTarget.src = "https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=120&q=80";
              }}
            />
            <div className="absolute -bottom-1 -right-1 bg-emerald-500 border border-white h-3.5 w-3.5 rounded-full flex items-center justify-center" title="Hệ thống trực tuyến">
              <span className="animate-ping absolute inline-flex h-2 w-2 rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-500"></span>
            </div>
          </div>
          <div>
            <div className="text-sm font-extrabold text-indigo-950 uppercase tracking-wide flex items-center gap-2">
              <span>Nền tảng Trí tuệ VSIC &amp; Kiểm soát Dữ liệu Liên ngành</span>
              <span className="bg-indigo-100 text-indigo-800 text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wider hidden sm:inline-block">v4.0 Pro</span>
            </div>
            <div className="text-xs text-slate-500 font-medium">
              Xử lý chuẩn hóa mã ngành, rà soát lô-gích thống kê toàn diện &amp; đối chiếu niên độ động cao cấp
            </div>
          </div>
        </div>

        {/* Center: Live Interactive/Animated Analytics & Processing Pipeline */}
        <div className="hidden lg:flex items-center gap-5 bg-white/70 border border-indigo-100/40 px-4 py-1.5 rounded-2xl shadow-[0_2px_10px_-3px_rgba(79,70,229,0.06)] max-w-sm xl:max-w-md flex-1 mx-4">
          {/* Live indicator & Pulse Line */}
          <div className="flex flex-col shrink-0 select-none">
            <div className="flex items-center gap-1.5">
              <span className="relative flex h-2 w-2">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-500 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-indigo-600"></span>
              </span>
              <span className="text-[10px] text-indigo-950 font-extrabold uppercase tracking-wider">Bộ xử lý trung tâm</span>
            </div>
            
            {/* Real-time stats text */}
            <div className="text-[11px] text-slate-500 font-bold flex items-center gap-1.5 mt-0.5">
              <span className="text-emerald-600">Đang hoạt động</span>
              <span className="text-slate-300">•</span>
              <span className="text-indigo-650 font-extrabold" title="Tổng số lượt tra cứu và thao tác phân tích dữ liệu">{totalQueries.toLocaleString("vi-VN")} lượt</span>
            </div>
          </div>

          {/* Animated SVG Sparkline Wave */}
          <div className="flex-1 h-8 relative flex items-center justify-center min-w-[120px] overflow-hidden rounded-lg bg-indigo-50/20 px-2 border border-indigo-100/20">
            <svg viewBox="0 0 100 30" className="w-full h-full overflow-visible">
              <defs>
                <linearGradient id="waveGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                  <stop offset="0%" stopColor="#3b82f6" stopOpacity="0.2" />
                  <stop offset="50%" stopColor="#4f46e5" stopOpacity="0.9" />
                  <stop offset="100%" stopColor="#06b6d4" stopOpacity="0.2" />
                </linearGradient>
                <linearGradient id="fillGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stopColor="#4f46e5" stopOpacity="0.1" />
                  <stop offset="100%" stopColor="#06b6d4" stopOpacity="0.0" />
                </linearGradient>
              </defs>
              
              {/* Background area wave */}
              <path 
                d="M 0 22 C 15 12, 25 32, 40 15 C 55 2, 65 25, 80 10 C 90 0, 95 15, 100 12 L 100 30 L 0 30 Z" 
                fill="url(#fillGrad)"
                className="animate-pulse"
              />
              
              {/* Glowing active path line */}
              <path 
                d="M 0 22 C 15 12, 25 32, 40 15 C 55 2, 65 25, 80 10 C 90 0, 95 15, 100 12" 
                fill="none" 
                stroke="url(#waveGrad)" 
                strokeWidth="2"
                strokeLinecap="round"
                style={{
                  strokeDasharray: '200',
                  strokeDashoffset: '0',
                  animation: 'shiftPath 4s linear infinite'
                }}
              />
              
              {/* Moving dot tracker */}
              <circle r="2.5" fill="#4f46e5" className="animate-ping" style={{ transformOrigin: 'center' }}>
                <animateMotion 
                  path="M 0 22 C 15 12, 25 32, 40 15 C 55 2, 65 25, 80 10 C 90 0, 95 15, 100 12" 
                  dur="4s" 
                  repeatCount="indefinite" 
                />
              </circle>
              <circle r="1.5" fill="#06b6d4">
                <animateMotion 
                  path="M 0 22 C 15 12, 25 32, 40 15 C 55 2, 65 25, 80 10 C 90 0, 95 15, 100 12" 
                  dur="4s" 
                  repeatCount="indefinite" 
                />
              </circle>
            </svg>

            {/* Custom keyframe animation style tag */}
            <style>{`
              @keyframes shiftPath {
                0% { strokeDashoffset: 200; }
                100% { strokeDashoffset: 0; }
              }
            `}</style>
          </div>

          {/* Equalizer Wave simulation */}
          <div className="flex items-end gap-1 h-7 px-1 shrink-0">
            <div className="w-1 bg-indigo-500 rounded-full animate-pulse h-3" style={{ animationDelay: '0.1s', animationDuration: '0.8s' }}></div>
            <div className="w-1 bg-violet-500 rounded-full animate-pulse h-5" style={{ animationDelay: '0.3s', animationDuration: '1.2s' }}></div>
            <div className="w-1 bg-sky-500 rounded-full animate-pulse h-6" style={{ animationDelay: '0.5s', animationDuration: '0.9s' }}></div>
            <div className="w-1 bg-emerald-500 rounded-full animate-pulse h-4" style={{ animationDelay: '0.2s', animationDuration: '1.1s' }}></div>
            <div className="w-1 bg-teal-500 rounded-full animate-pulse h-2" style={{ animationDelay: '0.4s', animationDuration: '0.7s' }}></div>
          </div>
        </div>

        {/* Right: Banner Image & User Controls */}
        <div className="flex items-center gap-4 flex-wrap sm:flex-nowrap">
          {/* Banner decoration */}
          <div className="hidden md:block relative h-11 w-48 rounded-lg overflow-hidden border border-indigo-100/50 shadow-inner">
            <img 
              src={bannerImg} 
              alt="VSIC Banner" 
              className="h-full w-full object-cover opacity-90 hover:opacity-100 transition-opacity"
              onError={(e) => {
                e.currentTarget.src = "https://images.unsplash.com/photo-1460925895917-afdab827c52f?auto=format&fit=crop&w=240&q=80";
              }}
            />
          </div>

          {/* Hệ thống Công khai - Không yêu cầu Tài khoản */}
          <div className="flex items-center gap-2 bg-emerald-50 border border-emerald-200/60 px-3.5 py-1.5 rounded-xl shadow-sm">
            <div className="text-right">
              <div className="text-[10px] text-emerald-600 font-extrabold uppercase tracking-wider flex items-center gap-1 justify-end">
                <span className="relative flex h-1.5 w-1.5">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-500"></span>
                </span>
                Hệ Thống
              </div>
              <div className="text-xs font-extrabold text-emerald-800 flex items-center gap-1">
                <Globe className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                Mở / Công khai
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  {/* Thanh Menu Ngang (Sub-header Navigation) sang trọng thay thế hoàn toàn Sidebar dọc */}
  const renderHorizontalMenu = () => {
    return (
      <div className="bg-indigo-50/95 backdrop-blur-md border-b border-indigo-200/80 shadow-md relative z-30 px-6 py-1.5 flex flex-nowrap items-center justify-between gap-3 overflow-x-auto lg:overflow-visible select-none animate-fade-in custom-scrollbar">
        <div className="flex items-center gap-1.5 flex-nowrap shrink-0">
          
          {/* Nút TRANG CHỦ */}
          <button 
            onClick={(e) => {
              e.stopPropagation();
              setActiveTab("trangchu");
              setOpenDropdown(null);
            }}
            className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
              activeTab === "trangchu" 
                ? "bg-indigo-600 text-white shadow-md border border-indigo-700" 
                : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
            }`}
          >
            <Home className="w-4 h-4 shrink-0 text-indigo-500" />
            Trang chủ
          </button>

          {/* NÚT CHÍNH: PHIẾU KHẢO SÁT & CHỌN MẪU BIỂU */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "phieukhaosat" ? null : "phieukhaosat");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["dataentry", "chonmau"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "phieukhaosat"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "bg-emerald-50 text-emerald-800 hover:bg-emerald-100 border border-emerald-200 shadow-sm"
              }`}
            >
              <span className="relative flex h-2 w-2">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
              </span>
              <span>✍️ Phiếu Khảo sát</span>
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 transition-transform ${openDropdown === "phieukhaosat" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "phieukhaosat" && (
              <div 
                className="absolute left-0 mt-1.5 w-64 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("dataentry"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "dataentry" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <FileText className="w-4 h-4 text-emerald-500 shrink-0" />
                  Khai báo &amp; Ký số Hưng Yên
                </button>
                <button 
                  onClick={() => { setActiveTab("chonmau"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "chonmau" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <FileCheck className="w-4 h-4 text-orange-500 shrink-0" />
                  Thiết lập Biểu mẫu khảo sát
                </button>
              </div>
            )}
          </div>

          {/* DROPDOWN 1: TRẠM DỮ LIỆU */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "quanlytep" ? null : "quanlytep");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["xemdulieu", "ghepnoi", "tachfile", "sosanh"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "quanlytep"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
              }`}
            >
              <FileSpreadsheet className="w-4 h-4 shrink-0 text-indigo-500" />
              📂 Trạm Dữ liệu
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 text-slate-500 transition-transform ${openDropdown === "quanlytep" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "quanlytep" && (
              <div 
                className="absolute left-0 mt-1.5 w-60 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("xemdulieu"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "xemdulieu" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <FileSpreadsheet className="w-4 h-4 text-sky-500 shrink-0" />
                  Cấu trúc &amp; Định nghĩa cột
                </button>
                <button 
                  onClick={() => { setActiveTab("ghepnoi"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "ghepnoi" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <GitMerge className="w-4 h-4 text-emerald-500 shrink-0" />
                  Hợp nhất Dữ liệu (Left Join &amp; Batch)
                </button>
                <button 
                  onClick={() => { setActiveTab("tachfile"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "tachfile" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Scissors className="w-4 h-4 text-purple-500 shrink-0" />
                  Phân rã File Hàng loạt
                </button>
                <button 
                  onClick={() => { setActiveTab("sosanh"); setOpenDropdown(null); }}
                  className={`w-full flex items-center justify-between px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "sosanh" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <div className="flex items-center gap-2.5">
                    <Database className="w-4 h-4 text-amber-500 shrink-0" />
                    <span>So sánh Dữ liệu giữa các kỳ</span>
                  </div>
                  <span className="bg-red-500 text-white text-[8px] px-1 py-0.5 rounded-full font-bold uppercase shrink-0 animate-pulse">MỚI</span>
                </button>
              </div>
            )}
          </div>

          {/* DROPDOWN 2: RÀ SOÁT & KIỂM TRA */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "rasoat" ? null : "rasoat");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["kiemtralogic", "outliers", "rulesstudio"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "rasoat"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
              }`}
            >
              <CheckSquare className="w-4 h-4 shrink-0 text-indigo-500" />
              🔍 Rà soát &amp; Kiểm tra
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 text-slate-500 transition-transform ${openDropdown === "rasoat" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "rasoat" && (
              <div 
                className="absolute left-0 mt-1.5 w-64 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("kiemtralogic"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "kiemtralogic" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Brain className="w-4 h-4 text-rose-500 shrink-0" />
                  Kiểm tra logic đa điều kiện
                </button>
                <button 
                  onClick={() => { setActiveTab("outliers"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "outliers" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Activity className="w-4 h-4 text-amber-500 shrink-0" />
                  Quét lệch quy luật phân phối
                </button>
                <button 
                  onClick={() => { setActiveTab("rulesstudio"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "rulesstudio" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Sliders className="w-4 h-4 text-blue-500 shrink-0" />
                  Quản trị quy tắc logic
                </button>
              </div>
            )}
          </div>

          {/* DROPDOWN 3: TRA CỨU & CHUẨN HÓA VSIC */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "vsic" ? null : "vsic");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["chuanhoanganh", "danhmucvsic"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "vsic"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
              }`}
            >
              <BrainCircuit className="w-4 h-4 shrink-0 text-indigo-500" />
              🏷️ Chuẩn hóa VSIC
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 text-slate-500 transition-transform ${openDropdown === "vsic" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "vsic" && (
              <div 
                className="absolute left-0 mt-1.5 w-60 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("chuanhoanganh"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "chuanhoanganh" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Sparkles className="w-4 h-4 text-emerald-500 shrink-0" />
                  Chuẩn hóa khớp ngành VSIC
                </button>
                <button 
                  onClick={() => { setActiveTab("danhmucvsic"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "danhmucvsic" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Search className="w-4 h-4 text-sky-500 shrink-0" />
                  Tra cứu danh mục VSIC chuẩn
                </button>
              </div>
            )}
          </div>

          {/* DROPDOWN 4: PHÂN TÍCH & TỔNG HỢP */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "phantich" ? null : "phantich");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["tonghop", "tansuat", "tuongquan"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "phantich"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
              }`}
            >
              <BarChart3 className="w-4 h-4 shrink-0 text-indigo-500" />
              📊 Phân tích &amp; Tổng hợp
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 text-slate-500 transition-transform ${openDropdown === "phantich" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "phantich" && (
              <div 
                className="absolute left-0 mt-1.5 w-60 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("tonghop"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "tonghop" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <PieChart className="w-4 h-4 text-sky-500 shrink-0" />
                  Báo cáo &amp; Tổng hợp biểu đồ
                </button>
                <button 
                  onClick={() => { setActiveTab("tansuat"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "tansuat" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Activity className="w-4 h-4 text-rose-500 shrink-0" />
                  Phân tích tần suất mẫu
                </button>
                <button 
                  onClick={() => { setActiveTab("tuongquan"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "tuongquan" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <ArrowRightLeft className="w-4 h-4 text-purple-500 shrink-0" />
                  Phân tích tương quan dữ liệu
                </button>
              </div>
            )}
          </div>

          {/* DROPDOWN 5: TIỆN ÍCH & CỘNG TÁC */}
          <div className="relative">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setOpenDropdown(openDropdown === "congtac" ? null : "congtac");
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all duration-150 cursor-pointer ${
                ["pdf2word", "videoroom", "admindashboard"].includes(activeTab)
                  ? "bg-indigo-600 text-white shadow-md border border-indigo-700"
                  : openDropdown === "congtac"
                    ? "bg-indigo-100 text-indigo-900 border border-indigo-200"
                    : "text-slate-700 hover:bg-indigo-100/50 hover:text-indigo-900 border border-transparent"
              }`}
            >
              <Users className="w-4 h-4 shrink-0 text-indigo-500" />
              🛠️ Tiện ích &amp; Cộng tác
              <ChevronDown className={`w-3.5 h-3.5 shrink-0 text-slate-500 transition-transform ${openDropdown === "congtac" ? "rotate-180" : ""}`} />
            </button>
            {openDropdown === "congtac" && (
              <div 
                className="absolute right-0 lg:left-auto mt-1.5 w-60 bg-white text-slate-850 rounded-xl shadow-2xl border border-indigo-100 py-2.5 z-50 animate-fade-in"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  onClick={() => { setActiveTab("pdf2word"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "pdf2word" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <FileText className="w-4 h-4 text-amber-500 shrink-0" />
                  Đọc PDF &amp; Chuyển sang Word
                </button>
                <button 
                  onClick={() => { setActiveTab("videoroom"); setOpenDropdown(null); }}
                  className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "videoroom" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                >
                  <Video className="w-4 h-4 text-emerald-500 shrink-0" />
                  Phòng họp video trực tuyến
                </button>
                {userRole === "admin" && (
                  <button 
                    onClick={() => { setActiveTab("admindashboard"); setOpenDropdown(null); }}
                    className={`w-full flex items-center gap-2.5 px-4 py-2 text-left text-xs font-bold transition-colors hover:bg-indigo-50/60 ${activeTab === "admindashboard" ? "text-indigo-600 bg-indigo-50" : "text-slate-700"}`}
                  >
                    <Shield className="w-4 h-4 text-indigo-500 shrink-0" />
                    Trình quản trị hệ thống
                  </button>
                )}
              </div>
            )}
          </div>

        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] flex flex-col">
      <div className="sticky top-0 z-50 flex flex-col shadow-md">
        {renderHeaderPanel()}
        {renderHorizontalMenu()}
      </div>

      {/* Main Layout split: Sidebar + Content */}
      <div className="flex flex-1 overflow-hidden">
        
        {/* Content Area */}
        <main className="flex-1 bg-[#f1f5f9] overflow-y-auto p-6 md:p-8 custom-scrollbar">
          
          {/* Lớp hiển thị nạp dữ liệu/ tiến trình hệ thống khi chạy */}
          {loading && (
            <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-white border border-slate-200 rounded-2xl max-w-md w-full p-6 text-center space-y-4 shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 h-1 bg-gradient-to-r from-indigo-500 via-sky-500 to-emerald-400 transition-all duration-300" style={{ width: `${progress}%` }}></div>
                <Loader2 className="w-12 h-12 text-indigo-600 mx-auto animate-spin" />
                <h3 className="text-lg font-bold text-slate-850 font-sans">Đang xử lý dữ liệu</h3>
                <p className="text-sm text-slate-500 font-mono leading-relaxed min-h-[40px]">{statusMessage}</p>
                
                <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden border border-slate-200">
                  <div className="bg-indigo-600 h-2.5 rounded-full transition-all duration-300" style={{ width: `${progress}%` }}></div>
                </div>
                <div className="text-xs font-bold text-indigo-600 tracking-wider font-mono">{progress}% Hoàn Thành</div>
              </div>
            </div>
          )}

          {/* Lớp hiển thị đổi mật khẩu truy cập */}
          {showPasswordChangeModal && (
            <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-6">
              <div className="bg-white border border-slate-200 rounded-2xl max-w-sm w-full p-6 space-y-4 shadow-2xl relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-indigo-500 to-sky-500"></div>
                <div className="text-center space-y-1">
                  <div className="mx-auto w-10.5 h-10.5 bg-indigo-50 border border-indigo-100 rounded-xl flex items-center justify-center">
                    <KeyRound className="w-5.5 h-5.5 text-indigo-600" />
                  </div>
                  <h3 className="text-base font-bold text-slate-850 pt-1">ĐỔI MẬT KHẨU BẢO VỆ</h3>
                  <p className="text-xs text-slate-500 text-center font-sans">Thiết lập mật khẩu riêng tư cho trình quản lý</p>
                </div>

                <div className="space-y-4">
                  <div className="space-y-1.5 font-sans">
                    <label className="text-[10.5px] font-bold text-slate-700 block">MẬT KHẨU MỚI TIN CẬY:</label>
                    <input
                      type="text"
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-mono shadow-inner"
                      placeholder="Nhập mật khẩu mới..."
                      value={newPasswordVal}
                      onChange={(e) => setNewPasswordVal(e.target.value)}
                      autoFocus
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-2 pt-2">
                    <button
                      onClick={() => setShowPasswordChangeModal(false)}
                      className="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold text-xs py-2 px-3 rounded-lg border border-slate-200 transition-all cursor-pointer font-sans"
                    >
                      Hủy Bỏ
                    </button>
                    <button
                      onClick={handleChangePassword}
                      className="w-full bg-gradient-to-r from-indigo-500 to-sky-500 hover:from-indigo-600 hover:to-sky-600 text-white font-bold text-xs py-2 px-3 rounded-lg shadow-md transition-all cursor-pointer font-sans active:scale-95"
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
              {/* 2. PIPELINE STEPS */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 border-b border-slate-150 pb-2">
                  <Layers className="w-5 h-5 text-indigo-600" />
                  <h3 className="text-base font-black text-slate-800 uppercase tracking-tight">
                    ⚙️ Quy trình xử lý dữ liệu khép kín
                  </h3>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {PIPELINE_STEPS.map((step) => {
                    const isActive = selectedPipelineStep === step.id;
                    return (
                      <div 
                        key={step.id}
                        onClick={() => setSelectedPipelineStep(step.id)}
                        className={`p-4 rounded-xl border transition-all duration-200 cursor-pointer text-left relative overflow-hidden ${
                          isActive 
                            ? "bg-indigo-600 text-white border-indigo-700 shadow-md ring-2 ring-indigo-500/20" 
                            : "bg-white text-slate-700 border-slate-200 hover:border-indigo-300 hover:bg-indigo-50/20"
                        }`}
                      >
                        <div className="flex items-start justify-between">
                          <span className={`text-[10px] font-mono font-bold px-2 py-0.5 rounded-full ${
                            isActive ? "bg-indigo-700 text-white" : "bg-slate-100 text-slate-600"
                          }`}>
                            BƯỚC 0{step.id}
                          </span>
                          {isActive && <CheckCircle className="w-4 h-4 text-emerald-300 shrink-0" />}
                        </div>
                        <h4 className="font-bold text-sm mt-3 tracking-tight">{step.title}</h4>
                        <p className={`text-[11px] mt-1 leading-snug ${isActive ? "text-indigo-100" : "text-slate-500"}`}>
                          {step.shortDesc}
                        </p>
                      </div>
                    );
                  })}
                </div>

                {/* Step Details display */}
                {selectedPipelineStep && (() => {
                  const activeStepObj = PIPELINE_STEPS.find(s => s.id === selectedPipelineStep);
                  if (!activeStepObj) return null;
                  return (
                    <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-5 md:p-6 space-y-4 animate-slide-up">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-1">
                          <h4 className="text-sm font-extrabold text-slate-800 flex items-center gap-2">
                            <span className="bg-indigo-100 text-indigo-700 text-xs font-mono font-bold w-5 h-5 rounded-full flex items-center justify-center">
                              {activeStepObj.id}
                            </span>
                            {activeStepObj.fullTitle}
                          </h4>
                          <p className="text-xs text-slate-600 max-w-3xl leading-relaxed">
                            {activeStepObj.fullDesc}
                          </p>
                        </div>
                        <button
                          onClick={() => setActiveTab(activeStepObj.targetTab as any)}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-5 py-2.5 rounded-xl shadow-md flex items-center justify-center gap-1.5 shrink-0 cursor-pointer border-0 transition-colors"
                        >
                          {activeStepObj.actionText}
                        </button>
                      </div>

                      <div className="border-t border-slate-200/60 pt-4">
                        <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block mb-2 font-mono">Điểm nhấn quy trình:</span>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                          {activeStepObj.highlights.map((h, i) => (
                            <div key={i} className="flex items-start gap-2 bg-white border border-slate-200/50 p-2.5 rounded-xl">
                              <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
                              <span className="text-xs text-slate-700 font-medium leading-relaxed">{h}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>

              {/* 3. SCENARIOS GUIDE WITH OPERATIONAL SEQUENCE */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 border-b border-slate-150 pb-2">
                  <Compass className="w-5 h-5 text-indigo-600" />
                  <h3 className="text-base font-black text-slate-800 uppercase tracking-tight">
                    🎯 Trình tự thao tác và hướng dẫn theo tình huống
                  </h3>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                  {/* Left Column: List of scenarios */}
                  <div className="lg:col-span-5 space-y-2 max-h-[500px] overflow-y-auto pr-1 scrollbar-thin">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block px-1 font-mono">Chọn nhu cầu của bạn:</span>
                    {GUIDE_SCENARIOS.map((scenario) => {
                      const isActive = selectedWizardScenario === scenario.id;
                      return (
                        <button
                          key={scenario.id}
                          onClick={() => setSelectedWizardScenario(scenario.id)}
                          className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all duration-150 cursor-pointer ${
                            isActive 
                              ? "bg-purple-50 text-purple-950 border-purple-300 shadow-sm" 
                              : "bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300"
                          }`}
                        >
                          <span className="text-xl shrink-0">{scenario.icon}</span>
                          <span className={`text-xs font-bold leading-tight ${isActive ? "text-purple-800" : "text-slate-700"}`}>
                            {scenario.buttonText}
                          </span>
                        </button>
                      );
                    })}
                  </div>

                  {/* Right Column: Detailed sequence of actions */}
                  <div className="lg:col-span-7 bg-white border border-slate-200 rounded-2xl p-5 md:p-6 shadow-sm min-h-[400px] flex flex-col justify-between">
                    {selectedWizardScenario ? (() => {
                      const activeScenarioObj = GUIDE_SCENARIOS.find(s => s.id === selectedWizardScenario);
                      if (!activeScenarioObj) return null;
                      return (
                        <div className="space-y-5 flex-1 flex flex-col justify-between">
                          <div className="space-y-4">
                            {/* Scenario Title */}
                            <div className="flex items-start gap-3 pb-3 border-b border-slate-100">
                              <span className="text-3xl shrink-0">{activeScenarioObj.icon}</span>
                              <div className="space-y-1">
                                <h4 className="text-sm font-black text-slate-800 uppercase tracking-tight">
                                  {activeScenarioObj.title}
                                </h4>
                                <p className="text-[11px] text-slate-500 leading-relaxed font-medium">
                                  {activeScenarioObj.intro}
                                </p>
                              </div>
                            </div>

                            {/* Detailed steps timeline */}
                            <div className="space-y-4">
                              <span className="text-[10px] font-bold text-indigo-500 uppercase tracking-wider block font-mono">Trình tự thực hiện chi tiết:</span>
                              <div className="relative border-l border-slate-200/80 ml-3.5 pl-5 space-y-4">
                                {activeScenarioObj.steps.map((stepText, idx) => (
                                  <div key={idx} className="relative">
                                    {/* Number Circle Badge */}
                                    <span className="absolute -left-[29px] top-0 bg-indigo-50 border border-indigo-200 text-indigo-700 text-[10px] font-mono font-extrabold w-5 h-5 rounded-full flex items-center justify-center shadow-sm">
                                      {idx + 1}
                                    </span>
                                    <div className="space-y-1">
                                      <p className="text-xs text-slate-700 leading-relaxed font-medium font-sans">
                                        {stepText}
                                      </p>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>

                          {/* Action Buttons */}
                          <div className="pt-4 border-t border-slate-100 mt-6 flex flex-wrap items-center gap-3">
                            {activeScenarioObj.actionButtons.map((btn, btnIdx) => (
                              <button
                                key={btnIdx}
                                onClick={() => {
                                  setActiveTab(btn.tab as any);
                                  if (btn.stepId) {
                                    setSelectedPipelineStep(btn.stepId);
                                  }
                                }}
                                className={`font-bold text-xs px-5 py-2.5 rounded-xl transition-all shadow-md cursor-pointer border-0 flex items-center gap-2 ${
                                  btn.style === "secondary" 
                                    ? "bg-slate-100 hover:bg-slate-200 text-slate-800" 
                                    : "bg-purple-600 hover:bg-purple-700 text-white"
                                }`}
                              >
                                {btn.text} <ArrowRight className="w-3.5 h-3.5" />
                              </button>
                            ))}
                          </div>
                        </div>
                      );
                    })() : (
                      <div className="h-full flex flex-col items-center justify-center text-center p-8 text-slate-400 space-y-2">
                        <Compass className="w-10 h-10 stroke-1 text-slate-300" />
                        <p className="text-xs font-bold">Hãy chọn một tình huống ở cột bên trái để hiển thị trình tự thao tác!</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 2. TAB FILE VIEWER & COLUMN MAPPING */}
          {activeTab === "xemdulieu" && (
            <div className="space-y-6 animate-fade-in font-sans">
              
              {/* Box Upload chính */}
              <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-4 shadow-sm">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <FileSpreadsheet className="w-5 h-5 text-indigo-600" /> FILE DỮ LIỆU NGUỒN CHÍNH
                    </h3>
                    <p className="text-xs text-slate-500">Tải lên tệp dữ liệu chính (Excel/CSV) của bạn hoặc định nghĩa nhanh các cột chỉ định bên dưới.</p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
                    <label className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-indigo-900/10 flex items-center gap-2 cursor-pointer w-full md:w-auto justify-center">
                      <FileUp className="w-4 h-4" /> TẢI FILE DỮ LIỆU CHÍNH (EXCEL, CSV)
                      <input 
                        type="file" 
                        accept=".xlsx, .xls, .csv, .txt" 
                        onChange={(e) => handleFileUpload(e, "main")} 
                        className="hidden" 
                      />
                    </label>

                    {rawImportedData.length > 0 && (
                      <button
                        onClick={clearData}
                        className="bg-red-50 hover:bg-red-100 border border-red-200 text-red-600 hover:text-red-700 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer w-full md:w-auto"
                        title="Xóa dữ liệu giao dịch hiện tại (Danh mục Chuẩn VSIC giữ nguyên)"
                      >
                        <Trash2 className="w-4 h-4" /> XÓA DỮ LIỆU ĐÃ NẠP
                      </button>
                    )}
                  </div>
                </div>

                {/* HIỂN THỊ THÔNG TIN FILE ĐÃ NẠP */}
                {rawImportedData.length > 0 && (
                  <div className="bg-indigo-50/50 border border-indigo-100 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 animate-fade-in">
                    <div className="flex items-center gap-3 text-xs">
                      <div className="p-2 bg-white rounded-lg border border-indigo-150 shadow-sm text-indigo-600">
                        <FileCheck className="w-5 h-5" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="font-extrabold text-slate-800">Tệp đang nạp:</span>
                          <span className="font-mono text-indigo-700 font-bold bg-indigo-100/50 px-2 py-0.5 rounded border border-indigo-100 max-w-[280px] sm:max-w-md truncate" title={fileName}>
                            {fileName || "Dữ liệu nguồn chính"}
                          </span>
                        </div>
                        <p className="text-[11px] text-slate-500 mt-1 flex items-center gap-2">
                          <span>Quy mô: <b>{rawImportedData.length.toLocaleString()}</b> dòng dữ liệu</span>
                          <span className="text-slate-300">|</span>
                          <span>Đã phát hiện: <b>{columns.length}</b> cột tiêu đề ban đầu</span>
                          <span className="text-slate-300">|</span>
                          <span className="text-emerald-600 font-bold">✓ Danh mục VSIC &amp; AI Rules được giữ an toàn</span>
                        </p>
                      </div>
                    </div>
                  </div>
                )}

                {/* HIỂN THỊ CẤU HÌNH GHÉP CÁC SHEET KHI PHÁT HIỆN TỆP NHIỀU SHEET */}
                {detectedSheets.length > 1 && (
                  <div className="bg-amber-50/60 rounded-2xl p-5 border border-amber-200/80 space-y-4 animate-slide-up mt-4">
                    <div className="flex items-center gap-2 border-b border-amber-100 pb-3">
                      <div className="p-1.5 bg-amber-100 rounded-lg border border-amber-200">
                        <FileSpreadsheet className="w-5 h-5 text-amber-700" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono">
                          ⚡ Phát hiện File có nhiều Sheet ({detectedSheets.length} Sheets)
                        </h4>
                        <p className="text-[11px] text-amber-800">
                          Bạn có thể ghép (gộp) dữ liệu của nhiều Sheet này lại với nhau dựa trên một cột chung (ví dụ: Mã số thuế, Mã định danh, ID,...).
                        </p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-white p-4 rounded-xl border border-amber-200/50">
                      <div>
                        <div className="flex items-center justify-between mb-1.5">
                          <label className="text-[11px] font-bold text-slate-800 block font-mono">
                            1. CHỌN CÁC SHEET MUỐN GHÉP:
                          </label>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={() => setSelectedSheetsToMerge([...detectedSheets])}
                              className="text-[10px] text-amber-600 hover:text-amber-700 hover:underline cursor-pointer bg-transparent border-0 font-bold"
                            >
                              Chọn tất cả
                            </button>
                            <span className="text-slate-300 text-[10px]">|</span>
                            <button
                              type="button"
                              onClick={() => setSelectedSheetsToMerge([])}
                              className="text-[10px] text-slate-500 hover:text-slate-700 hover:underline cursor-pointer bg-transparent border-0 font-bold"
                            >
                              Bỏ chọn cả
                            </button>
                          </div>
                        </div>
                        <div className="max-h-[140px] overflow-y-auto space-y-1.5 p-2 bg-slate-50 rounded-lg border border-slate-200">
                          {detectedSheets.map(sheet => {
                            const isSelected = selectedSheetsToMerge.includes(sheet);
                            return (
                              <label key={sheet} className="flex items-center gap-2 text-xs text-slate-700 hover:text-slate-950 cursor-pointer select-none">
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
                                  className="rounded border-slate-300 bg-white text-amber-600 focus:ring-amber-500"
                                />
                                <span className={isSelected ? "text-amber-700 font-bold" : ""}>{sheet}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>

                      <div className="space-y-4 flex flex-col justify-between">
                        <div className="space-y-2">
                          <label className="text-[11px] font-bold text-slate-800 block font-mono">
                            2. CHỌN CỘT CHUNG (ID/MST) ĐỘNG:
                          </label>
                          <select
                            value={sheetMergeCommonKey}
                            onChange={(e) => setSheetMergeCommonKey(e.target.value)}
                            className="w-full bg-white border border-slate-250 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:border-amber-600 focus:ring-1 focus:ring-amber-600 font-mono shadow-inner font-bold"
                          >
                            <option value="">-- Chọn cột định danh dùng để gộp dòng --</option>
                            {columns.map(c => (
                              <option key={c} value={c}>
                                🔑 Cột: {c}
                              </option>
                            ))}
                          </select>
                          <p className="text-[10px] text-slate-500 italic">
                            Hệ thống sẽ đồng nhất, phối hợp các thông tin cột của dòng từ các Sheet dựa theo giá trị trùng khớp tại cột này.
                          </p>
                        </div>

                        <button
                          type="button"
                          onClick={handleMergeWorkbookSheets}
                          disabled={selectedSheetsToMerge.length < 2 || !sheetMergeCommonKey}
                          className="w-full bg-gradient-to-r from-amber-500 to-yellow-600 hover:from-amber-600 hover:to-yellow-700 disabled:from-slate-200 disabled:to-slate-300 disabled:text-slate-400 disabled:cursor-not-allowed text-white font-bold text-xs py-2.5 px-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer border-0"
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
                <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm">
                  <div className="flex flex-col lg:flex-row lg:items-center justify-between border-b border-slate-200 pb-4 gap-4">
                    <div>
                      <h4 className="text-base font-bold text-slate-800 flex items-center gap-2">
                        <BrainCircuit className="w-5.5 h-5.5 text-indigo-600 animate-pulse" /> 🧠 TRỢ LÝ AI HỌC LỆNH & QUẢN LÝ ÁNH XẠ CỘT THÔNG MINH
                      </h4>
                      <p className="text-xs text-slate-500 mt-1">
                        Dạy AI nhận dạng mẫu tiêu đề cột bằng giọng nói/văn bản tự nhiên, hoặc kích hoạt thư viện lệnh học định cấu hình mẫu một lần cho mọi file sau.
                      </p>
                    </div>

                    <div className="flex flex-wrap gap-2 items-center">
                      <label className="bg-slate-50 hover:bg-slate-100 text-slate-700 font-bold text-xs px-3.5 py-2 rounded-xl transition-all border border-slate-200 flex items-center gap-1.5 cursor-pointer shadow-sm">
                        <Upload className="w-3.5 h-3.5 text-indigo-600" /> Nhập lệnh học (.json)
                        <input
                          type="file"
                          accept=".json"
                          onChange={handleImportCommands}
                          className="hidden"
                        />
                      </label>
                      <button
                        onClick={handleExportCommands}
                        className="bg-slate-50 hover:bg-slate-100 text-slate-700 font-bold text-xs px-3.5 py-2 rounded-xl transition-all border border-slate-200 flex items-center gap-1.5 cursor-pointer shadow-sm"
                        title="Tải tệp JSON chứa toàn bộ lệnh học của bạn để backup hoặc chia sẻ"
                      >
                        <Download className="w-3.5 h-3.5 text-indigo-600" /> Xuất thư viện lệnh (.json)
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
                    {/* Cột 1: Huấn luyện AI */}
                    <div className="xl:col-span-7 space-y-4 bg-slate-50 p-4 rounded-xl border border-slate-200 shadow-sm">
                      <div>
                        <div className="flex items-center justify-between mb-1.5 flex-wrap gap-2">
                          <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                            🗣️ Nhập khẩu lệnh của bạn hoặc chọn các mẫu gợi ý:
                          </label>
                          <button
                            type="button"
                            onClick={toggleMicCol}
                            className={`flex items-center gap-1 text-[10px] font-extrabold px-2 py-0.5 rounded border transition-all cursor-pointer ${
                              isRecordingColMic 
                                ? "bg-rose-100 hover:bg-rose-200 text-rose-700 border-rose-300 animate-pulse font-bold" 
                                : "bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border-emerald-200"
                            }`}
                            title="Bấm để nói bằng tiếng Việt"
                          >
                            {isRecordingColMic ? (
                              <>
                                <MicOff className="w-3 h-3 text-rose-600 shrink-0" /> Dừng nghe
                              </>
                            ) : (
                              <>
                                <Mic className="w-3 h-3 text-emerald-600 shrink-0" /> Ghi âm (Nói)
                              </>
                            )}
                          </button>
                        </div>
                        <textarea
                          rows={3}
                          value={aiColLearnPrompt}
                          onChange={(e) => setAiColLearnPrompt(e.target.value)}
                          placeholder="Ví dụ: Đặt tên cho MST thành 'Mã Số Thuế' và gán vai trò idCol, cột DoanhThu mới tên là 'Doanh Thu 2024' vai trò doanhthu, loại bỏ các cột không dùng khác..."
                          className="w-full bg-white border border-slate-200 hover:border-indigo-400 rounded-xl px-3 py-2.5 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 placeholder-slate-400 leading-relaxed font-sans shadow-inner"
                        />
                      </div>

                      {/* Gợi ý Lệnh nhanh */}
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Đổi tên cột MaST thành Mã Số Thuế gán vai trò idCol, cột Xa thành Địa bàn Xã vai trò xa, MoTa thành Mô Tả Hoạt Động vai trò mota.")}
                          className="text-[10px] bg-indigo-50 hover:bg-indigo-100 text-indigo-700 px-2.5 py-1 rounded-lg border border-indigo-200/50 transition-all cursor-pointer font-bold"
                        >
                          📌 Cú pháp thuế chuẩn
                        </button>
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Chỉ giữ lại cột Mã Số Thuế và Mô tả hoạt động kinh doanh, loại bỏ tất cả các cột dư thừa khác ra khỏi file mới.")}
                          className="text-[10px] bg-indigo-50 hover:bg-indigo-100 text-indigo-700 px-2.5 py-1 rounded-lg border border-indigo-200/50 transition-all cursor-pointer font-bold"
                        >
                          📌 Rút gọn giữ MST & Mô tả
                        </button>
                        <button
                          type="button"
                          onClick={() => setAiColLearnPrompt("Việt hóa có dấu thật gọn cho mọi tiêu đề cột, gán đúng vai trò số cho DoanhThu và LaoDong.")}
                          className="text-[10px] bg-indigo-50 hover:bg-indigo-100 text-indigo-700 px-2.5 py-1 rounded-lg border border-indigo-200/50 transition-all cursor-pointer font-bold"
                        >
                          📌 Việt hóa gọn gàng tự động
                        </button>
                      </div>

                      <div className="flex gap-2">
                        <button
                          onClick={() => handleAiColLearn(true)}
                          disabled={isLearningColAi}
                          className={`flex-1 bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 disabled:opacity-80 disabled:from-slate-100 disabled:to-slate-200 disabled:text-slate-400 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-md shadow-indigo-900/10 relative overflow-hidden border-0 ${
                            isLearningColAi ? "animate-pulse" : ""
                          }`}
                        >
                          {isLearningColAi ? (
                            <>
                              <Loader2 className="w-4 h-4 text-indigo-400 animate-spin" />
                              <span>ĐANG HUẤN LUYỆN GEMINI...</span>
                            </>
                          ) : (
                            <>
                              <BrainCircuit className="w-4 h-4 text-white" />
                              <span>🧠 HUẤN LUYỆN QUA AI (GEMINI)</span>
                            </>
                          )}
                        </button>
                        
                        <button
                          onClick={() => handleAiColLearn(false)}
                          disabled={isLearningColAi}
                          className={`bg-slate-200 hover:bg-slate-300 text-slate-800 disabled:opacity-50 disabled:bg-slate-100 disabled:text-slate-400 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer border border-slate-300 shadow-sm ${
                            isLearningColAi ? "opacity-60" : ""
                          }`}
                          title="Học lệnh tức thì bằng bộ phân tích từ khóa tiếng Việt không cần API key"
                        >
                          {isLearningColAi ? (
                            <Loader2 className="w-4 h-4 text-slate-500 animate-spin" />
                          ) : (
                            <Zap className="w-4 h-4 text-amber-600" />
                          )}
                          <span>HỌC ĐỊNH DẠNG TRỰC TIẾP</span>
                        </button>
                      </div>

                      {/* Log học lệnh */}
                      {learningColLogs.length > 0 && (
                        <div className="bg-slate-950 rounded-lg p-3 border border-slate-850 max-h-[120px] overflow-y-auto space-y-1 font-mono text-[10px] text-green-400">
                          {learningColLogs.map((log, lidx) => (
                            <div key={lidx} className="leading-relaxed whitespace-pre-wrap">{log}</div>
                          ))}
                        </div>
                      )}

                      {/* Nút chạy áp dụng trực tiếp sau khi học */}
                      <div className="pt-3 border-t border-slate-200 space-y-2">
                        <button
                          onClick={handleApplyColumnRedefinition}
                          className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs px-4 py-3 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-emerald-950/10 border-0"
                          title="Áp dụng ngay các cột đã định nghĩa để đổi tên và lọc dữ liệu chính"
                        >
                          <FileCheck className="w-4 h-4 text-emerald-100" />
                          ⚡ CHẠY ÁP DỤNG LỆNH & TẠO FILE SẠCH NGAY LẬP TỨC
                        </button>
                        <p className="text-[10px] text-slate-500 text-center leading-relaxed">
                          (Nhấn nút này để thực thi việc đổi tên, khớp nối lọc cột và chuyển kết quả sang tab <span className="text-indigo-600 font-bold">Xem Dữ Liệu</span>)
                        </p>
                      </div>
                    </div>

                    {/* Cột 2: Thư viện Lệnh học đã tích lũy */}
                    <div className="xl:col-span-5 space-y-4 bg-slate-50 p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <div className="space-y-3">
                        <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                          <label className="text-xs font-bold text-slate-800 block uppercase font-mono">
                            🎓 THƯ VIỆN LỆNH HỌC TÍCH LŨY ({colLearnedCommands.length}):
                          </label>
                        </div>

                        <div className="space-y-2 max-h-[160px] overflow-y-auto pr-1">
                          {colLearnedCommands.map((cmd) => (
                            <div key={cmd.id} className="bg-white border border-slate-200 p-2.5 rounded-lg flex items-start justify-between gap-3 text-xs shadow-sm">
                              <div className="space-y-0.5">
                                <div className="font-bold text-slate-900">{cmd.name}</div>
                                <div className="text-[10px] text-slate-500 leading-relaxed text-wrap">{cmd.description}</div>
                              </div>
                              <div className="flex gap-1.5 flex-shrink-0">
                                <button
                                  type="button"
                                  onClick={() => applyLearnedCommand(cmd)}
                                  className="bg-indigo-50 hover:bg-indigo-100 text-[10px] text-indigo-700 font-bold px-2 py-1 rounded cursor-pointer transition-all border border-indigo-200/40"
                                  title="Áp dụng mẫu gán nhãn cột này lên bảng tính hiện thời"
                                >
                                  Áp dụng
                                </button>
                                {!cmd.id.startsWith("default-") && (
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteCommand(cmd.id)}
                                    className="text-[10px] hover:text-red-600 text-slate-400 font-bold px-1.5 py-1 rounded cursor-pointer transition-all"
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
                      <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-2 mt-2 shadow-sm">
                        <label className="text-[11px] font-bold text-slate-800 block uppercase font-mono">
                          💾 Lưu cấu hình bảng hiện tại thành lệnh học mới:
                        </label>
                        <div className="flex gap-2 font-sans">
                          <input
                            type="text"
                            value={newColCommandName}
                            onChange={(e) => setNewColCommandName(e.target.value)}
                            placeholder="Tên lệnh học, vd: Cấu hình bảng xã Tân Bình..."
                            className="flex-1 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono shadow-inner"
                          />
                          <button
                            type="button"
                            onClick={handleSaveCurrentAsCommand}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3.5 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1 border-0"
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
                  <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-5 shadow-sm">
                    
                    <div className="border-b border-slate-200 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <div className="text-xs font-bold text-indigo-700 tracking-wider uppercase font-mono flex items-center gap-1.5 cursor-pointer select-none" onClick={() => setIsConfigExpanded(!isConfigExpanded)}>
                          <Database className="w-5 h-5 text-indigo-600 animate-pulse" /> ĐỊNH NGHĨA LẠI TÊN CỘT DỄ NHỚ & LỌC CỘT THỪA {isConfigExpanded ? "▼" : "▲"}
                        </div>
                        <p className="text-xs text-slate-500 mt-1">
                          Sửa đổi các từ viết tắt khó nhớ thành tiếng Việt rõ ràng. Cột nào chưa chọn sẽ bị loại khỏi bảng để giữ bộ dữ liệu sạch nhất.
                        </p>
                      </div>
                      
                      <div className="flex gap-2 items-center">
                        <button
                          onClick={() => setIsConfigExpanded(!isConfigExpanded)}
                          className="bg-slate-50 hover:bg-slate-100 text-slate-700 font-bold text-[11px] px-3.5 py-1.5 rounded-lg border border-slate-200 cursor-pointer transition-all flex items-center gap-1 shadow-sm"
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
                              className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-slate-300 cursor-pointer shadow-sm"
                              title="Hoàn tác tất cả tên cột về tên gốc"
                            >
                              Khôi Phục Tên Gốc
                            </button>
                            <button 
                              onClick={() => {
                                const origNames = customColConfigs.map(c => c.originalName);
                                const uniqueRoles = getUniqueRoleAssignments(origNames);
                                
                                const prefilled = customColConfigs.map(cfg => {
                                  const beautified = beautifyColumnName(cfg.originalName);
                                  const recRole = uniqueRoles[cfg.originalName] || "";
                                  
                                  // Tinh chỉnh tên Việt hóa dựa trên vai trò hệ thống quy chuẩn nếu có
                                  let finalName = beautified;
                                  if (recRole === "idCol") finalName = "Mã Số Thuế";
                                  else if (recRole === "mota") finalName = "Mô Tả Hoạt Động";
                                  else if (recRole === "manganh") finalName = "Mã Ngành ĐK";
                                  else if (recRole === "xa") finalName = "Địa bàn (Xã)";
                                  else if (recRole === "doanhthu") finalName = "Doanh Thu";
                                  else if (recRole === "laodong") finalName = "Số Lao Động";
                                  
                                  return { 
                                    ...cfg, 
                                    newName: finalName,
                                    role: recRole as any
                                  };
                                });
                                setCustomColConfigs(prefilled);

                                // Đồng bộ hóa ngay lập tức vào state mapping hệ thống chính
                                setMapping(prev => {
                                  const next = { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" };
                                  prefilled.forEach(p => {
                                    if (p.role && p.role in next) {
                                      next[p.role as keyof typeof next] = p.originalName;
                                    }
                                  });
                                  return next;
                                });
                              }}
                              className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-indigo-200/50 cursor-pointer shadow-sm flex items-center gap-1 active:scale-95"
                              title="Tự động dịch nghĩa cột tiếng Việt và phân loại vai trò hệ thống duy nhất cho mỗi cột không trùng lặp"
                            >
                              <Brain className="w-3.5 h-3.5 text-indigo-600" /> Tự Động Đề Xuất Tên Việt Hóa & Vai Trò
                            </button>
                          </>
                        )}
                      </div>
                    </div>

                    {isConfigExpanded ? (
                      <>
                        {/* Hướng dẫn chi tiết */}
                        <div className="bg-indigo-50/50 border border-indigo-100 rounded-xl p-4 text-xs text-slate-750 space-y-1.5 leading-relaxed shadow-inner">
                          <div className="font-bold text-indigo-800 flex items-center gap-1.5">
                            ⚙️ Cách thức vận hành (Định nghĩa trực quan):
                          </div>
                          <ul className="list-disc list-inside space-y-1 text-[11px] text-slate-600 pl-1">
                            <li><strong>Đặt tên cột dễ nhớ:</strong> Viết trực tiếp vào ô nhập bên dưới để thay đổi tên cột hiển thị theo từ ngữ dễ thuộc của riêng bạn.</li>
                            <li><strong>Lọc cột thừa:</strong> Bạn có thể bỏ tích ở cột không cần thiết, khi bấm áp dụng hệ thống sẽ sinh ra một <strong>Bảng dữ liệu mới hoàn hảo</strong> chỉ chứa các cột thích hợp.</li>
                            <li><strong>Gán vai trò (Mục tiêu):</strong> Gán vai trò cho cột giúp các thuật toán (Báo cáo xã, nhóm ngành, xử lý lỗi logic bằng AI) tự động tìm đúng dữ liệu mà không bị đứt gãy.</li>
                          </ul>
                        </div>

                        {/* Bảng Danh sách Cấu hình Cột */}
                        <div className="overflow-x-auto max-h-[500px] overflow-y-auto border border-slate-200 rounded-xl bg-slate-50 shadow-inner relative scrollbar-thin">
                          <table className="w-full text-left text-xs min-w-[700px] border-separate border-spacing-0">
                            <thead>
                              <tr className="bg-slate-100 border-b border-slate-200 text-slate-600 font-mono text-[11px] sticky top-0 z-10">
                                <th className="p-3 text-center w-[120px] bg-slate-100 border-b border-slate-200 shadow-sm select-none sticky top-0 z-10">
                                  <div className="flex items-center justify-center gap-2">
                                    <input
                                      type="checkbox"
                                      checked={customColConfigs.length > 0 && customColConfigs.every(c => c.use && c.newName.trim() !== "")}
                                      ref={(el) => {
                                        if (el) {
                                          const someUse = customColConfigs.some(c => c.use && c.newName.trim() !== "");
                                          const allUse = customColConfigs.every(c => c.use && c.newName.trim() !== "");
                                          el.indeterminate = someUse && !allUse;
                                        }
                                      }}
                                      onChange={(e) => {
                                        const checked = e.target.checked;
                                        const updated = customColConfigs.map(c => ({
                                          ...c,
                                          use: checked,
                                          newName: checked ? (c.newName.trim() || c.originalName) : ""
                                        }));
                                        setCustomColConfigs(updated);
                                      }}
                                      className="w-3.5 h-3.5 rounded border-slate-300 text-purple-600 focus:ring-purple-500 cursor-pointer"
                                      title="Chọn tất cả / Bỏ chọn tất cả"
                                    />
                                    <span>SỬ DỤNG</span>
                                  </div>
                                </th>
                                <th className="p-3 text-center w-[50px] bg-slate-100 border-b border-slate-200 shadow-sm sticky top-0 z-10">STT</th>
                                <th className="p-3 bg-slate-100 border-b border-slate-200 shadow-sm sticky top-0 z-10">TÊN CỘT GỐC TRONG FILE (NHẤP ĐÚP ĐỂ CHỌN NHANH ⚡)</th>
                                <th className="p-3 bg-slate-100 border-b border-slate-200 shadow-sm sticky top-0 z-10">TÊN MỚI ĐỊNH NGHĨA (ĐỂ TRỐNG = LOẠI BỎ KHỎI FILE)</th>
                                <th className="p-3 w-[260px] bg-slate-100 border-b border-slate-200 shadow-sm sticky top-0 z-10">VAI TRÒ HỆ THỐNG (MỤC TIÊU)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-200 font-sans bg-white">
                              {customColConfigs.map((cfg, idx) => {
                                const isIncluded = cfg.newName.trim() !== "";
                                return (
                                  <tr 
                                    key={cfg.originalName} 
                                    className={`transition-colors hover:bg-slate-50 ${
                                      isIncluded ? "bg-purple-50/40" : "bg-slate-100 opacity-60"
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
                                        className="w-4 h-4 rounded border-slate-300 bg-white text-purple-600 focus:ring-purple-500 accent-purple-600 cursor-pointer"
                                      />
                                    </td>

                                    {/* STT */}
                                    <td className="p-3 text-center text-slate-550 font-mono text-[11px]">
                                      {idx + 1}
                                    </td>

                                    {/* Tên Gốc */}
                                    <td 
                                      className="p-3 font-semibold text-slate-800 font-mono cursor-pointer hover:text-purple-600 transition-all"
                                      title="Nhấn đúp vào đây để chọn nhanh giữ tên cột gốc làm định nghĩa!"
                                    >
                                      <div className="flex items-center gap-2">
                                        <span className="bg-slate-50 px-2.5 py-1 rounded text-slate-800 border border-slate-200 max-w-[250px] truncate block shadow-sm">
                                          {cfg.originalName}
                                        </span>
                                        <span className="text-[10px] text-slate-500 hover:text-purple-650 select-none">
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
                                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-purple-500 font-bold placeholder-slate-400 font-mono shadow-sm"
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
                                        className="w-full bg-white border border-slate-300 hover:border-purple-400 rounded-lg px-2 py-1.5 text-[11px] text-slate-700 focus:outline-none focus:ring-1 focus:ring-purple-500 font-bold font-sans cursor-pointer shadow-sm"
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
                            className="bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-250 font-bold text-[11px] px-3.5 py-2 rounded-xl transition-all cursor-pointer shadow-sm"
                          >
                            ❌ Xóa hết định nghĩa (Để trống tất cả)
                          </button>

                          <button
                            onClick={handleApplyColumnRedefinition}
                            className="bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-sm flex items-center gap-2 cursor-pointer border border-purple-500/20 hover:scale-[1.02] active:scale-[0.98]"
                          >
                            <FileCheck className="w-4 h-4" />⚡ XÁC NHẬN ĐỊNH NGHĨA & LỌC GỌN NHẸ TỔ TẠO FILE MỚI
                          </button>
                        </div>
                      </>
                    ) : (
                      <div className="flex items-center justify-between text-xs text-slate-500 bg-slate-50 p-3 rounded-xl border border-slate-150 shadow-inner">
                        <span>💡 Bảng cấu hình định nghĩa tên cột đang được thu gọn để nhường lại không gian biểu diễn danh sách dữ liệu.</span>
                        <button
                          onClick={() => setIsConfigExpanded(true)}
                          className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs px-3.5 py-1.5 rounded-lg border border-indigo-200 cursor-pointer transition-all shadow-sm"
                        >
                          ⚙️ Hiện bảng cấu hình
                        </button>
                      </div>
                    )}
                  </div>
                )}

              {/* PHÉP TÍNH CỘT VỚI CỘT (COLUMN FORMULA CALCULATOR) */}
              {rawImportedData.length > 0 && (
                <div className="bg-white rounded-2xl p-6 border border-slate-200 space-y-4 animate-slide-up shadow-sm">
                  <div className="border-b border-slate-100 pb-3">
                    <div className="text-xs font-bold text-indigo-900 tracking-wider uppercase font-sans flex items-center gap-1.5 font-medium">
                      <Database className="w-5 h-5 text-indigo-600 animate-pulse" /> 🧮 CÔNG CỤ TÌNH PHÉP TÍNH CỘT VỚI CỘT & TẠO CỘT MỚI
                    </div>
                    <p className="text-xs text-slate-500 mt-1">
                      Tính toán số liệu nâng cao trực tiếp trên bảng tính của bạn. Bạn có thể cộng, trừ, nhân, chia 2 cột với nhau hoặc tính toán với một hằng số cố định, hoặc ghép nội dung cột văn bản.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end bg-slate-50 p-4 rounded-xl border border-slate-200">
                    
                    {/* Tên cột kết quả */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-slate-700 font-bold text-xs block font-mono">
                        ✍️ 1. TÊN CỘT KẾT QUẢ MỚI:
                      </label>
                      <input
                        type="text"
                        value={calcColName}
                        onChange={(e) => setCalcColName(e.target.value)}
                        placeholder="VD: DoanhThuBinhQuan, Cong_X_Y"
                        className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono font-medium shadow-sm"
                      />
                    </div>

                    {/* Cột 1 */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-slate-700 font-bold text-xs block font-mono">
                        📂 2. CỘT THỨ NHẤT (A):
                      </label>
                      <select
                        value={calcCol1}
                        onChange={(e) => setCalcCol1(e.target.value)}
                        className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono shadow-sm"
                      >
                        <option value="">-- Chọn Cột A --</option>
                        {columns.map(c => (
                          <option key={c} value={c}>📊 {c}</option>
                        ))}
                      </select>
                    </div>

                    {/* Phép toán */}
                    <div className="md:col-span-2 space-y-1.5">
                      <label className="text-slate-700 font-bold text-xs block font-mono text-center">
                        ➕ PHÉP TOÁN:
                      </label>
                      <select
                        value={calcOperator}
                        onChange={(e) => setCalcOperator(e.target.value as any)}
                        className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 text-center focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono font-bold shadow-sm"
                      >
                        <option value="+">➕ Cộng (+)</option>
                        <option value="-">➖ Trừ (-)</option>
                        <option value="*">✖️ Nhân (*)</option>
                        <option value="/">➗ Chia (/)</option>
                        <option value="concat">🔗 Ghép chữ</option>
                      </select>
                    </div>

                    {/* Loại cột 2: Cột hay Hằng số */}
                    <div className="md:col-span-2 space-y-1.5">
                      <label className="text-slate-700 font-bold text-xs block font-mono text-center">
                        🎯 ĐỐI TƯỢNG B:
                      </label>
                      <select
                        value={calcType}
                        onChange={(e) => setCalcType(e.target.value as any)}
                        className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono shadow-sm"
                      >
                        <option value="column">Cột khác</option>
                        <option value="constant">Hằng số</option>
                      </select>
                    </div>

                    {/* Cột 2 hoặc Hằng số */}
                    <div className="md:col-span-2 space-y-1.5">
                      {calcType === "column" ? (
                        <>
                          <label className="text-slate-700 font-bold text-xs block font-mono">
                            📂 3. CỘT THỨ HAI (B):
                          </label>
                          <select
                            value={calcCol2}
                            onChange={(e) => setCalcCol2(e.target.value)}
                            className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono shadow-sm"
                          >
                            <option value="">-- Chọn Cột B --</option>
                            {columns.map(c => (
                              <option key={c} value={c}>📊 {c}</option>
                            ))}
                          </select>
                        </>
                      ) : (
                        <>
                          <label className="text-slate-700 font-bold text-xs block font-mono">
                            🔢 3. NHẬP GIÁ TRỊ HẰNG SỐ (B):
                          </label>
                          <input
                            type="text"
                            value={calcConstant}
                            onChange={(e) => setCalcConstant(e.target.value)}
                            placeholder="VD: 1000000, 1.2, Chuỗi chữ"
                            className="w-full bg-white border border-slate-300 hover:border-indigo-500/50 rounded-lg px-2.5 py-2 text-xs text-slate-850 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono shadow-sm"
                          />
                        </>
                      )}
                    </div>
                  </div>

                  {/* Lựa chọn làm tròn & nút xử lý */}
                  <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-1">
                    <div className="flex flex-wrap items-center gap-4 text-xs">
                      <span className="font-bold text-slate-600 font-mono">🎯 CHẾ ĐỘ LÀM TRÒN:</span>
                      <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 hover:text-slate-900 font-medium">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "none"}
                          onChange={() => setCalcRounding("none")}
                          className="text-indigo-600 focus:ring-indigo-500 border-slate-300 bg-white"
                        />
                        Không làm tròn
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 hover:text-slate-900 font-medium">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "int"}
                          onChange={() => setCalcRounding("int")}
                          className="text-indigo-600 focus:ring-indigo-500 border-slate-300 bg-white"
                        />
                        Số nguyên
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 hover:text-slate-900 font-medium">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "1dec"}
                          onChange={() => setCalcRounding("1dec")}
                          className="text-indigo-600 focus:ring-indigo-500 border-slate-300 bg-white"
                        />
                        1 chữ số thập phân
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 hover:text-slate-900 font-medium">
                        <input
                          type="radio"
                          name="calc_rounding"
                          checked={calcRounding === "2dec"}
                          onChange={() => setCalcRounding("2dec")}
                          className="text-indigo-600 focus:ring-indigo-500 border-slate-300 bg-white"
                        />
                        2 chữ số thập phân
                      </label>
                    </div>

                    <button
                      onClick={handleCalculateColumn}
                      className="bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-black text-xs px-6 py-3.5 rounded-xl transition-all shadow-md flex items-center gap-2 cursor-pointer border-b-4 border-indigo-700 active:scale-95 uppercase tracking-wide"
                    >
                      ⚡ THỰC HIỆN PHÉP TÍNH & THÊM CỘT
                    </button>
                  </div>
                </div>
              )}

              {mainData.length > 0 ? (
                <div className="space-y-4 font-sans animate-fade-in">
                  {rowFilterLabel && (
                    <div className="bg-amber-500/10 border border-amber-500/20 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-amber-800 bg-amber-55 animate-slide-up">
                      <div className="flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-amber-500 animate-pulse"></span>
                        <div>
                          <span>Bộ lọc đang được kích hoạt: </span>
                          <strong className="text-slate-800 bg-amber-100 px-2 py-0.5 rounded border border-amber-200 font-mono ml-1">{rowFilterLabel}</strong>
                        </div>
                      </div>
                      <button 
                        onClick={() => {
                          setRowIndicesFilter(null);
                          setRowFilterLabel(null);
                          setSearchTerm("");
                          setAiCommandResult(null);
                        }}
                        className="bg-amber-600 text-white hover:bg-amber-700 font-bold px-3 py-1.5 rounded-lg border border-amber-600 transition-all cursor-pointer text-[11px]"
                      >
                        Hủy lọc (Xem tất cả {mainData.length} dòng)
                      </button>
                    </div>
                  )}

                  <div className="w-full">
                    {/* Main Data Table */}
                    <MainDataInlinePreview 
                      data={filteredData}
                      columns={columns}
                      title="DỮ LIỆU NGUỒN CHÍNH HIỆN TẠI"
                      subtitle={rowFilterLabel ? `Đang hiển thị nhóm dữ liệu đã lọc (${filteredData.length} dòng).` : "Hệ thống hỗ trợ chọn cột/dòng bằng checkbox, bảng có sticky header cố định dòng tiêu đề."}
                      mapping={mapping}
                      onExportExcel={handleExportExcel}
                      enableSelection={true}
                      selectedColumns={selectedColumns}
                      onSelectedColumnsChange={setSelectedColumns}
                      selectedRows={selectedRows}
                      onSelectedRowsChange={setSelectedRows}
                    />
                  </div>
                </div>
              ) : (
                <div className="bg-white border-2 border-dashed border-slate-300 p-12 text-center rounded-2xl space-y-4 font-sans">
                  <Database className="w-12 h-12 text-slate-400 mx-auto animate-pulse" />
                  <div>
                    <h4 className="text-base font-bold text-slate-700">Chưa có cơ sở dữ liệu nạp vào</h4>
                    <p className="text-xs text-slate-500 max-w-md mx-auto pt-1 leading-relaxed">
                      Hãy chọn "Tải tệp dữ liệu chính" ở ô phía trên để nạp bảng tài liệu và kích hoạt toàn bộ cơ cấu.
                    </p>
                  </div>
                </div>
              )}
            </div>
          )}

          
          {/* 3. TAB GHÉP NỐI DỮ LIỆU */}
          {activeTab === "ghepnoi" && (
            <FileMerger
              mainData={mainData}
              fileName={fileName}
              setMainData={setMainData}
              setRawImportedData={setRawImportedData}
              setColumns={setColumns}
              setFileName={setFileName}
              setCustomColConfigs={setCustomColConfigs}
              setMapping={setMapping}
              setLoading={setLoading}
              setProgress={setProgress}
              setStatusMessage={setStatusMessage}
              onExportExcel={handleExportExcel}
            />
          )}

          {/* 4. TAB SO SÁNH ĐỐI CHIẾU HAI NIÊN ĐỘ (DIFF) */}
          {activeTab === "sosanh" && (
            <DataComparison
              mainData={mainData}
              fileName={fileName}
              setMainData={setMainData}
              setRawImportedData={setRawImportedData}
              setColumns={setColumns}
              setFileName={setFileName}
              setCustomColConfigs={setCustomColConfigs}
              setMapping={setMapping}
              setLoading={setLoading}
              setProgress={setProgress}
              setStatusMessage={setStatusMessage}
              onExportExcel={handleExportExcel}
            />
          )}

          {/* 5. TAB TÁCH DỮ LIỆU THEO CỘT */}
          {activeTab === "tachfile" && (
            <div className="space-y-6 animate-fade-in font-sans">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-4 text-slate-800">
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <Scissors className="w-5 h-5 text-pink-500" /> TÁCH FILE HÀNG LOẠT THEO CỘT CHỈ ĐỊNH
                </h3>
                <p className="text-xs text-slate-500">Chia nhỏ bảng tính lớn của bạn thành nhiều file Excel riêng biệt dựa trên giá trị cột đã chọn (ví dụ: tách theo từng Địa Phương Xã) và đóng gói tải xuống ZIP.</p>

                {mainData.length > 0 ? (
                  <div className="max-w-md space-y-4 bg-slate-50 rounded-xl p-5 border border-slate-100 shadow-sm">
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-700 block">Chọn cột để định nghĩa tách file</label>
                      <select 
                        value={splitCol} 
                        onChange={(e) => setSplitCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800"
                      >
                        <option value="">-- Chọn cột --</option>
                        {columns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <button 
                      onClick={handleSplitData}
                      className="w-full bg-gradient-to-r from-pink-500 to-pink-600 hover:from-pink-600 hover:to-pink-700 text-white font-bold text-xs px-6 py-3.5 rounded-xl border-b-4 border-pink-700 active:scale-95 transition-all shadow-md flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      <Scissors className="w-4 h-4" /> KHỞI CHẠY BẮT ĐẦU TÁCH HÀNG LOẠT & ZIP DOWNLOAD
                    </button>
                  </div>
                ) : (
                  <div className="bg-amber-50 rounded-xl p-6 text-center text-xs text-amber-800 border border-amber-200 font-sans">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước ở Tab "Xem & Định Nghĩa Cột"!
                  </div>
                )}
              </div>

              {activeTab === "tachfile" && mainData.length > 0 && (
                <MainDataInlinePreview 
                  data={mainData} 
                  columns={columns} 
                  title="DỮ LIỆU NGUỒN CHUẨN BỊ TÁCH FILE" 
                  subtitle="Xem nhanh danh sách dữ liệu chính sẽ được phân chia hệ thống."
                  mapping={mapping}
                />
              )}
            </div>
          )}

          {/* 6. TAB TỔNG HỢP BÁO CÁO ĐỘNG */}
          {activeTab === "tonghop" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-6 text-slate-850">
                <div>
                  <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-indigo-500" /> TỔNG HỢP GỘP NHÓM &amp; PHÁT TRIỂN BÁO CÁO LINH HOẠT
                  </h3>
                  <p className="text-xs text-slate-500 font-sans mt-0.5">
                    Hệ thống hạch toán đa năng không khóa cứng cột. Cho phép bạn gộp nhóm dữ liệu gốc theo địa bàn xã, phân cấp mã ngành VSIC (Cấp 1 &amp; Cấp 2) hoặc phân nhóm trực tiếp từ bất kỳ tiêu chí dữ liệu nông nghiệp, dân số, công nghiệp nào để làm các loại điều tra thống kê khác nhau.
                  </p>
                </div>

                {/* 1. KHU VỰC QUẢN LÝ LỆNH TỔNG HỢP (LƯU LỆNH / XUẤT LỆNH) */}
                <div className="bg-slate-50 p-5 rounded-xl border border-slate-100 space-y-4 shadow-sm">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-200">
                    <div>
                      <span className="text-xs font-bold text-indigo-600 tracking-wider uppercase font-mono block">
                        ⚡ HỆ THỐNG LỆNH CẤU HÌNH TỔNG HỢP (HỌC LỆNH VÀ DI CHUYỂN NHANH)
                      </span>
                      <span className="text-[10.5px] text-slate-500 font-sans mt-0.5 block">
                        Lưu lại và xuất các thiết lập chọn cột và phép toán để tái sử dụng ngay lập tức cho các tệp dữ liệu khác nhau.
                      </span>
                    </div>

                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        onClick={handleSaveTongHopCommand}
                        className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-sm animate-pulse"
                        title="Lưu tất cả thiết lập hiện tại thành một lệnh mới"
                      >
                        <Save className="w-3.5 h-3.5" /> Lưu lệnh hiện tại
                      </button>

                      <button
                        onClick={handleExportTongHopCommands}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
                        title="Xuất danh sách lệnh ra tệp tin cấu hình .json"
                      >
                        <Download className="w-3.5 h-3.5" /> Xuất lệnh (.json)
                      </button>

                      <label className="bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-bold text-xs px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-sm">
                        <Upload className="w-3.5 h-3.5 text-slate-500" /> Nhập lệnh
                        <input
                          type="file"
                          accept=".json"
                          onChange={handleImportTongHopCommands}
                          className="hidden"
                        />
                      </label>
                    </div>
                  </div>

                  {savedTongHopCommands.length === 0 ? (
                    <p className="text-[11px] text-slate-400 italic">Chưa có lệnh nào được lưu trong bộ nhớ trình duyệt.</p>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 max-h-[140px] overflow-y-auto pr-1">
                      {savedTongHopCommands.map((cmd) => (
                        <div
                          key={cmd.id}
                          onClick={() => handleApplyTongHopCommand(cmd)}
                          className="group bg-white hover:bg-amber-50/50 border border-slate-200 hover:border-amber-400 p-2.5 rounded-xl cursor-pointer transition-all flex items-center justify-between gap-2 text-xs shadow-sm"
                          title={`Click để tải cấu hình: ${cmd.name}`}
                        >
                          <div className="min-w-0 flex-1">
                            <p className="font-bold text-slate-700 group-hover:text-amber-600 transition-colors truncate">{cmd.name}</p>
                            <p className="text-[9.5px] text-slate-400 truncate mt-0.5">
                              Xã: {cmd.quickReportXaCol || "mặc định"} | Ngành: {cmd.quickReportManganhCol || "mặc định"}
                            </p>
                          </div>
                          <button
                            onClick={(e) => handleDeleteTongHopCommand(cmd.id, e)}
                            className="text-slate-400 hover:text-red-600 p-1 rounded hover:bg-red-50 transition-colors cursor-pointer shrink-0"
                            title="Xóa lệnh này"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* 2. KHU VỰC QUẢN LÝ ĐA TỆP TIN & NẠP THÊM FILE */}
                <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 space-y-4 shadow-sm text-slate-800">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-200">
                    <div>
                      <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                        📂 DANH SÁCH TỆP TIN DỮ LIỆU TỔNG HỢP &amp; NẠP THÊM FILE
                      </span>
                      <span className="text-[10.5px] text-slate-500 font-sans mt-0.5 block">
                        Nạp thêm các tệp tin Excel/CSV khác nhau của nhiều năm, nhiều loại điều tra để chạy tổng hợp gộp chung hoặc làm phép toán liên cột.
                      </span>
                    </div>

                    <div className="flex flex-wrap gap-2 items-center">
                      {aggregateFiles.length > 0 && (
                        <button
                          onClick={() => {
                            setCustomConfirmModal({
                              isOpen: true,
                              title: "Xác nhận xóa tệp nạp thêm",
                              message: "Bạn có chắc chắn muốn xóa toàn bộ các tệp tin nạp thêm không?",
                              note: "Tệp dữ liệu chính và danh mục ngành vẫn được giữ nguyên.",
                              confirmText: "XÓA TỆP NẠP THÊM",
                              onConfirm: () => {
                                setAggregateFiles([]);
                                setSelectedFileIdToAggregate("main_data_file");
                                setStatusMessage("Đã gỡ bỏ toàn bộ tệp tin nạp thêm thành công!");
                              }
                            });
                          }}
                          className="bg-red-50 hover:bg-red-100 border border-red-200 text-red-600 hover:text-red-700 font-bold text-xs px-3.5 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm active:scale-95"
                        >
                          <Trash2 className="w-3.5 h-3.5" /> XÓA FILE NẠP THÊM
                        </button>
                      )}

                      <label className="bg-gradient-to-r from-sky-500 to-sky-600 hover:from-sky-600 hover:to-sky-700 text-white font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-md border-b-4 border-sky-700 active:scale-95">
                        <Plus className="w-4 h-4 shrink-0" /> Nạp thêm tệp tin dữ liệu...
                        <input
                          type="file"
                          multiple
                          accept=".xlsx,.xls,.csv"
                          onChange={handleAggregateFileUpload}
                          className="hidden"
                        />
                      </label>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                    {/* Cột trái: Chọn tệp hiện tại để tổng hợp */}
                    <div className="lg:col-span-4 space-y-1.5">
                      <label className="text-xs font-bold text-slate-700 block">Tệp tin đang hoạt động (Để chạy tổng hợp):</label>
                      <select
                        value={selectedFileIdToAggregate}
                        onChange={(e) => {
                          const val = e.target.value;
                          setSelectedFileIdToAggregate(val);
                          const fileObj = allAvailableFiles.find(f => f.id === val);
                          if (fileObj) {
                            const cols = fileObj.columns;
                            const autoMng = cols.find(c => /mã\s*ngành|manganh|vsic|mã\s*nghe|manghe|ngành/i.test(c)) || "";
                            const autoXa = cols.find(c => /xã|phường|địa\s*bàn|dia_ban/i.test(c)) || "";
                            if (autoMng) setQuickReportManganhCol(autoMng);
                            if (autoXa) setQuickReportXaCol(autoXa);
                          }
                        }}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-xs text-slate-800 focus:ring-1 focus:ring-sky-500 font-sans"
                      >
                        {allAvailableFiles.map((f) => (
                          <option key={f.id} value={f.id}>
                            {f.id === "main_data_file" ? "📂 " : "📄 "} {f.name} ({f.data.length} dòng)
                          </option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 leading-relaxed mt-1">
                        Mặc định hệ thống sử dụng Tệp chính nạp từ trang đầu. Bạn có thể chuyển sang tệp phụ bất kỳ vừa nạp để cài đặt tiêu chí tổng hợp tương ứng.
                      </p>
                    </div>

                    {/* Cột phải: Danh sách tệp đang lưu trữ trong bộ nhớ tạm */}
                    <div className="lg:col-span-8">
                      <label className="text-xs font-bold text-slate-600 block mb-1">Tệp tin trong bộ nhớ tạm ({allAvailableFiles.length}):</label>
                      <div className="border border-slate-200 rounded-lg bg-white max-h-[140px] overflow-y-auto p-2 space-y-1.5 shadow-inner">
                        {allAvailableFiles.length === 0 ? (
                          <div className="text-[11px] text-slate-400 italic text-center py-4">Chưa có tệp tin nào được nạp. Hãy nạp tệp chính hoặc nạp thêm tệp phụ!</div>
                        ) : (
                          allAvailableFiles.map((file) => (
                            <div key={file.id} className="flex items-center justify-between bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200 hover:border-slate-300 text-xs text-slate-700">
                              <span className="truncate font-medium flex items-center gap-1.5 max-w-[80%]" title={file.name}>
                                <FileSpreadsheet className="w-3.5 h-3.5 text-sky-500 shrink-0" />
                                <span className="truncate text-slate-700">{file.name}</span>
                              </span>
                              <div className="flex items-center gap-2 shrink-0">
                                <span className="text-[10px] font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200 shrink-0">
                                  {file.data.length} dòng | {file.columns.length} cột
                                </span>
                                {file.id !== "main_data_file" && (
                                  <button
                                    onClick={() => {
                                      if (confirm(`Bạn có muốn gỡ bỏ tệp "${file.name}" khỏi bộ nhớ tổng hợp?`)) {
                                        setAggregateFiles(prev => prev.filter(f => f.id !== file.id));
                                        if (selectedFileIdToAggregate === file.id) {
                                          setSelectedFileIdToAggregate("main_data_file");
                                        }
                                      }
                                    }}
                                    className="text-slate-400 hover:text-red-500 cursor-pointer p-0.5 font-bold text-sm"
                                    title="Gỡ bỏ tệp"
                                  >
                                    ×
                                  </button>
                                )}
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {allAvailableFiles.length > 0 ? (
                  <div className="space-y-6">
                    {/* BỘ LỰA CHỌN CỘT THỦ CÔNG */}
                    <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 space-y-5 shadow-sm text-slate-800">
                      <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                        ⚙️ Cấu hình các cột phân nhóm cho tệp đang chọn
                      </span>
                      
                      <div className="bg-sky-50 border border-sky-100 px-3 py-2 rounded-lg text-xs text-sky-700 flex items-center gap-1.5">
                        <Zap className="w-4 h-4 text-sky-500 animate-pulse shrink-0" />
                        <span>Đang cấu hình cho tệp: <strong className="text-slate-950">{(allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || allAvailableFiles[0])?.name}</strong></span>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="text-xs font-semibold text-slate-700 block mb-1">Cột Mã Ngành Hoặc Phân Nhóm Chính:</label>
                          <select 
                            value={quickReportManganhCol} 
                            onChange={(e) => setQuickReportManganhCol(e.target.value)}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột chính phân nhóm --</option>
                            {(allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || allAvailableFiles[0])?.columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <p className="text-[10px] text-slate-500 mt-1 font-sans">
                            Chọn cột mã ngành để quy nạp lên cấp 1, cấp 2 (VSIC) hoặc cột đặc tính gốc để phân tích trực tiếp.
                          </p>
                        </div>

                        <div>
                          <label className="text-xs font-semibold text-slate-700 block mb-1">Cột Xã / Địa Bàn / Đơn vị:</label>
                          <select 
                            value={quickReportXaCol} 
                            onChange={(e) => setQuickReportXaCol(e.target.value)}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium font-sans"
                          >
                            <option value="">-- Click chọn cột xã/phường/địa bàn --</option>
                            {(allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || allAvailableFiles[0])?.columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <p className="text-[10px] text-slate-500 mt-1 font-sans">
                            Dữ liệu sẽ được gộp và hiển thị lũy kế chi tiết theo từng giá trị địa bàn này.
                          </p>
                        </div>
                      </div>

                      {/* KHU VỰC THIẾT LẬP CHỈ TIÊU CỘNG DỒN ĐỘNG - KHÔNG KHÓA CỨNG */}
                      <div className="border-t border-slate-200 pt-4 space-y-3">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                          <div>
                            <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                              📊 Chọn cột chỉ tiêu số để cộng dồn (Không hạn chế số lượng)
                            </span>
                            <span className="text-[10.5px] text-slate-500">
                              Chọn nhiều chỉ tiêu tùy thích để phần mềm thực hiện cộng tổng cho từng nhóm (Ví dụ: Doanh thu, Lao động, Sản lượng, Vốn...).
                            </span>
                          </div>
                          
                          <div className="flex items-center gap-1.5 self-start">
                            <button
                              type="button"
                              onClick={() => {
                                const currentFile = allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || allAvailableFiles[0];
                                if (!currentFile) return;
                                const firstRow = currentFile.data[0] || {};
                                const numericCols = currentFile.columns.filter(col => {
                                  const val = String(firstRow[col] || "");
                                  return val && !isNaN(parseFloat(val.replace(/[^0-9.\-]/g, "")));
                                });
                                const cleanNumerics = numericCols.filter(col => {
                                  const isManganhOrXa = col === quickReportManganhCol || col === quickReportXaCol;
                                  const isIdOrCode = /mã|mst|code|id|phone|đt|điện\s*thoại|tel|fax|stt|index|key|serial|no\./i.test(col);
                                  return !isManganhOrXa && !isIdOrCode;
                                });
                                setQuickReportSumCols(cleanNumerics.length > 0 ? cleanNumerics : currentFile.columns.slice(0, 5).filter(col => !/mã|mst|code|id|phone|đt|điện\s*thoại|tel|fax|stt|index|key|serial|no\./i.test(col)));
                              }}
                              className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded px-2 py-1 text-[10px] font-bold cursor-pointer transition-all active:scale-95"
                            >
                              ⚙️ Tự động chọn cột số
                            </button>
                            <button
                              type="button"
                              onClick={() => setQuickReportSumCols([])}
                              className="bg-slate-200 hover:bg-slate-300 text-slate-700 rounded px-2 py-1 text-[10px] font-bold cursor-pointer transition-all active:scale-95"
                            >
                              ❌ Xóa tất cả lựa chọn
                            </button>
                          </div>
                        </div>

                        {/* LIST CHECKBOX ĐỘNG ĐỂ LỰA CHỌN CỘT CHỈ TIÊU */}
                        <div className="bg-white border border-slate-200 rounded-xl p-3 max-h-[160px] overflow-y-auto grid grid-cols-2 md:grid-cols-4 gap-2 shadow-inner">
                          {(allAvailableFiles.find(f => f.id === selectedFileIdToAggregate) || allAvailableFiles[0])?.columns.map(col => {
                            const isChecked = quickReportSumCols.includes(col);
                            const isIdOrCode = /mã|mst|code|id|phone|đt|điện\s*thoại|tel|fax|stt|index|key|serial|no\./i.test(col);
                            return (
                              <label 
                                key={col} 
                                className={`flex items-center justify-between gap-1.5 p-1.5 rounded-lg border text-xs cursor-pointer select-none transition-all ${
                                  isChecked 
                                    ? "bg-indigo-50 border-indigo-300 text-indigo-800 font-semibold" 
                                    : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-900"
                                }`}
                              >
                                <div className="flex items-center gap-1.5 min-w-0">
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
                                    className="rounded text-indigo-600 focus:ring-indigo-500 bg-white border-slate-300 w-3.5 h-3.5 shrink-0"
                                  />
                                  <span className="truncate" title={col}>{col}</span>
                                </div>
                                {isIdOrCode && (
                                  <span className="text-[8px] bg-amber-100 text-amber-700 px-1 py-0.5 rounded font-bold font-sans shrink-0" title="Mã số/ID thường không phù hợp để tính tổng cộng">
                                    ⚠️ Mã/ID
                                  </span>
                                )}
                              </label>
                            );
                          })}
                        </div>

                        {quickReportSumCols.length > 0 && (
                          <div className="flex flex-wrap gap-1.5 items-center bg-slate-100 p-2 rounded-lg border border-slate-200">
                            <span className="text-[10px] text-slate-500 uppercase font-bold font-mono">Đang chọn ({quickReportSumCols.length}):</span>
                            {quickReportSumCols.map(col => (
                              <span key={col} className="bg-white text-indigo-700 border border-slate-200 px-2 py-0.5 rounded text-[10px] font-mono flex items-center gap-1 shadow-sm">
                                {col}
                                <span 
                                  onClick={() => setQuickReportSumCols(prev => prev.filter(c => c !== col))}
                                  className="hover:text-red-500 cursor-pointer text-xs leading-none font-bold ml-1"
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
                    <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 space-y-3 shadow-sm text-slate-800">
                      <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                        Cấu hình định dạng hạch toán đầu ra
                      </span>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <label className="flex items-start gap-3 text-xs text-slate-600 hover:text-slate-950 cursor-pointer select-none bg-white p-3 rounded-lg border border-slate-200 shadow-sm">
                          <input 
                            type="radio" 
                            name="quickReportFormatPivot"
                            checked={reportType === "pivot"} 
                            onChange={() => setReportType("pivot")}
                            className="mt-1 text-indigo-600 focus:ring-indigo-500 bg-white border-slate-300"
                          />
                          <div>
                            <div className="font-bold text-slate-800 font-sans">Bảng xoay ngang Pivot (Khuyên dùng)</div>
                            <div className="text-[10.5px] text-slate-500 mt-1 font-sans">
                              Mỗi xã địa bàn hiển thị thành một hàng ngang. Các nhóm phân loại và các chỉ tiêu được xoay thành các cột mở rộng liền kề song song hỗ trợ rà soát nhanh chóng.
                            </div>
                          </div>
                        </label>

                        <label className="flex items-start gap-3 text-xs text-slate-600 hover:text-slate-950 cursor-pointer select-none bg-white p-3 rounded-lg border border-slate-200 shadow-sm">
                          <input 
                            type="radio" 
                            name="quickReportFormatPivot"
                            checked={reportType === "flat"} 
                            onChange={() => setReportType("flat")}
                            className="mt-1 text-indigo-600 focus:ring-indigo-500 bg-white border-slate-300"
                          />
                          <div>
                            <div className="font-bold text-slate-800 font-sans">Bảng phẳng danh sách truyền thống</div>
                            <div className="text-[10.5px] text-slate-500 mt-1 font-sans">
                              Dạng danh mục phẳng chuẩn hóa. Mỗi dòng tương ứng một cặp địa bàn xã và nhóm phân loại với các chỉ tiêu cộng dồn xếp dọc.
                            </div>
                          </div>
                        </label>
                      </div>
                    </div>

                    {/* NÚT THỰC THI CHẠY TỔNG HỢP VỚI CẤU HÌNH PHÂN CẤP ĐỘNG */}
                    <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 space-y-4 shadow-sm text-slate-800">
                      <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                        CHẠY TỔNG HỢP BÁO CÁO ĐA NĂNG
                      </span>
                      
                      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 items-end">
                        <div>
                          <label className="text-xs font-semibold text-slate-700 block mb-1">Mức độ phân cấp gộp nhóm:</label>
                          <select
                            value={quickReportLevel}
                            onChange={(e) => setQuickReportLevel(Number(e.target.value))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-sans font-medium"
                          >
                            <option value={0}>📊 Gom nhóm trực tiếp bằng nội dung gốc trong cột (Không tra cứu VSIC)</option>
                            <option value={6}>💼 Tổng hợp theo Nhóm ngành chính (Công nghiệp, Xây dựng, Thương mại, Vận tải, Dịch vụ)</option>
                            <option value={1}>📈 Phân cấp Ngành Cấp 1 (VSIC - Chữ cái A-U)</option>
                            <option value={2}>📈 Phân cấp Ngành Cấp 2 (VSIC - 2 chữ số)</option>
                            <option value={3}>📈 Phân cấp Ngành Cấp 3 (VSIC - 3 chữ số)</option>
                            <option value={4}>📈 Phân cấp Ngành Cấp 4 (VSIC - 4 chữ số)</option>
                            <option value={5}>📈 Phân cấp Ngành Cấp 5 (VSIC - 5 chữ số)</option>
                          </select>
                        </div>

                        <button 
                          onClick={() => handleQuickReport(quickReportLevel)}
                          className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3 rounded-xl transition-all shadow-md cursor-pointer flex items-center justify-center gap-2 font-sans active:scale-95 border-b-4 border-emerald-800 h-[38px]"
                        >
                          ⚡ Chạy Tổng Hợp Phân Cấp Đã Chọn
                        </button>

                        <button 
                          onClick={() => handleQuickReport(6)}
                          className="w-full bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-700 hover:to-indigo-800 text-white font-black text-xs py-3 rounded-xl transition-all shadow-md hover:shadow-indigo-950/30 cursor-pointer flex items-center justify-center gap-2 font-sans active:scale-95 border-b-4 border-indigo-900 h-[38px]"
                        >
                          💼 Tổng hợp theo nhóm ngành chính
                        </button>
                      </div>
                    </div>

                    {/* BẢNG HIỂN THỊ KẾT QUẢ ĐẦU RA (ĐÃ ĐƯỢC DI CHUYỂN LÊN TRÊN PHÉP TÍNH PHỨC TẠP) */}
                    {activeTab === "tonghop" && quickReportResultRows.length > 0 && (
                      <div className="space-y-8 pt-4 animate-fade-in text-slate-800">
                        <BeautifulReportTable
                          rows={quickReportResultRows}
                          cols={quickReportResultCols}
                          level={quickReportLevel}
                          reportType={reportType}
                          onExport={handleExportQuickReport}
                        />
                      </div>
                    )}

                    {/* 3. KHU VỰC PHÉP TÍNH PHỨC TẠP CỘNG TRỪ NHÂN CHIA GIỮA CỘT CỦA CÁC FILE KHÁC NHAU (ĐƯA XUỐNG DƯỚI CÙNG CÁC BƯỚC) */}
                    <ComplexCalculations
                      mathFileAId={mathFileAId}
                      setMathFileAId={setMathFileAId}
                      mathFileBId={mathFileBId}
                      setMathFileBId={setMathFileBId}
                      mathKeyA={mathKeyA}
                      setMathKeyA={setMathKeyA}
                      mathKeyA2={mathKeyA2}
                      setMathKeyA2={setMathKeyA2}
                      mathKeyB={mathKeyB}
                      setMathKeyB={setMathKeyB}
                      mathKeyB2={mathKeyB2}
                      setMathKeyB2={setMathKeyB2}
                      mathColA={mathColA}
                      setMathColA={setMathColA}
                      mathColA2={mathColA2}
                      setMathColA2={setMathColA2}
                      mathColA3={mathColA3}
                      setMathColA3={setMathColA3}
                      mathColB={mathColB}
                      setMathColB={setMathColB}
                      mathColB2={mathColB2}
                      setMathColB2={setMathColB2}
                      mathColB3={mathColB3}
                      setMathColB3={setMathColB3}
                      mathOp={mathOp}
                      setMathOp={setMathOp}
                      mathOp2={mathOp2}
                      setMathOp2={setMathOp2}
                      mathOp3={mathOp3}
                      setMathOp3={setMathOp3}
                      mathNewColName={mathNewColName}
                      setMathNewColName={setMathNewColName}
                      mathNewColName2={mathNewColName2}
                      setMathNewColName2={setMathNewColName2}
                      mathNewColName3={mathNewColName3}
                      setMathNewColName3={setMathNewColName3}
                      mathFilterA={mathFilterA}
                      setMathFilterA={setMathFilterA}
                      mathFilterB={mathFilterB}
                      setMathFilterB={setMathFilterB}
                      allAvailableFiles={allAvailableFiles}
                      handlePerformCrossFileMath={handlePerformCrossFileMath}
                    />
                  </div>
                ) : (
                  <div className="bg-amber-50 rounded-xl p-6 text-center text-xs text-amber-800 border border-amber-200 font-sans">
                    ⚠️ Vui lòng nạp dữ liệu chính ở trang đầu tiên hoặc bấm nút nạp thêm tệp tin ở trên để tiến hành hạch toán tổng hợp.
                  </div>
                )}
              </div>
            </div>
          )}


          {/* 7. TAB KIỂM TRA & PHÂN TÍCH NGÀNH */}
          {activeTab === "chuanhoanganh" && (
            <div className="space-y-6 animate-fade-in font-sans">
              {renderAiMacroCognitiveCenter()}
              <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 animate-fade-in shadow-sm">
                <div>
                  <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                    <Brain className="w-5 h-5 text-indigo-600 animate-pulse" /> CÔNG CỤ TỰ ĐỘNG PHÂN TÍCH BẤT NHẤT QUÁN QUY LUẬT VSIC
                  </h3>
                  <p className="text-xs text-slate-500">Rà soát chéo thông minh: Quét toàn bộ tệp tin dữ liệu chính, phát hiện mâu thuẫn phân gán lỗi giữa Mô tả hoạt động thực tế và Bản mã ngành VSIC.</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="space-y-6 border-t border-slate-200 pt-6">
                    
                    {/* KHU VỰC THIẾT LẬP CHỌN 2 CỘT ĐỀ PHÂN TÍCH */}
                    <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 space-y-4">
                      <h4 className="text-xs font-bold text-slate-700 uppercase tracking-widest flex items-center gap-1.5 pb-2 border-b border-slate-200">
                        <Sliders className="w-4 h-4 text-indigo-500" /> THIẾT LẬP 2 CỘT RÀ SÁT CHÉO
                      </h4>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="space-y-1.5">
                          <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                            Chọn Cột chứa Mô tả/Tên ngành thực tế:
                          </label>
                          <select
                            value={stdDescriptionCol || mapping.mota || ""}
                            onChange={(e) => setStdDescriptionCol(e.target.value)}
                            className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2.5 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-sm font-bold"
                          >
                            <option value="">-- Chọn cột mô tả --</option>
                            {columns.map(col => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                          <p className="text-[10px] text-slate-500">Chọn cột text chứa nội dung chi tiết hoạt động kinh tế.</p>
                        </div>

                        <div className="space-y-1.5">
                          <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                            Chọn Cột chứa Mã Ngành VSIC cấp 5:
                          </label>
                          <select
                            value={stdIndustryCol || mapping.manganh || ""}
                            onChange={(e) => setStdIndustryCol(e.target.value)}
                            className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2.5 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-sm font-bold"
                          >
                            <option value="">-- Chọn cột mã ngành --</option>
                            {columns.map(col => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                          <p className="text-[10px] text-slate-500">Cột chứa chuỗi mã định dạng cấp 5 (hoặc các cấp tự liên hợp).</p>
                        </div>
                      </div>
                    </div>

                    {/* HAI PANEL KẾT QUẢ ĐỐI SÁNH SONG SONG */}
                    <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 pt-2">
                      
                      {/* BẢNG 1: CÙNG MÔ TẢ -> KHÁC MÃ NGÀNH */}
                      <div className="bg-slate-50 rounded-2xl p-5 border border-slate-200 space-y-4 shadow-sm">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="bg-indigo-50 border border-indigo-200 text-indigo-700 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Cùng Mô tả / Khác Mã</span>
                            <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                              📋 CÙNG MÔ TẢ ➔ LỆCH KHÁC MÃ VSIC ({inconAnalysis.descToCodes.length})
                            </h4>
                          </div>
                          <p className="text-[11px] text-slate-500 mt-1">
                            Tìm thấy các hộ khai nội dung kinh doanh giống hệt nhau nhưng bị cán bộ nhập / phân gán lệch sang các mã nghề khác nhau.
                          </p>
                          {inconAnalysis.descToCodes.length > 0 && (
                            <div className="flex flex-wrap gap-2 pt-2">
                              <button
                                onClick={() => {
                                  const allIndices = inconAnalysis.descToCodes.flatMap(item => item.codes.flatMap(c => c.rows));
                                  if (allIndices.length === 0) {
                                    alert("Không có dòng mâu thuẫn nào để lọc!");
                                    return;
                                  }
                                  setRowIndicesFilter(allIndices);
                                  setRowFilterLabel("Tất cả dòng mâu thuẫn (Cùng mô tả nhưng khác mã ngành)");
                                  setViewPage(1);
                                  setIsConfigExpanded(false);
                                  setActiveTab("xemdulieu");
                                }}
                                className="bg-indigo-550/10 hover:bg-indigo-600 text-indigo-700 hover:text-white px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer border border-indigo-200 flex items-center gap-1.5 active:scale-95 shadow-sm"
                                title="Lọc tất cả các dòng có mô tả trùng nhau nhưng bị gán mã khác nhau để xem tập trung"
                              >
                                <Search className="w-3.5 h-3.5 text-indigo-500" />
                                Lọc tổng {inconAnalysis.descToCodes.reduce((acc, item) => acc + item.occurrences, 0)} dòng mâu thuẫn
                              </button>
                              <button
                                onClick={() => handleExportInconsistentExcel(inconAnalysis.descToCodes)}
                                className="bg-emerald-50 hover:bg-emerald-600 text-emerald-700 hover:text-white px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer border border-emerald-200 flex items-center gap-1.5 active:scale-95 shadow-sm"
                                title="Xuất excel chi tiết danh sách mâu thuẫn cùng mô tả khác mã ngành"
                              >
                                <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                                Xuất Excel mâu thuẫn chi tiết
                              </button>
                            </div>
                          )}
                        </div>

                        <div className="bg-white border border-slate-200 rounded-xl p-4 min-h-[250px] max-h-[450px] overflow-y-auto space-y-3 shadow-inner">
                          {inconAnalysis.descToCodes.length === 0 ? (
                            <div className="text-xs text-emerald-600 flex items-center justify-center h-44 gap-1.5 font-mono">
                              <CheckCircle2 className="w-4 h-4 text-emerald-500 animate-pulse" /> Tuyệt vời! Không phát hiện mâu thuẫn "Cùng một mô tả gán khác mã".
                            </div>
                          ) : (
                            <div className="space-y-3 divide-y divide-slate-100">
                              {inconAnalysis.descToCodes.slice(0, visibleDescInconCount).map((item, idx) => (
                                <div key={idx} className="pt-3 first:pt-0 flex flex-col justify-between gap-2.5 text-xs">
                                  <div className="space-y-1.5 flex-1">
                                    <div className="font-bold text-slate-800 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-200 leading-relaxed font-mono">
                                      📝 "{item.motaText}" <span className="text-slate-400 text-[10px] ml-1 font-normal font-sans">({item.occurrences} dòng giống nhau)</span>
                                    </div>
                                    <div className="flex flex-wrap gap-1.5 items-center text-[11px]">
                                      <span className="text-slate-500 font-medium">Được phân gán xen kẽ (Click để lọc riêng):</span>
                                      {(() => {
                                        const maxCount = Math.max(...item.codes.map(c => c.count));
                                        return item.codes.map((c, cidx) => {
                                          const isMinority = c.count < maxCount;
                                          return (
                                            <button
                                              key={cidx}
                                              onClick={() => {
                                                setRowIndicesFilter(c.rows);
                                                setRowFilterLabel(`Mô tả: "${item.motaText}" ➔ Mã: ${c.code}`);
                                                setViewPage(1);
                                                setIsConfigExpanded(false);
                                                setActiveTab("xemdulieu");
                                              }}
                                              className={`font-mono flex items-center gap-1 transition-all cursor-pointer text-[10px] px-2.5 py-0.5 rounded border ${
                                                isMinority 
                                                  ? "bg-amber-50 hover:bg-amber-100 text-amber-700 hover:text-amber-800 border-amber-300" 
                                                  : "bg-indigo-50 hover:bg-indigo-100 text-indigo-700 hover:text-indigo-800 border-indigo-250"
                                              }`}
                                              title={isMinority ? `Mã gán thiểu số ít hơn (${c.count} dòng < Đa số ${maxCount} dòng). Click để lọc riêng.` : `Mã gán đa số phổ biến nhất (${c.count} dòng). Click để lọc riêng.`}
                                            >
                                              {isMinority && <span className="text-[10px]" title="Mã thiểu số gán ít nhất - nghi ngờ sai">⚠️</span>}
                                              <strong>{c.code}</strong> 
                                              <span className="text-[9px] opacity-70">({c.count} d)</span>
                                            </button>
                                          );
                                        });
                                      })()}
                                    </div>
                                  </div>
                                  <button
                                    onClick={() => {
                                      const allIndices = item.codes.flatMap(c => c.rows);
                                      setRowIndicesFilter(allIndices);
                                      setRowFilterLabel(`Cùng mô tả: "${item.motaText}"`);
                                      setViewPage(1);
                                      setIsConfigExpanded(false);
                                      setActiveTab("xemdulieu");
                                    }}
                                    className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 px-2.5 py-1 rounded text-[10px] font-bold self-start cursor-pointer transition-all flex items-center gap-1 active:scale-95 shrink-0"
                                    title="Lọc tất cả các dòng của nhóm mô tả này"
                                  >
                                    <Search className="w-3 h-3 text-indigo-600" /> Lọc tổng {item.occurrences} dòng gốc
                                  </button>
                                </div>
                              ))}

                              {inconAnalysis.descToCodes.length > visibleDescInconCount && (
                                <div className="pt-3 pb-1 text-center">
                                  <button
                                    onClick={() => setVisibleDescInconCount(prev => prev + 100)}
                                    className="w-full bg-slate-100 hover:bg-slate-200 text-indigo-600 border border-slate-300 rounded-lg py-2 text-xs font-bold font-sans cursor-pointer transition-all"
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
                      <div className="bg-slate-50 rounded-2xl p-5 border border-slate-200 space-y-4 shadow-sm">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="bg-indigo-50 border border-indigo-200 text-indigo-700 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Cùng Mã / Khác Mô tả</span>
                            <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                              🏷️ GOM CÙNG MÃ VSIC ➔ KHÁC BIỆT MÔ TẢ CHỮ ({inconAnalysis.codeToDescs.length})
                            </h4>
                          </div>
                          <p className="text-[11px] text-slate-500 mt-1 font-sans">
                            Gom nhóm các dòng ghi nhận cùng mã VSIC nhưng nội dung văn bản mô tả hoạt động thực tế lại khác nhau.
                          </p>
                        </div>

                        <div className="bg-white border border-slate-200 rounded-xl p-4 min-h-[250px] max-h-[450px] overflow-y-auto space-y-3 shadow-inner">
                          {inconAnalysis.codeToDescs.length === 0 ? (
                            <div className="text-xs text-indigo-600 flex items-center justify-center h-44 gap-1.5 font-mono">
                              <CheckCircle2 className="w-4 h-4 text-indigo-500 animate-pulse" /> Tuyệt vời! Không phát hiện trường hợp đồng mã lệch chữ.
                            </div>
                          ) : (
                            <div className="space-y-3 divide-y divide-slate-100">
                              {inconAnalysis.codeToDescs.slice(0, visibleCodeInconCount).map((item, idx) => (
                                <div key={idx} className="pt-3 first:pt-0 flex flex-col justify-between gap-2.5 text-xs">
                                  <div className="space-y-1.5 flex-1">
                                    <div className="font-bold text-slate-800 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-200 leading-relaxed font-mono flex items-center justify-between">
                                      <span>🏷️ Mã VSIC: <strong className="text-slate-900 text-sm shrink-0 underline decoration-indigo-500/50 font-bold">{item.codeValue}</strong></span>
                                      <span className="text-slate-400 text-[10px] font-normal font-sans">({item.occurrences} dòng)</span>
                                    </div>
                                    <div className="space-y-1">
                                      <div className="text-[10px] text-slate-500 uppercase font-bold font-sans">Văn bản chi tiết lệch biệt (Click để lọc riêng):</div>
                                      <div className="max-h-[140px] overflow-y-auto pl-2 space-y-1 border-l-2 border-indigo-500/30">
                                        {item.descriptions.map((descObj, descIdx) => (
                                          <button
                                            key={descIdx}
                                            onClick={() => {
                                              setRowIndicesFilter(descObj.rows);
                                              setRowFilterLabel(`Mã VSIC: ${item.codeValue} ➔ Mô tả: "${descObj.desc}"`);
                                              setViewPage(1);
                                              setIsConfigExpanded(false);
                                              setActiveTab("xemdulieu");
                                            }}
                                            className="text-slate-700 bg-slate-50 hover:bg-indigo-50 hover:text-indigo-800 rounded p-1.5 flex items-start justify-between gap-2 border border-slate-200 text-[11px] w-full text-left transition-all cursor-pointer animate-fade-in"
                                            title={`Click để lọc riêng ${descObj.count} dòng mang mô tả này`}
                                          >
                                            <span className="italic font-medium">"{descObj.desc}"</span>
                                            <span className="text-[9px] shrink-0 text-indigo-700 font-mono bg-indigo-50 px-1 py-0.5 rounded border border-indigo-100 font-bold">
                                              ({descObj.count} dòng)
                                            </span>
                                          </button>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                  <button
                                    onClick={() => {
                                      const allIndices = item.descriptions.flatMap(d => d.rows);
                                      setRowIndicesFilter(allIndices);
                                      setRowFilterLabel(`Cùng mã VSIC: ${item.codeValue}`);
                                      setViewPage(1);
                                      setIsConfigExpanded(false);
                                      setActiveTab("xemdulieu");
                                    }}
                                    className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 px-2.5 py-1 rounded text-[10px] font-bold self-start cursor-pointer transition-all flex items-center gap-1 active:scale-95 shrink-0"
                                    title="Lọc tất cả các dòng của nhóm mã này"
                                  >
                                    <Search className="w-3 h-3 text-indigo-600" /> Lọc tổng {item.occurrences} dòng gốc
                                  </button>
                                </div>
                              ))}

                              {inconAnalysis.codeToDescs.length > visibleCodeInconCount && (
                                <div className="pt-3 pb-1 text-center">
                                  <button
                                    onClick={() => setVisibleCodeInconCount(prev => prev + 100)}
                                    className="w-full bg-slate-100 hover:bg-slate-200 text-indigo-600 border border-slate-300 rounded-lg py-2 text-xs font-bold font-sans cursor-pointer transition-all"
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
                  <div className="bg-slate-50 rounded-xl p-6 text-center text-xs text-amber-700 border border-slate-200">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>

              {mainData.length > 0 && (
                <div className="space-y-4">
                  {rowFilterLabel && (
                    <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-amber-800 animate-slide-up">
                      <div className="flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-amber-500 animate-pulse"></span>
                        <div>
                          <span>Hệ thống đang hiển thị chính xác các dòng của nhóm bất nhất: </span>
                          <strong className="text-white bg-amber-950/40 px-2 py-0.5 rounded border border-amber-800/30 font-mono ml-1">{rowFilterLabel}</strong>
                        </div>
                      </div>
                      <button 
                        onClick={() => {
                          setRowIndicesFilter(null);
                          setRowFilterLabel(null);
                          setSearchTerm("");
                        }}
                        className="bg-amber-500/25 hover:bg-amber-500/35 text-amber-200 hover:text-white font-bold px-3 py-1.5 rounded-lg border border-amber-500/30 transition-all cursor-pointer text-[11px]"
                      >
                        Hủy lọc (Xem tất cả {mainData.length} dòng)
                      </button>
                    </div>
                  )}
                  <MainDataInlinePreview 
                    data={augmentedFilteredData} 
                    columns={viewerColumns} 
                    title="BẢNG DỮ LIỆU CHUẨN HÓA VSIC" 
                    subtitle={rowFilterLabel ? `Đang hiển thị nhóm dữ liệu đã lọc (${filteredData.length} dòng).` : "Xem nhanh danh mục dữ liệu chính sau khi đã chuẩn hóa và khai bổ sung các cấp ngành VSIC."}
                    mapping={mapping}
                    onExportExcel={handleExportExcel}
                  />
                </div>
              )}
            </div>
          )}

          
          {/* 8. TAB KIỂM TRA LOGIC ĐA ĐIỀU KIỆN */}
          {activeTab === "kiemtralogic" && (
            <LogicChecking
              mainData={mainData}
              columns={columns}
              setMainData={setMainData}
              setColumns={setColumns}
              fileName={fileName}
              setFileName={setFileName}
              mapping={mapping}
              onExportExcel={handleExportExcel}
              saveAppState={saveAppState}
              setActiveTab={setActiveTab}
              setLoading={setLoading}
              setProgress={setProgress}
              setStatusMessage={setStatusMessage}
              rawImportedData={rawImportedData}
              customColConfigs={customColConfigs}
              quickReportManganhCol={quickReportManganhCol}
              setQuickReportManganhCol={setQuickReportManganhCol}
              quickReportXaCol={quickReportXaCol}
              setQuickReportXaCol={setQuickReportXaCol}
              quickReportDoanhThuCol={quickReportDoanhThuCol}
              setQuickReportDoanhThuCol={setQuickReportDoanhThuCol}
              quickReportLaoDongCol={quickReportLaoDongCol}
              setQuickReportLaoDongCol={setQuickReportLaoDongCol}
              handleQuickReport={handleQuickReport}
            />
          )}

          {/* 9. TAB CHỌN MẪU KHẢO SÁT CHUYÊN ĐỀ */}
          {activeTab === "chonmau" && (
            <SamplingSelection
              mainData={mainData}
              columns={columns}
              mapping={mapping}
              setLoading={setLoading}
              setStatusMessage={setStatusMessage}
              sampCorpData={sampCorpData}
              setSampCorpData={setSampCorpData}
              sampCorpFileName={sampCorpFileName}
              setSampCorpFileName={setSampCorpFileName}
              sampIndData={sampIndData}
              setSampIndData={setSampIndData}
              sampIndFileName={sampIndFileName}
              setSampIndFileName={setSampIndFileName}
            />
          )}

          {/* TAB PHÂN TÍCH TẦN SUẤT */}
          {activeTab === "tansuat" && (
            <FrequencyAnalysis mainData={mainData} columns={columns} />
          )}

          {/* TAB PHÂN TÍCH TƯƠNG QUAN */}
          {activeTab === "tuongquan" && (
            <CorrelationAnalysis 
              mainData={mainData} 
              columns={columns} 
              setRowIndicesFilter={setRowIndicesFilter} 
              setRowFilterLabel={setRowFilterLabel} 
              setViewPage={setViewPage} 
              setActiveTab={setActiveTab} 
            />
          )}

          {/* 10. TAB TRA CỨU DANH MỤC NGÀNH VSIC CHUẨN */}
          {activeTab === "danhmucvsic" && (
            <div className="space-y-6 animate-fade-in">
              <VsicCatalogExplorer />
            </div>
          )}

          {/* 11. TAB ĐỌC PDF & CHUYỂN SANG WORD */}
          {activeTab === "pdf2word" && (
            <div className="space-y-6 animate-fade-in">
              <PdfToWord />
            </div>
          )}

          {/* KHỐI 2: BỘ QUÉT LỆCH QUY LUẬT PHÂN PHỐI */}
          {activeTab === "outliers" && (
            <div className="space-y-6 animate-fade-in">
              <StatisticalOutliers 
                mainData={mainData} 
                columns={columns} 
                mapping={mapping} 
                onFilterRows={(indices) => {
                  setRowIndicesFilter(indices);
                  setActiveTab("xemdulieu");
                }}
                onUpdateMainData={(newData) => {
                  setMainData(newData);
                  saveAppState({
                    mainData: newData,
                    rawImportedData,
                    columns,
                    fileName,
                    mapping,
                    customColConfigs
                  }, true);
                }}
              />
            </div>
          )}



          {/* KHỐI 4: TRUNG TÂM QUẢN TRỊ QUY TẮC LOGIC */}
          {activeTab === "rulesstudio" && (
            <div className="space-y-6 animate-fade-in">
              <RulesStudio 
                mainData={mainData} 
                columns={columns} 
                mapping={mapping} 
                onFilterRows={(indices) => {
                  setRowIndicesFilter(indices);
                  setActiveTab("xemdulieu");
                }}
                onUpdateMainData={(newData) => {
                  setMainData(newData);
                  saveAppState({
                    mainData: newData,
                    rawImportedData,
                    columns,
                    fileName,
                    mapping,
                    customColConfigs
                  }, true);
                }}
              />
            </div>
          )}

          {/* HỆ THỐNG CỘNG TÁC CLOUD - TỰ ĐỘNG ĐƯỢC CHÈN VÀO ROUTER */}
          {activeTab === "dataentry" && (
            <div className="space-y-6 animate-fade-in">
              <DataEntry />
            </div>
          )}

          {activeTab === "videoroom" && (
            <div className="space-y-6 animate-fade-in">
              <VideoRoom />
            </div>
          )}

          {activeTab === "admindashboard" && (
            <div className="space-y-6 animate-fade-in">
              <AdminDashboard />
            </div>
          )}

        </main>
      </div>

      {/* CUSTOM CONFIRMATION MODAL - Tránh lỗi window.confirm bị chặn trong Iframe */}
      {customConfirmModal && customConfirmModal.isOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-950/70 backdrop-blur-sm p-4 animate-fade-in" id="custom-app-confirm-modal">
          <div className="bg-white rounded-2xl max-w-md w-full shadow-2xl border border-rose-100 overflow-hidden animate-scale-in">
            <div className="p-6 text-center space-y-4">
              <div className="w-16 h-16 bg-rose-50 text-rose-600 rounded-full flex items-center justify-center mx-auto border border-rose-100 shadow-inner">
                <Trash2 className="w-8 h-8 animate-pulse" />
              </div>
              <div className="space-y-1.5">
                <h3 className="text-base font-black text-slate-800 uppercase tracking-tight">{customConfirmModal.title}</h3>
                <p className="text-xs text-slate-500 leading-relaxed">
                  {customConfirmModal.message}
                </p>
                {customConfirmModal.note && (
                  <p className="text-[11px] bg-slate-50 border border-slate-100 p-2.5 rounded-xl text-slate-600 leading-relaxed">
                    ⚠️ <b>LƯU Ý:</b> {customConfirmModal.note}
                  </p>
                )}
              </div>
            </div>
            <div className="bg-slate-50 px-6 py-4 flex items-center justify-end gap-2.5 border-t border-slate-100">
              <button
                onClick={() => setCustomConfirmModal(null)}
                className="px-4 py-2 text-xs font-bold text-slate-600 hover:text-slate-800 hover:bg-slate-100 rounded-xl transition-all cursor-pointer"
              >
                HỦY BỎ
              </button>
              <button
                onClick={() => {
                  customConfirmModal.onConfirm();
                  setCustomConfirmModal(null);
                }}
                className="px-5 py-2.5 text-xs font-black text-white bg-rose-600 hover:bg-rose-700 shadow-md shadow-rose-900/15 hover:shadow-lg rounded-xl transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              >
                <Trash2 className="w-3.5 h-3.5" /> {customConfirmModal.confirmText || "XÁC NHẬN"}
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

function AppWrapper() {
  const { user, loading } = useAuth();

  if (loading) {
    return (
      <div className="min-h-screen bg-slate-900 flex flex-col justify-center items-center p-4 font-sans text-slate-100 animate-fade-in">
        <div className="text-center space-y-3">
          <div className="w-10 h-10 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin mx-auto" />
          <p className="text-xs text-slate-400 font-semibold">Đang liên kết cổng liên ngành...</p>
        </div>
      </div>
    );
  }

  if (!user) {
    return <AuthScreen />;
  }

  return <MainAppContent />;
}

// KHỞI CHẠY ỨNG DỤNG ĐƯỢC BAO BỌC BỞI AUTHENTICATION PROVIDER
export default function App() {
  return (
    <AuthProvider>
      <AppWrapper />
    </AuthProvider>
  );
}
