import React, { useState, useMemo, useEffect } from "react";
import {
  RefreshCw,
  FileUp,
  Calendar,
  TrendingUp,
  HelpCircle,
  ArrowRight,
  Download,
  BarChart2,
  CheckCircle2,
  Database,
  Trash2,
  Save,
  Sparkles,
  FolderOpen,
  AlertCircle,
  Plus
} from "lucide-react";
import * as XLSX from "xlsx";
import { parseCSV } from "../utils/sharedHelpers";
import { MainDataInlinePreview } from "./MainDataInlinePreview";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend
} from "recharts";

interface DataComparisonProps {
  mainData: any[];
  fileName: string;
  setMainData: (data: any[]) => void;
  setRawImportedData: (data: any[]) => void;
  setColumns: (cols: string[]) => void;
  setFileName: (name: string) => void;
  setCustomColConfigs: (configs: any[]) => void;
  setMapping: (mapping: any) => void;
  setLoading: (loading: boolean) => void;
  setProgress: (progress: number) => void;
  setStatusMessage: (msg: string) => void;
  onExportExcel: () => void;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// --- LONG-TERM INDEXEDDB DATABASE FOR SAVED SNAPSHOTS ---
export interface SavedSnapshot {
  id: string;
  name: string;
  month?: number;
  year: number;
  cycleType?: "month" | "quarter" | "year";
  cycleValue?: number;
  category?: string;
  fileName: string;
  data: any[];
  columns: string[];
  createdAt: string;
}

export function getFormattedPeriod(snap: SavedSnapshot): string {
  const type = snap.cycleType || "month";
  const val = snap.cycleValue ?? snap.month ?? 1;
  if (type === "quarter") return `Quý ${val}/${snap.year}`;
  if (type === "year") return `Năm ${snap.year}`;
  return `Tháng ${val}/${snap.year}`;
}

export function openSnapshotsDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    try {
      if (typeof window === "undefined" || !window.indexedDB) {
        reject(new Error("IndexedDB is not supported in this environment"));
        return;
      }
      const request = indexedDB.open("VTongSnapshotsDatabase", 1);
      request.onerror = () => reject(request.error || new Error("Failed to open snapshots DB"));
      request.onsuccess = () => resolve(request.result);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains("snapshots")) {
          db.createObjectStore("snapshots", { keyPath: "id" });
        }
      };
    } catch (e) {
      reject(e);
    }
  });
}

export function saveSnapshotToDB(snapshot: SavedSnapshot): Promise<void> {
  return openSnapshotsDB().then((db) => {
    return new Promise<void>((resolve, reject) => {
      const transaction = db.transaction("snapshots", "readwrite");
      const store = transaction.objectStore("snapshots");
      const request = store.put(snapshot);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error("Failed to save snapshot"));
    });
  });
}

export function getAllSnapshotsFromDB(): Promise<SavedSnapshot[]> {
  return openSnapshotsDB().then((db) => {
    return new Promise<SavedSnapshot[]>((resolve, reject) => {
      const transaction = db.transaction("snapshots", "readonly");
      const store = transaction.objectStore("snapshots");
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error || new Error("Failed to load snapshots"));
    });
  });
}

export function deleteSnapshotFromDB(id: string): Promise<void> {
  return openSnapshotsDB().then((db) => {
    return new Promise<void>((resolve, reject) => {
      const transaction = db.transaction("snapshots", "readwrite");
      const store = transaction.objectStore("snapshots");
      const request = store.delete(id);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error || new Error("Failed to delete snapshot"));
    });
  });
}

export default function DataComparison({
  mainData,
  fileName,
  setMainData,
  setRawImportedData,
  setColumns,
  setFileName,
  setCustomColConfigs,
  setMapping,
  setLoading,
  setProgress,
  setStatusMessage,
  onExportExcel,
}: DataComparisonProps) {
  // Mode switcher: diff (2 separate files) vs time-series (Month-by-Month / YoY) vs saved-snapshots
  const [subMode, setSubMode] = useState<"diff" | "time-series" | "saved-snapshots">("diff");

  // === PHẦN 3: KHO LƯU TRỮ VÀ SO SÁNH LIÊN KỲ ===
  const [savedSnapshots, setSavedSnapshots] = useState<SavedSnapshot[]>([]);
  
  // State variables for saving new snapshot
  const [newSnapFile, setNewSnapFile] = useState<any[] | null>(null);
  const [newSnapFileName, setNewSnapFileName] = useState<string>("");
  const [newSnapName, setNewSnapName] = useState<string>("");
  const [newSnapMonth, setNewSnapMonth] = useState<number>(new Date().getMonth() + 1);
  const [newSnapYear, setNewSnapYear] = useState<number>(new Date().getFullYear());
  const [newSnapColumns, setNewSnapColumns] = useState<string[]>([]);

  // New states for flexible cycle & category folders
  const [newSnapCycleType, setNewSnapCycleType] = useState<"month" | "quarter" | "year">("month");
  const [newSnapCycleValue, setNewSnapCycleValue] = useState<number>(new Date().getMonth() + 1);
  const [newSnapCategory, setNewSnapCategory] = useState<string>("Mặc định");
  const [newCustomCategory, setNewCustomCategory] = useState<string>("");
  const [collapsedFolders, setCollapsedFolders] = useState<Record<string, boolean>>({});

  // State variables for comparing 2 snapshots
  const [snapAId, setSnapAId] = useState<string>("");
  const [snapBId, setSnapBId] = useState<string>("");
  const [snapAKeyCol, setSnapAKeyCol] = useState<string>("");
  const [snapAMetricCol, setSnapAMetricCol] = useState<string>("");
  const [snapBKeyCol, setSnapBKeyCol] = useState<string>("");
  const [snapBMetricCol, setSnapBMetricCol] = useState<string>("");

  const [snapCompareResult, setSnapCompareResult] = useState<any[] | null>(null);

  // Load snapshots on mount
  useEffect(() => {
    getAllSnapshotsFromDB().then((snaps) => {
      setSavedSnapshots(snaps || []);
    }).catch(err => {
      console.error("Lỗi khi tải kho dữ liệu:", err);
    });
  }, []);

  const uniqueCategories = useMemo(() => {
    const cats = new Set<string>(["Mặc định", "Lưu Cá Thể", "Lưu Doanh Nghiệp", "Hợp Tác Xã"]);
    savedSnapshots.forEach(s => {
      if (s.category) cats.add(s.category);
    });
    return Array.from(cats).sort();
  }, [savedSnapshots]);

  const groupedSnapshots = useMemo<Record<string, SavedSnapshot[]>>(() => {
    const groups: Record<string, SavedSnapshot[]> = {};
    savedSnapshots.forEach(snap => {
      const cat = snap.category || "Mặc định";
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(snap);
    });
    return groups;
  }, [savedSnapshots]);

  const guessMonthYearFromFileName = (name: string) => {
    const clean = name.toLowerCase();
    // Look for patterns like T2, T02, thang 2, thang02, month 2, month02
    const mMatch = clean.match(/(?:t|thang|thg|month|m)\s*(\d+)/) || clean.match(/(?:_|-|q)(\d+)/);
    if (mMatch) {
      const parsedM = parseInt(mMatch[1]);
      if (parsedM >= 1 && parsedM <= 12) {
        setNewSnapMonth(parsedM);
      }
    }
    // Look for year 2020 to 2030
    const yMatch = clean.match(/(20\d{2})/);
    if (yMatch) {
      setNewSnapYear(parseInt(yMatch[1]));
    }
  };

  const handleSnapFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang đọc tệp tin: ${file.name}...`);

    const isCSV = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");

    if (isCSV) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const text = evt.target?.result as string;
          if (!text) throw new Error("Không thể đọc nội dung!");
          const data = parseCSV(text);
          if (data.length === 0) {
            alert("Tệp rỗng hoặc sai định dạng!");
            setLoading(false);
            return;
          }
          const keys = Object.keys(data[0] || {});
          setNewSnapFile(data);
          setNewSnapFileName(file.name);
          setNewSnapName(file.name.replace(/\.[^/.]+$/, ""));
          setNewSnapColumns(keys);
          
          guessMonthYearFromFileName(file.name);
          setStatusMessage(`Đã nạp tạm thời ${data.length} dòng. Vui lòng chọn Kỳ báo cáo và lưu lại.`);
        } catch (err: any) {
          alert("Lỗi khi nạp file CSV: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsText(file, "UTF-8");
    } else {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          if (!arrayBuffer) throw new Error("Không thể đọc nội dung!");
          const wb = XLSX.read(arrayBuffer, { type: "array" });
          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];
          if (data.length === 0) {
            alert("Tệp Excel rỗng!");
            setLoading(false);
            return;
          }
          const keys = Object.keys(data[0] || {});
          setNewSnapFile(data);
          setNewSnapFileName(file.name);
          setNewSnapName(file.name.replace(/\.[^/.]+$/, ""));
          setNewSnapColumns(keys);

          guessMonthYearFromFileName(file.name);
          setStatusMessage(`Đã nạp tạm thời ${data.length} dòng. Vui lòng chọn Kỳ báo cáo và lưu lại.`);
        } catch (err: any) {
          alert("Lỗi khi nạp file Excel: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const handleSaveActiveDataAsSnapshot = () => {
    if (mainData.length === 0) {
      alert("Không có dữ liệu chính nào đang chạy trong hệ thống để lưu trữ!");
      return;
    }
    const cols = Object.keys(mainData[0] || {});
    setNewSnapFile(mainData);
    setNewSnapFileName(fileName || "DuLieuHeThong.xlsx");
    setNewSnapName(fileName ? fileName.replace(/\.[^/.]+$/, "") : "Dữ liệu đang chạy");
    setNewSnapColumns(cols);
    guessMonthYearFromFileName(fileName || "");
    alert("Đã lấy dữ liệu đang chạy trên màn hình chính! Hãy kiểm tra thông tin Kỳ báo cáo, Tên Snapshot bên dưới và bấm 'Lưu vào Kho' để hoàn tất.");
  };

  const handleSaveSnapshot = async () => {
    if (!newSnapFile || newSnapFile.length === 0) {
      alert("Vui lòng tải một file hoặc lấy dữ liệu hiện tại trước khi lưu!");
      return;
    }
    if (!newSnapName.trim()) {
      alert("Vui lòng đặt tên gợi nhớ cho bản ghi lưu trữ này!");
      return;
    }

    setLoading(true);
    setStatusMessage("Đang lưu bản ghi vào kho lưu trữ lâu dài...");
    await sleep(200);

    try {
      const id = "snap_" + Date.now() + "_" + Math.random().toString(36).substr(2, 9);
      const finalCategory = newSnapCategory === "custom" 
        ? (newCustomCategory.trim() || "Mặc định") 
        : newSnapCategory;
      const finalCycleValue = newSnapCycleType === "year" ? 1 : newSnapCycleValue;

      const newSnap: SavedSnapshot = {
        id,
        name: newSnapName.trim(),
        month: newSnapCycleType === "month" ? finalCycleValue : 1,
        year: newSnapYear,
        cycleType: newSnapCycleType,
        cycleValue: finalCycleValue,
        category: finalCategory,
        fileName: newSnapFileName,
        data: newSnapFile,
        columns: newSnapColumns,
        createdAt: new Date().toLocaleString("vi-VN")
      };

      await saveSnapshotToDB(newSnap);
      
      const updated = await getAllSnapshotsFromDB();
      setSavedSnapshots(updated || []);

      setNewSnapFile(null);
      setNewSnapFileName("");
      setNewSnapName("");
      setNewCustomCategory("");
      setNewSnapCategory("Mặc định");
      setStatusMessage("Đã lưu trữ dữ liệu lâu dài thành công!");
      alert(`Đã lưu thành công Bản ghi "${newSnap.name}" vào mục "${finalCategory}" (Chu kỳ: ${getFormattedPeriod(newSnap)}) vào bộ nhớ lưu trữ lâu dài của hệ thống!`);
    } catch (err: any) {
      alert("Không thể lưu trữ do lỗi: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteSnapshot = async (id: string, name: string) => {
    if (!confirm(`Bạn có chắc chắn muốn xóa vĩnh viễn bản ghi lưu trữ: "${name}" khỏi kho hệ thống không?`)) {
      return;
    }
    setLoading(true);
    setStatusMessage("Đang xóa bản ghi...");
    try {
      await deleteSnapshotFromDB(id);
      const updated = await getAllSnapshotsFromDB();
      setSavedSnapshots(updated || []);
      setStatusMessage("Xóa thành công.");
    } catch (err: any) {
      alert("Lỗi khi xóa: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleLoadSnapshotAsActive = async (snap: SavedSnapshot) => {
    if (!snap.data || snap.data.length === 0) return;
    setLoading(true);
    setStatusMessage(`Đang tải dữ liệu "${snap.name}" lên màn hình chính...`);
    await sleep(200);

    try {
      setMainData(snap.data);
      setRawImportedData(snap.data);
      setColumns(snap.columns);
      setFileName(snap.fileName || `${snap.name}.xlsx`);
      
      setMapping({
        mota: "",
        manganh: "",
        xa: "",
        doanhthu: "",
        laodong: "",
        idCol: "",
      });

      const initConfigs = snap.columns.map((c) => ({
        originalName: c,
        use: true,
        newName: c,
        role: "" as any,
      }));
      setCustomColConfigs(initConfigs);
      setStatusMessage(`Đã nạp thành công ${snap.data.length} dòng từ Kho Lưu trữ.`);
      alert(`Đã nạp toàn bộ ${snap.data.length} dòng của Snapshot "${snap.name}" vào bảng phân tích chính! Bạn có thể xem và thực hiện các nghiệp vụ khác ngay bây giờ.`);
    } catch (err: any) {
      alert("Lỗi khi tải dữ liệu: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // Auto detect columns for Snapshot A and Snapshot B when selected
  useEffect(() => {
    const snapA = savedSnapshots.find(s => s.id === snapAId);
    if (snapA && snapA.columns.length > 0) {
      const guessedKey = snapA.columns.find(c => {
        const l = c.toLowerCase();
        return l.includes("id") || l.includes("mst") || l.includes("mã") || l.includes("key") || l.includes("code") || l.includes("định danh");
      }) || snapA.columns[0];
      setSnapAKeyCol(guessedKey);

      const guessedMetric = snapA.columns.find(c => {
        const l = c.toLowerCase();
        return l.includes("doanh thu") || l.includes("doanhthu") || l.includes("lao động") || l.includes("laodong") || l.includes("số lượng") || l.includes("soluong") || l.includes("sản lượng") || l.includes("giá trị") || l.includes("value") || l.includes("amount") || l.includes("tong") || l.includes("tổng");
      }) || snapA.columns[snapA.columns.length - 1] || "";
      setSnapAMetricCol(guessedMetric);
    }
  }, [snapAId, savedSnapshots]);

  useEffect(() => {
    const snapB = savedSnapshots.find(s => s.id === snapBId);
    if (snapB && snapB.columns.length > 0) {
      const guessedKey = snapB.columns.find(c => {
        const l = c.toLowerCase();
        return l.includes("id") || l.includes("mst") || l.includes("mã") || l.includes("key") || l.includes("code") || l.includes("định danh");
      }) || snapB.columns[0];
      setSnapBKeyCol(guessedKey);

      const guessedMetric = snapB.columns.find(c => {
        const l = c.toLowerCase();
        return l.includes("doanh thu") || l.includes("doanhthu") || l.includes("lao động") || l.includes("laodong") || l.includes("số lượng") || l.includes("soluong") || l.includes("sản lượng") || l.includes("giá trị") || l.includes("value") || l.includes("amount") || l.includes("tong") || l.includes("tổng");
      }) || snapB.columns[snapB.columns.length - 1] || "";
      setSnapBMetricCol(guessedMetric);
    }
  }, [snapBId, savedSnapshots]);

  const handleCompareTwoSnapshots = async () => {
    if (!snapAId || !snapBId) {
      alert("Vui lòng chọn cả hai Bản ghi cần so sánh đối chiếu!");
      return;
    }
    if (snapAId === snapBId) {
      alert("Bạn phải chọn hai Bản ghi lưu trữ khác nhau để so sánh!");
      return;
    }
    if (!snapAKeyCol || !snapBKeyCol) {
      alert("Vui lòng cấu hình Cột khóa định danh khớp nối cho cả hai Bản ghi!");
      return;
    }
    if (!snapAMetricCol || !snapBMetricCol) {
      alert("Vui lòng cấu hình Cột chỉ tiêu đo lường số lượng cho cả hai Bản ghi!");
      return;
    }

    const snapA = savedSnapshots.find(s => s.id === snapAId);
    const snapB = savedSnapshots.find(s => s.id === snapBId);

    if (!snapA || !snapB) {
      alert("Không tìm thấy dữ liệu Bản ghi trong hệ thống!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Đang thực hiện phân tích đối chiếu liên kỳ...");
    await sleep(300);

    try {
      const mapA = new Map<string, any>();
      snapA.data.forEach(row => {
        const k = String(row[snapAKeyCol] || "").trim();
        if (k) mapA.set(k, row);
      });

      const mapB = new Map<string, any>();
      snapB.data.forEach(row => {
        const k = String(row[snapBKeyCol] || "").trim();
        if (k) mapB.set(k, row);
      });

      const results: any[] = [];
      const allKeys = Array.from(new Set([...mapA.keys(), ...mapB.keys()]));
      const totalKeys = allKeys.length;
      const batchSize = Math.max(1, Math.floor(totalKeys / 20));

      for (let i = 0; i < totalKeys; i++) {
        const key = allKeys[i];
        const rowA = mapA.get(key);
        const rowB = mapB.get(key);

        let valA = 0;
        let valB = 0;

        if (rowA) {
          const rawA = rowA[snapAMetricCol];
          valA = parseFloat(String(rawA || "").replace(/[^0-9.-]/g, "")) || 0;
        }
        if (rowB) {
          const rawB = rowB[snapBMetricCol];
          valB = parseFloat(String(rawB || "").replace(/[^0-9.-]/g, "")) || 0;
        }

        let diff = valB - valA;
        let percent = valA !== 0 ? (diff / valA) * 100 : 0;
        let statusStr = "";

        if (rowA && !rowB) {
          statusStr = "❌ Đã xóa";
        } else if (!rowA && rowB) {
          statusStr = "✅ Mới thêm";
        } else {
          statusStr = diff > 0 ? "📈 Tăng" : diff < 0 ? "📉 Giảm" : "💡 Không đổi";
        }

        results.push({
          [snapBKeyCol || "Khóa định danh"]: key,
          "Kỳ đối chiếu": getFormattedPeriod(snapA),
          [`[Gốc] ${snapAMetricCol || "Chỉ tiêu"}`]: parseFloat(valA.toFixed(2)),
          "Kỳ báo cáo": getFormattedPeriod(snapB),
          [`[So khớp] ${snapBMetricCol || "Chỉ tiêu"}`]: parseFloat(valB.toFixed(2)),
          [`Chênh lệch (${snapBMetricCol || "Chỉ tiêu"})`]: parseFloat(diff.toFixed(2)),
          "% Thay đổi": parseFloat(percent.toFixed(2)),
          "Trạng thái": statusStr
        });

        if (i % batchSize === 0 || i === totalKeys - 1) {
          setProgress(Math.floor((i / totalKeys) * 100));
          await sleep(5);
        }
      }

      setSnapCompareResult(results);
      
      const resCols = Object.keys(results[0] || {});
      setMainData(results);
      setRawImportedData(results);
      setColumns(resCols);
      
      const cleanFileA = snapA.name.replace(/\s+/g, "_");
      const cleanFileB = snapB.name.replace(/\s+/g, "_");
      setFileName(`SoSanh_${cleanFileB}_vs_${cleanFileA}.xlsx`);

      setProgress(100);
      setStatusMessage(`Đã đối chiếu thành công! Tìm thấy ${results.length} thực thể khớp nối.`);
      alert(`Đã đối chiếu thành công Bản ghi "${snapB.name}" (${getFormattedPeriod(snapB)}) so với "${snapA.name}" (${getFormattedPeriod(snapA)})!\nKết quả so sánh đã được nạp trực tiếp vào Bảng Dữ liệu Phân tích chính.`);
    } catch (err: any) {
      alert("Lỗi khi đối chiếu: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // === PHẦN 1: SO SÁNH HAI NIÊN ĐỘ (DIFF CŨ - MỚI) ===
  const [oldData, setOldData] = useState<any[]>([]);
  const [oldFileName, setOldFileName] = useState<string>("");
  const [newData, setNewData] = useState<any[]>([]);
  const [newFileName, setNewFileName] = useState<string>("");
  const [diffKey, setDiffKey] = useState<string>("");
  const [compareResultData, setCompareResultData] = useState<any[] | null>(null);

  // === PHẦN 2: SO SÁNH CHUỖI THỜI GIAN (Yêu cầu Mới) ===
  const [tsData, setTsData] = useState<any[]>([]);
  const [tsFileName, setTsFileName] = useState<string>("");
  const [tsKey, setTsKey] = useState<string>(""); // Khóa định danh địa bàn/DN
  
  // Cấu hình thời gian
  const [timeConfigType, setTimeConfigType] = useState<"single" | "separate">("separate");
  const [singleDateCol, setSingleDateCol] = useState<string>("");
  const [monthCol, setMonthCol] = useState<string>("");
  const [yearCol, setYearCol] = useState<string>("");
  const [metricCol, setMetricCol] = useState<string>(""); // Cột chỉ tiêu số lượng (Doanh thu, lao động...)
  const [comparisonType, setComparisonType] = useState<"adjacent" | "yoy">("adjacent");
  
  // Kết quả so sánh chuỗi thời gian
  const [tsResultRows, setTsResultRows] = useState<any[]>([]);
  const [selectedChartEntity, setSelectedChartEntity] = useState<string>(""); // Bộ lọc biểu đồ theo thực thể

  // Đọc file tải lên cho phần 1 & 2
  const handleLocalFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "old" | "new" | "ts") => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang nạp file: ${file.name}...`);

    const isCSV = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");

    if (isCSV) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const text = evt.target?.result as string;
          if (!text) throw new Error("Không thể đọc nội dung tệp tin!");

          const data = parseCSV(text);
          if (data.length === 0) {
            alert("Tệp rỗng hoặc sai định dạng!");
            setLoading(false);
            return;
          }

          const keys = Object.keys(data[0] || {});
          if (type === "old") {
            setOldData(data);
            setOldFileName(file.name);
            if (keys.length > 0 && !diffKey) setDiffKey(keys[0]);
          } else if (type === "new") {
            setNewData(data);
            setNewFileName(file.name);
            if (keys.length > 0 && !diffKey) setDiffKey(keys[0]);
          } else {
            setTsData(data);
            setTsFileName(file.name);
            autoGuessTSConfigs(data);
          }
          setStatusMessage(`Đã nạp thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file CSV: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsText(file, "UTF-8");
    } else {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          if (!arrayBuffer) throw new Error("Không thể đọc nội dung!");

          const wb = XLSX.read(arrayBuffer, {
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false,
          });

          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          if (data.length === 0) {
            alert("Tệp Excel rỗng hoặc sai định dạng!");
            setLoading(false);
            return;
          }

          const keys = Object.keys(data[0] || {});
          if (type === "old") {
            setOldData(data);
            setOldFileName(file.name);
            if (keys.length > 0 && !diffKey) setDiffKey(keys[0]);
          } else if (type === "new") {
            setNewData(data);
            setNewFileName(file.name);
            if (keys.length > 0 && !diffKey) setDiffKey(keys[0]);
          } else {
            setTsData(data);
            setTsFileName(file.name);
            autoGuessTSConfigs(data);
          }
          setStatusMessage(`Đã nạp thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file Excel: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Tự động phân tích và đoán cột thời gian, khóa định danh
  const autoGuessTSConfigs = (data: any[]) => {
    if (data.length === 0) return;
    const cols = Object.keys(data[0] || {});
    
    // Tìm khóa định danh
    const guessedKey = cols.find(c => {
      const l = c.toLowerCase();
      return l.includes("id") || l.includes("mst") || l.includes("mã") || l.includes("key") || l.includes("code") || l.includes("định danh");
    });
    if (guessedKey) setTsKey(guessedKey);

    // Tìm cột tháng, năm riêng biệt hoặc chung
    const guessedMonth = cols.find(c => c.toLowerCase().includes("tháng") || c.toLowerCase() === "thang" || c.toLowerCase() === "month" || c.toLowerCase() === "thg");
    const guessedYear = cols.find(c => c.toLowerCase().includes("năm") || c.toLowerCase() === "nam" || c.toLowerCase() === "year" || c.toLowerCase() === "nm");
    const guessedDate = cols.find(c => c.toLowerCase().includes("ngày") || c.toLowerCase().includes("date") || c.toLowerCase().includes("thời gian") || c.toLowerCase().includes("period"));

    if (guessedMonth && guessedYear) {
      setTimeConfigType("separate");
      setMonthCol(guessedMonth);
      setYearCol(guessedYear);
    } else if (guessedDate) {
      setTimeConfigType("single");
      setSingleDateCol(guessedDate);
    } else {
      // Dự phòng
      setTimeConfigType("separate");
      if (cols.length > 1) {
        setMonthCol(cols[1]);
        setYearCol(cols[2] || cols[1]);
      }
    }

    // Dự đoán cột chỉ tiêu số học
    const guessedMetric = cols.find(c => {
      const l = c.toLowerCase();
      return l.includes("doanh thu") || l.includes("doanhthu") || l.includes("lao động") || l.includes("laodong") || l.includes("số lượng") || l.includes("soluong") || l.includes("sản lượng") || l.includes("giá trị") || l.includes("value") || l.includes("amount") || l.includes("tong") || l.includes("tổng");
    });
    if (guessedMetric) {
      setMetricCol(guessedMetric);
    } else {
      // Tìm cột đầu tiên có kiểu số hoặc giữ cột cuối
      setMetricCol(cols[cols.length - 1] || "");
    }
  };

  // Logic 1: So sánh đối chiếu cũ - mới (Diff)
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
    oldData.forEach((row) => {
      if (!row || typeof row !== "object") return;
      const k = String(row[diffKey] || "").trim();
      if (k) oldMap.set(k, row);
    });

    const newMap = new Map();
    newData.forEach((row) => {
      if (!row || typeof row !== "object") return;
      const k = String(row[diffKey] || "").trim();
      if (k) newMap.set(k, row);
    });

    const resultRows: any[] = [];
    const allKeys = Array.from(new Set([...oldMap.keys(), ...newMap.keys()]));
    const batchSize = Math.max(1, Math.floor(allKeys.length / 20));

    const firstOldRow = oldData.find((r) => r && typeof r === "object") || {};
    const firstNewRow = newData.find((r) => r && typeof r === "object") || {};
    const oldCols = Object.keys(firstOldRow);
    const newCols = Object.keys(firstNewRow);
    const unionCols = Array.from(new Set([...oldCols, ...newCols])).filter((c) => c !== diffKey);

    for (let i = 0; i < allKeys.length; i++) {
      const key = allKeys[i];
      const oldRow = oldMap.get(key);
      const newRow = newMap.get(key);

      const combined: any = { [diffKey]: key };

      if (oldRow && !newRow) {
        unionCols.forEach((col) => {
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = "";
        });
        combined["TrangThai_SoSanh"] = "❌ Đã xóa";
      } else if (!oldRow && newRow) {
        unionCols.forEach((col) => {
          combined[`${col}_Cu`] = "";
          combined[`${col}_Moi`] = newRow[col] || "";
        });
        combined["TrangThai_SoSanh"] = "✅ Mới thêm";
      } else {
        const changedCols: string[] = [];
        unionCols.forEach((col) => {
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
        setStatusMessage(`Đang so sánh đối chiếu: ${i}/${allKeys.length} dòng...`);
        await sleep(10);
      }
    }

    const compareCols = Object.keys(resultRows[0] || {});
    setMainData(resultRows);
    setRawImportedData(resultRows);
    setColumns(compareCols);
    setFileName(`SoSanhDiff_${oldFileName.replace(/\.[^/.]+$/, "")}_vs_${newFileName.replace(/\.[^/.]+$/, "")}.xlsx`);
    setCompareResultData(resultRows);
    setMapping({
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: "",
    });

    const initCompareConfigs = compareCols.map((c) => {
      return {
        originalName: c,
        use: true,
        newName: c,
        role: "" as any,
      };
    });
    setCustomColConfigs(initCompareConfigs);

    setProgress(100);
    setStatusMessage(`So sánh thành công! Tìm thấy tổng cộng ${resultRows.length} khóa định danh.`);
    await sleep(400);
    setLoading(false);
  };

  // Logic 2: So sánh Chuỗi Thời gian (Adjacent & YoY same month)
  const handleTimeSeriesCompare = async () => {
    const activeDataset = tsData.length > 0 ? tsData : mainData;
    const activeName = tsFileName || fileName || "Hệ thống hiện tại";

    if (activeDataset.length === 0) {
      alert("Không có dữ liệu nguồn để so sánh chuỗi thời gian! Vui lòng nạp dữ liệu.");
      return;
    }

    if (!tsKey) {
      alert("Vui lòng chọn Cột Khóa định danh địa bàn/thực thể!");
      return;
    }

    if (!metricCol) {
      alert("Vui lòng chọn Cột Chỉ tiêu số liệu cần đối chiếu!");
      return;
    }

    if (timeConfigType === "separate" && (!monthCol || !yearCol)) {
      alert("Vui lòng cấu hình đầy đủ cột Tháng và cột Năm!");
      return;
    }

    if (timeConfigType === "single" && !singleDateCol) {
      alert("Vui lòng chọn cột Ngày/Tháng thời gian!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu xử lý chuỗi thời gian...");
    await sleep(300);

    // BƯỚC 1: PARSE THỜI GIAN VÀ GROUP DỮ LIỆU
    // Cấu trúc map lưu trữ: Key -> Year -> Month -> Value
    const groupedMap = new Map<string, Map<number, Map<number, number>>>();
    const allYearsSet = new Set<number>();
    const allMonthsSet = new Set<number>();
    const allEntitiesSet = new Set<string>();

    activeDataset.forEach((row) => {
      const entityId = String(row[tsKey] || "").trim();
      if (!entityId) return;

      allEntitiesSet.add(entityId);

      let m = 0;
      let y = 0;

      if (timeConfigType === "separate") {
        m = parseInt(String(row[monthCol]).replace(/[^0-9]/g, "")) || 0;
        y = parseInt(String(row[yearCol]).replace(/[^0-9]/g, "")) || 0;
      } else {
        // Tự động phân tích trường thời gian chung (ví dụ: '02/2026', '2026-02-15')
        const rawDateStr = String(row[singleDateCol] || "").trim();
        if (rawDateStr) {
          // Định dạng MM/YYYY
          if (rawDateStr.includes("/")) {
            const parts = rawDateStr.split("/");
            if (parts.length >= 2) {
              m = parseInt(parts[0]) || 0;
              y = parseInt(parts[1]) || 0;
            }
          } else if (rawDateStr.includes("-")) {
            // Định dạng YYYY-MM-DD
            const parts = rawDateStr.split("-");
            if (parts[0].length === 4) {
              y = parseInt(parts[0]) || 0;
              m = parseInt(parts[1]) || 0;
            } else if (parts[2]?.length === 4) {
              // DD-MM-YYYY
              y = parseInt(parts[2]) || 0;
              m = parseInt(parts[1]) || 0;
            }
          } else {
            // Chuyển trực tiếp thành số
            const parsedVal = parseInt(rawDateStr);
            if (parsedVal > 100000) { // Giả lập timestamp hoặc YYMMDD
              const d = new Date(parsedVal);
              m = d.getMonth() + 1;
              y = d.getFullYear();
            } else {
              y = parsedVal;
            }
          }
        }
      }

      // Đảm bảo tháng và năm hợp lệ
      if (m < 1 || m > 12) m = 1;
      if (y < 1900 || y > 2100) y = new Date().getFullYear();

      allYearsSet.add(y);
      allMonthsSet.add(m);

      const rawMetricVal = row[metricCol];
      const metricVal = parseFloat(String(rawMetricVal).replace(/[^0-9.-]/g, "")) || 0;

      // Đưa vào nhóm
      let yearMap = groupedMap.get(entityId);
      if (!yearMap) {
        yearMap = new Map();
        groupedMap.set(entityId, yearMap);
      }

      let monthMap = yearMap.get(y);
      if (!monthMap) {
        monthMap = new Map();
        yearMap.set(y, monthMap);
      }

      // Cộng dồn nếu có nhiều bản ghi cùng kỳ
      const existingVal = monthMap.get(m) || 0;
      monthMap.set(m, existingVal + metricVal);
    });

    // BƯỚC 2: TIẾN HÀNH ĐỐI CHIẾU SO SÁNH THEO LOẠI YÊU CẦU
    const results: any[] = [];
    const entityList = Array.from(allEntitiesSet);
    const sortedYears = Array.from(allYearsSet).sort((a, b) => a - b);
    const sortedMonths = Array.from(allMonthsSet).sort((a, b) => a - b);

    entityList.forEach((entityId) => {
      const yearMap = groupedMap.get(entityId);
      if (!yearMap) return;

      sortedYears.forEach((y) => {
        const monthMap = yearMap.get(y);
        if (!monthMap) return;

        sortedMonths.forEach((m) => {
          const valCurrent = monthMap.get(m);
          if (valCurrent === undefined) return;

          let valPrior = 0;
          let priorLabel = "";
          let hasPrior = false;

          if (comparisonType === "adjacent") {
            // 1. SO SÁNH THÁNG LIÊN TIẾP GẦN NHAU (T vs T-1)
            let priorM = m - 1;
            let priorY = y;
            if (priorM === 0) {
              priorM = 12;
              priorY = y - 1;
            }

            const priorYearMap = yearMap.get(priorY);
            if (priorYearMap) {
              const val = priorYearMap.get(priorM);
              if (val !== undefined) {
                valPrior = val;
                hasPrior = true;
              }
            }
            priorLabel = `Tháng ${priorM}/${priorY}`;
          } else {
            // 2. SO SÁNH CÙNG KỲ NĂM TRƯỚC (YoY) - Ví dụ Tháng 2 năm nay vs Tháng 2 năm trước
            const priorY = y - 1;
            const priorYearMap = yearMap.get(priorY);
            if (priorYearMap) {
              const val = priorYearMap.get(m);
              if (val !== undefined) {
                valPrior = val;
                hasPrior = true;
              }
            }
            priorLabel = `Tháng ${m}/${priorY}`;
          }

          if (hasPrior) {
            const diff = valCurrent - valPrior;
            const percent = valPrior !== 0 ? (diff / valPrior) * 100 : 0;
            const statusStr = diff > 0 ? "📈 Tăng" : diff < 0 ? "📉 Giảm" : "💡 Không đổi";

            results.push({
              [tsKey]: entityId,
              "Kỳ hiện tại": `Tháng ${m}/${y}`,
              "Giá trị kỳ này": valCurrent,
              "Kỳ đối chiếu": priorLabel,
              "Giá trị kỳ trước": valPrior,
              "Chênh lệch tuyệt đối": parseFloat(diff.toFixed(2)),
              "% Thay đổi": parseFloat(percent.toFixed(2)),
              "Trạng thái": statusStr,
              "_year": y,
              "_month": m
            });
          }
        });
      });
    });

    setTsResultRows(results);
    if (entityList.length > 0) {
      setSelectedChartEntity(entityList[0]);
    }

    // Nạp kết quả vào bảng chính để xem
    if (results.length > 0) {
      const resCols = Object.keys(results[0]).filter(k => !k.startsWith("_"));
      setMainData(results);
      setRawImportedData(results);
      setColumns(resCols);
      setFileName(`SoSanhChuoiThoiGian_${comparisonType === "adjacent" ? "LienTiep" : "CungKy"}_${activeName.replace(/\.[^/.]+$/, "")}.xlsx`);
    }

    setProgress(100);
    setStatusMessage(`Đối chiếu chuỗi thời gian thành công! Tạo ra ${results.length} dòng chênh lệch đối chiếu.`);
    await sleep(400);
    setLoading(false);
  };

  // Lọc dữ liệu hiển thị trên biểu đồ theo thực thể được chọn
  const chartDataForSelected = useMemo(() => {
    if (!selectedChartEntity) return [];
    return tsResultRows
      .filter(row => row[tsKey] === selectedChartEntity)
      .sort((a, b) => {
        if (a._year !== b._year) return a._year - b._year;
        return a._month - b._month;
      })
      .map(row => ({
        name: row["Kỳ hiện tại"],
        "Kỳ này": row["Giá trị kỳ này"],
        "Kỳ trước": row["Giá trị kỳ trước"],
        "Thay đổi %": row["% Thay đổi"]
      }));
  }, [tsResultRows, selectedChartEntity, tsKey]);

  return (
    <div className="space-y-6 animate-fade-in font-sans">
      
      {/* Pill switcher for Sub modes */}
      <div className="flex border-b border-slate-200">
        <button
          onClick={() => setSubMode("diff")}
          className={`py-3 px-6 text-sm font-bold border-b-2 transition-all flex items-center gap-2 cursor-pointer ${
            subMode === "diff"
              ? "border-sky-500 text-sky-600 bg-sky-50/40"
              : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
          }`}
        >
          <RefreshCw className="w-4 h-4 text-sky-500" />
          So sánh Hai Niên Độ (Diff)
        </button>
        <button
          onClick={() => setSubMode("time-series")}
          className={`py-3 px-6 text-sm font-bold border-b-2 transition-all flex items-center gap-2 cursor-pointer ${
            subMode === "time-series"
              ? "border-sky-500 text-sky-600 bg-sky-50/40"
              : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
          }`}
        >
          <Calendar className="w-4 h-4 text-amber-500" />
          So sánh Chuỗi Thời gian (Tháng liên tiếp & Cùng kỳ YoY)
          <span className="bg-amber-100 text-amber-800 text-[10px] px-1.5 py-0.5 rounded-full font-mono font-bold">Mới</span>
        </button>
        <button
          onClick={() => setSubMode("saved-snapshots")}
          className={`py-3 px-6 text-sm font-bold border-b-2 transition-all flex items-center gap-2 cursor-pointer ${
            subMode === "saved-snapshots"
              ? "border-sky-500 text-sky-600 bg-sky-50/40"
              : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
          }`}
        >
          <Database className="w-4 h-4 text-emerald-500" />
          Kho Lưu trữ Lâu dài & Đối chiếu Liên kỳ
          <span className="bg-emerald-100 text-emerald-800 text-[10px] px-1.5 py-0.5 rounded-full font-mono font-bold">Lưu lâu dài</span>
        </button>
      </div>

      {/* CHẾ ĐỘ 1: ĐỐI CHIẾU HAI NIÊN ĐỘ (DIFF CŨ - MỚI) */}
      {subMode === "diff" && (
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-6 text-slate-800">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <RefreshCw className="w-5 h-5 text-sky-500 animate-spin-slow" /> SO SÁNH ĐỐI CHIẾU HAI NIÊN ĐỘ (DIFF)
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Hệ thống tự động rà soát đối chiếu chéo hai bảng dữ liệu Cũ và Mới để tìm ra phần tử Mới thêm, Đã xóa hoặc Thay đổi thuộc tính giữa hai thời kỳ.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* BẢNG CŨ */}
            <div className="bg-slate-50 border border-slate-100 p-5 rounded-xl space-y-4">
              <span className="text-xs font-bold text-amber-600 tracking-wider uppercase font-mono block">
                📁 1. BẢNG DỮ LIỆU CŨ (MỐC ĐỐI CHIẾU)
              </span>

              <label className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 cursor-pointer w-full justify-center">
                <FileUp className="w-4 h-4 text-amber-500" /> CHỌN BẢNG CŨ (.xlsx, .xls, .csv)
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => handleLocalFileUpload(e, "old")}
                  className="hidden"
                />
              </label>

              {mainData && mainData.length > 0 && (
                <button
                  onClick={() => {
                    setOldData(mainData);
                    setOldFileName(fileName || "Dữ liệu chính hệ thống");
                  }}
                  type="button"
                  className="w-full bg-amber-50 hover:bg-amber-100 text-amber-700 font-bold text-[11px] px-3 py-2 rounded-xl border border-amber-200 transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                >
                  📥 SỬ DỤNG FILE CHÍNH ĐÃ NẠP
                </button>
              )}

              {oldFileName && (
                <div className="bg-white border border-slate-200 p-3 rounded-lg text-xs flex justify-between items-center text-slate-700 shadow-sm">
                  <span className="truncate max-w-[200px]" title={oldFileName}>📄 {oldFileName}</span>
                  <span className="font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">{oldData.length} dòng</span>
                </div>
              )}
            </div>

            {/* BẢNG MỚI */}
            <div className="bg-slate-50 border border-slate-100 p-5 rounded-xl space-y-4">
              <span className="text-xs font-bold text-sky-600 tracking-wider uppercase font-mono block">
                📁 2. BẢNG DỮ LIỆU MỚI (CẬP NHẬT)
              </span>

              <label className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 cursor-pointer w-full justify-center">
                <FileUp className="w-4 h-4 text-sky-500" /> CHỌN BẢNG MỚI (.xlsx, .xls, .csv)
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => handleLocalFileUpload(e, "new")}
                  className="hidden"
                />
              </label>

              {mainData && mainData.length > 0 && (
                <button
                  onClick={() => {
                    setNewData(mainData);
                    setNewFileName(fileName || "Dữ liệu chính hệ thống");
                  }}
                  type="button"
                  className="w-full bg-sky-50 hover:bg-sky-100 text-sky-700 font-bold text-[11px] px-3 py-2 rounded-xl border border-sky-200 transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                >
                  📥 SỬ DỤNG FILE CHÍNH ĐÃ NẠP
                </button>
              )}

              {newFileName && (
                <div className="bg-white border border-slate-200 p-3 rounded-lg text-xs flex justify-between items-center text-slate-700 shadow-sm">
                  <span className="truncate max-w-[200px]" title={newFileName}>📄 {newFileName}</span>
                  <span className="font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">{newData.length} dòng</span>
                </div>
              )}
            </div>
          </div>

          {/* KHÓA CHÍNH ĐỂ ĐỐI CHIẾU */}
          {(oldData.length > 0 || newData.length > 0) && (
            <div className="bg-slate-50 border border-slate-100 p-5 rounded-xl space-y-3 shadow-sm">
              <label className="text-xs font-bold text-slate-700 block">
                🔑 Chọn Cột Khóa Chính Định Danh Độc Nhất (Unique Key):
              </label>
              <p className="text-[10px] text-slate-400">
                Chọn cột thông tin duy nhất dùng để đối chiếu so khớp từng dòng (ví dụ: Mã Số Thuế, Số định danh, ID doanh nghiệp...).
              </p>
              <select
                value={diffKey}
                onChange={(e) => setDiffKey(e.target.value)}
                className="w-full md:max-w-md bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-sky-500"
              >
                <option value="">-- Chọn cột khóa chính --</option>
                {Object.keys(newData[0] || oldData[0] || {}).map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>
          )}

          <button
            onClick={handleCompare}
            className="w-full bg-gradient-to-r from-sky-500 to-sky-600 hover:from-sky-600 hover:to-sky-700 text-white font-bold text-xs px-6 py-3.5 rounded-xl border-b-4 border-sky-700 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
          >
            <RefreshCw className="w-4 h-4" /> THỰC HIỆN SO SÁNH ĐỐI CHIẾU & NẠP VÀO HỆ THỐNG XEM CHÍNH
          </button>
        </div>
      )}

      {/* CHẾ ĐỘ 2: SO SÁNH CHUỖI THỜI GIAN (DÀNH CHO THÁNG GẦN NHAU / THÁNG 2 NĂM NAY VS THÁNG 2 NĂM TRƯỚC) */}
      {subMode === "time-series" && (
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-6 text-slate-800 animate-fade-in">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <Calendar className="w-5 h-5 text-amber-500" /> ĐỐI CHIẾU CHUỖI THỜI GIAN (LIÊN TIẾP & CÙNG KỲ NĂM TRƯỚC)
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Hỗ trợ phân tích so sánh động chỉ tiêu đo lường giữa các tháng liên tiếp kề nhau hoặc so sánh cùng kỳ tháng này năm nay với tháng này năm ngoái (ví dụ: Tháng 2 năm nay so với Tháng 2 năm trước).
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            
            {/* Cột 1: File nguồn và Loại đối chiếu */}
            <div className="space-y-4 md:col-span-1 border-r border-slate-100 pr-0 md:pr-6">
              <span className="text-xs font-extrabold text-amber-700 tracking-wider uppercase font-mono block">
                📂 1. CHỌN DỮ LIỆU NGUỒN CHUỖI THỜI GIAN
              </span>

              <label className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 cursor-pointer w-full justify-center">
                <FileUp className="w-4 h-4 text-amber-500" /> TẢI LÊN FILE CHUỖI THỜI GIAN
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => handleLocalFileUpload(e, "ts")}
                  className="hidden"
                />
              </label>

              {mainData && mainData.length > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setTsData(mainData);
                    setTsFileName(fileName || "Dữ liệu chính hệ thống");
                    autoGuessTSConfigs(mainData);
                  }}
                  className="w-full bg-amber-50 hover:bg-amber-100 text-amber-700 font-bold text-[11px] px-3 py-2 rounded-xl border border-amber-200 transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                >
                  📥 SỬ DỤNG FILE CHÍNH ĐÃ NẠP
                </button>
              )}

              {tsFileName && (
                <div className="bg-amber-50/50 border border-amber-100 p-3 rounded-lg text-xs flex justify-between items-center text-slate-700 shadow-xs">
                  <span className="truncate max-w-[180px] font-semibold">📄 {tsFileName}</span>
                  <span className="font-mono text-amber-800 bg-amber-100/50 px-1.5 py-0.5 rounded">{tsData.length} dòng</span>
                </div>
              )}

              {/* Loại đối chiếu */}
              <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 space-y-3 shadow-xs">
                <div className="space-y-1">
                  <label className="text-[11px] font-black text-slate-500 uppercase tracking-wider block font-mono">Phương thức so sánh:</label>
                  <div className="space-y-1.5">
                    <label className="flex items-center gap-2 text-xs font-bold text-slate-700 cursor-pointer">
                      <input
                        type="radio"
                        name="comparisonType"
                        checked={comparisonType === "adjacent"}
                        onChange={() => setComparisonType("adjacent")}
                        className="text-amber-500 focus:ring-amber-400 h-3.5 w-3.5"
                      />
                      So sánh giữa các tháng gần nhau kề tiếp
                    </label>
                    <label className="flex items-center gap-2 text-xs font-bold text-slate-700 cursor-pointer">
                      <input
                        type="radio"
                        name="comparisonType"
                        checked={comparisonType === "yoy"}
                        onChange={() => setComparisonType("yoy")}
                        className="text-amber-500 focus:ring-amber-400 h-3.5 w-3.5"
                      />
                      Cùng kỳ năm trước (YoY) (Ví dụ: Tháng 2 vs Tháng 2 năm ngoái)
                    </label>
                  </div>
                </div>
              </div>
            </div>

            {/* Cột 2: Cấu hình Cột định danh và Cột thời gian */}
            <div className="space-y-4 md:col-span-1 border-r border-slate-100 pr-0 md:pr-6">
              <span className="text-xs font-extrabold text-indigo-700 tracking-wider uppercase font-mono block">
                ⚙️ 2. ÁNH XẠ CỘT CHUỖI THỜI GIAN
              </span>

              {/* Khóa định danh địa bàn */}
              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-700 block">
                  🔑 Khóa định danh địa bàn/thực thể (Entity Key):
                </label>
                <select
                  value={tsKey}
                  onChange={(e) => setTsKey(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500"
                >
                  <option value="">-- Chọn cột định danh địa bàn --</option>
                  {Object.keys((tsData.length > 0 ? tsData : mainData)[0] || {}).map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <p className="text-[10px] text-slate-400">Dùng mã này để nhóm các giá trị của một địa bàn qua thời kỳ.</p>
              </div>

              {/* Cột chỉ tiêu định lượng */}
              <div className="space-y-1 pt-1">
                <label className="text-xs font-bold text-slate-700 block">
                  📈 Cột Chỉ tiêu cần phân tích:
                </label>
                <select
                  value={metricCol}
                  onChange={(e) => setMetricCol(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500"
                >
                  <option value="">-- Chọn cột chỉ tiêu số liệu --</option>
                  {Object.keys((tsData.length > 0 ? tsData : mainData)[0] || {}).map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <p className="text-[10px] text-slate-400">Ví dụ: Doanh thu, số lượng doanh nghiệp đạt, v.v.</p>
              </div>

              {/* Cấu hình thời gian */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 shadow-xs">
                <div className="space-y-1">
                  <label className="text-[11px] font-black text-slate-500 uppercase tracking-wider block font-mono">Dạng thức thời gian:</label>
                  <div className="flex gap-4">
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer">
                      <input
                        type="radio"
                        checked={timeConfigType === "separate"}
                        onChange={() => setTimeConfigType("separate")}
                        className="text-amber-500"
                      />
                      Cột Tháng & Năm riêng biệt
                    </label>
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer">
                      <input
                        type="radio"
                        checked={timeConfigType === "single"}
                        onChange={() => setTimeConfigType("single")}
                        className="text-amber-500"
                      />
                      Một cột ngày tháng chung
                    </label>
                  </div>
                </div>

                {timeConfigType === "separate" ? (
                  <div className="grid grid-cols-2 gap-2 pt-1">
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-slate-600 block">Cột chứa Tháng:</label>
                      <select
                        value={monthCol}
                        onChange={(e) => setMonthCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800"
                      >
                        <option value="">-- Tháng --</option>
                        {Object.keys((tsData.length > 0 ? tsData : mainData)[0] || {}).map((c) => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-slate-600 block">Cột chứa Năm:</label>
                      <select
                        value={yearCol}
                        onChange={(e) => setYearCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800"
                      >
                        <option value="">-- Năm --</option>
                        {Object.keys((tsData.length > 0 ? tsData : mainData)[0] || {}).map((c) => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-1 pt-1">
                    <label className="text-[10px] font-bold text-slate-600 block">Cột Ngày Tháng chung:</label>
                    <select
                      value={singleDateCol}
                      onChange={(e) => setSingleDateCol(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1 text-xs text-slate-800"
                    >
                      <option value="">-- Chọn cột thời gian --</option>
                      {Object.keys((tsData.length > 0 ? tsData : mainData)[0] || {}).map((c) => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            </div>

            {/* Cột 3: Thực hiện hành động & Xem trước xu hướng nhanh */}
            <div className="space-y-4 md:col-span-1">
              <span className="text-xs font-extrabold text-emerald-800 tracking-wider uppercase font-mono block">
                ⚡ 3. TIẾN HÀNH PHÂN TÍCH
              </span>

              <p className="text-xs text-slate-500 leading-relaxed">
                Sau khi ánh xạ thành công các thuộc tính cấu trúc thời gian và mã khóa, bấm nút dưới để hệ thống khởi chạy thuật toán tính toán tăng trưởng.
              </p>

              <button
                onClick={handleTimeSeriesCompare}
                className="w-full bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-white font-bold text-xs px-6 py-4 rounded-xl border-b-4 border-amber-800 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
              >
                <TrendingUp className="w-4 h-4" /> BẮT ĐẦU ĐỐI CHIẾU CHUỖI THỜI GIAN
              </button>

              {tsResultRows.length > 0 && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 shadow-xs">
                  <span className="text-[11px] font-black text-slate-500 block uppercase font-mono">📈 Xu hướng địa bàn cụ thể:</span>
                  <div className="space-y-2">
                    <select
                      value={selectedChartEntity}
                      onChange={(e) => setSelectedChartEntity(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800"
                    >
                      {Array.from(new Set(tsResultRows.map(r => r[tsKey]))).map(entity => (
                        <option key={entity} value={entity}>{entity}</option>
                      ))}
                    </select>

                    <div className="h-28">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={chartDataForSelected} margin={{ top: 5, right: 5, left: -25, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} />
                          <XAxis dataKey="name" tick={{ fontSize: 9 }} />
                          <YAxis tick={{ fontSize: 9 }} />
                          <Tooltip contentStyle={{ fontSize: 10 }} />
                          <Line type="monotone" dataKey="Kỳ này" stroke="#d97706" strokeWidth={2} name="Kỳ này" dot={{ r: 3 }} />
                          <Line type="monotone" dataKey="Kỳ trước" stroke="#475569" strokeWidth={1.5} name="Kỳ trước" dot={{ r: 2 }} strokeDasharray="4 4" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              )}
            </div>

          </div>
        </div>
      )}

      {/* CHẾ ĐỘ 3: KHO LƯU TRỮ VÀ SO SÁNH LIÊN KỲ ĐỘNG */}
      {subMode === "saved-snapshots" && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            
            {/* Cột trái (Lớp quản lý kho lưu trữ) */}
            <div className="lg:col-span-5 space-y-6">
              
              {/* Thêm mới vào kho */}
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-5 space-y-4 text-slate-800">
                <h3 className="text-sm font-extrabold text-emerald-800 tracking-wider uppercase font-mono flex items-center gap-2">
                  <Database className="w-5 h-5 text-emerald-600" /> 1. Lưu trữ dữ liệu mới vào kho
                </h3>
                <p className="text-xs text-slate-500 leading-relaxed">
                  Bạn có thể tải lên một tệp Excel/CSV bất kỳ để lưu trữ lâu dài vào cơ sở dữ liệu trình duyệt, hoặc lưu trực tiếp bảng dữ liệu đang phân tích trên màn hình.
                </p>

                {/* Drag drop or selector area */}
                <div className="border-2 border-dashed border-emerald-200 hover:border-emerald-400 bg-emerald-50/20 hover:bg-emerald-50/50 rounded-xl p-4 transition-all text-center space-y-2 relative">
                  <input
                    type="file"
                    onChange={handleSnapFileUpload}
                    accept=".xlsx,.xls,.csv"
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  />
                  <div className="flex flex-col items-center justify-center gap-1">
                    <FileUp className="w-8 h-8 text-emerald-500" />
                    <span className="text-xs font-bold text-slate-700">Kéo thả hoặc Click để nạp tệp tin mới</span>
                    <span className="text-[10px] text-slate-400">Chấp nhận .XLSX, .XLS, .CSV</span>
                  </div>
                </div>

                {/* Hoặc lấy dữ liệu hệ thống */}
                <div className="flex items-center justify-between gap-2 pt-1">
                  <div className="h-px bg-slate-200 flex-1"></div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase font-mono">Hoặc</span>
                  <div className="h-px bg-slate-200 flex-1"></div>
                </div>

                <button
                  onClick={handleSaveActiveDataAsSnapshot}
                  className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs px-3 py-2 rounded-lg border border-slate-300 transition-all cursor-pointer flex items-center justify-center gap-1.5"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-500" /> Lấy từ Dữ liệu đang chạy ngoài màn hình chính
                </button>

                {/* Thông tin metadata tệp nạp tạm */}
                {newSnapFile && (
                  <div className="bg-amber-50/60 border border-amber-200 rounded-xl p-4 space-y-3 text-slate-800">
                    <div className="flex items-start gap-2">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 mt-0.5" />
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-slate-800 block leading-tight">Đã tải tạm tệp tin thành công!</span>
                        <span className="text-[10px] text-slate-500 block font-mono">Tên file: {newSnapFileName} ({newSnapFile.length} dòng)</span>
                      </div>
                    </div>

                    <div className="space-y-3.5 pt-2 border-t border-amber-200">
                      <div>
                        <label className="text-[10px] font-black text-slate-600 uppercase block mb-1">Tên gợi nhớ của Snapshot:</label>
                        <input
                          type="text"
                          value={newSnapName}
                          onChange={(e) => setNewSnapName(e.target.value)}
                          placeholder="Ví dụ: Cá thể quý 1, Doanh nghiệp lớn tháng 5..."
                          className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-emerald-500 font-bold"
                        />
                      </div>

                      {/* PHÂN LOẠI MỤC LƯU TRỮ */}
                      <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-600 uppercase block">Thư mục / Mục lưu trữ dữ liệu:</label>
                        <div className="grid grid-cols-1 gap-2">
                          <select
                            value={newSnapCategory}
                            onChange={(e) => setNewSnapCategory(e.target.value)}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-emerald-500 font-semibold"
                          >
                            {uniqueCategories.map((cat) => (
                              <option key={cat} value={cat}>📁 {cat}</option>
                            ))}
                            <option value="custom">➕ Tạo mục lưu trữ mới...</option>
                          </select>

                          {newSnapCategory === "custom" && (
                            <div className="animate-fade-in space-y-1">
                              <input
                                type="text"
                                value={newCustomCategory}
                                onChange={(e) => setNewCustomCategory(e.target.value)}
                                placeholder="Nhập tên mục mới (Ví dụ: Lưu Cá Thể, Lưu Doanh Nghiệp...)"
                                className="w-full bg-white border border-emerald-300 rounded-lg px-2.5 py-1.5 text-xs text-emerald-800 focus:ring-1 focus:ring-emerald-500 font-bold animate-pulse"
                              />
                              <p className="text-[9px] text-slate-400">Các tệp tin lưu vào cùng mục này sẽ được gộp chung vào một thư mục trực quan.</p>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* CHU KỲ BÁO CÁO */}
                      <div className="space-y-1.5 bg-white p-2.5 rounded-lg border border-slate-200 shadow-xs">
                        <label className="text-[10px] font-black text-slate-500 uppercase block mb-1">Chu kỳ đối chiếu:</label>
                        <div className="flex gap-1.5 mb-2">
                          {(["month", "quarter", "year"] as const).map((type) => (
                            <button
                              key={type}
                              type="button"
                              onClick={() => {
                                setNewSnapCycleType(type);
                                if (type === "quarter") {
                                  setNewSnapCycleValue(1);
                                } else if (type === "month") {
                                  setNewSnapCycleValue(new Date().getMonth() + 1);
                                }
                              }}
                              className={`flex-1 text-[10px] font-bold py-1 px-1.5 rounded-md border transition-all ${
                                newSnapCycleType === type
                                  ? "bg-emerald-50 border-emerald-400 text-emerald-700 shadow-xs"
                                  : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                              }`}
                            >
                              {type === "month" ? "Theo Tháng" : type === "quarter" ? "Theo Quý" : "Theo Năm"}
                            </button>
                          ))}
                        </div>

                        <div className="grid grid-cols-2 gap-2">
                          {newSnapCycleType === "month" && (
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block mb-0.5">Chọn Tháng:</label>
                              <select
                                value={newSnapCycleValue}
                                onChange={(e) => setNewSnapCycleValue(parseInt(e.target.value))}
                                className="w-full bg-slate-50 border border-slate-300 rounded-md px-2 py-1 text-xs font-semibold text-slate-800"
                              >
                                {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
                                  <option key={m} value={m}>Tháng {m}</option>
                                ))}
                              </select>
                            </div>
                          )}

                          {newSnapCycleType === "quarter" && (
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block mb-0.5">Chọn Quý:</label>
                              <select
                                value={newSnapCycleValue}
                                onChange={(e) => setNewSnapCycleValue(parseInt(e.target.value))}
                                className="w-full bg-slate-50 border border-slate-300 rounded-md px-2 py-1 text-xs font-semibold text-slate-800"
                              >
                                {Array.from({ length: 4 }, (_, i) => i + 1).map((q) => (
                                  <option key={q} value={q}>Quý {q}</option>
                                ))}
                              </select>
                            </div>
                          )}

                          <div className={newSnapCycleType === "year" ? "col-span-2" : ""}>
                            <label className="text-[9px] font-bold text-slate-500 block mb-0.5 font-mono">Chọn Năm:</label>
                            <select
                              value={newSnapYear}
                              onChange={(e) => setNewSnapYear(parseInt(e.target.value))}
                              className="w-full bg-slate-50 border border-slate-300 rounded-md px-2 py-1 text-xs font-semibold text-slate-800"
                            >
                              {Array.from({ length: 11 }, (_, i) => 2020 + i).map((y) => (
                                <option key={y} value={y}>Năm {y}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                      </div>

                      <button
                        onClick={handleSaveSnapshot}
                        className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2.5 rounded-lg border-b-2 border-emerald-800 transition-all cursor-pointer flex items-center justify-center gap-1.5 mt-2"
                      >
                        <Save className="w-3.5 h-3.5" /> LƯU VÀO KHO LƯU TRỮ VĨNH VIỄN
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Hướng dẫn ngắn */}
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-6 space-y-4 text-slate-800 shadow-sm">
                <span className="text-xs font-black text-indigo-750 uppercase tracking-wider font-mono block flex items-center gap-1.5">
                  💡 QUY TRÌNH 4 BƯỚC THAO TÁC SO SÁNH LIÊN KỲ:
                </span>
                
                <div className="space-y-4 relative pl-3 border-l border-indigo-200 ml-1.5">
                  <div className="relative">
                    <div className="absolute -left-[17px] top-0.5 bg-indigo-600 text-white font-mono text-[9px] font-bold w-4 h-4 rounded-full flex items-center justify-center shadow-sm">1</div>
                    <p className="text-xs font-bold text-slate-800 leading-none">Bước 1: Nạp dữ liệu nguồn</p>
                    <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                      Kéo thả file Excel/CSV bất kỳ vào khung phía trên, hoặc nhấn nút <b>"Lấy từ Dữ liệu đang chạy ngoài màn hình chính"</b> để nạp bảng đang làm việc.
                    </p>
                  </div>

                  <div className="relative">
                    <div className="absolute -left-[17px] top-0.5 bg-indigo-600 text-white font-mono text-[9px] font-bold w-4 h-4 rounded-full flex items-center justify-center shadow-sm">2</div>
                    <p className="text-xs font-bold text-slate-800 leading-none">Bước 2: Gán Kỳ báo cáo & Lưu kho</p>
                    <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                      Kiểm tra thông tin, đặt tên gợi nhớ dễ hiểu (ví dụ: <i>Doanh thu tháng 5</i>), chọn đúng <b>Tháng</b> và <b>Năm báo cáo</b>, sau đó nhấn <b>"LƯU VÀO KHO LƯU TRỮ VĨNH VIỄN"</b>.
                    </p>
                  </div>

                  <div className="relative">
                    <div className="absolute -left-[17px] top-0.5 bg-indigo-600 text-white font-mono text-[9px] font-bold w-4 h-4 rounded-full flex items-center justify-center shadow-sm">3</div>
                    <p className="text-xs font-bold text-slate-800 leading-none">Bước 3: Chọn kỳ đối chiếu & Ánh xạ</p>
                    <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                      Chọn 2 kỳ bạn cần đối chiếu ở mục 3 (<b>Bản gốc</b> đại diện kỳ trước, <b>Bản so khớp</b> đại diện kỳ này). Hệ thống tự động gợi ý chọn Cột khóa định danh và Cột chỉ tiêu muốn rà soát.
                    </p>
                  </div>

                  <div className="relative">
                    <div className="absolute -left-[17px] top-0.5 bg-indigo-600 text-white font-mono text-[9px] font-bold w-4 h-4 rounded-full flex items-center justify-center shadow-sm">4</div>
                    <p className="text-xs font-bold text-slate-800 leading-none">Bước 4: Chạy rà soát & Xuất Excel</p>
                    <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                      Nhấn <b>"TIẾN HÀNH ĐỐI CHIẾU 2 KỲ BÁO CÁO"</b> để xem bảng chênh lệch chi tiết, tỷ lệ % tăng trưởng cùng trạng thái phân loại trực quan, hỗ trợ kết xuất tệp Excel sạch ngay sau khi chạy xong.
                    </p>
                  </div>
                </div>
              </div>

            </div>

            {/* Cột phải (Danh sách kho lưu trữ và giao diện so sánh đối chiếu) */}
            <div className="lg:col-span-7 space-y-6">
              
              {/* Danh sách Kho dữ liệu */}
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-5 space-y-4 text-slate-800">
                <h3 className="text-sm font-extrabold text-slate-800 tracking-wider uppercase font-mono flex items-center gap-2">
                  <FolderOpen className="w-5 h-5 text-amber-500" /> 2. Danh sách kho dữ liệu lưu trữ ({savedSnapshots.length})
                </h3>

                {/* Reassurance Security Banner */}
                <div className="bg-emerald-50/50 border border-emerald-150 rounded-xl p-3 flex items-start gap-3">
                  <span className="text-xl">🛡️</span>
                  <div className="space-y-1">
                    <h4 className="text-xs font-bold text-emerald-800 leading-none">Lưu trữ Bảo mật tại chỗ (An toàn 100%)</h4>
                    <p className="text-[10.5px] text-emerald-700/90 leading-relaxed">
                      Dữ liệu được lưu trực tiếp vào <b>bộ nhớ cục bộ IndexedDB bảo mật của Trình duyệt</b> trên thiết bị này. 
                      Hệ thống <b>hoàn toàn không tải dữ liệu lên Máy chủ, Đám mây (Firebase) hay Đĩa cứng của bạn</b>, đảm bảo bảo mật tuyệt đối và không phát sinh bất kỳ đường dẫn vật lý nào xâm phạm hệ thống.
                    </p>
                  </div>
                </div>

                {savedSnapshots.length === 0 ? (
                  <div className="border border-dashed border-slate-200 rounded-xl p-8 text-center text-slate-400 space-y-1">
                    <AlertCircle className="w-8 h-8 text-slate-300 mx-auto" />
                    <p className="text-xs font-bold">Kho lưu trữ đang trống</p>
                    <p className="text-[10px]">Vui lòng nạp tệp tin và lưu trữ ở khung bên trái để bắt đầu tích lũy dữ liệu.</p>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {(Object.entries(groupedSnapshots) as [string, SavedSnapshot[]][]).map(([catName, snaps]) => {
                      const isCollapsed = collapsedFolders[catName] ?? false;
                      return (
                        <div key={catName} className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                          {/* Folder Header */}
                          <button
                            type="button"
                            onClick={() => setCollapsedFolders(prev => ({ ...prev, [catName]: !isCollapsed }))}
                            className="w-full flex items-center justify-between bg-slate-50 hover:bg-slate-100/80 px-4 py-3 text-left transition-colors border-b border-slate-200"
                          >
                            <div className="flex items-center gap-2.5">
                              <span className="text-amber-500 font-bold text-lg">📁</span>
                              <div>
                                <span className="font-extrabold text-xs text-slate-800 uppercase tracking-wider">{catName}</span>
                                <span className="text-[10px] text-slate-500 ml-2 font-mono bg-slate-200/60 px-1.5 py-0.5 rounded-full">{snaps.length} bản ghi</span>
                              </div>
                            </div>
                            <span className="text-slate-400 text-xs font-bold font-mono">
                              {isCollapsed ? "Mở rộng ▽" : "Thu gọn △"}
                            </span>
                          </button>

                          {/* Folder Contents */}
                          {!isCollapsed && (
                            <div className="overflow-x-auto max-h-60 overflow-y-auto">
                              <table className="w-full text-left border-collapse text-xs">
                                <thead>
                                  <tr className="bg-slate-50/40 border-b border-slate-150 font-mono text-[9px] font-bold text-slate-500">
                                    <th className="px-3 py-2">Chu kỳ</th>
                                    <th className="px-3 py-2">Tên gợi nhớ / Tệp gốc</th>
                                    <th className="px-3 py-2 text-right">Số dòng</th>
                                    <th className="px-3 py-2 text-center">Thao tác</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-100">
                                  {snaps.map((snap) => (
                                    <tr key={snap.id} className="hover:bg-slate-50/30 transition-all text-slate-700">
                                      <td className="px-3 py-2.5 font-bold">
                                        <span className="bg-indigo-50 border border-indigo-100 text-indigo-700 text-[10px] font-mono font-bold px-2 py-0.5 rounded">
                                          {getFormattedPeriod(snap)}
                                        </span>
                                      </td>
                                      <td className="px-3 py-2.5">
                                        <span className="font-bold block text-slate-800">{snap.name}</span>
                                        <span className="text-[9px] text-slate-400 block font-mono truncate max-w-[200px]" title={snap.fileName}>
                                          {snap.fileName}
                                        </span>
                                      </td>
                                      <td className="px-3 py-2.5 text-right font-mono font-bold text-slate-600">
                                        {snap.data.length.toLocaleString()}
                                      </td>
                                      <td className="px-3 py-2.5">
                                        <div className="flex items-center justify-center gap-1.5">
                                          <button
                                            onClick={() => handleLoadSnapshotAsActive(snap)}
                                            title="Nạp vào màn hình chính để phân tích hoặc chỉnh sửa"
                                            className="bg-sky-50 text-sky-700 hover:bg-sky-100 hover:text-sky-800 font-bold text-[10px] px-2 py-1 rounded transition-all cursor-pointer"
                                          >
                                            Nạp bảng chính
                                          </button>
                                          <button
                                            onClick={() => handleDeleteSnapshot(snap.id, snap.name)}
                                            title="Xóa vĩnh viễn khỏi kho"
                                            className="text-red-500 hover:bg-red-50 hover:text-red-700 p-1 rounded transition-all cursor-pointer"
                                          >
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        </div>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Giao diện Đối chiếu kỳ */}
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-5 space-y-4 text-slate-800">
                <h3 className="text-sm font-extrabold text-slate-800 tracking-wider uppercase font-mono flex items-center gap-2">
                  <TrendingUp className="w-5 h-5 text-sky-500 animate-pulse" /> 3. Đối chiếu liên kỳ động giữa các tháng/năm
                </h3>
                <p className="text-xs text-slate-500 leading-relaxed">
                  Chọn 2 bản ghi từ kho dữ liệu, cấu hình các cột tương khớp và bấm bắt đầu để so sánh sự thay đổi doanh số/lao động giữa 2 kỳ.
                </p>

                {savedSnapshots.length < 2 ? (
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center text-slate-500 text-xs font-mono">
                    ⚠️ Cần ít nhất 2 Bản ghi lưu trữ trong Kho để thực hiện chức năng so sánh đối chiếu này.
                  </div>
                ) : (
                  <div className="space-y-4">
                    {/* Select Snapshots */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Bản A - Kỳ trước */}
                      <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-3">
                        <span className="text-[10px] font-black text-slate-400 uppercase font-mono block">A. Bản ghi gốc (Kỳ trước/Năm ngoái):</span>
                        <select
                          value={snapAId}
                          onChange={(e) => setSnapAId(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-800"
                        >
                          <option value="">-- Chọn Snapshot gốc --</option>
                          {(Object.entries(groupedSnapshots) as [string, SavedSnapshot[]][]).map(([catName, snaps]) => (
                            <optgroup key={catName} label={`📂 ${catName.toUpperCase()}`}>
                              {snaps.map((s) => (
                                <option key={s.id} value={s.id}>
                                  [{getFormattedPeriod(s)}] - {s.name} ({s.data.length} dòng)
                                </option>
                              ))}
                            </optgroup>
                          ))}
                        </select>

                        {snapAId && (
                          <div className="grid grid-cols-2 gap-2 pt-1">
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block">Cột khóa định danh:</label>
                              <select
                                value={snapAKeyCol}
                                onChange={(e) => setSnapAKeyCol(e.target.value)}
                                className="w-full bg-white border border-slate-300 rounded-md px-1.5 py-1 text-[11px]"
                              >
                                {savedSnapshots.find(s => s.id === snapAId)?.columns.map((col) => (
                                  <option key={col} value={col}>{col}</option>
                                ))}
                              </select>
                            </div>
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block">Cột chỉ tiêu số:</label>
                              <select
                                value={snapAMetricCol}
                                onChange={(e) => setSnapAMetricCol(e.target.value)}
                                className="w-full bg-white border border-slate-300 rounded-md px-1.5 py-1 text-[11px]"
                              >
                                {savedSnapshots.find(s => s.id === snapAId)?.columns.map((col) => (
                                  <option key={col} value={col}>{col}</option>
                                ))}
                              </select>
                            </div>
                          </div>
                        )}
                      </div>

                      {/* Bản B - Kỳ này */}
                      <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-3">
                        <span className="text-[10px] font-black text-slate-400 uppercase font-mono block">B. Bản ghi so khớp (Kỳ này/Năm nay):</span>
                        <select
                          value={snapBId}
                          onChange={(e) => setSnapBId(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-800"
                        >
                          <option value="">-- Chọn Snapshot so sánh --</option>
                          {(Object.entries(groupedSnapshots) as [string, SavedSnapshot[]][]).map(([catName, snaps]) => (
                            <optgroup key={catName} label={`📂 ${catName.toUpperCase()}`}>
                              {snaps.map((s) => (
                                <option key={s.id} value={s.id}>
                                  [{getFormattedPeriod(s)}] - {s.name} ({s.data.length} dòng)
                                </option>
                              ))}
                            </optgroup>
                          ))}
                        </select>

                        {snapBId && (
                          <div className="grid grid-cols-2 gap-2 pt-1">
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block">Cột khóa định danh:</label>
                              <select
                                value={snapBKeyCol}
                                onChange={(e) => setSnapBKeyCol(e.target.value)}
                                className="w-full bg-white border border-slate-300 rounded-md px-1.5 py-1 text-[11px]"
                              >
                                {savedSnapshots.find(s => s.id === snapBId)?.columns.map((col) => (
                                  <option key={col} value={col}>{col}</option>
                                ))}
                              </select>
                            </div>
                            <div>
                              <label className="text-[9px] font-bold text-slate-500 block">Cột chỉ tiêu số:</label>
                              <select
                                value={snapBMetricCol}
                                onChange={(e) => setSnapBMetricCol(e.target.value)}
                                className="w-full bg-white border border-slate-300 rounded-md px-1.5 py-1 text-[11px]"
                              >
                                {savedSnapshots.find(s => s.id === snapBId)?.columns.map((col) => (
                                  <option key={col} value={col}>{col}</option>
                                ))}
                              </select>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Compare Button */}
                    <button
                      onClick={handleCompareTwoSnapshots}
                      className="w-full bg-gradient-to-r from-sky-500 to-sky-600 hover:from-sky-600 hover:to-sky-700 text-white font-bold text-xs px-6 py-3.5 rounded-xl border-b-4 border-sky-800 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
                    >
                      <TrendingUp className="w-4 h-4 animate-bounce" /> TIẾN HÀNH ĐỐI CHIẾU 2 KỲ BÁO CÁO
                    </button>
                  </div>
                )}
              </div>

            </div>

          </div>
        </div>
      )}

      {/* HIỂN THỊ KẾT QUẢ ĐỐI CHIẾU KHO LƯU TRỮ */}
      {subMode === "saved-snapshots" && snapCompareResult && snapCompareResult.length > 0 && (
        <div className="space-y-6">
          
          {/* Dashboard phân tích nhanh */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-slate-800">
            
            <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col justify-between">
              <span className="text-[10px] font-black text-slate-400 uppercase font-mono">Tổng số đơn vị khớp:</span>
              <span className="text-2xl font-black text-slate-800 font-mono mt-1">{snapCompareResult.length.toLocaleString()}</span>
              <span className="text-[10px] text-slate-400 mt-2 block">Tổng số liên kết được tìm thấy</span>
            </div>

            <div className="bg-emerald-50/50 border border-emerald-100 rounded-2xl p-4 shadow-xs flex flex-col justify-between">
              <span className="text-[10px] font-black text-emerald-600 uppercase font-mono">Tăng trưởng dương (📈):</span>
              <span className="text-2xl font-black text-emerald-700 font-mono mt-1">
                {snapCompareResult.filter(r => r["Trạng thái"] === "📈 Tăng").length.toLocaleString()}
              </span>
              <span className="text-[10px] text-emerald-600/80 mt-2 block">Số lượng đơn vị tăng chỉ tiêu</span>
            </div>

            <div className="bg-red-50/50 border border-red-100 rounded-2xl p-4 shadow-xs flex flex-col justify-between">
              <span className="text-[10px] font-black text-red-600 uppercase font-mono">Tăng trưởng âm (📉):</span>
              <span className="text-2xl font-black text-red-700 font-mono mt-1">
                {snapCompareResult.filter(r => r["Trạng thái"] === "📉 Giảm").length.toLocaleString()}
              </span>
              <span className="text-[10px] text-red-600/80 mt-2 block">Số lượng đơn vị giảm chỉ tiêu</span>
            </div>

            <div className="bg-sky-50/50 border border-sky-100 rounded-2xl p-4 shadow-xs flex flex-col justify-between">
              <span className="text-[10px] font-black text-sky-600 uppercase font-mono">Mới xuất hiện (✅):</span>
              <span className="text-2xl font-black text-sky-700 font-mono mt-1">
                {snapCompareResult.filter(r => r["Trạng thái"] === "✅ Mới thêm").length.toLocaleString()}
              </span>
              <span className="text-[10px] text-sky-600/80 mt-2 block">Đơn vị mới bổ sung kỳ này</span>
            </div>

          </div>

          {/* Bảng dữ liệu chính */}
          <MainDataInlinePreview
            data={snapCompareResult}
            columns={Object.keys(snapCompareResult[0] || {})}
            title="KẾT QUẢ ĐỐI CHIẾU KHO LƯU TRỮ LIÊN KỲ"
            subtitle={`Đã hoàn tất phân tích đối chiếu chéo! Tìm thấy ${snapCompareResult.length} đơn vị có khóa liên kết tương ứng.`}
            onExportExcel={onExportExcel}
          />
        </div>
      )}

      {/* HIỂN THỊ KẾT QUẢ SO SÁNH NIÊN ĐỘ TRUYỀN THỐNG */}
      {subMode === "diff" && compareResultData && compareResultData.length > 0 && (
        <MainDataInlinePreview
          data={compareResultData}
          columns={Object.keys(compareResultData[0] || {})}
          title="KẾT QUẢ SO SÁNH ĐỐI CHIẾU HAI NIÊN ĐỘ"
          subtitle={`Đã so sánh đối chiếu thành công! Tìm thấy tổng số: ${compareResultData.length} dòng dữ liệu khóa liên kết với trạng thái thay đổi tương ứng.`}
          onExportExcel={onExportExcel}
        />
      )}

      {/* HIỂN THỊ KẾT QUẢ SO SÁNH CHUỖI THỜI GIAN ĐỘNG */}
      {subMode === "time-series" && tsResultRows.length > 0 && (
        <div className="space-y-6">
          {/* Biểu đồ động đầy đủ */}
          <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-4">
            <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2 uppercase font-mono">
              <BarChart2 className="w-4 h-4 text-amber-500 animate-pulse" /> Biểu đồ so sánh tăng trưởng - Địa bàn: {selectedChartEntity}
            </h4>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartDataForSelected} margin={{ top: 20, right: 10, left: -10, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} />
                  <XAxis dataKey="name" />
                  <YAxis />
                  <Tooltip />
                  <Legend />
                  <Bar dataKey="Kỳ này" fill="#f59e0b" name="Số liệu Kỳ này" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Kỳ trước" fill="#94a3b8" name="Số liệu Kỳ trước" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <MainDataInlinePreview
            data={tsResultRows.map(row => {
              // Loại bỏ các trường ẩn phục vụ vẽ biểu đồ trước khi đưa ra preview
              const { _year, _month, ...previewRow } = row;
              return previewRow;
            })}
            columns={Object.keys(tsResultRows[0] || {}).filter(k => !k.startsWith("_"))}
            title={`BẢNG SO SÁNH CHI TIẾT CHUỖI THỜI GIAN (${comparisonType === "adjacent" ? "LIÊN TIẾP" : "CÙNG KỲ NĂM TRƯỚC"})`}
            subtitle={`Đã xử lý chuỗi thời gian hoàn tất! Tìm thấy tổng số: ${tsResultRows.length} điểm đối chiếu liên kết có chênh lệch tăng trưởng.`}
            onExportExcel={onExportExcel}
          />
        </div>
      )}

      {/* XEM TRƯỚC FILE NGUỒN CŨ (Trong Mode Diff) */}
      {subMode === "diff" && oldData.length > 0 && (
        <MainDataInlinePreview
          data={oldData}
          columns={Object.keys(oldData[0] || {})}
          title="DỮ LIỆU NGUỒN CŨ"
          subtitle="Xem trước bảng niên độ cũ đang chuẩn bị đem so sánh."
        />
      )}
    </div>
  );
}
