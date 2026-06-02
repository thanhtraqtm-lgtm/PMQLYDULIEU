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
        resolve(request.result || null);
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
  Database,
  RefreshCw,
  Search,
  Plus,
  Trash2,
  FileCheck
} from "lucide-react";

import { 
  vsicRawData, 
  normalizeSectorCode, 
  getSectorHierarchy, 
  smartSuggestSectorByDescription,
  getSectorLevel,
  getParentSectorCode
} from "./data/vsic";

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

  // Storage chính
  const [mainData, setMainData] = useState<any[]>([]);
  const [rawImportedData, setRawImportedData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>("");

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

  // Biến trạng thái hiển thị của báo cáo nhanh
  const [quickReportLevel, setQuickReportLevel] = useState<number>(1);

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

  // Thuật toán dọn dẹp dải ô Workbook trống thừa thãi cực kỳ tối ưu và an toàn
  const optimizeSheetRange = (ws: XLSX.WorkSheet) => {
    if (!ws || !ws['!ref']) return;
    
    let range;
    try {
      range = XLSX.utils.decode_range(ws['!ref']);
    } catch (e) {
      return;
    }

    const rowCount = range.e.r - range.s.r + 1;
    
    // Nếu số dòng dưới 30,000, Object.keys(ws) hoạt động cực nhanh và an toàn
    if (rowCount < 30000) {
      let minRow = Infinity, maxRow = -Infinity;
      let minCol = Infinity, maxCol = -Infinity;
      
      const cellKeys = Object.keys(ws).filter(k => k[0] !== '!');
      if (cellKeys.length === 0) return;
      
      cellKeys.forEach(key => {
        try {
          const cell = XLSX.utils.decode_cell(key);
          const cellObj = ws[key];
          if (cellObj && cellObj.v !== undefined && cellObj.v !== null && String(cellObj.v).trim() !== "") {
            if (cell.r < minRow) minRow = cell.r;
            if (cell.r > maxRow) maxRow = cell.r;
            if (cell.c < minCol) minCol = cell.c;
            if (cell.c > maxCol) maxCol = cell.c;
          }
        } catch (e) {}
      });
      
      if (minRow !== Infinity) {
        ws['!ref'] = XLSX.utils.encode_range({
          s: { r: minRow, c: minCol },
          e: { r: maxRow, c: maxCol }
        });
      }
      return;
    }

    // Với các tệp siêu lớn (30,000+ dòng đến 1 triệu dòng):
    // KHÔNG dùng Object.keys(ws) vì sẽ lỗi "Too many properties to enumerate" hoặc làm đơ trình duyệt.
    // Chúng ta tối ưu bằng cách quét từ cuối hàng (hàng cuối cùng trong dải ô) ngược lên để phát hiện dữ liệu thực tế.
    
    const startCol = range.s.c;
    const endCol = Math.min(range.e.c, 100); // Giới hạn kiểm tra 100 cột đầu tiên để đảm bảo tốc độ
    
    const checkRowHasData = (r: number): boolean => {
      for (let c = startCol; c <= endCol; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cellObj = ws[cellRef];
        if (cellObj && cellObj.v !== undefined && cellObj.v !== null && String(cellObj.v).trim() !== "") {
          return true;
        }
      }
      return false;
    };

    // Kiểm tra nhanh xem dải ô có thực sự bị phình to (bloated) không.
    // Nếu kiểm tra vài dòng ở cuối cùng (ví dụ dòng range.e.r) mà có dữ liệu -> đây là tệp lớn thật sự, dải ô hoàn toàn chuẩn xác!
    let isGenuineLargeFile = false;
    for (let r = Math.max(range.s.r, range.e.r - 20); r <= range.e.r; r++) {
      if (checkRowHasData(r)) {
        isGenuineLargeFile = true;
        break;
      }
    }
    
    // Nếu tệp lớn thật sự, chúng ta giữ nguyên dải ô (!ref) và thoát ngay để tránh tính toán thừa.
    if (isGenuineLargeFile) {
      return;
    }

    // Nếu dải ô bị phình to vô lý (cuối dải ô trống hoàn toàn), tiến hành dò tìm hàng cuối cùng có dữ liệu ngược lên.
    // Sử dụng thuật toán nhảy bước để tìm hàng cuối có dữ liệu siêu nhanh mà không tốn tài nguyên.
    let realLastRow = range.s.r;
    let step = 10000;
    let currentHigh = range.e.r;
    
    while (currentHigh > range.s.r) {
      let foundDataInStep = false;
      const checkLimit = Math.max(range.s.r, currentHigh - step);
      for (let r = currentHigh; r >= checkLimit; r -= Math.max(1, Math.floor(step / 10))) {
        if (checkRowHasData(r)) {
          realLastRow = r;
          foundDataInStep = true;
          break;
        }
      }
      if (foundDataInStep) {
        for (let r = realLastRow; r <= Math.min(range.e.r, realLastRow + step); r++) {
          if (checkRowHasData(r)) {
            realLastRow = r;
          }
        }
        break;
      }
      currentHigh -= step;
    }

    // Cập nhật lại dải ô mới an toàn
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: range.s.r, c: range.s.c },
      e: { r: realLastRow, c: range.e.c }
    });
  };

  // Mock data generator
  const loadMockData = () => {
    const mock = [
      { STT: "1", MaST: "0100021312", TenDN: "Công ty Cổ phần Lúa Gạo Việt Nam", MoTa: "Hoạt động trồng lúa nước chất lượng cao và chăn nuôi heo thịt", MaNganhDTV: "1110", Xa: "Xã Mỹ Lộc", DoanhThu: "4500", LaoDong: "32" },
      { STT: "2", MaST: "0200843213", TenDN: "Đại lý Thương mại Lộc Phát", MoTa: "Bán buôn lúa gạo hữu cơ, bán lẻ dầu ăn bánh kẹo sữa", MaNganhDTV: "46321", Xa: "Xã An Hòa", DoanhThu: "1280", LaoDong: "5" },
      { STT: "3", MaST: "0301132345", TenDN: "Công ty Đầu tư Xây dựng Hoàn Mỹ", MoTa: "Thi công công trình xây dựng nhà ở dân dụng các loại", MaNganhDTV: "4100", Xa: "Xã Mỹ Lộc", DoanhThu: "8900", LaoDong: "120" },
      { STT: "4", MaST: "0401833441", TenDN: "Cơ sở Khai thác Đá Quý An Bình", MoTa: "Khai thác cát sỏi đất sét làm vật liệu xây dựng", MaNganhDTV: "811", Xa: "Xã Mỹ Lộc", DoanhThu: "620", LaoDong: "18" },
      { STT: "5", MaST: "0502123984", TenDN: "Phòng Khám Đa Khoa Sức Khỏe Vàng", MoTa: "Đại lý bán buôn thiết bị y tế dân dụng gia đình", MaNganhDTV: "4659", Xa: "Xã An Hòa", DoanhThu: "3100", LaoDong: "12" },
      { STT: "6", MaST: "0601243124", TenDN: "Nhà hàng Trúc Lâm Quán", MoTa: "Dịch vụ nhà hàng ăn uống phục vụ lưu động du khách", MaNganhDTV: "56100", Xa: "Xã Tân Bình", DoanhThu: "1800", LaoDong: "25" },
      { STT: "7", MaST: "0701982736", TenDN: "Cơ sở dệt may Hoàng Gia", MoTa: "Chăn nuôi trâu bò và trồng ngô sắn gia đình tự tiêu", MaNganhDTV: "98100", Xa: "Xã Tân Bình", DoanhThu: "450", LaoDong: "2" }
    ];
    setRawImportedData(mock);
    setMainData(mock);
    setColumns(Object.keys(mock[0]));
    setFileName("Du_Lieu_Doanh_Nghiep_Mau.xlsx");

    // Khởi tạo danh sách cấu hình cột động
    const initConfigs = Object.keys(mock[0]).map(c => {
      let role: "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" | "" = "";
      if (c === "MoTa") role = "mota";
      else if (c === "MaNganhDTV") role = "manganh";
      else if (c === "Xa") role = "xa";
      else if (c === "DoanhThu") role = "doanhthu";
      else if (c === "LaoDong") role = "laodong";
      else if (c === "MaST") role = "idCol";
      
      return {
        originalName: c,
        use: true,
        newName: c === "MoTa" ? "Mô Tả Hoạt Động" :
                 c === "MaNganhDTV" ? "Mã Ngành Đăng Ký" :
                 c === "Xa" ? "Địa bàn (Xã)" :
                 c === "DoanhThu" ? "Doanh Thu" :
                 c === "LaoDong" ? "Tổng số Lao Động" :
                 c === "MaST" ? "Mã Số Thuế" : c,
        role
      };
    });
    setCustomColConfigs(initConfigs);

    // Tự động suy đoán Mapping cột ban đầu
    setMapping({
      mota: "Mô Tả Hoạt Động",
      manganh: "Mã Ngành Đăng Ký",
      xa: "Địa bàn (Xã)",
      doanhthu: "Doanh Thu",
      laodong: "Tổng số Lao Động",
      idCol: "Mã Số Thuế"
    });
    setActiveTab("xemdulieu");
  };

  // Đọc file CSV hoặc Excel bằng xlsx
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "main" | "old" | "new" | "left" | "right") => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang tải tệp: ${file.name}...`);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        if (!arrayBuffer) {
          throw new Error("Không thể đọc nội dung tệp tin!");
        }

        const wb = XLSX.read(arrayBuffer, { 
          type: "array",
          cellFormula: false,
          cellHTML: false,
          cellStyles: false
        });

        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];

        // Tối ưu hóa dải ô để tránh quá tải bộ nhớ và lỗi "Too many properties to enumerate"
        optimizeSheetRange(ws);

        const data = XLSX.utils.sheet_to_json(ws);

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

          // Khởi tạo danh sách cấu hình cột động từ tệp vừa nạp
          const initConfigs = cols.map(c => {
            const low = c.toLowerCase();
            let role: "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol" | "" = "";
            if (low.includes("mô tả") || low.includes("mota") || low.includes("hoatdong")) role = "mota";
            if (low.includes("mã ngành") || low.includes("manganh") || low.includes("ma_nganh") || low.includes("dtv") || low.includes("vsic")) role = "manganh";
            if (low.includes("xã") || low.includes("xa") || low.includes("diaban") || low.includes("dia_ban")) role = "xa";
            if (low.includes("doanh thu") || low.includes("doanhthu") || low.includes("thu_nhap")) role = "doanhthu";
            if (low.includes("lao động") || low.includes("laodong") || low.includes("nhan_su")) role = "laodong";
            if (low.includes("mst") || low.includes("mã số") || low.includes("ident") || low.includes("id")) role = "idCol";

            return {
              originalName: c,
              use: true,
              newName: c, // giữ nguyên tên ban đầu, cho phép người dùng sửa đổi trực tiếp
              role
            };
          });
          setCustomColConfigs(initConfigs);

          // Thử tìm tự động mapping từ tên cột tương tự
          const autoMap: ColumnMapping = { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" };
          initConfigs.forEach(cfg => {
            if (cfg.role) {
              autoMap[cfg.role] = cfg.newName;
            }
          });
          setMapping(autoMap);
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
        alert("Lỗi khi đọc file: " + err.message);
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
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

    // Cập nhật cấu hình mapping chỉ tay từ hệ thống tới các tên cột mới định nghĩa
    const newMapping: ColumnMapping = {
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: ""
    };

    activeConfigs.forEach(cfg => {
      if (cfg.role) {
        newMapping[cfg.role] = cfg.newName.trim();
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
    if (!mainData.length || !mapping.mota || !mapping.manganh) {
      return { descToCodes: [], codeToDescs: [] };
    }

    // 1. CÙNG MÔ TẢ -> KHÁC MÃ LIÊN KẾT
    const descMap = new Map<string, Array<{ code: string; rowIdx: number; row: any }>>();
    mainData.forEach((row, idx) => {
      const rawMota = String(row[mapping.mota] || "").trim();
      const cleanMota = rawMota.toLowerCase().replace(/\s+/g, " ");
      if (!cleanMota) return;
      const code = normalizeSectorCode(row[mapping.manganh]);
      
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
      const codeCounts = new Map<string, number[]>();
      occurrences.forEach(occ => {
        if (!codeCounts.has(occ.code)) {
          codeCounts.set(occ.code, []);
        }
        codeCounts.get(occ.code)!.push(occ.rowIdx);
      });

      if (codeCounts.size > 1) {
        const originalText = occurrences[0].row[mapping.mota] || cleanMota;
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
    });

    // 2. CÙNG MÃ -> MÔ TẢ TRÁI QUY LUẬT KINH DOANH (Sản xuất vs Thương mại)
    const codeMap = new Map<string, Array<{ desc: string; rowIdx: number; row: any }>>();
    mainData.forEach((row, idx) => {
      const code = normalizeSectorCode(row[mapping.manganh]);
      if (!code) return;
      const rawMota = String(row[mapping.mota] || "").trim();
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

    codeMap.forEach((items, code) => {
      const tradeKeywords = ["bán", "mua", "thương mại", "đại lý", "cửa hàng", "phân phối", "wholesale", "retail", "shop", "siêu thị", "chợ"];
      const industrialKeywords = ["sản xuất", "gia công", "chế tạo", "làm mộc", "chế biến", "lắp ráp", "chế tác", "luyện kim", "nhà xưởng", "nhà máy"];
      
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
  }, [mainData, mapping.mota, mapping.manganh]);

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

    setMainData(mergedResults);
    setColumns(Object.keys(mergedResults[0] || {}));
    setFileName(`GhepNoi_${leftFileName}_vs_${rightFileName}.xlsx`);
    
    // Auto map idCol
    setMapping(prev => ({ ...prev, idCol: leftKey }));

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
      const k = String(row[diffKey] || "").trim();
      if (k) oldMap.set(k, row);
    });

    const newMap = new Map();
    newData.forEach(row => {
      const k = String(row[diffKey] || "").trim();
      if (k) newMap.set(k, row);
    });

    const resultRows: any[] = [];
    const allKeys = Array.from(new Set([...oldMap.keys(), ...newMap.keys()]));
    const batchSize = Math.max(1, Math.floor(allKeys.length / 20));

    // Lấy tập hợp tất cả các cột của cả hai file
    const oldCols = Object.keys(oldData[0] || {});
    const newCols = Object.keys(newData[0] || {});
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

    setMainData(resultRows);
    setColumns(Object.keys(resultRows[0] || {}));
    setFileName(`SoSanhDiff_${oldFileName}_vs_${newFileName}.xlsx`);
    setMapping(prev => ({ ...prev, idCol: diffKey }));

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
        if (col === "_virtual_sector_cap2" && mapping.manganh) {
          const mng = normalizeSectorCode(row[mapping.manganh]);
          const sec2Code = mng ? mng.slice(0, 2) : "";
          const sec2Name = vsicRawData[sec2Code] || "Ngành cấp 2 chưa định nghĩa";
          compositeKeyObj["Ngành Cấp 2"] = sec2Code ? `${sec2Code} - ${sec2Name}` : "Chưa xác định";
        } else if (col === "_virtual_sector_cap1" && mapping.manganh) {
          const mng = normalizeSectorCode(row[mapping.manganh]);
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

  // 5. CHỨC NĂNG BÁO CÁO NHANH THEO PHÂN CẤP NGÀNH & XÃ CHUẨN XÁC
  const handleQuickReport = async (level: number) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi chạy báo cáo nhanh.");
      return;
    }
    if (!mapping.manganh || !mapping.xa) {
      alert("Yêu cầu định cấu hình cột 'Mã Ngành' và 'Địa bàn (Xã)' ở trang cơ sở đầu vào trước!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang tạo báo cáo nhanh Ngành Cấp ${level} kết hợp Xã...`);
    await sleep(200);

    // Gom dữ liệu mỏng và phân tách bằng bộ nhớ mã ngành chuẩn
    const processedData = mainData.map(row => {
      const mng = normalizeSectorCode(row[mapping.manganh]);
      
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
        _tempXa: String(row[mapping.xa] || "Khác").trim()
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
            if (mapping.doanhthu) {
              const val = parseFloat(String(r[mapping.doanhthu]).replace(/[^0-9.\-]/g, ""));
              if (!isNaN(val)) sumDoanhThu += val;
            }
            if (mapping.laodong) {
              const val = parseFloat(String(r[mapping.laodong]).replace(/[^0-9.\-]/g, ""));
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
          if (mapping.doanhthu) {
            const val = parseFloat(String(r[mapping.doanhthu]).replace(/[^0-9.\-]/g, ""));
            if (!isNaN(val)) sumDoanhThu += val;
          }
          if (mapping.laodong) {
            const val = parseFloat(String(r[mapping.laodong]).replace(/[^0-9.\-]/g, ""));
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

    setMainData(finalReportRows);
    setColumns(Object.keys(finalReportRows[0] || {}));
    setFileName(`BaoCaoDynamic_NganhCap${level}_Va_Xa_${reportType}.xlsx`);

    setProgress(100);
    setStatusMessage(`Tạo báo cáo nhanh ${reportType === "pivot" ? "xoay cột Pivot" : "dạng phẳng"} Ngành Cấp ${level} thành công!`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
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

      // Đọc trong danh mục bộ nhớ chuẩn mã ngành cấp 5 của DTV nhập
      const stdCap5Ten = vsicRawData[maDtvVal] || "";

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

      if (!vsicRawData[maDtvVal]) {
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
    if (mainData.length === 0) {
      alert("Không có dữ liệu để xuất file!");
      return;
    }
    setLoading(true);
    setStatusMessage("Đang tạo tệp Excel phục vụ tải xuống...");

    setTimeout(() => {
      try {
        const ws = XLSX.utils.json_to_sheet(mainData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Tinh_Toan");
        XLSX.writeFile(wb, fileName || "Ket_Qua_Bao_Cao.xlsx");
        setStatusMessage("Đã tải xuống file Excel thành công.");
      } catch (e: any) {
        alert("Lỗi khi kết xuất Excel: " + e.message);
      } finally {
        setLoading(false);
      }
    }, 200);
  };

  return (
    <div className="flex flex-col min-h-screen bg-[#111827] text-gray-100 font-sans selection:bg-purple-600 selection:text-white">
      
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

          <button 
            onClick={loadMockData}
            className="bg-[#374151] hover:bg-[#4b5563] text-gray-200 text-xs font-semibold px-4 py-2 rounded-lg transition-all border border-[#4b5563] flex items-center gap-1.5 cursor-pointer shadow-sm"
          >
            <Database className="w-3.5 h-3.5 text-blue-400" /> Nạp dữ liệu mẫu
          </button>
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

          {/* 1. TAB TRANG CHỦ */}
          {activeTab === "trangchu" && (
            <div className="space-y-8 animate-fade-in">
              <div className="bg-gradient-to-r from-purple-900/40 via-[#1f2937] to-[#1f2937] border border-purple-500/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-center gap-6 justify-between">
                <div className="space-y-3 max-w-2xl">
                  <span className="bg-purple-900/50 border border-purple-500/30 text-purple-400 text-xs font-mono font-bold px-3 py-1 rounded-full uppercase tracking-wider">
                    Phát hành chuẩn mực V38.5
                  </span>
                  <h2 className="text-2xl md:text-3xl font-extrabold text-white tracking-tight">
                    Hệ Thống Phân Tích & Chuẩn Hóa Dữ Liệu Ngành Quốc Gia
                  </h2>
                  <p className="text-gray-300 text-sm leading-relaxed">
                    Công cụ chuyên sâu hỗ trợ thống kê dữ liệu doanh nghiệp, ghép tách tệp lớn, so khớp, rà soát logic đa chỉ tiêu và xử lý liên kết ngành kinh tế Việt Nam (VSIC) tự động áp dụng giải pháp tối ưu hóa cao cấp.
                  </p>
                  <div className="pt-2 flex items-center gap-4">
                    <button 
                      onClick={loadMockData}
                      className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-sm px-6 py-2.5 rounded-xl transition-all shadow-md shadow-purple-900/30 flex items-center gap-2 cursor-pointer"
                    >
                      Bắt đầu nhanh với Dữ liệu mẫu <ArrowRight className="w-4 h-4" />
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
                    <Layers className="w-5 h-5 text-purple-400 animate-pulse" /> BẢN ĐỒ CẨM NANG & KÍCH HOẠT CHỨC NĂNG NHANH
                  </h3>
                  <p className="text-xs text-gray-400 mt-1">
                    Hệ thống tích hợp đầy đủ 7 phân hệ cốt lõi chuyên sâu. Quý khách có thể xem nhanh hướng dẫn và nhấp trực tiếp vào bất kỳ thẻ nào dưới đây để bắt đầu ngay:
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

                  {/* CHỨC NĂNG 6: Chuẩn hóa VSIC & AI */}
                  <div className="bg-[#1f2937]/50 border border-indigo-500/20 rounded-2xl p-6 flex flex-col justify-between space-y-4 hover:border-indigo-500/40 transition-all group">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="bg-indigo-950/50 border border-indigo-500/30 p-2.5 rounded-xl text-indigo-400">
                          <Brain className="w-5 h-5" />
                        </div>
                        <span className="bg-indigo-900/40 text-indigo-300 border border-indigo-500/20 text-[10px] font-bold px-2 py-0.5 rounded-md uppercase font-mono tracking-wider">
                          TRÍ TUỆ NHÂN TẠO
                        </span>
                      </div>
                      <h4 className="text-base font-bold text-white group-hover:text-indigo-300 transition-colors">
                        🧠 Chuẩn Hóa VSIC &amp; AI
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Tự động rã tách mã ngành 1-5 cấp đầy đủ không đứt gãy phả hệ liên kết. Sử dụng thuật toán NLP tinh lọc so sánh ngữ nghĩa văn bản của "Mô tả hoạt động thực tế" so với "Mã ngành" đăng ký để phát hiện và chỉnh lỗi mã sai lệch.
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("chuanhoanganh")}
                      className="w-full bg-[#111827] hover:bg-indigo-900/30 text-indigo-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-indigo-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Chuẩn Hóa & AI <ArrowRight className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  {/* CHỨC NĂNG 7: Cỗ máy kiểm tra logic */}
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
                        Control Panel 🛂 Cỗ Máy Kiểm Tra Logic
                      </h4>
                      <p className="text-xs text-gray-300 leading-relaxed">
                        Thiết lập các quy định ràng buộc điều kiện thông minh dạng <code>NẾU (điều kiện) THÌ PHẢI (điều kiện kia)</code> tùy ý. Bộ máy sẽ lập tức rà soát toàn bộ tệp Excel thu về danh sách hồ sơ bất tuần tự (ví dụ: mô tả có chứa xi-măng nhưng mã ngành lại đăng ký nông nghiệp, hoặc lao động siêu lớn nhưng doanh thu rỗng).
                      </p>
                    </div>
                    <button 
                      onClick={() => setActiveTab("kiemtralogic")}
                      className="w-full bg-[#111827] hover:bg-emerald-900/30 text-emerald-400 hover:text-white font-bold text-xs py-2 rounded-xl transition-all border border-[#374151] hover:border-emerald-500/30 cursor-pointer flex items-center justify-center gap-1"
                    >
                      Mở Cỗ Máy Rà Quét Logic <ArrowRight className="w-3.5 h-3.5" />
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
                    <span className="font-bold text-white block">02. Chuẩn hóa & Rà rà logic</span>
                    <p className="text-gray-400">
                      Sử dụng <strong>🧠 Chuẩn Hóa VSIC & AI</strong> để hoàn thiện liên kết 5 cấp ngành nghề; tiếp theo sử dụng <strong>🛂 Cỗ Máy Logic</strong> thiết lập các quy chuẩn kiểm tra để lọc sạch các bản ghi lỗi hoặc dị thường.
                    </p>
                  </div>
                  <div className="space-y-1">
                    <span className="font-bold text-white block">03. Tổng hợp báo cáo & Chuyển giao</span>
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

                {/* Phần cấu hình định nghĩa lại tên cột theo phong cách của người dùng (CUSTOM RE-DEFINITION GRID) */}
                {rawImportedData.length > 0 && (
                  <div className="bg-[#111827]/90 rounded-2xl p-5 border border-purple-500/20 space-y-5 animate-slide-up">
                    <div className="border-b border-gray-800 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <div className="text-xs font-bold text-purple-400 tracking-wider uppercase font-mono flex items-center gap-1.5">
                          <Database className="w-5 h-5 text-purple-400 animate-pulse" /> ĐỊNH NGHĨA LẠI TÊN CỘT DỄ NHỚ & LỌC CỘT THỪA
                        </div>
                        <p className="text-xs text-gray-400 mt-1">
                          Sửa đổi các từ viết tắt khó nhớ thành tiếng Việt rõ ràng. Cột nào chưa chọn sẽ bị loại khỏi bảng để giữ bộ dữ liệu sạch nhất.
                        </p>
                      </div>
                      
                      <div className="flex gap-2">
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
                      </div>
                    </div>

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
                            <th className="p-3">TÊN CỘT GỐC (KÝ HIỆU TRONG FILE)</th>
                            <th className="p-3">TÊN MỚI DỄ HIỂU, DỄ NHỚ ĐỊNH NGHĨA LẠI</th>
                            <th className="p-3 w-[250px]">VAI TRÒ NGHIỆP VỤ HỆ THỐNG</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-800/60 font-sans">
                          {customColConfigs.map((cfg, idx) => {
                            const isSystemMappable = ["mota", "manganh", "xa", "doanhthu", "laodong", "idCol"].includes(cfg.role);
                            return (
                              <tr 
                                key={cfg.originalName} 
                                className={`transition-colors hover:bg-gray-800/15 ${
                                  cfg.use ? "bg-transparent" : "bg-gray-950/30 opacity-50"
                                }`}
                              >
                                {/* Cột Checkbox Sử Dụng */}
                                <td className="p-3 text-center">
                                  <input 
                                    type="checkbox"
                                    checked={cfg.use}
                                    onChange={(e) => {
                                      const updated = [...customColConfigs];
                                      updated[idx].use = e.target.checked;
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
                                <td className="p-3">
                                  <div className="flex items-center gap-1.5">
                                    <span className="font-mono bg-gray-800 px-2 py-0.5 rounded text-gray-300 border border-gray-700 max-w-[200px] truncate block" title={cfg.originalName}>
                                      {cfg.originalName}
                                    </span>
                                  </div>
                                </td>

                                {/* Input Tên Mới */}
                                <td className="p-3">
                                  <input 
                                    type="text"
                                    value={cfg.newName}
                                    disabled={!cfg.use}
                                    onChange={(e) => {
                                      const updated = [...customColConfigs];
                                      updated[idx].newName = e.target.value;
                                      setCustomColConfigs(updated);
                                    }}
                                    className="w-full bg-gray-900 border border-gray-700 rounded-lg px-2.5 py-1 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-medium placeholder-gray-600 disabled:opacity-40 disabled:bg-gray-950"
                                    placeholder="Nhập tên cột mới dễ nhớ..."
                                  />
                                </td>

                                {/* Dropdown vai trò nghiệp vụ */}
                                <td className="p-3">
                                  <select
                                    value={cfg.role}
                                    disabled={!cfg.use}
                                    onChange={(e) => {
                                      const updated = [...customColConfigs];
                                      const val = e.target.value as any;
                                      
                                      // Tránh việc gán trùng vai trò duy nhất cho cột khác
                                      if (val !== "") {
                                        updated.forEach((c, cIdx) => {
                                          if (cIdx !== idx && c.role === val) {
                                            c.role = "";
                                          }
                                        });
                                      }
                                      
                                      updated[idx].role = val;
                                      setCustomColConfigs(updated);
                                    }}
                                    className={`w-full bg-slate-900 border text-xs rounded-lg px-2 py-1 focus:outline-none focus:ring-1 focus:ring-purple-500 ${
                                      isSystemMappable 
                                        ? "border-purple-500 text-purple-300 font-bold" 
                                        : "border-gray-700 text-gray-400"
                                    } disabled:opacity-40`}
                                  >
                                    <option value="">-- Không gán vai trò --</option>
                                    <option value="mota">Mô tả hoạt động KD</option>
                                    <option value="manganh">Mã ngành thực tế (VSIC)</option>
                                    <option value="xa">Địa bàn (Xã)</option>
                                    <option value="doanhthu">Doanh thu</option>
                                    <option value="laodong">Lao động</option>
                                    <option value="idCol">Mã doanh nghiệp (Unique keys)</option>
                                  </select>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    {/* Nút Kích hoạt Tái cấu trúc bảng */}
                    <div className="flex justify-end pt-2">
                      <button
                        onClick={handleApplyColumnRedefinition}
                        className="bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-md shadow-purple-950/40 flex items-center gap-2 cursor-pointer border border-purple-500/20 hover:scale-[1.02] active:scale-[0.98]"
                      >
                        <FileCheck className="w-4 h-4" />⚡ XÁC NHẬN ĐỊNH NGHĨA & TÁI TẠO BẢNG DỮ LIỆU MỚI
                      </button>
                    </div>
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
                  <div className="overflow-x-auto max-h-[500px]">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#111827] text-gray-400 border-b border-gray-800 font-mono">
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
                      Hãy chọn "Tải tệp dữ liệu chính" ở ô phía trên hoặc nhấp nút "Nạp dữ liệu mẫu" ở góc trên để trải nghiệm thử nghiệm nhanh toàn bộ cơ cấu.
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
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
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
                    <BarChart3 className="w-5 h-5 text-amber-400" /> TỔNG HỢP DỮ LIỆU ĐA CHIỀU (PIVOT SUMMARY)
                  </h3>
                  <p className="text-xs text-gray-400">Lắp ráp các công thức tổng hợp, gom nhóm tính Tổng, Đếm dộc nhất, Đếm tần suất hoặc Tìm Trung Bình từ tệp dữ liệu đã nạp.</p>
                </div>

                {mainData.length > 0 ? (
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 border-t border-gray-800 pt-6">
                    
                    {/* BƯỚC 1: CHỌN CỘT TRỤC GROUP BY */}
                    <div className="bg-[#111827]/60 rounded-xl p-5 border border-[#374151] space-y-4">
                      <h4 className="text-xs font-bold text-amber-400 uppercase tracking-wider font-mono">1. Cột gom nhóm chính (Group By)</h4>
                      <p className="text-[11px] text-gray-400">Chọn 1 hoặc nhiều cột để làm trục phân cấp (ví dụ: Địa_Bàn_Xã, MaNganh):</p>
                      
                      <div className="max-h-[160px] overflow-y-auto border border-gray-850 p-3 rounded-lg space-y-2">
                        {/* TÙY CHỌN GOM NHÓM THEO NGÀNH KINH TẾ (VSIC) CHUẨN */}
                        {mapping.manganh && (
                          <>
                            <label className="flex items-center gap-2.5 text-xs text-amber-300 font-bold hover:text-amber-200 cursor-pointer select-none bg-[#1f2937]/50 p-1.5 rounded border border-amber-950">
                              <input 
                                type="checkbox"
                                checked={groupByCols.includes("_virtual_sector_cap2")}
                                onChange={() => {
                                  if (groupByCols.includes("_virtual_sector_cap2")) {
                                    setGroupByCols(groupByCols.filter(c => c !== "_virtual_sector_cap2"));
                                  } else {
                                    setGroupByCols([...groupByCols, "_virtual_sector_cap2"]);
                                  }
                                }}
                                className="rounded text-amber-500 focus:ring-amber-500 bg-gray-900 border-gray-750"
                              />
                              ⚡ Ngành Cấp 2 (2 số + tên tự ghép)
                            </label>
                            
                            <label className="flex items-center gap-2.5 text-xs text-amber-300 font-bold hover:text-amber-200 cursor-pointer select-none bg-[#1f2937]/50 p-1.5 rounded border border-amber-950">
                              <input 
                                type="checkbox"
                                checked={groupByCols.includes("_virtual_sector_cap1")}
                                onChange={() => {
                                  if (groupByCols.includes("_virtual_sector_cap1")) {
                                    setGroupByCols(groupByCols.filter(c => c !== "_virtual_sector_cap1"));
                                  } else {
                                    setGroupByCols([...groupByCols, "_virtual_sector_cap1"]);
                                  }
                                }}
                                className="rounded text-amber-500 focus:ring-amber-500 bg-gray-900 border-gray-750"
                              />
                              ⚡ Ngành Cấp 1 (Chữ cái + tên tự ghép)
                            </label>
                          </>
                        )}

                        {columns.map(col => {
                          const isChecked = groupByCols.includes(col);
                          return (
                            <label key={col} className="flex items-center gap-2.5 text-xs text-gray-300 hover:text-white cursor-pointer select-none">
                              <input 
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => {
                                  if (isChecked) {
                                    setGroupByCols(groupByCols.filter(c => c !== col));
                                  } else {
                                    setGroupByCols([...groupByCols, col]);
                                  }
                                }}
                                className="rounded text-amber-500 focus:ring-amber-500 bg-gray-900 border-gray-700"
                              />
                              {col}
                            </label>
                          );
                        })}
                      </div>
                    </div>

                    {/* BƯỚC 2: CÔNG THỨC CHI TIẾT */}
                    <div className="bg-[#111827]/60 rounded-xl p-5 border border-[#374151] space-y-4 flex flex-col justify-between">
                      <div className="space-y-4">
                        <h4 className="text-xs font-bold text-amber-400 uppercase tracking-wider font-mono">2. Cấu hình phép tính</h4>
                        
                        <div className="grid grid-cols-2 gap-3">
                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-gray-400 block">Chọn Cột</label>
                            <select 
                              value={newAggCol} 
                              onChange={(e) => setNewAggCol(e.target.value)}
                              className="w-full bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1 text-xs text-white"
                            >
                              <option value="">-- Chọn cột --</option>
                              {columns.map(c => <option key={c} value={c}>{c}</option>)}
                            </select>
                          </div>

                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-gray-400 block">Chọn Phép Tính</label>
                            <select 
                              value={newAggOp} 
                              onChange={(e) => setNewAggOp(e.target.value)}
                              className="w-full bg-[#1f2937] border border-[#374151] rounded-lg px-2 py-1 text-xs text-white"
                            >
                              <option value="sum">Tổng cộng (SUM)</option>
                              <option value="mean">Trung bình (AVERAGE)</option>
                              <option value="count">Đếm số dòng (COUNT)</option>
                              <option value="nunique">Đếm mục độc nhất (NUNIQUE)</option>
                              <option value="min">Nhỏ nhất (MIN)</option>
                              <option value="max">Lớn nhất (MAX)</option>
                            </select>
                          </div>
                        </div>

                        <button 
                          onClick={addAggRule}
                          className="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-4 py-2 rounded-lg transition-all w-full flex items-center justify-center gap-1 cursor-pointer"
                        >
                          <Plus className="w-4 h-4" /> Thêm quy tắc tính toán
                        </button>

                        {/* Danh sách quy tắc đã thêm */}
                        <div className="space-y-2 border-t border-gray-800 pt-3">
                          <label className="text-[10.5px] font-bold text-gray-400 block uppercase">Danh sách chỉ tiêu tổng hợp:</label>
                          {aggRules.length === 0 ? (
                            <div className="text-[11px] text-gray-500 italic">Chưa có chỉ tiêu nào được lập...</div>
                          ) : (
                            <div className="space-y-1 max-h-[140px] overflow-y-auto">
                              {aggRules.map((rule, idx) => (
                                <div key={idx} className="flex justify-between items-center bg-[#111827] px-3 py-1.5 rounded-lg border border-gray-800 text-xs">
                                  <span className="text-gray-300 font-mono">
                                    🔴 Phép <strong className="text-amber-400">{rule.op.toUpperCase()}</strong> trên cột <strong className="text-white">{rule.col}</strong>
                                  </span>
                                  <button onClick={() => removeAggRule(idx)} className="text-red-400 hover:text-red-300 cursor-pointer">
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>

                      <button 
                        onClick={handleRunSummary}
                        className="bg-amber-500 hover:bg-amber-600 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all w-full mt-4 flex items-center justify-center gap-1 cursor-pointer"
                      >
                        <BarChart3 className="w-4 h-4" /> BẮT ĐẦU CHẠY PHÂN TÍCH TỔNG HỢP NHÓM
                      </button>
                    </div>

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}

                {/* 2. CHỨC NĂNG BÁO CÁO NHANH CHUYÊN SÂU THEO VSIC PHÂN CẤP (KHẮC PHỤC LỖI KHÔNG KHỚP BAN ĐẦU) */}
                <div className="border-t border-[#374151] pt-6 space-y-4">
                  <h4 className="text-sm font-bold text-white flex items-center gap-2">
                    <Layers className="w-4 h-4 text-emerald-400" /> ⚡ 2. Báo cáo nhanh phân cấp ngành kinh tế (VSIC) & Xã địa bàn
                  </h4>
                  <p className="text-xs text-gray-400">Hệ thống áp dụng thuật toán rà quét đệ quy liên phân vùng (Khắc phục hoàn hảo lỗi lệch khớp mã ở đồ án Python cũ) thu gọn và tổng hợp chỉ số doanh thu và lao động theo nhóm ngành.</p>
                  
                  {mainData.length > 0 ? (
                    <div className="space-y-4 max-w-2xl">
                      
                      {/* BỘ LỰA CHỌN ĐỊNH DẠNG ĐẦU RA */}
                      <div className="bg-[#111827]/80 p-4 rounded-xl border border-emerald-500/10 space-y-2.5">
                        <span className="text-[10px] font-bold text-emerald-400 tracking-widest uppercase font-mono block">Cấu hình định dạng báo cáo đầu ra</span>
                        <div className="space-y-2">
                          <label className="flex items-start gap-2.5 text-xs text-gray-300 hover:text-white cursor-pointer select-none">
                            <input 
                              type="radio" 
                              name="quickReportTypeSetting"
                              checked={reportType === "pivot"} 
                              onChange={() => setReportType("pivot")}
                              className="mt-0.5 rounded-full text-emerald-500 focus:ring-emerald-500 bg-gray-900 border-gray-750 accent-emerald-500"
                            />
                            <div>
                              <div className="font-bold text-gray-200">Bảng xoay ngang Pivot (Khuyên dùng)</div>
                              <div className="text-[10.5px] text-gray-400">Các xã làm dòng, chi tiết các ngành và chỉ số tương ứng [DT], [LĐ] xếp thành cột kề nhau (Ví dụ: Ngành C - DT, Ngành C - LĐ, Ngành I - DT, Ngành I - LĐ).</div>
                            </div>
                          </label>

                          <label className="flex items-start gap-2.5 text-xs text-gray-300 hover:text-white cursor-pointer select-none border-t border-gray-800/60 pt-2 block">
                            <input 
                              type="radio" 
                              name="quickReportTypeSetting"
                              checked={reportType === "flat"} 
                              onChange={() => setReportType("flat")}
                              className="mt-0.5 rounded-full text-emerald-500 focus:ring-emerald-500 bg-gray-900 border-gray-750 accent-emerald-500"
                            />
                            <div>
                              <div className="font-bold text-gray-200">Danh sách bảng phẳng truyền thống</div>
                              <div className="text-[10.5px] text-gray-400">Mỗi hàng hiển thị một cặp khớp Ngành Cấp X kết hợp với Xã địa phương tương ứng.</div>
                            </div>
                          </label>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#111827]/60 p-4 rounded-xl border border-gray-850">
                        
                        <div className="bg-[#111827] rounded-lg p-4 border border-gray-800 text-center space-y-3">
                          <div className="text-emerald-400 font-bold text-xs uppercase font-mono">Báo cáo ngành Cấp 1</div>
                          <p className="text-[11px] text-gray-400 font-sans">Doanh thu và quy mô lao động cộng dồn theo các đại lĩnh vực chữ cái (A-U) phân chia theo địa phận xã.</p>
                          <button 
                            onClick={() => handleQuickReport(1)}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-lg transition-all w-full cursor-pointer shadow-sm"
                          >
                            Chạy Tổng Hợp Ngành Cấp 1 & Xã
                          </button>
                        </div>

                        <div className="bg-[#111827] rounded-lg p-4 border border-gray-800 text-center space-y-3">
                          <div className="text-emerald-400 font-bold text-xs uppercase font-mono">Báo cáo ngành Cấp 2</div>
                          <p className="text-[11px] text-gray-400 font-sans">Doanh thu và nhân lực tập trung theo ngành 2 chữ số chi tiết (01-99) phân chia theo địa phận xã.</p>
                          <button 
                            onClick={() => handleQuickReport(2)}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-lg transition-all w-full cursor-pointer shadow-sm"
                          >
                            Chạy Tổng Hợp Ngành Cấp 2 & Xã
                          </button>
                        </div>

                      </div>
                    </div>
                  ) : (
                    <div className="text-xs text-gray-500 italic">Vui lòng nạp dữ liệu nguồn chính trước...</div>
                  )}
                </div>

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
                    
                    <div className="bg-[#111827]/60 rounded-xl p-5 border border-indigo-500/10 grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-1">
                        <label className="text-xs font-bold text-indigo-400 block font-mono uppercase">Mã ngành thực nhập (Lấy từ map)</label>
                        <div className="text-xs text-gray-300 font-medium">Cột gắn nhãn: <span className="text-emerald-400 font-mono font-bold">"{mapping.manganh || "Chưa chọn!"}"</span></div>
                        <p className="text-[11px] text-gray-400 font-sans">Hệ thống sẽ phân tích mã này thành các cột phân cấp: Nganh_Cap_1, Nganh_Cap_2, Nganh_Cap_3, Nganh_Cap_4, Nganh_Cap_5 và tự so sánh cha con đệ quy chuẩn 100%.</p>
                      </div>

                      <div className="space-y-1">
                        <label className="text-xs font-bold text-indigo-400 block font-mono uppercase">mô tả hoạt động kinh doanh (Lấy từ map)</label>
                        <div className="text-xs text-gray-300 font-medium">Cột gắn nhãn: <span className="text-emerald-400 font-mono font-bold">"{mapping.mota || "Chưa chọn!"}"</span></div>
                        <p className="text-[11px] text-gray-400 font-sans">Chuỗi mô tả hoạt động sẽ được bóc tách từ khóa chính để hệ thống đối chiếu xem có lệch lĩnh vực, sai ngành đăng ký hay không.</p>
                      </div>
                    </div>

                    {/* LỰA CHỌN BỘ PHÂN PHÂN TÍCH */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      
                      {/* CÁCH 1: ALGORITHM MATCHER - SIÊU TỐC */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-gray-800 flex flex-col justify-between">
                        <div className="space-y-2">
                          <div className="flex items-center gap-2">
                            <span className="bg-emerald-950/40 border border-emerald-500/30 text-emerald-400 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Khuyên dùng</span>
                            <h4 className="text-sm font-bold text-white">Cách 1: Thuật toán so khớp thông minh (Siêu tốc 100% Offline)</h4>
                          </div>
                          <p className="text-xs text-gray-400 leading-relaxed font-sans">
                            Áp dụng động cơ <strong>Client-Side Smart Matcher</strong> tự thiết kế. Tự động rà quét hàng nghìn dòng, phân tách cụm từ tiếng Việt chuyên sâu tích lũy, sắp xếp phân nhóm. Đảm bảo chạy mượt mà tức thì, không bị dính mạng chậm, hiển thị tiến trình % chính xác.
                          </p>
                        </div>
                        <button 
                          onClick={() => handleStandardizeSectors(false)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all w-full mt-6 shadow-md shadow-emerald-900/10 cursor-pointer"
                        >
                          Chạy nhanh bằng thuật toán thông minh
                        </button>
                      </div>

                      {/* CÁCH 2: GEMINI API SERVER-SIDE PROXY */}
                      <div className="bg-[#111827] rounded-xl p-5 border border-[#374151] flex flex-col justify-between">
                        <div className="space-y-2">
                          <div className="flex items-center gap-1.5">
                            <span className="bg-indigo-950/40 border border-indigo-500/30 text-indigo-400 font-mono text-[10px] uppercase font-bold px-2 py-0.5 rounded-md">Trí tuệ nhân tạo</span>
                            <h4 className="text-sm font-bold text-white">Cách 2: Giải pháp phân tích lập luận Gemini 3.5 AI</h4>
                          </div>
                          <p className="text-xs text-gray-400 leading-relaxed font-sans">
                            Hệ thống sẽ gửi mô tả hoạt động qua cổng dịch vụ an toàn (server-side secrets) của Gemini 3.5 Flash để dịch hoạt động, suy luận logic lĩnh vực chuẩn xác và lý giải nguyên nhân tự nhiên. Thích hợp cho mẫu dữ liệu nhỏ hoặc muốn rà soát chiều sâu tinh xảo.
                          </p>
                        </div>
                        
                        <button 
                          onClick={() => handleStandardizeSectors(true)}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all w-full mt-6 shadow-md shadow-indigo-900/10 cursor-pointer flex items-center justify-center gap-1"
                        >
                          <Brain className="w-4 h-4 text-pink-400" /> Chạy bằng Trí Tuệ Nhân Tạo (Gemini 3.5 API)
                        </button>
                      </div>

                    </div>

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

        </main>
      </div>

    </div>
  );
}
