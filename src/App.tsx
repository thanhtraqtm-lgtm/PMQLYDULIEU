import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
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
  FileCheck,
  Calculator // Icon cho phép tính
} from "lucide-react";

import { 
  vsicRawData, 
  normalizeSectorCode, 
  getSectorHierarchy, 
  smartSuggestSectorByDescription,
  getSectorLevel,
  getParentSectorCode,
  formatVSICName 
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

// Interface mới cho cấu hình cột tính toán
interface CalculatedColumnConfig {
    use: boolean; // Có sử dụng cột này không
    newName: string; // Tên mới cho cột
    formula: string; // Công thức tính toán (ví dụ: "DoanhThu - LaoDong")
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

  // Cấu hình cột gốc (tên gốc, có dùng hay không, tên mới)
  const [customColConfigs, setCustomColConfigs] = useState<{
    originalName: string;
    use: boolean;
    newName: string;
  }[]>([]);

  // Cấu hình cột tính toán
  const [calculatedColConfigs, setCalculatedColConfigs] = useState<CalculatedColumnConfig[]>([]);
  const [newCalcColName, setNewCalcColName] = useState<string>("");
  const [newCalcColFormula, setNewCalcColFormula] = useState<string>("");

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
  const pageSize = 50;

  // Mock data generator - ĐÃ LOẠI BỎ PHẦN MOCK DATA CỨNG
  const loadMockData = () => {
    // Phần mock data đã được xóa bỏ theo yêu cầu. Nút này giờ chỉ đơn giản là gợi ý người dùng tải file.
    // Khi người dùng nhấn nút này, nó sẽ kích hoạt input file upload.
    // Nếu bạn muốn nút này làm gì đó khác, vui lòng cho biết.
    alert("Vui lòng chọn file dữ liệu của bạn để bắt đầu.");
    // Có thể thêm logic để kích hoạt input file upload nếu cần thiết
    // document.getElementById('file-upload-input')?.click(); 
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
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
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

          // Khởi tạo danh sách cấu hình cột gốc, TÊN MỚI BAN ĐẦU LÀ TÊN GỐC
          const initConfigs = cols.map(c => ({
            originalName: c,
            use: true,
            newName: c, // Tên mới ban đầu chính là tên gốc
          }));
          setCustomColConfigs(initConfigs);
          // Reset cấu hình cột tính toán khi nạp tệp mới
          setCalculatedColConfigs([]); 

          // Thử tìm tự động mapping từ tên cột tương tự (logic này vẫn giữ nguyên để gợi ý)
          const autoMap: ColumnMapping = { mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" };
          initConfigs.forEach(cfg => {
            const low = cfg.originalName.toLowerCase();
            if (low.includes("mô tả") || low.includes("mota") || low.includes("hoatdong")) {
              autoMap.mota = cfg.newName;
            }
            if (low.includes("mã ngành") || low.includes("manganh") || low.includes("ma_nganh") || low.includes("dtv") || low.includes("vsic")) {
              autoMap.manganh = cfg.newName;
            }
            if (low.includes("xã") || low.includes("xa") || low.includes("diaban") || low.includes("dia_ban")) {
              autoMap.xa = cfg.newName;
            }
            if (low.includes("doanh thu") || low.includes("doanhthu") || low.includes("thu_nhap")) {
              autoMap.doanhthu = cfg.newName;
            }
            if (low.includes("lao động") || low.includes("laodong") || low.includes("nhan_su")) {
              autoMap.laodong = cfg.newName;
            }
            if (low.includes("mst") || low.includes("mã số") || low.includes("ident") || low.includes("id")) {
              autoMap.idCol = cfg.newName;
            }
          });
          setMapping(autoMap);
          setActiveTab("xemdulieu");
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
    reader.readAsBinaryString(file);
  };

  // Reset toàn bộ dữ liệu và cấu hình
  const clearData = () => {
    setMainData([]);
    setRawImportedData([]);
    setColumns([]);
    setFileName("");
    setMapping({ mota: "", manganh: "", xa: "", doanhthu: "", laodong: "", idCol: "" });
    setCustomColConfigs([]);
    setCalculatedColConfigs([]); // Reset cột tính toán
    setActiveTab("trangchu"); 
  };

  // --- LỖI VÀ CẢNH BÁO CẦN XỬ LÝ KHI TẠO CỘT TÍNH TOÁN ---
  const evaluateFormula = (formula: string, row: any, availableCols: string[]): number | string | null => {
    try {
      // Tách công thức thành các phần tử: cột, toán tử, cột/số
      // Ví dụ: "DoanhThu - 500" -> ["DoanhThu", "-", "500"]
      // Cần xử lý trường hợp tên cột có khoảng trắng hoặc ký tự đặc biệt
      
      // Sử dụng regex để tách an toàn hơn, bao gồm cả tên cột có khoảng trắng
      const parts = formula.match(/([a-zA-ZÀ-ỹ\s]+)\s*([\+\-\*\/])\s*([a-zA-ZÀ-ỹ\s]+|\d+(\.\d+)?)/);
      
      if (!parts || parts.length < 4) {
          console.error("Invalid formula format:", formula);
          return null; // Công thức không hợp lệ
      }

      let col1Name = parts[1].trim();
      const operator = parts[2];
      let operand2 = parts[3].trim(); // Có thể là tên cột hoặc số

      // Tìm tên cột thực tế trong availableCols nếu operand2 là tên cột
      let col2Value: number | null = null;
      let col2Name = '';

      // Kiểm tra xem operand2 có phải là tên cột hợp lệ không
      const foundCol2 = availableCols.find(col => col.toLowerCase() === operand2.toLowerCase());
      if (foundCol2) {
        col2Name = foundCol2;
        col2Value = parseFloat(String(row[col2Name]).replace(/[^0-9.\-]/g, ""));
      } else {
        // Nếu không phải tên cột, thử parse thành số
        col2Value = parseFloat(operand2.replace(/[^0-9.\-]/g, ""));
      }
      
      // Lấy giá trị cột thứ nhất
      const col1Value = parseFloat(String(row[col1Name]).replace(/[^0-9.\-]/g, ""));

      // Kiểm tra NaN và thực hiện phép tính
      if (isNaN(col1Value) || (col2Value === null && !foundCol2) || (col2Value !== null && isNaN(col2Value) && !foundCol2) ) {
        // Nếu một trong hai không phải số hợp lệ (và operand2 không phải tên cột)
        return null; // Không thể tính toán
      }

      let result: number | null = null;
      const val2 = foundCol2 ? col2Value : parseFloat(operand2); // Dùng giá trị số nếu operand2 là số

      if (isNaN(val2)) return null; // Nếu operand2 vẫn không parse được thành số

      switch (operator) {
        case '+': result = col1Value + val2; break;
        case '-': result = col1Value - val2; break;
        case '*': result = col1Value * val2; break;
        case '/': 
          if (val2 === 0) {
              result = Infinity; // Hoặc NaN, tùy cách xử lý
          } else {
              result = col1Value / val2;
          }
          break;
        default: return null; // Toán tử không hợp lệ
      }

      // Làm tròn kết quả để tránh số thập phân dài
      return result !== null && !isNaN(result) ? Math.round(result * 100) / 100 : null;

    } catch (error) {
      console.error("Error evaluating formula:", formula, error);
      return null; // Lỗi khi đánh giá công thức
    }
  };


  // Áp dụng định nghĩa lại tên cột & tái cấu trúc bảng dữ liệu mới
  const handleApplyColumnRedefinition = async () => {
    if (rawImportedData.length === 0) {
      alert("Không tìm thấy dữ liệu tệp gốc để tái cấu trúc! Hãy nạp tệp chính trước.");
      return;
    }

    // Lọc các cấu hình cột gốc được chọn sử dụng
    const activeBaseConfigs = customColConfigs.filter(cfg => cfg.use && cfg.newName.trim() !== "");
    
    // Lọc các cấu hình cột tính toán được người dùng định nghĩa
    const activeCalcConfigs = calculatedColConfigs.filter(cfg => cfg.use && cfg.newName.trim() !== "");

    if (activeBaseConfigs.length === 0 && activeCalcConfigs.length === 0) {
      alert("Vui lòng chọn ít nhất một cột để sử dụng hoặc định nghĩa một cột tính toán!");
      return;
    }

    setLoading(true);
    setProgress(10);
    setStatusMessage("Đang lọc cột, đổi tên và tạo cột tính toán...");
    await sleep(200);

    // Bước 1: Tái cấu trúc bảng dựa trên các cột gốc được chọn
    const restructuredRows = rawImportedData.map(row => {
      const newRow: any = {};
      activeBaseConfigs.forEach(cfg => {
        const val = row[cfg.originalName];
        newRow[cfg.newName.trim()] = val !== undefined && val !== null ? val : "";
      });
      return newRow;
    });

    // Bước 2: Tạo các cột tính toán mới
    let finalRows = [...restructuredRows]; // Bắt đầu với các hàng đã tái cấu trúc
    const currentColumns = Object.keys(restructuredRows[0] || {}); // Các cột hiện có sau khi tái cấu trúc

    if (activeCalcConfigs.length > 0) {
        finalRows = finalRows.map(row => {
            const newRow = { ...row };
            activeCalcConfigs.forEach(calcCfg => {
                // Đánh giá công thức cho từng dòng
                const calculatedValue = evaluateFormula(calcCfg.formula, row, currentColumns);
                newRow[calcCfg.newName.trim()] = calculatedValue !== null ? calculatedValue : "";
            });
            return newRow;
        });
    }

    // Cập nhật cấu hình mapping chỉ tay từ hệ thống tới các tên cột mới định nghĩa
    const newMapping: ColumnMapping = {
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: ""
    };

    // Update mapping based on user's role assignments from the select dropdowns
    Object.keys(newMapping).forEach((key) => {
        const roleKey = key as keyof ColumnMapping;
        const mappedColumnName = mapping[roleKey];
        if (mappedColumnName) {
            // Find the corresponding config in activeBaseConfigs
            const matchingConfig = activeBaseConfigs.find(cfg => cfg.newName === mappedColumnName);
            if (matchingConfig) {
                newMapping[roleKey] = matchingConfig.newName;
            } else {
              // Check also in calculated columns if needed, though typically mapping is for base columns
              const matchingCalcConfig = activeCalcConfigs.find(cfg => cfg.newName === mappedColumnName);
              if (matchingCalcConfig) {
                newMapping[roleKey] = matchingConfig.newName;
              }
            }
        }
    });

    // Cố gắng tự động map các cột tính toán vào mapping nếu phù hợp
    activeCalcConfigs.forEach(calcCfg => {
        const formulaParts = calcCfg.formula.match(/([a-zA-ZÀ-ỹ\s]+)\s*([\+\-\*\/])\s*([a-zA-ZÀ-ỹ\s]+|\d+(\.\d+)?)/);
        if (formulaParts && formulaParts.length >= 4) {
            const op1Name = formulaParts[1].trim();
            const operator = formulaParts[2];
            const op2Str = formulaParts[3].trim();

            // Nếu là phép trừ và tên cột mới là "ChenhLenh" hoặc tương tự, có thể map vào đó
            // Logic này cần được tinh chỉnh thêm tùy thuộc vào quy tắc đặt tên của bạn
            if (operator === '-' && calcCfg.newName.toLowerCase().includes('chenh') && !newMapping.doanhthu) { // Ví dụ: map DoanhThu_ChenhLenh vào mapping.doanhthu nếu chưa có
               // Tạm thời không map vào mapping chuẩn vì nó có thể gây nhầm lẫn.
               // Mapping chuẩn chỉ nên dành cho các cột dữ liệu gốc.
            }
        }
    });


    setMainData(finalRows);
    setColumns(Object.keys(finalRows[0] || {}));
    setMapping(newMapping); // Update mapping based on user's role assignments
    setCalculatedColConfigs(activeCalcConfigs); // Lưu lại các cấu hình cột tính toán đã áp dụng
    setViewPage(1);

    setProgress(100);
    setStatusMessage(`Tái cấu trúc bảng thành công! Đã giữ lại ${activeBaseConfigs.length} cột gốc và tạo ${activeCalcConfigs.length} cột tính toán.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

  // Bộ lọc dữ liệu viewer
  const filteredData = useMemo(() => {
    if (!searchTerm) return mainData;
    const term = searchTerm.toLowerCase();
    return mainData.filter(row => {
      return Object.values(row).some(val => String(val).toLowerCase().includes(term));
    });
  }, [mainData, searchTerm]);

  // Phân trang dữ liệu hiển thị
  const paginatedData = useMemo(() => {
    const startIdx = (viewPage - 1) * pageSize;
    return filteredData.slice(startIdx, startIdx + pageSize);
  }, [filteredData, viewPage]);

  // Tổng số trang
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;

  // Tạm nghỉ bằng Promise cho việc Render mượt mà & hiển thị progress bar
  const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

  // --- Các hàm xử lý khác (Merge, Compare, Split, Summary, QuickReport, StandardizeSectors, LogicCheck) giữ nguyên ---
  // (Các hàm này được copy từ phiên bản trước và không thay đổi logic cốt lõi liên quan đến việc hiển thị dữ liệu cuối cùng)

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

    const rightMap = new Map();
    const batchSize = Math.max(1, Math.floor(rightData.length / 10));
    
    for (let i = 0; i < rightData.length; i++) {
      const row = rightData[i];
      const kv = String(row[rightKey] || "").trim();
      if (kv) {
        rightMap.set(kv, row);
      }
      if (i % batchSize === 0 || i === rightData.length - 1) {
        const pct = Math.floor((i / rightData.length) * 40); 
        setProgress(pct);
        setStatusMessage(`Đang lập chỉ mục bảng phải: ${i}/${rightData.length} dòng...`);
        await sleep(10);
      }
    }

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
        const finalColName = mainCols.includes(rc) ? `${rc}_Phai` : rc;
        mergedRow[finalColName] = matchedRight ? matchedRight[rc] : "";
      });

      mergedResults.push(mergedRow);

      if (j % stepSize === 0 || j === leftData.length - 1) {
        const pct = 40 + Math.floor((j / leftData.length) * 60);
        setProgress(pct);
        setStatusMessage(`Đang ánh xạ & ghép dòng: ${j}/${leftData.length} dòng...`);
        await sleep(10);
      }
    }

    setMainData(mergedResults);
    setColumns(Object.keys(mergedResults[0] || {}));
    setFileName(`GhepNoi_${leftFileName}_vs_${rightFileName}.xlsx`);
    
    if (!mapping.idCol) { // Auto map idCol nếu chưa map
        setMapping(prev => ({ ...prev, idCol: leftKey }));
    }

    setProgress(100);
    setStatusMessage(`Ghép nối thành công hoàn tất! Thu được ${mergedResults.length} dòng dữ liệu.`);
    await sleep(400);
    setLoading(false);
    setActiveTab("xemdulieu");
  };

// 2. CHỨC NĂNG SO SÁNH CŨ - MỚI (DIFF)
  const handleCompare = async () => {
    if (!oldData.length || !newData.length) {
        alert("Vui lòng tải lên cả hai tệp dữ liệu cũ và mới để so sánh.");
        return;
    }
    if (!diffKey) {
        alert("Vui lòng chọn cột khóa để đối chiếu.");
        return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Khởi động so sánh dữ liệu...");
    await sleep(200);

    // --- BẮT ĐẦU LOGIC MỚI: PHÂN TÍCH CỘT SỐ & TÍNH CHÊNH LỆCH TỰ ĐỘNG ---

    // Xác định các cột có thể là số dựa trên mapping (doanh thu, lao động) hoặc phân tích dữ liệu
    const potentialNumericCols = new Set<string>();
    // Thêm các cột từ mapping nếu chúng tồn tại trong cả hai file
    [mapping.doanhthu, mapping.laodong].filter(Boolean).forEach(col => {
        if (oldData.length > 0 && newData.length > 0 && oldData[0].hasOwnProperty(col) && newData[0].hasOwnProperty(col)) {
            potentialNumericCols.add(col);
        }
    });

    // Nếu chưa đủ, thử phân tích thêm các cột khác từ unionCols
    const oldCols = Object.keys(oldData[0] || {});
    const newCols = Object.keys(newData[0] || {});
    const unionCols = Array.from(new Set([...oldCols, ...newCols])).filter(c => c !== diffKey); // Lấy tất cả cột trừ khóa

    if (potentialNumericCols.size < 3 && unionCols.length > 0) { // Giới hạn phân tích thêm để tránh tốn thời gian
        const sampleSize = Math.min(100, oldData.length, newData.length); // Lấy mẫu 100 dòng đầu
        
        unionCols.forEach(col => {
            // Chỉ phân tích nếu chưa được xác định là số từ mapping và tên cột không chứa "_Cu", "_Moi", "_ChenhLenh", "TrangThai"
            if (!potentialNumericCols.has(col) && !col.toLowerCase().includes('_cu') && !col.toLowerCase().includes('_moi') && !col.toLowerCase().includes('_chenh') && !col.toLowerCase().includes('trangthai')) {
                let isLikelyNumeric = true;
                for (let i = 0; i < sampleSize; i++) {
                    const oldVal = oldData[i]?.[col];
                    const newVal = newData[i]?.[col];
                    
                    // Check if oldVal is numeric or empty/null
                    const isOldNumeric = !isNaN(parseFloat(String(oldVal).replace(/[^0-9.\-]/g, "")));
                    const isOldEmpty = oldVal === null || oldVal === undefined || String(oldVal).trim() === "";

                    // Check if newVal is numeric or empty/null
                    const isNewNumeric = !isNaN(parseFloat(String(newVal).replace(/[^0-9.\-]/g, "")));
                    const isNewEmpty = newVal === null || newVal === undefined || String(newVal).trim() === "";

                    // If a value exists and is not numeric, then this column is not purely numeric
                    if (!isOldEmpty && !isOldNumeric) { 
                        isLikelyNumeric = false;
                        break;
                    }
                     if (!isNewEmpty && !isNewNumeric) {
                        isLikelyNumeric = false;
                        break;
                    }
                }
                if (isLikelyNumeric) {
                    potentialNumericCols.add(col);
                }
            }
        });
    }
    const numericColsSet = potentialNumericCols; // Set các cột được xác định là số

    // --- Hết phần phân tích tự động ---

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

    // const oldCols = Object.keys(oldData[0] || {}); // Đã định nghĩa ở trên
    // const newCols = Object.keys(newData[0] || {}); // Đã định nghĩa ở trên
    // const unionCols = Array.from(new Set([...oldCols, ...newCols])).filter(c => c !== diffKey); // Đã định nghĩa ở trên

    for (let i = 0; i < allKeys.length; i++) {
      const key = allKeys[i];
      const oldRow = oldMap.get(key);
      const newRow = newMap.get(key);

      const combined: any = { [diffKey]: key };

      if (oldRow && !newRow) { // Đã xóa
        unionCols.forEach(col => {
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = "";
          // Nếu là cột số, cột chênh lệch sẽ là giá trị âm của cột cũ
          if (numericColsSet.has(col)) {
              const numCu = parseFloat(String(oldRow[col]).replace(/[^0-9.\-]/g, ""));
              combined[`${col}_ChenhLenh`] = !isNaN(numCu) ? -numCu : "";
          }
        });
        combined["TrangThai_SoSanh"] = "❌ Đã xóa";
      } else if (!oldRow && newRow) { // Mới thêm
        unionCols.forEach(col => {
          combined[`${col}_Cu`] = "";
          combined[`${col}_Moi`] = newRow[col] || "";
           // Nếu là cột số, cột chênh lệch sẽ là giá trị mới
          if (numericColsSet.has(col)) {
              const numMoi = parseFloat(String(newRow[col]).replace(/[^0-9.\-]/g, ""));
              combined[`${col}_ChenhLenh`] = !isNaN(numMoi) ? numMoi : "";
          }
        });
        combined["TrangThai_SoSanh"] = "✅ Mới thêm";
      } else { // Tồn tại ở cả 2 - cần kiểm tra thay đổi
        const changedCols: string[] = [];
        unionCols.forEach(col => {
          const valCu = String(oldRow[col] !== undefined ? oldRow[col] : "").trim();
          const valMoi = String(newRow[col] !== undefined ? newRow[col] : "").trim();
          
          combined[`${col}_Cu`] = oldRow[col] || "";
          combined[`${col}_Moi`] = newRow[col] || "";

          if (valCu !== valMoi) {
            changedCols.push(col);
            // Nếu là cột số, tính chênh lệch
            if (numericColsSet.has(col)) {
              const numCu = parseFloat(valCu.replace(/[^0-9.\-]/g, ""));
              const numMoi = parseFloat(valMoi.replace(/[^0-9.\-]/g, ""));
              const chenhLenh = !isNaN(numCu) && !isNaN(numMoi) ? Math.round((numMoi - numCu) * 100) / 100 : ""; // Use "" if calculation fails
              combined[`${col}_ChenhLenh`] = chenhLenh;
            }
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
    if (!mapping.idCol) { // Auto map idCol nếu chưa map
        setMapping(prev => ({ ...prev, idCol: diffKey }));
    }

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

    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const val = String(row[splitCol] || "Rong").trim();
      const safeVal = val.replace(/[^a-zA-Z0-9_\-À-ỹ\s]/g, "").replace(/\s+/g, "_"); 
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

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");

      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
      
      const buf = new ArrayBuffer(wbout.length);
      const view = new Uint8Array(buf);
      for (let j = 0; j < wbout.length; j++) {
        view[j] = wbout.charCodeAt(j) & 0xFF;
      }

      zip.file(`Tach_File_${key}.xlsx`, buf);

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
    setActiveTab("xemdulieu");
  };

  // 4. CHỨC NĂNG TỔNG HỢP BÁO CÁO ĐỘNG (DYNAMIC PIVOT & AGGREGATE)
  const addAggRule = () => {
    if (!newAggCol) {
      alert("Vui lòng chọn cột cần tính toán!");
      return;
    }
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

    const groups = new Map<string, any[]>();
    mainData.forEach(row => {
      const compositeKeyObj: any = {};
      groupByCols.forEach(col => {
        compositeKeyObj[col] = row[col] !== undefined ? String(row[col]) : "[Rỗng]";
      });
      const keyStr = JSON.stringify(compositeKeyObj);
      if (!groups.has(keyStr)) {
        groups.set(keyStr, []);
      }
      groups.get(keyStr)?.push(row);
    });

    const summaryRows: any[] = [];
    const keysArray = Array.from(groups.keys());
    const batchSize = Math.max(1, Math.floor(keysArray.length / 10));

    for (let k = 0; k < keysArray.length; k++) {
      const keyStr = keysArray[k];
      const rows = groups.get(keyStr) || [];
      const groupValueObj = JSON.parse(keyStr);

      const resultRow: any = { ...groupValueObj };

      aggRules.forEach(rule => {
        const { col, op } = rule;
        const colValues = rows.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== "");
        const numValues = colValues.map(v => Number(v)).filter(v => !isNaN(v));

        let calcVal: number | string = 0; 
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
    // Removed the hardcoded check for mapping.manganh and mapping.xa
    // Instead, we'll handle cases where they might be empty or not mapped.

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang tạo báo cáo nhanh Ngành Cấp ${level} kết hợp Xã...`);
    await sleep(200);

    const maNganhCol = mapping.manganh || ""; // Use mapped column or empty string
    const xaCol = mapping.xa || "";       // Use mapped column or empty string

    const processedData = mainData.map(row => {
      const mngRaw = maNganhCol ? row[maNganhCol] : ""; // Use column only if mapped
      const mng = normalizeSectorCode(mngRaw); 
      
      const hierarchy = getSectorHierarchy(mng);
      const cap1Info = hierarchy["1"];
      const cap2Info = hierarchy["2"];
      const cap5Info = hierarchy["5"];

      const tenNganhCap5 = cap5Info?.ten || (mng ? `[Ngành Cấp 5 không tìm thấy: ${mng}]` : "[Mã ngành rỗng]");
      const tenNganhCap2 = cap2Info?.ten || (mng && mng.length >= 2 ? `[Ngành Cấp 2 không tìm thấy: ${mng.slice(0, 2)}]` : "[Ngành cấp 2 chưa xác định]");
      const tenNganhCap1 = cap1Info?.ten || (mng && mng.length >= 1 ? `[Ngành Cấp 1 không tìm thấy: ${mng[0]}]` : "[Ngành cấp 1 chưa xác định]");

      let tenNganhLabel = ""; 
      if (level === 2) {
        const sec2Code = mng ? mng.slice(0, 2) : "";
        tenNganhLabel = sec2Code ? `${sec2Code} - ${tenNganhCap2}` : "[Chưa xác định Cấp 2]";
      } else { // level === 1
        const sec1Code = cap1Info?.ma || "";
        tenNganhLabel = sec1Code ? `${sec1Code} - ${tenNganhCap1}` : "[Chưa xác định Cấp 1]";
      }

      return {
        ...row,
        _tenNganhCap1: tenNganhCap1, 
        _tenNganhCap2: tenNganhCap2, 
        _tenNganhCap5: tenNganhCap5, 
        _maNganhCap1: cap1Info?.ma || "", 
        _maNganhCap2: cap2Info?.ma || "", 
        _maNganhCap5: mng, 
        _tenNganhLabel: tenNganhLabel, 
        _tempXa: xaCol ? String(row[xaCol] || "Khác").trim() : "Khác" // Use column only if mapped
      };
    });

    let finalReportRows: any[] = [];

    if (reportType === "pivot") {
      setStatusMessage("Đang tiến hành xoay (Pivot) gom nhóm theo từng Ngành Kinh Tế làm cột...");
      await sleep(150);

      const communes = Array.from(new Set(processedData.map(r => r._tempXa))).sort();
      const sectorLabels = Array.from(new Set(processedData.map(r => r._tenNganhLabel))).sort(); 

      communes.forEach((commune, cIdx) => {
        const communeObj: any = {
          "Địa_Bàn_Xã": commune 
        };

        let totalCommuneDN = 0;
        let totalCommuneDoanhThu = 0;
        let totalCommuneLaoDong = 0;

        sectorLabels.forEach(sectorLabel => {
          const matchedRows = processedData.filter(r => r._tempXa === commune && r._tenNganhLabel === sectorLabel);
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

          communeObj[`${sectorLabel} - Tổng Doanh Thu`] = Math.round(sumDoanhThu * 100) / 100;
          communeObj[`${sectorLabel} - Tổng Lao Động`] = Math.round(sumLaoDong);

          totalCommuneDN += matchedRows.length; 
          totalCommuneDoanhThu += sumDoanhThu; 
          totalCommuneLaoDong += sumLaoDong; 
        });

        communeObj["Số_DN_Địa_Phương"] = totalCommuneDN;
        communeObj["Tổng_Doanh_Thu_Địa_Phương"] = Math.round(totalCommuneDoanhThu * 100) / 100;
        communeObj["Tổng_Lao_Động_Địa_Phương"] = Math.round(totalCommuneLaoDong);

        finalReportRows.push(communeObj);
      });
    } else { // reportType === "flat"
      const groups = new Map<string, any[]>();
      processedData.forEach(row => {
        const key = JSON.stringify({
          Ngành_Cấp_1: row._tenNganhCap1,
          Mã_Ngành_Cấp_1: row._maNganhCap1,
          Ngành_Cấp_2: row._tenNganhCap2,
          Mã_Ngành_Cấp_2: row._maNganhCap2,
          Ngành_Cấp_5: row._tenNganhCap5,
          Mã_Ngành_Cấp_5: row._maNganhCap5,
          Địa_Bàn_Xã: row._tempXa
        });
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
          [`Ngành_Cấp_${level}`]: level === 1 ? dims.Ngành_Cấp_1 : dims.Ngành_Cấp_2, 
          "Mã_Ngành_Cấp_Báo_Cáo": level === 1 ? dims.Mã_Ngành_Cấp_1 : dims.Mã_Ngành_Cấp_2, 
          "Tên_Ngành_Cấp_Báo_Cáo": level === 1 ? dims.Ngành_Cấp_1 : dims.Ngành_Cấp_2, 
          "Địa_Bàn_Xã": dims.Địa_Bàn_Xã,
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

  // 6. CHUẨN HÓA & PHÂN TÍCH NGÀNH (KHỚP MÃ THÔNG MINH + GỢI Ý AI VỚI TIẾN TRÌNH THỰC TẾ)
  const handleStandardizeSectors = async (useAI: boolean) => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện.");
      return;
    }
    if (!mapping.manganh || !mapping.mota) {
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
      const maNganhRaw = mapping.manganh ? row[mapping.manganh] : ""; // Use mapped column or empty string
      const maDtvVal = normalizeSectorCode(maNganhRaw); 
      
      const hierarchy = getSectorHierarchy(maDtvVal);
      const cap1Info = hierarchy["1"];
      const cap2Info = hierarchy["2"];
      const cap5Info = hierarchy["5"];

      let goiyMa = "";
      let goiyTen = "";
      let diemTuongDong = "0.00";
      let giaiThich = "";
      let linhvucSuggest = ""; 

      if (useAI) {
        setStatusMessage(`[Phân tích AI] Đang dịch nghĩa dòng ${index + 1}/${mainData.length}: "${motaVal.slice(0, 30)}..."`);
        try {
          // Assume an API endpoint /api/gemini/analyze exists
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
            giaiThich = "Phân tích từ AI";
            diemTuongDong = "0.95"; 
          } else {
            console.warn("AI API call failed or returned error, falling back to local matcher.", data?.error);
            const local = smartSuggestSectorByDescription(motaVal);
            if (local) {
              goiyMa = normalizeSectorCode(local.ma);
              goiyTen = local.ten;
              diemTuongDong = local.diem.toFixed(2);
              giaiThich = `Mô hình cục bộ đề xuất phân nhóm (AI Server fallback)`;
              linhvucSuggest = getSectorHierarchy(goiyMa)["1"]?.ma || ""; 
            }
          }
        } catch (e) {
          console.error("Network error or API unavailable, falling back to local matcher.", e);
          const local = smartSuggestSectorByDescription(motaVal);
          if (local) {
            goiyMa = normalizeSectorCode(local.ma);
            goiyTen = local.ten;
            diemTuongDong = local.diem.toFixed(2);
            giaiThich = `Mô hình cục bộ đề xuất phân nhóm (Network error fallback)`;
            linhvucSuggest = getSectorHierarchy(goiyMa)["1"]?.ma || "";
          }
        }
      } else {
        const local = smartSuggestSectorByDescription(motaVal);
        if (local) {
          goiyMa = normalizeSectorCode(local.ma);
          goiyTen = local.ten;
          diemTuongDong = local.diem.toFixed(2);
          giaiThich = `Khớp từ khóa thông minh thành công đạt hiệu số tích hợp`;
          
          const sugHier = getSectorHierarchy(goiyMa);
          linhvucSuggest = sugHier["1"]?.ma || "";
        }
      }

      const stdCap5Ten = vsicRawData[maDtvVal] || `[Không tìm thấy tên Cấp 5 cho mã ${maDtvVal}]`;
      const stdCap2Ten = cap2Info?.ten || (maDtvVal && maDtvVal.length >= 2 ? `[Không tìm thấy tên Cấp 2 cho mã ${maDtvVal.slice(0, 2)}]` : "[Ngành cấp 2 chưa xác định]");
      const stdCap1Ten = cap1Info?.ten || (maDtvVal && maDtvVal.length >= 1 ? `[Không tìm thấy tên Cấp 1 cho mã ${maDtvVal[0]}]` : "[Ngành cấp 1 chưa xác định]");

      let trangThai = "✅ Hợp lệ"; 
      
      const dtvLinhVuc = cap1Info?.ma || ""; 

      const lcMota = motaVal.toLowerCase(); 
      
      const hasTradeKeywords = ["bán", "mua", "thương mại", "đại lý", "cửa hàng", "phân phối", "bán lẻ", "kinh doanh", "buôn bán", "nhà buôn", "tiệm", "shop", "market", "sales", "trade", "retail", "wholesale"].some(kw => lcMota.includes(kw));
      const hasIndustrialKeywords = ["sản xuất", "gia công", "chế tạo", "chế biến", "lắp ráp", "cơ khí", "dệt", "may", "in ấn", "nhà máy", "xưởng", "công nghiệp", "sản phẩm", "hàng hóa", "manufacture", "produce", "factory", "workshop", "industrial", "processing", "crafting"].some(kw => lcMota.includes(kw));

      const activeSub2 = maDtvVal ? maDtvVal.slice(0, 2) : "";
      const sub2Num = parseInt(activeSub2, 10);
      const isIndustrialCode = (!isNaN(sub2Num) && sub2Num >= 10 && sub2Num <= 33) || dtvLinhVuc === "C";

      if (!vsicRawData[maDtvVal]) { 
        trangThai = "❌ Lỗi: Mã ngành ĐTV không tồn tại trên danh mục VSIC chuẩn";
      } else if (hasTradeKeywords && !hasIndustrialKeywords && isIndustrialCode) {
        trangThai = "❌ Lỗi: Mô tả KD nghiêng về TM/Bán lẻ nhưng Mã đăng ký lại thuộc nhóm Công nghiệp";
      } else if (hasIndustrialKeywords && !isIndustrialCode && dtvLinhVuc !== "C") {
        trangThai = "❌ Lỗi: Mô tả KD nghiêng về SX/Công nghiệp nhưng Mã đăng ký lại thuộc nhóm ngoài Công nghiệp";
      } else if (linhvucSuggest && dtvLinhVuc && linhvucSuggest !== dtvLinhVuc) {
        trangThai = `❌ Lỗi (LỆCH LĨNH VỰC): Mô tả KD thiên về Nhóm [${linhvucSuggest}] nhưng Mã đăng ký thuộc Nhóm [${dtvLinhVuc}]`;
      } else if (goiyMa && parseFloat(diemTuongDong) > 0.6) { 
        const regCap2 = cap2Info?.ma || ""; 
        const sugHier = getSectorHierarchy(goiyMa); 
        const sugCap2 = sugHier["2"]?.ma || ""; 

        if (regCap2 && sugCap2 && regCap2 !== sugCap2) {
          trangThai = `⚠️ Cảnh báo (LỆCH CHI TIẾT CẤP 2): Gợi ý mã [${goiyMa}] (${sugHier["2"]?.ten}), nhưng đăng ký thực tế mã [${maDtvVal}]`;
        }
      }

      standardizedResults.push({
        ...row, 
        "Mã_Ngành_Đăng_Ký_Chuẩn": maDtvVal, 
        "Tên_Ngành_Cấp_5_VSIC": stdCap5Ten, 
        "Mã_Ngành_Cấp_2": cap2Info?.ma || "", 
        "Tên_Ngành_Cấp_2": stdCap2Ten, 
        "Mã_Ngành_Cấp_1": cap1Info?.ma || "", 
        "Tên_Ngành_Cấp_1": stdCap1Ten, 
        "Trạng_Thái_Kiểm_Tra_VSIC": trangThai, 
        "Mô_Tả_Hoạt_Động_Gốc": motaVal, 
        "Mã_Ngành_Gợi_Ý_AI": goiyMa, 
        "Tên_Ngành_Gợi_Ý": goiyTen, 
        "Độ_Tin_Cậy_Matcher": diemTuongDong, 
        "Giải_Thích_Phân_Tích": giaiThich, 
        "Lĩnh_Vực_KD_Gợi_Ý_C1": linhvucSuggest 
      });

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
    setStatusMessage(`Phân tích & Chuẩn hóa hoàn tất! Đã rà soát và phân tách thông tin ngành cho ${standardizedResults.length} dòng dữ liệu.`);
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

    const checkValue = (rowVal: any, op: string, compareVal: string) => {
      const v1 = String(rowVal !== undefined && rowVal !== null ? rowVal : "").trim().toLowerCase();
      const v2 = String(compareVal).trim().toLowerCase();

      if (op === "trống") return v1 === "";
      if (op === "không trống") return v1 !== "";

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

      if (op === "==") return v1 === v2;
      if (op === "!=") return v1 !== v2;
      if (op === "chứa") return v1.includes(v2);
      if (op === "không chứa") return !v1.includes(v2);
      
      return false; 
    };

    const results = mainData.map((row, index) => {
      const ifMatches = ifRules.map(r => checkValue(row[r.col], r.op, r.val));
      const satisfiesIf = ifCombine === "AND" 
        ? ifMatches.every(v => v === true) 
        : ifMatches.some(v => v === true);

      const thenMatches = thenRules.map(r => checkValue(row[r.col], r.op, r.val));
      const satisfiesThen = thenCombine === "AND"
        ? thenMatches.every(v => v === true)
        : thenMatches.some(v => v === true);

      let biViPham = false;
      let noteLoi = "";

      if (satisfiesIf && !satisfiesThen) {
        biViPham = true;
        const descriptIf = ifRules.map(r => `(${r.col} ${r.op} '${r.val}')`).join(` ${ifCombine} `);
        const descriptThen = thenRules.map(r => `(${r.col} ${r.op} '${r.val}')`).join(` ${thenCombine} `);
        noteLoi = `[VI PHẠM LOGIC] NẾU thỏa mãn: { ${descriptIf} } THÌ BẮT BUỘC PHẢI: { ${descriptThen} }; `;
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
    setStatusMessage(`Kiểm tra hoàn tất! Đã phân tích ${results.length} dòng và ghi nhận các vi phạm logic.`);
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
        XLSX.utils.book_append_sheet(wb, ws, "Bao_Cao_Ket_Qua"); 
        XLSX.writeFile(wb, fileName || "Ket_Qua_Xu_Ly_Du_Lieu.xlsx"); 
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
      
      <header className="border-b border-[#374151] bg-[#1f2937]/90 backdrop-blur-md sticky top-0 z-40 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-tr from-purple-600 to-indigo-500 p-2.5 rounded-xl shadow-lg shadow-purple-900/30 ring-1 ring-purple-400/50">
            <Layers className="w-6 h-6 text-white animate-pulse" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-white flex items-center gap-2">
              Công Cụ Hỗ Trợ Quản Lý Dữ Liệu <span className="bg-purple-600 text-[10px] uppercase font-bold px-2 py-0.5 rounded-full tracking-widest">V38.5</span>
            </h1>
            <p className="text-xs text-gray-400 font-mono">Hỗ trợ bạn xử lý dữ liệu nhanh chóng và hiệu quả</p>
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

          {/* Nút Nạp File được đặt ở đây để dễ truy cập */}
          <label className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition-all shadow-md shadow-purple-900/20 flex items-center gap-2 cursor-pointer">
            <FileUp className="w-4 h-4" /> NẠP DỮ LIỆU TỪ FILE
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv, .txt" 
              onChange={(e) => handleFileUpload(e, "main")} 
              className="hidden" 
              id="file-upload-input" // Thêm id để có thể kích hoạt bằng JS nếu cần
            />
          </label>
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden">
        
        <aside className="w-72 bg-[#1f2937]/60 border-r border-[#374151] p-5 space-y-2 flex flex-col justify-between">
          <div className="space-y-1.5">
            {/* Phần giới thiệu ban đầu khi chưa nạp file */}
            {activeTab === "trangchu" && mainData.length === 0 && (
              <div className="bg-[#1f2937]/50 border border-gray-700 rounded-xl p-4 text-xs text-gray-300 space-y-2.5 leading-relaxed animate-slide-in-bottom">
                <div className="font-bold text-purple-300 flex items-center gap-1.5 text-lg">
                  <Info className="w-5 h-5 text-purple-400" /> HƯỚNG DẪN NHANH
                </div>
                <p className="leading-relaxed">
                  Chào mừng bạn đến với **Công Cụ Hỗ Trợ Quản Lý Dữ Liệu**! Để bắt đầu, vui lòng thực hiện theo các bước sau:
                </p>
                <ul className="list-decimal list-inside space-y-1 text-[11px] text-gray-400 pl-1">
                  <li><strong>Nạp Dữ Liệu:</strong> Bấm vào nút <span className="font-bold text-purple-300">"NẠP DỮ LIỆU TỪ FILE"</span> ở góc trên bên phải để tải lên tệp Excel hoặc CSV của bạn.</li>
                  <li><strong>Định Nghĩa Cột:</strong> Sau khi nạp file, bạn có thể xem lại danh sách cột, đổi tên cột cho dễ nhớ và chọn những cột cần thiết.</li>
                  <li><strong>Sử Dụng Công Cụ:</strong> Lựa chọn các chức năng xử lý dữ liệu như Ghép nối, So sánh, Tách file, Tổng hợp báo cáo, hoặc Khớp mã ngành/Kiểm tra logic.</li>
                  <li><strong>Xuất Kết Quả:</strong> Khi hoàn tất, bạn có thể xuất dữ liệu đã xử lý ra file Excel.</li>
                </ul>
                <p className="leading-relaxed text-purple-300 font-semibold">
                  Chúng tôi luôn nỗ lực để mang đến trải nghiệm tốt nhất cho bạn.
                </p>
              </div>
            )}

            {/* Các nút điều hướng chính */}
            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 mb-2 pt-4">Chức năng chính</div>
            
            <button 
              onClick={() => setActiveTab("trangchu")}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                activeTab === "trangchu" 
                  ? "bg-purple-600/15 text-purple-400 border border-purple-500/20 shadow-sm" 
                  : "text-gray-300 hover:bg-[#374151]/50 hover:text-white"
              }`}
            >
              <Home className="w-4 h-4" /> 🏠 Tổng Quan & Hướng Dẫn
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
                <FileSpreadsheet className="w-4 h-4" /> 📂 Xem Dữ Liệu & Định Nghĩa Cột
              </span>
              <span className="text-[10px] font-mono bg-[#111827] text-gray-400 px-1.5 py-0.5 rounded-md">{mainData.length}</span>
            </button>

            <div className="text-[11px] font-bold text-gray-500 tracking-wider uppercase font-mono px-3 pt-4 mb-2">Công cụ xử lý nâng cao</div>

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

          <div className="bg-[#111827]/80 rounded-xl p-3 border border-[#374151] text-[10px] text-gray-400 font-mono leading-relaxed space-y-1">
            <div>⚙️ Engine: JS Client Native</div>
            <div>💡 Core Fix: VSIC Levels Ancestors Included</div>
            <div>⚡ Performance: Batch processing & async pipeline</div>
          </div>
        </aside>

        <main className="flex-1 bg-[#111827] overflow-y-auto p-6 md:p-8">
          
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

          {/* 1. TAB TRANG CHỦ (HIỂN THỊ HƯỚNG DẪN BAN ĐẦU HOẶC GIỚI THIỆU) */}
          {activeTab === "trangchu" && (
            <div className="space-y-8 animate-fade-in">
              {/* Phần giới thiệu ban đầu khi chưa nạp file */}
              {mainData.length === 0 && (
                <div className="bg-[#1f2937]/50 border border-gray-700 rounded-2xl p-6 space-y-4 leading-relaxed animate-slide-in-bottom">
                  <div className="font-bold text-purple-300 flex items-center gap-1.5 text-lg">
                    <Info className="w-5 h-5 text-purple-400" /> CHÀO MỪNG BẠN ĐẾN VỚI CÔNG CỤ HỖ TRỢ QUẢN LÝ DỮ LIỆU
                  </div>
                  <p className="text-xs text-gray-400">
                    Chúng tôi cung cấp một giải pháp toàn diện giúp bạn xử lý, phân tích và quản lý dữ liệu một cách hiệu quả.
                  </p>
                  <p className="text-xs text-gray-400">
                    Để bắt đầu, vui lòng thực hiện theo các bước đơn giản sau:
                  </p>
                  <ul className="list-decimal list-inside space-y-1 text-[11px] text-gray-400 pl-1">
                    <li><strong>Nạp Dữ Liệu:</strong> Bấm vào nút <span className="font-bold text-purple-300">"NẠP DỮ LIỆU TỪ FILE"</span> ở góc trên bên phải để tải lên tệp Excel hoặc CSV của bạn.</li>
                    <li><strong>Định Nghĩa Cột:</strong> Sau khi nạp file, bạn có thể xem lại danh sách cột, đổi tên cột cho dễ nhớ và chọn những cột cần thiết.</li>
                    <li><strong>Sử Dụng Công Cụ:</strong> Lựa chọn các chức năng xử lý dữ liệu như Ghép nối, So sánh, Tách file, Tổng hợp báo cáo, hoặc Khớp mã ngành/Kiểm tra logic.</li>
                    <li><strong>Xuất Kết Quả:</strong> Khi hoàn tất, bạn có thể xuất dữ liệu đã xử lý ra file Excel.</li>
                  </ul>
                  <p className="leading-relaxed text-purple-300 font-semibold">
                    Chúng tôi luôn nỗ lực để mang đến trải nghiệm tốt nhất cho bạn.
                  </p>
                </div>
              )}

              {/* Phần giới thiệu chung khi đã có dữ liệu hoặc muốn xem lại */}
              {mainData.length > 0 && (
                <div className="bg-gradient-to-r from-purple-900/40 via-[#1f2937] to-[#1f2937] border border-purple-500/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-center gap-6 justify-between">
                  <div className="space-y-3 max-w-2xl">
                    <span className="bg-purple-900/50 border border-purple-500/30 text-purple-400 text-xs font-mono font-bold px-3 py-1 rounded-full uppercase tracking-wider">
                      Phiên bản V38.5
                    </span>
                    <h2 className="text-2xl md:text-3xl font-extrabold text-white tracking-tight">
                      Công Cụ Hỗ Trợ Quản Lý Dữ Liệu
                    </h2>
                    <p className="text-gray-300 text-sm leading-relaxed">
                      Công cụ này được thiết kế để giúp bạn xử lý, phân tích và chuẩn hóa dữ liệu một cách hiệu quả, với các chức năng như ghép nối, so sánh, tách file, tổng hợp báo cáo và kiểm tra logic.
                    </p>
                    <div className="pt-2 flex items-center gap-4">
                      <button 
                        onClick={() => setActiveTab("xemdulieu")} // Chuyển đến tab xem dữ liệu
                        className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-sm px-6 py-2.5 rounded-xl transition-all shadow-md shadow-purple-900/30 flex items-center gap-2 cursor-pointer"
                      >
                        Bắt đầu xử lý dữ liệu <ArrowRight className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                  <div className="w-full md:w-auto flex justify-center">
                    <div className="bg-gradient-to-tr from-[#374151] to-purple-800/20 border border-[#4b5563] p-6 rounded-2xl text-center space-y-2 min-w-[200px] shadow-sm">
                      <div className="text-4xl font-extrabold text-white font-mono">{mainData.length}</div>
                      <div className="text-[11px] font-bold text-gray-400 tracking-wider uppercase font-mono">Dòng dữ liệu đang xử lý</div>
                      <div className="text-[10px] text-green-400 font-mono">{columns.length} cột</div>
                    </div>
                  </div>
                </div>
              )}

              {/* Các phần giới thiệu khác hoặc thông tin bổ sung có thể thêm ở đây */}
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
                      <FileSpreadsheet className="w-5 h-5 text-purple-400" /> QUẢN LÝ DỮ LIỆU NGUỒN
                    </h3>
                    <p className="text-xs text-gray-400">Nạp file dữ liệu của bạn hoặc xem lại file đã nạp. Bạn có thể định nghĩa lại tên cột và chọn các cột cần sử dụng.</p>
                  </div>

                  <label className="bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-purple-900/20 flex items-center gap-2 cursor-pointer self-start w-full md:w-auto justify-center">
                    <FileUp className="w-4 h-4" /> TẢI LÊN FILE DỮ LIỆU (EXCEL, CSV)
                    <input 
                      type="file" 
                      accept=".xlsx, .xls, .csv, .txt" 
                      onChange={(e) => handleFileUpload(e, "main")} 
                      className="hidden" 
                      id="file-upload-input" // Thêm id để có thể kích hoạt bằng JS nếu cần
                    />
                  </label>
                </div>

                {rawImportedData.length > 0 && (
                  <div className="bg-[#111827]/90 rounded-2xl p-5 border border-purple-500/20 space-y-5 animate-slide-up">
                    <div className="border-b border-gray-800 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <div className="text-xs font-bold text-purple-400 tracking-wider uppercase font-mono flex items-center gap-1.5">
                          <Database className="w-5 h-5 text-purple-400 animate-pulse" /> ĐỊNH NGHĨA LẠI TÊN CỘT & CHỌN CỘT SỬ DỤNG
                        </div>
                        <p className="text-xs text-gray-400 mt-1">
                          Bạn có thể sửa đổi tên cột cho dễ nhớ. Bỏ chọn các cột không cần thiết để dữ liệu được gọn gàng hơn.
                        </p>
                      </div>
                      
                      <div className="flex gap-2">
                        <button 
                          onClick={() => {
                            // Reset all configurations back to their original state
                            const resetConfigs = customColConfigs.map(c => ({
                              ...c,
                              use: true, 
                              newName: c.originalName 
                            }));
                            setCustomColConfigs(resetConfigs);
                            setCalculatedColConfigs([]); // Reset cột tính toán
                          }}
                          className="bg-gray-800 hover:bg-gray-700 text-gray-300 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-gray-700 cursor-pointer"
                          title="Hoàn tác tất cả tên cột về tên gốc"
                        >
                          Khôi Phục Tên Gốc
                        </button>
                        <button 
                          onClick={() => {
                            const prefilled = customColConfigs.map(cfg => {
                              const c = cfg.originalName;
                              let newN = c; 
                              // Keep default naming convention only for specific known columns
                              // ĐÃ BỎ CÁC ĐỊNH NGHĨA CỨNG NHẮC Ở ĐÂY
                              return { ...cfg, newName: newN };
                            });
                            setCustomColConfigs(prefilled);
                          }}
                          className="bg-purple-950/40 hover:bg-purple-900/40 text-purple-300 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all border border-purple-800/30 cursor-pointer"
                        >
                          Tự Động Đề Xuất Tên (Nếu cần)
                        </button>
                      </div>
                    </div>

                    <div className="bg-[#1e1b4b]/30 border border-purple-900/30 rounded-xl p-4 text-xs text-gray-300 space-y-1.5 leading-relaxed">
                      <div className="font-bold text-purple-300 flex items-center gap-1.5">
                        ⚙️ Hướng dẫn định nghĩa cột:
                      </div>
                      <ul className="list-disc list-inside space-y-1 text-[11px] text-gray-400 pl-1">
                        <li><strong>Đổi tên cột:</strong> Nhập tên mới vào ô "Tên mới dễ hiểu" để thay đổi cách hiển thị cột.</li>
                        <li><strong>Chọn cột sử dụng:</strong> Bỏ chọn ở ô "SỬ DỤNG" nếu bạn không muốn giữ lại cột này trong bảng dữ liệu cuối cùng.</li>
                        <li><strong>Gán vai trò hệ thống:</strong> Ở cột "VAI TRÒ NGHIỆP VỤ HỆ THỐNG", hãy chọn vai trò phù hợp cho từng cột (ví dụ: Mã ngành, Doanh thu). Điều này giúp các chức năng khác hoạt động chính xác.</li>
                      </ul>
                    </div>

                    <div className="overflow-x-auto border border-gray-800 rounded-xl bg-[#0f172a]/40">
                      <table className="w-full text-left text-xs min-w-[700px]">
                        <thead>
                          <tr className="bg-[#1f2937]/50 border-b border-gray-800 text-gray-400 font-mono text-[11px]">
                            <th className="p-3 text-center w-[70px]">SỬ DỤNG</th>
                            <th className="p-3 text-center w-[50px]">STT</th>
                            <th className="p-3">TÊN CỘT GỐC (TRONG FILE)</th>
                            <th className="p-3">TÊN MỚI (ĐỔI TÊN TẠI ĐÂY)</th>
                            <th className="p-3 w-[250px]">VAI TRÒ HỆ THỐNG (QUAN TRỌNG)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-800/60 font-sans">
                          {customColConfigs.map((cfg, idx) => {
                            // Role assignment is now handled by the select dropdowns below, not stored directly in customColConfigs
                            return (
                              <tr 
                                key={cfg.originalName} 
                                className={`transition-colors hover:bg-gray-800/15 ${
                                  cfg.use ? "bg-transparent" : "bg-gray-950/30 opacity-50"
                                }`}
                              >
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

                                <td className="p-3 text-center text-gray-500 font-mono text-[11px]">
                                  {idx + 1}
                                </td>

                                <td className="p-3">
                                  <div className="flex items-center gap-1.5">
                                    <span className="font-mono bg-gray-800 px-2 py-0.5 rounded text-gray-300 border border-gray-700 max-w-[200px] truncate block" title={cfg.originalName}>
                                      {cfg.originalName}
                                    </span>
                                  </div>
                                </td>

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
                                    placeholder="Nhập tên mới..."
                                  />
                                </td>

                                <td className="p-3">
                                  <select
                                    value={mapping[cfg.role as keyof ColumnMapping]} // Use mapping directly
                                    onChange={(e) => {
                                      const selectedRole = e.target.value as keyof ColumnMapping;
                                      const newMapping = { ...mapping };
                                      
                                      // Clear previous mapping for this role if it exists and is not the current one
                                      Object.keys(newMapping).forEach((key) => {
                                        if (key === selectedRole && newMapping[key] === cfg.newName) {
                                          // If the current column is already mapped to this role, do nothing
                                        } else if (newMapping[key as keyof ColumnMapping] === cfg.newName) {
                                          newMapping[key as keyof ColumnMapping] = "";
                                        }
                                      });

                                      // Set new mapping if a role is selected
                                      if (selectedRole !== "") { // Assuming "" or "-- Không gán vai trò --" maps to empty string
                                          newMapping[selectedRole as keyof ColumnMapping] = cfg.newName;
                                      }
                                      setMapping(newMapping);
                                    }}
                                    className={`w-full bg-slate-900 border text-xs rounded-lg px-2 py-1 focus:outline-none focus:ring-1 focus:ring-purple-500 ${
                                      // Highlight system-mappable roles
                                      ["mota", "manganh", "xa", "doanhthu", "laodong", "idCol"].includes(cfg.role as any) // Cast to 'any' to satisfy TS, as role is not in customColConfigs state anymore
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

                    {/* --- PHẦN TẠO CỘT TÍNH TOÁN --- */}
                    <div className="bg-[#111827]/90 rounded-2xl p-5 border border-blue-500/20 space-y-5 animate-slide-up mt-6">
                      <div className="border-b border-gray-800 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div>
                          <div className="text-xs font-bold text-blue-400 tracking-wider uppercase font-mono flex items-center gap-1.5">
                            <Calculator className="w-5 h-5 text-blue-400" /> TẠO CỘT TÍNH TOÁN TỪ CÔNG THỨC
                          </div>
                          <p className="text-xs text-gray-400 mt-1">
                            Định nghĩa các cột mới dựa trên phép toán (+, -, *, /) giữa các cột số hiện có.
                          </p>
                        </div>
                      </div>

                      {/* Input cho công thức mới */}
                      <div className="space-y-3">
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                          <div className="col-span-1 md:col-span-1">
                            <label className="text-xs font-bold text-gray-400 block mb-1">Tên cột mới</label>
                            <input
                              type="text"
                              value={newCalcColName}
                              onChange={(e) => setNewCalcColName(e.target.value)}
                              placeholder="Ví dụ: Lợi nhuận"
                              className="w-full bg-gray-900 border border-gray-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-blue-500 font-medium placeholder-gray-600"
                            />
                          </div>
                          <div className="col-span-1 md:col-span-2">
                             <label className="text-xs font-bold text-gray-400 block mb-1">Công thức (Cột1 ToánTử Cột2 hoặc Số)</label>
                             <input
                               type="text"
                               value={newCalcColFormula}
                               onChange={(e) => setNewCalcColFormula(e.target.value)}
                               placeholder="Ví dụ: DoanhThu - LaoDong hoặc DoanhThu * 0.1"
                               className="w-full bg-gray-900 border border-gray-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-blue-500 font-medium placeholder-gray-600"
                            />
                          </div>
                        </div>
                        
                        <button 
                          onClick={() => {
                            if (!newCalcColName.trim() || !newCalcColFormula.trim()) {
                                alert("Vui lòng nhập tên cột mới và công thức!");
                                return;
                            }
                            // Kiểm tra xem tên cột đã tồn tại chưa
                            if (columns.includes(newCalcColName.trim()) || 
                                customColConfigs.some(c => c.newName.trim() === newCalcColName.trim()) ||
                                calculatedColConfigs.some(c => c.newName.trim() === newCalcColName.trim())) {
                                alert(`Tên cột "${newCalcColName.trim()}" đã tồn tại. Vui lòng chọn tên khác.`);
                                return;
                            }
                            
                            setCalculatedColConfigs([...calculatedColConfigs, { use: true, newName: newCalcColName.trim(), formula: newCalcColFormula.trim() }]);
                            setNewCalcColName(""); // Clear input
                            setNewCalcColFormula(""); // Clear input
                          }}
                          className="bg-blue-700 hover:bg-blue-800 text-white text-xs font-bold px-4 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer"
                        >
                          <Plus className="w-4 h-4" /> Thêm cột tính toán
                        </button>
                      </div>

                      {/* Danh sách các cột tính toán đã thêm */}
                      {calculatedColConfigs.length > 0 && (
                        <div className="mt-4 space-y-2 border-t border-gray-800 pt-3">
                          <label className="text-[10.5px] font-bold text-gray-400 block uppercase">Các cột tính toán đã định nghĩa:</label>
                          <div className="max-h-[150px] overflow-y-auto space-y-1 pr-2">
                             {calculatedColConfigs.map((calcCfg, idx) => (
                                <div key={idx} className="flex justify-between items-center bg-[#111827] px-3 py-1.5 rounded-lg border border-blue-800/30 text-xs">
                                  <span className="text-blue-300 font-mono">
                                    <strong className="text-white">{calcCfg.newName}</strong> = {calcCfg.formula}
                                  </span>
                                  <button onClick={() => {
                                      setCalculatedColConfigs(calculatedColConfigs.filter((_, i) => i !== idx));
                                  }} className="text-red-400 hover:text-red-300 cursor-pointer">
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                             ))}
                          </div>
                        </div>
                      )}

                    </div>

                    <div className="flex justify-end pt-2 mt-4 border-t border-[#374151] pt-4">
                      <button
                        onClick={handleApplyColumnRedefinition}
                        className="bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs px-6 py-3 rounded-xl transition-all shadow-md shadow-purple-950/40 flex items-center gap-2 cursor-pointer border border-purple-500/20 hover:scale-[1.02] active:scale-[0.98]"
                      >
                        <FileCheck className="w-4 h-4" />⚡ ÁP DỤNG ĐỊNH NGHĨA & TẠO BẢNG DỮ LIỆU MỚI
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {mainData.length > 0 ? (
                <div className="bg-[#1f2937] border border-[#374151] rounded-2xl overflow-hidden shadow-sm space-y-4 p-4">
                  
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
                      Hãy chọn "Tải lên file dữ liệu" để bắt đầu xử lý.
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
                  <GitMerge className="w-5 h-5 text-blue-400" /> GHÉP NỐI HAI BẢNG DỮ LIỆU
                </h3>
                <p className="text-xs text-gray-400">Kết hợp hai tệp dữ liệu dựa theo trường khóa liên kết tương ứng (left outer join).</p>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  
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
                  
                  <div className="bg-[#111827]/60 rounded-xl p-5 border border-gray-800 space-y-4 text-center">
                    <h4 className="text-sm font-bold text-gray-400">📁 FILE DỮ LIỆU BẢN CŨ</h4>
                    <label className="inline-block bg-[#1f2937] hover:bg-[#374151] text-xs text-white border border-[#4b5563] font-semibold px-4 py-2.5 rounded-lg cursor-pointer transition-all">
                      Tải File Cũ
                      <input type="file" onChange={(e) => handleFileUpload(e, "old")} className="hidden" />
                    </label>
                    <div className="text-xs text-gray-400 font-mono select-none">{oldFileName ? `📂 ${oldFileName} (${oldData.length} dòng)` : "Chưa tải file cũ"}</div>
                  </div>

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
                    
                    <div className="bg-[#111827]/60 rounded-xl p-5 border border-[#374151] space-y-4">
                      <h4 className="text-xs font-bold text-amber-400 uppercase tracking-wider font-mono">1. Cột gom nhóm chính (Group By)</h4>
                      <p className="text-[11px] text-gray-400">Chọn 1 hoặc nhiều cột để làm trục phân cấp (ví dụ: Địa_Bàn_Xã, Mã Ngành Cấp 1):</p>
                      
                      <div className="max-h-[160px] overflow-y-auto border border-gray-850 p-3 rounded-lg space-y-2">
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
                              <option value="sum">Tổng (Sum)</option>
                              <option value="count">Đếm (Count)</option>
                              <option value="nunique">Đếm độc nhất (Unique Count)</option>
                              <option value="mean">Trung bình (Average)</option>
                              <option value="min">Nhỏ nhất (Min)</option>
                              <option value="max">Lớn nhất (Max)</option>
                            </select>
                          </div>
                        </div>
                        
                        <button 
                          onClick={addAggRule}
                          className="w-full bg-amber-700 hover:bg-amber-800 text-white text-xs font-bold px-4 py-1.5 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer"
                        >
                          <Plus className="w-4 h-4" /> Thêm Quy Tắc Tính Toán
                        </button>
                      </div>

                      <div className="border-t border-gray-800 pt-3">
                        <label className="text-[10.5px] font-bold text-amber-400 block uppercase">Các quy tắc đã thêm:</label>
                        <div className="max-h-[150px] overflow-y-auto space-y-1 pr-2">
                           {aggRules.map((rule, idx) => (
                              <div key={idx} className="flex justify-between items-center bg-[#111827] px-3 py-1.5 rounded-lg border border-amber-800/30 text-xs">
                                <span className="text-amber-300 font-mono">
                                  <strong className="text-white">{rule.col}</strong> ({rule.op})
                                </span>
                                <button onClick={() => removeAggRule(idx)} className="text-red-400 hover:text-red-300 cursor-pointer">
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                           ))}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-end pt-4 border-t border-gray-800 pt-4">
                    <button 
                      onClick={handleRunSummary}
                      className="bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl transition-all shadow-md shadow-amber-900/30 font-sans cursor-pointer flex items-center gap-1.5"
                    >
                      <BarChart3 className="w-4 h-4" /> CHẠY TỔNG HỢP BÁO CÁO
                    </button>
                  </div>
                </div>
              ) : (
                <div className="bg-[#1f2937]/40 border-2 border-dashed border-[#374151] p-12 text-center rounded-2xl space-y-4">
                  <Database className="w-12 h-12 text-[#4b5563] mx-auto animate-pulse" />
                  <div>
                    <h4 className="text-base font-bold text-white">Chưa có dữ liệu để tổng hợp</h4>
                    <p className="text-xs text-gray-400 max-w-md mx-auto pt-1 leading-relaxed">
                      Hãy nạp tệp dữ liệu chính ở Tab "Xem Dữ Liệu & Định Nghĩa Cột" trước khi sử dụng chức năng này.
                    </p>
                  </div>
                </div>
              )}
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

                  </div>
                ) : (
                  <div className="bg-[#111827]/50 rounded-xl p-6 text-center text-xs text-amber-400 border border-amber-950">
                    ⚠️ Yêu cầu nạp dữ liệu nguồn chính trước!
                  </div>
                )}
              </div>
            </div>
          )}
