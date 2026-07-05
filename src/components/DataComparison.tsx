import React, { useState, useMemo } from "react";
import { RefreshCw, FileUp, Calendar, TrendingUp, HelpCircle, ArrowRight, Download, BarChart2, CheckCircle2 } from "lucide-react";
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
  // Mode switcher: diff (2 separate files) vs time-series (Month-by-Month / YoY)
  const [subMode, setSubMode] = useState<"diff" | "time-series">("diff");

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
