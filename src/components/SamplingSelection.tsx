import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  FileCheck,
  Sliders,
  Activity,
  FileUp,
  Search,
  Layers,
  AlertTriangle,
  Download,
  Play,
  Check,
  RefreshCw,
  Trash2,
  Building,
  Users,
  Database,
  ArrowRight,
  Info
} from "lucide-react";
import { vsicRawData } from "../data/vsic";

interface SamplingSelectionProps {
  mainData: any[];
  columns: string[];
  mapping: any;
  setLoading: (loading: boolean) => void;
  setStatusMessage: (msg: string) => void;
  sampCorpData: any[];
  setSampCorpData: (data: any[]) => void;
  sampCorpFileName: string;
  setSampCorpFileName: (name: string) => void;
  sampIndData: any[];
  setSampIndData: (data: any[]) => void;
  sampIndFileName: string;
  setSampIndFileName: (name: string) => void;
}

export default function SamplingSelection({
  mainData,
  columns,
  setLoading,
  setStatusMessage,
  sampCorpData,
  setSampCorpData,
  sampCorpFileName,
  setSampCorpFileName,
  sampIndData,
  setSampIndData,
  sampIndFileName,
  setSampIndFileName
}: SamplingSelectionProps) {

  // --- SOURCE TYPE SELECTION ("main" vs "uploaded") ---
  const [corpSourceType, setCorpSourceType] = useState<"main" | "uploaded">("main");
  const [indSourceType, setIndSourceType] = useState<"main" | "uploaded">("main");

  // --- STRICT MANUAL COLUMN SELECTIONS (Initialized to empty, absolutely NO auto-matching) ---
  const [corpManganhCol, setCorpManganhCol] = useState<string>("");
  const [corpDoanhThuCol, setCorpDoanhThuCol] = useState<string>("");
  const [corpIdCol, setCorpIdCol] = useState<string>("");
  const [corpXaCol, setCorpXaCol] = useState<string>("");

  const [indManganhCol, setIndManganhCol] = useState<string>("");
  const [indDoanhThuCol, setIndDoanhThuCol] = useState<string>("");
  const [indIdCol, setIndIdCol] = useState<string>("");
  const [indXaCol, setIndXaCol] = useState<string>("");

  // --- CALCULATING STATES ---
  const [corpHasFiltered, setCorpHasFiltered] = useState<boolean>(false);
  const [indHasFiltered, setIndHasFiltered] = useState<boolean>(false);

  // --- REGULAR CONFIG STATES ---
  const [entCutoffPercent, setEntCutoffPercent] = useState<number>(75);
  const [entMinGroupSize, setEntMinGroupSize] = useState<number>(1);
  const [entForceStates, setEntForceStates] = useState<boolean>(true);
  const [entForceMonthly, setEntForceMonthly] = useState<boolean>(true);
  const [entGroupScope, setEntGroupScope] = useState<"province" | "xa">("province");

  // Household configurations
  const [indSize1To5Value, setIndSize1To5Value] = useState<number>(5);
  const [indSize1To5All, setIndSize1To5All] = useState<boolean>(true);
  const [indSize6To100Value, setIndSize6To100Value] = useState<number>(5);
  const [indSize101To1000Value, setIndSize101To1000Value] = useState<number>(8);
  const [indSize1001PlusPercent, setIndSize1001PlusPercent] = useState<number>(1);
  const [indTransportPercent, setIndTransportPercent] = useState<number>(1.5);
  const [indTransportMaxCap, setIndTransportMaxCap] = useState<number>(50);

  // Dynamic sectors to filter industries on both sides
  const [corpSectors, setCorpSectors] = useState<string[]>(["congnghiep", "thuongmai", "vantai", "anuongluutru"]);
  const [indSectors, setIndSectors] = useState<string[]>(["congnghiep", "thuongmai", "vantai", "anuongluutru"]);

  // Search terms for output lists
  const [corpSearchTerm, setCorpSearchTerm] = useState<string>("");
  const [indSearchTerm, setIndSearchTerm] = useState<string>("");

  // --- FINAL CALCULATED RESULTS ---
  const [corpSelectedList, setCorpSelectedList] = useState<any[]>([]);
  const [corpBackupList, setCorpBackupList] = useState<any[]>([]);
  const [corpGroupStats, setCorpGroupStats] = useState<Record<string, any>>({});

  const [indSelectedList, setIndSelectedList] = useState<any[]>([]);
  const [indBackupList, setIndBackupList] = useState<any[]>([]);
  const [indGroupStats, setIndGroupStats] = useState<Record<string, any>>({});

  // Get lists of available columns depending on the selected source
  const corpColumns = useMemo(() => {
    if (corpSourceType === "main") return columns;
    if (sampCorpData.length > 0) return Object.keys(sampCorpData[0]);
    return [];
  }, [corpSourceType, columns, sampCorpData]);

  const indColumns = useMemo(() => {
    if (indSourceType === "main") return columns;
    if (sampIndData.length > 0) return Object.keys(sampIndData[0]);
    return [];
  }, [indSourceType, columns, sampIndData]);

  // Helper to check if a VSIC code matches selected sectors
  const isSectorSelected = (vsicL2: string, sectors: string[]): boolean => {
    const l2Num = parseInt(vsicL2, 10);
    if (isNaN(l2Num)) return false;

    let matched = false;
    if (sectors.includes("congnghiep") && l2Num >= 5 && l2Num <= 33) matched = true;
    if (sectors.includes("thuongmai") && l2Num >= 45 && l2Num <= 46) matched = true;
    if (sectors.includes("vantai") && l2Num === 49) matched = true;
    if (sectors.includes("anuongluutru") && (l2Num === 55 || l2Num === 56)) matched = true;

    return matched;
  };

  // High-performance CSV parser
  const parseCSV = (rawText: string): any[] => {
    let text = rawText;
    if (text.charCodeAt(0) === 0xFEFF) {
      text = text.substring(1);
    }

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
          start = i + 1;
        } else {
          if (i + 1 < length && text[i + 1] === '"') {
            i++;
          } else {
            inQuotes = false;
          }
        }
      } else if (!inQuotes) {
        if (char === delimiter) {
          let cell = text.substring(start, i);
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

  // Handle uploaded files without auto-detecting column roles
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "corp" | "ind") => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang tải tệp tin: ${file.name}...`);
    
    // Reset filtered state so they must manually re-filter
    if (type === "corp") {
      setCorpHasFiltered(false);
      setCorpManganhCol("");
      setCorpDoanhThuCol("");
      setCorpIdCol("");
      setCorpXaCol("");
    } else {
      setIndHasFiltered(false);
      setIndManganhCol("");
      setIndDoanhThuCol("");
      setIndIdCol("");
      setIndXaCol("");
    }

    const isCSV = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");

    if (isCSV) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const text = evt.target?.result as string;
          if (!text) throw new Error("Tệp không có nội dung!");

          const data = parseCSV(text);
          if (data.length === 0) {
            alert("Tệp trống hoặc không đúng định dạng!");
            setLoading(false);
            return;
          }

          if (type === "corp") {
            setSampCorpData(data);
            setSampCorpFileName(file.name);
          } else {
            setSampIndData(data);
            setSampIndFileName(file.name);
          }
          setStatusMessage(`Đã nạp ${data.length} dòng dữ liệu từ ${file.name}. Vui lòng chỉ định các cột.`);
        } catch (err: any) {
          alert("Lỗi đọc file CSV: " + err.message);
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
          if (!arrayBuffer) throw new Error("Không có dữ liệu ArrayBuffer!");

          const wb = XLSX.read(arrayBuffer, { 
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false
          });

          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          if (data.length === 0) {
            alert("Tệp Excel trống!");
            setLoading(false);
            return;
          }

          if (type === "corp") {
            setSampCorpData(data);
            setSampCorpFileName(file.name);
          } else {
            setSampIndData(data);
            setSampIndFileName(file.name);
          }
          setStatusMessage(`Đã nạp ${data.length} dòng từ Excel: ${file.name}. Vui lòng chỉ định các cột.`);
        } catch (err: any) {
          alert("Lỗi đọc file Excel: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // --- ACTION: PERFORM ENTERPRISE SAMPLING FILTER ---
  const handleFilterCorp = () => {
    if (!corpManganhCol || !corpDoanhThuCol) {
      alert("⚠️ Bạn bắt buộc phải chỉ ra 'Cột mã ngành' và 'Cột doanh thu' của Doanh nghiệp để chương trình thực hiện lọc!");
      return;
    }

    setLoading(true);
    setStatusMessage("Đang thực hiện lọc & chọn mẫu doanh nghiệp...");

    setTimeout(() => {
      try {
        const sourceData = corpSourceType === "main" ? mainData : sampCorpData;
        if (sourceData.length === 0) {
          alert("Dữ liệu nguồn doanh nghiệp trống!");
          setLoading(false);
          return;
        }

        // 1. Process and extract VSIC Level 2 codes (cắt lấy 2 số đầu)
        const processedRows = sourceData.map((row, index) => {
          const manganhVal = String(row[corpManganhCol] || "").trim();
          const cleanDigits = manganhVal.replace(/\D/g, "");
          const vsicL2 = cleanDigits.slice(0, 2);

          const idVal = corpIdCol ? String(row[corpIdCol] || "") : `DN_${index}`;
          const xaVal = corpXaCol ? String(row[corpXaCol] || "") : "30000";
          const revVal = parseFloat(String(row[corpDoanhThuCol] || "0").replace(/,/g, "")) || 0;

          // Deduce name dynamically
          const nameVal = String(row["Tên doanh nghiệp"] || row["Tên đơn vị"] || row["Tên"] || `Doanh nghiệp ${index}`);

          return {
            id: idVal,
            name: nameVal,
            xaCode: xaVal,
            vsicL2,
            vsicFull: manganhVal,
            revenue: revVal,
            originalRow: row
          };
        });

        // 2. Filter rows that fall within selected industry sectors
        const filteredBySector = processedRows.filter(row => isSectorSelected(row.vsicL2, corpSectors));

        // 3. Grouping logic depending on entGroupScope (province or commune level)
        const groups: Record<string, any[]> = {};
        filteredBySector.forEach(row => {
          const groupKey = entGroupScope === "province" ? row.vsicL2 : `${row.xaCode}-${row.vsicL2}`;
          if (!groups[groupKey]) groups[groupKey] = [];
          groups[groupKey].push(row);
        });

        const selectedSet = new Set<string>();
        const backupSet = new Set<string>();
        const resultsGroupStats: Record<string, any> = {};

        // 4. Run the cumulative 75% cutoff calculations per group
        Object.entries(groups).forEach(([groupKey, list]) => {
          const sorted = [...list].sort((a, b) => b.revenue - a.revenue);
          const totalGroupRevenue = sorted.reduce((sum, item) => sum + item.revenue, 0);
          const targetCutoffRevenue = totalGroupRevenue * (entCutoffPercent / 100);

          let runningSum = 0;
          const groupSelected: any[] = [];
          const groupBackup: any[] = [];

          // Separate always selected (DNNN)
          const alwaysSelected = sorted.filter(row => {
            if (entForceStates && row.originalRow) {
              const isDnnn = Object.entries(row.originalRow).some(([key, val]) => {
                const kLow = key.toLowerCase();
                const vLow = String(val || "").toLowerCase();
                const matchesKey = kLow.includes("loại hình") || kLow.includes("loại doanh nghiệp") || kLow.includes("dnnn") || kLow.includes("loai hinh");
                const matchesVal = vLow.includes("nhà nước") || vLow.includes("nha nuoc") || vLow === "có" || vLow === "co" || vLow === "yes";
                return matchesKey && matchesVal;
              });
              if (isDnnn) return true;
            }
            if (sorted.length <= entMinGroupSize) return true;
            return false;
          });

          const alwaysSelectedIds = new Set(alwaysSelected.map(x => x.id));
          const standardCandidates = sorted.filter(x => !alwaysSelectedIds.has(x.id));

          // First, add all forced DNNN and minimum sized groups
          alwaysSelected.forEach(row => {
            runningSum += row.revenue;
            groupSelected.push({ ...row, selectionType: "Chọn toàn bộ (DNNN / Ngưỡng nhóm cực nhỏ)" });
            selectedSet.add(row.id);
          });

          // Second, add other candidates until cutoff is reached
          standardCandidates.forEach(row => {
            const isCentralForce = entForceMonthly && row.originalRow && Object.entries(row.originalRow).some(([key, val]) => {
              const kLow = key.toLowerCase();
              const vLow = String(val || "").toLowerCase();
              const matchesKey = kLow.includes("mẫu trung ương") || kLow.includes("mẫu tu") || kLow.includes("mau tu") || kLow.includes("mẫu t.ư");
              const matchesVal = vLow === "có" || vLow === "co" || vLow === "yes" || vLow === "1" || vLow === "true" || vLow.includes("trung ương");
              return matchesKey && matchesVal;
            });

            if (isCentralForce || runningSum < targetCutoffRevenue) {
              runningSum += row.revenue;
              groupSelected.push({ ...row, selectionType: isCentralForce ? "Ưu tiên mẫu trung ương" : "Doanh thu lũy kế" });
              selectedSet.add(row.id);
            } else {
              groupBackup.push({ ...row, selectionType: "Dự phòng" });
              backupSet.add(row.id);
            }
          });

          resultsGroupStats[groupKey] = {
            totalN: list.length,
            totalRevenue: totalGroupRevenue,
            selectedCount: groupSelected.length,
            backupCount: groupBackup.length,
            runningSumPercent: totalGroupRevenue > 0 ? (runningSum / totalGroupRevenue) * 100 : 0
          };
        });

        // 5. Save output lists
        const finalSelected = filteredBySector.filter(row => selectedSet.has(row.id));
        const finalBackup = filteredBySector.filter(row => backupSet.has(row.id));

        setCorpSelectedList(finalSelected);
        setCorpBackupList(finalBackup);
        setCorpGroupStats(resultsGroupStats);
        setCorpHasFiltered(true);

        setStatusMessage(`Đã lọc xong Doanh nghiệp: Chọn ${finalSelected.length} mẫu chính thức, ${finalBackup.length} mẫu dự phòng.`);
      } catch (err: any) {
        alert("Lỗi khi tính toán chọn mẫu doanh nghiệp: " + err.message);
      } finally {
        setLoading(false);
      }
    }, 100);
  };

  // --- ACTION: PERFORM INDIVIDUAL SAMPLING FILTER ---
  const handleFilterInd = () => {
    if (!indManganhCol || !indDoanhThuCol) {
      alert("⚠️ Bạn bắt buộc phải chỉ ra 'Cột mã ngành' và 'Cột doanh thu' của Hộ cá thể để chương trình thực hiện lọc!");
      return;
    }

    setLoading(true);
    setStatusMessage("Đang thực hiện lọc & chọn mẫu hộ cá thể theo chuẩn GSO...");

    setTimeout(() => {
      try {
        const sourceData = indSourceType === "main" ? mainData : sampIndData;
        if (sourceData.length === 0) {
          alert("Dữ liệu nguồn hộ cá thể trống!");
          setLoading(false);
          return;
        }

        // 1. Process and extract VSIC Level 2 codes (cắt lấy 2 số đầu)
        const processedRows = sourceData.map((row, index) => {
          const manganhVal = String(row[indManganhCol] || "").trim();
          const cleanDigits = manganhVal.replace(/\D/g, "");
          const vsicL2 = cleanDigits.slice(0, 2);

          const idVal = indIdCol ? String(row[indIdCol] || "") : `IND_${index}`;
          const xaVal = indXaCol ? String(row[indXaCol] || "") : "30000";
          const revVal = parseFloat(String(row[indDoanhThuCol] || "0").replace(/,/g, "")) || 0;

          // Deduce name dynamically
          const nameVal = String(row["Tên hộ"] || row["Tên chủ hộ"] || row["Tên đơn vị"] || row["Tên"] || `Hộ cá thể ${index}`);

          return {
            id: idVal,
            name: nameVal,
            xaCode: xaVal,
            vsicL2,
            vsicFull: manganhVal,
            revenue: revVal,
            originalRow: row
          };
        });

        // 2. Filter rows by selected industry sectors
        const filteredBySector = processedRows.filter(row => isSectorSelected(row.vsicL2, indSectors));

        // 3. Group by Ward (Địa bàn xã) + Industry Level 2
        const groups: Record<string, any[]> = {};
        filteredBySector.forEach(row => {
          const groupKey = `${row.xaCode}-${row.vsicL2}`;
          if (!groups[groupKey]) groups[groupKey] = [];
          groups[groupKey].push(row);
        });

        const selectedSet = new Set<string>();
        const backupSet = new Set<string>();
        const resultsGroupStats: Record<string, any> = {};

        // 4. Calculate target sample sizes for each commune-VSIC group according to GSO rules
        Object.entries(groups).forEach(([groupKey, list]) => {
          const [xaCode, vsicL2] = groupKey.split("-");
          const sorted = [...list].sort((a, b) => b.revenue - a.revenue);
          const totalN = sorted.length;
          const totalRevenue = sorted.reduce((sum, item) => sum + item.revenue, 0);

          let targetSize = 0;
          const isTransport = ["49", "50", "51", "52", "53"].includes(vsicL2);

          if (totalN <= 5) {
            targetSize = indSize1To5All ? totalN : Math.min(indSize1To5Value, totalN);
          } else if (totalN <= 100) {
            targetSize = Math.min(indSize6To100Value, totalN);
          } else if (totalN <= 1000) {
            targetSize = Math.min(indSize101To1000Value, totalN);
          } else {
            // Over 1001
            if (isTransport) {
              targetSize = Math.min(indTransportMaxCap, Math.max(8, Math.round(totalN * (indTransportPercent / 100))));
            } else {
              targetSize = Math.max(8, Math.round(totalN * (indSize1001PlusPercent / 100)));
            }
          }

          // Assign official samples and replacements based on sorted revenue rank
          sorted.forEach((row, rank) => {
            if (rank < targetSize) {
              selectedSet.add(row.id);
            } else {
              backupSet.add(row.id);
            }
          });

          resultsGroupStats[groupKey] = {
            totalN,
            totalRevenue,
            selectedCount: Math.min(targetSize, totalN),
            backupCount: Math.max(0, totalN - targetSize)
          };
        });

        // 5. Save output lists
        const finalSelected = filteredBySector.filter(row => selectedSet.has(row.id));
        const finalBackup = filteredBySector.filter(row => backupSet.has(row.id));

        setIndSelectedList(finalSelected);
        setIndBackupList(finalBackup);
        setIndGroupStats(resultsGroupStats);
        setIndHasFiltered(true);

        setStatusMessage(`Đã lọc xong Hộ cá thể: Chọn ${finalSelected.length} mẫu chính thức, ${finalBackup.length} mẫu dự phòng.`);
      } catch (err: any) {
        alert("Lỗi khi tính toán chọn mẫu hộ cá thể: " + err.message);
      } finally {
        setLoading(false);
      }
    }, 100);
  };

  // --- EXPORT DOWNLOAD HANDLERS ---
  const handleExportCSV = (type: "corp" | "ind") => {
    try {
      const isCorp = type === "corp";
      const selectedList = isCorp ? corpSelectedList : indSelectedList;
      const backupList = isCorp ? corpBackupList : indBackupList;
      const groupStats = isCorp ? corpGroupStats : indGroupStats;

      if (selectedList.length === 0) {
        alert("⚠️ Chưa có danh sách mẫu đã lọc để xuất tệp tin! Vui lòng thực hiện lọc trước.");
        return;
      }

      const headers = [
        "Mã định danh/MST",
        "Tên cơ sở/đơn vị",
        "Mã địa bàn xã/phường",
        "Mã ngành VSIC",
        "Mã ngành Cấp 2",
        "Tên ngành Cấp 2",
        "Doanh thu / Sản lượng",
        "Phân loại",
        "Trạng thái chọn mẫu"
      ];

      const csvRows: string[][] = [];

      // Add selected official samples
      selectedList.forEach(item => {
        const vsicL2Name = vsicRawData[item.vsicL2] || `Ngành cấp 2 (${item.vsicL2})`;
        csvRows.push([
          item.id,
          item.name,
          item.xaCode,
          item.vsicFull,
          item.vsicL2,
          vsicL2Name,
          String(item.revenue),
          isCorp ? "Doanh nghiệp" : "Hộ cá thể",
          "Mẫu chính thức"
        ]);
      });

      // Add backup samples
      backupList.forEach(item => {
        const vsicL2Name = vsicRawData[item.vsicL2] || `Ngành cấp 2 (${item.vsicL2})`;
        csvRows.push([
          item.id,
          item.name,
          item.xaCode,
          item.vsicFull,
          item.vsicL2,
          vsicL2Name,
          String(item.revenue),
          isCorp ? "Doanh nghiệp" : "Hộ cá thể",
          "Mẫu dự phòng"
        ]);
      });

      // Encode CSV content
      const headers_escaped = headers.map(h => `"${h}"`);
      let csvContent = "\uFEFF"; // BOM for UTF-8
      csvContent += headers_escaped.join(",") + "\n";

      csvRows.forEach(r => {
        const escaped = r.map(v => `"${String(v).replace(/"/g, '""')}"`);
        csvContent += escaped.join(",") + "\n";
      });

      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `Mau_Khao_Sat_${isCorp ? "Doanh_Nghiep" : "Ho_Ca_The"}_${new Date().toISOString().slice(0, 10)}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err: any) {
      alert("Lỗi khi xuất tệp tin CSV: " + err.message);
    }
  };

  // Filter lists based on users search inputs
  const displayedCorpList = useMemo(() => {
    if (!corpSearchTerm) return corpSelectedList.slice(0, 50);
    const term = corpSearchTerm.toLowerCase();
    return corpSelectedList.filter(item => 
      item.id.toLowerCase().includes(term) ||
      item.name.toLowerCase().includes(term) ||
      item.vsicFull.toLowerCase().includes(term) ||
      item.xaCode.toLowerCase().includes(term)
    ).slice(0, 50);
  }, [corpSelectedList, corpSearchTerm]);

  const displayedIndList = useMemo(() => {
    if (!indSearchTerm) return indSelectedList.slice(0, 50);
    const term = indSearchTerm.toLowerCase();
    return indSelectedList.filter(item => 
      item.id.toLowerCase().includes(term) ||
      item.name.toLowerCase().includes(term) ||
      item.vsicFull.toLowerCase().includes(term) ||
      item.xaCode.toLowerCase().includes(term)
    ).slice(0, 50);
  }, [indSelectedList, indSearchTerm]);

  return (
    <div className="space-y-8 animate-fade-in font-sans text-slate-800">
      
      {/* HEADER BANNER CARD */}
      <div className="bg-gradient-to-r from-indigo-900 via-indigo-950 to-slate-900 text-white rounded-3xl p-6 shadow-xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-80 h-80 bg-indigo-500/10 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none" />
        <div className="relative z-10 space-y-2">
          <span className="bg-indigo-500/20 text-indigo-300 text-[10px] font-extrabold uppercase tracking-wider px-3.5 py-1.5 rounded-full border border-indigo-500/30">
            🖥️ PHÂN HỆ KHẢO SÁT CHUYÊN BIỆT
          </span>
          <h2 className="text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
            <FileCheck className="w-6.5 h-6.5 text-indigo-400" /> BỘ ĐIỀU CHỈNH CHỌN MẪU SONG SONG
          </h2>
          <p className="text-xs text-slate-300 max-w-4xl leading-relaxed">
            Hệ thống rà soát và tạo mẫu điều tra độc lập cho cả hai phân hệ <b>Doanh nghiệp</b> & <b>Hộ cá thể</b> đồng thời. 
            <span className="text-emerald-400 font-bold ml-1">Quy tắc tuyệt đối: Không tự động gán trước bất cứ cột nào</span>, người dùng chỉ định thủ công các vai trò cột để đảm bảo độ tin cậy tuyệt đối 100%.
          </p>
        </div>
      </div>

      {/* 2-COLUMN SPLIT WORKSPACE */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-stretch">
        
        {/* ====================================================================== */}
        {/* LEFT COLUMN: ENTERPRISE SAMPLING (DOANH NGHIỆP) */}
        {/* ====================================================================== */}
        <div className="bg-white border-2 border-indigo-100 rounded-3xl p-6 shadow-lg flex flex-col justify-between space-y-6 relative hover:border-indigo-300 transition-all">
          <div className="absolute top-4 right-4 bg-indigo-50 text-indigo-700 text-[10px] font-black uppercase tracking-widest px-2.5 py-1 rounded-md border border-indigo-100">
            Màn 1: Doanh nghiệp
          </div>

          <div className="space-y-6">
            <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
              <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-2xl border border-indigo-100">
                <Building className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black text-indigo-950 uppercase tracking-wider">🏢 PHÂN HỆ DOANH NGHIỆP</h3>
                <p className="text-[11px] text-slate-500">Thiết lập nạp tệp, chọn các cột doanh thu lũy kế 75% tối ưu.</p>
              </div>
            </div>

            {/* BƯỚC 1: CHỌN NGUỒN DỮ LIỆU */}
            <div className="space-y-3">
              <label className="block text-xs font-extrabold text-indigo-950 tracking-wide uppercase">
                Bước 1: Chọn Nguồn Dữ Liệu Doanh Nghiệp
              </label>
              
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => {
                    setCorpSourceType("main");
                    setCorpHasFiltered(false);
                  }}
                  className={`px-4 py-3 rounded-xl border text-xs font-bold transition-all text-center flex items-center justify-center gap-1.5 cursor-pointer ${
                    corpSourceType === "main"
                      ? "bg-indigo-600 text-white border-indigo-700 shadow-md"
                      : "bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200"
                  }`}
                >
                  <Database className="w-4 h-4" /> Dữ liệu chính ({mainData.length.toLocaleString()})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setCorpSourceType("uploaded");
                    setCorpHasFiltered(false);
                  }}
                  className={`px-4 py-3 rounded-xl border text-xs font-bold transition-all text-center flex items-center justify-center gap-1.5 cursor-pointer ${
                    corpSourceType === "uploaded"
                      ? "bg-indigo-600 text-white border-indigo-700 shadow-md"
                      : "bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200"
                  }`}
                >
                  <FileUp className="w-4 h-4" /> Tự nạp tệp riêng {sampCorpFileName ? "✔️" : ""}
                </button>
              </div>

              {corpSourceType === "uploaded" && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                  <label className="bg-white hover:bg-indigo-50/20 border-2 border-dashed border-indigo-200 text-indigo-950 font-bold text-xs p-3.5 rounded-xl transition-all flex flex-col items-center justify-center gap-1.5 cursor-pointer">
                    <FileUp className="w-5 h-5 text-indigo-500" />
                    <span>Nạp tệp Excel/CSV doanh nghiệp</span>
                    <input
                      type="file"
                      accept=".xlsx, .xls, .csv"
                      onChange={(e) => handleFileUpload(e, "corp")}
                      className="hidden"
                    />
                  </label>
                  {sampCorpFileName && (
                    <div className="text-[11px] text-indigo-900 bg-white border border-indigo-100 p-2.5 rounded-lg flex items-center justify-between">
                      <span className="truncate font-semibold max-w-[180px]">📁 {sampCorpFileName}</span>
                      <span className="font-mono font-bold bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded">
                        {sampCorpData.length} dòng
                      </span>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* BƯỚC 2: CHỈ ĐỊNH CÁC CỘT (STRICTLY MANUAL) */}
            <div className="bg-indigo-50/40 border border-indigo-100 rounded-2xl p-4.5 space-y-4">
              <div className="flex items-start gap-2">
                <AlertTriangle className="w-4.5 h-4.5 text-indigo-600 shrink-0 mt-0.5" />
                <div className="space-y-0.5">
                  <span className="text-xs font-black text-indigo-950 block">CHỈ ĐỊNH CỘT DỮ LIỆU DOANH NGHIỆP</span>
                  <span className="text-[10.5px] text-slate-500 block">Không được tự động điền sẵn. Hãy chọn đúng tên cột tương ứng trong bảng dữ liệu:</span>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="block text-slate-700 font-extrabold mb-1">1. Cột Mã Ngành VSIC *</label>
                  <select
                    value={corpManganhCol}
                    onChange={(e) => {
                      setCorpManganhCol(e.target.value);
                      setCorpHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm font-sans"
                  >
                    <option value="">-- Click để Chọn Cột Mã Ngành --</option>
                    {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-700 font-extrabold mb-1">2. Cột Doanh thu / Sản lượng *</label>
                  <select
                    value={corpDoanhThuCol}
                    onChange={(e) => {
                      setCorpDoanhThuCol(e.target.value);
                      setCorpHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm font-sans"
                  >
                    <option value="">-- Click để Chọn Cột Doanh Thu --</option>
                    {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-500 font-bold mb-1">3. Cột Định danh (Mã số thuế) (Tùy chọn)</label>
                  <select
                    value={corpIdCol}
                    onChange={(e) => {
                      setCorpIdCol(e.target.value);
                      setCorpHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột ID --</option>
                    {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-500 font-bold mb-1">4. Cột Địa bàn xã / Phường (Tùy chọn)</label>
                  <select
                    value={corpXaCol}
                    onChange={(e) => {
                      setCorpXaCol(e.target.value);
                      setCorpHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột xã/phương --</option>
                    {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>
            </div>

            {/* BƯỚC 3: CẤU HÌNH THAM SỐ CHỌN MẪU */}
            <div className="space-y-4">
              <label className="block text-xs font-extrabold text-indigo-950 tracking-wide uppercase">
                Bước 3: Tinh Chỉnh Tham Số Lọc Doanh Nghiệp
              </label>

              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4.5 space-y-4 text-xs font-sans">
                <div>
                  <label className="block text-slate-700 font-extrabold mb-1">Phạm vi gom nhóm</label>
                  <select
                    value={entGroupScope}
                    onChange={(e) => {
                      setEntGroupScope(e.target.value as "province" | "xa");
                      setCorpHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm font-sans"
                  >
                    <option value="province">Gom nhóm toàn tỉnh (Đăng ký cấp ngành 2 - Đề xuất ✨)</option>
                    <option value="xa">Gom nhóm chi tiết từng địa bàn Xã/Phường</option>
                  </select>
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-slate-700 font-bold">Ngưỡng tích lũy đóng góp doanh thu (%):</span>
                    <span className="text-indigo-600 font-black text-sm">{entCutoffPercent}%</span>
                  </div>
                  <input
                    type="range"
                    min="50"
                    max="95"
                    step="5"
                    value={entCutoffPercent}
                    onChange={(e) => {
                      setEntCutoffPercent(parseInt(e.target.value));
                      setCorpHasFiltered(false);
                    }}
                    className="w-full accent-indigo-600 cursor-pointer"
                  />
                  <p className="text-[10px] text-slate-400 leading-normal">
                    Chương trình sắp xếp doanh nghiệp từ lớn đến nhỏ. Lấy các đơn vị chiếm đóng góp đầu bảng cho tới khi đạt tích lũy {entCutoffPercent}% tổng doanh thu.
                  </p>
                </div>

                <div className="grid grid-cols-2 gap-4 pt-2 border-t border-slate-200">
                  <label className="flex items-center gap-2 cursor-pointer text-slate-700 font-bold text-[11px]">
                    <input
                      type="checkbox"
                      checked={entForceStates}
                      onChange={(e) => {
                        setEntForceStates(e.target.checked);
                        setCorpHasFiltered(false);
                      }}
                      className="accent-indigo-600 h-3.5 w-3.5 rounded"
                    />
                    <span>Chọn 100% DNNN</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-slate-700 font-bold text-[11px]">
                    <input
                      type="checkbox"
                      checked={entForceMonthly}
                      onChange={(e) => {
                        setEntForceMonthly(e.target.checked);
                        setCorpHasFiltered(false);
                      }}
                      className="accent-indigo-600 h-3.5 w-3.5 rounded"
                    />
                    <span>Ưu tiên mẫu T.Ư</span>
                  </label>
                </div>

                {/* Nhóm ngành chọn lọc */}
                <div className="pt-2.5 border-t border-slate-200">
                  <span className="block text-slate-700 font-extrabold mb-1.5">Lọc nhóm ngành khảo sát:</span>
                  <div className="flex flex-wrap gap-2.5">
                    {["congnghiep", "thuongmai", "vantai", "anuongluutru"].map(sec => {
                      const labels: Record<string, string> = {
                        congnghiep: "Công nghiệp (05-33)",
                        thuongmai: "Thương mại (45-46)",
                        vantai: "Vận tải (49)",
                        anuongluutru: "Dịch vụ/Ăn uống (55-56)"
                      };
                      return (
                        <label key={sec} className="flex items-center gap-1.5 cursor-pointer text-[10.5px]">
                          <input
                            type="checkbox"
                            checked={corpSectors.includes(sec)}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setCorpSectors(prev => [...prev, sec]);
                              } else {
                                setCorpSectors(prev => prev.filter(x => x !== sec));
                              }
                              setCorpHasFiltered(false);
                            }}
                            className="accent-indigo-600 rounded"
                          />
                          <span>{labels[sec]}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* CHẠY LỌC DOANH NGHIỆP ACTION BUTTON */}
          <div className="pt-4 border-t border-slate-100 space-y-4">
            <button
              type="button"
              onClick={handleFilterCorp}
              className="w-full py-4 px-6 bg-gradient-to-r from-indigo-600 to-indigo-800 hover:from-indigo-700 hover:to-indigo-900 text-white font-black rounded-2xl text-xs uppercase tracking-widest shadow-lg flex items-center justify-center gap-2 transition-all cursor-pointer hover:scale-[1.01]"
            >
              <Play className="w-4 h-4 fill-white" /> Thực Hiện Lọc &amp; Chọn Mẫu Doanh Nghiệp
            </button>

            {/* RESULT COMPARTMENT FOR ENTERPRISE */}
            {corpHasFiltered ? (
              <div className="bg-emerald-50/60 border border-emerald-200 rounded-2xl p-4.5 space-y-4.5 animate-scale-in text-xs font-sans">
                <div className="flex items-center justify-between">
                  <span className="font-extrabold text-emerald-950 flex items-center gap-1.5 uppercase">
                    <Check className="w-4 h-4 text-emerald-600" /> Kết quả mẫu Doanh nghiệp
                  </span>
                  <button
                    type="button"
                    onClick={() => handleExportCSV("corp")}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold px-3 py-1.5 rounded-lg text-[10.5px] transition-all flex items-center gap-1 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" /> Tải về CSV
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-3 text-center">
                  <div className="bg-white border border-emerald-100 p-3 rounded-xl shadow-sm">
                    <span className="block text-[9.5px] font-bold text-slate-500 uppercase tracking-wider">MẪU CHÍNH THỨC</span>
                    <span className="block text-xl font-black text-emerald-600 font-mono mt-0.5">{corpSelectedList.length}</span>
                  </div>
                  <div className="bg-white border border-emerald-100 p-3 rounded-xl shadow-sm">
                    <span className="block text-[9.5px] font-bold text-slate-500 uppercase tracking-wider">MẪU DỰ PHÒNG</span>
                    <span className="block text-xl font-black text-orange-500 font-mono mt-0.5">{corpBackupList.length}</span>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                    <input
                      type="text"
                      placeholder="Tìm MST, tên doanh nghiệp..."
                      value={corpSearchTerm}
                      onChange={(e) => setCorpSearchTerm(e.target.value)}
                      className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-1.5 text-[11px] text-slate-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 shadow-inner"
                    />
                  </div>

                  <div className="overflow-x-auto border border-slate-100 rounded-lg">
                    <table className="w-full text-left text-[10px] text-slate-700">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-100">
                          <th className="p-2 font-bold uppercase">ID/MST</th>
                          <th className="p-2 font-bold uppercase">Tên Doanh nghiệp</th>
                          <th className="p-2 font-bold uppercase">Mã Xã</th>
                          <th className="p-2 font-bold uppercase">Mã Ngành</th>
                          <th className="p-2 font-bold text-right uppercase">Doanh thu</th>
                        </tr>
                      </thead>
                      <tbody>
                        {displayedCorpList.map((item, idx) => (
                          <tr key={idx} className="border-b border-slate-50 hover:bg-emerald-50/20">
                            <td className="p-2 font-mono font-semibold text-slate-900">{item.id}</td>
                            <td className="p-2 truncate max-w-[140px]" title={item.name}>{item.name}</td>
                            <td className="p-2 text-slate-500">{item.xaCode}</td>
                            <td className="p-2 font-mono text-indigo-700 font-bold">{item.vsicFull}</td>
                            <td className="p-2 text-right font-mono text-emerald-700 font-bold">{item.revenue.toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-[10px] text-slate-400 italic text-center">Hiển thị tối đa 50 dòng kết quả đầu tiên.</p>
                </div>
              </div>
            ) : (
              <div className="bg-slate-50 border border-slate-200/60 p-5 rounded-2xl text-center space-y-1.5">
                <span className="text-xl">📊</span>
                <p className="text-[11.5px] font-bold text-slate-600">Chưa thực hiện lọc mẫu doanh nghiệp</p>
                <p className="text-[10.5px] text-slate-400 max-w-sm mx-auto leading-normal">
                  Vui lòng nạp tệp (hoặc dùng DB chính), chỉ định cột <b>Mã ngành & Doanh thu</b> ở trên, và ấn nút Chạy lọc để hiển thị danh sách kết quả.
                </p>
              </div>
            )}
          </div>
        </div>

        {/* ====================================================================== */}
        {/* RIGHT COLUMN: INDIVIDUAL SAMPLING (HỘ CÁ THỂ) */}
        {/* ====================================================================== */}
        <div className="bg-white border-2 border-amber-100 rounded-3xl p-6 shadow-lg flex flex-col justify-between space-y-6 relative hover:border-amber-300 transition-all">
          <div className="absolute top-4 right-4 bg-amber-50 text-amber-700 text-[10px] font-black uppercase tracking-widest px-2.5 py-1 rounded-md border border-amber-100">
            Màn 2: Hộ cá thể
          </div>

          <div className="space-y-6">
            <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
              <div className="p-2.5 bg-amber-50 text-amber-600 rounded-2xl border border-amber-100">
                <Users className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black text-amber-950 uppercase tracking-wider">🏡 PHÂN HỆ HỘ CÁ THỂ</h3>
                <p className="text-[11px] text-slate-500">Gom nhóm Địa bàn xã &amp; Mã ngành cấp 2, chọn mẫu định mức Chuẩn GSO.</p>
              </div>
            </div>

            {/* BƯỚC 1: CHỌN NGUỒN DỮ LIỆU */}
            <div className="space-y-3">
              <label className="block text-xs font-extrabold text-amber-950 tracking-wide uppercase">
                Bước 1: Chọn Nguồn Dữ Liệu Hộ Cá Thể
              </label>
              
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => {
                    setIndSourceType("main");
                    setIndHasFiltered(false);
                  }}
                  className={`px-4 py-3 rounded-xl border text-xs font-bold transition-all text-center flex items-center justify-center gap-1.5 cursor-pointer ${
                    indSourceType === "main"
                      ? "bg-amber-600 text-white border-amber-700 shadow-md"
                      : "bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200"
                  }`}
                >
                  <Database className="w-4 h-4" /> Dữ liệu chính ({mainData.length.toLocaleString()})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setIndSourceType("uploaded");
                    setIndHasFiltered(false);
                  }}
                  className={`px-4 py-3 rounded-xl border text-xs font-bold transition-all text-center flex items-center justify-center gap-1.5 cursor-pointer ${
                    indSourceType === "uploaded"
                      ? "bg-amber-600 text-white border-amber-700 shadow-md"
                      : "bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200"
                  }`}
                >
                  <FileUp className="w-4 h-4" /> Tự nạp tệp riêng {sampIndFileName ? "✔️" : ""}
                </button>
              </div>

              {indSourceType === "uploaded" && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                  <label className="bg-white hover:bg-amber-50/20 border-2 border-dashed border-amber-200 text-amber-950 font-bold text-xs p-3.5 rounded-xl transition-all flex flex-col items-center justify-center gap-1.5 cursor-pointer">
                    <FileUp className="w-5 h-5 text-amber-500" />
                    <span>Nạp tệp Excel/CSV Hộ cá thể</span>
                    <input
                      type="file"
                      accept=".xlsx, .xls, .csv"
                      onChange={(e) => handleFileUpload(e, "ind")}
                      className="hidden"
                    />
                  </label>
                  {sampIndFileName && (
                    <div className="text-[11px] text-amber-900 bg-white border border-amber-100 p-2.5 rounded-lg flex items-center justify-between">
                      <span className="truncate font-semibold max-w-[180px]">📁 {sampIndFileName}</span>
                      <span className="font-mono font-bold bg-amber-50 text-amber-700 px-1.5 py-0.5 rounded">
                        {sampIndData.length} dòng
                      </span>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* BƯỚC 2: CHỈ ĐỊNH CÁC CỘT (STRICTLY MANUAL) */}
            <div className="bg-amber-50/40 border border-amber-100 rounded-2xl p-4.5 space-y-4">
              <div className="flex items-start gap-2">
                <AlertTriangle className="w-4.5 h-4.5 text-amber-600 shrink-0 mt-0.5" />
                <div className="space-y-0.5">
                  <span className="text-xs font-black text-amber-950 block">CHỈ ĐỊNH CỘT DỮ LIỆU HỘ CÁ THỂ</span>
                  <span className="text-[10.5px] text-slate-500 block">Không được tự động điền sẵn. Hãy chọn đúng tên cột tương ứng trong bảng dữ liệu:</span>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="block text-slate-700 font-extrabold mb-1">1. Cột Mã Ngành VSIC *</label>
                  <select
                    value={indManganhCol}
                    onChange={(e) => {
                      setIndManganhCol(e.target.value);
                      setIndHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 outline-none shadow-sm font-sans"
                  >
                    <option value="">-- Click để Chọn Cột Mã Ngành --</option>
                    {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-700 font-extrabold mb-1">2. Cột Doanh thu / Sản lượng *</label>
                  <select
                    value={indDoanhThuCol}
                    onChange={(e) => {
                      setIndDoanhThuCol(e.target.value);
                      setIndHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 outline-none shadow-sm font-sans"
                  >
                    <option value="">-- Click để Chọn Cột Doanh Thu --</option>
                    {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-500 font-bold mb-1">3. Cột Định danh (MST/Tên chủ) (Tùy chọn)</label>
                  <select
                    value={indIdCol}
                    onChange={(e) => {
                      setIndIdCol(e.target.value);
                      setIndHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột định danh Hộ --</option>
                    {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-500 font-bold mb-1">4. Cột Địa bàn xã / Phường (Tùy chọn)</label>
                  <select
                    value={indXaCol}
                    onChange={(e) => {
                      setIndXaCol(e.target.value);
                      setIndHasFiltered(false);
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột xã/phương --</option>
                    {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>
            </div>

            {/* BƯỚC 3: CẤU HÌNH THAM SỐ GSO CHỌN MẪU */}
            <div className="space-y-4">
              <div className="flex justify-between items-center">
                <label className="block text-xs font-extrabold text-amber-950 tracking-wide uppercase">
                  Bước 3: Định Mức Phân Tầng Chọn Mẫu (Chuẩn GSO)
                </label>
                <button
                  type="button"
                  onClick={() => {
                    setIndSize1To5Value(5);
                    setIndSize1To5All(true);
                    setIndSize6To100Value(5);
                    setIndSize101To1000Value(8);
                    setIndSize1001PlusPercent(1);
                    setIndTransportPercent(1.5);
                    setIndTransportMaxCap(50);
                    setIndHasFiltered(false);
                  }}
                  className="text-[10px] text-amber-800 font-bold hover:underline"
                >
                  🔄 Khôi phục mặc định
                </button>
              </div>

              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4.5 space-y-3 text-xs font-sans">
                {/* 1 - 5 cơ sở */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800">1. Nhóm cực nhỏ (1-5 cơ sở):</span>
                    <label className="flex items-center gap-1 font-bold text-[10px] cursor-pointer text-amber-800">
                      <input
                        type="checkbox"
                        checked={indSize1To5All}
                        onChange={(e) => {
                          setIndSize1To5All(e.target.checked);
                          setIndHasFiltered(false);
                        }}
                        className="accent-amber-600 rounded"
                      />
                      <span>Lấy hết 100%</span>
                    </label>
                  </div>
                  {!indSize1To5All && (
                    <div className="flex items-center gap-3">
                      <input
                        type="range"
                        min="1"
                        max="5"
                        value={indSize1To5Value}
                        onChange={(e) => {
                          setIndSize1To5Value(parseInt(e.target.value));
                          setIndHasFiltered(false);
                        }}
                        className="flex-1 accent-amber-600"
                      />
                      <span className="font-mono font-bold bg-white border px-1.5 py-0.5 rounded text-[10.5px] shrink-0 text-amber-800">{indSize1To5Value} hộ</span>
                    </div>
                  )}
                </div>

                {/* 6 - 100 cơ sở */}
                <div className="space-y-1">
                  <div className="flex justify-between">
                    <span className="font-semibold text-slate-700">2. Nhóm nhỏ (6-100 cơ sở):</span>
                    <span className="font-mono font-bold text-amber-700">{indSize6To100Value} hộ</span>
                  </div>
                  <input
                    type="range"
                    min="1"
                    max="30"
                    value={indSize6To100Value}
                    onChange={(e) => {
                      setIndSize6To100Value(parseInt(e.target.value));
                      setIndHasFiltered(false);
                    }}
                    className="w-full accent-amber-600"
                  />
                </div>

                {/* 101 - 1000 cơ sở */}
                <div className="space-y-1">
                  <div className="flex justify-between">
                    <span className="font-semibold text-slate-700">3. Nhóm vừa (101-1000 cơ sở):</span>
                    <span className="font-mono font-bold text-amber-700">{indSize101To1000Value} hộ</span>
                  </div>
                  <input
                    type="range"
                    min="1"
                    max="100"
                    value={indSize101To1000Value}
                    onChange={(e) => {
                      setIndSize101To1000Value(parseInt(e.target.value));
                      setIndHasFiltered(false);
                    }}
                    className="w-full accent-amber-600"
                  />
                </div>

                {/* Lớn 1001+ */}
                <div className="space-y-1">
                  <div className="flex justify-between">
                    <span className="font-semibold text-slate-700">4. Nhóm lớn (&gt;1001 cơ sở):</span>
                    <span className="font-mono font-bold text-amber-700">{indSize1001PlusPercent}%</span>
                  </div>
                  <input
                    type="range"
                    min="0.1"
                    max="10"
                    step="0.1"
                    value={indSize1001PlusPercent}
                    onChange={(e) => {
                      setIndSize1001PlusPercent(parseFloat(e.target.value));
                      setIndHasFiltered(false);
                    }}
                    className="w-full accent-amber-600"
                  />
                </div>

                {/* Vận tải đặc thù */}
                <div className="bg-amber-50/40 p-2.5 rounded-xl border border-amber-200/50 space-y-1.5">
                  <div className="flex justify-between text-[11px] font-extrabold text-amber-900">
                    <span>📍 Riêng Vận tải (VSIC 49):</span>
                    <span>{indTransportPercent}% (Tối đa {indTransportMaxCap})</span>
                  </div>
                  <input
                    type="range"
                    min="0.5"
                    max="15"
                    step="0.1"
                    value={indTransportPercent}
                    onChange={(e) => {
                      setIndTransportPercent(parseFloat(e.target.value));
                      setIndHasFiltered(false);
                    }}
                    className="w-full accent-amber-700"
                  />
                  <input
                    type="range"
                    min="10"
                    max="150"
                    value={indTransportMaxCap}
                    onChange={(e) => {
                      setIndTransportMaxCap(parseInt(e.target.value));
                      setIndHasFiltered(false);
                    }}
                    className="w-full accent-amber-700"
                  />
                </div>

                {/* Nhóm ngành chọn lọc */}
                <div className="pt-2.5 border-t border-slate-200">
                  <span className="block text-slate-700 font-extrabold mb-1.5">Lọc nhóm ngành khảo sát:</span>
                  <div className="flex flex-wrap gap-2.5">
                    {["congnghiep", "thuongmai", "vantai", "anuongluutru"].map(sec => {
                      const labels: Record<string, string> = {
                        congnghiep: "Công nghiệp (05-33)",
                        thuongmai: "Thương mại (45-46)",
                        vantai: "Vận tải (49)",
                        anuongluutru: "Dịch vụ/Ăn uống (55-56)"
                      };
                      return (
                        <label key={sec} className="flex items-center gap-1.5 cursor-pointer text-[10.5px]">
                          <input
                            type="checkbox"
                            checked={indSectors.includes(sec)}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setIndSectors(prev => [...prev, sec]);
                              } else {
                                setIndSectors(prev => prev.filter(x => x !== sec));
                              }
                              setIndHasFiltered(false);
                            }}
                            className="accent-amber-600 rounded"
                          />
                          <span>{labels[sec]}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* CHẠY LỌC HỘ CÁ THỂ ACTION BUTTON */}
          <div className="pt-4 border-t border-slate-100 space-y-4">
            <button
              type="button"
              onClick={handleFilterInd}
              className="w-full py-4 px-6 bg-gradient-to-r from-amber-600 to-amber-800 hover:from-amber-700 hover:to-amber-900 text-white font-black rounded-2xl text-xs uppercase tracking-widest shadow-lg flex items-center justify-center gap-2 transition-all cursor-pointer hover:scale-[1.01]"
            >
              <Play className="w-4 h-4 fill-white" /> Thực Hiện Lọc &amp; Chọn Mẫu Hộ Cá Thể
            </button>

            {/* RESULT COMPARTMENT FOR HOUSEHOLD */}
            {indHasFiltered ? (
              <div className="bg-emerald-50/60 border border-emerald-200 rounded-2xl p-4.5 space-y-4.5 animate-scale-in text-xs font-sans">
                <div className="flex items-center justify-between">
                  <span className="font-extrabold text-emerald-950 flex items-center gap-1.5 uppercase">
                    <Check className="w-4 h-4 text-emerald-600" /> Kết quả mẫu Hộ cá thể
                  </span>
                  <button
                    type="button"
                    onClick={() => handleExportCSV("ind")}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold px-3 py-1.5 rounded-lg text-[10.5px] transition-all flex items-center gap-1 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" /> Tải về CSV
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-3 text-center">
                  <div className="bg-white border border-emerald-100 p-3 rounded-xl shadow-sm">
                    <span className="block text-[9.5px] font-bold text-slate-500 uppercase tracking-wider">MẪU CHÍNH THỨC</span>
                    <span className="block text-xl font-black text-emerald-600 font-mono mt-0.5">{indSelectedList.length}</span>
                  </div>
                  <div className="bg-white border border-emerald-100 p-3 rounded-xl shadow-sm">
                    <span className="block text-[9.5px] font-bold text-slate-500 uppercase tracking-wider">MẪU DỰ PHÒNG</span>
                    <span className="block text-xl font-black text-orange-500 font-mono mt-0.5">{indBackupList.length}</span>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                    <input
                      type="text"
                      placeholder="Tìm ID, tên chủ hộ..."
                      value={indSearchTerm}
                      onChange={(e) => setIndSearchTerm(e.target.value)}
                      className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-1.5 text-[11px] text-slate-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 shadow-inner"
                    />
                  </div>

                  <div className="overflow-x-auto border border-slate-100 rounded-lg">
                    <table className="w-full text-left text-[10px] text-slate-700">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-100">
                          <th className="p-2 font-bold uppercase">ID</th>
                          <th className="p-2 font-bold uppercase">Chủ hộ</th>
                          <th className="p-2 font-bold uppercase">Mã Xã</th>
                          <th className="p-2 font-bold uppercase">Mã Ngành</th>
                          <th className="p-2 font-bold text-right uppercase">Doanh thu</th>
                        </tr>
                      </thead>
                      <tbody>
                        {displayedIndList.map((item, idx) => (
                          <tr key={idx} className="border-b border-slate-50 hover:bg-emerald-50/20">
                            <td className="p-2 font-mono font-semibold text-slate-900">{item.id}</td>
                            <td className="p-2 truncate max-w-[140px]" title={item.name}>{item.name}</td>
                            <td className="p-2 text-slate-500">{item.xaCode}</td>
                            <td className="p-2 font-mono text-amber-700 font-bold">{item.vsicFull}</td>
                            <td className="p-2 text-right font-mono text-emerald-700 font-bold">{item.revenue.toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-[10px] text-slate-400 italic text-center">Hiển thị tối đa 50 dòng kết quả đầu tiên.</p>
                </div>
              </div>
            ) : (
              <div className="bg-slate-50 border border-slate-200/60 p-5 rounded-2xl text-center space-y-1.5">
                <span className="text-xl">📊</span>
                <p className="text-[11.5px] font-bold text-slate-600">Chưa thực hiện lọc mẫu hộ cá thể</p>
                <p className="text-[10.5px] text-slate-400 max-w-sm mx-auto leading-normal">
                  Vui lòng nạp tệp (hoặc dùng DB chính), chỉ định cột <b>Mã ngành & Doanh thu</b> ở trên, và ấn nút Chạy lọc để hiển thị danh sách kết quả.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* JOINT SYSTEM STATISTICS PANEL */}
      <div className="bg-indigo-50/60 border border-indigo-100 rounded-3xl p-6 shadow-md space-y-4 font-sans text-xs">
        <h4 className="text-sm font-black text-indigo-950 uppercase tracking-wider flex items-center gap-2">
          <Info className="w-4.5 h-4.5 text-indigo-600" /> THÔNG TIN KHẢO SÁT &amp; CÁCH TÍNH CHUẨN CỦA TỔNG CỤC THỐNG KÊ (GSO)
        </h4>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 leading-relaxed text-slate-600">
          <div className="space-y-2">
            <span className="text-[11px] font-black text-indigo-900 block uppercase">1. Quy Tắc Lọc Doanh Nghiệp (Cutoff 75%)</span>
            <p>
              Dữ liệu doanh nghiệp được sắp xếp theo doanh thu giảm dần trong từng địa bàn hoặc toàn tỉnh (tùy thuộc lựa chọn của bạn). 
              Các doanh nghiệp đứng đầu bảng có đóng góp doanh thu cộng dồn chiếm tối thiểu <b>75% tổng doanh thu</b> của nhóm ngành đó sẽ được đưa vào mẫu chính thức. 
              Các đơn vị còn lại đóng vai trò là mẫu dự phòng thay thế.
            </p>
          </div>
          <div className="space-y-2">
            <span className="text-[11px] font-black text-amber-900 block uppercase">2. Định mức Phân tầng Hộ cá thể</span>
            <p>
              Hộ cá thể kinh doanh trong xã được phân loại theo quy mô mật độ đơn vị trong từng nhóm xã &amp; ngành cấp 2 (VSIC):
            </p>
            <ul className="list-disc pl-4 space-y-1 mt-1 text-[11px]">
              <li><b>1 - 5 cơ sở:</b> Lấy toàn bộ 100% để đảm bảo phủ sóng địa bàn.</li>
              <li><b>6 - 100 cơ sở:</b> Chọn cố định tối đa 5 cơ sở doanh thu cao nhất.</li>
              <li><b>101 - 1000 cơ sở:</b> Chọn cố định 8 cơ sở.</li>
              <li><b>Từ 1001 cơ sở trở lên:</b> Chọn tỷ lệ 1% trên tổng quy mô đơn vị trong nhóm.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
