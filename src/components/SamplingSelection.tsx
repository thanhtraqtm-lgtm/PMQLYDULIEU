import React, { useState, useMemo, useEffect } from "react";
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
  Trash2
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
  mapping,
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
  // --- STATES FOR SAMPLING MODULE ---
  const [entCutoffPercent, setEntCutoffPercent] = useState<number>(75);
  const [entMinGroupSize, setEntMinGroupSize] = useState<number>(1);
  const [entForceStates, setEntForceStates] = useState<boolean>(true);
  const [entForceMonthly, setEntForceMonthly] = useState<boolean>(true);
  const [entGroupScope, setEntGroupScope] = useState<"province" | "xa">("province");
  const [indSamplingMode, setIndSamplingMode] = useState<"GSO" | "Custom">("GSO");
  const [indMaxCap, setIndMaxCap] = useState<number>(10);
  const [indCustomMode, setIndCustomMode] = useState<"fixed" | "percent">("fixed");
  const [indCustomCountValue, setIndCustomCountValue] = useState<number>(5);
  const [indCustomPercentValue, setIndCustomPercentValue] = useState<number>(10);

  // --- STATES FOR MANUAL TRIGGER & USER CONFIDENCE ---
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [calculationSuccess, setCalculationSuccess] = useState<boolean>(false);
  const [calculationDetails, setCalculationDetails] = useState<{
    totalSelected: number;
    totalBackup: number;
    time: string;
  } | null>(null);

  // States for dynamic sector filters (Công nghiệp, Thương mại, Vận tải, Lưu trú ăn uống)
  const [selectedSectors, setSelectedSectors] = useState<string[]>(["congnghiep"]);
  const [customSectorRange, setCustomSectorRange] = useState<string>("05-33");

  // Helper to check if a level 2 VSIC code matches selected sectors
  const isSectorSelected = (vsicL2: string): boolean => {
    const l2Num = parseInt(vsicL2, 10);
    if (isNaN(l2Num)) return false;

    let matched = false;
    if (selectedSectors.includes("congnghiep") && l2Num >= 5 && l2Num <= 33) {
      matched = true;
    }
    if (selectedSectors.includes("thuongmai") && l2Num >= 45 && l2Num <= 46) {
      matched = true;
    }
    if (selectedSectors.includes("vantai") && l2Num === 49) {
      matched = true;
    }
    if (selectedSectors.includes("anuongluutru") && (l2Num === 55 || l2Num === 56)) {
      matched = true;
    }
    if (selectedSectors.includes("custom") && customSectorRange) {
      const parts = customSectorRange.split(",").map(p => p.trim());
      for (const part of parts) {
        if (part.includes("-")) {
          const [startStr, endStr] = part.split("-").map(p => p.trim());
          const start = parseInt(startStr, 10);
          const end = parseInt(endStr, 10);
          if (!isNaN(start) && !isNaN(end) && l2Num >= start && l2Num <= end) {
            matched = true;
            break;
          }
        } else {
          const singleVal = parseInt(part, 10);
          if (!isNaN(singleVal) && l2Num === singleVal) {
            matched = true;
            break;
          }
        }
      }
    }
    return matched;
  };

  // Column configuration
  const [sampIdCol, setSampIdCol] = useState<string>("");
  const [sampXaCol, setSampXaCol] = useState<string>("");
  const [sampManganhCol, setSampManganhCol] = useState<string>("");
  const [sampDoanhThuCol, setSampDoanhThuCol] = useState<string>("");
  const [sampTypeCol, setSampTypeCol] = useState<string>("");
  const [sampFilterType, setSampFilterType] = useState<"all_ent" | "all_ind" | "by_col">("all_ent");
  const [sampTypeEnterpriseValue, setSampTypeEnterpriseValue] = useState<string>("DN");
  const [sampTypeHouseholdValue, setSampTypeHouseholdValue] = useState<string>("Hộ");

  // Search & Navigation
  const [sampSearchTerm, setSampSearchTerm] = useState<string>("");
  const [sampActiveDetailGroup, setSampActiveDetailGroup] = useState<string>("");

  const samplingColumns = useMemo(() => {
    const colsSet = new Set<string>();
    columns.forEach(c => colsSet.add(c));
    if (sampCorpData.length > 0) {
      Object.keys(sampCorpData[0] || {}).forEach(c => colsSet.add(c));
    }
    if (sampIndData.length > 0) {
      Object.keys(sampIndData[0] || {}).forEach(c => colsSet.add(c));
    }
    return Array.from(colsSet);
  }, [columns, sampCorpData, sampIndData]);

  // Sync columns with mapping whenever columns or mapping change
  useEffect(() => {
    if (mapping.idCol && !sampIdCol) setSampIdCol(mapping.idCol);
    if (mapping.xa && !sampXaCol) setSampXaCol(mapping.xa);
    if (mapping.manganh && !sampManganhCol) setSampManganhCol(mapping.manganh);
    if (mapping.doanhthu && !sampDoanhThuCol) setSampDoanhThuCol(mapping.doanhthu);
  }, [mapping, samplingColumns]);

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

  const handleSamplingFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "corp" | "ind") => {
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
          if (!text) throw new Error("Không thể đọc nội dung tệp tin!");

          const data = parseCSV(text);
          if (data.length === 0) {
            alert("Tệp trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const cols = Object.keys(data[0] as any);
          if (type === "corp") {
            setSampCorpData(data);
            setSampCorpFileName(file.name);
          } else {
            setSampIndData(data);
            setSampIndFileName(file.name);
          }

          if (cols.length > 0) {
            cols.forEach(c => {
              const cLow = c.toLowerCase();
              if ((cLow.includes("mst") || cLow.includes("mã số thuế") || cLow.includes("định danh") || cLow.includes("id")) && !sampIdCol) {
                setSampIdCol(c);
              }
              if ((cLow.includes("xã") || cLow.includes("địa bàn") || cLow.includes("mã xã")) && !sampXaCol) {
                setSampXaCol(c);
              }
              if ((cLow.includes("ngành") || cLow.includes("mã ngành") || cLow.includes("vsic")) && !sampManganhCol) {
                setSampManganhCol(c);
              }
              if ((cLow.includes("doanh thu") || cLow.includes("sản lượng") || cLow.includes("doanhthu")) && !sampDoanhThuCol) {
                setSampDoanhThuCol(c);
              }
            });
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
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          if (data.length === 0) {
            alert("Tệp Excel trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const cols = Object.keys(data[0] as any);
          if (type === "corp") {
            setSampCorpData(data);
            setSampCorpFileName(file.name);
          } else {
            setSampIndData(data);
            setSampIndFileName(file.name);
          }

          if (cols.length > 0) {
            cols.forEach(c => {
              const cLow = c.toLowerCase();
              if ((cLow.includes("mst") || cLow.includes("mã số thuế") || cLow.includes("định danh") || cLow.includes("id")) && !sampIdCol) {
                setSampIdCol(c);
              }
              if ((cLow.includes("xã") || cLow.includes("địa bàn") || cLow.includes("mã xã")) && !sampXaCol) {
                setSampXaCol(c);
              }
              if ((cLow.includes("ngành") || cLow.includes("mã ngành") || cLow.includes("vsic")) && !sampManganhCol) {
                setSampManganhCol(c);
              }
              if ((cLow.includes("doanh thu") || cLow.includes("sản lượng") || cLow.includes("doanhthu")) && !sampDoanhThuCol) {
                setSampDoanhThuCol(c);
              }
            });
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

  const enterpriseList = useMemo(() => {
    const sourceData = sampCorpData.length > 0 ? sampCorpData : mainData;
    if (sourceData.length === 0) return [];

    const processRow = (row: any, index: number) => {
      const manganhVal = String(row[sampManganhCol] || "").trim();
      const cleanDigits = manganhVal.replace(/\D/g, "");
      const vsicL2 = cleanDigits.slice(0, 2);
      
      if (!isSectorSelected(vsicL2)) return null;
      
      const idVal = String(row[sampIdCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
      const nameVal = String(row["Tên doanh nghiệp"] || row["Tên đơn vị"] || row["Tên"] || row["Tên hộ"] || "Doanh nghiệp " + index);
      const xaVal = String(row[sampXaCol] || "30000");
      const revVal = parseFloat(String(row[sampDoanhThuCol] || "0").replace(/,/g, "")) || 0;
      const vsicL2Name = vsicRawData[vsicL2] || `Ngành cấp 2 (${vsicL2})`;
      
      return {
        id: idVal,
        name: nameVal,
        xaCode: xaVal,
        vsicL2,
        vsicL2Name,
        vsicFull: manganhVal,
        revenue: revVal,
        type: "DN",
        status: "active",
        originalRow: row
      };
    };

    if (sampCorpData.length > 0) {
      return sampCorpData
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }

    if (sampFilterType === "all_ent") {
      return mainData
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }
    if (sampFilterType === "by_col" && sampTypeCol) {
      return mainData
        .filter(row => {
          const val = String(row[sampTypeCol] || "").trim().toLowerCase();
          const entMatch = sampTypeEnterpriseValue.trim().toLowerCase();
          return val === entMatch || val.includes("doanh nghiệp") || val.includes("dn") || val.includes("công ty");
        })
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }
    return [];
  }, [mainData, sampCorpData, sampFilterType, sampTypeCol, sampTypeEnterpriseValue, sampIdCol, sampXaCol, sampManganhCol, sampDoanhThuCol, selectedSectors, customSectorRange]);

  const individualList = useMemo(() => {
    const sourceData = sampIndData.length > 0 ? sampIndData : mainData;
    if (sourceData.length === 0) return [];

    const processRow = (row: any, index: number) => {
      const manganhVal = String(row[sampManganhCol] || "").trim();
      const cleanDigits = manganhVal.replace(/\D/g, "");
      const vsicL2 = cleanDigits.slice(0, 2);
      
      if (!isSectorSelected(vsicL2)) return null;
      
      const idVal = String(row[sampIdCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
      const nameVal = String(row["Tên hộ"] || row["Tên đơn vị"] || row["Tên"] || "Hộ " + index);
      const xaVal = String(row[sampXaCol] || "30000");
      const revVal = parseFloat(String(row[sampDoanhThuCol] || "0").replace(/,/g, "")) || 0;
      const vsicL2Name = vsicRawData[vsicL2] || `Ngành cấp 2 (${vsicL2})`;
      
      return {
        id: idVal,
        name: nameVal,
        xaCode: xaVal,
        vsicL2,
        vsicL2Name,
        vsicFull: manganhVal,
        revenue: revVal,
        type: "HO",
        status: "active",
        originalRow: row
      };
    };

    if (sampIndData.length > 0) {
      return sampIndData
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }

    if (sampFilterType === "all_ind") {
      return mainData
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }
    if (sampFilterType === "by_col" && sampTypeCol) {
      return mainData
        .filter(row => {
          const val = String(row[sampTypeCol] || "").trim().toLowerCase();
          const indMatch = sampTypeHouseholdValue.trim().toLowerCase();
          return val === indMatch || val.includes("hộ") || val.includes("cá thể") || val.includes("hct");
        })
        .map((row, index) => processRow(row, index))
        .filter((item): item is NonNullable<typeof item> => item !== null);
    }
    return [];
  }, [mainData, sampIndData, sampFilterType, sampTypeCol, sampTypeHouseholdValue, sampIdCol, sampXaCol, sampManganhCol, sampDoanhThuCol, selectedSectors, customSectorRange]);

  const corporateSamplingResults = useMemo(() => {
    const results: {
      groups: Record<string, {
        xaCode: string;
        vsicL2: string;
        totalN: number;
        totalRevenue: number;
        targetCutoff: number;
        runningSum: number;
        selectedCandidates: any[];
        backupCandidates: any[];
        isAlwaysSelectedGroup: boolean;
      }>;
      selectedIDs: Set<string>;
      backupIDs: Set<string>;
    } = { groups: {}, selectedIDs: new Set(), backupIDs: new Set() };

    if (enterpriseList.length === 0) return results;

    // Map to hold temporary selection details for each enterprise
    const annotatedEnterprisesMap = new Map<string, {
      selectionType: string;
      isSelected: boolean;
      cumulativeRevenuePercent: number;
    }>();

    if (entGroupScope === "province") {
      // Group purely by vsicL2 for province-wide selection
      const provGroups: Record<string, any[]> = {};
      enterpriseList.forEach(ent => {
        const key = ent.vsicL2;
        if (!provGroups[key]) provGroups[key] = [];
        provGroups[key].push(ent);
      });

      Object.entries(provGroups).forEach(([vsicL2, list]) => {
        const sortedByRev = [...list].sort((a, b) => b.revenue - a.revenue);
        const totalN = sortedByRev.length;
        const isAlwaysSelectedGroup = totalN <= entMinGroupSize;

        const alwaysSelectedCandidates = sortedByRev.filter(ent => {
          if (entForceStates && ent.originalRow) {
            const isDnnn = Object.entries(ent.originalRow).some(([key, val]) => {
              const kLow = key.toLowerCase();
              const vLow = String(val || "").toLowerCase();
              const matchesKey = kLow.includes("loại hình") || kLow.includes("loại doanh nghiệp") || kLow.includes("hình thức") || kLow.includes("dnnn") || kLow.includes("loai hinh") || kLow.includes("loai doanh nghiep") || kLow.includes("hinh thuc");
              const matchesVal = vLow.includes("nhà nước") || vLow.includes("nha nuoc") || vLow === "có" || vLow === "co" || vLow === "yes";
              return matchesKey && matchesVal;
            });
            if (isDnnn) return true;
          }
          if (isAlwaysSelectedGroup) return true;
          return false;
        });

        const alwaysSelectedIDs = new Set(alwaysSelectedCandidates.map(c => c.id));
        const nonAlwaysCandidates = sortedByRev.filter(ent => !alwaysSelectedIDs.has(ent.id));
        const totalGroupRevenue = sortedByRev.reduce((sum, item) => sum + item.revenue, 0);
        const targetCutoffRevenue = totalGroupRevenue * (entCutoffPercent / 100);

        let runningSum = 0;

        alwaysSelectedCandidates.forEach(ent => {
          runningSum += ent.revenue;
          annotatedEnterprisesMap.set(ent.id, {
            selectionType: "Chọn toàn bộ (DNNN / Ngưỡng tối thiểu)",
            isSelected: true,
            cumulativeRevenuePercent: totalGroupRevenue > 0 ? (runningSum / totalGroupRevenue) * 100 : 0
          });
          results.selectedIDs.add(ent.id);
        });

        nonAlwaysCandidates.forEach((ent) => {
          const isCentralForce = entForceMonthly && ent.originalRow && Object.entries(ent.originalRow).some(([key, val]) => {
            const kLow = key.toLowerCase();
            const vLow = String(val || "").toLowerCase();
            const matchesKey = kLow.includes("mẫu trung ương") || kLow.includes("mẫu tu") || kLow.includes("mau trung uong") || kLow.includes("mau tu") || kLow.includes("mẫu t.ư");
            const matchesVal = vLow === "có" || vLow === "co" || vLow === "yes" || vLow === "1" || vLow === "true" || vLow.includes("trung ương") || vLow.includes("trung uong");
            return matchesKey && matchesVal;
          });

          if (isCentralForce || runningSum < targetCutoffRevenue) {
            if (ent.status === "active") {
              runningSum += ent.revenue;
              annotatedEnterprisesMap.set(ent.id, {
                selectionType: isCentralForce ? "Ưu tiên mẫu trung ương" : "Chọn theo doanh thu lũy kế",
                isSelected: true,
                cumulativeRevenuePercent: totalGroupRevenue > 0 ? (runningSum / totalGroupRevenue) * 100 : 0
              });
              results.selectedIDs.add(ent.id);
            }
          } else {
            if (ent.status === "active") {
              annotatedEnterprisesMap.set(ent.id, {
                selectionType: "Dự phòng",
                isSelected: false,
                cumulativeRevenuePercent: totalGroupRevenue > 0 ? ((runningSum + ent.revenue) / totalGroupRevenue) * 100 : 0
              });
              results.backupIDs.add(ent.id);
            }
          }
        });
      });
    } else {
      // Group by `${xaCode}-${vsicL2}` for commune-level selection (Old default)
      const groups: Record<string, any[]> = {};
      enterpriseList.forEach(ent => {
        const key = `${ent.xaCode}-${ent.vsicL2}`;
        if (!groups[key]) groups[key] = [];
        groups[key].push(ent);
      });

      Object.entries(groups).forEach(([groupKey, list]) => {
        const sortedByRev = [...list].sort((a, b) => b.revenue - a.revenue);
        const totalN = sortedByRev.length;
        const isAlwaysSelectedGroup = totalN <= entMinGroupSize;

        const alwaysSelectedCandidates = sortedByRev.filter(ent => {
          if (entForceStates && ent.originalRow) {
            const isDnnn = Object.entries(ent.originalRow).some(([key, val]) => {
              const kLow = key.toLowerCase();
              const vLow = String(val || "").toLowerCase();
              const matchesKey = kLow.includes("loại hình") || kLow.includes("loại doanh nghiệp") || kLow.includes("hình thức") || kLow.includes("dnnn") || kLow.includes("loai hinh") || kLow.includes("loai doanh nghiep") || kLow.includes("hinh thuc");
              const matchesVal = vLow.includes("nhà nước") || vLow.includes("nha nuoc") || vLow === "có" || vLow === "co" || vLow === "yes";
              return matchesKey && matchesVal;
            });
            if (isDnnn) return true;
          }
          if (isAlwaysSelectedGroup) return true;
          return false;
        });

        const alwaysSelectedIDs = new Set(alwaysSelectedCandidates.map(c => c.id));
        const nonAlwaysCandidates = sortedByRev.filter(ent => !alwaysSelectedIDs.has(ent.id));
        const totalGroupRevenue = sortedByRev.reduce((sum, item) => sum + item.revenue, 0);
        const targetCutoffRevenue = totalGroupRevenue * (entCutoffPercent / 100);

        let runningSum = 0;

        alwaysSelectedCandidates.forEach(ent => {
          runningSum += ent.revenue;
          annotatedEnterprisesMap.set(ent.id, {
            selectionType: "Chọn toàn bộ (DNNN / Ngưỡng tối thiểu)",
            isSelected: true,
            cumulativeRevenuePercent: totalGroupRevenue > 0 ? (runningSum / totalGroupRevenue) * 100 : 0
          });
          results.selectedIDs.add(ent.id);
        });

        nonAlwaysCandidates.forEach((ent) => {
          const isCentralForce = entForceMonthly && ent.originalRow && Object.entries(ent.originalRow).some(([key, val]) => {
            const kLow = key.toLowerCase();
            const vLow = String(val || "").toLowerCase();
            const matchesKey = kLow.includes("mẫu trung ương") || kLow.includes("mẫu tu") || kLow.includes("mau trung uong") || kLow.includes("mau tu") || kLow.includes("mẫu t.ư");
            const matchesVal = vLow === "có" || vLow === "co" || vLow === "yes" || vLow === "1" || vLow === "true" || vLow.includes("trung ương") || vLow.includes("trung uong");
            return matchesKey && matchesVal;
          });

          if (isCentralForce || runningSum < targetCutoffRevenue) {
            if (ent.status === "active") {
              runningSum += ent.revenue;
              annotatedEnterprisesMap.set(ent.id, {
                selectionType: isCentralForce ? "Ưu tiên mẫu trung ương" : "Chọn theo doanh thu lũy kế",
                isSelected: true,
                cumulativeRevenuePercent: totalGroupRevenue > 0 ? (runningSum / totalGroupRevenue) * 100 : 0
              });
              results.selectedIDs.add(ent.id);
            }
          } else {
            if (ent.status === "active") {
              annotatedEnterprisesMap.set(ent.id, {
                selectionType: "Dự phòng",
                isSelected: false,
                cumulativeRevenuePercent: totalGroupRevenue > 0 ? ((runningSum + ent.revenue) / totalGroupRevenue) * 100 : 0
              });
              results.backupIDs.add(ent.id);
            }
          }
        });
      });
    }

    // Always group back into `${ent.xaCode}-${ent.vsicL2}` so that presentation layers remain unchanged
    const displayGroups: Record<string, any[]> = {};
    enterpriseList.forEach(ent => {
      const key = `${ent.xaCode}-${ent.vsicL2}`;
      if (!displayGroups[key]) displayGroups[key] = [];
      const ann = annotatedEnterprisesMap.get(ent.id) || { selectionType: "Dự phòng", isSelected: false, cumulativeRevenuePercent: 0 };
      displayGroups[key].push({
        ...ent,
        selectionType: ann.selectionType,
        cumulativeRevenuePercent: ann.cumulativeRevenuePercent,
        isSelected: ann.isSelected
      });
    });

    Object.entries(displayGroups).forEach(([groupKey, list]) => {
      const [xaCode, vsicL2] = groupKey.split("-");
      const sortedByRev = [...list].sort((a, b) => b.revenue - a.revenue);
      const totalN = sortedByRev.length;
      const isAlwaysSelectedGroup = totalN <= entMinGroupSize;
      const totalGroupRevenue = sortedByRev.reduce((sum, item) => sum + item.revenue, 0);

      const selectedCandidates = sortedByRev.filter(ent => ent.isSelected);
      const backupCandidates = sortedByRev.filter(ent => !ent.isSelected);

      results.groups[groupKey] = {
        xaCode,
        vsicL2,
        totalN,
        totalRevenue: totalGroupRevenue,
        targetCutoff: totalGroupRevenue * (entCutoffPercent / 100),
        runningSum: selectedCandidates.reduce((sum, item) => sum + item.revenue, 0),
        selectedCandidates,
        backupCandidates,
        isAlwaysSelectedGroup
      };
    });

    return results;
  }, [enterpriseList, entCutoffPercent, entMinGroupSize, entForceStates, entForceMonthly, entGroupScope]);

  const individualSamplingResults = useMemo(() => {
    const results: {
      groups: Record<string, {
        xaCode: string;
        vsicL2: string;
        totalN: number;
        totalRevenue: number;
        targetSampleSize: number;
        selectedCandidates: any[];
        backupCandidates: any[];
      }>;
      selectedIDs: Set<string>;
      backupIDs: Set<string>;
    } = { groups: {}, selectedIDs: new Set(), backupIDs: new Set() };

    if (individualList.length === 0) return results;

    const groups: Record<string, any[]> = {};
    individualList.forEach(ind => {
      const key = `${ind.xaCode}-${ind.vsicL2}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(ind);
    });

    Object.entries(groups).forEach(([groupKey, list]) => {
      const [xaCode, vsicL2] = groupKey.split('-');
      const sortedByRev = [...list].sort((a, b) => b.revenue - a.revenue);
      const totalN = sortedByRev.length;
      const totalGroupRevenue = sortedByRev.reduce((sum, item) => sum + item.revenue, 0);

      let targetSampleSize = 0;
      if (indSamplingMode === 'GSO') {
        if (totalN <= 5) {
          targetSampleSize = totalN;
        } else if (totalN <= 100) {
          targetSampleSize = Math.min(5, totalN);
        } else if (totalN <= 1000) {
          targetSampleSize = Math.min(8, totalN);
        } else {
          targetSampleSize = Math.min(indMaxCap, Math.max(1, Math.round(totalN * 0.01)));
        }
      } else {
        if (indCustomMode === 'fixed') {
          targetSampleSize = Math.min(indCustomCountValue, totalN);
        } else {
          targetSampleSize = Math.min(indMaxCap, Math.max(1, Math.round(totalN * (indCustomPercentValue / 100))));
        }
      }

      const selectedCandidatesList: any[] = [];
      const backupCandidatesList: any[] = [];

      let selectedActiveCount = 0;
      sortedByRev.forEach((ind) => {
        if (ind.status === 'active') {
          if (selectedActiveCount < targetSampleSize) {
            selectedCandidatesList.push({
              ...ind,
              selectionType: "Mẫu chính thức",
              cumulativeRevenuePercent: 100
            });
            results.selectedIDs.add(ind.id);
            selectedActiveCount++;
          } else {
            backupCandidatesList.push({
              ...ind,
              selectionType: "Mẫu dự phòng",
              cumulativeRevenuePercent: 0
            });
            results.backupIDs.add(ind.id);
          }
        }
      });

      results.groups[groupKey] = {
        xaCode,
        vsicL2,
        totalN,
        totalRevenue: totalGroupRevenue,
        targetSampleSize,
        selectedCandidates: selectedCandidatesList,
        backupCandidates: backupCandidatesList
      };
    });
    return results;
  }, [individualList, indSamplingMode, indCustomMode, indCustomCountValue, indCustomPercentValue, indMaxCap]);

  const allSamplingGroups = useMemo(() => {
    const keys = new Set([
      ...Object.keys(corporateSamplingResults.groups),
      ...Object.keys(individualSamplingResults.groups)
    ]);
    
    return Array.from(keys).map(key => {
      const [xaCode, vsicL2] = key.split("-");
      const corpGrp = corporateSamplingResults.groups[key];
      const indGrp = individualSamplingResults.groups[key];
      
      const totalN = (corpGrp?.totalN || 0) + (indGrp?.totalN || 0);
      const totalRevenue = (corpGrp?.totalRevenue || 0) + (indGrp?.totalRevenue || 0);
      const selectedCount = (corpGrp?.selectedCandidates.length || 0) + (indGrp?.selectedCandidates.length || 0);
      const backupCount = (corpGrp?.backupCandidates.length || 0) + (indGrp?.backupCandidates.length || 0);
      const vsicL2Name = vsicRawData[vsicL2] || `Ngành công nghiệp cấp 2 (${vsicL2})`;
      
      return {
        key,
        xaCode,
        vsicL2,
        vsicL2Name,
        totalN,
        totalRevenue,
        selectedCount,
        backupCount,
        corpGrp,
        indGrp
      };
    });
  }, [corporateSamplingResults, individualSamplingResults]);

  const filteredSamplingGroups = useMemo(() => {
    if (!sampSearchTerm) return allSamplingGroups;
    const term = sampSearchTerm.toLowerCase();
    return allSamplingGroups.filter(g => 
      g.xaCode.toLowerCase().includes(term) || 
      g.vsicL2.toLowerCase().includes(term)
    );
  }, [allSamplingGroups, sampSearchTerm]);

  const handleManualRun = () => {
    setIsCalculating(true);
    setCalculationSuccess(false);
    setTimeout(() => {
      setIsCalculating(false);
      setCalculationSuccess(true);
      
      const selectedCount = corporateSamplingResults.selectedIDs.size + individualSamplingResults.selectedIDs.size;
      const backupCount = corporateSamplingResults.backupIDs.size + individualSamplingResults.backupIDs.size;
      
      setCalculationDetails({
        totalSelected: selectedCount,
        totalBackup: backupCount,
        time: new Date().toLocaleTimeString('vi-VN')
      });
      
      setTimeout(() => {
        setCalculationSuccess(false);
      }, 8000);
    }, 700);
  };

  return (
    <div className="space-y-6 animate-fade-in font-sans text-slate-800">
      {/* Header Tab */}
      <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="space-y-1">
          <h2 className="text-xl font-bold tracking-tight text-slate-900 flex items-center gap-2 font-sans">
            <FileCheck className="w-5 h-5 text-indigo-600" /> BỘ CHỌN MẪU KHẢO SÁT CHUYÊN ĐỀ (DOANH NGHIỆP &amp; HỘ CÁ THỂ)
          </h2>
          <p className="text-xs text-slate-500 font-medium font-sans">
            Tự động lọc các nhóm ngành được chọn. Chọn mẫu theo quy mô doanh thu lũy kế {entCutoffPercent}% (Doanh nghiệp) và định mức quy mô chuẩn của Tổng cục Thống kê (Hộ cá thể).
          </p>
        </div>
        {(mainData.length > 0 || sampCorpData.length > 0 || sampIndData.length > 0) && (
          <button
            onClick={() => {
              try {
                const sourceData = sampCorpData.length > 0 || sampIndData.length > 0
                  ? [...sampCorpData, ...sampIndData]
                  : mainData;

                if (sourceData.length === 0) {
                  alert("Chưa có dữ liệu nguồn để xuất!");
                  return;
                }
                
                const headers = [
                  "Mã số định danh/MST",
                  "Tên đơn vị/đối tượng",
                  "Mã xã/Địa bàn",
                  "Mã ngành VSIC",
                  "Mã ngành Cấp 2",
                  "Tên ngành Cấp 2",
                  "Doanh thu/Sản lượng",
                  "Phân loại",
                  "Trạng thái chọn mẫu",
                  "Chi tiết chọn mẫu",
                  "Doanh thu lũy kế (%)"
                ];
                
                const rows = sourceData
                  .map((row, index) => {
                    const idVal = String(row[sampIdCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
                    const nameVal = String(row["Tên doanh nghiệp"] || row["Tên đơn vị"] || row["Tên"] || row["Tên hộ"] || "Đơn vị " + index);
                    const xaVal = String(row[sampXaCol] || "30000");
                    const manganhVal = String(row[sampManganhCol] || "").trim();
                    const cleanDigits = manganhVal.replace(/\D/g, "");
                    const vsicL2 = cleanDigits.slice(0, 2) || "00";
                    const revVal = parseFloat(String(row[sampDoanhThuCol] || "0").replace(/,/g, "")) || 0;
                    
                    const entItem = enterpriseList.find(e => e.id === idVal);
                    const indItem = individualList.find(i => i.id === idVal);
                    
                    if (!entItem && !indItem) return null;

                    let classification = "Không xác định";
                    let samplingStatus = "Không được chọn";
                    let samplingDetail = "Không lọt mẫu";
                    let cumulativePercent = 0;
                    const vsicL2Name = vsicRawData[vsicL2] || `Ngành công nghiệp cấp 2 (${vsicL2})`;
                    
                    if (entItem) {
                      classification = "Doanh nghiệp";
                      const groupKey = `${xaVal}-${vsicL2}`;
                      const grp = corporateSamplingResults.groups[groupKey];
                      if (grp) {
                        const isSelected = grp.selectedCandidates.find(c => c.id === idVal);
                        const isBackup = grp.backupCandidates.find(c => c.id === idVal);
                        if (isSelected) {
                          samplingStatus = "Mẫu chính thức";
                          samplingDetail = isSelected.selectionType;
                          cumulativePercent = isSelected.cumulativeRevenuePercent;
                        } else if (isBackup) {
                          samplingStatus = "Mẫu dự phòng";
                          samplingDetail = "Dự phòng";
                          cumulativePercent = isBackup.cumulativeRevenuePercent;
                        }
                      }
                    } else if (indItem) {
                      classification = "Hộ cá thể";
                      const groupKey = `${xaVal}-${vsicL2}`;
                      const grp = individualSamplingResults.groups[groupKey];
                      if (grp) {
                        const isSelected = grp.selectedCandidates.find(c => c.id === idVal);
                        const isBackup = grp.backupCandidates.find(c => c.id === idVal);
                        if (isSelected) {
                          samplingStatus = "Mẫu chính thức";
                          samplingDetail = "Mẫu chính thức (Chuẩn GSO)";
                        } else if (isBackup) {
                          samplingStatus = "Mẫu dự phòng";
                          samplingDetail = "Mẫu dự phòng";
                        }
                      }
                    }
                    
                    return [
                      idVal,
                      nameVal,
                      xaVal,
                      manganhVal,
                      vsicL2,
                      vsicL2Name,
                      revVal,
                      classification,
                      samplingStatus,
                      samplingDetail,
                      cumulativePercent.toFixed(1) + "%"
                    ];
                  })
                  .filter((r): r is NonNullable<typeof r> => r !== null);
                
                const headers_escaped = headers.map(h => `"${h}"`);
                let csvContent = "\uFEFF"; // BOM for UTF-8
                csvContent += headers_escaped.join(",") + "\n";
                rows.forEach(r => {
                  const escaped = r.map(v => {
                    const s = String(v).replace(/"/g, '""');
                    return `"${s}"`;
                  });
                  csvContent += escaped.join(",") + "\n";
                });
                
                const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
                const url = URL.createObjectURL(blob);
                const link = document.createElement("a");
                link.setAttribute("href", url);
                link.setAttribute("download", `Ket_Qua_Chon_Mau_Khao_Sat_${new Date().toISOString().slice(0, 10)}.csv`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
              } catch (err: any) {
                alert("Lỗi khi xuất file chọn mẫu: " + err.message);
              }
            }}
            className="bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 hover:scale-[1.02] active:scale-95 text-white font-bold px-5 py-2.5 rounded-xl shadow-md flex items-center justify-center gap-2 transition-all cursor-pointer font-sans uppercase tracking-wider text-xs"
          >
            <Download className="w-4 h-4 text-white" /> XUẤT KẾT QUẢ CHỌN MẪU KHẢO SÁT (.CSV)
          </button>
        )}
      </div>

      {/* KHỐI NẠP FILE TRỰC TIẾP CHO PHÂN HỆ CHỌN MẪU */}
      <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h4 className="text-sm font-black text-slate-800 tracking-tight">📂 QUẢN LÝ CÁC TỆP CHỌN MẪU RIÊNG</h4>
          <p className="text-[11px] text-slate-500 mt-0.5">Nạp hoặc gỡ bỏ các tệp dữ liệu dùng riêng cho phân hệ chọn mẫu (Doanh nghiệp &amp; Hộ cá thể). Danh mục chuẩn VSIC chuẩn vẫn giữ nguyên vẹn.</p>
        </div>
        {(sampCorpData.length > 0 || sampIndData.length > 0) && (
          <button
            onClick={() => {
              if (window.confirm("Bạn có chắc chắn muốn xóa toàn bộ danh sách tệp dữ liệu chọn mẫu (Doanh nghiệp & Hộ cá thể) đã nạp? (Hệ thống sẽ giữ lại nguyên vẹn các danh mục hệ thống và thiết lập khác)")) {
                setSampCorpData([]);
                setSampCorpFileName("");
                setSampIndData([]);
                setSampIndFileName("");
                alert("Đã xóa các file dữ liệu chọn mẫu thành công! (Danh mục chuẩn VSIC được giữ nguyên)");
              }
            }}
            className="bg-red-50 hover:bg-red-100 border border-red-200 text-red-600 hover:text-red-700 font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer self-start sm:self-center"
            title="Xóa toàn bộ file dữ liệu đã nạp cho phân hệ chọn mẫu"
          >
            <Trash2 className="w-4 h-4" /> XÓA DỮ LIỆU CHỌN MẪU
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
        {/* Khối nạp Doanh nghiệp */}
        <div className="bg-slate-50 border border-slate-200 p-5 rounded-2xl space-y-4 shadow-inner flex flex-col justify-between">
          <div className="space-y-1">
            <span className="text-xs font-black text-indigo-700 tracking-wider uppercase font-mono flex items-center gap-1.5">
              🏢 KHỐI 1: DANH SÁCH DOANH NGHIỆP CƠ SỞ
            </span>
            <p className="text-[11px] text-slate-500 font-medium font-sans leading-normal">
              Nạp danh sách các doanh nghiệp trên địa bàn để thực hiện chọn mẫu theo quy mô doanh thu lũy kế.
            </p>
          </div>
          
          <div className="space-y-3 pt-2">
            <label className="bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-3 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer w-full border-dashed border-2 hover:border-indigo-500">
              <FileUp className="w-4 h-4 text-indigo-500" /> 
              {sampCorpFileName ? "THAY ĐỔI FILE DOANH NGHIỆP" : "CHỌN FILE DOANH NGHIỆP (.xlsx, .xls, .csv)"}
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv" 
                onChange={(e) => handleSamplingFileUpload(e, "corp")} 
                className="hidden" 
              />
            </label>

            {sampCorpFileName && (
              <div className="bg-white border border-slate-200 p-3 rounded-xl text-xs flex justify-between items-center text-slate-700 shadow-sm">
                <span className="truncate max-w-[220px] font-bold text-indigo-700" title={sampCorpFileName}>📄 {sampCorpFileName}</span>
                <div className="flex items-center gap-2">
                  <span className="font-mono text-xs text-indigo-600 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded-lg font-bold">{sampCorpData.length.toLocaleString()} dòng</span>
                  <button
                    onClick={() => {
                      setSampCorpData([]);
                      setSampCorpFileName("");
                    }}
                    className="text-red-500 hover:text-red-700 font-bold text-sm bg-transparent border-0 cursor-pointer p-1"
                    title="Xóa file này"
                  >
                    ×
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Khối nạp Hộ cá thể */}
        <div className="bg-slate-50 border border-slate-200 p-5 rounded-2xl space-y-4 shadow-inner flex flex-col justify-between">
          <div className="space-y-1">
            <span className="text-xs font-black text-orange-700 tracking-wider uppercase font-mono flex items-center gap-1.5">
              🏪 KHỐI 2: DANH SÁCH HỘ KINH DOANH CÁ THỂ
            </span>
            <p className="text-[11px] text-slate-500 font-medium font-sans leading-normal">
              Nạp danh sách các hộ cá thể công nghiệp để chọn mẫu theo định mức GSO của Tổng cục Thống kê.
            </p>
          </div>
          
          <div className="space-y-3 pt-2">
            <label className="bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-3 rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer w-full border-dashed border-2 hover:border-orange-500">
              <FileUp className="w-4 h-4 text-orange-500" /> 
              {sampIndFileName ? "THAY ĐỔI FILE HỘ CÁ THỂ" : "CHỌN FILE HỘ CÁ THỂ (.xlsx, .xls, .csv)"}
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv" 
                onChange={(e) => handleSamplingFileUpload(e, "ind")} 
                className="hidden" 
              />
            </label>

            {sampIndFileName && (
              <div className="bg-white border border-slate-200 p-3 rounded-xl text-xs flex justify-between items-center text-slate-700 shadow-sm">
                <span className="truncate max-w-[220px] font-bold text-orange-700" title={sampIndFileName}>📄 {sampIndFileName}</span>
                <div className="flex items-center gap-2">
                  <span className="font-mono text-xs text-orange-600 bg-orange-50 border border-orange-100 px-2 py-0.5 rounded-lg font-bold">{sampIndData.length.toLocaleString()} dòng</span>
                  <button
                    onClick={() => {
                      setSampIndData([]);
                      setSampIndFileName("");
                    }}
                    className="text-red-500 hover:text-red-700 font-bold text-sm bg-transparent border-0 cursor-pointer p-1"
                    title="Xóa file này"
                  >
                    ×
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {(mainData.length > 0 || sampCorpData.length > 0 || sampIndData.length > 0) ? (
        <>
          {/* KHỐI ĐIỀU KHIỂN & CHẠY CHỌN MẪU HỆ THỐNG */}
          <div className="bg-white border-2 border-indigo-500 rounded-2xl p-6 shadow-md space-y-4 relative overflow-hidden">
            <div className="absolute top-0 right-0 bg-indigo-500 text-white font-mono text-[9px] font-black px-3 py-1 uppercase rounded-bl-xl tracking-widest animate-pulse">
              ⚡ LIVE CALCULATION ACTIVE
            </div>
            
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
              <div className="space-y-1.5 max-w-2xl">
                <h3 className="text-base font-black text-indigo-900 tracking-tight flex items-center gap-2 uppercase font-sans">
                  <Play className="w-5 h-5 text-indigo-600 fill-indigo-600 animate-pulse" /> KHỐI ĐIỀU KHIỂN &amp; CHẠY CHỌN MẪU HỆ THỐNG
                </h3>
                <p className="text-xs text-slate-600 leading-relaxed font-sans">
                  Mặc định hệ thống luôn tự động tính toán tức thì (Real-time) mỗi khi bạn nạp file hoặc thay đổi bộ lọc. Tuy nhiên, bạn có thể <b>bấm nút kích hoạt bên phải</b> để hệ thống tái cơ cấu, phân phối lại tỷ lệ và đồng bộ kết xuất danh sách tối ưu hóa mẫu mới nhất.
                </p>
              </div>

              <div className="flex flex-col sm:flex-row items-center gap-3 shrink-0">
                <button
                  id="btn-run-sampling"
                  onClick={handleManualRun}
                  disabled={isCalculating}
                  className={`w-full sm:w-auto px-6 py-3.5 rounded-xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-2.5 transition-all shadow-md active:scale-95 cursor-pointer border ${
                    isCalculating
                      ? "bg-indigo-50 border-indigo-200 text-indigo-400 cursor-not-allowed"
                      : "bg-indigo-600 hover:bg-indigo-700 hover:shadow-lg border-indigo-700 text-white"
                  }`}
                >
                  {isCalculating ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin text-indigo-500" />
                      <span>ĐANG TÍNH TOÁN MẪU...</span>
                    </>
                  ) : (
                    <>
                      <Play className="w-4 h-4 text-white fill-white" />
                      <span>KÍCH HOẠT CHẠY CHỌN MẪU</span>
                    </>
                  )}
                </button>
              </div>
            </div>

            {/* Trạng thái thông báo thành công hoặc đang chạy */}
            {isCalculating && (
              <div className="bg-indigo-50/50 border border-indigo-100 rounded-xl p-4 flex items-center gap-3 animate-pulse">
                <div className="h-2 w-2 rounded-full bg-indigo-600 animate-ping"></div>
                <span className="text-xs text-indigo-800 font-medium font-sans">Hệ thống đang chạy thuật toán phân tầng đối xứng, lọc dải mã ngành cấp 2, xếp hạng quy mô doanh thu tích lũy và áp định mức của Tổng cục Thống kê...</span>
              </div>
            )}

            {calculationSuccess && calculationDetails && (
              <div className="bg-emerald-50 border-2 border-emerald-200 rounded-xl p-4 space-y-2">
                <div className="flex items-center gap-2 text-emerald-800 font-bold text-xs font-sans uppercase">
                  <span className="bg-emerald-500 text-white p-1 rounded-full text-xs">
                    <Check className="w-4 h-4 stroke-[3]" />
                  </span>
                  <span>Đã chạy và tối ưu hóa mẫu thành công lúc {calculationDetails.time}!</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs font-sans text-slate-700 pt-1 border-t border-emerald-100">
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">⚡ Chế độ:</span>
                    <span className="font-extrabold text-emerald-700 bg-white px-2 py-0.5 rounded border border-emerald-200 font-mono text-[10px]">Tối ưu phân tầng</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">🎯 Mẫu chính thức:</span>
                    <span className="font-extrabold text-indigo-700 bg-white px-2 py-0.5 rounded border border-indigo-200 font-mono text-[11px]">{calculationDetails.totalSelected.toLocaleString()} đơn vị</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">🛡️ Danh sách dự phòng:</span>
                    <span className="font-extrabold text-orange-700 bg-white px-2 py-0.5 rounded border border-orange-200 font-mono text-[11px]">{calculationDetails.totalBackup.toLocaleString()} đơn vị</span>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Cấu hình Nhóm ngành & Hướng dẫn File Excel */}
          <div className="bg-gradient-to-br from-indigo-50/50 to-slate-50/50 border border-indigo-100 rounded-2xl p-6 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* Left Col: Sector Selector */}
            <div className="lg:col-span-5 space-y-4">
              <div className="flex items-center gap-2 border-b border-indigo-100 pb-2.5">
                <Layers className="w-4 h-4 text-indigo-600" />
                <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-700">Bộ chọn Nhóm ngành Khảo sát</h3>
              </div>
              <p className="text-[11px] text-slate-500 leading-relaxed font-sans">
                Chọn một hoặc nhiều nhóm ngành kinh tế dưới đây để tự động lọc danh sách đối tượng chọn mẫu:
              </p>
              
              <div className="space-y-2">
                {[
                  { id: "congnghiep", label: "Công nghiệp (Mã 05 - 33)", color: "border-indigo-200 text-indigo-700 bg-indigo-50/50" },
                  { id: "thuongmai", label: "Thương mại (Mã 45 - 46)", color: "border-emerald-200 text-emerald-700 bg-emerald-50/50" },
                  { id: "vantai", label: "Vận tải (Mã 49)", color: "border-teal-200 text-teal-700 bg-teal-50/50" },
                  { id: "anuongluutru", label: "Ăn uống & Lưu trú (Mã 55, 56)", color: "border-orange-200 text-orange-700 bg-orange-50/50" }
                ].map((sector) => {
                  const isSelected = selectedSectors.includes(sector.id);
                  return (
                    <button
                      key={sector.id}
                      onClick={() => {
                        if (isSelected) {
                          setSelectedSectors(prev => prev.filter(x => x !== sector.id));
                        } else {
                          setSelectedSectors(prev => [...prev, sector.id]);
                        }
                      }}
                      className={`w-full text-left p-2.5 rounded-xl border flex items-center justify-between transition-all font-sans text-xs ${
                        isSelected 
                          ? `${sector.color} border-2 font-bold shadow-sm` 
                          : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"
                      }`}
                    >
                      <div className="flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          readOnly
                          className="accent-indigo-600 h-3.5 w-3.5"
                        />
                        <span>{sector.label}</span>
                      </div>
                      <span className={`text-[10px] font-mono px-1.5 py-0.5 rounded font-bold ${isSelected ? "bg-white border" : "bg-slate-100"}`}>
                        {isSelected ? "ĐANG BẬT" : "TẮT"}
                      </span>
                    </button>
                  );
                })}

                {/* Custom input */}
                <div className={`p-3 rounded-xl border transition-all ${selectedSectors.includes("custom") ? "bg-indigo-50/30 border-indigo-200" : "bg-white border-slate-200"}`}>
                  <label className="flex items-center gap-2 cursor-pointer mb-2 text-xs text-slate-700 font-semibold">
                    <input
                      type="checkbox"
                      checked={selectedSectors.includes("custom")}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setSelectedSectors(prev => [...prev, "custom"]);
                        } else {
                          setSelectedSectors(prev => prev.filter(x => x !== "custom"));
                        }
                      }}
                      className="accent-indigo-600 h-3.5 w-3.5"
                    />
                    <span>Tùy chọn / Tự nhập dải mã ngành</span>
                  </label>
                  {selectedSectors.includes("custom") && (
                    <div className="space-y-1.5">
                      <input
                        type="text"
                        value={customSectorRange}
                        onChange={(e) => setCustomSectorRange(e.target.value)}
                        placeholder="Ví dụ: 05-33, 45-46, 49, 55-56"
                        className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 outline-none"
                      />
                      <p className="text-[10px] text-slate-500 font-sans leading-normal">
                        Ngăn cách các dải mã bằng dấu phẩy. Chấp nhận dạng dải <code>05-33</code> hoặc mã đơn lẻ <code>49</code>.
                      </p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Right Col: Excel Prep Guide */}
            <div className="lg:col-span-7 space-y-4">
              <div className="flex items-center gap-2 border-b border-indigo-100 pb-2.5">
                <FileCheck className="w-4 h-4 text-emerald-600" />
                <h3 className="text-xs font-bold uppercase tracking-wider text-emerald-700">Hướng dẫn Cấu trúc File Excel &amp; Tiêu đề Cột</h3>
              </div>
              <p className="text-[11px] text-slate-500 leading-relaxed font-sans">
                Để hệ thống tự động nhận dạng tính năng <b>Ưu tiên mẫu Trung ương</b> và <b>Doanh nghiệp Nhà nước (DNNN) 100%</b>, hãy thiết lập file Excel của bạn khớp với các chuẩn cột dưới đây:
              </p>

              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-inner bg-white">
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-50 text-[10px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-200 font-sans">
                      <th className="p-2.5">Tính năng ưu tiên</th>
                      <th className="p-2.5">Tiêu đề cột hỗ trợ</th>
                      <th className="p-2.5">Giá trị cần ghi trong ô</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-sans text-slate-700 text-[11px]">
                    <tr className="hover:bg-slate-50/50">
                      <td className="p-2.5 font-bold text-emerald-700">
                        ⭐ Ưu tiên mẫu Trung ương
                      </td>
                      <td className="p-2.5">
                        <code className="bg-slate-100 text-slate-800 px-1.5 py-0.5 rounded font-mono font-bold text-[10px]">Mẫu trung ương</code><br/>
                        <span className="text-[9px] text-slate-400">hoặc <code>Mẫu TU</code>, <code>Mau trung uong</code>, <code>Mau TU</code></span>
                      </td>
                      <td className="p-2.5 font-semibold text-slate-800">
                        <code className="bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-100">Có</code> <span className="text-slate-400 font-normal">hoặc</span> <code className="bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-100">Yes</code>
                      </td>
                    </tr>
                    <tr className="hover:bg-slate-50/50">
                      <td className="p-2.5 font-bold text-indigo-700">
                        🏛️ Ưu tiên 100% DNNN
                      </td>
                      <td className="p-2.5">
                        <code className="bg-slate-100 text-slate-800 px-1.5 py-0.5 rounded font-mono font-bold text-[10px]">Loại hình</code><br/>
                        <span className="text-[9px] text-slate-400">hoặc <code>Loại doanh nghiệp</code>, <code>Hình thức</code>, <code>DNNN</code></span>
                      </td>
                      <td className="p-2.5 font-semibold text-slate-800">
                        <code className="bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded border border-indigo-100">Nhà nước</code> <span className="text-slate-400 font-normal">hoặc</span> <code className="bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded border border-indigo-100">Có</code><br/>
                        <span className="text-[9px] text-slate-400 font-normal">Chấp nhận giá trị chứa từ khóa "Nhà nước"</span>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="bg-indigo-50 text-indigo-800 border border-indigo-100 p-2.5 rounded-xl text-[10px] leading-relaxed flex items-start gap-1.5 font-sans">
                <span className="text-xs">💡</span>
                <span><b>Mẹo thiết lập:</b> Hệ thống sử dụng bộ khớp thông minh không phân biệt chữ hoa/thường hay dấu tiếng Việt. Bạn chỉ cần điền đúng từ khóa hoặc sử dụng cột <b>DNNN = Có</b> và <b>Mẫu TU = Có</b> trong file Excel là hệ thống tự nhận dạng tức thì!</span>
              </div>
            </div>
          </div>

          {/* Setup parameters / mapping */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            
            {/* Column mappings */}
            <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm hover:scale-[1.01] transition-all">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2.5">
                <Sliders className="w-4 h-4 text-indigo-600" />
                <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-600">1. Cấu hình cột khảo sát</h3>
              </div>
              
              <div className="space-y-3 text-xs">
                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Cột Khóa định danh (MST/ID)</label>
                  <select
                    value={sampIdCol}
                    onChange={(e) => setSampIdCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột định danh --</option>
                    {samplingColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Cột Địa bàn xã / Địa bàn</label>
                  <select
                    value={sampXaCol}
                    onChange={(e) => setSampXaCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột địa bàn --</option>
                    {samplingColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Cột Mã ngành VSIC</label>
                  <select
                    value={sampManganhCol}
                    onChange={(e) => setSampManganhCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột mã ngành --</option>
                    {samplingColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Cột Doanh thu / Sản lượng</label>
                  <select
                    value={sampDoanhThuCol}
                    onChange={(e) => setSampDoanhThuCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                  >
                    <option value="">-- Chọn cột doanh thu --</option>
                    {samplingColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div className="bg-amber-50 border border-amber-200 p-2.5 rounded-lg text-[10px] text-amber-800 leading-relaxed mt-2 shadow-sm">
                  <span className="font-bold">⚠️ Lưu ý Nhóm ngành lọc:</span> Hệ thống tự động tách lấy 2 chữ số đầu của mã ngành và lọc giữ lại theo <b>Bộ chọn Nhóm ngành Khảo sát</b> đang kích hoạt ở trên. Các dòng không thuộc nhóm ngành đã chọn sẽ tự động được loại bỏ khỏi danh sách mẫu.
                </div>

                <div className="border-t border-slate-200 pt-3 space-y-2">
                  <label className="block text-slate-700 font-bold text-xs">Phân loại đối tượng mẫu</label>
                  <div className="flex flex-col gap-2 text-slate-700">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        checked={sampFilterType === "all_ent"}
                        onChange={() => setSampFilterType("all_ent")}
                        className="accent-indigo-600 h-3.5 w-3.5"
                      />
                      <span>Xem toàn bộ là Doanh nghiệp (DN)</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        checked={sampFilterType === "all_ind"}
                        onChange={() => setSampFilterType("all_ind")}
                        className="accent-indigo-600 h-3.5 w-3.5"
                      />
                      <span>Xem toàn bộ là Hộ cá thể (Hộ)</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        checked={sampFilterType === "by_col"}
                        onChange={() => setSampFilterType("by_col")}
                        className="accent-indigo-600 h-3.5 w-3.5"
                      />
                      <span>Phân chia theo cột dữ liệu</span>
                    </label>
                  </div>
                </div>

                {sampFilterType === "by_col" && (
                  <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-200 space-y-2 shadow-inner">
                    <div>
                      <label className="block text-[11px] text-slate-700 font-bold mb-1">Cột phân loại</label>
                      <select
                        value={sampTypeCol}
                        onChange={(e) => setSampTypeCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800 focus:outline-none"
                      >
                        <option value="">-- Chọn cột phân loại --</option>
                        {samplingColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div className="grid grid-cols-2 gap-2 text-slate-700 font-medium">
                      <div>
                        <label className="block text-[10px] text-slate-500 font-bold">Giá trị DN</label>
                        <input
                          type="text"
                          value={sampTypeEnterpriseValue}
                          onChange={(e) => setSampTypeEnterpriseValue(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 outline-none"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] text-slate-500 font-bold">Giá trị Hộ</label>
                        <input
                          type="text"
                          value={sampTypeHouseholdValue}
                          onChange={(e) => setSampTypeHouseholdValue(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 outline-none"
                        />
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* Enterprise selection rules */}
            <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm hover:scale-[1.01] transition-all">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2.5">
                <Activity className="w-4 h-4 text-emerald-600" />
                <h3 className="text-xs font-bold uppercase tracking-wider text-emerald-600">2. Chọn mẫu doanh nghiệp</h3>
              </div>
              
              <div className="space-y-4 text-xs">
                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Phạm vi gom nhóm doanh thu lũy kế</label>
                  <select
                    value={entGroupScope}
                    onChange={(e) => setEntGroupScope(e.target.value as "province" | "xa")}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-emerald-500 focus:border-emerald-500 outline-none shadow-sm font-sans"
                  >
                    <option value="province">Toàn tỉnh (Gom theo Ngành cấp 2 của tỉnh - Đề xuất ✨)</option>
                    <option value="xa">Địa bàn xã (Gom theo Địa bàn xã + Ngành cấp 2)</option>
                  </select>
                  <p className="text-[10px] text-slate-500 mt-1 leading-normal">
                    {entGroupScope === "province" 
                      ? "✨ Khuyên dùng: Sắp xếp và chọn mẫu doanh thu 75% trên phạm vi toàn tỉnh của từng ngành. Tránh tình trạng chọn quá nhiều do phân mảnh địa bàn nhỏ lẻ." 
                      : "Sắp xếp và chọn mẫu doanh thu 75% riêng biệt cho từng địa bàn xã. Thích hợp khi cần đại diện chi tiết từng địa bàn xã."
                    }
                  </p>
                </div>

                <div>
                  <div className="flex justify-between mb-1">
                    <span className="text-slate-700 font-bold text-xs">Ngưỡng đóng góp doanh thu lũy kế (%)</span>
                    <span className="text-emerald-600 font-mono font-bold">{entCutoffPercent}%</span>
                  </div>
                  <input
                    type="range"
                    min="50"
                    max="95"
                    step="5"
                    value={entCutoffPercent}
                    onChange={(e) => setEntCutoffPercent(parseInt(e.target.value))}
                    className="w-full accent-emerald-500 cursor-pointer"
                  />
                  <p className="text-[10px] text-slate-500 mt-1 leading-normal">
                    Xếp doanh nghiệp từ lớn tới nhỏ. Cộng dồn doanh thu cho đến khi chiếm tối thiểu {entCutoffPercent}% tổng doanh thu của ngành tại địa bàn.
                  </p>
                </div>

                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Số Doanh nghiệp tối thiểu trong nhóm</label>
                  <input
                    type="number"
                    min="1"
                    max="20"
                    value={entMinGroupSize}
                    onChange={(e) => setEntMinGroupSize(parseInt(e.target.value) || 1)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                  />
                  <p className="text-[10px] text-slate-500 mt-1">
                    Nếu số doanh nghiệp tại địa bàn cho 1 ngành từ {entMinGroupSize} trở xuống: Chọn toàn bộ 100% không loại trừ.
                  </p>
                </div>

                <div className="space-y-2 pt-2 border-t border-slate-200">
                  <label className="flex items-center gap-2 cursor-pointer text-slate-700">
                    <input
                      type="checkbox"
                      checked={entForceStates}
                      onChange={(e) => setEntForceStates(e.target.checked)}
                      className="accent-emerald-500 h-3.5 w-3.5 rounded"
                    />
                    <span>Ưu tiên chọn 100% Doanh nghiệp Nhà nước (DNNN)</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-slate-700">
                    <input
                      type="checkbox"
                      checked={entForceMonthly}
                      onChange={(e) => setEntForceMonthly(e.target.checked)}
                      className="accent-emerald-500 h-3.5 w-3.5 rounded"
                    />
                    <span>Ưu tiên mẫu Trung ương ("Mẫu trung ương" có ghi "Có")</span>
                  </label>
                </div>
              </div>
            </div>

            {/* Household selection rules */}
            <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm hover:scale-[1.01] transition-all">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2.5">
                <Sliders className="w-4 h-4 text-orange-600" />
                <h3 className="text-xs font-bold uppercase tracking-wider text-orange-600">3. Chọn mẫu Hộ cá thể</h3>
              </div>
              
              <div className="space-y-4 text-xs">
                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1.5">Cách thức chọn mẫu hộ</label>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      onClick={() => setIndSamplingMode("GSO")}
                      className={`py-2 px-3 rounded-lg text-xs font-bold transition-all border ${
                        indSamplingMode === "GSO" 
                          ? "bg-orange-50 text-orange-700 border-orange-300 shadow-sm" 
                          : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50 hover:text-slate-950 shadow-sm"
                      }`}
                    >
                      Theo Chuẩn TCTK
                    </button>
                    <button
                      onClick={() => setIndSamplingMode("Custom")}
                      className={`py-2 px-3 rounded-lg text-xs font-bold transition-all border ${
                        indSamplingMode === "Custom" 
                          ? "bg-orange-50 text-orange-700 border-orange-300 shadow-sm" 
                          : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50 hover:text-slate-950 shadow-sm"
                      }`}
                    >
                      Tùy chọn thiết lập
                    </button>
                  </div>
                  <p className="text-[10px] text-slate-500 mt-1.5 leading-normal">
                    {indSamplingMode === "GSO" 
                      ? "Áp dụng định mức chuẩn: Dưới 5 hộ chọn hết; từ 6-100 hộ chọn 5 hộ lớn nhất; 101-1000 chọn 8 hộ lớn nhất; trên 1000 chọn 1%."
                      : "Cho phép cấu hình thủ công số lượng hộ hoặc tỷ lệ % hộ đại diện lớn nhất."}
                  </p>
                </div>

                {indSamplingMode === "Custom" && (
                  <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-3 shadow-inner">
                    <div className="flex gap-4 text-slate-700 font-semibold">
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input
                          type="radio"
                          checked={indCustomMode === "fixed"}
                          onChange={() => setIndCustomMode("fixed")}
                          className="accent-orange-500 h-3.5 w-3.5"
                          name="indCustomMode"
                        />
                        <span>Cố định hộ</span>
                      </label>
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input
                          type="radio"
                          checked={indCustomMode === "percent"}
                          onChange={() => setIndCustomMode("percent")}
                          className="accent-orange-500 h-3.5 w-3.5"
                          name="indCustomMode"
                        />
                        <span>Tỷ lệ %</span>
                      </label>
                    </div>

                    {indCustomMode === "fixed" ? (
                      <div>
                        <label className="block text-[10px] text-slate-500 font-bold mb-1">Số lượng hộ lấy tối đa trong nhóm</label>
                        <input
                          type="number"
                          min="1"
                          max="100"
                          value={indCustomCountValue}
                          onChange={(e) => setIndCustomCountValue(parseInt(e.target.value) || 1)}
                          className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 outline-none"
                        />
                      </div>
                    ) : (
                      <div>
                        <label className="block text-[10px] text-slate-500 font-bold mb-1">Tỷ lệ % lấy mẫu hộ</label>
                        <input
                          type="number"
                          min="1"
                          max="100"
                          value={indCustomPercentValue}
                          onChange={(e) => setIndCustomPercentValue(parseInt(e.target.value) || 1)}
                          className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 outline-none"
                        />
                      </div>
                    )}
                  </div>
                )}

                <div>
                  <label className="block text-slate-700 font-bold text-xs mb-1">Giới hạn mẫu tối đa / ngành địa bàn</label>
                  <input
                    type="number"
                    min="1"
                    max="200"
                    value={indMaxCap}
                    onChange={(e) => setIndMaxCap(parseInt(e.target.value) || 10)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none"
                  />
                  <p className="text-[10px] text-slate-500 mt-1">
                    Ngưỡng an toàn chặn trên để tránh bùng nổ số lượng mẫu khảo sát quá tải tại các địa bàn lớn đặc thù.
                  </p>
                </div>
              </div>
            </div>

          </div>

          {/* Summary widgets */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-1 shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all">
              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">DOANH NGHIỆP NGUỒN</div>
              <div className="text-2xl font-black text-indigo-600 font-mono">{enterpriseList.length}</div>
            </div>
            <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-1 shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all">
              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">HỘ CÁ THỂ NGUỒN</div>
              <div className="text-2xl font-black text-amber-600 font-mono">{individualList.length}</div>
            </div>
            <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-1 shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all">
              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">ĐƯỢC CHỌN CHÍNH THỨC</div>
              <div className="text-2xl font-black text-emerald-600 font-mono">
                {corporateSamplingResults.selectedIDs.size + individualSamplingResults.selectedIDs.size}
              </div>
            </div>
            <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-1 shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all">
              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">DỰ PHÒNG XẾP HẠNG</div>
              <div className="text-2xl font-black text-orange-500 font-mono">
                {corporateSamplingResults.backupIDs.size + individualSamplingResults.backupIDs.size}
              </div>
            </div>
            <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-1 shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all col-span-2 md:col-span-1">
              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">ĐIỀU TRA CHI TIẾT NGÀNH XÃ</div>
              <div className="text-2xl font-black text-teal-600 font-mono">{allSamplingGroups.length}</div>
            </div>
          </div>

          {/* Group exploration table & detail */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            
            {/* Groups list */}
            <div className="lg:col-span-5 bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm">
              <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-3">
                <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider font-sans">Danh sách nhóm địa bàn &amp; ngành</h4>
                <span className="text-[10px] font-mono bg-indigo-50 text-indigo-600 px-2.5 py-0.5 rounded-full border border-indigo-200 font-bold">
                  {filteredSamplingGroups.length} nhóm
                </span>
              </div>

              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                <input
                  type="text"
                  placeholder="Lọc mã xã hoặc mã ngành VSIC..."
                  value={sampSearchTerm}
                  onChange={(e) => setSampSearchTerm(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-lg pl-9 pr-4 py-2 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 shadow-inner"
                />
              </div>

              <div className="overflow-y-auto max-h-[450px] custom-scrollbar space-y-1.5 pr-1 text-xs">
                {filteredSamplingGroups.length > 0 ? (
                  filteredSamplingGroups.map(g => {
                    const isActive = sampActiveDetailGroup === g.key;
                    return (
                      <button
                        key={g.key}
                        onClick={() => setSampActiveDetailGroup(g.key)}
                        className={`w-full text-left p-3 rounded-xl transition-all border flex items-center justify-between ${
                          isActive 
                            ? "bg-indigo-50 border-indigo-400 text-indigo-900 font-semibold shadow-sm" 
                            : "bg-slate-50/50 border-slate-200 hover:bg-slate-100 hover:border-slate-300 text-slate-700"
                        }`}
                      >
                        <div className="space-y-1 max-w-[70%]">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="bg-white text-slate-600 border border-slate-200 px-1.5 py-0.5 rounded font-mono text-[10px] font-bold shadow-sm">Xã: {g.xaCode}</span>
                            <span className="bg-indigo-100/50 text-indigo-700 px-1.5 py-0.5 rounded font-mono text-[10px] border border-indigo-200 font-bold">VSIC: {g.vsicL2}</span>
                          </div>
                          <div className="text-[10px] text-indigo-600 font-semibold truncate max-w-[240px]" title={g.vsicL2Name}>
                            {g.vsicL2Name}
                          </div>
                          <div className="text-[10px] text-slate-500 flex items-center gap-2">
                            <span>Nguồn: <b className="text-slate-800 font-semibold">{g.totalN}</b></span>
                            <span>•</span>
                            <span>Doanh thu: <b className="text-slate-850 font-semibold">{g.totalRevenue.toLocaleString()}</b></span>
                          </div>
                        </div>
                        <div className="text-right space-y-1 shrink-0">
                          <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 text-[10px] font-extrabold px-2.5 py-0.5 rounded-full block">
                            Mẫu: {g.selectedCount}
                          </span>
                          {g.backupCount > 0 && (
                            <span className="text-[9px] text-slate-500 block font-medium">
                              Dự phòng: {g.backupCount}
                            </span>
                          )}
                        </div>
                      </button>
                    );
                  })
                ) : (
                  <div className="p-8 text-center text-slate-400 border border-dashed border-slate-200 rounded-xl bg-slate-50">
                    Không tìm thấy nhóm địa bàn nào thỏa mãn
                  </div>
                )}
              </div>
            </div>

            {/* Group Details */}
            <div className="lg:col-span-7 bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm">
              {sampActiveDetailGroup ? (
                (() => {
                  const grp = allSamplingGroups.find(g => g.key === sampActiveDetailGroup);
                  if (!grp) return <div className="text-slate-400 text-center p-8">Lỗi tải dữ liệu nhóm</div>;
                  
                  const corpSelected = grp.corpGrp?.selectedCandidates || [];
                  const corpBackup = grp.corpGrp?.backupCandidates || [];
                  const indSelected = grp.indGrp?.selectedCandidates || [];
                  const indBackup = grp.indGrp?.backupCandidates || [];
                  
                  const allSelected = [...corpSelected, ...indSelected];
                  const allBackup = [...corpBackup, ...indBackup];
                  
                  return (
                    <div className="space-y-4 text-xs">
                      
                      {/* Detail header */}
                      <div className="border-b border-slate-100 pb-3 space-y-2">
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-sm font-bold text-slate-800 flex items-center gap-1.5 font-sans uppercase tracking-tight">
                            <Layers className="w-4 h-4 text-indigo-600" /> CHI TIẾT ĐỊA BÀN XÃ {grp.xaCode} - {grp.vsicL2Name} ({grp.vsicL2})
                          </span>
                          <span className="text-[10px] bg-slate-100 text-slate-600 font-mono px-2 py-0.5 rounded font-bold border border-slate-200 shadow-sm">
                            Tổng {grp.totalN} đơn vị
                          </span>
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-[11px] bg-slate-50 p-2.5 rounded-xl border border-slate-200 shadow-inner">
                          <div>
                            <span className="text-slate-500 font-semibold">Tổng doanh thu nhóm:</span>
                            <span className="block text-slate-900 font-bold font-mono mt-0.5">{grp.totalRevenue.toLocaleString()}</span>
                          </div>
                          <div>
                            <span className="text-slate-500 font-semibold">Doanh nghiệp chọn mẫu:</span>
                            <span className="block text-emerald-600 font-bold font-mono mt-0.5">{corpSelected.length}/{grp.corpGrp?.totalN || 0}</span>
                          </div>
                          <div>
                            <span className="text-slate-500 font-semibold">Hộ cá thể chọn mẫu:</span>
                            <span className="block text-amber-600 font-bold font-mono mt-0.5">{indSelected.length}/{grp.indGrp?.totalN || 0}</span>
                          </div>
                        </div>
                      </div>

                      {/* Members list */}
                      <div className="space-y-3">
                        <div className="text-[11px] font-bold text-slate-700 uppercase tracking-wider font-sans">Đơn vị đã chọn làm mẫu chính thức ({allSelected.length})</div>
                        <div className="space-y-1.5 max-h-[160px] overflow-y-auto custom-scrollbar pr-1">
                          {allSelected.length > 0 ? (
                            allSelected.map(item => (
                              <div key={item.id} className="bg-emerald-50/50 border border-emerald-200 rounded-xl p-2.5 flex items-center justify-between gap-2 hover:scale-[1.01] transition-all">
                                <div className="space-y-1 max-w-[75%]">
                                  <div className="font-bold text-slate-900 truncate" title={item.name}>{item.name}</div>
                                  <div className="text-[10px] text-slate-500 font-mono flex flex-wrap items-center gap-x-2 gap-y-0.5">
                                    <span>MST/ID: {item.id}</span>
                                    <span>•</span>
                                    <span>VSIC: {item.vsicFull || grp.vsicL2}</span>
                                    {item.originalRow?.["Loại hình"] && (
                                      <>
                                        <span>•</span>
                                        <span className="text-indigo-600 font-medium">{item.originalRow["Loại hình"]}</span>
                                      </>
                                    )}
                                  </div>
                                </div>
                                <div className="text-right shrink-0">
                                  <div className="font-extrabold text-emerald-700 font-mono">{item.revenue.toLocaleString()}</div>
                                  <div className="text-[9px] text-emerald-600 font-bold uppercase tracking-wider mt-0.5">{item.selectionType}</div>
                                </div>
                              </div>
                            ))
                          ) : (
                            <div className="p-4 text-center text-slate-400 border border-dashed border-slate-200 rounded-xl bg-slate-50">Không có mẫu nào được chọn</div>
                          )}
                        </div>
                      </div>

                      {/* Backup list */}
                      <div className="space-y-3 pt-2">
                        <div className="text-[11px] font-bold text-slate-700 uppercase tracking-wider font-sans">Đơn vị xếp vào danh sách dự phòng ({allBackup.length})</div>
                        <div className="space-y-1.5 max-h-[160px] overflow-y-auto custom-scrollbar pr-1">
                          {allBackup.length > 0 ? (
                            allBackup.map((item, idx) => (
                              <div key={item.id} className="bg-amber-50/30 border border-amber-200 rounded-xl p-2.5 flex items-center justify-between gap-2 hover:scale-[1.01] transition-all">
                                <div className="space-y-1 max-w-[75%]">
                                  <div className="font-bold text-slate-800 truncate" title={item.name}>{item.name}</div>
                                  <div className="text-[10px] text-slate-500 font-mono">
                                    MST/ID: {item.id} • Thứ tự xếp hạng dự phòng: #{idx + 1}
                                  </div>
                                </div>
                                <div className="text-right shrink-0">
                                  <div className="font-extrabold text-amber-700 font-mono">{item.revenue.toLocaleString()}</div>
                                  <div className="text-[9px] text-amber-600/75 font-bold mt-0.5">Xếp hạng dự bị #{idx + 1}</div>
                                </div>
                              </div>
                            ))
                          ) : (
                            <div className="p-4 text-center text-slate-400 border border-dashed border-slate-200 rounded-xl bg-slate-50">Không có đơn vị dự phòng</div>
                          )}
                        </div>
                      </div>

                    </div>
                  );
                })()
              ) : (
                <div className="h-full flex flex-col items-center justify-center p-12 text-center space-y-3 bg-slate-50/50 border border-dashed border-slate-200 rounded-2xl">
                  <div className="bg-indigo-50 p-4 rounded-full text-indigo-600 border border-indigo-100 shadow-sm">
                    <Layers className="w-8 h-8" />
                  </div>
                  <div>
                    <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider font-sans">Xem chi tiết đối tượng lọt mẫu</h4>
                    <p className="text-[11px] text-slate-500 max-w-sm mt-1 leading-relaxed">
                      Hãy chọn một nhóm địa bàn &amp; mã ngành bên trái để phân tích danh sách chi tiết, theo dõi thứ tự đóng góp doanh thu và thứ hạng dự phòng.
                    </p>
                  </div>
                </div>
              )}
            </div>

          </div>
        </>
      ) : (
        <div className="bg-slate-50 rounded-2xl p-12 text-center text-xs text-amber-700 border border-slate-200 font-condensed space-y-4 shadow-sm">
          <div className="flex justify-center">
            <AlertTriangle className="w-8 h-8 text-amber-500 animate-bounce" />
          </div>
          <div>
            <h3 className="text-sm font-bold uppercase tracking-wider text-slate-800">CHƯA CÓ DỮ LIỆU NGUỒN KHẢO SÁT</h3>
            <p className="text-xs text-slate-500 max-w-md mx-auto mt-1 font-sans">
              Hãy nạp file Excel hoặc CSV danh sách đơn vị cơ sở khảo sát (doanh nghiệp/hộ cá thể) của bạn tại Trang Chủ trước để hệ thống tiến hành tính toán phân tích và thiết lập mẫu.
            </p>
          </div>
        </div>
      )}
    </div>
  );
}
