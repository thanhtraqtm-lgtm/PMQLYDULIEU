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
  const [manualIndGroupSamples, setManualIndGroupSamples] = useState<Record<string, number>>({});
  const [sampFilterTab, setSampFilterTab] = useState<"corp" | "ind">("corp");

  // Cấu hình kéo thả nâng cao theo quy mô nhóm ngành của Hộ cá thể (Individual Sliders)
  const [indSize1To5Value, setIndSize1To5Value] = useState<number>(5); // Nhóm có 1-5 hộ: Chọn toàn bộ (mặc định tối đa 5 hoặc toàn bộ)
  const [indSize1To5All, setIndSize1To5All] = useState<boolean>(true); // Mặc định chọn toàn bộ
  const [indSize6To100Value, setIndSize6To100Value] = useState<number>(5); // Nhóm có 6-100: Chọn 5 cơ sở (kéo từ 1 đến 50)
  const [indSize101To1000Value, setIndSize101To1000Value] = useState<number>(8); // Nhóm có 101-1000: Chọn 8 cơ sở (kéo từ 1 đến 100)
  const [indSize1001PlusPercent, setIndSize1001PlusPercent] = useState<number>(1); // Nhóm từ 1001 trở lên: Chọn 1% (kéo % từ 0.1 đến 10.0)
  const [indTransportPercent, setIndTransportPercent] = useState<number>(1.5); // Riêng vận tải kho bãi: Chọn 1.5% (kéo % từ 0.5 đến 15.0)
  const [indTransportMaxCap, setIndTransportMaxCap] = useState<number>(50); // Tối đa không quá 50 cơ sở cho vận tải kho bãi (kéo từ 5 đến 200)

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
  
  // Independent Column configurations to completely separate Enterprise and Household
  const [corpIdCol, setCorpIdCol] = useState<string>("");
  const [corpXaCol, setCorpXaCol] = useState<string>("");
  const [corpManganhCol, setCorpManganhCol] = useState<string>("");
  const [corpDoanhThuCol, setCorpDoanhThuCol] = useState<string>("");

  const [indIdCol, setIndIdCol] = useState<string>("");
  const [indXaCol, setIndXaCol] = useState<string>("");
  const [indManganhCol, setIndManganhCol] = useState<string>("");
  const [indDoanhThuCol, setIndDoanhThuCol] = useState<string>("");

  const [sampTypeCol, setSampTypeCol] = useState<string>("");
  const [sampFilterType, setSampFilterType] = useState<"all_ent" | "all_ind" | "by_col">("all_ent");
  const [sampTypeEnterpriseValue, setSampTypeEnterpriseValue] = useState<string>("DN");
  const [sampTypeHouseholdValue, setSampTypeHouseholdValue] = useState<string>("Hộ");

  // Search & Navigation
  const [sampSearchTerm, setSampSearchTerm] = useState<string>("");
  const [sampActiveDetailGroup, setSampActiveDetailGroup] = useState<string>("");

  const corpColumns = useMemo(() => {
    const colsSet = new Set<string>();
    columns.forEach(c => colsSet.add(c));
    if (sampCorpData.length > 0) {
      Object.keys(sampCorpData[0] || {}).forEach(c => colsSet.add(c));
    }
    return Array.from(colsSet);
  }, [columns, sampCorpData]);

  const indColumns = useMemo(() => {
    const colsSet = new Set<string>();
    columns.forEach(c => colsSet.add(c));
    if (sampIndData.length > 0) {
      Object.keys(sampIndData[0] || {}).forEach(c => colsSet.add(c));
    }
    return Array.from(colsSet);
  }, [columns, sampIndData]);

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
    if (mapping.idCol) {
      if (!corpIdCol) setCorpIdCol(mapping.idCol);
      if (!indIdCol) setIndIdCol(mapping.idCol);
      if (!sampIdCol) setSampIdCol(mapping.idCol);
    }
    if (mapping.xa) {
      if (!corpXaCol) setCorpXaCol(mapping.xa);
      if (!indXaCol) setIndXaCol(mapping.xa);
      if (!sampXaCol) setSampXaCol(mapping.xa);
    }
    if (mapping.manganh) {
      if (!corpManganhCol) setCorpManganhCol(mapping.manganh);
      if (!indManganhCol) setIndManganhCol(mapping.manganh);
      if (!sampManganhCol) setSampManganhCol(mapping.manganh);
    }
    if (mapping.doanhthu) {
      if (!corpDoanhThuCol) setCorpDoanhThuCol(mapping.doanhthu);
      if (!indDoanhThuCol) setIndDoanhThuCol(mapping.doanhthu);
      if (!sampDoanhThuCol) setSampDoanhThuCol(mapping.doanhthu);
    }
  }, [mapping]);

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

  const autoDetectColumns = (cols: string[], type: "corp" | "ind") => {
    cols.forEach(c => {
      const cLow = c.toLowerCase();
      const isId = cLow.includes("mst") || cLow.includes("mã số thuế") || cLow.includes("định danh") || cLow.includes("id");
      const isXa = cLow.includes("xã") || cLow.includes("địa bàn") || cLow.includes("mã xã");
      const isManganh = cLow.includes("ngành") || cLow.includes("mã ngành") || cLow.includes("vsic");
      const isDoanhThu = cLow.includes("doanh thu") || cLow.includes("sản lượng") || cLow.includes("doanhthu");

      if (type === "corp") {
        if (isId) setCorpIdCol(c);
        if (isXa) setCorpXaCol(c);
        if (isManganh) setCorpManganhCol(c);
        if (isDoanhThu) setCorpDoanhThuCol(c);
      } else {
        if (isId) setIndIdCol(c);
        if (isXa) setIndXaCol(c);
        if (isManganh) setIndManganhCol(c);
        if (isDoanhThu) setIndDoanhThuCol(c);
      }

      // Also set the legacy fallback shared states
      if (isId && !sampIdCol) setSampIdCol(c);
      if (isXa && !sampXaCol) setSampXaCol(c);
      if (isManganh && !sampManganhCol) setSampManganhCol(c);
      if (isDoanhThu && !sampDoanhThuCol) setSampDoanhThuCol(c);
    });
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
            autoDetectColumns(cols, type);
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
            autoDetectColumns(cols, type);
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
      const manganhVal = String(row[corpManganhCol || sampManganhCol] || "").trim();
      const cleanDigits = manganhVal.replace(/\D/g, "");
      const vsicL2 = cleanDigits.slice(0, 2);
      
      if (!isSectorSelected(vsicL2)) return null;
      
      const idVal = String(row[corpIdCol || sampIdCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
      const nameVal = String(row["Tên doanh nghiệp"] || row["Tên đơn vị"] || row["Tên"] || row["Tên hộ"] || "Doanh nghiệp " + index);
      const xaVal = String(row[corpXaCol || sampXaCol] || "30000");
      const revVal = parseFloat(String(row[corpDoanhThuCol || sampDoanhThuCol] || "0").replace(/,/g, "")) || 0;
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
  }, [mainData, sampCorpData, sampFilterType, sampTypeCol, sampTypeEnterpriseValue, corpIdCol, corpXaCol, corpManganhCol, corpDoanhThuCol, sampIdCol, sampXaCol, sampManganhCol, sampDoanhThuCol, selectedSectors, customSectorRange]);

  const individualList = useMemo(() => {
    const sourceData = sampIndData.length > 0 ? sampIndData : mainData;
    if (sourceData.length === 0) return [];

    const processRow = (row: any, index: number) => {
      const manganhVal = String(row[indManganhCol || sampManganhCol] || "").trim();
      const cleanDigits = manganhVal.replace(/\D/g, "");
      const vsicL2 = cleanDigits.slice(0, 2);
      
      if (!isSectorSelected(vsicL2)) return null;
      
      const idVal = String(row[indIdCol || sampIdCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
      const nameVal = String(row["Tên hộ"] || row["Tên đơn vị"] || row["Tên"] || "Hộ " + index);
      const xaVal = String(row[indXaCol || sampXaCol] || "30000");
      const revVal = parseFloat(String(row[indDoanhThuCol || sampDoanhThuCol] || "0").replace(/,/g, "")) || 0;
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
  }, [mainData, sampIndData, sampFilterType, sampTypeCol, sampTypeHouseholdValue, indIdCol, indXaCol, indManganhCol, indDoanhThuCol, sampIdCol, sampXaCol, sampManganhCol, sampDoanhThuCol, selectedSectors, customSectorRange]);

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
        const isTransport = ["49", "50", "51", "52", "53"].includes(vsicL2);
        if (totalN <= 5) {
          targetSampleSize = indSize1To5All ? totalN : Math.min(indSize1To5Value, totalN);
        } else if (totalN <= 100) {
          targetSampleSize = Math.min(indSize6To100Value, totalN);
        } else if (totalN <= 1000) {
          targetSampleSize = Math.min(indSize101To1000Value, totalN);
        } else {
          if (isTransport) {
            targetSampleSize = Math.min(indTransportMaxCap, Math.max(8, Math.round(totalN * (indTransportPercent / 100))));
          } else {
            targetSampleSize = Math.max(8, Math.round(totalN * (indSize1001PlusPercent / 100)));
          }
        }
      } else {
        if (indCustomMode === 'fixed') {
          targetSampleSize = Math.min(indCustomCountValue, totalN);
        } else {
          targetSampleSize = Math.min(indMaxCap, Math.max(1, Math.round(totalN * (indCustomPercentValue / 100))));
        }
      }

      // Ghi đè bằng giá trị kéo/chọn thủ công của người dùng nếu có
      if (manualIndGroupSamples[groupKey] !== undefined) {
        targetSampleSize = Math.min(manualIndGroupSamples[groupKey], totalN);
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
  }, [individualList, indSamplingMode, indCustomMode, indCustomCountValue, indCustomPercentValue, indMaxCap, manualIndGroupSamples, indSize1To5Value, indSize1To5All, indSize6To100Value, indSize101To1000Value, indSize1001PlusPercent, indTransportPercent, indTransportMaxCap]);

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
    let result = allSamplingGroups;
    
    // Lọc theo tab hiển thị tách riêng
    if (sampFilterTab === "corp") {
      result = result.filter(g => g.corpGrp && g.corpGrp.totalN > 0);
    } else if (sampFilterTab === "ind") {
      result = result.filter(g => g.indGrp && g.indGrp.totalN > 0);
    }

    if (!sampSearchTerm) return result;
    const term = sampSearchTerm.toLowerCase();
    return result.filter(g => 
      g.xaCode.toLowerCase().includes(term) || 
      g.vsicL2.toLowerCase().includes(term)
    );
  }, [allSamplingGroups, sampSearchTerm, sampFilterTab]);

  const handleManualRun = () => {
    setIsCalculating(true);
    setCalculationSuccess(false);
    setTimeout(() => {
      setIsCalculating(false);
      setCalculationSuccess(true);
      
      const isCorp = sampFilterTab === "corp";
      const selectedCount = isCorp ? corporateSamplingResults.selectedIDs.size : individualSamplingResults.selectedIDs.size;
      const backupCount = isCorp ? corporateSamplingResults.backupIDs.size : individualSamplingResults.backupIDs.size;
      
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
      <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1">
          <h2 className="text-xl font-bold tracking-tight text-slate-900 flex items-center gap-2 font-sans">
            <FileCheck className="w-5 h-5 text-indigo-600" /> BỘ CHỌN MẪU KHẢO SÁT CHUYÊN BIỆT
          </h2>
          <p className="text-xs text-slate-500 font-medium font-sans">
            {sampFilterTab === "corp" 
              ? `Doanh nghiệp: Tự động lọc các nhóm ngành được chọn. Chọn mẫu theo quy mô doanh thu lũy kế ${entCutoffPercent}% tích lũy tối ưu.`
              : `Hộ cá thể: Tự động gom nhóm theo Địa bàn xã & Mã ngành cấp 2. Chọn mẫu theo quy mô và định mức GSO.`}
          </p>
        </div>
        {(sampFilterTab === "corp" ? sampCorpData.length > 0 : sampIndData.length > 0) && (
          <button
            onClick={() => {
              try {
                const isCorp = sampFilterTab === "corp";
                const sourceData = isCorp ? sampCorpData : sampIndData;
                const idCol = isCorp ? (corpIdCol || sampIdCol) : (indIdCol || sampIdCol);
                const xaCol = isCorp ? (corpXaCol || sampXaCol) : (indXaCol || sampXaCol);
                const manganhCol = isCorp ? (corpManganhCol || sampManganhCol) : (indManganhCol || sampManganhCol);
                const doanhthuCol = isCorp ? (corpDoanhThuCol || sampDoanhThuCol) : (indDoanhThuCol || sampDoanhThuCol);

                if (sourceData.length === 0) {
                  alert(`Chưa có dữ liệu nguồn ${isCorp ? "doanh nghiệp" : "hộ cá thể"} để xuất!`);
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
                
                const listData = isCorp ? enterpriseList : individualList;
                const results = isCorp ? corporateSamplingResults : individualSamplingResults;

                const rows = sourceData
                  .map((row, index) => {
                    const idVal = String(row[idCol] || row["Mã số thuế"] || row["MST"] || row["id"] || index);
                    const nameVal = isCorp 
                      ? String(row["Tên doanh nghiệp"] || row["Tên đơn vị"] || row["Tên"] || "Doanh nghiệp " + index)
                      : String(row["Tên hộ"] || row["Tên đơn vị"] || row["Tên"] || "Hộ " + index);
                    const xaVal = String(row[xaCol] || "30000");
                    const manganhVal = String(row[manganhCol] || "").trim();
                    const cleanDigits = manganhVal.replace(/\D/g, "");
                    const vsicL2 = cleanDigits.slice(0, 2) || "00";
                    const revVal = parseFloat(String(row[doanhthuCol] || "0").replace(/,/g, "")) || 0;
                    
                    const item = listData.find(x => x.id === idVal);
                    if (!item) return null;

                    let classification = isCorp ? "Doanh nghiệp" : "Hộ cá thể";
                    let samplingStatus = "Không được chọn";
                    let samplingDetail = "Không lọt mẫu";
                    let cumulativePercent = 0;
                    const vsicL2Name = vsicRawData[vsicL2] || `Ngành công nghiệp cấp 2 (${vsicL2})`;
                    
                    const groupKey = `${xaVal}-${vsicL2}`;
                    const grp = results.groups[groupKey];
                    if (grp) {
                      const isSelected = grp.selectedCandidates.find(c => c.id === idVal);
                      const isBackup = grp.backupCandidates.find(c => c.id === idVal);
                      if (isSelected) {
                        samplingStatus = "Mẫu chính thức";
                        samplingDetail = isCorp ? isSelected.selectionType : "Mẫu chính thức (Chuẩn GSO)";
                        if (isCorp) cumulativePercent = isSelected.cumulativeRevenuePercent;
                      } else if (isBackup) {
                        samplingStatus = "Mẫu dự phòng";
                        samplingDetail = "Dự phòng";
                        if (isCorp) cumulativePercent = isBackup.cumulativeRevenuePercent;
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
                link.setAttribute("download", `Ket_Qua_Chon_Mau_${isCorp ? "Doanh_Nghiep" : "Ho_Ca_The"}_${new Date().toISOString().slice(0, 10)}.csv`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
              } catch (err: any) {
                alert("Lỗi khi xuất file chọn mẫu: " + err.message);
              }
            }}
            className={`font-bold px-5 py-2.5 rounded-xl shadow-md flex items-center justify-center gap-2 transition-all cursor-pointer font-sans uppercase tracking-wider text-xs text-white ${
              sampFilterTab === "corp" 
                ? "bg-gradient-to-r from-indigo-500 to-indigo-700 hover:from-indigo-600 hover:to-indigo-800" 
                : "bg-gradient-to-r from-amber-500 to-amber-700 hover:from-amber-600 hover:to-amber-800"
            }`}
          >
            <Download className="w-4 h-4 text-white" /> XUẤT KẾT QUẢ MẪU {sampFilterTab === "corp" ? "DOANH NGHIỆP" : "HỘ CÁ THỂ"} (.CSV)
          </button>
        )}
      </div>

      {/* TABS PHÂN HỆ CHUYÊN BIỆT CHÍNH Ở ĐẦU TRANG */}
      <div className="flex bg-slate-100 p-2 rounded-2xl border border-slate-300 shadow-md max-w-2xl mx-auto gap-3">
        <button
          type="button"
          onClick={() => {
            setSampFilterTab("corp");
            setSampActiveDetailGroup("");
          }}
          className={`flex-1 text-sm font-black py-4 px-6 rounded-xl transition-all flex items-center justify-center gap-2.5 cursor-pointer hover:scale-[1.01] ${
            sampFilterTab === "corp"
              ? "bg-indigo-600 text-white shadow-lg border border-indigo-700"
              : "text-slate-600 hover:text-slate-900 hover:bg-white"
          }`}
        >
          <span className="text-lg">🏢</span> PHÂN HỆ DOANH NGHIỆP
        </button>
        <button
          type="button"
          onClick={() => {
            setSampFilterTab("ind");
            setSampActiveDetailGroup("");
          }}
          className={`flex-1 text-sm font-black py-4 px-6 rounded-xl transition-all flex items-center justify-center gap-2.5 cursor-pointer hover:scale-[1.01] ${
            sampFilterTab === "ind"
              ? "bg-amber-600 text-white shadow-lg border border-amber-700"
              : "text-slate-600 hover:text-slate-900 hover:bg-white"
          }`}
        >
          <span className="text-lg">🏡</span> PHÂN HỆ HỘ CÁ THỂ
        </button>
      </div>

      {/* KHỐI NẠP FILE TRỰC TIẾP CHO PHÂN HỆ CHỌN MẪU - ĐỘNG THEO TAB CHỦ */}
      {sampFilterTab === "corp" ? (
        /* NẠP FILE DOANH NGHIỆP */
        <div className="bg-white border-2 border-dashed border-indigo-200 hover:border-indigo-400 bg-gradient-to-br from-indigo-50/20 to-white rounded-2xl p-6 shadow-sm space-y-4 transition-all">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-indigo-100/50 pb-4">
            <div className="space-y-1">
              <span className="text-xs font-black text-indigo-700 tracking-wider uppercase font-mono flex items-center gap-1.5">
                🏢 THIẾT LẬP NGUỒN: DANH SÁCH DOANH NGHIỆP CƠ SỞ
              </span>
              <p className="text-[11px] text-slate-500 font-medium leading-normal">
                Nạp tệp Excel/CSV chứa danh sách các doanh nghiệp để thực hiện chọn mẫu theo quy mô doanh thu lũy kế tích lũy.
              </p>
            </div>
            {sampCorpFileName && (
              <button
                type="button"
                onClick={() => {
                  if (window.confirm("Bạn có chắc muốn gỡ bỏ dữ liệu doanh nghiệp hiện tại?")) {
                    setSampCorpData([]);
                    setSampCorpFileName("");
                  }
                }}
                className="bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" /> Gỡ bỏ tệp
              </button>
            )}
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-center pt-1">
            <label className="bg-white hover:bg-indigo-50/30 text-indigo-800 border-2 border-dashed border-indigo-200 shadow-sm font-extrabold text-xs px-5 py-5 rounded-2xl transition-all flex flex-col items-center justify-center gap-2 cursor-pointer hover:border-indigo-500">
              <FileUp className="w-6 h-6 text-indigo-500 animate-bounce" /> 
              <span>{sampCorpFileName ? "THAY ĐỔI FILE DOANH NGHIỆP (.XLSX, .CSV)" : "NẠP TỆP DỮ LIỆU DOANH NGHIỆP CHUYÊN BIỆT"}</span>
              <span className="text-[10px] text-slate-400 font-normal">Kéo thả hoặc click để duyệt tìm tệp tin nguồn</span>
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv" 
                onChange={(e) => handleSamplingFileUpload(e, "corp")} 
                className="hidden" 
              />
            </label>

            <div className="p-4 bg-indigo-50/40 rounded-2xl border border-indigo-100/60 text-xs text-indigo-950 space-y-2 h-full flex flex-col justify-center">
              {sampCorpFileName ? (
                <>
                  <div className="font-bold flex items-center gap-1.5 text-indigo-900">
                    <span className="text-sm">📄</span> Tệp đã nhận: <span className="underline truncate max-w-[180px]">{sampCorpFileName}</span>
                  </div>
                  <div className="font-mono text-[11px] bg-white border border-indigo-100 rounded-lg px-2.5 py-1.5 inline-block font-bold">
                    Tổng số bản ghi nạp: <span className="text-indigo-600 text-sm">{sampCorpData.length.toLocaleString()}</span> doanh nghiệp
                  </div>
                </>
              ) : (
                <div className="text-slate-500 flex flex-col items-center py-4 text-center">
                  <span className="text-lg">📁</span>
                  <span className="font-bold text-[11px] mt-1 text-slate-600">Chưa nạp tệp dữ liệu Doanh nghiệp</span>
                  <span className="text-[10px] text-slate-400 max-w-[220px] mt-0.5">Vui lòng nạp tệp danh sách cơ sở khảo sát của Doanh nghiệp để thiết lập bộ chọn mẫu độc lập</span>
                </div>
              )}
            </div>
          </div>
        </div>
      ) : (
        /* NẠP FILE HỘ CÁ THỂ */
        <div className="bg-white border-2 border-dashed border-amber-200 hover:border-amber-400 bg-gradient-to-br from-amber-50/20 to-white rounded-2xl p-6 shadow-sm space-y-4 transition-all">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-amber-100/50 pb-4">
            <div className="space-y-1">
              <span className="text-xs font-black text-amber-700 tracking-wider uppercase font-mono flex items-center gap-1.5">
                🏡 THIẾT LẬP NGUỒN: DANH SÁCH HỘ KINH DOANH CÁ THỂ
              </span>
              <p className="text-[11px] text-slate-500 font-medium leading-normal">
                Nạp tệp Excel/CSV chứa danh sách các hộ cá thể để chọn mẫu phân tầng theo định mức chuẩn của Tổng cục Thống kê (GSO).
              </p>
            </div>
            {sampIndFileName && (
              <button
                type="button"
                onClick={() => {
                  if (window.confirm("Bạn có chắc muốn gỡ bỏ dữ liệu hộ cá thể hiện tại?")) {
                    setSampIndData([]);
                    setSampIndFileName("");
                  }
                }}
                className="bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" /> Gỡ bỏ tệp
              </button>
            )}
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-center pt-1">
            <label className="bg-white hover:bg-amber-50/30 text-amber-800 border-2 border-dashed border-amber-200 shadow-sm font-extrabold text-xs px-5 py-5 rounded-2xl transition-all flex flex-col items-center justify-center gap-2 cursor-pointer hover:border-amber-500">
              <FileUp className="w-6 h-6 text-amber-500 animate-bounce" /> 
              <span>{sampIndFileName ? "THAY ĐỔI FILE HỘ CÁ THỂ (.XLSX, .CSV)" : "NẠP TỆP DỮ LIỆU HỘ CÁ THỂ CHUYÊN BIỆT"}</span>
              <span className="text-[10px] text-slate-400 font-normal">Kéo thả hoặc click để duyệt tìm tệp tin nguồn</span>
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv" 
                onChange={(e) => handleSamplingFileUpload(e, "ind")} 
                className="hidden" 
              />
            </label>

            <div className="p-4 bg-amber-50/40 rounded-2xl border border-amber-100/60 text-xs text-amber-950 space-y-2 h-full flex flex-col justify-center">
              {sampIndFileName ? (
                <>
                  <div className="font-bold flex items-center gap-1.5 text-amber-900">
                    <span className="text-sm">📄</span> Tệp đã nhận: <span className="underline truncate max-w-[180px]">{sampIndFileName}</span>
                  </div>
                  <div className="font-mono text-[11px] bg-white border border-amber-100 rounded-lg px-2.5 py-1.5 inline-block font-bold">
                    Tổng số bản ghi nạp: <span className="text-amber-600 text-sm">{sampIndData.length.toLocaleString()}</span> hộ cá thể
                  </div>
                </>
              ) : (
                <div className="text-slate-500 flex flex-col items-center py-4 text-center">
                  <span className="text-lg">📁</span>
                  <span className="font-bold text-[11px] mt-1 text-slate-600">Chưa nạp tệp dữ liệu Hộ cá thể</span>
                  <span className="text-[10px] text-slate-400 max-w-[220px] mt-0.5">Vui lòng nạp tệp danh sách cơ sở khảo sát của Hộ cá thể để thiết lập bộ chọn mẫu độc lập</span>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {(sampFilterTab === "corp" ? sampCorpData.length > 0 : sampIndData.length > 0) ? (
        <>
          {/* KHỐI ĐIỀU KHIỂN & CHẠY CHỌN MẪU HỆ THỐNG */}
          <div className={`bg-white border-2 rounded-2xl p-6 shadow-md space-y-4 relative overflow-hidden ${
            sampFilterTab === "corp" ? "border-indigo-500" : "border-amber-500"
          }`}>
            <div className={`absolute top-0 right-0 text-white font-mono text-[9px] font-black px-3 py-1 uppercase rounded-bl-xl tracking-widest animate-pulse ${
              sampFilterTab === "corp" ? "bg-indigo-500" : "bg-amber-500"
            }`}>
              ⚡ LIVE CALCULATION ACTIVE
            </div>
            
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
              <div className="space-y-1.5 max-w-2xl">
                <h3 className={`text-base font-black tracking-tight flex items-center gap-2 uppercase font-sans ${
                  sampFilterTab === "corp" ? "text-indigo-900" : "text-amber-900"
                }`}>
                  <Play className={`w-5 h-5 animate-pulse ${
                    sampFilterTab === "corp" ? "text-indigo-600 fill-indigo-600" : "text-amber-600 fill-amber-600"
                  }`} /> 
                  KHỐI ĐIỀU KHIỂN CHẠY CHỌN MẪU {sampFilterTab === "corp" ? "DOANH NGHIỆP" : "HỘ CÁ THỂ"}
                </h3>
                <p className="text-xs text-slate-600 leading-relaxed font-sans">
                  Hệ thống tự động đồng bộ tức thì. Bạn có thể <b>bấm nút kích hoạt bên phải</b> để hệ thống tái cơ cấu phân bổ, tối ưu hóa dải phân bố dải mẫu và xếp hạng nguồn đơn vị cập nhật mới nhất.
                </p>
              </div>

              <div className="flex flex-col sm:flex-row items-center gap-3 shrink-0">
                <button
                  id="btn-run-sampling"
                  onClick={handleManualRun}
                  disabled={isCalculating}
                  className={`w-full sm:w-auto px-6 py-3.5 rounded-xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-2.5 transition-all shadow-md active:scale-95 cursor-pointer border ${
                    isCalculating
                      ? "bg-slate-50 text-slate-400 cursor-not-allowed border-slate-200"
                      : sampFilterTab === "corp"
                        ? "bg-indigo-600 hover:bg-indigo-700 hover:shadow-lg border-indigo-700 text-white"
                        : "bg-amber-600 hover:bg-amber-700 hover:shadow-lg border-amber-700 text-white"
                  }`}
                >
                  {isCalculating ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin text-slate-400" />
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
              <div className={`rounded-xl p-4 flex items-center gap-3 animate-pulse ${
                sampFilterTab === "corp" ? "bg-indigo-50/50 border border-indigo-100" : "bg-amber-50/50 border border-amber-100"
              }`}>
                <div className={`h-2 w-2 rounded-full bg-current ${
                  sampFilterTab === "corp" ? "text-indigo-600 animate-ping" : "text-amber-600 animate-ping"
                }`}></div>
                <span className={`text-xs font-medium font-sans ${
                  sampFilterTab === "corp" ? "text-indigo-800" : "text-amber-800"
                }`}>Hệ thống đang chạy thuật toán phân tách, lọc dải mã ngành và phân tầng đối xứng...</span>
              </div>
            )}

            {calculationSuccess && calculationDetails && (
              <div className={`border-2 rounded-xl p-4 space-y-2 ${
                sampFilterTab === "corp" ? "bg-indigo-50/30 border-indigo-200" : "bg-amber-50/30 border-amber-200"
              }`}>
                <div className={`flex items-center gap-2 font-bold text-xs font-sans uppercase ${
                  sampFilterTab === "corp" ? "text-indigo-800" : "text-amber-800"
                }`}>
                  <span className={`text-white p-1 rounded-full text-xs ${
                    sampFilterTab === "corp" ? "bg-indigo-500" : "bg-amber-500"
                  }`}>
                    <Check className="w-4 h-4 stroke-[3]" />
                  </span>
                  <span>Đã tối ưu hóa danh sách mẫu {sampFilterTab === "corp" ? "doanh nghiệp" : "hộ cá thể"} thành công lúc {calculationDetails.time}!</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs font-sans text-slate-700 pt-1 border-t border-slate-100">
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">⚡ Chế độ:</span>
                    <span className={`font-extrabold bg-white px-2 py-0.5 rounded border font-mono text-[10px] ${
                      sampFilterTab === "corp" ? "text-indigo-700 border-indigo-200" : "text-amber-700 border-amber-200"
                    }`}>Tối ưu phân tầng</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">🎯 Mẫu chính thức:</span>
                    <span className={`font-extrabold bg-white px-2 py-0.5 rounded border font-mono text-[11px] ${
                      sampFilterTab === "corp" ? "text-indigo-700 border-indigo-200" : "text-amber-700 border-amber-200"
                    }`}>{calculationDetails.totalSelected.toLocaleString()} đơn vị</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-400">🛡️ Danh sách dự phòng:</span>
                    <span className={`font-extrabold bg-white px-2 py-0.5 rounded border font-mono text-[11px] ${
                      sampFilterTab === "corp" ? "text-indigo-700 border-indigo-200" : "text-amber-700 border-amber-200"
                    }`}>{calculationDetails.totalBackup.toLocaleString()} đơn vị</span>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Cấu hình Nhóm ngành & Hướng dẫn File tương ứng với từng Phân hệ */}
          {sampFilterTab === "corp" ? (
            <div className="bg-gradient-to-br from-indigo-50/50 to-slate-50/50 border border-indigo-100 rounded-2xl p-6 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6">
              {/* Left Col: Sector Selector */}
              <div className="lg:col-span-5 space-y-4">
                <div className="flex items-center gap-2 border-b border-indigo-100 pb-2.5">
                  <Layers className="w-4 h-4 text-indigo-600" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-700">Bộ chọn Nhóm ngành Khảo sát Doanh nghiệp</h3>
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
                  <h3 className="text-xs font-bold uppercase tracking-wider text-emerald-700">Hướng dẫn Cấu trúc File Excel &amp; Tiêu đề Cột Doanh nghiệp</h3>
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
          ) : (
            <div className="bg-gradient-to-br from-amber-50/50 to-slate-50/50 border border-amber-200 rounded-2xl p-6 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6">
              {/* Left Col: Sector Selector */}
              <div className="lg:col-span-5 space-y-4">
                <div className="flex items-center gap-2 border-b border-amber-200 pb-2.5">
                  <Layers className="w-4 h-4 text-amber-600" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-amber-700">Bộ chọn Nhóm ngành Khảo sát Hộ cá thể</h3>
                </div>
                <p className="text-[11px] text-slate-500 leading-relaxed font-sans">
                  Chọn một hoặc nhiều nhóm ngành kinh tế dưới đây để tự động lọc danh sách đối tượng chọn mẫu:
                </p>
                
                <div className="space-y-2">
                  {[
                    { id: "congnghiep", label: "Công nghiệp (Mã 05 - 33)", color: "border-amber-200 text-amber-700 bg-amber-50/50" },
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
                            className="accent-amber-600 h-3.5 w-3.5"
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
                  <div className={`p-3 rounded-xl border transition-all ${selectedSectors.includes("custom") ? "bg-amber-50/30 border-amber-200" : "bg-white border-slate-200"}`}>
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
                        className="accent-amber-600 h-3.5 w-3.5"
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
                          className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 outline-none"
                        />
                        <p className="text-[10px] text-slate-500 font-sans leading-normal">
                          Ngăn cách các dải mã bằng dấu phẩy. Chấp nhận dạng dải <code>05-33</code> hoặc mã đơn lẻ <code>49</code>.
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Right Col: Household Sampling Guide */}
              <div className="lg:col-span-7 space-y-4">
                <div className="flex items-center gap-2 border-b border-amber-200 pb-2.5">
                  <FileCheck className="w-4 h-4 text-amber-600" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-amber-700">Nguyên tắc Chọn Mẫu Hộ cá thể (Chuẩn TCTK)</h3>
                </div>
                <p className="text-[11px] text-slate-500 leading-relaxed font-sans">
                  Việc chọn mẫu hộ cá thể được cấu trúc chặt chẽ dựa trên số lượng cơ sở thực tế của từng địa bàn xã &amp; nhóm ngành tương ứng để tối ưu dung lượng mẫu điều tra:
                </p>

                <div className="border border-amber-200 rounded-xl overflow-hidden shadow-inner bg-white">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-amber-50 text-[10px] font-bold text-amber-700 uppercase tracking-wider border-b border-amber-200 font-sans">
                        <th className="p-2.5">Quy mô nhóm ngành</th>
                        <th className="p-2.5">Nguyên tắc chọn mẫu mặc định</th>
                        <th className="p-2.5">Tùy chỉnh linh hoạt</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-sans text-slate-700 text-[11px]">
                      <tr className="hover:bg-amber-50/10">
                        <td className="p-2.5 font-bold text-slate-800">
                          1 - 5 cơ sở
                        </td>
                        <td className="p-2.5 text-amber-800 font-medium">
                          Chọn toàn bộ cơ sở
                        </td>
                        <td className="p-2.5 text-slate-500">
                          Sử dụng thanh kéo hoặc chọn tất cả
                        </td>
                      </tr>
                      <tr className="hover:bg-amber-50/10">
                        <td className="p-2.5 font-bold text-slate-800">
                          6 - 100 cơ sở
                        </td>
                        <td className="p-2.5 text-amber-800 font-medium">
                          Chọn cố định 5 cơ sở
                        </td>
                        <td className="p-2.5 text-slate-500">
                          Kéo tăng giảm định mức từ 1 - 30
                        </td>
                      </tr>
                      <tr className="hover:bg-amber-50/10">
                        <td className="p-2.5 font-bold text-slate-800">
                          101 - 1000 cơ sở
                        </td>
                        <td className="p-2.5 text-amber-800 font-medium">
                          Chọn cố định 8 cơ sở
                        </td>
                        <td className="p-2.5 text-slate-500">
                          Kéo tăng giảm định mức từ 1 - 100
                        </td>
                      </tr>
                      <tr className="hover:bg-amber-50/10">
                        <td className="p-2.5 font-bold text-slate-800">
                          Từ 1001 cơ sở trở lên
                        </td>
                        <td className="p-2.5 text-amber-800 font-medium">
                          Chọn tỷ lệ 1% tổng số cơ sở
                        </td>
                        <td className="p-2.5 text-slate-500">
                          Kéo tăng giảm tỷ lệ từ 0.1% - 10%
                        </td>
                      </tr>
                      <tr className="hover:bg-amber-50/10">
                        <td className="p-2.5 font-bold text-indigo-800">
                          📍 Đặc thù: Vận tải, kho bãi (49)
                        </td>
                        <td className="p-2.5 text-indigo-700 font-bold">
                          Chọn 1.5% (Tối đa không quá 50 cơ sở)
                        </td>
                        <td className="p-2.5 text-slate-500">
                          Kéo tỷ lệ 0.5% - 15% &amp; Tối đa 10 - 200 cơ sở
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div className="bg-amber-50 text-amber-800 border border-amber-100 p-2.5 rounded-xl text-[10px] leading-relaxed flex items-start gap-1.5 font-sans">
                  <span className="text-xs">💡</span>
                  <span><b>Phân hệ Chuyên biệt:</b> Dữ liệu hộ cá thể sẽ tự động áp dụng các quy tắc phân tầng trên. Bạn có thể sử dụng các thanh trượt và nút kéo bên dưới để tinh chỉnh tham số tính toán của Tổng cục Thống kê ngay lập tức!</span>
                </div>
              </div>
            </div>
          )}

          {/* Setup parameters / mapping */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            
            {/* Column mappings */}
            <div className={`bg-white border rounded-2xl p-5 space-y-4 shadow-sm hover:scale-[1.01] transition-all ${
              sampFilterTab === "corp" ? "border-indigo-200" : "border-amber-200"
            }`}>
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2.5">
                <Sliders className={`w-4 h-4 ${sampFilterTab === "corp" ? "text-indigo-600" : "text-amber-600"}`} />
                <h3 className={`text-xs font-bold uppercase tracking-wider ${sampFilterTab === "corp" ? "text-indigo-600" : "text-amber-600"}`}>
                  1. Cấu hình cột khảo sát
                </h3>
              </div>
              
              <div className="space-y-3 text-xs">
                {sampFilterTab === "corp" ? (
                  /* COLUMN MAPPINGS FOR ENTERPRISE */
                  <>
                    <div>
                      <label className="block text-indigo-950 font-bold text-xs mb-1">Cột Khóa định danh (MST/ID)</label>
                      <select
                        value={corpIdCol}
                        onChange={(e) => setCorpIdCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột định danh DN --</option>
                        {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-indigo-950 font-bold text-xs mb-1">Cột Địa bàn xã / Địa bàn</label>
                      <select
                        value={corpXaCol}
                        onChange={(e) => setCorpXaCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột địa bàn DN --</option>
                        {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-indigo-950 font-bold text-xs mb-1">Cột Mã ngành VSIC</label>
                      <select
                        value={corpManganhCol}
                        onChange={(e) => setCorpManganhCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột mã ngành DN --</option>
                        {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-indigo-950 font-bold text-xs mb-1">Cột Doanh thu / Sản lượng</label>
                      <select
                        value={corpDoanhThuCol}
                        onChange={(e) => setCorpDoanhThuCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột doanh thu DN --</option>
                        {corpColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                  </>
                ) : (
                  /* COLUMN MAPPINGS FOR HOUSEHOLD */
                  <>
                    <div>
                      <label className="block text-amber-950 font-bold text-xs mb-1">Cột Khóa định danh (MST/ID)</label>
                      <select
                        value={indIdCol}
                        onChange={(e) => setIndIdCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 focus:border-amber-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột định danh Hộ --</option>
                        {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-amber-950 font-bold text-xs mb-1">Cột Địa bàn xã / Địa bàn</label>
                      <select
                        value={indXaCol}
                        onChange={(e) => setIndXaCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 focus:border-amber-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột địa bàn Hộ --</option>
                        {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-amber-950 font-bold text-xs mb-1">Cột Mã ngành VSIC</label>
                      <select
                        value={indManganhCol}
                        onChange={(e) => setIndManganhCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 focus:border-amber-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột mã ngành Hộ --</option>
                        {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="block text-amber-950 font-bold text-xs mb-1">Cột Doanh thu / Sản lượng</label>
                      <select
                        value={indDoanhThuCol}
                        onChange={(e) => setIndDoanhThuCol(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800 focus:ring-1 focus:ring-amber-500 focus:border-amber-500 outline-none shadow-sm"
                      >
                        <option value="">-- Chọn cột doanh thu Hộ --</option>
                        {indColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                  </>
                )}

                <div className={`p-2.5 rounded-lg text-[10px] leading-relaxed mt-2 shadow-sm border ${
                  sampFilterTab === "corp" 
                    ? "bg-indigo-50/50 text-indigo-900 border-indigo-100" 
                    : "bg-amber-50/50 text-amber-900 border-amber-100"
                }`}>
                  <span className="font-bold">⚠️ Lưu ý Nhóm ngành lọc:</span> Hệ thống tự động tách lấy 2 chữ số đầu của mã ngành và lọc giữ lại theo <b>Bộ chọn Nhóm ngành Khảo sát</b> đang kích hoạt ở trên. Các dòng không thuộc nhóm ngành đã chọn sẽ tự động được loại bỏ khỏi danh sách mẫu.
                </div>
              </div>
            </div>

            {/* Cột 2 & 3: Nội dung thay đổi động dựa theo Tab chuyển đổi */}
            <div className="lg:col-span-2">
              {(sampFilterTab === "corp" || sampFilterTab === "all") ? (
                /* PHÂN HỆ DOANH NGHIỆP */
                <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm hover:scale-[1.01] transition-all h-full">
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
              ) : (
                /* PHÂN HỆ HỘ CÁ THỂ - TRANG CẤU HÌNH KÉO THẢ RIÊNG BIỆT */
                <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-5 shadow-sm hover:scale-[1.01] transition-all h-full">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                    <div className="flex items-center gap-2">
                      <Sliders className="w-4 h-4 text-amber-600" />
                      <h3 className="text-xs font-bold uppercase tracking-wider text-amber-600">3. Cấu hình chọn mẫu Hộ cá thể</h3>
                    </div>
                    <span className="text-[9.5px] bg-amber-100 text-amber-800 font-black font-sans px-2.5 py-0.5 rounded-lg border border-amber-200 shadow-inner">
                      BẢNG ĐIỀU KHIỂN KÉO THẢ ĐỘNG
                    </span>
                  </div>

                  {/* Giới thiệu cách tính */}
                  <div className="bg-amber-50/50 border border-amber-100 p-3.5 rounded-xl text-[10.5px] text-amber-900 leading-relaxed space-y-1 shadow-sm font-sans">
                    <p className="font-extrabold flex items-center gap-1">
                      <span>💡</span> Hướng dẫn xác định cỡ mẫu cá thể động:
                    </p>
                    <p className="pl-4">
                      Hộ cá thể tự động gom nhóm theo <b>Địa bàn xã và Mã ngành cấp 2 (VSIC)</b>. Với mỗi nhóm, số lượng mẫu được xác định dựa theo các phân khúc quy mô dưới đây. Hãy kéo các nút trượt để điều chỉnh trực tiếp số lượng mẫu mong muốn.
                    </p>
                  </div>

                  <div className="space-y-4 text-xs font-sans">
                    {/* Phân khúc 1: 1 - 5 cơ sở */}
                    <div className="bg-slate-50/50 border border-slate-200/60 p-3.5 rounded-xl space-y-2.5 shadow-inner">
                      <div className="flex items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-[11.5px] font-black text-slate-800 block">1. Nhóm ngành cực nhỏ (1 - 5 cơ sở)</span>
                          <span className="text-[9.5px] text-slate-500 block leading-normal">Chuẩn TCTK: Chọn toàn bộ (100%) để đảm bảo đại diện nhóm ngành.</span>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700 bg-white border border-slate-200 px-3 py-1.5 rounded-xl shadow-sm hover:bg-slate-50">
                            <input
                              type="checkbox"
                              checked={indSize1To5All}
                              onChange={(e) => setIndSize1To5All(e.target.checked)}
                              className="accent-amber-600 h-3.5 w-3.5 rounded"
                            />
                            <span>Chọn toàn bộ</span>
                          </label>
                        </div>
                      </div>

                      {!indSize1To5All && (
                        <div className="space-y-1.5 pt-2 border-t border-slate-200/50">
                          <div className="flex justify-between items-center text-xs">
                            <span className="text-slate-500 font-medium">Số hộ lấy tối đa trong nhóm:</span>
                            <span className="font-extrabold text-amber-700 bg-amber-50 px-2 py-0.5 rounded font-mono border border-amber-100">{indSize1To5Value} hộ</span>
                          </div>
                          <input
                            type="range"
                            min="1"
                            max="5"
                            value={indSize1To5Value}
                            onChange={(e) => setIndSize1To5Value(parseInt(e.target.value) || 1)}
                            className="w-full accent-amber-600 cursor-pointer h-1.5 bg-slate-200 rounded"
                          />
                        </div>
                      )}
                    </div>

                    {/* Phân khúc 2: 6 - 100 cơ sở */}
                    <div className="bg-slate-50/50 border border-slate-200/60 p-3.5 rounded-xl space-y-2.5 shadow-inner">
                      <div className="flex justify-between items-center gap-4">
                        <div className="space-y-0.5">
                          <span className="text-[11.5px] font-black text-slate-800 block">2. Nhóm ngành nhỏ (6 - 100 cơ sở)</span>
                          <span className="text-[9.5px] text-slate-500 block leading-normal">Chuẩn TCTK: Chọn cố định 5 cơ sở/nhóm ngành.</span>
                        </div>
                        <span className="font-black text-amber-700 bg-amber-50 px-2.5 py-1 rounded-xl font-mono text-xs border border-amber-100 shrink-0">
                          {indSize6To100Value} cơ sở
                        </span>
                      </div>
                      <div className="space-y-1">
                        <input
                          type="range"
                          min="1"
                          max="30"
                          value={indSize6To100Value}
                          onChange={(e) => setIndSize6To100Value(parseInt(e.target.value) || 5)}
                          className="w-full accent-amber-600 cursor-pointer h-1.5 bg-slate-200 rounded"
                        />
                        <div className="flex justify-between text-[9px] text-slate-400 font-bold font-mono">
                          <span>1 cơ sở</span>
                          <span>Mặc định: 5</span>
                          <span>30 cơ sở</span>
                        </div>
                      </div>
                    </div>

                    {/* Phân khúc 3: 101 - 1000 cơ sở */}
                    <div className="bg-slate-50/50 border border-slate-200/60 p-3.5 rounded-xl space-y-2.5 shadow-inner">
                      <div className="flex justify-between items-center gap-4">
                        <div className="space-y-0.5">
                          <span className="text-[11.5px] font-black text-slate-800 block">3. Nhóm ngành trung bình (101 - 1000 cơ sở)</span>
                          <span className="text-[9.5px] text-slate-500 block leading-normal">Chuẩn TCTK: Chọn cố định 8 cơ sở/nhóm ngành.</span>
                        </div>
                        <span className="font-black text-amber-700 bg-amber-50 px-2.5 py-1 rounded-xl font-mono text-xs border border-amber-100 shrink-0">
                          {indSize101To1000Value} cơ sở
                        </span>
                      </div>
                      <div className="space-y-1">
                        <input
                          type="range"
                          min="1"
                          max="100"
                          value={indSize101To1000Value}
                          onChange={(e) => setIndSize101To1000Value(parseInt(e.target.value) || 8)}
                          className="w-full accent-amber-600 cursor-pointer h-1.5 bg-slate-200 rounded"
                        />
                        <div className="flex justify-between text-[9px] text-slate-400 font-bold font-mono">
                          <span>1 cơ sở</span>
                          <span>Mặc định: 8</span>
                          <span>100 cơ sở</span>
                        </div>
                      </div>
                    </div>

                    {/* Phân khúc 4: Từ 1001 cơ sở trở lên */}
                    <div className="bg-slate-50/50 border border-slate-200/60 p-3.5 rounded-xl space-y-3.5 shadow-inner">
                      <div className="flex justify-between items-center gap-4">
                        <div className="space-y-0.5">
                          <span className="text-[11.5px] font-black text-slate-800 block">4. Nhóm ngành lớn (Từ 1001 cơ sở trở lên)</span>
                          <span className="text-[9.5px] text-slate-500 block leading-normal">Chuẩn TCTK: Chọn tỷ lệ 1% tổng số cơ sở trong nhóm ngành.</span>
                        </div>
                        <span className="font-black text-amber-700 bg-amber-50 px-2.5 py-1 rounded-xl font-mono text-xs border border-amber-100 shrink-0">
                          Tỷ lệ: {indSize1001PlusPercent}%
                        </span>
                      </div>
                      <div className="space-y-1">
                        <input
                          type="range"
                          min="0.1"
                          max="10"
                          step="0.1"
                          value={indSize1001PlusPercent}
                          onChange={(e) => setIndSize1001PlusPercent(parseFloat(e.target.value) || 1)}
                          className="w-full accent-amber-600 cursor-pointer h-1.5 bg-slate-200 rounded"
                        />
                        <div className="flex justify-between text-[9px] text-slate-400 font-bold font-mono">
                          <span>0.1%</span>
                          <span>Mặc định: 1%</span>
                          <span>10%</span>
                        </div>
                      </div>

                      {/* Phân khúc đặc thù Vận tải, kho bãi */}
                      <div className="bg-amber-50/40 border border-amber-200/60 p-3 rounded-xl space-y-3.5">
                        <div className="flex justify-between items-center gap-4">
                          <div className="space-y-0.5">
                            <span className="text-[10.5px] font-extrabold text-amber-900 block">📍 Nhóm ngành Vận tải, kho bãi đặc thù (VSIC 49-53)</span>
                            <span className="text-[9px] text-amber-750 block leading-normal">Cơ sở nhỏ thường ít nên khuyến nghị chọn 1.5% tổng số cơ sở.</span>
                          </div>
                          <span className="font-black text-amber-800 bg-white px-2 py-0.5 rounded-lg font-mono text-[10.5px] border border-amber-200 shrink-0">
                            Tỷ lệ: {indTransportPercent}%
                          </span>
                        </div>
                        <div className="space-y-1">
                          <input
                            type="range"
                            min="0.5"
                            max="15"
                            step="0.1"
                            value={indTransportPercent}
                            onChange={(e) => setIndTransportPercent(parseFloat(e.target.value) || 1.5)}
                            className="w-full accent-amber-700 cursor-pointer h-1.5 bg-amber-200/30 rounded"
                          />
                        </div>

                        <div className="flex justify-between items-center gap-4 pt-1.5 border-t border-amber-200/40">
                          <span className="text-[10px] font-extrabold text-amber-900">Giới hạn tối đa không quá:</span>
                          <span className="font-black text-amber-800 bg-white px-2 py-0.5 rounded-lg font-mono text-[10.5px] border border-amber-200 shrink-0">
                            Tối đa {indTransportMaxCap} cơ sở
                          </span>
                        </div>
                        <div className="space-y-1">
                          <input
                            type="range"
                            min="10"
                            max="200"
                            value={indTransportMaxCap}
                            onChange={(e) => setIndTransportMaxCap(parseInt(e.target.value) || 50)}
                            className="w-full accent-amber-700 cursor-pointer h-1.5 bg-amber-200/30 rounded"
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Nút khôi phục mặc định */}
                  <div className="flex justify-end pt-2 border-t border-slate-100">
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
                      }}
                      className="text-[10.5px] text-amber-700 hover:text-amber-950 font-black hover:underline cursor-pointer flex items-center gap-1.5"
                    >
                      🔄 Khôi phục thiết lập mặc định của TCTK
                    </button>
                  </div>
                </div>
              )}
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

              {/* Bộ chỉ báo trạng thái phân hệ hiện hành của danh sách nhóm */}
              <div className={`p-2.5 rounded-xl border flex items-center justify-between font-sans text-[11px] font-bold ${
                sampFilterTab === "corp"
                  ? "bg-indigo-50/40 border-indigo-100 text-indigo-950"
                  : "bg-amber-50/40 border-amber-100 text-amber-950"
              }`}>
                <span className="flex items-center gap-1.5">
                  {sampFilterTab === "corp" ? "🏢 Phân hệ:" : "🏡 Phân hệ:"}
                  <span className={`px-2 py-0.5 rounded-lg text-[10px] font-black uppercase tracking-wider ${
                    sampFilterTab === "corp" ? "bg-indigo-600 text-white" : "bg-amber-600 text-white"
                  }`}>
                    {sampFilterTab === "corp" ? "Doanh nghiệp" : "Hộ cá thể"}
                  </span>
                </span>
                <span className="text-slate-400 font-normal">Tự động phân nhóm theo xã & VSIC</span>
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
                  
                  const allSelected = sampFilterTab === "corp" ? corpSelected : indSelected;
                  const allBackup = sampFilterTab === "corp" ? corpBackup : indBackup;
                  
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

                      {/* Thanh trượt kéo thả cấu hình mẫu cá thể trực quan */}
                      {grp.indGrp && grp.indGrp.totalN > 0 && (
                        <div className="bg-amber-50/60 border border-amber-200 rounded-2xl p-4 space-y-3 shadow-sm animate-fade-in">
                          <div className="flex items-center justify-between gap-2">
                            <div className="flex items-center gap-2 font-bold text-amber-800">
                              <span className="p-1 bg-amber-100 rounded-lg text-amber-700">🏡</span>
                              <span>Kéo điều chỉnh số lượng mẫu Hộ cá thể (Ngành {grp.vsicL2})</span>
                            </div>
                            <div className="text-right shrink-0">
                              <span className="text-[11px] bg-amber-100 text-amber-800 font-black font-mono px-2.5 py-1 rounded-lg border border-amber-200 shadow-inner">
                                {indSelected.length} / {grp.indGrp.totalN} hộ
                              </span>
                            </div>
                          </div>
                          
                          <div className="space-y-1.5">
                            <div className="flex items-center gap-3">
                              <span className="text-[10px] font-bold text-slate-400 font-mono w-3 text-center">0</span>
                              <input
                                type="range"
                                min="0"
                                max={grp.indGrp.totalN}
                                value={grp.indGrp.targetSampleSize}
                                onChange={(e) => {
                                  const val = parseInt(e.target.value) || 0;
                                  setManualIndGroupSamples(prev => ({
                                    ...prev,
                                    [grp.key]: val
                                  }));
                                }}
                                className="flex-1 accent-amber-600 cursor-pointer h-2 bg-amber-200/50 rounded-lg appearance-none"
                              />
                              <span className="text-[10px] font-bold text-slate-400 font-mono w-6 text-center">{grp.indGrp.totalN}</span>
                            </div>
                            
                            <div className="flex justify-between items-center text-[10.5px]">
                              <span className="text-slate-500 font-medium">
                                Mặc định GSO: <strong className="text-slate-700 font-extrabold">{
                                  (() => {
                                    const totalN = grp.indGrp.totalN;
                                    const isTransport = ["49", "50", "51", "52", "53"].includes(grp.vsicL2);
                                    if (totalN <= 5) return totalN;
                                    if (totalN <= 100) return Math.min(5, totalN);
                                    if (totalN <= 1000) return Math.min(8, totalN);
                                    if (isTransport) return Math.min(50, Math.max(8, Math.round(totalN * 0.015)));
                                    return Math.max(8, Math.round(totalN * 0.01));
                                  })()
                                } hộ</strong>
                              </span>
                              
                              {manualIndGroupSamples[grp.key] !== undefined && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setManualIndGroupSamples(prev => {
                                      const updated = { ...prev };
                                      delete updated[grp.key];
                                      return updated;
                                    });
                                  }}
                                  className="text-amber-700 hover:text-amber-900 font-black hover:underline cursor-pointer flex items-center gap-0.5"
                                >
                                  🔄 Khôi phục mặc định
                                </button>
                              )}
                            </div>
                          </div>
                        </div>
                      )}

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
