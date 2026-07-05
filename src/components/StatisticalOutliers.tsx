import React, { useState, useEffect, useMemo } from "react";
import { 
  Sliders, 
  AlertTriangle, 
  TrendingUp, 
  BarChart3, 
  Download, 
  Info, 
  Search, 
  Zap, 
  CheckCircle2,
  ChevronDown,
  ChevronUp
} from "lucide-react";
import { 
  ResponsiveContainer, 
  ScatterChart, 
  Scatter, 
  XAxis, 
  YAxis, 
  ZAxis, 
  CartesianGrid, 
  Tooltip, 
  ReferenceLine,
  Cell,
  Label
} from "recharts";
import * as XLSX from "xlsx";
import { LiveTenRowPreview } from "./LiveTenRowPreview";

interface ColumnMapping {
  mota: string;
  manganh: string;
  xa: string;
  doanhthu: string;
  laodong: string;
  idCol: string;
}

interface StatisticalOutliersProps {
  mainData: any[];
  columns: string[];
  mapping: ColumnMapping;
  onFilterRows?: (indices: number[], label: string) => void;
  onExportExcel?: (data: any[], fileName: string) => void;
  onUpdateMainData?: (newData: any[]) => void;
}

export default function StatisticalOutliers({ 
  mainData, 
  columns, 
  mapping, 
  onFilterRows, 
  onExportExcel,
  onUpdateMainData
}: StatisticalOutliersProps) {
  // Config states
  const [targetCol, setTargetCol] = useState<string>("");
  const [groupCol, setGroupCol] = useState<string>("");
  const [ratioCol, setRatioCol] = useState<string>("");
  const [sigmaThreshold, setSigmaThreshold] = useState<number>(3);
  const [minSampleCount, setMinSampleCount] = useState<number>(3);
  const [groupLevel, setGroupLevel] = useState<string>("all"); // "all" for exact, or number of digits if VSIC

  // UI states
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [hasRun, setHasRun] = useState<boolean>(false);
  const [outlierResults, setOutlierResults] = useState<any[]>([]);
  const [overallStats, setOverallStats] = useState<any>(null);
  const [searchFilter, setSearchFilter] = useState<string>("");
  const [groupStatsList, setGroupStatsList] = useState<any[]>([]);
  const [showConfig, setShowConfig] = useState<boolean>(true);

  // Auto-select columns based on mapping without overwriting active user selections
  useEffect(() => {
    if (columns && columns.length > 0) {
      setTargetCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        return mapping.doanhthu || columns.find(c => {
          const lower = c.toLowerCase();
          return lower.includes("doanh thu") || lower.includes("doanhthu") || lower.includes("thu nhập") || lower.includes("sản lượng") || lower.includes("trị giá");
        }) || "";
      });

      setGroupCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        return mapping.manganh || columns.find(c => {
          const lower = c.toLowerCase();
          return lower.includes("ngành") || lower.includes("manganh") || lower.includes("vsic") || lower.includes("phân loại");
        }) || "";
      });

      setRatioCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        return mapping.laodong || columns.find(c => {
          const lower = c.toLowerCase();
          return lower.includes("lao động") || lower.includes("laodong") || lower.includes("nhân sự") || lower.includes("số người");
        }) || "";
      });
    }
  }, [columns, mapping]);

  // Execute Anomaly detection
  const handleAnalyze = () => {
    if (mainData.length === 0) return;
    if (!targetCol) {
      alert("Vui lòng chọn cột số lượng cần phân tích!");
      return;
    }

    setIsAnalyzing(true);

    setTimeout(() => {
      try {
        const validRows: any[] = [];
        const groupValuesMap: { [key: string]: number[] } = {};
        const allValues: number[] = [];

        // 1. Process rows and extract numerical values
        mainData.forEach((row, idx) => {
          // Parse numerical target
          const rawTargetStr = String(row[targetCol] !== undefined && row[targetCol] !== null ? row[targetCol] : "").replace(/,/g, "").replace(/[^0-9.\-]/g, "");
          const rawTarget = parseFloat(rawTargetStr);
          if (isNaN(rawTarget)) return;

          // Parse ratio column if active
          let ratioNum = 1;
          if (ratioCol) {
            const rawRatioStr = String(row[ratioCol] !== undefined && row[ratioCol] !== null ? row[ratioCol] : "").replace(/,/g, "").replace(/[^0-9.\-]/g, "");
            const parsedRatio = parseFloat(rawRatioStr);
            if (!isNaN(parsedRatio) && parsedRatio > 0) {
              ratioNum = parsedRatio;
            }
          }

          const finalValue = ratioCol ? rawTarget / ratioNum : rawTarget;

          // Extract group identifier
          let groupKey = "TẤT CẢ";
          if (groupCol) {
            const rawGrp = String(row[groupCol] || "Không phân lớp").trim();
            if (groupLevel !== "all" && /^\d+$/.test(rawGrp)) {
              const digitCount = parseInt(groupLevel, 10);
              groupKey = rawGrp.substring(0, digitCount);
            } else {
              groupKey = rawGrp;
            }
          }

          validRows.push({
            originalIdx: idx,
            row,
            group: groupKey,
            rawVal: rawTarget,
            ratioVal: ratioNum,
            finalVal: finalValue
          });

          allValues.push(finalValue);
          if (!groupValuesMap[groupKey]) {
            groupValuesMap[groupKey] = [];
          }
          groupValuesMap[groupKey].push(finalValue);
        });

        if (validRows.length === 0) {
          alert(`Không tìm thấy giá trị số hợp lệ nào trong cột [${targetCol}]!`);
          setIsAnalyzing(false);
          return;
        }

        // 2. Compute overall parameters
        const overallCount = allValues.length;
        const overallSum = allValues.reduce((a, b) => a + b, 0);
        const overallMean = overallSum / overallCount;
        const overallVariance = overallCount > 1
          ? allValues.reduce((acc, v) => acc + Math.pow(v - overallMean, 2), 0) / (overallCount - 1)
          : 0;
        const overallStdDev = Math.sqrt(overallVariance);

        // 3. Compute group parameters
        const groupStats: { [group: string]: { mean: number; stdDev: number; count: number; useGroup: boolean } } = {};
        const groupStatsArr: any[] = [];

        Object.keys(groupValuesMap).forEach(groupName => {
          const vals = groupValuesMap[groupName];
          const count = vals.length;

          if (count >= minSampleCount) {
            const gSum = vals.reduce((a, b) => a + b, 0);
            const gMean = gSum / count;
            const gVariance = count > 1
              ? vals.reduce((acc, v) => acc + Math.pow(v - gMean, 2), 0) / (count - 1)
              : 0;
            const gStdDev = Math.sqrt(gVariance);

            groupStats[groupName] = {
              mean: gMean,
              stdDev: gStdDev,
              count,
              useGroup: gStdDev > 0
            };

            groupStatsArr.push({
              groupName,
              mean: gMean,
              stdDev: gStdDev,
              count,
              min: Math.min(...vals),
              max: Math.max(...vals)
            });
          } else {
            // Sample too small, fall back to global stats
            groupStats[groupName] = {
              mean: overallMean,
              stdDev: overallStdDev,
              count,
              useGroup: false
            };
          }
        });

        // 4. Filter outliers
        const detectedOutliers: any[] = [];
        validRows.forEach(item => {
          const stats = groupStats[item.group] || { mean: overallMean, stdDev: overallStdDev, useGroup: false };
          const mean = stats.mean;
          const stdDev = stats.stdDev || 1; // Safeguard against zero std dev

          const zScore = (item.finalVal - mean) / stdDev;
          const absZ = Math.abs(zScore);

          if (absZ >= sigmaThreshold) {
            // Calculate ratio to group mean
            const ratioToMean = item.finalVal / (mean || 1);
            let suspicionMsg = "Mức độ trung bình";
            let suspicionColor = "text-amber-600";
            let labelShort = "Vượt lệch";

            if (ratioToMean >= 10) {
              suspicionColor = "text-red-600 font-bold";
              if (item.finalVal % 10 === 0 && (item.finalVal / 10) % 10 === 0) {
                suspicionMsg = `Nghi vấn gõ thừa số 0 cực kỳ cao (Gấp ${Math.round(ratioToMean)} lần bình quân nhóm). Giá trị tròn trăm/tròn chục.`;
                labelShort = "Nghi thừa số 0 (X100)";
              } else {
                suspicionMsg = `Nghi vấn sai lệch lớn (Gấp ${Math.round(ratioToMean)} lần bình quân nhóm)`;
                labelShort = "Gấp >10 lần trung bình";
              }
            } else if (ratioToMean <= 0.05) {
              suspicionColor = "text-sky-600 font-bold";
              suspicionMsg = `Nghi vấn ghi thiếu số hoặc sai đơn vị tính (Chỉ bằng ${Math.round(ratioToMean * 100)}% bình quân nhóm)`;
              labelShort = "Nghi thiếu số / sai đơn vị";
            } else {
              suspicionMsg = `Bất thường vượt quá ${sigmaThreshold} lần độ lệch chuẩn nhóm.`;
              labelShort = "Sai lệch phân phối";
            }

            detectedOutliers.push({
              ...item.row,
              _rowIdx: item.originalIdx, // 0-based for tracking
              _rowNum: item.originalIdx + 1, // 1-based for UI
              _finalVal: item.finalVal,
              _rawVal: item.rawVal,
              _ratioVal: item.ratioVal,
              _group: item.group,
              _groupMean: mean,
              _groupStdDev: stdDev,
              _zScore: zScore,
              _zScoreAbs: absZ,
              _ratioToMean: ratioToMean,
              _isGroupStat: stats.useGroup,
              _suspicionMsg: suspicionMsg,
              _suspicionColor: suspicionColor,
              _labelShort: labelShort
            });
          }
        });

        // Sort by absolute Z-Score descending
        detectedOutliers.sort((a, b) => b._zScoreAbs - a._zScoreAbs);

        setOutlierResults(detectedOutliers);
        setGroupStatsList(groupStatsArr.sort((a, b) => b.count - a.count));
        setOverallStats({
          totalScanned: validRows.length,
          outliersFound: detectedOutliers.length,
          outlierRate: validRows.length > 0 ? ((detectedOutliers.length / validRows.length) * 100).toFixed(2) + "%" : "0%",
          mean: overallMean,
          stdDev: overallStdDev,
          targetCol,
          groupCol,
          ratioCol
        });
        setHasRun(true);
        setShowConfig(false); // Collapsed settings to save vertical space
      } catch (err: any) {
        console.error(err);
        alert("Lỗi thực hiện rà quét sai lệch phân phối: " + err.message);
      } finally {
        setIsAnalyzing(false);
      }
    }, 150);
  };

  // Quiet auto-recalculate of outliers when mainData changes
  useEffect(() => {
    if (hasRun && mainData.length > 0 && targetCol) {
      try {
        const validRows: any[] = [];
        const groupValuesMap: { [key: string]: number[] } = {};
        const allValues: number[] = [];

        mainData.forEach((row, idx) => {
          const rawTargetStr = String(row[targetCol] !== undefined && row[targetCol] !== null ? row[targetCol] : "").replace(/,/g, "").replace(/[^0-9.\-]/g, "");
          const rawTarget = parseFloat(rawTargetStr);
          if (isNaN(rawTarget)) return;

          let ratioNum = 1;
          if (ratioCol) {
            const rawRatioStr = String(row[ratioCol] !== undefined && row[ratioCol] !== null ? row[ratioCol] : "").replace(/,/g, "").replace(/[^0-9.\-]/g, "");
            const parsedRatio = parseFloat(rawRatioStr);
            if (!isNaN(parsedRatio) && parsedRatio > 0) {
              ratioNum = parsedRatio;
            }
          }

          const finalValue = ratioCol ? rawTarget / ratioNum : rawTarget;

          let groupKey = "TẤT CẢ";
          if (groupCol) {
            const rawGrp = String(row[groupCol] || "Không phân lớp").trim();
            if (groupLevel !== "all" && /^\d+$/.test(rawGrp)) {
              const digitCount = parseInt(groupLevel, 10);
              groupKey = rawGrp.substring(0, digitCount);
            } else {
              groupKey = rawGrp;
            }
          }

          validRows.push({
            originalIdx: idx,
            row,
            group: groupKey,
            rawVal: rawTarget,
            ratioVal: ratioNum,
            finalVal: finalValue
          });

          allValues.push(finalValue);
          if (!groupValuesMap[groupKey]) {
            groupValuesMap[groupKey] = [];
          }
          groupValuesMap[groupKey].push(finalValue);
        });

        if (validRows.length === 0) return;

        const overallCount = allValues.length;
        const overallSum = allValues.reduce((a, b) => a + b, 0);
        const overallMean = overallSum / overallCount;
        const overallVariance = overallCount > 1
          ? allValues.reduce((acc, v) => acc + Math.pow(v - overallMean, 2), 0) / (overallCount - 1)
          : 0;
        const overallStdDev = Math.sqrt(overallVariance);

        const groupStats: { [group: string]: { mean: number; stdDev: number; count: number; useGroup: boolean } } = {};
        const groupStatsArr: any[] = [];

        Object.keys(groupValuesMap).forEach(groupName => {
          const vals = groupValuesMap[groupName];
          const count = vals.length;

          if (count >= minSampleCount) {
            const gSum = vals.reduce((a, b) => a + b, 0);
            const gMean = gSum / count;
            const gVariance = count > 1
              ? vals.reduce((acc, v) => acc + Math.pow(v - gMean, 2), 0) / (count - 1)
              : 0;
            const gStdDev = Math.sqrt(gVariance);

            groupStats[groupName] = {
              mean: gMean,
              stdDev: gStdDev,
              count,
              useGroup: gStdDev > 0
            };

            groupStatsArr.push({
              groupName,
              mean: gMean,
              stdDev: gStdDev,
              count,
              min: Math.min(...vals),
              max: Math.max(...vals)
            });
          } else {
            groupStats[groupName] = {
              mean: overallMean,
              stdDev: overallStdDev,
              count,
              useGroup: false
            };
          }
        });

        const detectedOutliers: any[] = [];
        validRows.forEach(item => {
          const stats = groupStats[item.group] || { mean: overallMean, stdDev: overallStdDev, useGroup: false };
          const mean = stats.mean;
          const stdDev = stats.stdDev || 1;

          const zScore = (item.finalVal - mean) / stdDev;
          const absZ = Math.abs(zScore);

          if (absZ >= sigmaThreshold) {
            const ratioToMean = item.finalVal / (mean || 1);
            let suspicionMsg = "Mức độ trung bình";
            let suspicionColor = "text-amber-600";
            let labelShort = "Vượt lệch";

            if (ratioToMean >= 10) {
              suspicionColor = "text-red-600 font-bold";
              if (item.finalVal % 10 === 0 && (item.finalVal / 10) % 10 === 0) {
                suspicionMsg = `Nghi vấn gõ thừa số 0 cực kỳ cao (Gấp ${Math.round(ratioToMean)} lần bình quân nhóm). Giá trị tròn trăm/tròn chục.`;
                labelShort = "Nghi thừa số 0 (X100)";
              } else {
                suspicionMsg = `Nghi vấn sai lệch lớn (Gấp ${Math.round(ratioToMean)} lần bình quân nhóm)`;
                labelShort = "Gấp >10 lần trung bình";
              }
            } else if (ratioToMean <= 0.05) {
              suspicionColor = "text-sky-600 font-bold";
              suspicionMsg = `Nghi vấn ghi thiếu số hoặc sai đơn vị tính (Chỉ bằng ${Math.round(ratioToMean * 100)}% bình quân nhóm)`;
              labelShort = "Nghi thiếu số / sai đơn vị";
            } else {
              suspicionMsg = `Bất thường vượt quá ${sigmaThreshold} lần độ lệch chuẩn nhóm.`;
              labelShort = "Sai lệch phân phối";
            }

            detectedOutliers.push({
              ...item.row,
              _rowIdx: item.originalIdx,
              _rowNum: item.originalIdx + 1,
              _finalVal: item.finalVal,
              _rawVal: item.rawVal,
              _ratioVal: item.ratioVal,
              _group: item.group,
              _groupMean: mean,
              _groupStdDev: stdDev,
              _zScore: zScore,
              _zScoreAbs: absZ,
              _ratioToMean: ratioToMean,
              _isGroupStat: stats.useGroup,
              _suspicionMsg: suspicionMsg,
              _suspicionColor: suspicionColor,
              _labelShort: labelShort
            });
          }
        });

        detectedOutliers.sort((a, b) => b._zScoreAbs - a._zScoreAbs);

        setOutlierResults(detectedOutliers);
        setGroupStatsList(groupStatsArr.sort((a, b) => b.count - a.count));
        setOverallStats({
          totalScanned: validRows.length,
          outliersFound: detectedOutliers.length,
          outlierRate: validRows.length > 0 ? ((detectedOutliers.length / validRows.length) * 100).toFixed(2) + "%" : "0%",
          mean: overallMean,
          stdDev: overallStdDev,
          targetCol,
          groupCol,
          ratioCol
        });
      } catch (err) {
        console.error("Auto re-run anomalies failed:", err);
      }
    }
  }, [mainData, targetCol, groupCol, ratioCol, sigmaThreshold, minSampleCount, groupLevel, hasRun]);

  // Filter outlier results by text search
  const filteredOutliers = useMemo(() => {
    if (!searchFilter) return outlierResults;
    const query = searchFilter.toLowerCase();
    return outlierResults.filter(o => {
      return Object.values(o).some(val => String(val).toLowerCase().includes(query));
    });
  }, [outlierResults, searchFilter]);

  // Export Outlier report to Excel
  const handleExportOutliersExcel = () => {
    if (outlierResults.length === 0) return;
    
    // Prepare friendly data for Excel export
    const excelRows = outlierResults.map(o => {
      const excelItem: any = {};
      excelItem["Dòng số"] = o._rowNum;
      excelItem["Nhóm phân lớp"] = o._group;
      
      if (ratioCol) {
        excelItem[`Giá trị cột số (${targetCol})`] = o._rawVal;
        excelItem[`Giá trị chia cột tỉ lệ (${ratioCol})`] = o._ratioVal;
        excelItem["Tỷ lệ thực tế (Đơn vị tính)"] = o._finalVal;
      } else {
        excelItem[`Giá trị (${targetCol})`] = o._finalVal;
      }

      excelItem["Trung bình bình quân nhóm"] = Math.round(o._groupMean * 100) / 100;
      excelItem["Độ lệch chuẩn nhóm (Sigma)"] = Math.round(o._groupStdDev * 100) / 100;
      excelItem["Hệ số lệch Z-Score"] = Math.round(o._zScore * 100) / 100;
      excelItem["Tỷ lệ so với trung bình nhóm"] = Math.round(o._ratioToMean * 100) + "%";
      excelItem["Kết luận phân tích lỗi"] = o._suspicionMsg;

      // Add other original columns
      columns.forEach(col => {
        if (!col.startsWith("_") && col !== "Loi_Logic") {
          excelItem[col] = o[col];
        }
      });

      return excelItem;
    });

    if (onExportExcel) {
      onExportExcel(excelRows, `Báo_Cáo_Dị_Biệt_Phân_Phối_${fileNameClean()}`);
    } else {
      const worksheet = XLSX.utils.json_to_sheet(excelRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sai lệch phân phối");
      XLSX.writeFile(workbook, `Báo_Cáo_Dị_Biệt_Phân_Phối_${fileNameClean()}.xlsx`);
    }
  };

  const fileNameClean = () => {
    return new Date().toISOString().slice(0, 10).replace(/-/g, "");
  };

  // Convert outlier results into scatter plot format for recharts
  const scatterPlotData = useMemo(() => {
    if (outlierResults.length === 0 && mainData.length === 0) return [];
    
    // We sample a subset of normal records and mix with outliers to keep scatter chart fast
    const sampledNormal: any[] = [];
    const outlierIndices = new Set(outlierResults.map(o => o._rowIdx));
    const targetColClean = targetCol;
    const ratioColClean = ratioCol;

    let indexCount = 0;
    mainData.forEach((row, idx) => {
      if (outlierIndices.has(idx)) return; // Skip outliers, we plot them explicitly
      if (sampledNormal.length >= 120) return; // Keep chart size moderate for speed
      
      // Every Nth normal row
      if (idx % Math.max(1, Math.floor(mainData.length / 120)) !== 0) return;

      const rawTargetStr = String(row[targetColClean] || "").replace(/[^0-9.\-]/g, "");
      const rawTarget = parseFloat(rawTargetStr);
      if (isNaN(rawTarget)) return;

      let ratioNum = 1;
      if (ratioColClean) {
        const rawRatioStr = String(row[ratioColClean] || "").replace(/[^0-9.\-]/g, "");
        const parsedRatio = parseFloat(rawRatioStr);
        if (!isNaN(parsedRatio) && parsedRatio > 0) {
          ratioNum = parsedRatio;
        }
      }

      sampledNormal.push({
        x: idx + 1,
        y: ratioColClean ? rawTarget / ratioNum : rawTarget,
        isOutlier: false,
        name: `Dòng ${idx + 1}`
      });
    });

    const plottedOutliers = outlierResults.map(o => ({
      x: o._rowNum,
      y: o._finalVal,
      isOutlier: true,
      name: `Dòng ${o._rowNum} (LỆCH: ${Math.round(o._zScoreAbs * 10) / 10}σ)`
    }));

    return [...sampledNormal, ...plottedOutliers];
  }, [outlierResults, mainData, targetCol, ratioCol]);

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm font-sans">
      {/* Tab Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
        <div>
          <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
            <TrendingUp className="w-5 h-5 text-indigo-600 animate-pulse" /> KHỐI 2: BỘ QUÉT LỆCH QUY LUẬT PHÂN PHỐI (STATISTICAL OUTLIERS)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Tự động phân tích xu hướng quy luật của tệp tin vừa nạp. Nhận diện các đơn vị đột biến vượt ngưỡng phân phối (Z-Score &gt; {sigmaThreshold} Sigma) để phát hiện sai số gõ thừa/thiếu số 0 hoặc sai đơn vị tính.
          </p>
        </div>
        
        {hasRun && (
          <button
            onClick={() => setShowConfig(!showConfig)}
            className="flex items-center gap-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-lg border border-slate-200 transition-all cursor-pointer self-start"
          >
            <Sliders className="w-3.5 h-3.5" />
            {showConfig ? "Ẩn cấu hình rà quét" : "Hiện cấu hình rà quét"}
            {showConfig ? <ChevronUp className="w-3" /> : <ChevronDown className="w-3" />}
          </button>
        )}
      </div>

      {mainData.length > 0 ? (
        <div className="space-y-6">
          {/* Live 10 Row Preview inside StatisticalOutliers */}
          <LiveTenRowPreview
            data={mainData}
            columns={columns}
            onUpdateData={onUpdateMainData}
            highlightedIndices={outlierResults.map(item => item._rowIdx)}
            highlightLabel="Dòng dị biệt"
            title="BẢNG XEM NHANH & SỬA TRỰC TIẾP 10 DÒNG (RÀ SOÁT LỆCH PHÂN PHỐI)"
          />

          {/* CONFIG SECTION */}
          {showConfig && (
            <div className="bg-slate-50 rounded-2xl p-5 border border-slate-200 space-y-4 animate-fade-in">
              <h4 className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-slate-200">
                <Sliders className="w-4 h-4 text-indigo-500" /> THIẾT LẬP THAM SỐ PHÂN TÍCH PHÂN PHỐI DỮ LIỆU
              </h4>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {/* Chọn cột rà soát */}
                <div className="space-y-1">
                  <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
                    <span className="w-2 h-2 rounded-full bg-indigo-500"></span>
                    Cột số liệu cần quét lệch phân phối:
                  </label>
                  <select
                    value={targetCol}
                    onChange={(e) => setTargetCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-sm font-semibold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {columns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                  <p className="text-[10px] text-slate-500">Chọn cột Doanh thu, Sản lượng, Quy mô, Vốn...</p>
                </div>

                {/* Chọn cột phân nhóm */}
                <div className="space-y-1">
                  <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
                    <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                    Cột phân nhóm/ngành (Group By):
                  </label>
                  <select
                    value={groupCol}
                    onChange={(e) => setGroupCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-sm font-semibold"
                  >
                    <option value="">-- Toàn bộ tệp (Không phân nhóm) --</option>
                    {columns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                  <p className="text-[10px] text-slate-500">Giúp tính bình quân riêng cho từng mã ngành VSIC, Địa bàn...</p>
                </div>

                {/* Chọn cột chia tỉ lệ */}
                <div className="space-y-1">
                  <label className="text-xs text-slate-700 font-bold flex items-center gap-1.5">
                    <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                    Chia tỷ lệ cho (Normalize/Ratio) (Tùy chọn):
                  </label>
                  <select
                    value={ratioCol}
                    onChange={(e) => setRatioCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-sm font-semibold"
                  >
                    <option value="">-- Không chia (Phân tích giá trị trực tiếp) --</option>
                    {columns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                  <p className="text-[10px] text-slate-500">Ví dụ: Chọn cột [Lao động] để tính toán chỉ số [Doanh thu/lao động].</p>
                </div>
              </div>

              {/* Advanced Parameter Controls */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-2">
                <div className="space-y-1.5">
                  <div className="flex justify-between items-center text-xs">
                    <span className="text-slate-700 font-bold">Ngưỡng lọc lệch dị biệt (Sigma &sigma;):</span>
                    <span className="bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-md font-mono font-bold font-sans">
                      &gt;= {sigmaThreshold} Sigma
                    </span>
                  </div>
                  <input
                    type="range"
                    min="1.5"
                    max="5.0"
                    step="0.5"
                    value={sigmaThreshold}
                    onChange={(e) => setSigmaThreshold(parseFloat(e.target.value))}
                    className="w-full accent-indigo-600 cursor-pointer h-2 bg-slate-200 rounded-lg"
                  />
                  <p className="text-[10px] text-slate-400">
                    Theo quy tắc phân phối chuẩn: &gt; 3 Sigma là các giá trị cực kỳ đột biến (chiếm &lt; 0.3% tần suất xuất hiện tự nhiên).
                  </p>
                </div>

                <div className="space-y-1">
                  <label className="text-xs text-slate-700 font-bold">Quy tắc rút gọn nhóm ngành VSIC:</label>
                  <select
                    value={groupLevel}
                    onChange={(e) => setGroupLevel(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs text-slate-800 outline-none cursor-pointer"
                  >
                    <option value="all">Sử dụng nguyên bản chính xác trong cột</option>
                    <option value="2">Rút gọn lấy cấp 2 (Lấy 2 ký tự đầu mã VSIC)</option>
                    <option value="3">Rút gọn lấy cấp 3 (Lấy 3 ký tự đầu mã VSIC)</option>
                    <option value="4">Rút gọn lấy cấp 4 (Lấy 4 ký tự đầu mã VSIC)</option>
                  </select>
                  <p className="text-[10px] text-slate-500">Bình quân theo nhóm ngành rộng giúp mẫu so sánh dày dặn và có nghĩa hơn.</p>
                </div>

                <div className="space-y-1">
                  <label className="text-xs text-slate-700 font-bold">Số mẫu tối thiểu để lập nhóm bình quân:</label>
                  <input
                    type="number"
                    min="2"
                    max="50"
                    value={minSampleCount}
                    onChange={(e) => setMinSampleCount(parseInt(e.target.value, 10) || 3)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800"
                  />
                  <p className="text-[10px] text-slate-500">Nếu nhóm có ít hơn số mẫu này, hệ thống sẽ sử dụng số liệu bình quân toàn bộ tệp.</p>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex justify-end">
                <button
                  onClick={handleAnalyze}
                  disabled={isAnalyzing}
                  className="bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-bold text-xs py-2.5 px-6 rounded-xl cursor-pointer transition-all shadow-md active:scale-95 flex items-center gap-1.5 border-0 disabled:opacity-50"
                >
                  <Zap className="w-4 h-4 text-amber-300" />
                  {isAnalyzing ? "Đang chạy tính toán xu hướng & kiểm tra..." : "⚡ CHẠY QUÉT DỊ BIỆT PHÂN PHỐI"}
                </button>
              </div>
            </div>
          )}

          {/* RESULTS DASHBOARD */}
          {hasRun && (
            <div className="space-y-6 animate-fade-in">
              {/* Metric Cards Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-slate-50 border border-slate-200 p-4 rounded-2xl text-center space-y-1 shadow-2xs">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">🔍 Tổng số bản ghi quét</span>
                  <strong className="text-xl text-slate-800 font-mono font-bold">
                    {overallStats.totalScanned.toLocaleString()}
                  </strong>
                  <span className="text-[10px] text-slate-400 block">Dòng dữ liệu số hợp lệ</span>
                </div>

                <div className={`border p-4 rounded-2xl text-center space-y-1 shadow-2xs transition-all ${
                  overallStats.outliersFound > 0 ? "bg-amber-50/70 border-amber-200 text-amber-800" : "bg-emerald-50 border-emerald-200 text-emerald-800"
                }`}>
                  <span className="text-[10px] font-bold uppercase tracking-wider block">⚠️ Số dị biệt phát hiện</span>
                  <strong className="text-xl font-mono font-bold">
                    {overallStats.outliersFound.toLocaleString()}
                  </strong>
                  <span className="text-[10px] block opacity-90">Tỷ lệ: {overallStats.outlierRate}</span>
                </div>

                <div className="bg-slate-50 border border-slate-200 p-4 rounded-2xl text-center space-y-1 shadow-2xs">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">📊 Bình quân toàn tệp (Mean)</span>
                  <strong className="text-xl text-indigo-700 font-mono font-bold">
                    {Math.round(overallStats.mean * 100) / 100}
                  </strong>
                  <span className="text-[10px] text-slate-400 block truncate" title={ratioCol ? `${targetCol}/${ratioCol}` : targetCol}>
                    {ratioCol ? "Tỉ giá đơn vị bình quân" : "Bình quân số lượng trực tiếp"}
                  </span>
                </div>

                <div className="bg-slate-50 border border-slate-200 p-4 rounded-2xl text-center space-y-1 shadow-2xs">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">📈 Độ lệch tiêu chuẩn (Sigma)</span>
                  <strong className="text-xl text-slate-700 font-mono font-bold">
                    {Math.round(overallStats.stdDev * 100) / 100}
                  </strong>
                  <span className="text-[10px] text-slate-400 block">Mức độ biến thiên sai lệch</span>
                </div>
              </div>

              {/* Visual Scatter Chart & Statistical Explanatory */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Visual Chart Card */}
                <div className="lg:col-span-2 bg-white border border-slate-200 rounded-2xl p-4.5 shadow-sm space-y-3">
                  <div className="flex items-center justify-between">
                    <h5 className="text-[12px] font-extrabold text-slate-700 flex items-center gap-1.5 uppercase font-mono">
                      <BarChart3 className="w-4 h-4 text-indigo-600" /> Biểu đồ phân tán & trực quan hóa dị biệt
                    </h5>
                    <span className="text-[10px] bg-indigo-50 text-indigo-700 border border-indigo-200 px-2 py-0.5 rounded font-mono font-bold">
                      Trục Y: {ratioCol ? "Tỷ lệ đơn vị" : "Giá trị"}
                    </span>
                  </div>

                  {/* Scatter Plot */}
                  <div className="h-64 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ScatterChart margin={{ top: 10, right: 10, bottom: 10, left: 10 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                        <XAxis type="number" dataKey="x" name="Dòng số" stroke="#94a3b8" fontSize={10} />
                        <YAxis type="number" dataKey="y" name="Giá trị" stroke="#94a3b8" fontSize={10} />
                        <Tooltip 
                          cursor={{ strokeDasharray: '3 3' }} 
                          content={({ active, payload }) => {
                            if (active && payload && payload.length) {
                              const data = payload[0].payload;
                              return (
                                <div className="bg-slate-800 text-white p-2.5 rounded-lg text-[11px] font-sans space-y-1 shadow-md border border-slate-700">
                                  <p className="font-bold border-b border-slate-600 pb-1">{data.name}</p>
                                  <p><span className="text-slate-300">Giá trị/Tỉ giá:</span> <strong className="font-mono text-amber-400">{Math.round(data.y * 100) / 100}</strong></p>
                                  <p><span className="text-slate-300">Trạng thái:</span> <span className={data.isOutlier ? "text-red-400 font-bold" : "text-emerald-400"}>{data.isOutlier ? "⚠️ Bất thường" : "✅ Bình thường"}</span></p>
                                </div>
                              );
                            }
                            return null;
                          }}
                        />
                        <Scatter name="Phân phối dữ liệu" data={scatterPlotData}>
                          {scatterPlotData.map((entry: any, index: number) => (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={entry.isOutlier ? "#f59e0b" : "#4f46e5"} 
                              fillOpacity={entry.isOutlier ? 0.95 : 0.35}
                              r={entry.isOutlier ? 6 : 4} 
                            />
                          ))}
                        </Scatter>
                        {/* Reference lines */}
                        <ReferenceLine y={overallStats.mean} stroke="#10b981" strokeDasharray="5 5">
                          <Label value="Bình quân toàn tệp" position="insideBottomRight" fill="#10b981" fontSize={9} />
                        </ReferenceLine>
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Explanatory Statistical Table */}
                <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4.5 shadow-sm space-y-3 flex flex-col justify-between">
                  <div className="space-y-3">
                    <h5 className="text-[12px] font-extrabold text-slate-700 flex items-center gap-1.5 uppercase font-mono">
                      <Info className="w-4 h-4 text-emerald-600" /> Bản tin khoa học dữ liệu gốc
                    </h5>
                    
                    <div className="text-xs text-slate-600 space-y-2.5 leading-relaxed font-sans">
                      <p>
                        Hệ thống đã phân tích <strong>{overallStats.totalScanned.toLocaleString()}</strong> bản ghi. 
                        Phát hiện thấy <strong>{overallStats.outliersFound}</strong> phần tử lệch lệch mạnh so với trục phân phối quy chuẩn.
                      </p>
                      
                      <div className="bg-white border border-slate-200 p-3 rounded-xl space-y-1.5">
                        <span className="text-[10px] uppercase font-bold text-slate-500 block">💡 Cơ sở toán học cảnh báo:</span>
                        <div className="text-[10.5px] space-y-1 font-mono text-indigo-700">
                          <div>• Điểm số Z-Score = |x - &mu;| / &sigma;</div>
                          <div>• &mu; (Mean): Trung bình nhóm nghề</div>
                          <div>• &sigma; (Sigma): Khoảng lệch biến thiên</div>
                          <div>• Tiêu chuẩn: Nếu Z-Score &gt;= {sigmaThreshold}, lập tức cảnh báo.</div>
                        </div>
                      </div>

                      <p className="text-[11px] italic text-slate-500">
                        * Mẹo giám sát: Các dị biệt có đuôi số tròn chục (gấp 10 lần, 100 lần trung bình nhóm) thường là do điều tra viên gõ thừa số 0 khi nhập liệu hoặc tính sai đơn vị (ví dụ nghìn đồng vs triệu đồng).
                      </p>
                    </div>
                  </div>

                  <div className="pt-3 border-t border-slate-200">
                    <button
                      onClick={handleExportOutliersExcel}
                      className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-2 px-3 rounded-xl transition-all cursor-pointer flex items-center justify-center gap-1.5 border-0 shadow-sm active:scale-95"
                    >
                      <Download className="w-3.5 h-3.5" /> Xuất danh sách dị biệt (.xlsx)
                    </button>
                  </div>
                </div>
              </div>

              {/* LIST OF DETECTED OUTLIERS */}
              <div className="border border-slate-200 rounded-2xl overflow-hidden shadow-sm bg-white space-y-3 p-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                  <div>
                    <h5 className="text-[12px] font-extrabold text-slate-800 uppercase tracking-wider font-mono flex items-center gap-1.5">
                      🚨 DANH SÁCH {filteredOutliers.length} HỘ/BẢN GHI DỊ BIỆT BẤT THƯỜNG
                    </h5>
                    <p className="text-[10.5px] text-slate-500 font-sans">
                      Thỏa mãn tiêu chí lệch vượt quá {sigmaThreshold} lần độ lệch tiêu chuẩn (&sigma;). Click vào dòng hoặc lọc nhanh để xem.
                    </p>
                  </div>

                  {/* Search Bar for Outliers list */}
                  <div className="relative max-w-xs">
                    <input
                      type="text"
                      placeholder="Tìm nhanh dòng dị biệt..."
                      value={searchFilter}
                      onChange={(e) => setSearchFilter(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-250 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-800 outline-none focus:bg-white focus:ring-1 focus:ring-indigo-500 transition-all placeholder-slate-400"
                    />
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  </div>
                </div>

                {filteredOutliers.length === 0 ? (
                  <div className="p-8 text-center text-slate-400 italic bg-slate-50 rounded-xl border border-dashed border-slate-250 text-xs">
                    Không tìm thấy dòng dị biệt nào khớp với nội dung tìm kiếm.
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                          <th className="py-2.5 px-3 font-bold text-center w-12">STT</th>
                          <th className="py-2.5 px-3 font-bold w-16">Dòng số</th>
                          <th className="py-2.5 px-3 font-bold w-36">Phân nhóm</th>
                          <th className="py-2.5 px-3 font-bold w-36 text-right">Giá trị quét</th>
                          {ratioCol && <th className="py-2.5 px-3 font-bold w-32 text-right">Hệ số chuẩn hóa</th>}
                          <th className="py-2.5 px-3 font-bold w-24 text-right">Z-Score</th>
                          <th className="py-2.5 px-3 font-bold w-24 text-right">Gấp TB nhóm</th>
                          <th className="py-2.5 px-3 font-bold">Chẩn đoán thông minh sai số</th>
                          <th className="py-2.5 px-3 font-bold text-center w-24">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredOutliers.slice(0, 100).map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50/70 transition-colors">
                            <td className="py-2 px-3 text-center text-slate-400 font-mono">{idx + 1}</td>
                            <td className="py-2 px-3 font-bold text-indigo-700 font-mono">#{item._rowNum}</td>
                            <td className="py-2 px-3 truncate max-w-[150px]" title={item._group}>
                              <span className="bg-slate-100 text-slate-800 border border-slate-200 px-1.5 py-0.5 rounded font-mono text-[10px] font-bold">
                                {item._group}
                              </span>
                            </td>
                            <td className="py-2 px-3 font-mono font-bold text-right text-slate-900">
                              {item._finalVal.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                            </td>
                            {ratioCol && (
                              <td className="py-2 px-3 font-mono text-right text-slate-500">
                                {item._rawVal.toLocaleString()} / {item._ratioVal.toLocaleString()}
                              </td>
                            )}
                            <td className="py-2 px-3 text-right font-mono font-bold text-amber-600">
                              +{item._zScoreAbs.toFixed(1)} &sigma;
                            </td>
                            <td className="py-2 px-3 text-right font-mono text-indigo-700">
                              {item._ratioToMean >= 1 
                                ? `${item._ratioToMean.toFixed(1)}x` 
                                : `x1/${(1/item._ratioToMean).toFixed(1)}`}
                            </td>
                            <td className="py-2 px-3">
                              <span className={`text-[11px] font-medium flex items-center gap-1 ${item._suspicionColor}`}>
                                <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                                {item._suspicionMsg}
                              </span>
                            </td>
                            <td className="py-2 px-3 text-center">
                              {onFilterRows && (
                                <button
                                  onClick={() => onFilterRows([item._rowIdx], `Dòng dị biệt phân phối #${item._rowNum}`)}
                                  className="bg-indigo-50 hover:bg-indigo-600 hover:text-white text-indigo-700 px-2 py-1 rounded text-[10px] font-bold transition-all cursor-pointer border border-indigo-200"
                                >
                                  🔍 Xem chi tiết
                                </button>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    
                    {filteredOutliers.length > 100 && (
                      <div className="p-3 bg-slate-50 text-center text-[11px] text-slate-500 border-t border-slate-100">
                        Đang hiển thị 100 dòng dị biệt đầu tiên. Nhấp vào "Xuất danh sách dị biệt" để xem toàn bộ file báo cáo.
                      </div>
                    )}
                  </div>
                )}
                
                {onFilterRows && (
                  <div className="flex justify-end pt-2">
                    <button
                      onClick={() => {
                        const indices = filteredOutliers.map(o => o._rowIdx);
                        onFilterRows(indices, `Tất cả ${indices.length} dòng dị biệt bất thường phân phối`);
                      }}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-4 py-2 rounded-xl border-0 transition-all flex items-center gap-1 cursor-pointer active:scale-95 shadow-sm"
                    >
                      👁️ LỌC TẬP TRUNG TOÀN BỘ TRÊN BẢNG CHÍNH HỆ THỐNG
                    </button>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="p-10 text-center text-slate-400 bg-slate-50 border border-dashed border-slate-300 rounded-2xl font-sans">
          <Info className="w-8 h-8 text-slate-300 mx-auto mb-2" />
          Chưa có dữ liệu nguồn được nạp. Hãy nạp file Excel/CSV ở trang chủ để kích hoạt chức năng rà quét.
        </div>
      )}
    </div>
  );
}
