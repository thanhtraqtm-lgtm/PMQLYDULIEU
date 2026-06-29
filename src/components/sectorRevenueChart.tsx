import React, { useState, useMemo, useEffect } from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  Cell,
  LabelList
} from "recharts";
import { 
  normalizeSectorCode, 
  getParentSectorCode, 
  vsicRawData,
  isSummaryRow
} from "../data/vsic";
import { BarChart3, AlertCircle, CheckCircle2, ChevronRight, HelpCircle } from "lucide-react";

interface SectorRevenueChartProps {
  mainData: any[];
  columns: string[];
  mapping: {
    mota: string;
    manganh: string;
    xa: string;
    doanhthu: string;
    laodong: string;
    idCol: string;
  };
  reportLevel?: number; // 0: Direct raw groupings, 1: VSIC level 1, 2: VSIC level 2
}

interface Level1Summary {
  code: string;
  name: string;
  revenue: number;
  communeCount: number;
  recordCount: number;
}

export function parseRobustNumber(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return val;
  let str = String(val).trim();
  if (!str) return 0;

  str = str.replace(/\s+/g, "").replace(/[đđVNDvnd]/g, "");

  const hasComma = str.includes(",");
  const hasDot = str.includes(".");

  if (hasComma && hasDot) {
    if (str.indexOf(".") < str.indexOf(",")) {
      str = str.replace(/\./g, "").replace(/,/g, ".");
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (hasComma) {
    const commas = (str.match(/,/g) || []).length;
    if (commas > 1) {
      str = str.replace(/,/g, "");
    } else {
      const parts = str.split(",");
      if (parts[1].length === 3) {
        str = str.replace(/,/g, "");
      } else {
        str = str.replace(/,/g, ".");
      }
    }
  } else if (hasDot) {
    const dots = (str.match(/\./g) || []).length;
    if (dots > 1) {
      str = str.replace(/\./g, "");
    } else {
      const parts = str.split(".");
      if (parts[1].length === 3) {
        str = str.replace(/\./g, "");
      }
    }
  }

  const clean = str.replace(/[^0-9.\-]/g, "");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

const SectorRevenueChart = React.memo(function SectorRevenueChart({ mainData, columns, mapping, reportLevel = 1 }: SectorRevenueChartProps) {
  const [selectedManganh, setSelectedManganh] = useState<string>(mapping.manganh || "");
  const [selectedDoanhthu, setSelectedDoanhthu] = useState<string>(mapping.doanhthu || "");
  const [scale, setScale] = useState<"auto" | "raw" | "million" | "billion">("auto");
  const [layout, setLayout] = useState<"horizontal" | "vertical">("horizontal");
  const [showLabels, setShowLabels] = useState<boolean>(true);
  const [colorTheme, setColorTheme] = useState<"violet" | "indigo" | "emerald" | "amber">("indigo");

  useEffect(() => {
    setSelectedManganh(mapping.manganh || "");
    setSelectedDoanhthu(mapping.doanhthu || "");
  }, [mapping.manganh, mapping.doanhthu]);

  const aggregatedData = useMemo(() => {
    if (mainData.length === 0 || !selectedManganh) return [];

    const map: { [code: string]: { revenue: number; records: number; communes: Set<string> } } = {};

    mainData.forEach(row => {
      if (isSummaryRow(row)) return;

      const rawMng = row[selectedManganh];
      let groupKey = "";
      let groupName = "";
      
      if (reportLevel === 0) {
        groupKey = String(rawMng || "Chưa xác định / Bỏ trống").trim();
        groupName = groupKey;
      } else {
        const mng = normalizeSectorCode(rawMng);
        if (reportLevel === 2) {
          groupKey = mng ? mng.slice(0, 2) : "";
          groupName = groupKey ? `${groupKey} - ${vsicRawData[groupKey] || "Ngành cấp 2 chưa định nghĩa"}` : "Chưa xác định";
        } else {
          let l1Code = "";
          if (mng) {
            if (/^[a-zA-Z]$/.test(mng)) {
              l1Code = mng.toUpperCase();
            } else {
              l1Code = getParentSectorCode(mng) || "";
            }
          }
          groupKey = l1Code || "KHAC";
          groupName = vsicRawData[groupKey] || vsicRawData[groupKey.toUpperCase()] || (groupKey === "KHAC" ? "Mã ngành chưa khớp VSIC / Khác" : `Ngành cấp 1 chưa định nghĩa (${groupKey})`);
        }
      }

      if (!groupKey) {
        groupKey = "CHUA_XAC_DINH";
        groupName = "Chưa xác định / Bỏ trống";
      }

      let valNum = 0;
      if (selectedDoanhthu && row[selectedDoanhthu] !== undefined) {
        valNum = parseRobustNumber(row[selectedDoanhthu]);
      }

      const communeVal = mapping.xa && row[mapping.xa] ? String(row[mapping.xa]).trim() : "Khác";

      if (!map[groupKey]) {
        map[groupKey] = { revenue: 0, records: 0, communes: new Set() };
      }
      map[groupKey].revenue += valNum;
      map[groupKey].records += 1;
      if (communeVal) {
        map[groupKey].communes.add(communeVal);
      }
    });

    const results: Level1Summary[] = Object.keys(map).map(code => {
      let name = code;
      if (reportLevel === 2) {
        name = code ? `${code} - ${vsicRawData[code] || "Ngành cấp 2 chưa định nghĩa"}` : "Chưa xác định";
      } else if (reportLevel === 1) {
        name = vsicRawData[code] || vsicRawData[code.toUpperCase()] || (code === "KHAC" ? "Mã ngành chưa khớp VSIC / Khác" : `Ngành cấp 1 chưa định nghĩa (${code})`);
      } else {
        name = code;
      }
      return {
        code,
        name,
        revenue: Math.round(map[code].revenue * 100) / 100,
        communeCount: map[code].communes.size,
        recordCount: map[code].records
      };
    });

    return results.sort((a, b) => b.revenue - a.revenue);
  }, [mainData, selectedManganh, selectedDoanhthu, mapping.xa, reportLevel]);

  const allSectors = useMemo(() => {
    return aggregatedData.map(item => ({
      code: item.code,
      name: item.name,
      revenue: item.revenue
    }));
  }, [aggregatedData]);

  const [selectedSectors, setSelectedSectors] = useState<string[]>([]);
  const [hasInitializedSectors, setHasInitializedSectors] = useState<string>("");

  const currentDatasetId = `${mainData.length}_${selectedManganh}_${selectedDoanhthu}_${reportLevel}`;
  useEffect(() => {
    if (allSectors.length > 0 && hasInitializedSectors !== currentDatasetId) {
      setSelectedSectors(allSectors.map(s => s.code));
      setHasInitializedSectors(currentDatasetId);
    }
  }, [allSectors, hasInitializedSectors, currentDatasetId]);

  const handleToggleSector = (code: string) => {
    if (selectedSectors.includes(code)) {
      setSelectedSectors(selectedSectors.filter(c => c !== code));
    } else {
      setSelectedSectors([...selectedSectors, code]);
    }
  };

  const handleSelectAll = () => {
    setSelectedSectors(allSectors.map(s => s.code));
  };

  const handleDeselectAll = () => {
    setSelectedSectors([]);
  };

  const activeSectorsData = useMemo(() => {
    return aggregatedData.filter(item => selectedSectors.includes(item.code));
  }, [aggregatedData, selectedSectors]);

  const autoDetectedScale = useMemo(() => {
    if (activeSectorsData.length === 0) return { factor: 1, unit: "" };
    const maxVal = Math.max(...activeSectorsData.map(item => item.revenue));
    
    if (maxVal >= 1000000000) {
      return { factor: 1000000000, unit: "Tỷ" };
    } else if (maxVal >= 1000000) {
      return { factor: 1000000, unit: "Triệu" };
    } else if (maxVal >= 1000) {
      return { factor: 1000, unit: "Nghìn" };
    }
    return { factor: 1, unit: "" };
  }, [activeSectorsData]);

  const currentScaleInfo = useMemo(() => {
    if (scale === "raw") return { factor: 1, unit: "" };
    if (scale === "million") return { factor: 1000000, unit: "Triệu" };
    if (scale === "billion") return { factor: 1000000000, unit: "Tỷ" };
    return autoDetectedScale;
  }, [scale, autoDetectedScale]);

  const chartData = useMemo(() => {
    return activeSectorsData.map(item => {
      const scaledRev = item.revenue / currentScaleInfo.factor;
      return {
        ...item,
        scaledRevenue: Math.round(scaledRev * 100) / 100,
        codeLabel: item.code.length > 25 ? `${item.code.substring(0, 22)}...` : item.code,
        fullName: item.name,
        rawRevenueFormatted: new Intl.NumberFormat("vi-VN").format(item.revenue)
      };
    }).sort((a, b) => b.scaledRevenue - a.scaledRevenue);
  }, [activeSectorsData, currentScaleInfo]);

  const colors = {
    indigo: { bar: "#4f46e5", barHover: "#4338ca", bgHover: "#f5f3ff" },
    violet: { bar: "#7c3aed", barHover: "#6d28d9", bgHover: "#f5f3ff" },
    emerald: { bar: "#059669", barHover: "#047857", bgHover: "#ecfdf5" },
    amber: { bar: "#d97706", barHover: "#b45309", bgHover: "#fffbeb" }
  };

  const selectedColor = colors[colorTheme];

  const totalRevenueSelected = useMemo(() => {
    const total = activeSectorsData.reduce((sum, item) => sum + item.revenue, 0);
    return new Intl.NumberFormat("vi-VN").format(total);
  }, [activeSectorsData]);

  const maxRevenueSector = useMemo(() => {
    if (activeSectorsData.length === 0) return null;
    let maxItem = activeSectorsData[0];
    activeSectorsData.forEach(item => {
      if (item.revenue > maxItem.revenue) {
        maxItem = item;
      }
    });
    return maxItem;
  }, [activeSectorsData]);

  const chartTitle = useMemo(() => {
    if (reportLevel === 0) return "BIỂU ĐỒ TRỰC QUAN THEO NHÓM THỜI GIAN/ĐẶC TÍNH";
    if (reportLevel === 2) return "BIỂU ĐỒ TRỰC QUAN THEO NGÀNH CẤP 2 (VSIC 2018)";
    return "BIỂU ĐỒ TRỰC QUAN THEO NGÀNH CẤP 1 (VSIC 2018)";
  }, [reportLevel]);

  const chartDesc = useMemo(() => {
    if (reportLevel === 0) return `Phân tích gộp và so sánh chỉ tiêu [${selectedDoanhthu}] trực tiếp của các nhóm phân loại chính trong dữ liệu.`;
    return `Tổng hợp và so sánh tự động đại lượng [${selectedDoanhthu}] quy nạp lên chuẩn ngành kinh tế tương ứng quốc gia.`;
  }, [reportLevel, selectedDoanhthu]);

  return (
    <div className="space-y-6 animate-fade-in font-sans text-slate-800">
      {/* Configuration Header Panel */}
      <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4 shadow-xl">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <BarChart3 className="w-5 h-5 text-indigo-600" /> {chartTitle}
            </h3>
            <p className="text-[11px] text-slate-500 font-sans mt-0.5">
              {chartDesc}
            </p>
          </div>

          <div className="flex flex-wrap gap-2">
            {/* Column selectors */}
            <div className="flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
              <label className="text-[10px] uppercase font-bold text-slate-500 font-mono">
                {reportLevel === 0 ? "Cột Nhóm:" : "Cột Ngành:"}
              </label>
              <select
                value={selectedManganh}
                onChange={(e) => setSelectedManganh(e.target.value)}
                className="bg-transparent text-xs text-indigo-700 font-bold border-none outline-none focus:ring-0 max-w-[130px] cursor-pointer"
              >
                <option value="">-- Chọn cột --</option>
                {columns.map(col => (
                  <option key={col} value={col} className="bg-white text-slate-800">{col}</option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
              <label className="text-[10px] uppercase font-bold text-slate-500 font-mono">Cột Chỉ tiêu số:</label>
              <select
                disabled={!columns.length}
                value={selectedDoanhthu}
                onChange={(e) => setSelectedDoanhthu(e.target.value)}
                className="bg-transparent text-xs text-emerald-700 font-bold border-none outline-none focus:ring-0 max-w-[130px] cursor-pointer"
              >
                <option value="">-- Chọn cột --</option>
                {columns.map(col => (
                  <option key={col} value={col} className="bg-white text-slate-800">{col}</option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {(!selectedManganh || !selectedDoanhthu) && (
          <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex gap-3 text-xs text-amber-800">
            <AlertCircle className="w-5 h-5 shrink-0 text-amber-600" />
            <div>
              <p className="font-semibold">Chưa khớp hoặc thiếu cấu hình mapping chỉ tiêu số!</p>
              <p className="mt-1 text-[11px] text-slate-500">
                Hãy lựa chọn Cột phân nhóm và Cột giá trị số cộng gộp thích hợp ở biểu đồ để kích hoạt vẽ trực quan tức thì.
              </p>
            </div>
          </div>
        )}
      </div>

      {mainData.length > 0 && selectedManganh && selectedDoanhthu ? (
        <div className="grid grid-cols-1 xl:grid-cols-4 gap-6">
          
          {/* LEFT: Checklist filter sidebar */}
          <div className="xl:col-span-1 bg-white border border-slate-200 rounded-2xl p-4 flex flex-col space-y-4 max-h-[640px] overflow-hidden shadow-md">
            <div>
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider font-mono">
                  {reportLevel === 0 ? "BỘ LỌC CÁC NHÓM" : "BỘ LỌC NGÀNH KINH TẾ"} ({selectedSectors.length}/{allSectors.length})
                </span>
              </div>
              <p className="text-[10px] text-slate-500 leading-relaxed font-sans">
                Tích chọn các hạng mục muốn hiển thị so sánh trực tiếp trên đồ thị cột:
              </p>
            </div>

            {/* Quick multi-select buttons */}
            <div className="flex gap-2 shrink-0">
              <button
                onClick={handleSelectAll}
                className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[10px] py-1.5 rounded-lg border border-slate-200 transition-all cursor-pointer"
              >
                Chọn tất cả
              </button>
              <button
                onClick={handleDeselectAll}
                className="flex-1 bg-white hover:bg-slate-50 text-slate-500 font-bold text-[10px] py-1.5 rounded-lg border border-slate-200 transition-all cursor-pointer"
              >
                Bỏ chọn hết
              </button>
            </div>

            {/* Sector list display inside a scrollbar */}
            <div className="flex-1 overflow-y-auto space-y-1.5 pr-1 text-xs">
              {allSectors.length === 0 ? (
                <div className="text-center italic text-slate-400 py-8 font-sans">
                  Không quét được dữ liệu giá trị nào hợp lệ trong cột đã chọn.
                </div>
              ) : (
                allSectors.map(sector => {
                  const isChecked = selectedSectors.includes(sector.code);
                  const formattedRev = new Intl.NumberFormat("vi-VN").format(sector.revenue);
                  return (
                    <label
                      key={sector.code}
                      className={`flex items-start gap-2.5 p-2 rounded-xl border transition-all cursor-pointer ${
                        isChecked
                          ? "bg-indigo-50/50 border-indigo-200 text-indigo-900"
                          : "bg-transparent border-transparent text-slate-500 hover:bg-slate-50"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => handleToggleSector(sector.code)}
                        className="mt-0.5 rounded text-indigo-600 focus:ring-indigo-500 bg-white border-slate-300 cursor-pointer"
                      />
                      <div className="flex-1 min-w-0 font-sans">
                        <div className="font-bold flex items-center gap-1.5 shrink-0 text-slate-800">
                          {sector.code.length > 25 ? `${sector.code.substring(0,22)}...` : sector.code}
                        </div>
                        <div className="text-[10px] text-slate-500 mt-0.5 line-clamp-2 truncate-2-lines" title={sector.name}>
                          {sector.name}
                        </div>
                        <div className="text-[10px] text-emerald-700 font-mono font-bold mt-0.5">
                          Giá trị: {formattedRev}
                        </div>
                      </div>
                    </label>
                  );
                })
              )}
            </div>
          </div>

          {/* RIGHT: Actual chart visualization canvas */}
          <div className="xl:col-span-3 bg-white border border-slate-200 rounded-2xl p-5 flex flex-col justify-between space-y-4 shadow-xl relative min-h-[500px]">
            
            {/* Control panel for graphs */}
            <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div className="flex flex-wrap items-center gap-4 text-xs">
                {/* Scale selection */}
                <div className="flex items-center gap-1.5">
                  <span className="text-[10.5px] font-bold text-slate-500 font-mono">ĐƠN VỊ TỶ LỆ:</span>
                  <div className="bg-slate-100 border border-slate-200 rounded-lg p-0.5 flex gap-0.5">
                    {(["auto", "raw", "million", "billion"] as const).map(u => (
                      <button
                        key={u}
                        onClick={() => setScale(u)}
                        className={`px-2.5 py-1 text-[10px] font-bold rounded-lg transition-all cursor-pointer ${
                          scale === u
                            ? "bg-indigo-600 text-white shadow-sm"
                            : "text-slate-600 hover:text-slate-800 hover:bg-slate-200"
                        }`}
                      >
                        {u === "auto" ? "Tự động" : u === "raw" ? "Nguyên bản" : u === "million" ? "Triệu đ" : "Tỷ đ"}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Orientation switch */}
                <div className="flex items-center gap-1.5">
                  <span className="text-[10.5px] font-bold text-slate-500 font-mono">HƯỚNG BIỂU ĐỒ:</span>
                  <div className="bg-slate-100 border border-slate-200 rounded-lg p-0.5 flex gap-0.5">
                    <button
                      onClick={() => setLayout("horizontal")}
                      className={`px-2.5 py-1 text-[10px] font-bold rounded-lg transition-all cursor-pointer ${
                        layout === "horizontal" ? "bg-indigo-600 text-white shadow-sm" : "text-slate-600 hover:text-slate-800 hover:bg-slate-200"
                      }`}
                    >
                      Xếp đứng
                    </button>
                    <button
                      onClick={() => setLayout("vertical")}
                      className={`px-2.5 py-1 text-[10px] font-bold rounded-lg transition-all cursor-pointer ${
                        layout === "vertical" ? "bg-indigo-600 text-white shadow-sm" : "text-slate-600 hover:text-slate-800 hover:bg-slate-200"
                      }`}
                    >
                      Xếp ngang
                    </button>
                  </div>
                </div>
              </div>

              {/* Extras layout */}
              <div className="flex items-center gap-3">
                <label className="flex items-center gap-1.5 text-xs text-slate-650 hover:text-slate-800 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={showLabels}
                    onChange={(e) => setShowLabels(e.target.checked)}
                    className="rounded text-indigo-600 bg-white border-slate-300"
                  />
                  <span>Hiện số liệu dòng cột</span>
                </label>

                {/* Color Scheme selector */}
                <div className="flex items-center gap-1">
                  {(["indigo", "violet", "emerald", "amber"] as const).map(colorThemeName => (
                    <button
                      key={colorThemeName}
                      onClick={() => setColorTheme(colorThemeName)}
                      className={`w-3.5 h-3.5 rounded-full transition-all border cursor-pointer ${
                        colorTheme === colorThemeName
                          ? "border-slate-800 scale-125 ring-2 ring-indigo-600/20"
                          : "border-transparent opacity-60 hover:opacity-100"
                      }`}
                      style={{
                        backgroundColor: 
                          colorThemeName === "indigo" ? "#4f46e5" :
                          colorThemeName === "violet" ? "#7c3aed" :
                          colorThemeName === "emerald" ? "#059669" : "#d97706"
                      }}
                      title={`Tông màu ${colorThemeName}`}
                    />
                  ))}
                </div>
              </div>
            </div>

            {/* Recharts Render Stage */}
            <div className="flex-1 min-h-[360px] flex items-center justify-center bg-slate-50/50 rounded-xl border border-slate-100 p-2">
              {chartData.length === 0 ? (
                <div className="text-slate-400 italic text-xs py-16 text-center font-sans">
                  ⚠️ Không có hạng mục dư liệu phân loại nào đang được chọn hiển thị. Vui lòng tick chọn hộp bên trái.
                </div>
              ) : (
                <div className="w-full h-[400px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      layout={layout}
                      data={chartData}
                      margin={layout === "horizontal" 
                        ? { top: 20, right: 10, left: 10, bottom: 25 }
                        : { top: 10, right: 30, left: 40, bottom: 10 }
                      }
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" opacity={0.6} />
                      
                      {layout === "horizontal" ? (
                        <>
                          <XAxis 
                            dataKey="codeLabel" 
                            stroke="#475569" 
                            fontSize={10} 
                            tickLine={false} 
                            dy={10}
                            fontFamily="monospace"
                          />
                          <YAxis 
                            stroke="#475569" 
                            fontSize={10} 
                            tickLine={false}
                            dx={-5}
                            fontFamily="monospace"
                            unit={currentScaleInfo.unit}
                          />
                        </>
                      ) : (
                        <>
                          <XAxis 
                            type="number" 
                            stroke="#475569" 
                            fontSize={10} 
                            tickLine={false}
                            fontFamily="monospace"
                            unit={currentScaleInfo.unit}
                          />
                          <YAxis 
                            type="category"
                            dataKey="codeLabel" 
                            stroke="#475569" 
                            fontSize={10} 
                            tickLine={false}
                            dx={-5}
                            fontFamily="monospace"
                          />
                        </>
                      )}

                      <Tooltip
                        cursor={{ fill: "#f1f5f9", opacity: 0.4 }}
                        content={({ active, payload }) => {
                          if (active && payload && payload.length) {
                            const data = payload[0].payload;
                            return (
                              <div className="bg-white border border-slate-250 rounded-xl p-3 shadow-xl space-y-1.5 text-xs max-w-sm font-sans z-50 text-slate-800">
                                <div className="font-bold text-slate-900 flex items-center gap-1.5">
                                  <span className="bg-indigo-50 border border-indigo-200 text-indigo-700 font-mono text-[9px] px-1.5 py-0.5 rounded">
                                    {data.code}
                                  </span>
                                  {data.fullName}
                                </div>
                                <div className="text-[10px] text-slate-600">
                                  Giá trị thực tế: <strong className="text-emerald-700 font-mono">{data.rawRevenueFormatted} đ</strong>
                                </div>
                                <div className="text-[10px] text-slate-600">
                                  Số lượng nguồn: <span className="font-semibold text-purple-700 font-mono">{data.recordCount} dòng</span> / <span className="font-semibold text-blue-700 font-mono">{data.communeCount} địa bàn xã</span>
                                </div>
                              </div>
                            );
                          }
                          return null;
                        }}
                      />

                      <Bar 
                        dataKey="scaledRevenue" 
                        fill={selectedColor.bar} 
                        radius={layout === "horizontal" ? [6, 6, 0, 0] : [0, 6, 6, 0]}
                        animationDuration={1000}
                      >
                        {chartData.map((entry, index) => (
                          <Cell 
                            key={`cell-${index}`} 
                            fill={selectedColor.bar} 
                            className="transition-all duration-300 hover:opacity-85 cursor-pointer"
                          />
                        ))}
                        {showLabels && (
                          <LabelList 
                            dataKey="scaledRevenue" 
                            position={layout === "horizontal" ? "top" : "right"} 
                            fill="#0f172a" 
                            fontSize={10}
                            fontFamily="monospace"
                            fontWeight="bold"
                            formatter={(value: any) => `${value} ${currentScaleInfo.unit || "đ"}`}
                          />
                        )}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}
            </div>

            {/* Quick Metrics and analytics footer */}
            {chartData.length > 0 && (
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 shrink-0 font-sans">
                
                {/* Metric 1 */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider font-mono">TỔNG CHỈ TIÊU ĐANG CHỌN</div>
                  <div className="text-lg font-extrabold text-emerald-700 font-mono truncate">
                    {totalRevenueSelected} <span className="text-xs text-slate-500 font-semibold font-sans">đơn vị</span>
                  </div>
                  <div className="text-[10px] text-slate-500">Quy nạp từ {selectedSectors.length} nhóm phân loại</div>
                </div>

                {/* Metric 2 */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider font-mono">NHÓM CO' GIÁ TRỊ LỚN NHẤT</div>
                  {maxRevenueSector ? (
                    <>
                      <div className="text-xs font-bold text-amber-800 truncate flex items-center gap-1.5">
                        <span className="px-1.5 py-0.2 bg-amber-100 text-amber-800 text-[10px] font-mono rounded shrink-0">{maxRevenueSector.code}</span>
                        <span className="text-slate-800 truncate font-semibold">{maxRevenueSector.name}</span>
                      </div>
                      <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                        Giá trị: <span className="font-bold text-emerald-700">{new Intl.NumberFormat("vi-VN").format(maxRevenueSector.revenue)}đ</span>
                      </div>
                    </>
                  ) : (
                    <div className="text-xs text-slate-400 italic">Chưa xác định</div>
                  )}
                </div>

                {/* Metric 3 */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider font-mono">PHÂN PHỐI LƯỢNG MẪU NGUỒN</div>
                  <div className="text-xs text-slate-700 font-mono">
                    <span className="text-base font-bold text-indigo-700">{activeSectorsData.reduce((sum, i) => sum + i.recordCount, 0)}</span> dòng / ghi chép gốc
                  </div>
                  <div className="text-[10px] text-slate-500">
                    {reportLevel === 0 ? "Phân tích theo đặc tính phân nhóm trực tiếp" : "Tương thích phân lớp chuẩn VSIC 2018"}
                  </div>
                </div>

              </div>
            )}

          </div>

        </div>
      ) : (
        <div className="bg-white rounded-2xl p-16 text-center border border-slate-200 space-y-3 shadow-xl">
          <div className="p-4 bg-amber-50 rounded-full inline-block text-amber-600 border border-amber-200">
            <AlertCircle className="w-8 h-8" />
          </div>
          <div className="text-base font-bold text-slate-800">Yêu cầu chọn cấu hình và chỉ tiêu số</div>
          <p className="text-xs text-slate-500 max-w-sm mx-auto leading-relaxed">
            Biểu đồ trực quan yêu cầu tệp dữ liệu chính của bạn đã được tải lên thành công, và các cột <b>Nhóm/Mã Ngành</b>, <b>Chỉ tiêu số</b> phải được chỉ định khớp nối.
          </p>
        </div>
      )}
    </div>
  );
});

SectorRevenueChart.displayName = "SectorRevenueChart";

export default SectorRevenueChart;
