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
  vsicRawData 
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
}

// Interface for Level 1 Sector summary
interface Level1Summary {
  code: string;
  name: string;
  revenue: number;
  communeCount: number;
  recordCount: number;
}

export default function SectorRevenueChart({ mainData, columns, mapping }: SectorRevenueChartProps) {
  // Local state for configuration columns (fallback to mapped ones)
  const [selectedManganh, setSelectedManganh] = useState<string>(mapping.manganh || "");
  const [selectedDoanhthu, setSelectedDoanhthu] = useState<string>(mapping.doanhthu || "");
  
  // Format Scale (Auto, Original, Million VND, Billion VND)
  const [scale, setScale] = useState<"auto" | "raw" | "million" | "billion">("auto");
  // Chart orientation (horizontal = category on X, vertical = category on Y)
  const [layout, setLayout] = useState<"horizontal" | "vertical">("horizontal");
  // Show labels on top of bars
  const [showLabels, setShowLabels] = useState<boolean>(true);
  // Bar Color Theme
  const [colorTheme, setColorTheme] = useState<"violet" | "indigo" | "emerald" | "amber">("indigo");

  // Reset local state if global mapping changes
  React.useEffect(() => {
    if (mapping.manganh && !selectedManganh) {
      setSelectedManganh(mapping.manganh);
    }
    if (mapping.doanhthu && !selectedDoanhthu) {
      setSelectedDoanhthu(mapping.doanhthu);
    }
  }, [mapping]);

  // Aggregate Data to Level 1
  const aggregatedData = useMemo(() => {
    if (mainData.length === 0 || !selectedManganh) return [];

    const map: { [code: string]: { revenue: number; records: number; communes: Set<string> } } = {};

    mainData.forEach(row => {
      // Normalize raw sector code
      const rawMng = row[selectedManganh];
      const mng = normalizeSectorCode(rawMng);

      // Determine level 1 code (Letter)
      let l1Code = "";
      if (mng) {
        if (/^[a-zA-Z]$/.test(mng)) {
          l1Code = mng.toUpperCase();
        } else {
          // Standard Level 2 and down
          const l2Code = mng.slice(0, 2);
          l1Code = getParentSectorCode(l2Code) || "";
        }
      }

      if (!l1Code) {
        l1Code = "KHAC"; // Unknown / Others
      }

      // Parse Revenue
      let revVal = 0;
      if (selectedDoanhthu && row[selectedDoanhthu] !== undefined) {
        const rawRev = row[selectedDoanhthu];
        const num = parseFloat(String(rawRev).replace(/[^0-9.-]/g, ""));
        if (!isNaN(num)) {
          revVal = num;
        }
      }

      // Track communes if available
      const communeVal = mapping.xa && row[mapping.xa] ? String(row[mapping.xa]).trim() : "Khác";

      if (!map[l1Code]) {
        map[l1Code] = { revenue: 0, records: 0, communes: new Set() };
      }
      map[l1Code].revenue += revVal;
      map[l1Code].records += 1;
      if (communeVal) {
        map[l1Code].communes.add(communeVal);
      }
    });

    // Form summary structure
    const results: Level1Summary[] = Object.keys(map).map(code => {
      const name = vsicRawData[code] || (code === "KHAC" ? "Mã ngành chưa khớp VSIC / Khác" : `Ngành cấp 1 chưa định nghĩa (${code})`);
      return {
        code,
        name,
        revenue: Math.round(map[code].revenue * 100) / 100,
        communeCount: map[code].communes.size,
        recordCount: map[code].records
      };
    });

    // Sort by revenue descending
    return results.sort((a, b) => b.revenue - a.revenue);
  }, [mainData, selectedManganh, selectedDoanhthu, mapping.xa]);

  // List of all unique Level 1 codes in aggregate data
  const allSectors = useMemo(() => {
    return aggregatedData.map(item => ({
      code: item.code,
      name: item.name,
      revenue: item.revenue
    }));
  }, [aggregatedData]);

  // State for which sectors are selected to display in chart
  const [selectedSectors, setSelectedSectors] = useState<string[]>([]);
  // Keep selected sectors updated with all sectors initially
  const [hasInitializedSectors, setHasInitializedSectors] = useState<string>("");

  const currentDatasetId = `${mainData.length}_${selectedManganh}_${selectedDoanhthu}`;
  useEffect(() => {
    if (allSectors.length > 0 && hasInitializedSectors !== currentDatasetId) {
      setSelectedSectors(allSectors.map(s => s.code));
      setHasInitializedSectors(currentDatasetId);
    }
  }, [allSectors, hasInitializedSectors, currentDatasetId]);

  // Handle Sector Checklist Toggles
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

  // Determine optimal scaling factor based on max revenue of selected items
  const activeSectorsData = useMemo(() => {
    return aggregatedData.filter(item => selectedSectors.includes(item.code));
  }, [aggregatedData, selectedSectors]);

  const autoDetectedScale = useMemo(() => {
    if (activeSectorsData.length === 0) return { factor: 1, unit: "" };
    const maxVal = Math.max(...activeSectorsData.map(item => item.revenue));
    
    // If maximum value is in millions (10^6) or billions (10^9)
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
    return autoDetectedScale; // default auto
  }, [scale, autoDetectedScale]);

  // Scale data for graph rendering
  const chartData = useMemo(() => {
    return activeSectorsData.map(item => {
      const scaledRev = item.revenue / currentScaleInfo.factor;
      return {
        ...item,
        scaledRevenue: Math.round(scaledRev * 100) / 100,
        codeLabel: item.code, // Letter e.g. A, B
        fullName: `${item.code} - ${item.name}`,
        rawRevenueFormatted: new Intl.NumberFormat("vi-VN").format(item.revenue)
      };
    }).sort((a, b) => b.scaledRevenue - a.scaledRevenue); // Descending order for clean look
  }, [activeSectorsData, currentScaleInfo]);

  // Color map
  const colors = {
    indigo: { bar: "#6366f1", barHover: "#4f46e5", gradient: ["#4f46e5", "#818cf8"] },
    violet: { bar: "#8b5cf6", barHover: "#7c3aed", gradient: ["#7c3aed", "#a78bfa"] },
    emerald: { bar: "#10b981", barHover: "#059669", gradient: ["#059669", "#34d399"] },
    amber: { bar: "#f59e0b", barHover: "#d97706", gradient: ["#d97706", "#fbbf24"] }
  };

  const selectedColor = colors[colorTheme];

  // Quick Stats
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

  return (
    <div className="space-y-6">
      {/* Configuration Header Panel */}
      <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-5 space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h3 className="text-lg font-bold text-white flex items-center gap-2">
              <BarChart3 className="w-5 h-5 text-amber-400" /> BIỂU ĐỒ TRỰC QUAN DOANH THU THEO NGÀNH CẤP 1
            </h3>
            <p className="text-xs text-gray-400">
              Tổng hợp và so sánh tự động tổng doanh thu của các doanh nghiệp được quy nạp lên Ngành Cấp 1 toàn quốc (Danh mục VSIC 2018).
            </p>
          </div>

          <div className="flex flex-wrap gap-2.5">
            {/* Column selectors if users want to change setup dynamically */}
            <div className="flex items-center gap-1.5 bg-[#111827] px-3 py-1.5 rounded-xl border border-gray-800">
              <label className="text-[10px] uppercase font-bold text-gray-500 font-mono">Cột Mã Ngành:</label>
              <select
                value={selectedManganh}
                onChange={(e) => setSelectedManganh(e.target.value)}
                className="bg-transparent text-xs text-amber-400 font-bold border-none outline-none focus:ring-0 max-w-[130px]"
              >
                <option value="">-- Chọn cột --</option>
                {columns.map(col => (
                  <option key={col} value={col} className="bg-[#1f2937] text-white">{col}</option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-1.5 bg-[#111827] px-3 py-1.5 rounded-xl border border-gray-800">
              <label className="text-[10px] uppercase font-bold text-gray-500 font-mono">Cột Doanh Thu:</label>
              <select
                disabled={!columns.length}
                value={selectedDoanhthu}
                onChange={(e) => setSelectedDoanhthu(e.target.value)}
                className="bg-transparent text-xs text-emerald-400 font-bold border-none outline-none focus:ring-0 max-w-[130px]"
              >
                <option value="">-- Chọn cột --</option>
                {columns.map(col => (
                  <option key={col} value={col} className="bg-[#1f2937] text-white">{col}</option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* General Guidelines Warning */}
        {(!selectedManganh || !selectedDoanhthu) && (
          <div className="bg-amber-950/25 border border-amber-500/20 rounded-xl p-4 flex gap-3 text-xs text-amber-300">
            <AlertCircle className="w-5 h-5 shrink-0 text-amber-400" />
            <div>
              <p className="font-semibold">Chưa khớp hoặc thiếu cấu hình Mapping gốc!</p>
              <p className="mt-1 text-[11px] text-gray-400">
                Hãy lựa chọn cột Mã Ngành và cột Doanh Thu thích hợp từ bảng dữ liệu nguồn tải lên ở góc phải để hệ thống tự lọc quy nạp cột và vẽ đồ thị.
              </p>
            </div>
          </div>
        )}
      </div>

      {mainData.length > 0 && selectedManganh && selectedDoanhthu ? (
        <div className="grid grid-cols-1 xl:grid-cols-4 gap-6">
          
          {/* LEFT: Checklist filter sidebar */}
          <div className="xl:col-span-1 bg-[#1f2937]/50 border border-[#374151] rounded-2xl p-4 flex flex-col space-y-4 max-h-[640px] overflow-hidden">
            <div>
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-xs font-bold text-gray-300 uppercase tracking-wider font-mono">Chọn Ngành Cấp 1 ({selectedSectors.length}/{allSectors.length})</span>
              </div>
              <p className="text-[10px] text-gray-400 leading-relaxed">
                Tích chọn các ngành kinh tế muốn hiển thị đối sánh trực tiếp trên biểu đồ:
              </p>
            </div>

            {/* Quick multi-select buttons */}
            <div className="flex gap-2">
              <button
                onClick={handleSelectAll}
                className="flex-1 bg-gray-800 hover:bg-gray-700 text-gray-300 font-semibold text-[10px] py-1.5 rounded-lg border border-gray-700 hover:text-white transition-all cursor-pointer"
              >
                Chọn tất cả
              </button>
              <button
                onClick={handleDeselectAll}
                className="flex-1 bg-gray-900/40 hover:bg-gray-800 text-gray-400 font-semibold text-[10px] py-1.5 rounded-lg border border-gray-800 hover:text-gray-300 transition-all cursor-pointer"
              >
                Bỏ chọn hết
              </button>
            </div>

            {/* Sector list display inside a scrollbar */}
            <div className="flex-1 overflow-y-auto space-y-1.5 pr-1 text-xs">
              {allSectors.length === 0 ? (
                <div className="text-center italic text-gray-500 py-8">
                  Không quét được mã ngành cấp 1 nào hợp lệ trong cột dữ liệu chính.
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
                          ? "bg-purple-600/10 border-purple-500/20 text-purple-200"
                          : "bg-transparent border-transparent text-gray-400 hover:bg-gray-800/40"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => handleToggleSector(sector.code)}
                        className="mt-0.5 rounded border-gray-600 text-purple-600 focus:ring-purple-600 focus:ring-offset-[#111827] bg-gray-900 h-3.5 w-3.5 cursor-pointer"
                      />
                      <div className="space-y-0.5 min-w-0">
                        <div className="font-bold flex items-center gap-1.5">
                          <span className="text-amber-400 font-mono">{sector.code}</span>
                          <span className="truncate text-gray-300 group-hover:text-white transition-all text-[11px] block">{sector.name}</span>
                        </div>
                        <div className="text-[10px] text-gray-500 font-mono flex justify-between">
                          <span>Doanh thu:</span>
                          <span className="font-semibold text-emerald-400">{formattedRev}</span>
                        </div>
                      </div>
                    </label>
                  );
                })
              )}
            </div>
          </div>

          {/* RIGHT: Visual charts panel & tool toggles */}
          <div className="xl:col-span-3 space-y-6 flex flex-col">
            
            {/* Visual control options */}
            <div className="bg-[#1f2937]/40 border border-[#374151]/60 rounded-2xl p-4 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
              
              {/* Scale config */}
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider font-mono">Đơn vị đo lường</label>
                <select
                  value={scale}
                  onChange={(e) => setScale(e.target.value as any)}
                  className="bg-[#111827] border border-[#374151] rounded-xl px-3 py-1.5 text-xs text-white w-full font-bold focus:ring-purple-500"
                >
                  <option value="auto">Tự động (Khuyên dùng)</option>
                  <option value="raw">Đồng (VND gốc)</option>
                  <option value="million">Triệu đồng (1.000.000đ)</option>
                  <option value="billion">Tỷ đồng (1.000.000.000đ)</option>
                </select>
              </div>

              {/* Layout Config */}
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider font-mono">Hướng biểu đồ</label>
                <div className="grid grid-cols-2 gap-1 bg-[#111827] border border-[#374151] rounded-xl p-1">
                  <button
                    onClick={() => setLayout("horizontal")}
                    className={`text-xs py-1 rounded-lg font-bold transition-all cursor-pointer ${
                      layout === "horizontal" ? "bg-purple-600 text-white shadow" : "text-gray-400 hover:text-white"
                    }`}
                  >
                    Cột đứng
                  </button>
                  <button
                    onClick={() => setLayout("vertical")}
                    className={`text-xs py-1 rounded-lg font-bold transition-all cursor-pointer ${
                      layout === "vertical" ? "bg-purple-600 text-white shadow" : "text-gray-400 hover:text-white"
                    }`}
                  >
                    Thanh ngang
                  </button>
                </div>
              </div>

              {/* Labels on top */}
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider font-mono">Hiện số liệu trực tiếp</label>
                <div className="grid grid-cols-2 gap-1 bg-[#111827] border border-[#374151] rounded-xl p-1">
                  <button
                    onClick={() => setShowLabels(true)}
                    className={`text-xs py-1 rounded-lg font-bold transition-all cursor-pointer ${
                      showLabels ? "bg-emerald-600 text-white shadow" : "text-gray-400 hover:text-white"
                    }`}
                  >
                    Bật
                  </button>
                  <button
                    onClick={() => setShowLabels(false)}
                    className={`text-xs py-1 rounded-lg font-bold transition-all cursor-pointer ${
                      !showLabels ? "bg-emerald-600 text-white shadow" : "text-gray-400 hover:text-white"
                    }`}
                  >
                    Tắt
                  </button>
                </div>
              </div>

              {/* Color themes */}
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider font-mono">Tông màu chủ đạo</label>
                <select
                  value={colorTheme}
                  onChange={(e) => setColorTheme(e.target.value as any)}
                  className="bg-[#111827] border border-[#374151] rounded-xl px-3 py-1.5 text-xs text-white w-full font-bold focus:ring-purple-500"
                >
                  <option value="indigo">🌌 Twilight Indigo</option>
                  <option value="violet">🔮 Galactic Violet</option>
                  <option value="emerald">💚 Forest Emerald</option>
                  <option value="amber">💛 Honey Amber</option>
                </select>
              </div>

            </div>

            {/* CHART DISPLAY CORE CONTAINER */}
            <div className="bg-[#111827] border border-purple-950/15 rounded-2xl p-5 shadow-2xl relative flex-1 flex flex-col justify-center min-h-[420px]">
              {chartData.length === 0 ? (
                <div className="text-center space-y-2 py-20 z-10">
                  <span className="p-3 bg-gray-800/60 rounded-full inline-block text-amber-400 border border-amber-950/40">
                    <AlertCircle className="w-6 h-6 animate-pulse" />
                  </span>
                  <div className="text-sm font-semibold text-gray-300">Biểu đồ trống!</div>
                  <p className="text-xs text-gray-500 max-w-sm mx-auto">Vui lòng tích chọn ít nhất một ngành cấp 1 ở bảng điều khiển bên trái để bắt đầu lập biểu đồ.</p>
                </div>
              ) : (
                <div className="w-full h-[400px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={chartData}
                      layout={layout}
                      margin={{ top: 20, right: 35, left: 15, bottom: layout === "horizontal" ? 10 : 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={layout === "vertical"} />
                      
                      {layout === "horizontal" ? (
                        <>
                          <XAxis 
                            dataKey="codeLabel" 
                            stroke="#9ca3af" 
                            fontSize={11} 
                            tickLine={false}
                            axisLine={false} 
                            dy={5}
                          />
                          <YAxis 
                            stroke="#9ca3af" 
                            fontSize={10} 
                            tickLine={false}
                            axisLine={false} 
                            unit={` ${currentScaleInfo.unit}`}
                          />
                        </>
                      ) : (
                        <>
                          <XAxis 
                            type="number"
                            stroke="#9ca3af" 
                            fontSize={10} 
                            tickLine={false}
                            axisLine={false} 
                            unit={` ${currentScaleInfo.unit}`}
                          />
                          <YAxis 
                            type="category"
                            dataKey="codeLabel" 
                            stroke="#9ca3af" 
                            fontSize={11} 
                            tickLine={false}
                            axisLine={false} 
                            dx={-5}
                          />
                        </>
                      )}

                      <Tooltip
                        cursor={{ fill: '#374151', opacity: 0.2 }}
                        content={({ active, payload }) => {
                          if (active && payload && payload.length) {
                            const data = payload[0].payload;
                            return (
                              <div className="bg-[#1f2937] border border-purple-500/30 rounded-xl p-3.5 shadow-xl max-w-[280px] font-sans leading-relaxed text-xs">
                                <div className="font-bold flex items-center gap-1.5 border-b border-gray-700/60 pb-1.5 mb-1.5">
                                  <span className="px-1.5 py-0.5 bg-amber-500/10 text-amber-400 font-mono text-[10px] rounded border border-amber-500/20">{data.code}</span>
                                  <span className="truncate text-gray-200">{data.name}</span>
                                </div>
                                <div className="space-y-1">
                                  <div className="flex justify-between gap-4 text-gray-400">
                                    <span>Doanh thu:</span>
                                    <span className="font-mono text-emerald-400 font-semibold">{data.rawRevenueFormatted} đ</span>
                                  </div>
                                  <div className="flex justify-between gap-4 text-gray-400">
                                    <span>Bản ghi nguồn:</span>
                                    <span className="font-mono text-indigo-300 font-medium">{data.recordCount} dòng</span>
                                  </div>
                                  <div className="flex justify-between gap-4 text-gray-400">
                                    <span>Hơn:</span>
                                    <span className="font-mono text-pink-300 font-medium">{data.communeCount} địa bàn xã</span>
                                  </div>
                                </div>
                              </div>
                            );
                          }
                          return null;
                        }}
                      />

                      <Bar
                        dataKey="scaledRevenue"
                        radius={layout === "horizontal" ? [6, 6, 0, 0] : [0, 6, 6, 0]}
                        animationDuration={1000}
                      >
                        {chartData.map((entry, index) => (
                          <Cell 
                            key={`cell-${index}`} 
                            fill={selectedColor.bar} 
                            className="transition-all duration-300 hover:opacity-85"
                          />
                        ))}
                        {showLabels && (
                          <LabelList 
                            dataKey="scaledRevenue" 
                            position={layout === "horizontal" ? "top" : "right"} 
                            fill="#cbd5e1" 
                            fontSize={10}
                            fontFamily="monospace"
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
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                
                {/* Metric 1 */}
                <div className="bg-[#1f2937]/30 border border-[#374151]/40 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider font-mono">TỔNG DOANH THU ĐANG CHỌN</div>
                  <div className="text-lg font-extrabold text-emerald-400 font-mono truncate">{totalRevenueSelected} <span className="text-xs text-gray-400 font-semibold font-sans">đồng</span></div>
                  <div className="text-[10px] text-gray-400">Quy nạp từ {selectedSectors.length} ngành chính</div>
                </div>

                {/* Metric 2 */}
                <div className="bg-[#1f2937]/30 border border-[#374151]/40 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider font-mono">NGÀNH CÓ DOANH THU LỚN NHẤT</div>
                  {maxRevenueSector ? (
                    <>
                      <div className="text-sm font-bold text-amber-400 truncate flex items-center gap-1.5">
                        <span className="px-1.5 py-0.2 bg-amber-500/15 text-amber-400 text-[10px] font-mono rounded">{maxRevenueSector.code}</span>
                        <span className="text-gray-200">{maxRevenueSector.name}</span>
                      </div>
                      <div className="text-[10px] text-gray-400 font-mono">
                        Số tiền: <span className="font-semibold text-emerald-400">{new Intl.NumberFormat("vi-VN").format(maxRevenueSector.revenue)}đ</span>
                      </div>
                    </>
                  ) : (
                    <div className="text-xs text-gray-500 italic">Chưa xác định</div>
                  )}
                </div>

                {/* Metric 3 */}
                <div className="bg-[#1f2937]/30 border border-[#374151]/40 rounded-xl p-4 space-y-1 flex flex-col justify-center">
                  <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider font-mono">PHÂN PHỐI LƯỢNG MẪU NGUỒN</div>
                  <div className="text-xs text-gray-300 font-mono">
                    <span className="text-base font-bold text-purple-400">{activeSectorsData.reduce((sum, i) => sum + i.recordCount, 0)}</span> dòng / ghi chép gốc
                  </div>
                  <div className="text-[10px] text-gray-400">Tương thích phân lớp chuẩn VSIC 2018</div>
                </div>

              </div>
            )}

          </div>

        </div>
      ) : (
        <div className="bg-[#111827]/50 rounded-2xl p-16 text-center border border-[#1f2937]/60 space-y-3">
          <div className="p-4 bg-gray-800/40 rounded-full inline-block text-amber-400 border border-amber-950/40">
            <AlertCircle className="w-8 h-8 animate-bounce" />
          </div>
          <div className="text-base font-bold text-white">Yêu cầu hoàn tất nạp và cấu hình dữ liệu</div>
          <p className="text-xs text-gray-400 max-w-md mx-auto leading-relaxed">
            Biểu đồ trực quan yêu cầu tệp dữ liệu chính của bạn đã được tải lên thành công, và các cột <b>Mã Ngành</b>, <b>Doanh Thu</b> phải được chọn cấu hình khớp nối ở menu phía trên.
          </p>
        </div>
      )}
    </div>
  );
}
