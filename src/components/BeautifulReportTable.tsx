import React, { useState, useMemo } from "react";
import { Search, Eye, EyeOff, Download, ChevronRight, ChevronDown, CheckCircle } from "lucide-react";

interface BeautifulReportTableProps {
  rows: any[];
  cols: string[];
  level: number;
  reportType: "pivot" | "flat";
  onExport: () => void;
}

export const BeautifulReportTable = React.memo<BeautifulReportTableProps>(({
  rows,
  cols,
  level,
  reportType,
  onExport,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [hideZeroColumns, setHideZeroColumns] = useState(true);
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);

  // Filter rows based on search term
  const filteredRows = useMemo(() => {
    if (!searchTerm.trim()) return rows;
    const term = searchTerm.toLowerCase();
    return rows.filter((row) => {
      const commune = String(row["Địa_Bàn_Xã"] || row["Địa_bàn_Xã"] || "").toLowerCase();
      const sectorKey = level === 0 ? "Nhóm_Phân_Loại" : `Ngành_Cấp_${level}`;
      const sector = String(row[sectorKey] || "").toLowerCase();
      return commune.includes(term) || sector.includes(term);
    });
  }, [rows, searchTerm, level]);

  // Extract sectors and indicators dynamically from column list
  const pivotAnalysis = useMemo(() => {
    if (reportType !== "pivot" || rows.length === 0) {
      return { sectors: [], indicators: [], activeSectors: [] };
    }

    const sectorSet = new Set<string>();
    const indicatorSet = new Set<string>();

    cols.forEach((col) => {
      if (col.includes(" - Tổng ")) {
        const parts = col.split(" - Tổng ");
        if (parts.length === 2) {
          sectorSet.add(parts[0]);
          indicatorSet.add(parts[1]);
        }
      }
    });

    const sectors = Array.from(sectorSet).sort();
    const indicators = Array.from(indicatorSet); // Keep their selection order!

    // Determine which sectors have non-zero values in at least one row for at least one indicator
    const activeSectors = sectors.filter((sector) => {
      return rows.some((row) => {
        return indicators.some((ind) => {
          const val = row[`${sector} - Tổng ${ind}`] || 0;
          return val > 0;
        });
      });
    });

    return { sectors, indicators, activeSectors };
  }, [rows, cols, reportType]);

  // Columns to display based on toggle
  const visibleSectors = useMemo(() => {
    if (reportType !== "pivot") return [];
    return hideZeroColumns ? pivotAnalysis.activeSectors : pivotAnalysis.sectors;
  }, [pivotAnalysis, hideZeroColumns, reportType]);

  // Calculate dynamic totals across all columns
  const overallTotals = useMemo(() => {
    let grandDN = 0;
    const totalsByIndicator: { [ind: string]: number } = {};
    const sectorStats: { [sector: string]: { [ind: string]: number } } = {};
    
    rows.forEach((row) => {
      if (reportType === "pivot") {
        grandDN += row["Số_Dòng_Tổng_Hợp"] ?? row["Số_DN_Địa_Phương"] ?? 0;
        
        pivotAnalysis.indicators.forEach((ind) => {
          if (totalsByIndicator[ind] === undefined) totalsByIndicator[ind] = 0;
          totalsByIndicator[ind] += row[`Tổng_Cộng_${ind}_Toàn_Xã`] ?? row[`Tổng_Cộng_Toàn_Xã_${ind}`] ?? 0;
        });

        pivotAnalysis.sectors.forEach((sector) => {
          if (!sectorStats[sector]) {
            sectorStats[sector] = {};
          }
          pivotAnalysis.indicators.forEach((ind) => {
            if (sectorStats[sector][ind] === undefined) sectorStats[sector][ind] = 0;
            sectorStats[sector][ind] += row[`${sector} - Tổng ${ind}`] || 0;
          });
        });
      } else {
        // Flat mode
        grandDN += row["Số_Lượng_Bản_Ghi"] ?? row["Số_Lượng_Doanh_Nghiệp"] ?? 0;
        cols.forEach((col) => {
          if (col.startsWith("Tổng_")) {
            const ind = col.replace("Tổng_", "");
            if (totalsByIndicator[ind] === undefined) totalsByIndicator[ind] = 0;
            totalsByIndicator[ind] += row[col] || 0;
          }
        });
      }
    });

    return {
      dn: grandDN,
      totalsByIndicator,
      sectorStats
    };
  }, [rows, reportType, pivotAnalysis.sectors, pivotAnalysis.indicators, cols]);

  // Cell representation formatting (0 -> -)
  const formatCellValue = (val: any, isNumeric: boolean) => {
    if (isNumeric) {
      if (val === undefined || val === null || val === 0) {
        return <span className="text-slate-400 font-normal">—</span>;
      }
      return <span className="font-mono text-emerald-700 font-bold">{Math.round(val * 100) / 100 ? val.toLocaleString("vi-VN") : "—"}</span>;
    }
    return String(val ?? "");
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden space-y-4 p-5" id="beautiful_report_container">
      {/* HEADER BAR */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200 font-sans">
        <div>
          <span className="text-xs font-black text-slate-800 font-mono tracking-widest uppercase block mb-1">
            📊 KẾT QUẢ TỔNG HỢP {level === 0 ? "PHÂN NHÓM" : `DANH MỤC NGÀNH CẤP ${level}`} × ĐỊA BÀN
          </span>
          <span className="text-xs text-slate-600 block leading-normal">
            Bản hợp nhất động từ các tiêu chí và cột hạch toán do bạn tùy chỉnh. Đã tổng cộng dồn lũy kế tự động.
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          {/* SEARCH BAR */}
          <div className="relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none">
              <Search className="h-3.5 w-3.5 text-slate-400" />
            </span>
            <input
              type="text"
              placeholder="Tìm kiếm địa bàn xã..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="bg-white border border-slate-300 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-800 placeholder-slate-450 focus:outline-none focus:ring-1 focus:ring-indigo-500 w-52 font-medium"
            />
          </div>

          {/* HIDE EMPTY COLUMNS TOGGLE (ONLY FOR PIVOT) */}
          {reportType === "pivot" && pivotAnalysis.sectors.length > 0 && (
            <button
              onClick={() => setHideZeroColumns(!hideZeroColumns)}
              className={`text-xs font-bold py-1.5 px-3 rounded-lg border transition-all flex items-center gap-1.5 select-none md:w-auto w-full justify-center cursor-pointer ${
                hideZeroColumns
                  ? "bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100/60"
                  : "bg-slate-100 text-slate-700 border-slate-250 hover:bg-slate-200"
              }`}
              title="Nhấn để ẩn/hiện các cột ngành không phát sinh dữ liệu"
            >
              {hideZeroColumns ? <EyeOff className="w-3.5 h-3.5 text-amber-700" /> : <Eye className="w-3.5 h-3.5 text-slate-600" />}
              {hideZeroColumns ? `Đang ẩn cột rỗng (${pivotAnalysis.sectors.length - pivotAnalysis.activeSectors.length})` : "Hiện tất cả cột ngành"}
            </button>
          )}

          {/* EXCEL EXPORT BUTTON */}
          <button
            onClick={onExport}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-md inline-flex justify-center"
          >
            <Download className="w-3.5 h-3.5" /> Xuất Excel (.xlsx)
          </button>
        </div>
      </div>

      {/* DYNAMIC INSIGHTS / STATS CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 font-sans">
        {pivotAnalysis.indicators.slice(0, 2).map((ind, idx) => {
          const totalVal = overallTotals.totalsByIndicator[ind] || 0;
          return (
            <div 
              key={ind} 
              className={`p-4 rounded-xl border space-y-1 bg-gradient-to-br ${
                idx === 0 
                  ? "from-emerald-50/50 to-emerald-50/20 border-emerald-100" 
                  : "from-indigo-50/50 to-indigo-50/20 border-indigo-100"
              }`}
            >
              <div className={`text-[10px] uppercase tracking-wider font-mono font-bold flex items-center gap-1 ${idx === 0 ? "text-emerald-700" : "text-indigo-700"}`}>
                {idx === 0 ? "💰" : "👥"} TỔNG CỘNG: {ind}
              </div>
              <div className="text-lg font-black text-slate-900 mt-0.5 font-mono">
                {totalVal.toLocaleString("vi-VN")} <span className="text-xs font-normal text-slate-500 font-sans">tổng cộng</span>
              </div>
              <div className="text-[10px] text-slate-500">Tổng quy nạp lũy kế trên toàn bộ {rows.length} địa bàn</div>
            </div>
          );
        })}

        {/* Fallback cards if less than 2 indicators selected */}
        {pivotAnalysis.indicators.length === 0 && (
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-1">
            <div className="text-[10px] text-slate-500 uppercase font-mono font-bold">💰 CHƯA CÓ CHỈ TIÊU SỐ</div>
            <div className="text-lg font-bold text-slate-450">0</div>
            <div className="text-[10px] text-slate-400">Tick chọn chỉ tiêu ở bảng cấu hình</div>
          </div>
        )}

        {pivotAnalysis.indicators.length < 2 && pivotAnalysis.indicators.length > 0 && (
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-1">
            <div className="text-[10px] text-slate-500 uppercase font-mono font-bold">ℹ️ CHỈ TIÊU PHỤ</div>
            <div className="text-sm font-semibold text-slate-500 italic">Bổ sung chỉ tiêu số ở bảng gộp</div>
            <div className="text-[10px] text-slate-400">Chỉ chọn duy nhất 1 cột chỉ tiêu chính</div>
          </div>
        )}

        <div className="bg-gradient-to-br from-teal-50/60 to-teal-50/20 p-4 rounded-xl border border-teal-100 space-y-1">
          <div className="text-[10px] text-teal-700 uppercase tracking-wider font-mono font-bold flex items-center gap-1">
            🏢 TỔNG LƯỢNG MẪU NGUỒN
          </div>
          <div className="text-lg font-black text-slate-900 mt-0.5 font-mono">
            {overallTotals.dn.toLocaleString("vi-VN")} <span className="text-xs font-normal text-slate-500 font-sans">cơ sở/dòng d.liệu</span>
          </div>
          <div className="text-[10px] text-slate-500">Ghi nhận đại diện cho các địa bàn đã nạp</div>
        </div>
      </div>

      {/* THE MAIN TABLE CORES */}
      <div className="relative font-sans">
        {filteredRows.length === 0 ? (
          <div className="text-center py-10 bg-slate-50 border border-slate-200 rounded-xl text-slate-500 text-xs">
            Không tìm thấy địa bàn nào khớp với từ khóa tìm kiếm.
          </div>
        ) : reportType === "pivot" ? (
          /* ================== PIVOT MATRIX LAYOUT (DYNAMIC & HIGHLY POLISHED) ================== */
          <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white max-h-[500px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                {/* FIRST HEADER ROW: GROUPED SECTORS */}
                <tr className="bg-slate-100 border-b border-slate-200 text-slate-800 font-mono text-[10.5px] sticky top-0 z-20">
                  <th rowSpan={2} className="p-3 font-bold whitespace-nowrap bg-slate-100 border-r border-slate-200 align-middle text-center sticky left-0 z-30">
                    Địa Bàn Xã/Phường
                  </th>
                  {visibleSectors.map((sector) => (
                    <th 
                      key={sector} 
                      colSpan={pivotAnalysis.indicators.length} 
                      className="p-2 font-bold whitespace-nowrap text-center border-r border-b border-slate-200 bg-slate-50 font-sans text-[11px] text-slate-800"
                    >
                      {sector}
                    </th>
                  ))}
                  <th 
                    colSpan={1 + pivotAnalysis.indicators.length} 
                    className="p-2 font-bold whitespace-nowrap text-center bg-teal-50 text-teal-900 align-middle border-b border-slate-200"
                  >
                    TỔNG HỢP ĐỊA PHƯƠNG
                  </th>
                </tr>

                {/* SECOND HEADER ROW: INDICATORS */}
                <tr className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-mono text-[9px] sticky top-[31px] z-20 uppercase">
                  {visibleSectors.map((sector) => (
                    <React.Fragment key={`sub-${sector}`}>
                      {pivotAnalysis.indicators.map((ind, indIdx) => (
                        <th 
                          key={ind} 
                          className="p-2 font-medium text-right border-r bg-slate-50/55 whitespace-nowrap"
                          style={{ 
                            borderColor: indIdx === pivotAnalysis.indicators.length - 1 ? '#e2e8f0' : '#f1f5f9',
                            width: pivotAnalysis.indicators.length > 2 ? '80px' : '96px'
                          }}
                        >
                          {ind}
                        </th>
                      ))}
                    </React.Fragment>
                  ))}
                  <th className="p-2 font-semibold text-right border-r border-slate-200 bg-teal-50/40 text-teal-800 whitespace-nowrap w-20">
                    Số mẫu
                  </th>
                  {pivotAnalysis.indicators.map((ind) => (
                    <th key={`hdr-tot-${ind}`} className="p-2 font-semibold text-right border-r border-slate-200 bg-teal-50/40 text-teal-800 whitespace-nowrap">
                      {ind}
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100 text-slate-800 font-sans text-[11.5px]">
                {filteredRows.map((row, rIdx) => {
                  const isSelected = selectedRowIndex === rIdx;
                  const communeName = row["Địa_Bàn_Xã"] || row["Địa_bàn_Xã"] || "Khác";
                  return (
                    <React.Fragment key={`row-${rIdx}`}>
                      <tr
                        onClick={() => setSelectedRowIndex(isSelected ? null : rIdx)}
                        className={`hover:bg-slate-50 transition-colors cursor-pointer ${
                          isSelected ? "bg-amber-50/60 font-semibold border-l-4 border-amber-500" : rIdx % 2 === 1 ? "bg-slate-50/40" : ""
                        }`}
                      >
                        {/* Commune Name */}
                        <td className="p-3 whitespace-nowrap font-sans font-medium text-slate-900 bg-white border-r border-slate-200 sticky left-0 z-10 flex items-center gap-1.5 shadow-sm">
                          {isSelected ? <ChevronDown className="w-3 h-3 text-amber-600" /> : <ChevronRight className="w-3 h-3 text-slate-400" />}
                          {communeName}
                        </td>

                        {/* Sector Values */}
                        {visibleSectors.map((sector) => {
                          return (
                            <React.Fragment key={`val-${sector}`}>
                              {pivotAnalysis.indicators.map((ind, indIdx) => {
                                const val = row[`${sector} - Tổng ${ind}`];
                                const isHl = val > 0;
                                return (
                                  <td 
                                    key={ind} 
                                    className={`p-2.5 text-right border-r whitespace-nowrap ${isHl ? "bg-emerald-50/20 font-bold" : ""}`}
                                    style={{ borderColor: indIdx === pivotAnalysis.indicators.length - 1 ? '#e2e8f0' : '#f1f5f9' }}
                                  >
                                    {formatCellValue(val, true)}
                                  </td>
                                );
                              })}
                            </React.Fragment>
                          );
                        })}

                        {/* Summary Columns */}
                        <td className="p-2.5 text-right font-semibold text-indigo-700 border-r border-slate-200 bg-teal-50/20 whitespace-nowrap">
                          {formatCellValue(row["Số_Dòng_Tổng_Hợp"] ?? row["Số_DN_Địa_Phương"], false)}
                        </td>
                        {pivotAnalysis.indicators.map((ind, indIdx) => {
                          const totVal = row[`Tổng_Cộng_${ind}_Toàn_Xã`] ?? row[`Tổng_Cộng_Toàn_Xã_${ind}`] ?? row["Tổng_Doanh_Thu_Địa_Phương"];
                          return (
                            <td 
                              key={`overall-${ind}`} 
                              className="p-2.5 text-right font-bold text-teal-800 border-r border-slate-200 bg-teal-50/30 whitespace-nowrap font-mono"
                              style={{ borderRightColor: indIdx === pivotAnalysis.indicators.length - 1 ? '#e2e8f0' : '#f1f5f9' }}
                            >
                              {formatCellValue(totVal, true)}
                            </td>
                          );
                        })}
                      </tr>

                      {/* EXPANDED PANEL VIEW FOR THIS ROW */}
                      {isSelected && (
                        <tr>
                          <td colSpan={visibleSectors.length * pivotAnalysis.indicators.length + 2 + pivotAnalysis.indicators.length} className="bg-slate-50 p-4 border-y border-amber-200/50 animate-fade-in">
                            <div className="max-w-4xl mx-auto space-y-3 font-sans text-slate-800">
                              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                                <span className="text-xs font-bold text-amber-800 uppercase tracking-wider flex items-center gap-1.5">
                                  <CheckCircle className="w-4 h-4 text-emerald-600" /> phân tích tỉ trọng địa bàn: {communeName}
                                </span>
                                <span className="text-[10px] text-slate-500 font-mono">Click vào hàng xã một lần nữa để thu gọn</span>
                              </div>

                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                {/* Left stats list */}
                                <div className="space-y-2">
                                  <div className="bg-white p-3 rounded-lg border border-slate-200 space-y-3">
                                    <h5 className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Cơ cấu đóng góp của Xã vào tổng thể</h5>
                                    
                                    <div className="space-y-3">
                                      {/* Record count percentage */}
                                      <div className="space-y-1">
                                        <div className="flex justify-between text-xs text-slate-700">
                                          <span>🏢 Số dòng mẫu:</span>
                                          <strong className="text-cyan-700 font-mono">
                                            {(row["Số_Dòng_Tổng_Hợp"] ?? 0).toLocaleString()} / {overallTotals.dn.toLocaleString()} dòng ({overallTotals.dn > 0 ? (((row["Số_Dòng_Tổng_Hợp"] ?? 0) / overallTotals.dn) * 100).toFixed(2) : 0}%)
                                          </strong>
                                        </div>
                                        <div className="bg-slate-100 h-1 rounded overflow-hidden">
                                          <div className="bg-cyan-500 h-full text-xs" style={{ width: `${overallTotals.dn > 0 ? (((row["Số_Dòng_Tổng_Hợp"] ?? 0) / overallTotals.dn) * 100) : 0}%` }} />
                                        </div>
                                      </div>

                                      {/* Dynamic indicators percentage */}
                                      {pivotAnalysis.indicators.map((ind) => {
                                        const totVal = row[`Tổng_Cộng_${ind}_Toàn_Xã`] ?? row[`Tổng_Cộng_Toàn_Xã_${ind}`] ?? 0;
                                        const overallVal = overallTotals.totalsByIndicator[ind] || 1;
                                        const pct = (totVal / overallVal) * 100;
                                        return (
                                          <div key={ind} className="space-y-1">
                                            <div className="flex justify-between text-xs text-slate-700">
                                              <span>📊 Chỉ tiêu {ind}:</span>
                                              <strong className="text-emerald-700 font-mono">
                                                {totVal.toLocaleString("vi-VN")} ({pct.toFixed(2)}%)
                                              </strong>
                                            </div>
                                            <div className="bg-slate-100 h-1 rounded overflow-hidden">
                                              <div className="bg-emerald-600 h-full text-xs" style={{ width: `${pct}%` }} />
                                            </div>
                                          </div>
                                        );
                                      })}
                                    </div>
                                  </div>
                                </div>

                                {/* Right: Detailed actual list of sectors with value > 0 */}
                                <div className="space-y-1.5 bg-white p-3 rounded-lg border border-slate-200 max-h-[190px] overflow-y-auto">
                                  <h5 className="text-[11px] font-bold text-amber-800 uppercase tracking-wider">Cơ cấu đóng góp của ngành trong địa bàn</h5>
                                  <div className="divide-y divide-slate-100">
                                    {pivotAnalysis.sectors.map((sec) => {
                                      const hasActiveIndicator = pivotAnalysis.indicators.some(
                                        (ind) => (row[`${sec} - Tổng ${ind}`] || 0) > 0
                                      );
                                      if (!hasActiveIndicator) return null;

                                      return (
                                        <div key={sec} className="py-2 flex flex-col text-xs gap-1.5">
                                          <span className="text-slate-900 font-semibold leading-tight text-[11px] font-mono text-purple-700">
                                            🏢 NGÀNH: {sec}
                                          </span>
                                          <div className={`grid grid-cols-2 gap-3 text-[10.5px]`}>
                                            {pivotAnalysis.indicators.map((ind) => {
                                              const val = row[`${sec} - Tổng ${ind}`] || 0;
                                              const communeTotal = row[`Tổng_Cộng_${ind}_Toàn_Xã`] ?? row[`Tổng_Cộng_Toàn_Xã_${ind}`] ?? 1;
                                              const overallTotal = overallTotals.totalsByIndicator[ind] || 1;
                                              const propOfCommune = (val / communeTotal) * 100;
                                              const propOfGrand = (val / overallTotal) * 100;
                                              return (
                                                <div key={ind} className="bg-slate-50 p-1.5 rounded border border-slate-200 font-mono">
                                                  <span className="text-slate-500 block text-[9px] uppercase">{ind}:</span>
                                                  <strong className="text-emerald-700 font-bold block">{val.toLocaleString("vi-VN")}</strong>
                                                  <span className="text-slate-450 block text-[8px] mt-0.5">
                                                    Nhóm: {propOfCommune.toFixed(1)}% xã / {propOfGrand.toFixed(2)}% tổng
                                                  </span>
                                                </div>
                                              );
                                            })}
                                          </div>
                                        </div>
                                      );
                                    })}
                                    {/* Safeguard if no sector is active */}
                                    {pivotAnalysis.sectors.every(sec => pivotAnalysis.indicators.every(ind => (row[`${sec} - Tổng ${ind}`] || 0) === 0)) && (
                                      <p className="text-[11px] text-slate-400 italic py-2">Không ghi nhận hoạt động phát sinh.</p>
                                    )}
                                  </div>
                                </div>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>

              {/* TFOOT FOR PIVOT */}
              <tfoot>
                {/* DÒNG 1: TỔNG CỘNG TRỊ SỐ LŨY KẾ TOÀN BỘ PHÁT SINH */}
                <tr className="bg-slate-100 border-t border-slate-300 text-slate-900 font-bold text-[11px] font-mono sticky bottom-0 z-20 uppercase">
                  <td className="p-3 text-center bg-slate-100 border-r border-slate-200 sticky left-0 z-20 shadow-lg">
                    TỔNG CỘNG TOÀN HUYỆN
                  </td>
                  {visibleSectors.map((sector) => {
                    return (
                      <React.Fragment key={`tot-${sector}`}>
                        {pivotAnalysis.indicators.map((ind) => {
                          const sumVal = overallTotals.sectorStats[sector]?.[ind] || 0;
                          return (
                            <td key={ind} className="p-2.5 text-right border-r border-slate-200 bg-emerald-50/35 text-emerald-800 font-bold font-mono">
                              {sumVal > 0 ? sumVal.toLocaleString("vi-VN") : "—"}
                            </td>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                  <td className="p-2.5 text-right font-bold text-teal-800 border-r border-slate-200 bg-teal-50/40">
                    {overallTotals.dn.toLocaleString("vi-VN")}
                  </td>
                  {pivotAnalysis.indicators.map((ind) => {
                    const grandVal = overallTotals.totalsByIndicator[ind] || 0;
                    return (
                      <td key={`grand-tot-${ind}`} className="p-2.5 text-right font-extrabold text-teal-950 border-r border-slate-200 bg-teal-50/50 font-mono">
                        {grandVal.toLocaleString("vi-VN")}
                      </td>
                    );
                  })}
                </tr>

                {/* DÒNG 2: TỈ TRỌNG (%) CỦA TỪNG NGÀNH TRÊN TỔNG THỂ */}
                <tr className="bg-slate-50 text-amber-800 font-bold text-[10px] font-mono uppercase border-t border-slate-200">
                  <td className="p-3 text-center bg-slate-50 border-r border-slate-200 sticky left-0 z-20 shadow-lg">
                    TỈ TRỌNG CƠ CẤU (%)
                  </td>
                  {visibleSectors.map((sector) => {
                    return (
                      <React.Fragment key={`pct-${sector}`}>
                        {pivotAnalysis.indicators.map((ind) => {
                          const val = overallTotals.sectorStats[sector]?.[ind] || 0;
                          const overallTotal = overallTotals.totalsByIndicator[ind] || 0;
                          const pct = overallTotal > 0 ? (val / overallTotal) * 100 : 0;
                          return (
                            <td key={ind} className="p-2.5 text-right border-r border-slate-200 bg-amber-50/15 text-amber-700 font-semibold">
                              {pct > 0 ? `${pct.toFixed(2)}%` : "—"}
                            </td>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                  <td className="p-2.5 text-right border-r border-slate-200 bg-teal-50/20 text-slate-500 font-normal">
                    100.0%
                  </td>
                  {pivotAnalysis.indicators.map((ind) => (
                    <td key={`pct-grand-${ind}`} className="p-2.5 text-right border-r border-slate-200 bg-teal-50/25 text-slate-500 font-normal font-mono">
                      100.0%
                    </td>
                  ))}
                </tr>
              </tfoot>
            </table>
          </div>
        ) : (
          /* ================== STANDARD FLAT TABLE LAYOUT ================== */
          <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white max-h-[500px]">
            <table className="w-full text-left text-xs border-collapse font-sans">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-200 text-slate-800 font-mono text-[11px] sticky top-0 z-10 uppercase">
                  <th className="p-3 w-12 text-center text-slate-400 font-sans">STT</th>
                  {cols.map((col) => (
                    <th key={col} className="p-3 font-semibold whitespace-nowrap">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-slate-800 text-[11.5px]">
                {filteredRows.map((row, rIdx) => (
                  <tr key={rIdx} className={`hover:bg-slate-50 transition-colors ${rIdx % 2 === 1 ? "bg-slate-50/30" : ""}`}>
                    <td className="p-3 text-center text-slate-400 font-mono">{rIdx + 1}</td>
                    {cols.map((col) => {
                      const val = row[col];
                      const isNumeric = typeof val === "number";
                      let proportionStr = "";

                      if (isNumeric) {
                        if (col.startsWith("Tổng_")) {
                          const indName = col.replace("Tổng_", "");
                          const overallVal = overallTotals.totalsByIndicator[indName] || 0;
                          if (overallVal > 0) {
                            proportionStr = ` (${((val / overallVal) * 100).toFixed(2)}%)`;
                          }
                        } else if (col === "Số_Lượng_Bản_Ghi") {
                          const overallVal = overallTotals.dn || 0;
                          if (overallVal > 0) {
                            proportionStr = ` (${((val / overallVal) * 100).toFixed(2)}%)`;
                          }
                        }
                      }

                      return (
                        <td
                          key={col}
                          className={`p-3 whitespace-nowrap ${
                            isNumeric
                              ? "font-mono text-emerald-700 font-bold text-right"
                              : col === "Địa_Bàn_Xã" || col === "Địa_bàn_Xã"
                              ? "font-semibold text-slate-900"
                              : ""
                          }`}
                        >
                          {isNumeric ? (
                            <>
                              {val.toLocaleString("vi-VN")}
                              {proportionStr && <span className="text-[10px] text-slate-500 font-normal ml-1 bg-slate-50 border border-slate-200 px-1.2 py-0.5 rounded">{proportionStr}</span>}
                            </>
                          ) : String(val ?? "")}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="bg-slate-100 border-t-2 border-slate-300 text-slate-900 font-bold text-[11.5px] font-mono uppercase">
                  <td className="p-3 text-center text-slate-400">∑</td>
                  {cols.map((col, cIdx) => {
                    if (cIdx === 0) {
                      return <td key={col} colSpan={2} className="p-3 text-left">TỔNG CỘNG TOÀN BẢNG</td>;
                    }
                    if (col === "Địa_Bàn_Xã" || col === "Địa_bàn_Xã") {
                      return null;
                    }
                    if (col === "Số_Lượng_Bản_Ghi") {
                      return (
                        <td key={col} className="p-3 text-right font-semibold text-teal-800 bg-teal-50/20">
                          {overallTotals.dn.toLocaleString("vi-VN")} (100%)
                        </td>
                      );
                    }
                    if (col.startsWith("Tổng_")) {
                      const ind = col.replace("Tổng_", "");
                      const totalVal = overallTotals.totalsByIndicator[ind] || 0;
                      return (
                        <td key={col} className="p-3 text-right font-bold text-teal-900 bg-teal-50/30 font-mono">
                          {totalVal.toLocaleString("vi-VN")} (100%)
                        </td>
                      );
                    }
                    return <td key={col} className="p-3 bg-slate-50" />;
                  })}
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>
    </div>
  );
});

BeautifulReportTable.displayName = "BeautifulReportTable";
