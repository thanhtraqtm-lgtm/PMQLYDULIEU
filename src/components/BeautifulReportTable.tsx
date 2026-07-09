import React, { useState, useMemo, useEffect } from "react";
import { Search, Eye, EyeOff, Download, ChevronRight, ChevronDown, CheckCircle, Settings, Edit, Columns } from "lucide-react";

export function formatNumberVi(val: any): string {
  const num = typeof val === "number" ? val : parseFloat(String(val).replace(/[^0-9.\-eE+]/g, "")) || 0;
  if (isNaN(num)) return "0";
  if (!isFinite(num)) return "∞";
  
  // Nếu là số cực lớn (lớn hơn 10^12), hiển thị dạng số khoa học kèm cảnh báo để phát hiện MST/SĐT bị cộng nhầm
  if (Math.abs(num) > 1e12) {
    return num.toExponential(4) + " (⚠️ Số cực lớn)";
  }
  
  return num.toLocaleString("vi-VN", { maximumFractionDigits: 2 });
}

interface BeautifulReportTableProps {
  rows: any[];
  cols: string[];
  level: number;
  reportType: "pivot" | "flat";
  onExport: (exportRows?: any[], exportCols?: string[]) => void;
}

export const BeautifulReportTable = React.memo<BeautifulReportTableProps>(({
  rows,
  cols,
  level,
  reportType,
  onExport,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [hideZeroColumns, setHideZeroColumns] = useState(false);
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);
  
  // State for editable rows
  const [localRows, setLocalRows] = useState<any[]>([]);

  // Synchronize localRows with incoming rows prop
  useEffect(() => {
    setLocalRows(JSON.parse(JSON.stringify(rows)));
  }, [rows]);

  // Tự động phát hiện cột địa bàn (Xã/Khu vực), cột nhóm phân loại và cột số lượng bản ghi
  // để loại bỏ hoàn toàn việc gán cứng tên cột (giúp hoạt động với bất kỳ bộ dữ liệu nào)
  const communeCol = useMemo(() => {
    const standardKeys = ["Địa_Bàn_Xã", "Địa_bàn_Xã", "Xã", "Địa bàn", "Commune", "Phường", "Huyện", "Tỉnh", "Khu vực", "Area", "District"];
    for (const key of standardKeys) {
      const found = cols.find(c => c.toLowerCase() === key.toLowerCase());
      if (found) return found;
    }
    const sectorKey = level === 0 ? "Nhóm_Phân_Loại" : (level === 6 ? "Nhóm_Ngành_Chính" : `Ngành_Cấp_${level}`);
    const foundNonNumeric = cols.find(col => {
      if (col === sectorKey || col.toLowerCase().includes("ngành") || col.toLowerCase().includes("vsic")) return false;
      if (col.includes(" - Tổng ") || col.startsWith("Tổng_Cộng_") || col.startsWith("Tổng_")) return false;
      if (/số.*dòng|số.*mẫu|số.*cơ.*sở|bản.*ghi|record|count|số_dòng|số_dn|số_lượng_bản_ghi|Số_Dòng_Tổng_Hợp|Số_Lượng_Bản_Ghi/i.test(col)) return false;
      return true;
    });
    return foundNonNumeric || cols[0] || "Địa_Bàn_Xã";
  }, [cols, level]);

  const sectorKey = useMemo(() => {
    const standardKeys = [
      level === 0 ? "Nhóm_Phân_Loại" : (level === 6 ? "Nhóm_Ngành_Chính" : `Ngành_Cấp_${level}`),
      "Nhóm_Phân_Loại", "Nhóm_Ngành_Chính", `Ngành_Cấp_${level}`,
      "Ngành_Cấp_1", "Ngành_Cấp_2", "Ngành_Cấp_3", "Ngành_Cấp_4", "Ngành_Cấp_5", "Nhóm"
    ];
    for (const key of standardKeys) {
      if (cols.includes(key)) return key;
    }
    const found = cols.find(col => /ngành|phân.*loại|nhóm|sector|group|category/i.test(col));
    return found || cols[1] || cols[0] || "Nhóm_Phân_Loại";
  }, [cols, level]);

  const countCol = useMemo(() => {
    const standardKeys = ["Số_Dòng_Tổng_Hợp", "Số_Lượng_Bản_Ghi", "Số lượng dòng", "Số cơ sở", "Số_DN_Địa_Phương", "Số_Lượng_Doanh_Nghiệp"];
    for (const key of standardKeys) {
      if (cols.includes(key)) return key;
    }
    const found = cols.find(col => /số.*dòng|số.*mẫu|số.*cơ.*sở|bản.*ghi|record|count|số_dòng|số_dn|số_lượng_bản_ghi/i.test(col));
    return found || cols.find(col => !col.includes(" - Tổng ") && (col.toLowerCase().includes("số") || col.toLowerCase().includes("count"))) || "Số lượng dòng";
  }, [cols]);

  // Extract sectors and indicators dynamically from column list based on localRows
  const pivotAnalysis = useMemo(() => {
    if (reportType !== "pivot" || localRows.length === 0) {
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
      return localRows.some((row) => {
        return indicators.some((ind) => {
          const val = row[`${sector} - Tổng ${ind}`] || 0;
          return val !== 0;
        });
      });
    });

    return { sectors, indicators, activeSectors };
  }, [localRows, cols, reportType]);

  // Sector selection state for "Chỉ để lại các cột cần thiết" (Column Visibility Manager)
  const [selectedSectors, setSelectedSectors] = useState<string[]>([]);
  const [selectedFlatCols, setSelectedFlatCols] = useState<string[]>([]);
  const [showColumnSelector, setShowColumnSelector] = useState(false);

  // Initialize selected columns when pivotAnalysis sectors or flat cols change
  useEffect(() => {
    if (pivotAnalysis.sectors.length > 0) {
      setSelectedSectors(pivotAnalysis.sectors);
    }
  }, [pivotAnalysis.sectors]);

  useEffect(() => {
    setSelectedFlatCols(cols);
  }, [cols]);

  // Columns to display based on toggle & visibility checklist
  const visibleSectors = useMemo(() => {
    if (reportType !== "pivot") return [];
    const baseList = hideZeroColumns ? pivotAnalysis.activeSectors : pivotAnalysis.sectors;
    // Filter baseList to only keep the selected/necessary columns
    const filtered = baseList.filter(s => selectedSectors.includes(s));
    if (filtered.length === 0 && selectedSectors.length > 0) return selectedSectors;
    if (filtered.length === 0) return baseList; // Fallback
    return filtered;
  }, [pivotAnalysis, hideZeroColumns, selectedSectors, reportType]);

  const visibleFlatCols = useMemo(() => {
    if (reportType !== "flat") return [];
    return cols.filter(c => selectedFlatCols.includes(c));
  }, [cols, selectedFlatCols, reportType]);

  // Filter localRows based on search term
  const filteredRows = useMemo(() => {
    if (!searchTerm.trim()) return localRows;
    const term = searchTerm.toLowerCase();
    return localRows.filter((row) => {
      const commune = String(row[communeCol] || "").toLowerCase();
      const sector = String(row[sectorKey] || "").toLowerCase();
      return commune.includes(term) || sector.includes(term);
    });
  }, [localRows, searchTerm, communeCol, sectorKey]);

  // Calculate dynamic totals across all columns
  const overallTotals = useMemo(() => {
    let grandDN = 0;
    const totalsByIndicator: { [ind: string]: number } = {};
    const sectorStats: { [sector: string]: { [ind: string]: number } } = {};
    
    localRows.forEach((row) => {
      if (reportType === "pivot") {
        const dnRaw = row[countCol] ?? 0;
        const dnVal = typeof dnRaw === "number" ? dnRaw : parseFloat(String(dnRaw).replace(/[^0-9.\-]/g, "")) || 0;
        grandDN += isNaN(dnVal) ? 0 : dnVal;
        
        pivotAnalysis.indicators.forEach((ind) => {
          if (totalsByIndicator[ind] === undefined) totalsByIndicator[ind] = 0;
          const key = row[`Tổng_Cộng_${ind}_Toàn_Xã`] !== undefined 
            ? `Tổng_Cộng_${ind}_Toàn_Xã` 
            : (row[`Tổng_Cộng_Toàn_Xã_${ind}`] !== undefined 
               ? `Tổng_Cộng_Toàn_Xã_${ind}` 
               : cols.find(c => c.includes(ind) && (c.startsWith("Tổng_Cộng_") || c.endsWith("_Toàn_Xã"))) || `Tổng_Cộng_${ind}_Toàn_Xã`);
               
          const rawVal = row[key] ?? 0;
          const val = typeof rawVal === "number" ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.\-]/g, "")) || 0;
          totalsByIndicator[ind] += isNaN(val) ? 0 : val;
        });

        pivotAnalysis.sectors.forEach((sector) => {
          if (!sectorStats[sector]) {
            sectorStats[sector] = {};
          }
          pivotAnalysis.indicators.forEach((ind) => {
            if (sectorStats[sector][ind] === undefined) sectorStats[sector][ind] = 0;
            const rawVal = row[`${sector} - Tổng ${ind}`] ?? 0;
            const val = typeof rawVal === "number" ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.\-]/g, "")) || 0;
            sectorStats[sector][ind] += isNaN(val) ? 0 : val;
          });
        });
      } else {
        // Flat mode
        const dnRaw = row[countCol] ?? 0;
        const dnVal = typeof dnRaw === "number" ? dnRaw : parseFloat(String(dnRaw).replace(/[^0-9.\-]/g, "")) || 0;
        grandDN += isNaN(dnVal) ? 0 : dnVal;
        cols.forEach((col) => {
          if (col.startsWith("Tổng_")) {
            const ind = col.replace("Tổng_", "");
            if (totalsByIndicator[ind] === undefined) totalsByIndicator[ind] = 0;
            const rawVal = row[col] ?? 0;
            const val = typeof rawVal === "number" ? rawVal : parseFloat(String(rawVal).replace(/[^0-9.\-]/g, "")) || 0;
            totalsByIndicator[ind] += isNaN(val) ? 0 : val;
          }
        });
      }
    });

    return {
      dn: grandDN,
      totalsByIndicator,
      sectorStats
    };
  }, [localRows, reportType, pivotAnalysis.sectors, pivotAnalysis.indicators, cols, countCol]);

  // Handle direct numeric input cell updates with instant calculations
  const handleCellChange = (rIdx: number, field: string, newValue: string) => {
    const cleanNumStr = newValue.replace(/[^0-9.\-]/g, "");
    const numericVal = cleanNumStr === "" ? 0 : parseFloat(cleanNumStr) || 0;
    
    setLocalRows(prev => {
      const updated = [...prev];
      const row = { ...updated[rIdx] };
      row[field] = numericVal;
      
      // Auto recalculate row summary indicators in pivot mode
      if (reportType === "pivot") {
        pivotAnalysis.indicators.forEach(ind => {
          if (field.includes(` - Tổng ${ind}`)) {
            let sum = 0;
            pivotAnalysis.sectors.forEach(sec => {
              const key = `${sec} - Tổng ${ind}`;
              const rawV = key === field ? numericVal : (row[key] || 0);
              const v = typeof rawV === "number" ? rawV : parseFloat(String(rawV).replace(/[^0-9.\-]/g, "")) || 0;
              sum += isNaN(v) ? 0 : v;
            });
            const keyToUpdate = row[`Tổng_Cộng_${ind}_Toàn_Xã`] !== undefined 
              ? `Tổng_Cộng_${ind}_Toàn_Xã` 
              : (row[`Tổng_Cộng_Toàn_Xã_${ind}`] !== undefined 
                 ? `Tổng_Cộng_Toàn_Xã_${ind}` 
                 : cols.find(c => c.includes(ind) && (c.startsWith("Tổng_Cộng_") || c.endsWith("_Toàn_Xã"))) || `Tổng_Cộng_${ind}_Toàn_Xã`);
            row[keyToUpdate] = Math.round(sum * 100) / 100;
          }
        });
      }
      
      updated[rIdx] = row;
      return updated;
    });
  };

  // Handle text edits (e.g. commune name)
  const handleCellChangeText = (rIdx: number, field: string, newValue: string) => {
    setLocalRows(prev => {
      const updated = [...prev];
      const row = { ...updated[rIdx] };
      row[field] = newValue;
      updated[rIdx] = row;
      return updated;
    });
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden space-y-4 p-5" id="beautiful_report_container">
      {/* HEADER BAR */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200 font-sans">
        <div>
          <span className="text-xs font-black text-slate-800 font-mono tracking-widest uppercase block mb-1">
            📊 KẾT QUẢ TỔNG HỢP {level === 0 ? "PHÂN NHÓM" : (level === 6 ? "NHÓM NGÀNH CHÍNH" : `DANH MỤC NGÀNH CẤP ${level}`)} × ĐỊA BÀN
          </span>
          <span className="text-xs text-slate-600 block leading-normal">
            Nhấp chuột trực tiếp vào bất kỳ ô số liệu nào để <strong>chỉnh sửa số lại</strong>. Hệ thống sẽ tự động cộng dồn và cập nhật biểu mẫu toàn bảng.
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

          {/* CHOOSE NECESSARY COLUMNS MANAGER */}
          <button
            onClick={() => setShowColumnSelector(!showColumnSelector)}
            className={`text-xs font-bold py-1.5 px-3 rounded-lg border transition-all flex items-center gap-1.5 select-none md:w-auto w-full justify-center cursor-pointer ${
              showColumnSelector
                ? "bg-indigo-600 text-white border-indigo-600 shadow-sm"
                : "bg-white text-slate-700 border-slate-300 hover:bg-slate-50 shadow-sm"
            }`}
            title="Nhấn để chọn và chỉ giữ lại các cột cần thiết"
          >
            <Columns className="w-3.5 h-3.5" />
            {showColumnSelector ? "Đóng bộ lọc cột" : "⚙️ Chỉ giữ lại cột cần thiết"}
          </button>

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
              {hideZeroColumns ? `Đang ẩn cột rỗng` : "Hiện tất cả cột ngành"}
            </button>
          )}

          {/* EXCEL EXPORT BUTTON WITH CLEAN FILTERED DATA */}
          <button
            onClick={() => {
              const exportRowsClean = localRows.map(row => {
                const cleanRow: any = {};
                if (reportType === "pivot") {
                  cleanRow[communeCol] = row[communeCol] || "";
                  visibleSectors.forEach(sector => {
                    pivotAnalysis.indicators.forEach(ind => {
                      cleanRow[`${sector} - Tổng ${ind}`] = row[`${sector} - Tổng ${ind}`] || 0;
                    });
                  });
                  cleanRow[countCol] = row[countCol] ?? 0;
                  pivotAnalysis.indicators.forEach(ind => {
                    const key = row[`Tổng_Cộng_${ind}_Toàn_Xã`] !== undefined 
                      ? `Tổng_Cộng_${ind}_Toàn_Xã` 
                      : (row[`Tổng_Cộng_Toàn_Xã_${ind}`] !== undefined 
                         ? `Tổng_Cộng_Toàn_Xã_${ind}` 
                         : cols.find(c => c.includes(ind) && (c.startsWith("Tổng_Cộng_") || c.endsWith("_Toàn_Xã"))) || `Tổng_Cộng_${ind}_Toàn_Xã`);
                    cleanRow[key] = row[key] || 0;
                  });
                } else {
                  visibleFlatCols.forEach(col => {
                    cleanRow[col] = row[col];
                  });
                }
                return cleanRow;
              });
              onExport(exportRowsClean, reportType === "pivot" ? Object.keys(exportRowsClean[0] || {}) : visibleFlatCols);
            }}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-md inline-flex justify-center border-0"
          >
            <Download className="w-3.5 h-3.5" /> Xuất Excel đã sửa (.xlsx)
          </button>
        </div>
      </div>

      {/* COLUMN visibility CONTROLLER */}
      {showColumnSelector && (
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 font-sans animate-fade-in text-slate-800">
          <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-2 pb-2 border-b border-slate-200">
            <div>
              <span className="text-xs font-bold text-indigo-700 uppercase tracking-wider font-mono block">
                ⚙️ Bộ chọn cột hiển thị (Chỉ lưu lại các cột cần thiết)
              </span>
              <span className="text-[10.5px] text-slate-500 block">Tích chọn các cột/ngành cần thiết để giữ lại trên bảng hiển thị và tệp xuất Excel.</span>
            </div>
            <div className="flex gap-2 self-start sm:self-auto">
              <button
                onClick={() => {
                  if (reportType === "pivot") {
                    setSelectedSectors(pivotAnalysis.sectors);
                  } else {
                    setSelectedFlatCols(cols);
                  }
                }}
                className="bg-white border border-slate-300 text-slate-700 font-bold text-[10px] px-2.5 py-1 rounded hover:bg-slate-100 cursor-pointer"
              >
                Chọn tất cả
              </button>
              <button
                onClick={() => {
                  if (reportType === "pivot") {
                    setSelectedSectors([]);
                  } else {
                    setSelectedFlatCols([]);
                  }
                }}
                className="bg-white border border-slate-300 text-slate-500 font-bold text-[10px] px-2.5 py-1 rounded hover:bg-slate-100 cursor-pointer"
              >
                Bỏ chọn hết
              </button>
            </div>
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2 max-h-[140px] overflow-y-auto pr-1">
            {reportType === "pivot" ? (
              pivotAnalysis.sectors.map((sec) => {
                const isChecked = selectedSectors.includes(sec);
                return (
                  <label 
                    key={sec} 
                    className={`flex items-center gap-1.5 p-1.5 rounded border text-[11px] cursor-pointer transition-all select-none ${
                      isChecked 
                        ? "bg-indigo-50/70 border-indigo-200 text-indigo-900 font-bold" 
                        : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => {
                        if (isChecked) {
                          setSelectedSectors(prev => prev.filter(s => s !== sec));
                        } else {
                          setSelectedSectors(prev => [...prev, sec]);
                        }
                      }}
                      className="rounded text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer"
                    />
                    <span className="truncate" title={sec}>{sec}</span>
                  </label>
                );
              })
            ) : (
              cols.map((col) => {
                const isChecked = selectedFlatCols.includes(col);
                return (
                  <label 
                    key={col} 
                    className={`flex items-center gap-1.5 p-1.5 rounded border text-[11px] cursor-pointer transition-all select-none ${
                      isChecked 
                        ? "bg-indigo-50/70 border-indigo-200 text-indigo-900 font-bold" 
                        : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50"
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => {
                        if (isChecked) {
                          setSelectedFlatCols(prev => prev.filter(c => c !== col));
                        } else {
                          setSelectedFlatCols(prev => [...prev, col]);
                        }
                      }}
                      className="rounded text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer"
                    />
                    <span className="truncate" title={col}>{col}</span>
                  </label>
                );
              })
            )}
          </div>
        </div>
      )}

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
                {formatNumberVi(totalVal)} <span className="text-xs font-normal text-slate-500 font-sans">tổng cộng</span>
              </div>
              <div className="text-[10px] text-slate-500">Tổng quy nạp lũy kế trên toàn bộ {localRows.length} địa bàn</div>
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
            {formatNumberVi(overallTotals.dn)} <span className="text-xs font-normal text-slate-500 font-sans">cơ sở/dòng d.liệu</span>
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
                  <th rowSpan={2} className="p-3 font-bold whitespace-nowrap bg-slate-100 border-r border-slate-200 align-middle text-center sticky left-0 z-30 capitalize">
                    {communeCol.replace(/_/g, " ")}
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
                  <th className="p-2 font-semibold text-right border-r border-slate-200 bg-teal-50/40 text-teal-800 whitespace-nowrap w-20 capitalize">
                    {countCol.replace(/_/g, " ")}
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
                  const communeName = row[communeCol] || "Khác";
                  return (
                    <React.Fragment key={`row-${rIdx}`}>
                      <tr
                        className={`hover:bg-slate-50 transition-colors ${
                          isSelected ? "bg-amber-50/60 font-semibold border-l-4 border-amber-500" : rIdx % 2 === 1 ? "bg-slate-50/40" : ""
                        }`}
                      >
                        {/* Commune Name (Editable text) */}
                        <td className="p-1 whitespace-nowrap font-sans font-medium text-slate-900 bg-white border-r border-slate-200 sticky left-0 z-10 flex items-center gap-1 shadow-sm">
                          <button
                            type="button"
                            onClick={() => setSelectedRowIndex(isSelected ? null : rIdx)}
                            className="text-slate-400 hover:text-slate-600 p-1 rounded hover:bg-slate-100 transition-all shrink-0 cursor-pointer"
                          >
                            {isSelected ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
                          </button>
                          <input
                            type="text"
                            value={communeName}
                            onChange={(e) => handleCellChangeText(rIdx, communeCol, e.target.value)}
                            className="bg-transparent font-medium text-slate-900 outline-none focus:bg-amber-100/80 focus:ring-1 focus:ring-amber-500 rounded px-1.5 py-1 text-xs border-0 w-28 font-sans"
                            title="Nhấp đúp chuột để đổi tên địa bàn"
                          />
                        </td>

                        {/* Sector Values (Directly Editable inputs) */}
                        {visibleSectors.map((sector) => {
                          return (
                            <React.Fragment key={`val-${sector}`}>
                              {pivotAnalysis.indicators.map((ind, indIdx) => {
                                const val = row[`${sector} - Tổng ${ind}`];
                                const isHl = val > 0;
                                return (
                                  <td 
                                    key={ind} 
                                    className={`p-1 text-right border-r whitespace-nowrap ${isHl ? "bg-emerald-50/10" : ""}`}
                                    style={{ borderColor: indIdx === pivotAnalysis.indicators.length - 1 ? '#e2e8f0' : '#f1f5f9' }}
                                  >
                                    <input
                                      type="text"
                                      value={val !== undefined && val !== null ? val : ""}
                                      onChange={(e) => handleCellChange(rIdx, `${sector} - Tổng ${ind}`, e.target.value)}
                                      className="w-full bg-transparent text-right font-mono font-bold text-emerald-700 outline-none focus:bg-amber-100/70 focus:ring-1 focus:ring-amber-500 rounded px-1.5 py-1 text-xs border-0 transition-all"
                                      title="Click để sửa số của ngành này"
                                    />
                                  </td>
                                );
                              })}
                            </React.Fragment>
                          );
                        })}

                        {/* Summary Columns (Editable overrides) */}
                        <td className="p-1 text-right font-semibold text-indigo-700 border-r border-slate-200 bg-teal-50/10 whitespace-nowrap">
                          <input
                            type="text"
                            value={row[countCol] !== undefined && row[countCol] !== null ? row[countCol] : ""}
                            onChange={(e) => handleCellChange(rIdx, countCol, e.target.value)}
                            className="w-full bg-transparent text-right font-sans font-bold text-indigo-700 outline-none focus:bg-amber-100/70 focus:ring-1 focus:ring-amber-500 rounded px-1.5 py-1 text-xs border-0"
                            title="Sửa số lượng mẫu địa phương"
                          />
                        </td>
                        {pivotAnalysis.indicators.map((ind, indIdx) => {
                          const key = row[`Tổng_Cộng_${ind}_Toàn_Xã`] !== undefined ? `Tổng_Cộng_${ind}_Toàn_Xã` : `Tổng_Cộng_Toàn_Xã_${ind}`;
                          const totVal = row[key];
                          return (
                            <td 
                              key={`overall-${ind}`} 
                              className="p-1 text-right font-bold text-teal-850 border-r border-slate-200 bg-teal-50/20 whitespace-nowrap font-mono"
                              style={{ borderRightColor: indIdx === pivotAnalysis.indicators.length - 1 ? '#e2e8f0' : '#f1f5f9' }}
                            >
                              <input
                                type="text"
                                value={totVal !== undefined && totVal !== null ? totVal : ""}
                                onChange={(e) => handleCellChange(rIdx, key, e.target.value)}
                                className="w-full bg-transparent text-right font-mono font-bold text-teal-900 outline-none focus:bg-amber-100/70 focus:ring-1 focus:ring-amber-500 rounded px-1.5 py-1 text-xs border-0"
                                title="Sửa tổng cộng dồn của xã"
                              />
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
                                <span className="text-[10px] text-slate-500 font-mono">Bảng cơ cấu đã được đồng bộ với bất kỳ sửa đổi số liệu nào của bạn</span>
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
                                          <span className="capitalize">🏢 {countCol.replace(/_/g, " ")}:</span>
                                          <strong className="text-cyan-700 font-mono">
                                            {formatNumberVi(row[countCol] ?? 0)} / {formatNumberVi(overallTotals.dn)} dòng ({overallTotals.dn > 0 ? (((row[countCol] ?? 0) / overallTotals.dn) * 100).toFixed(2) : 0}%)
                                          </strong>
                                        </div>
                                        <div className="bg-slate-100 h-1 rounded overflow-hidden">
                                          <div className="bg-cyan-500 h-full text-xs" style={{ width: `${overallTotals.dn > 0 ? (((row[countCol] ?? 0) / overallTotals.dn) * 100) : 0}%` }} />
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
                                                {formatNumberVi(totVal)} ({pct.toFixed(2)}%)
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
                                                  <strong className="text-emerald-700 font-bold block">{formatNumberVi(val)}</strong>
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
                              {sumVal > 0 ? formatNumberVi(sumVal) : "—"}
                            </td>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                  <td className="p-2.5 text-right font-bold text-teal-800 border-r border-slate-200 bg-teal-50/40">
                    {formatNumberVi(overallTotals.dn)}
                  </td>
                  {pivotAnalysis.indicators.map((ind) => {
                    const grandVal = overallTotals.totalsByIndicator[ind] || 0;
                    return (
                      <td key={`grand-tot-${ind}`} className="p-2.5 text-right font-extrabold text-teal-950 border-r border-slate-200 bg-teal-50/50 font-mono">
                        {formatNumberVi(grandVal)}
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
                  {visibleFlatCols.map((col) => (
                    <th key={col} className="p-3 font-semibold whitespace-nowrap">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-slate-800 text-[11.5px]">
                {filteredRows.map((row, rIdx) => (
                  <tr key={rIdx} className={`hover:bg-slate-50 transition-colors ${rIdx % 2 === 1 ? "bg-slate-50/30" : ""}`}>
                    <td className="p-3 text-center text-slate-400 font-mono">{rIdx + 1}</td>
                    {visibleFlatCols.map((col) => {
                      const val = row[col];
                      const isNumeric = typeof val === "number" || !isNaN(Number(String(val).replace(/[^0-9.\-]/g, "")));
                      let proportionStr = "";

                      if (isNumeric) {
                        const parsedNum = Number(String(val).replace(/[^0-9.\-]/g, "")) || 0;
                        if (col.startsWith("Tổng_")) {
                          const indName = col.replace("Tổng_", "");
                          const overallVal = overallTotals.totalsByIndicator[indName] || 0;
                          if (overallVal > 0) {
                            proportionStr = ` (${((parsedNum / overallVal) * 105).toFixed(2)}%)`;
                          }
                        } else if (col === countCol) {
                          const overallVal = overallTotals.dn || 0;
                          if (overallVal > 0) {
                            proportionStr = ` (${((parsedNum / overallVal) * 100).toFixed(2)}%)`;
                          }
                        }
                      }

                      return (
                        <td
                          key={col}
                          className={`p-1 whitespace-nowrap ${
                            col === communeCol
                              ? "font-semibold text-slate-900 bg-white"
                              : ""
                          }`}
                        >
                          <input
                            type="text"
                            value={val !== undefined && val !== null ? val : ""}
                            onChange={(e) => {
                              if (isNumeric) {
                                handleCellChange(rIdx, col, e.target.value);
                              } else {
                                handleCellChangeText(rIdx, col, e.target.value);
                              }
                            }}
                            className={`w-full bg-transparent outline-none focus:bg-amber-100/70 focus:ring-1 focus:ring-amber-500 rounded px-1.5 py-1 text-xs border-0 ${
                              isNumeric 
                                ? "text-right font-mono text-emerald-700 font-bold" 
                                : "text-left text-slate-800 font-medium"
                            }`}
                          />
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="bg-slate-100 border-t-2 border-slate-300 text-slate-900 font-bold text-[11.5px] font-mono uppercase">
                  <td className="p-3 text-center text-slate-400">∑</td>
                  {visibleFlatCols.map((col, cIdx) => {
                    if (cIdx === 0) {
                      return <td key={col} colSpan={2} className="p-3 text-left">TỔNG CỘNG TOÀN BẢNG</td>;
                    }
                    if (col === communeCol) {
                      return null;
                    }
                    if (col === countCol) {
                      return (
                        <td key={col} className="p-3 text-right font-semibold text-teal-850 bg-teal-50/20 font-mono">
                          {formatNumberVi(overallTotals.dn)} (100%)
                        </td>
                      );
                    }
                    if (col.startsWith("Tổng_")) {
                      const ind = col.replace("Tổng_", "");
                      const totalVal = overallTotals.totalsByIndicator[ind] || 0;
                      return (
                        <td key={col} className="p-3 text-right font-bold text-teal-900 bg-teal-50/30 font-mono">
                          {formatNumberVi(totalVal)} (100%)
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
