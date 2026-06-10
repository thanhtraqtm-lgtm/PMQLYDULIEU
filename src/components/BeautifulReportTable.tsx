import React, { useState, useMemo } from "react";
import { Search, Eye, EyeOff, LayoutGrid, ListFilter, Download, ChevronRight, ChevronDown, CheckCircle } from "lucide-react";

interface BeautifulReportTableProps {
  rows: any[];
  cols: string[];
  level: number;
  reportType: "pivot" | "flat";
  onExport: () => void;
}

export const BeautifulReportTable: React.FC<BeautifulReportTableProps> = ({
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
      const sector = String(row[`Ngành_Cấp_${level}`] || "").toLowerCase();
      return commune.includes(term) || sector.includes(term);
    });
  }, [rows, searchTerm, level]);

  // If report type is Pivot, we can extract the active sectors and check if they are all zeros
  const pivotAnalysis = useMemo(() => {
    if (reportType !== "pivot" || rows.length === 0) {
      return { sectors: [], activeSectors: [] };
    }

    // Find all unique sectors in columns
    // Columns are structured as: `"${sector} - Tổng Doanh Thu"` and `"${sector} - Tổng Lao Động"`
    const sectorSet = new Set<string>();
    cols.forEach((col) => {
      if (col.includes(" - Tổng Doanh Thu")) {
        const sectorName = col.replace(" - Tổng Doanh Thu", "");
        sectorSet.add(sectorName);
      }
    });

    const sectors = Array.from(sectorSet).sort();

    // Determine which sectors have non-zero values in at least one row
    const activeSectors = sectors.filter((sector) => {
      return rows.some((row) => {
        const rev = row[`${sector} - Tổng Doanh Thu`] || 0;
        const lab = row[`${sector} - Tổng Lao Động`] || 0;
        return rev > 0 || lab > 0;
      });
    });

    return { sectors, activeSectors };
  }, [rows, cols, reportType]);

  // Columns to display based on toggles
  const visibleSectors = useMemo(() => {
    if (reportType !== "pivot") return [];
    return hideZeroColumns ? pivotAnalysis.activeSectors : pivotAnalysis.sectors;
  }, [pivotAnalysis, hideZeroColumns, reportType]);

  // Clean values for nice representation (0 -> -)
  const formatCellValue = (val: any, isNumeric: boolean) => {
    if (isNumeric) {
      if (val === undefined || val === null || val === 0) {
        return <span className="text-gray-600 font-normal">—</span>;
      }
      return <span className="font-mono text-emerald-400 font-bold">{val.toLocaleString("en-US")}</span>;
    }
    return String(val ?? "");
  };

  return (
    <div className="bg-[#111827]/80 rounded-2xl border border-gray-800 shadow-xl overflow-hidden animate-fade-in space-y-4 p-5" id="beautiful_report_container">
      {/* HEADER BAR */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-[#1f2937]/50 p-4 rounded-xl border border-gray-800">
        <div>
          <span className="text-xs font-bold text-amber-400 font-mono tracking-widest uppercase block mb-1">
            📊 KẾT QUẢ TỔNG HỢP KINH TẾ (NGÀNH CẤP {level} × ĐỊA BÀN XÃ/PHƯỜNG)
          </span>
          <span className="text-xs text-gray-400 font-sans block leading-normal">
            Bản hợp nhất từ danh mục ngành khớp trực tiếp trong bộ nhớ. Đã chạy đối soát dữ liệu thực tế.
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          {/* SEARCH BAR */}
          <div className="relative">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none">
              <Search className="h-3.5 w-3.5 text-gray-400" />
            </span>
            <input
              type="text"
              placeholder="Tìm kiếm địa bàn xã..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="bg-[#0f172a] border border-gray-700 rounded-lg pl-9 pr-3 py-1.5 text-xs text-white placeholder-gray-500 focus:outline-none focus:ring-1 focus:ring-amber-500 w-52"
            />
          </div>

          {/* HIDE EMPTY COLUMNS TOGGLE (ONLY FOR PIVOT) */}
          {reportType === "pivot" && pivotAnalysis.sectors.length > 0 && (
            <button
              onClick={() => setHideZeroColumns(!hideZeroColumns)}
              className={`text-xs font-bold py-1.5 px-3 rounded-lg border transition-all flex items-center gap-1.5 select-none md:w-auto w-full justify-center cursor-pointer ${
                hideZeroColumns
                  ? "bg-amber-950/40 text-amber-300 border-amber-800 hover:bg-amber-900/30"
                  : "bg-gray-800 text-gray-300 border-gray-750 hover:bg-gray-700"
              }`}
              title="Nhấn để ẩn/hiện các cột ngành không có số liệu phát sinh trên địa bàn"
            >
              {hideZeroColumns ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
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

      {/* DETAILED INSIGHTS / STATS CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div className="bg-gray-900/60 p-3 rounded-xl border border-gray-800">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider font-mono">Tổng Số Địa Bàn Xã</div>
          <div className="text-xl font-bold text-white mt-0.5">{rows.length} <span className="text-xs font-normal text-gray-500">địa phương</span></div>
        </div>
        <div className="bg-gray-900/60 p-3 rounded-xl border border-gray-800">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider font-mono">Phát Sinh Hoạt Động</div>
          <div className="text-xl font-bold text-emerald-400 mt-0.5">
            {reportType === "pivot" ? pivotAnalysis.activeSectors.length : "Nhiều Nhóm"} 
            <span className="text-xs font-normal text-gray-500"> nhóm ngành có phát sinh số liệu</span>
          </div>
        </div>
        <div className="bg-gray-900/60 p-3 rounded-xl border border-gray-800">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider font-mono">Ý nghĩa thiết kế</div>
          <div className="text-xs text-amber-300 mt-1 leading-normal font-sans">
            Màn hình thông minh đã được nén gọn lại. Bấm trực tiếp vào từng dòng để xem báo cáo dọc chi tiết của Xã đó.
          </div>
        </div>
      </div>

      {/* THE MAIN TABLE CORES */}
      <div className="relative">
        {filteredRows.length === 0 ? (
          <div className="text-center py-10 bg-gray-900/20 border border-gray-800 rounded-xl text-gray-400 text-xs">
            Không tìm thấy địa bàn nào khớp với từ khóa tìm kiếm.
          </div>
        ) : reportType === "pivot" ? (
          /* ================== PIVOT MATRIX LAYOUT (HIGHLY POLISHED) ================== */
          <div className="overflow-x-auto border border-gray-800 rounded-xl bg-gray-950/40 max-h-[480px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                {/* FIRST HEADER ROW: GROUPED SECTORS */}
                <tr className="bg-gray-900 border-b border-gray-800 text-gray-300 font-mono text-[10.5px] sticky top-0 z-20">
                  <th rowSpan={2} className="p-3 font-bold whitespace-nowrap bg-gray-900 border-r border-gray-800 align-middle text-center sticky left-0 z-30">
                    Địa Bàn Xã/Phường
                  </th>
                  {visibleSectors.map((sector) => (
                    <th key={sector} colSpan={2} className="p-2 font-bold whitespace-nowrap text-center border-r border-b border-gray-800 bg-[#1e293b]/70 font-sans text-[11px] text-yellow-300">
                      {sector}
                    </th>
                  ))}
                  <th colSpan={3} className="p-2 font-bold whitespace-nowrap text-center bg-teal-950/70 text-teal-300 align-middle border-b border-gray-800">
                    TỔNG HỢP ĐỊA PHƯƠNG
                  </th>
                </tr>

                {/* SECOND HEADER ROW: INDICATORS */}
                <tr className="bg-[#1f2937]/90 border-b border-gray-800 text-gray-400 font-mono text-[9px] sticky top-[31px] z-20 uppercase">
                  {visibleSectors.map((sector) => (
                    <React.Fragment key={`sub-${sector}`}>
                      <th className="p-2 font-medium text-right border-r border-gray-800/50 bg-[#1f2937]/50 whitespace-nowrap w-24">Doanh Thu</th>
                      <th className="p-2 font-medium text-right border-r border-gray-850 bg-[#1f2937]/50 whitespace-nowrap w-20">Lao Động</th>
                    </React.Fragment>
                  ))}
                  <th className="p-2 font-semibold text-right border-r border-gray-800/50 bg-teal-900/20 text-teal-400 whitespace-nowrap w-16">Số DN</th>
                  <th className="p-2 font-semibold text-right border-r border-gray-800/50 bg-teal-900/20 text-teal-400 whitespace-nowrap w-28">Tổng D.Thu</th>
                  <th className="p-2 font-semibold text-right bg-teal-900/20 text-teal-400 whitespace-nowrap w-24">Tổng L.Động</th>
                </tr>
              </thead>

              <tbody className="divide-y divide-gray-800/50 text-gray-200 font-sans text-[11.5px]">
                {filteredRows.map((row, rIdx) => {
                  const isSelected = selectedRowIndex === rIdx;
                  return (
                    <React.Fragment key={`row-${rIdx}`}>
                      <tr
                        onClick={() => setSelectedRowIndex(isSelected ? null : rIdx)}
                        className={`hover:bg-slate-800/40 transition-colors cursor-pointer ${
                          isSelected ? "bg-slate-800/60 font-semibold border-l-4 border-amber-500" : rIdx % 2 === 1 ? "bg-gray-900/10" : ""
                        }`}
                      >
                        {/* Commune Name */}
                        <td className="p-3 whitespace-nowrap font-sans font-medium text-gray-100 bg-gray-950/90 border-r border-gray-800 sticky left-0 z-10 flex items-center gap-1.5 shadow-md">
                          {isSelected ? <ChevronDown className="w-3 h-3 text-amber-400" /> : <ChevronRight className="w-3 h-3 text-gray-500" />}
                          {row["Địa_Bàn_Xã"] || row["Địa_bàn_Xã"] || "Khác"}
                        </td>

                        {/* Sector Values */}
                        {visibleSectors.map((sector) => {
                          const valRev = row[`${sector} - Tổng Doanh Thu`];
                          const valLab = row[`${sector} - Tổng Lao Động`];
                          const isHl = (valRev > 0 || valLab > 0);

                          return (
                            <React.Fragment key={`val-${sector}`}>
                              <td className={`p-2.5 text-right border-r border-gray-800/40 whitespace-nowrap ${isHl ? "bg-emerald-950/20" : ""}`}>
                                {formatCellValue(valRev, true)}
                              </td>
                              <td className={`p-2.5 text-right border-r border-gray-850 whitespace-nowrap ${isHl ? "bg-emerald-950/25" : ""}`}>
                                {formatCellValue(valLab, true)}
                              </td>
                            </React.Fragment>
                          );
                        })}

                        {/* Summary Columns */}
                        <td className="p-2.5 text-right font-semibold text-teal-300 border-r border-gray-800 bg-teal-950/20 whitespace-nowrap">
                          {row["Số_DN_Địa_Phương"] !== undefined ? row["Số_DN_Địa_Phương"].toLocaleString() : "—"}
                        </td>
                        <td className="p-2.5 text-right font-bold text-teal-400 border-r border-gray-800 bg-teal-950/25 whitespace-nowrap font-mono">
                          {row["Tổng_Doanh_Thu_Địa_Phương"] !== undefined ? row["Tổng_Doanh_Thu_Địa_Phương"].toLocaleString() : "—"}
                        </td>
                        <td className="p-2.5 text-right font-bold text-teal-400 bg-teal-950/30 whitespace-nowrap font-mono">
                          {row["Tổng_Lao_Động_Địa_Phương"] !== undefined ? row["Tổng_Lao_Động_Địa_Phương"].toLocaleString() : "—"}
                        </td>
                      </tr>

                      {/* EXPANDED PANEL VIEW FOR THIS ROW */}
                      {isSelected && (
                        <tr>
                          <td colSpan={visibleSectors.length * 2 + 4} className="bg-[#0f172a] p-4 border-y border-amber-950/45 animate-fade-in">
                            <div className="max-w-4xl mx-auto space-y-3 font-sans">
                              <div className="flex items-center justify-between border-b border-gray-800 pb-2">
                                <span className="text-xs font-bold text-amber-400 uppercase tracking-wider flex items-center gap-1.5">
                                  <CheckCircle className="w-4 h-4 text-emerald-400" /> BÁO CÁO PHÂN TÍCH THEO NGÀNH CỦA ĐỊA BÀN: {row["Địa_Bàn_Xã"] || "Khác"}
                                </span>
                                <span className="text-[10px] text-gray-500 font-mono">Ấn dòng tiêu đề lần nữa để thu gọn</span>
                              </div>

                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                {/* Left stats list */}
                                <div className="space-y-2">
                                  <div className="bg-[#1e293b]/70 p-3 rounded-lg border border-gray-800 space-y-1.5">
                                    <h5 className="text-[11px] font-bold text-gray-400 uppercase tracking-wider">Số liệu tổng kết sơ bộ tại địa bàn</h5>
                                    <div className="grid grid-cols-3 gap-2">
                                      <div className="text-center p-1.5 bg-slate-950/50 rounded">
                                        <p className="text-[9px] text-gray-500 uppercase">Doanh Nghiệp</p>
                                        <p className="text-sm font-bold text-cyan-400">{row["Số_DN_Địa_Phương"] || 0}</p>
                                      </div>
                                      <div className="text-center p-1.5 bg-slate-950/50 rounded">
                                        <p className="text-[9px] text-gray-500 uppercase">T.Doanh Thu</p>
                                        <p className="text-sm font-bold text-emerald-400 font-mono">{row["Tổng_Doanh_Thu_Địa_Phương"]?.toLocaleString() || 0}</p>
                                      </div>
                                      <div className="text-center p-1.5 bg-slate-950/50 rounded">
                                        <p className="text-[9px] text-gray-500 uppercase">T.Lao Động</p>
                                        <p className="text-sm font-bold text-yellow-500 font-mono">{row["Tổng_Lao_Động_Địa_Phương"]?.toLocaleString() || 0}</p>
                                      </div>
                                    </div>
                                  </div>
                                </div>

                                {/* Right: Detailed actual list of sectors with value > 0 */}
                                <div className="space-y-1.5 bg-[#1e293b]/30 p-3 rounded-lg border border-gray-800 max-h-[160px] overflow-y-auto">
                                  <h5 className="text-[11px] font-bold text-amber-300 uppercase tracking-wider">Danh sách chi tiết các ngành thực chi phát sinh</h5>
                                  <div className="divide-y divide-gray-800">
                                    {pivotAnalysis.sectors.map((sec) => {
                                      const dt = row[`${sec} - Tổng Doanh Thu`] || 0;
                                      const ld = row[`${sec} - Tổng Lao Động`] || 0;
                                      if (dt === 0 && ld === 0) return null;

                                      return (
                                        <div key={sec} className="py-2 flex items-center justify-between text-xs gap-3">
                                          <span className="text-white font-medium break-all pr-2 max-w-[200px] leading-tight">
                                            🏢 {sec}
                                          </span>
                                          <div className="flex gap-4 font-mono shrink-0">
                                            <span className="text-emerald-400 text-right"><span className="text-[9px] text-gray-500 block">D.Thu</span>{dt.toLocaleString()}</span>
                                            <span className="text-yellow-400 text-right"><span className="text-[9px] text-gray-500 block">L.Động</span>{ld.toLocaleString()}</span>
                                          </div>
                                        </div>
                                      );
                                    })}
                                    {/* Safeguard if no sector is active */}
                                    {pivotAnalysis.sectors.every(sec => (row[`${sec} - Tổng Doanh Thu`] || 0) === 0 && (row[`${sec} - Tổng Lao Động`] || 0) === 0) && (
                                      <p className="text-[11px] text-gray-500 italic">Không ghi nhận hoạt động phát sinh.</p>
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
            </table>
          </div>
        ) : (
          /* ================== STANDARD FLAT TABLE LAYOUT ================== */
          <div className="overflow-x-auto border border-gray-800 rounded-xl bg-gray-950/40 max-h-[480px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-gray-900 border-b border-gray-800 text-gray-300 font-mono text-[11px] sticky top-0 z-10 uppercase">
                  <th className="p-3 w-12 text-center text-gray-500">STT</th>
                  {cols.map((col) => (
                    <th key={col} className="p-3 font-semibold whitespace-nowrap">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-800/50 text-gray-200 font-sans text-[11.5px]">
                {filteredRows.map((row, rIdx) => (
                  <tr key={rIdx} className={`hover:bg-slate-800/30 transition-colors ${rIdx % 2 === 1 ? "bg-gray-900/10" : ""}`}>
                    <td className="p-3 text-center text-gray-500 font-mono">{rIdx + 1}</td>
                    {cols.map((col) => {
                      const val = row[col];
                      const isNumeric = typeof val === "number";
                      return (
                        <td
                          key={col}
                          className={`p-3 whitespace-nowrap ${
                            isNumeric
                              ? "font-mono text-emerald-400 font-bold text-right"
                              : col.includes("Địa_Bàn_Xã")
                              ? "font-semibold text-gray-100"
                              : ""
                          }`}
                        >
                          {isNumeric ? val.toLocaleString("en-US") : String(val ?? "")}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
