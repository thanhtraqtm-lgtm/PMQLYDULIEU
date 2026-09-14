import React, { useState, useMemo } from "react";
import { Search, Download, CheckSquare, Square, ChevronLeft, ChevronRight, Eye, Table } from "lucide-react";

export interface MainDataInlinePreviewProps {
  data: any[];
  columns: string[];
  title?: string;
  subtitle?: string;
  mapping?: any;
  onExportExcel?: () => void;
  enableSelection?: boolean;
  selectedColumns?: string[];
  onSelectedColumnsChange?: (cols: string[]) => void;
  selectedRows?: number[];
  onSelectedRowsChange?: (rows: number[]) => void;
  embedded?: boolean;
}

export const MainDataInlinePreview: React.FC<MainDataInlinePreviewProps> = ({
  data = [],
  columns = [],
  title = "XEM TRƯỚC DỮ LIỆU",
  subtitle = "",
  mapping,
  onExportExcel,
  enableSelection = false,
  selectedColumns = [],
  onSelectedColumnsChange,
  embedded = false,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);

  const activeCols = useMemo(() => {
    if (enableSelection && selectedColumns.length > 0) {
      return columns.filter(c => selectedColumns.includes(c));
    }
    return columns;
  }, [columns, enableSelection, selectedColumns]);

  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return data;
    const term = searchTerm.toLowerCase().trim();
    return data.filter(row => {
      return Object.values(row).some(val => {
        if (val === null || val === undefined) return false;
        return String(val).toLowerCase().includes(term);
      });
    });
  }, [data, searchTerm]);

  const totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
  const currentPage = Math.min(page, totalPages);

  const displayedRows = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  const toggleColumnSelection = (col: string) => {
    if (!onSelectedColumnsChange) return;
    if (selectedColumns.includes(col)) {
      onSelectedColumnsChange(selectedColumns.filter(c => c !== col));
    } else {
      onSelectedColumnsChange([...selectedColumns, col]);
    }
  };

  const selectAllColumns = () => {
    if (onSelectedColumnsChange) onSelectedColumnsChange([...columns]);
  };

  const deselectAllColumns = () => {
    if (onSelectedColumnsChange) onSelectedColumnsChange([]);
  };

  // Tính tổng dòng tổng cộng (Footer Totals) cho các cột số / cột đếm số lần trùng
  const columnTotals = useMemo(() => {
    const totals: Record<string, { sum: number; count: number; isNumeric: boolean; isDupCol: boolean; totalRows: number }> = {};
    activeCols.forEach(col => {
      let sum = 0;
      let numericCount = 0;
      let nonEmptyCount = 0;
      const lower = col.toLowerCase();
      const isDupCol = lower.includes("trung") || lower.includes("lần") || lower.includes("so_lan") || lower.includes("count") || lower.includes("dup");

      filteredData.forEach(row => {
        const val = row[col];
        if (val !== undefined && val !== null && val !== "") {
          nonEmptyCount++;
          const num = Number(val);
          if (!isNaN(num) && typeof val !== "boolean") {
            sum += num;
            numericCount++;
          }
        }
      });

      const isNumeric = (numericCount > 0 && numericCount >= nonEmptyCount * 0.7) || isDupCol;
      totals[col] = {
        sum,
        count: nonEmptyCount,
        isNumeric,
        isDupCol,
        totalRows: filteredData.length
      };
    });
    return totals;
  }, [activeCols, filteredData]);

  return (
    <div className={embedded ? "bg-sky-50/30 flex flex-col w-full" : "bg-sky-50/30 rounded-none shadow-xs border border-sky-200 overflow-hidden flex flex-col"}>
      {/* Header thanh tiêu đề của bảng danh sách: màu xanh tươi đẹp từ ảnh mẫu #286e42, chữ trắng */}
      <div className="px-4 sm:px-5 py-2.5 bg-[#286e42] text-white border-b border-[#1d4f2f] flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <Table className="w-4 h-4 text-emerald-200" />
          <h3 className="font-bold text-white text-xs uppercase tracking-wider">{title}</h3>
          <span className="bg-[#1d4f2f] text-emerald-100 text-[11px] font-bold px-2 py-0.5 rounded-none font-mono border border-emerald-400/30">
            {filteredData.length.toLocaleString("vi-VN")} dòng
          </span>
          {subtitle ? <span className="text-xs text-emerald-100 font-normal ml-1">• {subtitle}</span> : null}
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {onExportExcel && (
            <button
              onClick={onExportExcel}
              className="inline-flex items-center gap-1 px-2.5 py-1 bg-[#059669] hover:bg-[#047857] text-white text-xs font-bold rounded-none shadow-xs transition-colors cursor-pointer whitespace-nowrap border-0"
              title="Xuất bảng dữ liệu ra file Excel"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Excel</span>
            </button>
          )}

          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Tìm kiếm nhanh..."
              value={searchTerm}
              onChange={e => {
                setSearchTerm(e.target.value);
                setPage(1);
              }}
              className="pl-8 pr-2.5 py-1 text-xs bg-white text-slate-800 border border-slate-300 rounded-none focus:outline-none focus:border-cyan-400 w-44 sm:w-56 font-normal"
            />
          </div>
        </div>
      </div>

      {/* Column selector if enabled - nền xanh nhạt */}
      {enableSelection && columns.length > 0 && (
        <div className="px-4 sm:px-5 py-2 bg-sky-100/70 border-b border-sky-200 flex items-center gap-2 overflow-x-auto text-xs">
          <span className="font-bold text-slate-700 whitespace-nowrap">Chọn cột:</span>
          <button
            onClick={selectAllColumns}
            className="bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs px-2 py-0.5 rounded-none whitespace-nowrap cursor-pointer transition-colors shadow-2xs border-0"
          >
            Tất cả
          </button>
          <button
            onClick={deselectAllColumns}
            className="bg-slate-600 hover:bg-slate-700 text-white font-bold text-xs px-2 py-0.5 rounded-none whitespace-nowrap cursor-pointer transition-colors shadow-2xs border-0"
          >
            Bỏ chọn
          </button>
          <div className="flex items-center gap-1.5 pl-2">
            {columns.map(col => {
              const checked = selectedColumns.includes(col);
              return (
                <button
                  key={col}
                  onClick={() => toggleColumnSelection(col)}
                  className={`inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-bold transition-colors cursor-pointer whitespace-nowrap rounded-none ${
                    checked
                      ? "bg-[#0284c7] hover:bg-[#0369a1] text-white border-0 shadow-2xs"
                      : "bg-white border border-slate-300 text-slate-700 hover:bg-slate-100"
                  }`}
                >
                  {checked ? <CheckSquare className="w-3 h-3 text-white" /> : <Square className="w-3 h-3 text-slate-400" />}
                  {col}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* Table Container - nền xanh nhạt viền */}
      <div className="overflow-x-auto max-h-[540px] relative border-b border-sky-200">
        <table className="w-full text-left border-collapse text-xs">
          <thead className="bg-[#286e42] sticky top-0 z-10 text-white font-bold border-b border-[#1d4f2f] shadow-xs">
            <tr>
              <th className="py-2.5 px-3 w-12 text-center border-r border-[#1d4f2f] bg-[#215935] text-white font-mono font-bold">
                STT
              </th>
              {activeCols.map(col => {
                const isRoleCol = mapping && Object.values(mapping).includes(col);
                return (
                  <th
                    key={col}
                    className={`py-2.5 px-3 font-bold border-r border-[#1d4f2f] whitespace-nowrap ${
                      isRoleCol ? "bg-[#338252] text-emerald-100" : "bg-[#286e42] text-white"
                    }`}
                  >
                    {col}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody className="divide-y divide-sky-100 bg-white">
            {displayedRows.length === 0 ? (
              <tr>
                <td colSpan={activeCols.length + 1} className="py-12 text-center text-slate-400 bg-sky-50/20">
                  Không tìm thấy bản ghi dữ liệu phù hợp.
                </td>
              </tr>
            ) : (
              displayedRows.map((row, idx) => {
                const globalIndex = (currentPage - 1) * pageSize + idx + 1;
                return (
                  <tr key={idx} className="hover:bg-sky-100/70 transition-colors even:bg-sky-50/40 group">
                    <td className="py-2 px-3 text-center text-slate-500 border-r border-sky-100 font-mono text-[11px] bg-sky-50/70 font-semibold">
                      {globalIndex}
                    </td>
                    {activeCols.map(col => (
                      <td key={col} className="py-2 px-3 border-r border-sky-100 text-slate-800 whitespace-nowrap max-w-xs truncate font-medium">
                        {row[col] !== undefined && row[col] !== null ? String(row[col]) : ""}
                      </td>
                    ))}
                  </tr>
                );
              })
            )}
          </tbody>
          <tfoot className="bg-emerald-50/95 sticky bottom-0 z-10 border-t-2 border-[#286e42] shadow-sm">
            <tr className="font-bold text-xs text-slate-800">
              <td className="py-2.5 px-3 text-center border-r border-emerald-300 font-mono text-[11px] bg-[#286e42] text-white uppercase tracking-wider font-bold">
                TỔNG
              </td>
              {activeCols.map(col => {
                const tot = columnTotals[col];
                if (tot && tot.isNumeric) {
                  return (
                    <td key={col} className="py-2 px-3 border-r border-emerald-200 font-mono text-emerald-950 whitespace-nowrap bg-emerald-100/70 font-bold">
                      {tot.isDupCol ? (
                        <div className="flex flex-col">
                          <span className="text-emerald-950 font-black text-[12px]">{tot.totalRows.toLocaleString("vi-VN")}</span>
                          {tot.sum !== tot.totalRows && tot.sum > 0 && (
                            <span className="text-[10px] font-normal text-emerald-800 font-sans">
                              (Ghi nhận: {tot.sum.toLocaleString("vi-VN")})
                            </span>
                          )}
                        </div>
                      ) : (
                        <span className="text-[12px]">{tot.sum.toLocaleString("vi-VN")}</span>
                      )}
                    </td>
                  );
                }
                return (
                  <td key={col} className="py-2 px-3 border-r border-emerald-200 text-slate-400 font-normal italic text-[11px] bg-emerald-50/50">
                    -
                  </td>
                );
              })}
            </tr>
          </tfoot>
        </table>
      </div>

      {/* Pagination Footer - nền xanh nhạt */}
      <div className="px-4 sm:px-5 py-2.5 bg-sky-50/90 border-t border-sky-200 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-700">
        <div className="flex items-center gap-2">
          <span className="font-semibold text-slate-600">Hiển thị</span>
          <select
            value={pageSize}
            onChange={e => {
              setPageSize(Number(e.target.value));
              setPage(1);
            }}
            className="border border-sky-300 rounded-none px-2 py-1 bg-white text-xs font-semibold text-slate-700 focus:outline-none"
          >
            <option value={10}>10 dòng</option>
            <option value={20}>20 dòng</option>
            <option value={50}>50 dòng</option>
            <option value={100}>100 dòng</option>
          </select>
          <span className="text-slate-500">
            (Tổng <strong className="text-slate-700">{filteredData.length.toLocaleString("vi-VN")}</strong> dòng)
          </span>
        </div>

        <div className="flex items-center gap-2">
          <span className="text-slate-600 font-medium">
            Trang <strong className="text-sky-800">{currentPage}</strong> / {totalPages}
          </span>
          <div className="flex items-center gap-1">
            <button
              onClick={() => setPage(p => Math.max(1, p - 1))}
              disabled={currentPage <= 1}
              className="p-1 rounded-none bg-[#0284c7] hover:bg-[#0369a1] text-white disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed cursor-pointer transition-colors shadow-2xs border-0"
              title="Trang trước"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <button
              onClick={() => setPage(p => Math.min(totalPages, p + 1))}
              disabled={currentPage >= totalPages}
              className="p-1 rounded-none bg-[#0284c7] hover:bg-[#0369a1] text-white disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed cursor-pointer transition-colors shadow-2xs border-0"
              title="Trang sau"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
