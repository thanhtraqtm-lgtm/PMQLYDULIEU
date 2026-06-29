import React, { useState, useMemo } from "react";
import { Search, Download, FileSpreadsheet, ChevronDown, ChevronUp } from "lucide-react";
import * as XLSX from "xlsx";

interface MainDataInlinePreviewProps {
  data: any[];
  columns: string[];
  title: string;
  subtitle?: string;
  mapping?: {
    mota?: string;
    manganh?: string;
    xa?: string;
    doanhthu?: string;
    laodong?: string;
    idCol?: string;
  };
  onExportExcel?: () => void;
  enableSelection?: boolean;
  selectedColumns?: string[];
  onSelectedColumnsChange?: (cols: string[]) => void;
  selectedRows?: number[];
  onSelectedRowsChange?: (rows: number[]) => void;
}

export const MainDataInlinePreview = React.memo(function MainDataInlinePreview({
  data,
  columns,
  title,
  subtitle = "Xem nhanh kết quả bảng tính ngay tại tab thao tác hiện tại.",
  mapping = {},
  onExportExcel,
  enableSelection = false,
  selectedColumns = [],
  onSelectedColumnsChange,
  selectedRows = [],
  onSelectedRowsChange,
}: MainDataInlinePreviewProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [page, setPage] = useState(1);
  const [isExpanded, setIsExpanded] = useState(true);
  const rowsPerPage = 50;

  // Filter logic
  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return data;
    const term = searchTerm.toLowerCase();
    return data.filter((row) => {
      if (!row) return false;
      return Object.values(row).some((val) => {
        if (val === null || val === undefined) return false;
        return String(val).toLowerCase().includes(term);
      });
    });
  }, [data, searchTerm]);

  // Pagination logic
  const totalPages = Math.max(1, Math.ceil(filteredData.length / rowsPerPage));
  const paginatedData = useMemo(() => {
    const start = (page - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, page, rowsPerPage]);

  // Adjust page if it exceeds totalPages
  React.useEffect(() => {
    if (page > totalPages) {
      setPage(1);
    }
  }, [totalPages, page]);

  const handleExport = () => {
    if (onExportExcel) {
      onExportExcel();
      return;
    }
    // Default fallback simple export
    if (data.length === 0) return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "ResultData");
    XLSX.writeFile(wb, "KetQuaThaoTac.xlsx");
  };

  const finalColumns = (columns && columns.length > 0 ? columns : (data && data.length > 0 ? Object.keys(data[0] || {}) : []))
    .filter((col) => !col.startsWith("_"));

  // Check if all paginated rows are selected
  const allPaginatedRowsSelected = useMemo(() => {
    if (paginatedData.length === 0) return false;
    return paginatedData.every((row) => {
      const originalIndex = data.indexOf(row);
      return selectedRows.includes(originalIndex);
    });
  }, [paginatedData, data, selectedRows]);

  const handleToggleSelectAllRows = () => {
    if (!onSelectedRowsChange) return;
    const paginatedIndices = paginatedData.map((row) => data.indexOf(row));
    const allSelected = paginatedIndices.every((idx) => selectedRows.includes(idx));
    
    if (allSelected) {
      // Deselect all paginated rows
      onSelectedRowsChange(selectedRows.filter(idx => !paginatedIndices.includes(idx)));
    } else {
      // Select all paginated rows
      const newSelected = [...selectedRows];
      paginatedIndices.forEach(idx => {
        if (!newSelected.includes(idx)) newSelected.push(idx);
      });
      onSelectedRowsChange(newSelected);
    }
  };

  if (finalColumns.length === 0) {
    return null;
  }

  return (
    <div className="bg-white border border-slate-200 rounded-2xl shadow-lg overflow-hidden mt-6 animate-slide-up" id="inline-data-preview-container">
      {/* Header Bar */}
      <div 
        className="px-6 py-4 bg-slate-50 flex items-center justify-between border-b border-slate-200 cursor-pointer select-none"
        onClick={() => setIsExpanded(!isExpanded)}
        id="inline-data-preview-header"
      >
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 bg-indigo-50 border border-indigo-100 rounded-lg">
            <FileSpreadsheet className="w-5 h-5 text-indigo-600" />
          </div>
          <div>
            <h4 className="text-sm font-bold text-slate-800 tracking-wide uppercase font-mono">
              📋 {title} ({data.length} dòng)
            </h4>
            <p className="text-[11px] text-slate-500 mt-0.5">
              {subtitle}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3" onClick={(e) => e.stopPropagation()}>
          {enableSelection && (
            <div className="hidden sm:flex items-center gap-2 text-xs text-slate-650 bg-indigo-50/55 border border-indigo-100 px-3 py-1.5 rounded-lg">
              <span className="font-bold text-indigo-700">Đã chọn:</span>
              <span className="font-mono font-black text-indigo-800">{selectedColumns.length} cột</span>
              <span className="text-slate-300">|</span>
              <span className="font-mono font-black text-indigo-800">{selectedRows.length} dòng</span>
            </div>
          )}
          <button
            onClick={handleExport}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] px-3.5 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-sm shadow-emerald-100"
            title="Tải thành phẩm Excel trực tiếp"
            id="inline-btn-export-excel"
          >
            <Download className="w-3.5 h-3.5" /> Xuất Excel tab này
          </button>
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors cursor-pointer"
            id="inline-btn-expand-toggle"
          >
            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {isExpanded && (
        <div className="p-5 space-y-4" id="inline-data-preview-body">
          {/* Filtering Bar */}
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="relative w-full sm:max-w-xs">
              <Search className="absolute left-3 top-2.5 h-3.5 w-3.5 text-slate-400" />
              <input
                type="text"
                placeholder="Tìm lọc trực tiếp..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setPage(1);
                }}
                className="w-full bg-white border border-slate-300 rounded-xl pl-9 pr-4 py-2 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                id="inline-search-input"
              />
            </div>

            <div className="text-[11px] text-slate-500 font-mono">
              Hiển thị: <strong className="text-slate-800">{Math.min(filteredData.length, (page - 1) * rowsPerPage + 1)}-{Math.min(filteredData.length, page * rowsPerPage)}</strong> trên tổng số <strong className="text-indigo-600">{filteredData.length}</strong> dòng trùng khớp
            </div>
          </div>

          {/* Table View with Sticky Header and Selection */}
          <div className="overflow-x-auto overflow-y-auto max-h-[440px] relative border border-slate-200 rounded-xl bg-white">
            <table className="w-full text-left text-xs border-collapse" id="inline-preview-table">
              <thead>
                <tr className="bg-slate-100 text-slate-600 border-b border-slate-200 font-mono sticky top-0 z-20 shadow-xs">
                  {/* Row Selection Header Checkbox */}
                  {enableSelection && (
                    <th className="p-2.5 font-bold text-center whitespace-nowrap bg-slate-100 border-r border-slate-200 sticky left-0 z-30 min-w-[40px] w-[40px]">
                      <input
                        type="checkbox"
                        checked={allPaginatedRowsSelected}
                        onChange={handleToggleSelectAllRows}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer w-4 h-4"
                        title="Chọn tất cả dòng trang này"
                      />
                    </th>
                  )}

                  {finalColumns.map((col) => {
                    const isMota = col === mapping.mota;
                    const isManganh = col === mapping.manganh;
                    const isXa = col === mapping.xa;
                    const isId = col === mapping.idCol;
                    const isColSelected = selectedColumns.includes(col);

                    return (
                      <th
                        key={col}
                        className={`p-2.5 font-bold text-center whitespace-nowrap min-w-[130px] border-r border-slate-200 transition-colors ${
                          enableSelection && isColSelected 
                            ? "bg-indigo-50 text-indigo-900" 
                            : "bg-slate-100 text-slate-700"
                        }`}
                      >
                        <div className="flex items-center justify-between gap-1.5 px-1">
                          {enableSelection && (
                            <input
                              type="checkbox"
                              checked={isColSelected}
                              onChange={(e) => {
                                if (!onSelectedColumnsChange) return;
                                const checked = e.target.checked;
                                if (checked) {
                                  onSelectedColumnsChange([...selectedColumns, col]);
                                } else {
                                  onSelectedColumnsChange(selectedColumns.filter(c => c !== col));
                                }
                              }}
                              className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer w-3.5 h-3.5 shrink-0 mr-1.5"
                              title="Chọn cột này"
                            />
                          )}
                          <span className="flex items-center justify-center gap-1 mx-auto">
                            {isMota && "📝 "}
                            {isManganh && "🏷️ "}
                            {isXa && "🗺️ "}
                            {isId && "🔑 "}
                            {col}
                          </span>
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {paginatedData.length > 0 ? (
                  paginatedData.map((row, rIdx) => {
                    const originalIndex = data.indexOf(row);
                    const isRowSelected = selectedRows.includes(originalIndex);

                    return (
                      <tr
                        key={rIdx}
                        className={`border-b border-slate-100 transition-colors ${
                          enableSelection && isRowSelected
                            ? "bg-indigo-50/40 hover:bg-indigo-50/60"
                            : "odd:bg-white even:bg-slate-50/40 hover:bg-slate-100/50"
                        }`}
                      >
                        {/* Row Selection Body Checkbox */}
                        {enableSelection && (
                          <td className={`p-2.5 text-center border-r border-slate-200 sticky left-0 z-10 transition-colors ${
                            isRowSelected ? "bg-indigo-50/80" : "bg-slate-50/50"
                          }`}>
                            <input
                              type="checkbox"
                              checked={isRowSelected}
                              onChange={(e) => {
                                if (!onSelectedRowsChange) return;
                                const checked = e.target.checked;
                                if (checked) {
                                  onSelectedRowsChange([...selectedRows, originalIndex]);
                                } else {
                                  onSelectedRowsChange(selectedRows.filter(idx => idx !== originalIndex));
                                }
                              }}
                              className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer w-3.5 h-3.5"
                            />
                          </td>
                        )}

                        {finalColumns.map((col) => {
                          const cellValue = row[col];
                          const isMota = col === mapping.mota;
                          const isColSelected = selectedColumns.includes(col);

                          return (
                            <td
                              key={col}
                              className={`p-2.5 truncate max-w-[240px] text-center font-sans border-r border-slate-100/80 transition-colors ${
                                isMota ? "text-slate-900 text-left font-medium" : "text-slate-800"
                              } ${
                                enableSelection && isColSelected 
                                  ? "bg-indigo-50/20 font-semibold text-indigo-950" 
                                  : ""
                              }`}
                              title={String(cellValue)}
                            >
                              {cellValue === null || cellValue === undefined
                                ? ""
                                : String(cellValue)}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td
                      colSpan={finalColumns.length + (enableSelection ? 1 : 0)}
                      className="p-8 text-center text-xs text-slate-500 font-sans"
                    >
                      Không tìm thấy bản ghi nào khớp với điều kiện lọc.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Controls */}
          <div className="flex items-center justify-between border-t border-slate-200 pt-3 text-xs">
            <span className="text-slate-500">
              Trang <strong className="text-slate-800">{page}</strong> / {totalPages}
            </span>
            <div className="flex gap-2">
              <button
                disabled={page === 1}
                onClick={() => setPage((prev) => Math.max(1, prev - 1))}
                className={`px-3 py-1.5 rounded-lg border border-slate-200 font-semibold text-[11px] ${
                  page === 1
                    ? "bg-slate-50 text-slate-400 cursor-not-allowed"
                    : "bg-white hover:bg-slate-100 text-slate-700 cursor-pointer"
                }`}
                id="inline-btn-prev"
              >
                Trước
              </button>
              <button
                disabled={page === totalPages}
                onClick={() => setPage((prev) => Math.min(totalPages, prev + 1))}
                className={`px-3 py-1.5 rounded-lg border border-slate-200 font-semibold text-[11px] ${
                  page === totalPages
                    ? "bg-slate-50 text-slate-400 cursor-not-allowed"
                    : "bg-white hover:bg-slate-100 text-slate-700 cursor-pointer"
                }`}
                id="inline-btn-next"
              >
                Sau
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
});
