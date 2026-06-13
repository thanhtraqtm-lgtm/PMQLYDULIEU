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
}

export function MainDataInlinePreview({
  data,
  columns,
  title,
  subtitle = "Xem nhanh kết quả bảng tính ngay tại tab thao tác hiện tại.",
  mapping = {},
  onExportExcel,
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

  if (!data || data.length === 0) {
    return null;
  }

  const finalColumns = (columns && columns.length > 0 ? columns : Object.keys(data[0] || {}))
    .filter((col) => !col.startsWith("_"));

  return (
    <div className="bg-[#1f2937] border border-[#374151] rounded-2xl shadow-xl overflow-hidden mt-6 animate-slide-up" id="inline-data-preview-container">
      {/* Header Bar */}
      <div 
        className="px-6 py-4 bg-[#111827]/80 flex items-center justify-between border-b border-[#374151] cursor-pointer select-none"
        onClick={() => setIsExpanded(!isExpanded)}
        id="inline-data-preview-header"
      >
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 bg-purple-500/10 border border-purple-500/20 rounded-lg">
            <FileSpreadsheet className="w-5 h-5 text-purple-400" />
          </div>
          <div>
            <h4 className="text-sm font-bold text-white tracking-wide uppercase font-mono">
              📋 {title} ({data.length} dòng)
            </h4>
            <p className="text-[11px] text-gray-400 mt-0.5">
              {subtitle}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3" onClick={(e) => e.stopPropagation()}>
          <button
            onClick={handleExport}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] px-3.5 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-sm shadow-emerald-950/20"
            title="Tải thành phẩm Excel trực tiếp"
            id="inline-btn-export-excel"
          >
            <Download className="w-3.5 h-3.5" /> Xuất Excel tab này
          </button>
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1.5 text-gray-400 hover:text-white hover:bg-gray-800 rounded-lg transition-colors cursor-pointer"
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
              <Search className="absolute left-3 top-2.5 h-3.5 w-3.5 text-gray-400" />
              <input
                type="text"
                placeholder="Tìm lọc trực tiếp..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setPage(1);
                }}
                className="w-full bg-[#111827] border border-[#374151] rounded-xl pl-9 pr-4 py-2 text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                id="inline-search-input"
              />
            </div>

            <div className="text-[11px] text-gray-400 font-mono">
              Hiển thị: <strong className="text-white">{Math.min(filteredData.length, (page - 1) * rowsPerPage + 1)}-{Math.min(filteredData.length, page * rowsPerPage)}</strong> trên tổng số <strong className="text-purple-400">{filteredData.length}</strong> dòng trùng khớp
            </div>
          </div>

          {/* Table View */}
          <div className="overflow-x-auto max-h-[420px] relative border border-gray-800 rounded-xl bg-[#0f172a]/20">
            <table className="w-full text-left text-xs border-collapse" id="inline-preview-table">
              <thead>
                <tr className="bg-[#111827] text-gray-400 border-b border-gray-800 font-mono sticky top-0 z-10 shadow-sm">
                  {finalColumns.map((col) => {
                    const isMota = col === mapping.mota;
                    const isManganh = col === mapping.manganh;
                    const isXa = col === mapping.xa;
                    const isId = col === mapping.idCol;
                    return (
                      <th
                        key={col}
                        className="p-2.5 font-semibold text-center whitespace-nowrap min-w-[120px]"
                      >
                        <span className="flex items-center justify-center gap-1">
                          {isMota && "📝 "}
                          {isManganh && "🏷️ "}
                          {isXa && "🗺️ "}
                          {isId && "🔑 "}
                          {col}
                        </span>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {paginatedData.length > 0 ? (
                  paginatedData.map((row, rIdx) => (
                    <tr
                      key={rIdx}
                      className="border-b border-gray-800/40 hover:bg-gray-800/50 transition-colors"
                    >
                      {finalColumns.map((col) => {
                        const cellValue = row[col];
                        const isMota = col === mapping.mota;
                        return (
                          <td
                            key={col}
                            className={`p-2.5 truncate max-w-[240px] text-center font-sans ${
                              isMota ? "text-slate-100 text-left" : "text-gray-300"
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
                  ))
                ) : (
                  <tr>
                    <td
                      colSpan={finalColumns.length}
                      className="p-8 text-center text-xs text-gray-500 font-sans"
                    >
                      Không tìm thấy bản ghi nào khớp với điều kiện lọc.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Controls */}
          <div className="flex items-center justify-between border-t border-gray-800 pt-3 text-xs">
            <span className="text-gray-400">
              Trang <strong className="text-white">{page}</strong> / {totalPages}
            </span>
            <div className="flex gap-2">
              <button
                disabled={page === 1}
                onClick={() => setPage((prev) => Math.max(1, prev - 1))}
                className={`px-3 py-1.5 rounded-lg border border-gray-700 font-semibold text-[11px] ${
                  page === 1
                    ? "bg-[#111827] text-gray-600 cursor-not-allowed"
                    : "bg-[#111827] hover:bg-[#374151] text-gray-300 cursor-pointer"
                }`}
                id="inline-btn-prev"
              >
                Trước
              </button>
              <button
                disabled={page === totalPages}
                onClick={() => setPage((prev) => Math.min(totalPages, prev + 1))}
                className={`px-3 py-1.5 rounded-lg border border-gray-700 font-semibold text-[11px] ${
                  page === totalPages
                    ? "bg-[#111827] text-gray-600 cursor-not-allowed"
                    : "bg-[#111827] hover:bg-[#374151] text-gray-300 cursor-pointer"
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
}
