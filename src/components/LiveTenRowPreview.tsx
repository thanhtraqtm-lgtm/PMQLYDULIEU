import React, { useState, useMemo } from "react";
import { Edit2, Sparkles, Check, RefreshCw, AlertCircle } from "lucide-react";

interface LiveTenRowPreviewProps {
  data: any[];
  columns: string[];
  onUpdateData?: (newData: any[]) => void;
  highlightedIndices?: number[]; // indices of rows that have violations/outliers/matches
  highlightLabel?: string;       // label for the highlighted selection (e.g., "Dòng vi phạm", "Dòng dị biệt")
  title?: string;
}

export const LiveTenRowPreview = React.memo(function LiveTenRowPreview({
  data,
  columns,
  onUpdateData,
  highlightedIndices = [],
  highlightLabel = "Dòng phát hiện lỗi/kết quả",
  title = "BẢNG XEM NHANH & SỬA TRỰC TIẾP DỮ LIỆU"
}: LiveTenRowPreviewProps) {
  const [viewMode, setViewMode] = useState<"all" | "highlighted">("all");
  const [editingCell, setEditingCell] = useState<{ rowIdx: number; col: string } | null>(null);
  const [editValue, setEditValue] = useState<string>("");

  const filteredColumns = useMemo(() => {
    return columns.filter(col => !col.startsWith("_") && col !== "Loi_Logic");
  }, [columns]);

  // Determine which 10 rows to display
  const rowsToDisplay = useMemo(() => {
    if (viewMode === "highlighted" && highlightedIndices.length > 0) {
      return highlightedIndices.slice(0, 10).map(idx => ({
        originalIdx: idx,
        rowNum: idx + 1,
        row: data[idx]
      }));
    }

    // Default: Top 10 rows of the dataset
    return data.slice(0, 10).map((row, idx) => ({
      originalIdx: idx,
      rowNum: idx + 1,
      row
    }));
  }, [data, viewMode, highlightedIndices]);

  const handleStartEdit = (rowIdx: number, col: string, val: any) => {
    if (!onUpdateData) return;
    setEditingCell({ rowIdx, col });
    setEditValue(String(val ?? ""));
  };

  const handleSaveEdit = (originalIdx: number, col: string) => {
    if (!onUpdateData) return;
    
    const newData = [...data];
    newData[originalIdx] = {
      ...newData[originalIdx],
      [col]: editValue
    };
    
    onUpdateData(newData);
    setEditingCell(null);
  };

  const handleKeyPress = (e: React.KeyboardEvent, originalIdx: number, col: string) => {
    if (e.key === "Enter") {
      handleSaveEdit(originalIdx, col);
    } else if (e.key === "Escape") {
      setEditingCell(null);
    }
  };

  if (data.length === 0) return null;

  return (
    <div className="bg-slate-900 text-slate-100 rounded-2xl border border-slate-750 p-4 space-y-3.5 shadow-md">
      {/* Title block */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800 pb-3">
        <div className="space-y-0.5">
          <div className="flex items-center gap-2">
            <span className="bg-indigo-500/10 text-indigo-400 p-1 rounded-lg">
              <Sparkles className="w-4 h-4 text-indigo-400 animate-pulse" />
            </span>
            <h4 className="text-xs font-black tracking-wider uppercase text-slate-200">
              {title}
            </h4>
          </div>
          <p className="text-[10.5px] text-slate-400 font-sans">
            Sửa nhanh trực tiếp tại ô bất kỳ (Nhấn Enter để lưu, dữ liệu nguồn sẽ cập nhật tức thì).
          </p>
        </div>

        {/* Filter modes selection */}
        <div className="flex items-center gap-1.5 bg-slate-950 p-1 rounded-xl border border-slate-800 self-start sm:self-center">
          <button
            onClick={() => setViewMode("all")}
            className={`px-3 py-1.5 rounded-lg text-[11px] font-bold transition-all border-0 cursor-pointer ${
              viewMode === "all"
                ? "bg-indigo-600 text-white shadow-sm"
                : "bg-transparent text-slate-400 hover:text-slate-200"
            }`}
          >
            📋 10 dòng đầu
          </button>
          <button
            onClick={() => setViewMode("highlighted")}
            disabled={highlightedIndices.length === 0}
            className={`px-3 py-1.5 rounded-lg text-[11px] font-bold transition-all border-0 flex items-center gap-1 cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
              viewMode === "highlighted"
                ? "bg-amber-600 text-white shadow-sm"
                : "bg-transparent text-slate-400 hover:text-slate-200"
            }`}
          >
            ⚠️ {highlightLabel} ({highlightedIndices.length})
          </button>
        </div>
      </div>

      {/* Grid container */}
      <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-950 max-h-[300px]">
        <table className="w-full text-left border-collapse table-auto text-xs">
          <thead>
            <tr className="bg-slate-900 border-b border-slate-800">
              <th className="p-2.5 font-bold text-slate-400 text-center font-mono w-14 shrink-0">Dòng</th>
              {filteredColumns.map(col => (
                <th key={col} className="p-2.5 font-bold text-slate-300 min-w-[140px] border-l border-slate-800/50">
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/40">
            {rowsToDisplay.map(({ originalIdx, rowNum, row }) => {
              const isRowViolated = highlightedIndices.includes(originalIdx);
              return (
                <tr
                  key={originalIdx}
                  className={`hover:bg-slate-900/60 transition-colors ${
                    isRowViolated 
                      ? "bg-amber-950/20 hover:bg-amber-950/35 border-l-2 border-amber-500" 
                      : ""
                  }`}
                >
                  <td className="p-2.5 text-center font-mono font-bold text-slate-500 bg-slate-900/30">
                    {rowNum}
                  </td>
                  {filteredColumns.map(col => {
                    const value = row[col];
                    const isEditing = editingCell?.rowIdx === originalIdx && editingCell?.col === col;

                    return (
                      <td 
                        key={col} 
                        className="p-2 border-l border-slate-800/30 group relative cursor-pointer"
                        onClick={() => !isEditing && handleStartEdit(originalIdx, col, value)}
                      >
                        {isEditing ? (
                          <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                            <input
                              type="text"
                              value={editValue}
                              onChange={(e) => setEditValue(e.target.value)}
                              onKeyDown={(e) => handleKeyPress(e, originalIdx, col)}
                              className="bg-slate-800 text-white text-xs px-2 py-1 rounded border border-indigo-500 outline-none w-full font-mono"
                              autoFocus
                            />
                            <button
                              onClick={() => handleSaveEdit(originalIdx, col)}
                              className="p-1 rounded bg-indigo-600 hover:bg-indigo-500 text-white border-0 cursor-pointer shrink-0"
                            >
                              <Check className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center justify-between gap-1 min-h-[1.5rem]">
                            <span className={`font-mono truncate ${isRowViolated && col === "Loi_Logic" ? "text-amber-400 font-bold" : ""}`}>
                              {value === null || value === undefined ? "" : String(value)}
                            </span>
                            <Edit2 className="w-3 h-3 text-slate-600 group-hover:text-slate-300 opacity-0 group-hover:opacity-100 transition-all ml-1 shrink-0" />
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Row summary info */}
      <div className="flex items-center justify-between text-[11px] text-slate-400 font-sans px-1">
        <div className="flex items-center gap-1">
          <AlertCircle className="w-3.5 h-3.5 text-indigo-400 shrink-0" />
          <span>
            Đang hiển thị {rowsToDisplay.length} dòng ({viewMode === "all" ? "10 dòng đầu của bảng" : `10 dòng ${highlightLabel} đầu tiên`}).
          </span>
        </div>
        <span>Tổng cộng: <strong className="text-slate-200">{data.length} dòng</strong></span>
      </div>
    </div>
  );
});
