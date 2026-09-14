import React, { useState, useMemo } from "react";
import { Database, Search, Play, Download, Sparkles, Filter, Code2, CheckCircle2, AlertCircle } from "lucide-react";
import * as XLSX from "xlsx";

interface ExcelSqlAssistantProps {
  mainData: any[];
  fileName?: string;
}

export const ExcelSqlAssistant: React.FC<ExcelSqlAssistantProps> = ({ mainData = [], fileName = "data.xlsx" }) => {
  const [filterCol, setFilterCol] = useState<string>("");
  const [filterOp, setFilterOp] = useState<string>("contains");
  const [filterVal, setFilterVal] = useState<string>("");
  const [sortCol, setSortCol] = useState<string>("");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [searchPrompt, setSearchPrompt] = useState<string>("");
  const [appliedFilters, setAppliedFilters] = useState<{ col: string; op: string; val: string }[]>([]);

  const columns = useMemo(() => {
    if (mainData.length === 0) return [];
    return Object.keys(mainData[0]);
  }, [mainData]);

  // Handle Natural Language / Smart Query Suggestions
  const handleApplyPrompt = () => {
    if (!searchPrompt.trim()) return;
    const p = searchPrompt.toLowerCase();
    const newFilters: { col: string; op: string; val: string }[] = [];

    // Check for common column heuristics
    columns.forEach(c => {
      const cl = c.toLowerCase();
      if (cl.includes("thuế") || cl.includes("mst")) {
        const mstMatch = p.match(/\b\d{10}(?:-\d{3})?\b/);
        if (mstMatch) {
          newFilters.push({ col: c, op: "equals", val: mstMatch[0] });
        }
      }
      if (cl.includes("doanh thu") || cl.includes("doanhthu")) {
        const numMatch = p.match(/(?:trên|hơn|>|lớn hơn)\s*(\d+)/i);
        if (numMatch) {
          newFilters.push({ col: c, op: ">", val: numMatch[1] });
        }
      }
      if (cl.includes("xã") || cl.includes("địa bàn")) {
        const xaMatch = p.match(/(?:xã|phường|thị trấn)\s+([a-zA-ZÀ-ỹ\s]+)/i);
        if (xaMatch) {
          newFilters.push({ col: c, op: "contains", val: xaMatch[1].trim() });
        }
      }
    });

    if (newFilters.length > 0) {
      setAppliedFilters(prev => [...prev, ...newFilters]);
    } else {
      // General full text search
      setFilterCol(columns[0] || "");
      setFilterOp("contains");
      setFilterVal(searchPrompt.trim());
    }
  };

  const addFilter = () => {
    if (!filterCol || !filterVal) return;
    setAppliedFilters(prev => [...prev, { col: filterCol, op: filterOp, val: filterVal }]);
    setFilterVal("");
  };

  const removeFilter = (index: number) => {
    setAppliedFilters(prev => prev.filter((_, i) => i !== index));
  };

  const filteredData = useMemo(() => {
    let result = [...mainData];

    for (const f of appliedFilters) {
      const valLower = f.val.toLowerCase().trim();
      const numTarget = parseFloat(f.val.replace(/,/g, ""));

      result = result.filter(row => {
        const cell = row[f.col];
        if (cell === null || cell === undefined) return false;
        const cellStr = String(cell).toLowerCase().trim();
        const cellNum = parseFloat(String(cell).replace(/,/g, ""));

        switch (f.op) {
          case "equals":
            return cellStr === valLower;
          case "contains":
            return cellStr.includes(valLower);
          case "starts_with":
            return cellStr.startsWith(valLower);
          case ">":
            return !isNaN(cellNum) && !isNaN(numTarget) && cellNum > numTarget;
          case "<":
            return !isNaN(cellNum) && !isNaN(numTarget) && cellNum < numTarget;
          case ">=":
            return !isNaN(cellNum) && !isNaN(numTarget) && cellNum >= numTarget;
          case "<=":
            return !isNaN(cellNum) && !isNaN(numTarget) && cellNum <= numTarget;
          case "!=":
            return cellStr !== valLower;
          default:
            return true;
        }
      });
    }

    if (sortCol) {
      result.sort((a, b) => {
        const valA = a[sortCol];
        const valB = b[sortCol];
        const numA = parseFloat(String(valA).replace(/,/g, ""));
        const numB = parseFloat(String(valB).replace(/,/g, ""));

        if (!isNaN(numA) && !isNaN(numB)) {
          return sortDir === "asc" ? numA - numB : numB - numA;
        }
        const strA = String(valA || "");
        const strB = String(valB || "");
        return sortDir === "asc" ? strA.localeCompare(strB, "vi") : strB.localeCompare(strA, "vi");
      });
    }

    return result;
  }, [mainData, appliedFilters, sortCol, sortDir]);

  const handleExportExcel = () => {
    if (filteredData.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "KetQuaTruyVan");
    XLSX.writeFile(wb, `KetQua_SQL_${fileName.replace(/\.[^/.]+$/, "")}_${Date.now()}.xlsx`);
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-100">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-violet-100 flex items-center justify-center text-violet-600">
            <Code2 className="w-5 h-5" />
          </div>
          <div>
            <h3 className="font-bold text-slate-800 text-base">Trợ Lý Truy Vấn SQL & Lọc Dữ Liệu Thông Minh</h3>
            <p className="text-xs text-slate-500">
              Lọc trích xuất tập con, xếp hạng chỉ tiêu và truy vấn dữ liệu theo cấu trúc SQL đơn giản.
            </p>
          </div>
        </div>

        <button
          onClick={handleExportExcel}
          disabled={filteredData.length === 0}
          className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors cursor-pointer"
        >
          <Download className="w-4 h-4" />
          Xuất Kết Quả Lọc ({filteredData.length.toLocaleString("vi-VN")})
        </button>
      </div>

      {/* Natural Language Prompt Query */}
      <div className="p-4 bg-violet-50/60 rounded-xl border border-violet-100 space-y-2">
        <div className="flex items-center gap-2 text-violet-900 text-xs font-bold">
          <Sparkles className="w-4 h-4 text-violet-600" />
          Truy vấn bằng ngôn ngữ tự nhiên
        </div>
        <div className="flex gap-2">
          <input
            type="text"
            value={searchPrompt}
            onChange={e => setSearchPrompt(e.target.value)}
            onKeyDown={e => e.key === "Enter" && handleApplyPrompt()}
            placeholder="Ví dụ: lọc doanh thu trên 50000000 hoặc tìm mã số thuế 0101234567..."
            className="flex-1 px-3 py-2 text-xs bg-white border border-violet-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-violet-400"
          />
          <button
            onClick={handleApplyPrompt}
            className="px-4 py-2 bg-violet-600 hover:bg-violet-700 text-white text-xs font-semibold rounded-lg transition-colors cursor-pointer shrink-0"
          >
            Phân Tích & Lọc
          </button>
        </div>
      </div>

      {/* Manual Filter Controls */}
      <div className="space-y-3">
        <div className="flex flex-wrap items-center gap-2">
          <select
            value={filterCol}
            onChange={e => setFilterCol(e.target.value)}
            className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 bg-white font-medium text-slate-700"
          >
            <option value="">-- Chọn cột điều kiện --</option>
            {columns.map(c => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>

          <select
            value={filterOp}
            onChange={e => setFilterOp(e.target.value)}
            className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 bg-white font-medium text-slate-700"
          >
            <option value="contains">Chứa (Contains)</option>
            <option value="equals">Bằng (=)</option>
            <option value="starts_with">Bắt đầu bằng</option>
            <option value=">">Lớn hơn (&gt;)</option>
            <option value="<">Nhỏ hơn (&lt;)</option>
            <option value=">=">Lớn hơn hoặc bằng (&gt;=)</option>
            <option value="<=">Nhỏ hơn hoặc bằng (&lt;=)</option>
            <option value="!=">Khác (!=)</option>
          </select>

          <input
            type="text"
            value={filterVal}
            onChange={e => setFilterVal(e.target.value)}
            onKeyDown={e => e.key === "Enter" && addFilter()}
            placeholder="Giá trị so sánh..."
            className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 bg-white flex-1 min-w-40 focus:outline-none focus:ring-2 focus:ring-violet-400"
          />

          <button
            onClick={addFilter}
            className="px-3 py-1.5 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold rounded-lg transition-colors cursor-pointer"
          >
            + Thêm Điều Kiện (AND)
          </button>
        </div>

        {/* Applied filters tags */}
        {appliedFilters.length > 0 && (
          <div className="flex flex-wrap items-center gap-2 pt-2">
            <span className="text-xs font-semibold text-slate-500">Bộ lọc đang áp dụng:</span>
            {appliedFilters.map((f, i) => (
              <span
                key={i}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium bg-violet-100 text-violet-800 border border-violet-200"
              >
                <span>
                  <strong>{f.col}</strong> {f.op} "{f.val}"
                </span>
                <button
                  onClick={() => removeFilter(i)}
                  className="text-violet-600 hover:text-violet-900 font-bold ml-1 cursor-pointer"
                >
                  ×
                </button>
              </span>
            ))}
            <button
              onClick={() => setAppliedFilters([])}
              className="text-xs text-rose-600 hover:underline font-medium ml-2 cursor-pointer"
            >
              Xóa tất cả bộ lọc
            </button>
          </div>
        )}
      </div>

      {/* Sorting & Preview Table */}
      <div className="border border-slate-200 rounded-lg overflow-hidden">
        <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs text-slate-600">
          <div className="font-semibold">
            Kết quả: {filteredData.length.toLocaleString("vi-VN")} / {mainData.length.toLocaleString("vi-VN")} dòng
          </div>
          <div className="flex items-center gap-2">
            <span>Sắp xếp:</span>
            <select
              value={sortCol}
              onChange={e => setSortCol(e.target.value)}
              className="border border-slate-300 rounded px-2 py-1 bg-white text-xs"
            >
              <option value="">(Không sắp xếp)</option>
              {columns.map(c => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
            <button
              onClick={() => setSortDir(d => (d === "asc" ? "desc" : "asc"))}
              className="px-2 py-1 border border-slate-300 rounded bg-white font-semibold text-xs cursor-pointer hover:bg-slate-50"
            >
              {sortDir === "asc" ? "Tăng dần (ASC)" : "Giảm dần (DESC)"}
            </button>
          </div>
        </div>

        <div className="overflow-x-auto max-h-96">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="bg-slate-100 sticky top-0 text-slate-700 font-bold border-b border-slate-300">
              <tr>
                <th className="py-2 px-3 w-10 text-center border-r border-slate-200">#</th>
                {columns.map(c => (
                  <th key={c} className="py-2 px-3 border-r border-slate-200 whitespace-nowrap">
                    {c}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredData.slice(0, 50).map((row, idx) => (
                <tr key={idx} className="hover:bg-slate-50">
                  <td className="py-1.5 px-3 text-center text-slate-400 border-r border-slate-100 font-mono text-[11px]">
                    {idx + 1}
                  </td>
                  {columns.map(c => (
                    <td key={c} className="py-1.5 px-3 border-r border-slate-100 text-slate-700 whitespace-nowrap max-w-xs truncate">
                      {row[c] !== undefined && row[c] !== null ? String(row[c]) : ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
export default ExcelSqlAssistant;
