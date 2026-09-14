import React, { useState, useMemo } from "react";
import { BarChart3, Download, Search } from "lucide-react";
import * as XLSX from "xlsx";

export interface FrequencyAnalysisProps {
  mainData: any[];
  columns: string[];
}

export const FrequencyAnalysis: React.FC<FrequencyAnalysisProps> = ({ mainData = [], columns = [] }) => {
  const [selectedCol, setSelectedCol] = useState<string>(columns[0] || "");
  const [searchTerm, setSearchTerm] = useState("");
  const [sortOrder, setSortOrder] = useState<"desc" | "asc">("desc");

  const frequencyTable = useMemo(() => {
    if (!selectedCol || mainData.length === 0) return [];
    const counts: Record<string, number> = {};
    let totalValid = 0;

    mainData.forEach(row => {
      const val = row[selectedCol];
      const key = val !== null && val !== undefined && String(val).trim() !== "" ? String(val).trim() : "(Trống)";
      counts[key] = (counts[key] || 0) + 1;
      totalValid++;
    });

    let entries = Object.entries(counts).map(([label, count]) => ({
      label,
      count,
      percent: totalValid > 0 ? (count / totalValid) * 100 : 0
    }));

    entries.sort((a, b) => (sortOrder === "desc" ? b.count - a.count : a.count - b.count));

    let cumulative = 0;
    return entries.map(item => {
      cumulative += item.percent;
      return {
        ...item,
        cumulativePercent: cumulative
      };
    });
  }, [mainData, selectedCol, sortOrder]);

  const filteredFrequencies = useMemo(() => {
    if (!searchTerm.trim()) return frequencyTable;
    const term = searchTerm.toLowerCase();
    return frequencyTable.filter(item => item.label.toLowerCase().includes(term));
  }, [frequencyTable, searchTerm]);

  const handleExportExcel = () => {
    if (frequencyTable.length === 0) return;
    const exportData = frequencyTable.map(item => ({
      "Giá Trị": item.label,
      "Tần Số (Số Lượng)": item.count,
      "Tỷ Lệ (%)": Number(item.percent.toFixed(2)),
      "Tỷ Lệ Tích Lũy (%)": Number(item.cumulativePercent.toFixed(2))
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "TanSo");
    XLSX.writeFile(wb, `TanSo_${selectedCol}_${Date.now()}.xlsx`);
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ phong cách #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <BarChart3 className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Phân Tích Tần Số &amp; Phân Phối Dữ Liệu
              </span>
              <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                {frequencyTable.length} giá trị phân loại
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleExportExcel}
            disabled={frequencyTable.length === 0}
            className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 disabled:opacity-50"
            title="Xuất bảng tần số ra file Excel"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Xuất Excel</span>
          </button>
        </div>
      </div>

      {/* THANH ĐIỀU KHIỂN & LỌC */}
      <div className="mx-3.5 bg-white border border-sky-200 rounded-none p-3 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-3 flex-wrap">
          <label className="text-xs font-bold text-slate-700">Chọn cột phân tích:</label>
          <select
            value={selectedCol}
            onChange={e => setSelectedCol(e.target.value)}
            className="text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-semibold text-slate-800 focus:outline-none focus:border-sky-500 font-mono"
          >
            {columns.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>

          <button
            type="button"
            onClick={() => setSortOrder(o => (o === "desc" ? "asc" : "desc"))}
            className="px-2.5 py-1.5 border border-slate-300 bg-slate-50 hover:bg-slate-100 rounded-none text-xs font-semibold text-slate-700 cursor-pointer transition-colors"
          >
            {sortOrder === "desc" ? "Sắp xếp: Giảm dần" : "Sắp xếp: Tăng dần"}
          </button>
        </div>

        <div className="relative">
          <Search className="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Lọc giá trị..."
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)}
            className="pl-7 pr-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-none w-52 font-medium focus:outline-none focus:border-sky-500"
          />
        </div>
      </div>

      {/* BẢNG TẦN SỐ VỚI HEADER #286e42 */}
      <div className="mx-3.5 border border-sky-200 rounded-none overflow-hidden max-h-[500px] overflow-y-auto bg-white shadow-inner">
        <table className="w-full text-left border-collapse text-xs">
          <thead className="bg-[#286e42] sticky top-0 text-white font-bold border-b border-[#1d4f2f] z-10">
            <tr>
              <th className="py-2 px-3 w-12 text-center border-r border-[#1d4f2f]">#</th>
              <th className="py-2 px-3 border-r border-[#1d4f2f]">GIÁ TRỊ ({selectedCol})</th>
              <th className="py-2 px-3 border-r border-[#1d4f2f] text-right">TẦN SỐ</th>
              <th className="py-2 px-3 border-r border-[#1d4f2f] text-right">TỶ LỆ (%)</th>
              <th className="py-2 px-3 border-r border-[#1d4f2f] text-right">TÍCH LŨY (%)</th>
              <th className="py-2 px-3 w-48">BIỂU ĐỒ TỶ LỆ</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filteredFrequencies.length === 0 ? (
              <tr>
                <td colSpan={6} className="py-8 text-center text-slate-400 italic">
                  Không có dữ liệu phân tích.
                </td>
              </tr>
            ) : (
              filteredFrequencies.map((item, idx) => (
                <tr key={idx} className={`hover:bg-sky-50/50 ${idx % 2 === 1 ? "bg-slate-50/40" : ""}`}>
                  <td className="py-1.5 px-3 text-center text-slate-400 font-mono text-[11px] border-r border-slate-100">
                    {idx + 1}
                  </td>
                  <td className="py-1.5 px-3 font-semibold text-slate-800 border-r border-slate-100 font-mono">
                    {item.label}
                  </td>
                  <td className="py-1.5 px-3 text-right font-mono text-slate-800 font-bold border-r border-slate-100">
                    {item.count.toLocaleString("vi-VN")}
                  </td>
                  <td className="py-1.5 px-3 text-right font-mono text-[#0284c7] font-bold border-r border-slate-100">
                    {item.percent.toFixed(2)}%
                  </td>
                  <td className="py-1.5 px-3 text-right font-mono text-slate-600 border-r border-slate-100">
                    {item.cumulativePercent.toFixed(2)}%
                  </td>
                  <td className="py-1.5 px-3">
                    <div className="w-full bg-slate-100 rounded-none h-2.5 overflow-hidden border border-slate-200">
                      <div
                        className="bg-[#286e42] h-2.5 transition-all"
                        style={{ width: `${Math.min(100, item.percent)}%` }}
                      />
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default FrequencyAnalysis;
