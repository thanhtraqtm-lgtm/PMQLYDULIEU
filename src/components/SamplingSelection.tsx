import React, { useState } from "react";
import { Shuffle, Download } from "lucide-react";
import * as XLSX from "xlsx";

interface SamplingSelectionProps {
  mainData: any[];
  columns: string[];
  mapping?: any;
  setLoading?: (loading: boolean) => void;
  setStatusMessage?: (msg: string) => void;
  sampCorpData?: any[];
  setSampCorpData?: (data: any[]) => void;
  sampCorpFileName?: string;
  setSampCorpFileName?: (name: string) => void;
  sampIndData?: any[];
  setSampIndData?: (data: any[]) => void;
  sampIndFileName?: string;
  setSampIndFileName?: (name: string) => void;
  [key: string]: any;
}

export const SamplingSelection: React.FC<SamplingSelectionProps> = ({
  mainData = [],
  columns = [],
  mapping,
  sampCorpData = [],
  setSampCorpData,
}) => {
  const [sampleMethod, setSampleMethod] = useState<"random_percent" | "random_count" | "stratified">("random_percent");
  const [samplePercent, setSamplePercent] = useState<number>(10);
  const [sampleCount, setSampleCount] = useState<number>(100);
  const [strataCol, setStrataCol] = useState<string>(() => mapping?.xa || mapping?.manganh || columns[0] || "");
  const [sampleResults, setSampleResults] = useState<any[]>(sampCorpData);

  const handleDrawSample = () => {
    if (mainData.length === 0) {
      alert("Không có dữ liệu trong bảng chính để rút mẫu!");
      return;
    }

    let sampled: any[] = [];

    if (sampleMethod === "random_percent") {
      const targetSize = Math.max(1, Math.round((mainData.length * samplePercent) / 100));
      const shuffled = [...mainData].sort(() => 0.5 - Math.random());
      sampled = shuffled.slice(0, targetSize);
    } else if (sampleMethod === "random_count") {
      const targetSize = Math.min(mainData.length, Math.max(1, sampleCount));
      const shuffled = [...mainData].sort(() => 0.5 - Math.random());
      sampled = shuffled.slice(0, targetSize);
    } else if (sampleMethod === "stratified") {
      if (!strataCol) {
        alert("Vui lòng chọn cột phân tầng!");
        return;
      }
      const groups: Record<string, any[]> = {};
      mainData.forEach(row => {
        const k = String(row[strataCol] || "Chưa phân loại");
        if (!groups[k]) groups[k] = [];
        groups[k].push(row);
      });

      Object.values(groups).forEach(group => {
        const groupSampleCount = Math.max(1, Math.round((group.length * samplePercent) / 100));
        const shuffled = [...group].sort(() => 0.5 - Math.random());
        sampled.push(...shuffled.slice(0, groupSampleCount));
      });
    }

    setSampleResults(sampled);
    if (setSampCorpData) {
      setSampCorpData(sampled);
    }
  };

  const handleExportSampleExcel = () => {
    if (sampleResults.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(sampleResults);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MauChon");
    XLSX.writeFile(wb, `DanhSachMau_${sampleResults.length}_${Date.now()}.xlsx`);
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ phong cách #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <Shuffle className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Thiết Kế &amp; Rút Mẫu Điều Tra Thống Kê (Sampling)
              </span>
              <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                {sampleResults.length > 0 ? `${sampleResults.length.toLocaleString("vi-VN")} mẫu đã chọn` : "Chưa rút mẫu"}
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          {sampleResults.length > 0 && (
            <button
              type="button"
              onClick={handleExportSampleExcel}
              className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer whitespace-nowrap border-0"
              title="Xuất danh sách mẫu đã rút ra file Excel"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Danh Sách Mẫu ({sampleResults.length.toLocaleString("vi-VN")})</span>
            </button>
          )}
        </div>
      </div>

      {/* FORM CẤU HÌNH THAM SỐ RÚT MẪU */}
      <div className="mx-3.5 grid grid-cols-1 md:grid-cols-4 gap-2.5 bg-white p-3 rounded-none border border-sky-200 items-end">
        <div>
          <label className="block text-[11px] font-bold text-slate-700 mb-1">Phương pháp rút mẫu:</label>
          <select
            value={sampleMethod}
            onChange={e => setSampleMethod(e.target.value as any)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
          >
            <option value="random_percent">Ngẫu nhiên theo tỷ lệ (%)</option>
            <option value="random_count">Ngẫu nhiên số lượng cố định (N)</option>
            <option value="stratified">Phân tầng ngẫu nhiên (Stratified)</option>
          </select>
        </div>

        {sampleMethod === "random_percent" || sampleMethod === "stratified" ? (
          <div>
            <label className="block text-[11px] font-bold text-slate-700 mb-1">Tỷ lệ rút mẫu (%):</label>
            <input
              type="number"
              min="1"
              max="100"
              value={samplePercent}
              onChange={e => setSamplePercent(Number(e.target.value))}
              className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
            />
          </div>
        ) : (
          <div>
            <label className="block text-[11px] font-bold text-slate-700 mb-1">Số lượng đơn vị mẫu:</label>
            <input
              type="number"
              min="1"
              max={mainData.length || 1000}
              value={sampleCount}
              onChange={e => setSampleCount(Number(e.target.value))}
              className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
            />
          </div>
        )}

        {sampleMethod === "stratified" && (
          <div>
            <label className="block text-[11px] font-bold text-slate-700 mb-1">Cột biến phân tầng:</label>
            <select
              value={strataCol}
              onChange={e => setStrataCol(e.target.value)}
              className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500 font-mono"
            >
              {columns.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
        )}

        <div>
          <button
            type="button"
            onClick={handleDrawSample}
            className="w-full py-1.5 bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs rounded-none shadow-2xs transition-colors cursor-pointer flex items-center justify-center gap-1.5 border-0"
          >
            <Shuffle className="w-3.5 h-3.5" />
            <span>Rút Mẫu Ngay</span>
          </button>
        </div>
      </div>

      {/* DANH SÁCH MẪU ĐƯỢC RÚT VỚI BẢNG HEADER #286e42 */}
      <div className="mx-3.5 space-y-2">
        <div className="flex items-center justify-between text-xs font-bold text-slate-800">
          <span>
            Quy mô mẫu đã chọn: <strong className="text-sky-800 font-mono font-bold">{sampleResults.length.toLocaleString("vi-VN")}</strong> đơn vị
            {mainData.length > 0 && (
              <span className="text-slate-500 font-normal ml-1">
                ({((sampleResults.length / mainData.length) * 100).toFixed(1)}% tổng số {mainData.length.toLocaleString("vi-VN")} bản ghi)
              </span>
            )}
          </span>
        </div>

        <div className="border border-sky-200 rounded-none overflow-hidden max-h-80 overflow-y-auto bg-white shadow-inner">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="bg-[#286e42] sticky top-0 text-white font-bold border-b border-[#1d4f2f] z-10">
              <tr>
                <th className="py-2 px-3 w-12 text-center border-r border-[#1d4f2f]">#</th>
                {columns.slice(0, 6).map(c => (
                  <th key={c} className="py-2 px-3 border-r border-[#1d4f2f] whitespace-nowrap">{c}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sampleResults.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-8 text-center text-slate-400 italic">
                    Chưa thực hiện rút mẫu. Nhấn &quot;Rút Mẫu Ngay&quot; ở trên để khởi tạo.
                  </td>
                </tr>
              ) : (
                sampleResults.slice(0, 100).map((row, idx) => (
                  <tr key={idx} className={`hover:bg-sky-50/50 ${idx % 2 === 1 ? "bg-slate-50/40" : ""}`}>
                    <td className="py-1.5 px-3 text-center text-slate-500 font-mono text-[11px] border-r border-slate-100">
                      {idx + 1}
                    </td>
                    {columns.slice(0, 6).map(c => (
                      <td key={c} className="py-1.5 px-3 border-r border-slate-100 whitespace-nowrap max-w-xs truncate text-slate-800 font-medium">
                        {row[c] !== undefined && row[c] !== null ? String(row[c]) : ""}
                      </td>
                    ))}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default SamplingSelection;
