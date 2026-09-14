import React, { useState, useMemo } from "react";
import { AlertTriangle, Filter, Trash2 } from "lucide-react";

export interface StatisticalOutliersProps {
  mainData: any[];
  columns: string[];
  mapping?: any;
  onFilterRows?: (rowIndices: number[]) => void;
  onUpdateMainData?: (cleanedData: any[]) => void;
  [key: string]: any;
}

export const StatisticalOutliers: React.FC<StatisticalOutliersProps> = ({
  mainData = [],
  columns = [],
  mapping,
  onFilterRows,
  onUpdateMainData,
}) => {
  const [selectedCol, setSelectedCol] = useState<string>(columns[0] || "");
  const [method, setMethod] = useState<"iqr" | "zscore">("iqr");
  const [multiplier, setMultiplier] = useState<number>(1.5);

  const { outlierResults, stats } = useMemo(() => {
    if (!selectedCol || mainData.length === 0) return { outlierResults: [], stats: null };

    const validEntries: { val: number; index: number }[] = [];
    mainData.forEach((row, idx) => {
      const raw = row[selectedCol];
      if (raw !== null && raw !== undefined && raw !== "") {
        const num = typeof raw === "number" ? raw : parseFloat(String(raw).replace(/,/g, ""));
        if (!isNaN(num)) {
          validEntries.push({ val: num, index: idx });
        }
      }
    });

    if (validEntries.length < 4) return { outlierResults: [], stats: null };

    const sortedVals = [...validEntries].sort((a, b) => a.val - b.val);
    const n = sortedVals.length;

    let outliers: { index: number; val: number; reason: string }[] = [];

    if (method === "iqr") {
      const q1 = sortedVals[Math.floor(n * 0.25)].val;
      const median = sortedVals[Math.floor(n * 0.5)].val;
      const q3 = sortedVals[Math.floor(n * 0.75)].val;
      const iqr = q3 - q1;
      const lowerBound = q1 - multiplier * iqr;
      const upperBound = q3 + multiplier * iqr;

      validEntries.forEach(item => {
        if (item.val < lowerBound) {
          outliers.push({
            index: item.index,
            val: item.val,
            reason: `Nhỏ hơn ngưỡng dưới (${lowerBound.toFixed(2)})`,
          });
        } else if (item.val > upperBound) {
          outliers.push({
            index: item.index,
            val: item.val,
            reason: `Lớn hơn ngưỡng trên (${upperBound.toFixed(2)})`,
          });
        }
      });

      return {
        outlierResults: outliers,
        stats: { q1, median, q3, iqr, lowerBound, upperBound },
      };
    } else {
      const sum = validEntries.reduce((acc, curr) => acc + curr.val, 0);
      const mean = sum / n;
      const variance = validEntries.reduce((acc, curr) => acc + Math.pow(curr.val - mean, 2), 0) / n;
      const stdDev = Math.sqrt(variance);

      if (stdDev === 0) return { outlierResults: [], stats: { mean, stdDev } };

      validEntries.forEach(item => {
        const z = Math.abs((item.val - mean) / stdDev);
        if (z > multiplier) {
          outliers.push({
            index: item.index,
            val: item.val,
            reason: `Z-score (${z.toFixed(2)}) vượt quá ngưỡng ${multiplier}`,
          });
        }
      });

      return {
        outlierResults: outliers,
        stats: { mean, stdDev },
      };
    }
  }, [mainData, selectedCol, method, multiplier]);

  const handleFilterToViewer = () => {
    if (!onFilterRows || outlierResults.length === 0) return;
    const indices = outlierResults.map(o => o.index);
    onFilterRows(indices);
  };

  const handleRemoveOutliers = () => {
    if (!onUpdateMainData || outlierResults.length === 0) return;
    if (confirm(`Bạn có chắc muốn xóa ${outlierResults.length} bản ghi dị biệt này khỏi dữ liệu chính?`)) {
      const outlierSet = new Set(outlierResults.map(o => o.index));
      const cleaned = mainData.filter((_, idx) => !outlierSet.has(idx));
      onUpdateMainData(cleaned);
      alert(`Đã xóa ${outlierResults.length} dòng dị biệt thành công!`);
    }
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ chuẩn #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <AlertTriangle className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Bộ Quét Giá Trị Dị Biệt &amp; Ngoại Lai (Outliers)
              </span>
              <span className="text-[11px] font-mono text-rose-900 bg-rose-200/80 px-2 py-0.5 rounded-none font-bold border border-rose-300">
                {outlierResults.length} dị biệt phát hiện
              </span>
            </div>
          </div>
        </div>

        {outlierResults.length > 0 && (
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={handleFilterToViewer}
              className="bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer border-0"
            >
              <Filter className="w-3.5 h-3.5" />
              <span>Xem {outlierResults.length} Dòng Dị Biệt</span>
            </button>
            <button
              type="button"
              onClick={handleRemoveOutliers}
              className="bg-rose-700 hover:bg-rose-800 text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer border-0"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Xóa Các Dòng Dị Biệt</span>
            </button>
          </div>
        )}
      </div>

      {/* THANH ĐIỀU KHIỂN & CẤU HÌNH THAM SỐ */}
      <div className="mx-3.5 grid grid-cols-1 md:grid-cols-3 gap-3 bg-white p-3 rounded-none border border-sky-200">
        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1">Cột số cần kiểm tra dị biệt:</label>
          <select
            value={selectedCol}
            onChange={e => setSelectedCol(e.target.value)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-mono font-medium focus:outline-none focus:border-sky-500"
          >
            {columns.map(c => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1">Phương pháp kiểm định:</label>
          <select
            value={method}
            onChange={e => setMethod(e.target.value as any)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
          >
            <option value="iqr">Khoảng tứ phân vị IQR (Khuyên dùng)</option>
            <option value="zscore">Độ lệch chuẩn Z-Score</option>
          </select>
        </div>

        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1">
            Hệ số độ nhạy ({method === "iqr" ? "k × IQR" : "Z-threshold"}):
          </label>
          <input
            type="number"
            step="0.1"
            min="0.5"
            max="10"
            value={multiplier}
            onChange={e => setMultiplier(parseFloat(e.target.value) || 1.5)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
          />
        </div>
      </div>

      {/* THÔNG TIN THỐNG KÊ & BẢNG KẾT QUẢ VỚI HEADER #286e42 */}
      <div className="mx-3.5 space-y-2">
        <div className="flex items-center justify-between text-xs font-semibold text-slate-700">
          <span>
            Phát hiện: <strong className="text-rose-700 font-bold">{outlierResults.length}</strong> bản ghi dị biệt
          </span>
          {stats && (
            <span className="text-slate-500 font-mono text-[11px]">
              {"median" in stats
                ? `Trung vị: ${stats.median.toLocaleString("vi-VN")} | IQR: ${stats.iqr.toLocaleString("vi-VN")} [${stats.lowerBound.toFixed(1)} đến ${stats.upperBound.toFixed(1)}]`
                : `Trung bình: ${stats.mean.toLocaleString("vi-VN")} | Độ lệch chuẩn: ${stats.stdDev.toFixed(2)}`}
            </span>
          )}
        </div>

        <div className="border border-sky-200 rounded-none overflow-hidden max-h-80 overflow-y-auto bg-white shadow-inner">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="bg-[#286e42] sticky top-0 text-white font-bold border-b border-[#1d4f2f] z-10">
              <tr>
                <th className="py-2 px-3 w-16 text-center border-r border-[#1d4f2f]">DÒNG #</th>
                <th className="py-2 px-3 border-r border-[#1d4f2f]">GIÁ TRỊ CỘT ({selectedCol})</th>
                <th className="py-2 px-3">CĂN CỨ CHẨN ĐOÁN</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {outlierResults.length === 0 ? (
                <tr>
                  <td colSpan={3} className="py-8 text-center text-slate-400 italic">
                    Không phát hiện giá trị dị biệt nào với ngưỡng hiện tại.
                  </td>
                </tr>
              ) : (
                outlierResults.slice(0, 150).map((item, idx) => (
                  <tr key={idx} className="hover:bg-rose-50/50">
                    <td className="py-1.5 px-3 text-center text-slate-500 font-mono border-r border-slate-100">
                      {item.index + 1}
                    </td>
                    <td className="py-1.5 px-3 font-bold text-rose-700 font-mono border-r border-slate-100">
                      {item.val.toLocaleString("vi-VN")}
                    </td>
                    <td className="py-1.5 px-3 text-slate-700 font-medium">{item.reason}</td>
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

export default StatisticalOutliers;
