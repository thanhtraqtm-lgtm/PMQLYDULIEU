import React, { useState } from "react";
import { Calculator, Play, Download, Sparkles, Plus, ArrowRight, Layers } from "lucide-react";
import * as XLSX from "xlsx";

export const ComplexCalculations: React.FC<any> = (props) => {
  const {
    aggregateFiles = [],
    mainData = [],
    columns = [],
    setMainData,
    saveAppState
  } = props;

  const [calcMode, setCalcMode] = useState<"internal" | "cross_file">("internal");
  const [col1, setCol1] = useState<string>(columns[0] || "");
  const [col2, setCol2] = useState<string>(columns[1] || "");
  const [operator, setOperator] = useState<"+" | "-" | "*" | "/" | "%">("+");
  const [newColName, setNewColName] = useState<string>("KetQua_PhepTinh");
  const [status, setStatus] = useState<string | null>(null);

  const handleExecute = () => {
    if (!col1 || !col2 || !newColName.trim()) {
      alert("Vui lòng chọn đầy đủ 2 cột và đặt tên cho cột kết quả!");
      return;
    }

    if (!mainData || mainData.length === 0) {
      alert("Không có dữ liệu chính để thực hiện tính toán!");
      return;
    }

    const updated = mainData.map((row: any) => {
      const v1 = parseFloat(String(row[col1] || "").replace(/,/g, "")) || 0;
      const v2 = parseFloat(String(row[col2] || "").replace(/,/g, "")) || 0;
      let res = 0;

      switch (operator) {
        case "+": res = v1 + v2; break;
        case "-": res = v1 - v2; break;
        case "*": res = v1 * v2; break;
        case "/": res = v2 !== 0 ? v1 / v2 : 0; break;
        case "%": res = v2 !== 0 ? (v1 / v2) * 100 : 0; break;
      }

      return {
        ...row,
        [newColName.trim()]: operator === "%" ? Number(res.toFixed(2)) : Number(res.toFixed(4))
      };
    });

    if (setMainData) {
      setMainData(updated);
    }
    if (saveAppState) {
      saveAppState();
    }

    setStatus(`Đã tính toán thành công cột "${newColName.trim()}" cho ${updated.length.toLocaleString("vi-VN")} dòng!`);
    setTimeout(() => setStatus(null), 4000);
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-100">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-amber-100 flex items-center justify-center text-amber-600">
            <Calculator className="w-5 h-5" />
          </div>
          <div>
            <h3 className="font-bold text-slate-800 text-base">Tính Toán Phức Hợp & Công Thức Đa Cột</h3>
            <p className="text-xs text-slate-500">
              Tạo cột tính toán mới bằng công thức số học, tỷ trọng phần trăm hoặc ghép nối dữ liệu giữa các cột.
            </p>
          </div>
        </div>
      </div>

      {status && (
        <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800 text-xs font-semibold">
          {status}
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end bg-slate-50/70 p-5 rounded-xl border border-slate-200">
        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1.5">Cột thứ nhất (A):</label>
          <select
            value={col1}
            onChange={(e) => setCol1(e.target.value)}
            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 bg-white font-medium"
          >
            <option value="">-- Chọn cột A --</option>
            {columns.map((c: string) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1.5">Toán tử:</label>
          <select
            value={operator}
            onChange={(e) => setOperator(e.target.value as any)}
            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 bg-white font-bold text-center"
          >
            <option value="+">Cộng (+)</option>
            <option value="-">Trừ (-)</option>
            <option value="*">Nhân (×)</option>
            <option value="/">Chia (÷)</option>
            <option value="%">Tỷ lệ phần trăm (% A/B)</option>
          </select>
        </div>

        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1.5">Cột thứ hai (B):</label>
          <select
            value={col2}
            onChange={(e) => setCol2(e.target.value)}
            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 bg-white font-medium"
          >
            <option value="">-- Chọn cột B --</option>
            {columns.map((c: string) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-xs font-bold text-slate-700 mb-1.5">Tên cột kết quả mới:</label>
          <input
            type="text"
            value={newColName}
            onChange={(e) => setNewColName(e.target.value)}
            placeholder="KetQua_Moi..."
            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 bg-white font-medium focus:ring-2 focus:ring-amber-500/20"
          />
        </div>
      </div>

      <div className="flex justify-end">
        <button
          onClick={handleExecute}
          className="inline-flex items-center gap-2 px-5 py-2.5 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded-lg shadow-sm transition-colors cursor-pointer"
        >
          <Play className="w-4 h-4 fill-white" />
          Thực Hiện Tính Toán & Ghi Vào Bảng Chính
        </button>
      </div>
    </div>
  );
};
export default ComplexCalculations;
