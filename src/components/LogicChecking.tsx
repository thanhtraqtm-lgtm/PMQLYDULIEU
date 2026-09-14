import React, { useState, useMemo } from "react";
import { ShieldAlert, Download, AlertCircle, CheckCircle2 } from "lucide-react";
import * as XLSX from "xlsx";
import { normalizeSectorCode, vsicRawData } from "../data/vsic";

interface LogicCheckingProps {
  mainData: any[];
  columns: string[];
  setMainData?: (data: any[]) => void;
  setColumns?: (cols: string[]) => void;
  fileName?: string;
  setFileName?: (name: string) => void;
  mapping?: any;
  onExportExcel?: () => void;
  saveAppState?: any;
  setActiveTab?: (tab: string) => void;
  setLoading?: (loading: boolean) => void;
  setProgress?: (progress: number) => void;
  setStatusMessage?: (msg: string) => void;
  [key: string]: any;
}

export const LogicChecking: React.FC<LogicCheckingProps> = ({
  mainData = [],
  mapping,
}) => {
  const [selectedIssueType, setSelectedIssueType] = useState<string>("all");

  const issues = useMemo(() => {
    const list: {
      index: number;
      type: string;
      title: string;
      detail: string;
      severity: "error" | "warning";
      row: any;
    }[] = [];

    const dtCol = mapping?.doanhthu;
    const ldCol = mapping?.laodong;
    const mstCol = mapping?.idCol;
    const mngCol = mapping?.manganh;

    mainData.forEach((row, idx) => {
      // 1. Kiểm tra Doanh thu và Lao động
      if (dtCol && ldCol) {
        const dt = parseFloat(String(row[dtCol] || "").replace(/,/g, "")) || 0;
        const ld = parseFloat(String(row[ldCol] || "").replace(/,/g, "")) || 0;

        if (dt > 1000000000 && ld === 0) {
          list.push({
            index: idx,
            type: "dt_ld_conflict",
            title: "Doanh thu lớn nhưng không có lao động",
            detail: `Doanh thu = ${dt.toLocaleString("vi-VN")} nhưng Lao động = 0`,
            severity: "warning",
            row,
          });
        } else if (dt === 0 && ld > 20) {
          list.push({
            index: idx,
            type: "dt_ld_conflict",
            title: "Nhiều lao động nhưng doanh thu bằng 0",
            detail: `Lao động = ${ld} nhưng Doanh thu = 0`,
            severity: "warning",
            row,
          });
        }
      }

      // 2. Kiểm tra Mã số thuế (MST)
      if (mstCol && row[mstCol]) {
        const mst = String(row[mstCol]).trim();
        const cleanMst = mst.replace(/-/g, "");
        if (cleanMst.length > 0 && cleanMst.length !== 10 && cleanMst.length !== 13) {
          list.push({
            index: idx,
            type: "invalid_mst",
            title: "Mã số thuế không đúng chuẩn 10 hoặc 13 số",
            detail: `MST: "${mst}" (${cleanMst.length} ký tự)`,
            severity: "error",
            row,
          });
        }
      }

      // 3. Kiểm tra Mã ngành cấp 5
      if (mngCol && row[mngCol]) {
        const mng = normalizeSectorCode(row[mngCol]);
        if (mng && !vsicRawData[mng] && mng.length >= 4 && !vsicRawData[mng.slice(0, 4)]) {
          list.push({
            index: idx,
            type: "invalid_mng",
            title: "Mã ngành không tồn tại trong danh mục VSIC",
            detail: `Mã ngành: "${mng}"`,
            severity: "error",
            row,
          });
        }
      }
    });

    return list;
  }, [mainData, mapping]);

  const filteredIssues = useMemo(() => {
    if (selectedIssueType === "all") return issues;
    return issues.filter(i => i.type === selectedIssueType);
  }, [issues, selectedIssueType]);

  const handleExportIssues = () => {
    if (issues.length === 0) return;
    const rows = issues.map(i => ({
      "Dòng": i.index + 1,
      "Mức độ": i.severity === "error" ? "Lỗi nghiêm trọng" : "Cảnh báo",
      "Vấn đề": i.title,
      "Chi tiết": i.detail,
      ...i.row
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "LoiLogic");
    XLSX.writeFile(wb, `KiemTraLogic_${Date.now()}.xlsx`);
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ phong cách #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <ShieldAlert className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Kiểm Tra Mối Quan Hệ Logic Dữ Liệu Điều Tra
              </span>
              <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                {issues.length} mâu thuẫn / cảnh báo phát hiện
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleExportIssues}
            disabled={issues.length === 0}
            className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer whitespace-nowrap border-0 disabled:opacity-50"
            title="Xuất danh sách lỗi logic ra Excel"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Xuất Danh Sách Lỗi ({issues.length})</span>
          </button>
        </div>
      </div>

      {/* CÁC THẺ PHÂN LOẠI SỐ LIỆU */}
      <div className="mx-3.5 grid grid-cols-1 sm:grid-cols-4 gap-2.5">
        <button
          type="button"
          onClick={() => setSelectedIssueType("all")}
          className={`p-2.5 rounded-none border text-left transition-colors cursor-pointer ${
            selectedIssueType === "all" ? "bg-[#286e42] text-white border-[#1d4f2f] shadow-2xs" : "bg-white border-slate-300 hover:bg-slate-50"
          }`}
        >
          <p className="text-[11px] font-bold opacity-80 uppercase tracking-wide">Tổng phát hiện</p>
          <p className="text-lg font-bold font-mono mt-0.5">{issues.length.toLocaleString("vi-VN")}</p>
        </button>

        <button
          type="button"
          onClick={() => setSelectedIssueType("dt_ld_conflict")}
          className={`p-2.5 rounded-none border text-left transition-colors cursor-pointer ${
            selectedIssueType === "dt_ld_conflict" ? "bg-amber-600 text-white border-amber-700 shadow-2xs" : "bg-white border-slate-300 hover:bg-slate-50"
          }`}
        >
          <p className="text-[11px] font-bold opacity-80 uppercase tracking-wide">Mâu thuẫn DT - LĐ</p>
          <p className="text-lg font-bold font-mono mt-0.5 text-amber-700 group-hover:text-amber-800">
            {issues.filter(i => i.type === "dt_ld_conflict").length.toLocaleString("vi-VN")}
          </p>
        </button>

        <button
          type="button"
          onClick={() => setSelectedIssueType("invalid_mst")}
          className={`p-2.5 rounded-none border text-left transition-colors cursor-pointer ${
            selectedIssueType === "invalid_mst" ? "bg-rose-700 text-white border-rose-800 shadow-2xs" : "bg-white border-slate-300 hover:bg-slate-50"
          }`}
        >
          <p className="text-[11px] font-bold opacity-80 uppercase tracking-wide">Lỗi Mã số thuế</p>
          <p className="text-lg font-bold font-mono mt-0.5 text-rose-700">
            {issues.filter(i => i.type === "invalid_mst").length.toLocaleString("vi-VN")}
          </p>
        </button>

        <button
          type="button"
          onClick={() => setSelectedIssueType("invalid_mng")}
          className={`p-2.5 rounded-none border text-left transition-colors cursor-pointer ${
            selectedIssueType === "invalid_mng" ? "bg-indigo-600 text-white border-indigo-700 shadow-2xs" : "bg-white border-slate-300 hover:bg-slate-50"
          }`}
        >
          <p className="text-[11px] font-bold opacity-80 uppercase tracking-wide">Mã ngành ngoài VSIC</p>
          <p className="text-lg font-bold font-mono mt-0.5 text-indigo-700">
            {issues.filter(i => i.type === "invalid_mng").length.toLocaleString("vi-VN")}
          </p>
        </button>
      </div>

      {/* BẢNG DANH SÁCH LỖI VỚI HEADER #286e42 */}
      <div className="mx-3.5 border border-sky-200 rounded-none overflow-hidden max-h-[500px] overflow-y-auto bg-white shadow-inner">
        <table className="w-full text-left border-collapse text-xs">
          <thead className="bg-[#286e42] sticky top-0 text-white font-bold border-b border-[#1d4f2f] z-10">
            <tr>
              <th className="py-2 px-3 w-14 text-center border-r border-[#1d4f2f]">DÒNG #</th>
              <th className="py-2 px-3 w-28 border-r border-[#1d4f2f]">MỨC ĐỘ</th>
              <th className="py-2 px-3 w-72 border-r border-[#1d4f2f]">LOẠI VẤN ĐỀ</th>
              <th className="py-2 px-3">CHI TIẾT MÂU THUẪN</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filteredIssues.length === 0 ? (
              <tr>
                <td colSpan={4} className="py-10 text-center text-slate-400 italic">
                  Tuyệt vời! Không phát hiện mâu thuẫn logic nào trong dữ liệu.
                </td>
              </tr>
            ) : (
              filteredIssues.slice(0, 150).map((item, idx) => (
                <tr key={idx} className={`hover:bg-sky-50/50 ${idx % 2 === 1 ? "bg-slate-50/40" : ""}`}>
                  <td className="py-1.5 px-3 text-center text-slate-500 font-mono text-[11px] border-r border-slate-100">
                    {item.index + 1}
                  </td>
                  <td className="py-1.5 px-3 border-r border-slate-100">
                    <span
                      className={`inline-block px-1.5 py-0.5 rounded-none text-[10px] font-bold border ${
                        item.severity === "error" ? "bg-rose-100 text-rose-800 border-rose-300" : "bg-amber-100 text-amber-800 border-amber-300"
                      }`}
                    >
                      {item.severity === "error" ? "LỖI" : "CẢNH BÁO"}
                    </span>
                  </td>
                  <td className="py-1.5 px-3 font-semibold text-slate-800 border-r border-slate-100">
                    {item.title}
                  </td>
                  <td className="py-1.5 px-3 text-slate-700 font-mono text-[11px]">
                    {item.detail}
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

export default LogicChecking;
