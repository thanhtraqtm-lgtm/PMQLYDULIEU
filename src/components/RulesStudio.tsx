import React, { useState } from "react";
import { Sliders, Plus, Trash2, CheckCircle2, AlertTriangle, Filter } from "lucide-react";

interface RulesStudioProps {
  mainData: any[];
  columns: string[];
  mapping?: any;
  onFilterRows?: (indices: number[]) => void;
  onUpdateMainData?: (newData: any[]) => void;
}

interface RuleItem {
  id: string;
  name: string;
  col: string;
  condition: "not_empty" | "is_number" | "greater_than" | "less_than" | "len_equals" | "regex";
  param: string;
  description: string;
}

export const RulesStudio: React.FC<RulesStudioProps> = ({
  mainData = [],
  columns = [],
  mapping,
  onFilterRows,
}) => {
  const [rules, setRules] = useState<RuleItem[]>([
    {
      id: "r1",
      name: "Kiểm tra doanh thu âm",
      col: mapping?.doanhthu || columns[0] || "",
      condition: "less_than",
      param: "0",
      description: "Phát hiện các bản ghi có Doanh thu < 0"
    },
    {
      id: "r2",
      name: "Mã ngành chưa đủ 5 ký tự",
      col: mapping?.manganh || columns[1] || "",
      condition: "len_equals",
      param: "5",
      description: "Mã ngành cấp 5 phải chuẩn 5 số"
    }
  ]);

  const [ruleName, setRuleName] = useState("");
  const [selectedCol, setSelectedCol] = useState(columns[0] || "");
  const [condition, setCondition] = useState<RuleItem["condition"]>("not_empty");
  const [param, setParam] = useState("");

  const handleAddRule = () => {
    if (!selectedCol) return;
    const newRule: RuleItem = {
      id: "rule_" + Date.now(),
      name: ruleName.trim() || `Kiểm tra cột ${selectedCol}`,
      col: selectedCol,
      condition,
      param: param.trim(),
      description: `${selectedCol} [${condition}] ${param}`
    };
    setRules(prev => [...prev, newRule]);
    setRuleName("");
    setParam("");
  };

  const checkViolations = (r: RuleItem): number[] => {
    const violatedIndices: number[] = [];

    mainData.forEach((row, idx) => {
      const val = row[r.col];
      const strVal = val !== null && val !== undefined ? String(val).trim() : "";
      const numVal = parseFloat(strVal.replace(/,/g, ""));

      let isViolation = false;
      switch (r.condition) {
        case "not_empty":
          isViolation = strVal === "";
          break;
        case "is_number":
          isViolation = isNaN(numVal);
          break;
        case "greater_than":
          isViolation = !isNaN(numVal) && numVal <= parseFloat(r.param);
          break;
        case "less_than":
          isViolation = !isNaN(numVal) && numVal >= parseFloat(r.param);
          break;
        case "len_equals":
          isViolation = strVal.length !== parseInt(r.param);
          break;
        case "regex":
          try {
            const re = new RegExp(r.param);
            isViolation = !re.test(strVal);
          } catch {
            isViolation = false;
          }
          break;
      }

      if (isViolation) {
        violatedIndices.push(idx);
      }
    });

    return violatedIndices;
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ phong cách #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <Sliders className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Rules Studio - Thiết Kế &amp; Kiểm Định Luật Nghiệp Vụ
              </span>
              <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                {rules.length} quy tắc đang thiết lập
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* FORM THÊM LUẬT MỚI */}
      <div className="mx-3.5 grid grid-cols-1 md:grid-cols-4 gap-2.5 bg-white p-3 rounded-none border border-sky-200 items-end">
        <div>
          <label className="block text-[11px] font-bold text-slate-700 mb-1">Tên quy tắc kiểm tra:</label>
          <input
            type="text"
            value={ruleName}
            onChange={e => setRuleName(e.target.value)}
            placeholder="Ví dụ: Kiểm tra MST hợp lệ..."
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
          />
        </div>

        <div>
          <label className="block text-[11px] font-bold text-slate-700 mb-1">Cột áp dụng:</label>
          <select
            value={selectedCol}
            onChange={e => setSelectedCol(e.target.value)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-mono font-medium focus:outline-none focus:border-sky-500"
          >
            {columns.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-[11px] font-bold text-slate-700 mb-1">Điều kiện kiểm tra:</label>
          <select
            value={condition}
            onChange={e => setCondition(e.target.value as any)}
            className="w-full text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
          >
            <option value="not_empty">Không được để trống</option>
            <option value="is_number">Phải là dạng số</option>
            <option value="greater_than">Phải lớn hơn (&gt;)</option>
            <option value="less_than">Phải nhỏ hơn (&lt;)</option>
            <option value="len_equals">Độ dài ký tự bằng</option>
            <option value="regex">Khớp Regex</option>
          </select>
        </div>

        <div className="flex gap-2">
          {condition !== "not_empty" && condition !== "is_number" && (
            <input
              type="text"
              value={param}
              onChange={e => setParam(e.target.value)}
              placeholder="Giá trị tham số..."
              className="flex-1 text-xs border border-slate-300 rounded-none px-2.5 py-1.5 bg-white font-medium focus:outline-none focus:border-sky-500"
            />
          )}
          <button
            type="button"
            onClick={handleAddRule}
            className="px-3.5 py-1.5 bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs rounded-none transition-colors shadow-2xs cursor-pointer shrink-0 border-0 flex items-center gap-1"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Thêm Luật</span>
          </button>
        </div>
      </div>

      {/* DANH SÁCH LUẬT VỚI BẢNG HEADER #286e42 */}
      <div className="mx-3.5 space-y-2">
        <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide">
          Danh Sách Luật Nghiệp Vụ Đang Áp Dụng ({rules.length})
        </h4>

        <div className="divide-y divide-slate-100 border border-sky-200 rounded-none bg-white shadow-2xs">
          {rules.map((rule) => {
            const violations = checkViolations(rule);
            return (
              <div key={rule.id} className="p-3 flex flex-wrap items-center justify-between gap-3 hover:bg-sky-50/40">
                <div>
                  <h5 className="font-bold text-slate-900 text-xs flex items-center gap-2">
                    {rule.name}
                    <span className="text-[11px] font-mono text-sky-800 bg-sky-100 px-1.5 py-0.2 border border-sky-200">
                      {rule.col}
                    </span>
                  </h5>
                  <p className="text-[11px] text-slate-500 mt-0.5 font-mono">{rule.description}</p>
                </div>

                <div className="flex items-center gap-2.5">
                  {violations.length === 0 ? (
                    <span className="inline-flex items-center gap-1 text-xs font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 border border-emerald-200 rounded-none">
                      <CheckCircle2 className="w-3.5 h-3.5" />
                      100% Hợp lệ
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1 text-xs font-bold text-rose-700 bg-rose-50 px-2 py-0.5 border border-rose-200 rounded-none">
                      <AlertTriangle className="w-3.5 h-3.5" />
                      {violations.length.toLocaleString("vi-VN")} vi phạm
                    </span>
                  )}

                  {violations.length > 0 && onFilterRows && (
                    <button
                      type="button"
                      onClick={() => onFilterRows(violations)}
                      className="px-2.5 py-1 text-xs bg-sky-50 text-sky-800 hover:bg-sky-100 border border-sky-300 font-bold rounded-none cursor-pointer flex items-center gap-1"
                    >
                      <Filter className="w-3 h-3" />
                      <span>Lọc bản ghi</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => setRules(prev => prev.filter(r => r.id !== rule.id))}
                    className="p-1 text-slate-400 hover:text-rose-600 cursor-pointer border-0 bg-transparent"
                    title="Xóa quy tắc này"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

export default RulesStudio;
