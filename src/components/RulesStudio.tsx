import React, { useState, useEffect, useMemo } from "react";
import { 
  CheckSquare, 
  AlertTriangle, 
  Download, 
  Upload,
  Search, 
  Trash2, 
  Plus, 
  Save, 
  Sliders, 
  Zap, 
  CheckCircle2, 
  Brain, 
  Info,
  ChevronDown,
  ChevronUp,
  SlidersHorizontal,
  FileCheck,
  FolderOpen,
  Sparkles,
  Loader2
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import { LiveTenRowPreview } from "./LiveTenRowPreview";

interface ColumnMapping {
  mota: string;
  manganh: string;
  xa: string;
  doanhthu: string;
  laodong: string;
  idCol: string;
}

interface RulesStudioProps {
  mainData: any[];
  columns: string[];
  mapping: ColumnMapping;
  onFilterRows?: (indices: number[], label: string) => void;
  onExportExcel?: (data: any[], fileName: string) => void;
  onUpdateMainData?: (newData: any[]) => void;
}

interface LogicRuleCondition {
  col: string;
  op: "==" | "!=" | ">" | "<" | ">=" | "<=" | "chứa" | "không chứa" | "trống" | "không trống";
  val: string;
  isFieldCompare: boolean;
}

interface StudioRule {
  id: string;
  name: string;
  type: "manual" | "ai_expression";
  active: boolean;
  category?: string;
  // For manual
  ifRules?: LogicRuleCondition[];
  thenRules?: LogicRuleCondition[];
  ifCombine?: "AND" | "OR";
  thenCombine?: "AND" | "OR";
  logicRuleMode?: "must_satisfy" | "conflict";
  // For AI expression
  prompt?: string;
  expression?: string;
}

// Global evaluation helper for condition checking
const checkValue = (rowVal: any, op: string, compareVal: string) => {
  const v1 = String(rowVal !== undefined && rowVal !== null ? rowVal : "").trim();
  const v2 = String(compareVal).trim();

  const v1LC = v1.toLowerCase();
  const v2LC = v2.toLowerCase();

  if (op === "trống") return v1 === "";
  if (op === "không trống") return v1 !== "";

  // Extract digits for numeric comparisons
  const cleanV1 = v1.replace(/,/g, "").replace(/[^0-9.\-]/g, "");
  const cleanV2 = v2.replace(/,/g, "").replace(/[^0-9.\-]/g, "");

  const num1 = v1 === "" ? 0 : parseFloat(cleanV1);
  const num2 = parseFloat(cleanV2);

  const isNum2 = !isNaN(num2) && cleanV2 !== "";
  const isNum1 = !isNaN(num1) && (v1 === "" || cleanV1 !== "");

  if (isNum2 && isNum1) {
    if (op === "==") return num1 === num2;
    if (op === "!=") return num1 !== num2;
    if (op === ">") return num1 > num2;
    if (op === "<") return num1 < num2;
    if (op === ">=") return num1 >= num2;
    if (op === "<=") return num1 <= num2;
  }

  if (op === "==") return v1LC === v2LC;
  if (op === "!=") return v1LC !== v2LC;
  if (op === "chứa") return v1LC.includes(v2LC);
  if (op === "không chứa") return !v1LC.includes(v2LC);

  if (op === ">") return v1LC > v2LC;
  if (op === "<") return v1LC < v2LC;
  if (op === ">=") return v1LC >= v2LC;
  if (op === "<=") return v1LC <= v2LC;

  return false;
};

export default function RulesStudio({
  mainData,
  columns,
  mapping,
  onFilterRows,
  onExportExcel,
  onUpdateMainData
}: RulesStudioProps) {
  // Master Rules List State
  const [rules, setRules] = useState<StudioRule[]>(() => {
    const defaultRules: StudioRule[] = [
      {
        id: "r_1",
        name: "Doanh thu rỗng hoặc bằng 0 nhưng có ghi nhận Lao động",
        type: "manual",
        active: true,
        category: "Doanh thu & Lao động",
        ifRules: [
          { col: mapping.doanhthu || "Doanh thu", op: "trống", val: "", isFieldCompare: false }
        ],
        thenRules: [
          { col: mapping.laodong || "Lao động", op: "==", val: "0", isFieldCompare: false }
        ],
        ifCombine: "AND",
        thenCombine: "AND",
        logicRuleMode: "must_satisfy"
      },
      {
        id: "r_2",
        name: "Ngành nông nghiệp nhưng Doanh thu cực đại bất thường",
        type: "ai_expression",
        active: true,
        category: "Dữ liệu biên",
        prompt: "Ngành bắt đầu bằng mã 01 nhưng doanh thu vượt quá 1000 (triệu đồng)",
        expression: "String(row['" + (mapping.manganh || "Mã ngành") + "'] || '').startsWith('01') && parseFloat(String(row['" + (mapping.doanhthu || "Doanh thu") + "'] || '0').replace(/,/g, '')) > 1000"
      },
      {
        id: "r_3",
        name: "Địa bàn xã bị bỏ trống thông tin",
        type: "manual",
        active: true,
        category: "Cơ cấu hành chính",
        ifRules: [
          { col: mapping.xa || "Xã/Phường", op: "trống", val: "", isFieldCompare: false }
        ],
        thenRules: [],
        ifCombine: "AND",
        thenCombine: "AND",
        logicRuleMode: "must_satisfy"
      },
      {
        id: "r_4",
        name: "Lệch độ tuổi đi học so với lớp học (Từ Mầm non đến Đại học)",
        type: "ai_expression",
        active: true,
        category: "Dân số & Giáo dục",
        prompt: "Quét mâu thuẫn giữa cột Tuổi/Age và cột Lớp/Grade của học sinh (Hỗ trợ từ mầm non đến đại học)",
        expression: "(() => { const tuoi = parseInt(String(row['Tuổi'] || row['tuổi'] || row['Age'] || row['age'] || '0').replace(/[^0-9]/g, ''), 10); const lop = String(row['Lớp'] || row['lớp'] || row['Grade'] || row['grade'] || '').trim().toLowerCase(); if (tuoi > 0 && lop) { if (lop.includes('mầm') || lop.includes('mẫu giáo') || lop.includes('nhà trẻ') || lop.includes('chồi') || lop.includes('lá')) { return tuoi < 1 || tuoi > 6; } const lopMatch = lop.match(/\\b(1[0-2]|[1-9])\\b/); if (lopMatch) { const soLop = parseInt(lopMatch[1], 10); return tuoi < (soLop + 5) || tuoi > (soLop + 8); } if (lop.includes('đại học') || lop.includes('cao đẳng') || lop.includes('trung cấp') || lop.includes('sinh viên') || lop.includes('sv')) { if (tuoi < 17) return true; if (lop.includes('năm 1') || lop.includes('nam 1')) return tuoi < 17 || tuoi > 23; if (lop.includes('năm 2') || lop.includes('nam 2')) return tuoi < 18 || tuoi > 24; if (lop.includes('năm 3') || lop.includes('nam 3')) return tuoi < 19 || tuoi > 25; if (lop.includes('năm 4') || lop.includes('nam 4')) return tuoi < 20 || tuoi > 26; } } return false; })()"
      }
    ];

    try {
      const saved = localStorage.getItem("vsic_studio_rules");
      if (saved) {
        const parsed = JSON.parse(saved) as StudioRule[];
        if (Array.isArray(parsed)) {
          // If any of our default rules are missing from the parsed array, append them
          const parsedIds = new Set(parsed.map(r => r.id));
          const missingRules = defaultRules.filter(r => !parsedIds.has(r.id));
          if (missingRules.length > 0) {
            const merged = [...parsed, ...missingRules];
            localStorage.setItem("vsic_studio_rules", JSON.stringify(merged));
            return merged;
          }
          return parsed;
        }
      }
    } catch (e) {}
    
    return defaultRules;
  });

  // Saved scenarios/templates state
  const [templates, setTemplates] = useState<{ name: string; rules: StudioRule[] }[]>(() => {
    try {
      const saved = localStorage.getItem("vsic_studio_templates");
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return [
      {
        name: "Kịch bản Kiểm tra Doanh nghiệp cơ bản",
        rules: [
          {
            id: "r_1",
            name: "Doanh thu rỗng hoặc bằng 0 nhưng có ghi nhận Lao động",
            type: "manual",
            active: true,
            ifRules: [{ col: "Doanh thu", op: "trống", val: "", isFieldCompare: false }],
            thenRules: [{ col: "Lao động", op: "==", val: "0", isFieldCompare: false }],
            ifCombine: "AND",
            thenCombine: "AND",
            logicRuleMode: "must_satisfy"
          }
        ]
      },
      {
        name: "Kịch bản Kiểm tra Nông nghiệp - Thủy sản",
        rules: [
          {
            id: "r_ag_1",
            name: "Diện tích trồng trọt bằng 0 hoặc rỗng nhưng có Sản lượng thu hoạch",
            type: "manual",
            active: true,
            ifRules: [
              { col: "Diện tích gieo trồng", op: "trống", val: "", isFieldCompare: false }
            ],
            thenRules: [
              { col: "Sản lượng thu hoạch", op: "==", val: "0", isFieldCompare: false }
            ],
            ifCombine: "AND",
            thenCombine: "AND",
            logicRuleMode: "must_satisfy"
          },
          {
            id: "r_ag_2",
            name: "Mâu thuẫn: Có chi phí thức ăn nhưng số lượng vật nuôi bằng 0",
            type: "manual",
            active: true,
            ifRules: [
              { col: "Chi phí thức ăn chăn nuôi", op: ">", val: "0", isFieldCompare: false }
            ],
            thenRules: [
              { col: "Số lượng vật nuôi", op: "==", val: "0", isFieldCompare: false }
            ],
            ifCombine: "AND",
            thenCombine: "AND",
            logicRuleMode: "conflict"
          }
        ]
      },
      {
        name: "Kịch bản Kiểm tra Kinh tế Xã hội Hộ gia đình",
        rules: [
          {
            id: "r_se_1",
            name: "Có thu nhập từ sản xuất nông nghiệp nhưng số lao động nông nghiệp bằng 0",
            type: "manual",
            active: true,
            ifRules: [
              { col: "Thu nhập nông nghiệp", op: ">", val: "0", isFieldCompare: false }
            ],
            thenRules: [
              { col: "Lao động nông nghiệp", op: ">", val: "0", isFieldCompare: false }
            ],
            ifCombine: "AND",
            thenCombine: "AND",
            logicRuleMode: "must_satisfy"
          },
          {
            id: "r_se_2",
            name: "Mâu thuẫn: Thu nhập nông nghiệp lớn hơn tổng thu nhập",
            type: "manual",
            active: true,
            ifRules: [
              { col: "Thu nhập nông nghiệp", op: ">", val: "Tổng thu nhập", isFieldCompare: true }
            ],
            thenRules: [
              { col: "Tổng thu nhập", op: "<=", val: "0", isFieldCompare: false }
            ],
            ifCombine: "AND",
            thenCombine: "AND",
            logicRuleMode: "conflict"
          }
        ]
      }
    ];
  });

  const [activeTemplate, setActiveTemplate] = useState<string>("");
  const [newTemplateName, setNewTemplateName] = useState<string>("");

  // Rule Builders UI states
  const [builderTab, setBuilderTab] = useState<"manual" | "ai">("manual");
  const [newRuleName, setNewRuleName] = useState<string>("");
  const [newRuleCategory, setNewRuleCategory] = useState<string>("Rà soát chung");

  // Manual Builder states
  const [mIfRules, setMIfRules] = useState<LogicRuleCondition[]>([{ col: "", op: "==", val: "", isFieldCompare: false }]);
  const [mThenRules, setMThenRules] = useState<LogicRuleCondition[]>([]);
  const [mIfCombine, setMIfCombine] = useState<"AND" | "OR">("AND");
  const [mThenCombine, setMThenCombine] = useState<"AND" | "OR">("AND");
  const [mLogicMode, setMLogicMode] = useState<"must_satisfy" | "conflict">("must_satisfy");

  // AI Builder states
  const [aiPrompt, setAiPrompt] = useState<string>("");
  const [isTranslatingAi, setIsTranslatingAi] = useState<boolean>(false);
  const [aiExpression, setAiExpression] = useState<string>("");

  // Scan execution states
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [hasScanned, setHasScanned] = useState<boolean>(false);
  const [scanProgress, setScanProgress] = useState<number>(0);
  const [violationResults, setViolationResults] = useState<any[]>([]); // { rowIdx, row, violations: { ruleId, ruleName, detail }[] }
  const [scanStats, setScanStats] = useState<any>(null);
  const [resultFilter, setResultFilter] = useState<"violated" | "all">("violated");
  const [resultSearch, setResultSearch] = useState<string>("");

  // Save rules and templates to local storage on modification
  useEffect(() => {
    localStorage.setItem("vsic_studio_rules", JSON.stringify(rules));
  }, [rules]);

  useEffect(() => {
    localStorage.setItem("vsic_studio_templates", JSON.stringify(templates));
  }, [templates]);

  // Handle template selection
  const handleLoadTemplate = (tplName: string) => {
    const tpl = templates.find(t => t.name === tplName);
    if (tpl) {
      setRules(tpl.rules);
      setActiveTemplate(tplName);
      setHasScanned(false);
    }
  };

  // Save current active rules as template
  const handleSaveAsTemplate = () => {
    const name = newTemplateName.trim();
    if (!name) {
      alert("Vui lòng nhập tên kịch bản quét!");
      return;
    }

    const exists = templates.some(t => t.name === name);
    if (exists) {
      if (!confirm("Kịch bản quét này đã tồn tại. Bạn có muốn ghi đè không?")) {
        return;
      }
    }

    const updated = [
      ...templates.filter(t => t.name !== name),
      { name, rules }
    ];

    setTemplates(updated);
    setActiveTemplate(name);
    setNewTemplateName("");
    alert(`Lưu kịch bản quét [${name}] thành công!`);
  };

  // Delete saved template
  const handleDeleteTemplate = (tplName: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (confirm(`Bạn có chắc chắn muốn xóa kịch bản quét [${tplName}] không?`)) {
      setTemplates(templates.filter(t => t.name !== tplName));
      if (activeTemplate === tplName) {
        setActiveTemplate("");
      }
    }
  };

  // Import Scenario from JSON
  const handleImportJson = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target?.result as string);
        if (Array.isArray(parsed)) {
          // Rule list format
          setRules(parsed);
          setActiveTemplate("Tệp cấu hình vừa nhập");
          alert("Nhập kịch bản quét thành công!");
        } else if (parsed.name && Array.isArray(parsed.rules)) {
          // Single template format
          setTemplates([parsed, ...templates.filter(t => t.name !== parsed.name)]);
          setRules(parsed.rules);
          setActiveTemplate(parsed.name);
          alert(`Nhập kịch bản quét [${parsed.name}] thành công!`);
        } else {
          alert("Tệp JSON không đúng định dạng kịch bản quét!");
        }
      } catch (err: any) {
        alert("Lỗi phân tích tệp JSON kịch bản quét: " + err.message);
      }
    };
    reader.readAsText(file);
  };

  // Export current rules as JSON
  const handleExportJson = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify({
      name: activeTemplate || "Kịch_bản_quét_tùy_biến",
      rules: rules
    }, null, 2));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", `Kich_Ban_Quet_${(activeTemplate || "Studio").replace(/\s+/g, "_")}.json`);
    dlAnchorElem.click();
  };

  // Add rule toggles
  const handleToggleRule = (id: string) => {
    setRules(rules.map(r => r.id === id ? { ...r, active: !r.active } : r));
  };

  const handleDeleteRule = (id: string) => {
    setRules(rules.filter(r => r.id !== id));
  };

  const handleRestoreDefaults = () => {
    if (confirm("Bạn có chắc chắn muốn khôi phục danh sách quy tắc mặc định ban đầu không? (Các quy tắc do bạn tự viết thêm vẫn sẽ được giữ nguyên, chỉ có các quy tắc hệ thống gốc được bổ sung lại đầy đủ)")) {
      const defaultRules: StudioRule[] = [
        {
          id: "r_1",
          name: "Doanh thu rỗng hoặc bằng 0 nhưng có ghi nhận Lao động",
          type: "manual",
          active: true,
          category: "Doanh thu & Lao động",
          ifRules: [
            { col: mapping.doanhthu || "Doanh thu", op: "trống", val: "", isFieldCompare: false }
          ],
          thenRules: [
            { col: mapping.laodong || "Lao động", op: "==", val: "0", isFieldCompare: false }
          ],
          ifCombine: "AND",
          thenCombine: "AND",
          logicRuleMode: "must_satisfy"
        },
        {
          id: "r_2",
          name: "Ngành nông nghiệp nhưng Doanh thu cực đại bất thường",
          type: "ai_expression",
          active: true,
          category: "Dữ liệu biên",
          prompt: "Ngành bắt đầu bằng mã 01 nhưng doanh thu vượt quá 1000 (triệu đồng)",
          expression: "String(row['" + (mapping.manganh || "Mã ngành") + "'] || '').startsWith('01') && parseFloat(String(row['" + (mapping.doanhthu || "Doanh thu") + "'] || '0').replace(/,/g, '')) > 1000"
        },
        {
          id: "r_3",
          name: "Địa bàn xã bị bỏ trống thông tin",
          type: "manual",
          active: true,
          category: "Cơ cấu hành chính",
          ifRules: [
            { col: mapping.xa || "Xã/Phường", op: "trống", val: "", isFieldCompare: false }
          ],
          thenRules: [],
          ifCombine: "AND",
          thenCombine: "AND",
          logicRuleMode: "must_satisfy"
        },
        {
          id: "r_4",
          name: "Lệch độ tuổi đi học so với lớp học (Từ Mầm non đến Đại học)",
          type: "ai_expression",
          active: true,
          category: "Dân số & Giáo dục",
          prompt: "Quét mâu thuẫn giữa cột Tuổi/Age và cột Lớp/Grade của học sinh (Hỗ trợ từ mầm non đến đại học)",
          expression: "(() => { const tuoi = parseInt(String(row['Tuổi'] || row['tuổi'] || row['Age'] || row['age'] || '0').replace(/[^0-9]/g, ''), 10); const lop = String(row['Lớp'] || row['lớp'] || row['Grade'] || row['grade'] || '').trim().toLowerCase(); if (tuoi > 0 && lop) { if (lop.includes('mầm') || lop.includes('mẫu giáo') || lop.includes('nhà trẻ') || lop.includes('chồi') || lop.includes('lá')) { return tuoi < 1 || tuoi > 6; } const lopMatch = lop.match(/\\b(1[0-2]|[1-9])\\b/); if (lopMatch) { const soLop = parseInt(lopMatch[1], 10); return tuoi < (soLop + 5) || tuoi > (soLop + 8); } if (lop.includes('đại học') || lop.includes('cao đẳng') || lop.includes('trung cấp') || lop.includes('sinh viên') || lop.includes('sv')) { if (tuoi < 17) return true; if (lop.includes('năm 1') || lop.includes('nam 1')) return tuoi < 17 || tuoi > 23; if (lop.includes('năm 2') || lop.includes('nam 2')) return tuoi < 18 || tuoi > 24; if (lop.includes('năm 3') || lop.includes('nam 3')) return tuoi < 19 || tuoi > 25; if (lop.includes('năm 4') || lop.includes('nam 4')) return tuoi < 20 || tuoi > 26; } } return false; })()"
        }
      ];

      const customRules = rules.filter(r => !r.id.startsWith("r_"));
      const newRules = [...defaultRules, ...customRules];
      setRules(newRules);
      alert("Đã khôi phục các quy tắc rà soát hệ thống mặc định!");
    }
  };

  // Manual rule builders helpers
  const handleAddCondition = (type: "if" | "then") => {
    if (type === "if") {
      setMIfRules([...mIfRules, { col: columns[0] || "", op: "==", val: "", isFieldCompare: false }]);
    } else {
      setMThenRules([...mThenRules, { col: columns[0] || "", op: "==", val: "", isFieldCompare: false }]);
    }
  };

  const handleRemoveCondition = (type: "if" | "then", idx: number) => {
    if (type === "if") {
      setMIfRules(mIfRules.filter((_, i) => i !== idx));
    } else {
      setMThenRules(mThenRules.filter((_, i) => i !== idx));
    }
  };

  // Build and save manual rule
  const handleSaveManualRule = () => {
    const name = newRuleName.trim();
    if (!name) {
      alert("Vui lòng đặt tên cho Quy tắc mới!");
      return;
    }
    if (mIfRules.length === 0) {
      alert("Quy tắc vế NẾU không thể bỏ trống!");
      return;
    }

    const newRule: StudioRule = {
      id: "manual_" + Date.now(),
      name,
      category: newRuleCategory,
      type: "manual",
      active: true,
      ifRules: mIfRules.filter(r => r.col),
      thenRules: mThenRules.filter(r => r.col),
      ifCombine: mIfCombine,
      thenCombine: mThenCombine,
      logicRuleMode: mLogicMode
    };

    setRules([newRule, ...rules]);
    setNewRuleName("");
    setMIfRules([{ col: columns[0] || "", op: "==", val: "", isFieldCompare: false }]);
    setMThenRules([]);
    alert("Thêm quy tắc thủ công thành công!");
  };

  // Translate AI prompt with Gemini to Javascript expression
  const handleTranslateAiRule = async () => {
    const prompt = aiPrompt.trim();
    if (!prompt) {
      alert("Vui lòng nhập mô tả ràng buộc bằng ngôn ngữ tự nhiên!");
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Vui lòng cấu hình VITE_GEMINI_API_KEY trong Cài đặt > Khóa bí mật (Secrets) hoặc file .env để kích hoạt trí tuệ nhân tạo Gemini dịch luật!");
      return;
    }

    setIsTranslatingAi(true);
    setAiExpression("");

    try {
      const aiClient = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build"
          }
        }
      });

      const systemPrompt = `Bạn là chuyên gia về cấu trúc dữ liệu và lập trình hàm Javascript.
Nhiệm vụ: Hãy dịch yêu cầu rà soát lỗi dữ liệu của người dùng thành một biểu thức logic JavaScript viết trên một dòng.
Biểu thức logic này PHẢI trả về true nếu dòng dữ liệu bị LỖI (vi phạm ràng buộc), ngược lại trả về false.
Bạn sẽ làm việc trên một đối tượng dòng dữ liệu tên là 'row'. Các cột dữ liệu khả dụng của dòng là: [${columns.filter(c => !c.startsWith("_")).map(c => `'${c}'`).join(", ")}].

Truy cập cột bằng cú pháp chuẩn: row['Tên Cột'].
Nếu so sánh số, hãy ép kiểu và xóa dấu phẩy: parseFloat(String(row['Tên Cột'] || '0').replace(/,/g, ''))
Ví dụ:
1. "Doanh thu < 0" -> parseFloat(String(row['Doanh thu'] || '0').replace(/,/g, '')) < 0
2. "Số lao động bằng 0 nhưng doanh thu lớn hơn 100" -> parseFloat(String(row['Lao động'] || '0').replace(/,/g, '')) == 0 && parseFloat(String(row['Doanh thu'] || '0').replace(/,/g, '')) > 100

Yêu cầu cực kỳ quan trọng:
CHỈ trả về biểu thức logic Javascript trên một dòng duy nhất, không bọc trong markdown (\`\`\`), không giải thích gì thêm.`;

      const response = await aiClient.models.generateContent({
        model: "gemini-3.5-flash",
        contents: `Dịch luật: "${prompt}"`,
        config: {
          systemInstruction: systemPrompt,
          temperature: 0.1
        }
      });

      let expr = (response.text || "").trim();
      expr = expr.replace(/```javascript/gi, "").replace(/```js/gi, "").replace(/```/g, "").trim();
      setAiExpression(expr);
    } catch (err: any) {
      console.error(err);
      alert("Lỗi dịch AI Gemini: " + err.message);
    } finally {
      setIsTranslatingAi(false);
    }
  };

  // Add translated AI rule to rules list
  const handleSaveAiRule = () => {
    const name = newRuleName.trim() || aiPrompt.trim();
    if (!name) {
      alert("Vui lòng đặt tên hoặc nhập mô tả cho quy tắc AI!");
      return;
    }
    if (!aiExpression) {
      alert("Không có biểu thức Javascript tương ứng! Hãy nhấp nút dịch AI trước.");
      return;
    }

    const newRule: StudioRule = {
      id: "ai_" + Date.now(),
      name,
      category: newRuleCategory,
      type: "ai_expression",
      active: true,
      prompt: aiPrompt,
      expression: aiExpression
    };

    setRules([newRule, ...rules]);
    setNewRuleName("");
    setAiPrompt("");
    setAiExpression("");
    alert("Thêm quy tắc AI thành công!");
  };

  // RUN ALL ACTIVE RULES (THE CORE MASSIVE BATCH ENGINE)
  const handleRunStudioScan = () => {
    if (mainData.length === 0) return;
    
    const activeRules = rules.filter(r => r.active);
    if (activeRules.length === 0) {
      alert("Không có quy tắc nào đang được bật để rà quét! Hãy kích hoạt ít nhất 1 quy tắc.");
      return;
    }

    setIsScanning(true);
    setScanProgress(0);
    setViolationResults([]);

    setTimeout(() => {
      try {
        const violationsList: any[] = [];
        let totalViolationsCount = 0;

        // Process line by line
        mainData.forEach((row, rowIdx) => {
          const rowViolations: any[] = [];

          activeRules.forEach(rule => {
            let isViolated = false;
            let noteDetail = "";

            if (rule.type === "manual") {
              // 1. Evaluate If Conditions
              const ifMatches = (rule.ifRules || []).map(cond => {
                const compareValue = cond.isFieldCompare ? String(row[cond.val] || "") : cond.val;
                return checkValue(row[cond.col], cond.op, compareValue);
              });
              const satisfiesIf = rule.ifCombine === "OR" 
                ? ifMatches.some(v => v === true) 
                : ifMatches.every(v => v === true);

              // 2. Evaluate Then Conditions
              if ((rule.thenRules || []).length === 0) {
                // Independent Filter/Filter-only mode: if condition is met, it is a violation (or filter target)
                if (satisfiesIf) {
                  isViolated = true;
                  noteDetail = `Thỏa mãn bộ lọc rà soát: { ${rule.ifRules?.map(r => `[${r.col}] ${r.op} '${r.val}'`).join(" " + rule.ifCombine + " ")} }`;
                }
              } else {
                const thenMatches = (rule.thenRules || []).map(cond => {
                  const compareValue = cond.isFieldCompare ? String(row[cond.val] || "") : cond.val;
                  return checkValue(row[cond.col], cond.op, compareValue);
                });
                const satisfiesThen = rule.thenCombine === "OR"
                  ? thenMatches.some(v => v === true)
                  : thenMatches.every(v => v === true);

                if (rule.logicRuleMode === "conflict") {
                  // Conflict mode: violated if BOTH If and Then are satisfied
                  if (satisfiesIf && satisfiesThen) {
                    isViolated = true;
                    noteDetail = `Mâu thuẫn đồng thời: vế NẾU và vế ĐỒNG THỜI CÓ THÊM đạt thỏa mãn.`;
                  }
                } else {
                  // Default Must Satisfy mode: violated if satisfies If but fails Then
                  if (satisfiesIf && !satisfiesThen) {
                    isViolated = true;
                    noteDetail = `Có thỏa mãn vế NẾU nhưng vi phạm không đạt vế THÌ BẮT BUỘC PHẢI.`;
                  }
                }
              }

            } else if (rule.type === "ai_expression" && rule.expression) {
              // Evaluate JS Expression Rule
              try {
                const fn = new Function("row", `try { return (${rule.expression}); } catch(e) { return false; }`);
                if (fn(row) === true) {
                  isViolated = true;
                  noteDetail = `AI phân tích: Vi phạm biểu thức điều kiện logic.`;
                }
              } catch (e) {
                console.warn("Invalid JS expression for rule", rule.name, e);
              }
            }

            if (isViolated) {
              rowViolations.push({
                ruleId: rule.id,
                ruleName: rule.name,
                detail: noteDetail,
                category: rule.category || "Rà soát chung"
              });
              totalViolationsCount++;
            }
          });

          if (rowViolations.length > 0) {
            violationsList.push({
              rowIdx,
              rowNum: rowIdx + 1,
              row,
              violations: rowViolations
            });
          }

          if (rowIdx % 500 === 0) {
            setScanProgress(Math.round((rowIdx / mainData.length) * 100));
          }
        });

        const totalRows = mainData.length;
        const violatedRowsCount = violationsList.length;
        const passedRowsCount = totalRows - violatedRowsCount;

        setViolationResults(violationsList);
        setScanStats({
          totalRows,
          violatedRows: violatedRowsCount,
          passedRows: passedRowsCount,
          totalViolationsCount,
          violationRate: totalRows > 0 ? ((violatedRowsCount / totalRows) * 100).toFixed(2) + "%" : "0%"
        });
        setHasScanned(true);
        setIsScanning(false);
      } catch (err: any) {
        alert("Lỗi thực hiện rà quét kịch bản: " + err.message);
        setIsScanning(false);
      }
    }, 100);
  };

  // Auto-re-evaluate when mainData changes after a scan has been run
  useEffect(() => {
    if (hasScanned && mainData.length > 0) {
      const activeRules = rules.filter(r => r.active);
      if (activeRules.length > 0) {
        const violationsList: any[] = [];
        let totalViolationsCount = 0;

        mainData.forEach((row, rowIdx) => {
          const rowViolations: any[] = [];

          activeRules.forEach(rule => {
            let isViolated = false;
            let noteDetail = "";

            if (rule.type === "manual") {
              const ifMatches = (rule.ifRules || []).map(cond => {
                const compareValue = cond.isFieldCompare ? String(row[cond.val] || "") : cond.val;
                return checkValue(row[cond.col], cond.op, compareValue);
              });
              const satisfiesIf = rule.ifCombine === "OR" 
                ? ifMatches.some(v => v === true) 
                : ifMatches.every(v => v === true);

              if ((rule.thenRules || []).length === 0) {
                if (satisfiesIf) {
                  isViolated = true;
                  noteDetail = `Thỏa mãn bộ lọc rà soát: { ${rule.ifRules?.map(r => `[${r.col}] ${r.op} '${r.val}'`).join(" " + rule.ifCombine + " ")} }`;
                }
              } else {
                const thenMatches = (rule.thenRules || []).map(cond => {
                  const compareValue = cond.isFieldCompare ? String(row[cond.val] || "") : cond.val;
                  return checkValue(row[cond.col], cond.op, compareValue);
                });
                const satisfiesThen = rule.thenCombine === "OR"
                  ? thenMatches.some(v => v === true)
                  : thenMatches.every(v => v === true);

                if (rule.logicRuleMode === "conflict") {
                  if (satisfiesIf && satisfiesThen) {
                    isViolated = true;
                    noteDetail = `Mâu thuẫn đồng thời: vế NẾU và vế ĐỒNG THỜI CÓ THÊM đạt thỏa mãn.`;
                  }
                } else {
                  if (satisfiesIf && !satisfiesThen) {
                    isViolated = true;
                    noteDetail = `Có thỏa mãn vế NẾU nhưng vi phạm không đạt vế THÌ BẮT BUỘC PHẢI.`;
                  }
                }
              }
            } else if (rule.type === "ai_expression" && rule.expression) {
              try {
                const fn = new Function("row", `try { return (${rule.expression}); } catch(e) { return false; }`);
                if (fn(row) === true) {
                  isViolated = true;
                  noteDetail = `AI phân tích: Vi phạm biểu thức điều kiện logic.`;
                }
              } catch (e) {}
            }

            if (isViolated) {
              rowViolations.push({
                ruleId: rule.id,
                ruleName: rule.name,
                detail: noteDetail,
                category: rule.category || "Rà soát chung"
              });
              totalViolationsCount++;
            }
          });

          if (rowViolations.length > 0) {
            violationsList.push({
              rowIdx,
              rowNum: rowIdx + 1,
              row,
              violations: rowViolations
            });
          }
        });

        setViolationResults(violationsList);
        setScanStats({
          totalRows: mainData.length,
          violatedRows: violationsList.length,
          passedRows: mainData.length - violationsList.length,
          totalViolationsCount,
          violationRate: mainData.length > 0 ? ((violationsList.length / mainData.length) * 100).toFixed(2) + "%" : "0%"
        });
      } else {
        setViolationResults([]);
        setScanStats({
          totalRows: mainData.length,
          violatedRows: 0,
          passedRows: mainData.length,
          totalViolationsCount: 0,
          violationRate: "0%"
        });
      }
    }
  }, [mainData, rules, hasScanned]);

  // Export report to Excel
  const handleExportViolationsExcel = () => {
    if (violationResults.length === 0) return;

    const excelRows = violationResults.flatMap(item => {
      return item.violations.map((v: any) => {
        const excelItem: any = {};
        excelItem["Dòng số"] = item.rowNum;
        excelItem["Quy tắc vi phạm"] = v.ruleName;
        excelItem["Phân nhóm quy tắc"] = v.category;
        excelItem["Chi tiết vi phạm"] = v.detail;

        // Append original row columns
        columns.forEach(col => {
          if (!col.startsWith("_") && col !== "Loi_Logic") {
            excelItem[col] = item.row[col];
          }
        });

        return excelItem;
      });
    });

    if (onExportExcel) {
      onExportExcel(excelRows, `Báo_Cáo_Lỗi_Logic_Kịch_Bản_${activeTemplate || "Studio"}`);
    } else {
      const worksheet = XLSX.utils.json_to_sheet(excelRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Danh sách lỗi");
      XLSX.writeFile(workbook, `Báo_Cáo_Lỗi_Logic_Kịch_Bản_${activeTemplate || "Studio"}.xlsx`);
    }
  };

  // Filter violation results list
  const filteredViolations = useMemo(() => {
    let list = violationResults;

    if (resultSearch) {
      const q = resultSearch.toLowerCase();
      list = list.filter(item => {
        const matchesRow = Object.values(item.row).some(v => String(v).toLowerCase().includes(q));
        const matchesViolations = item.violations.some((v: any) => 
          v.ruleName.toLowerCase().includes(q) || v.detail.toLowerCase().includes(q)
        );
        return matchesRow || matchesViolations;
      });
    }

    return list;
  }, [violationResults, resultSearch]);

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm font-sans">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
        <div>
          <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
            <CheckSquare className="w-5 h-5 text-indigo-600 animate-pulse" /> KHỐI 4: RULES STUDIO (TRUNG TÂM QUẢN TRỊ QUY TẮC)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Xây dựng kịch bản rà quét thông minh đa quy tắc ("IF-THEN"). Lưu trữ dưới dạng Kịch bản quét (Templates) để tái sử dụng ngay khi nạp file mới, tạo quy tắc thủ công hoặc dịch bằng AI.
          </p>
        </div>
        
        {/* Export / Import buttons */}
        <div className="flex gap-2 self-start sm:self-center">
          <label className="flex items-center gap-1 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-lg border border-slate-300 transition-all cursor-pointer shadow-2xs">
            <Upload className="w-3.5 h-3.5" /> Nạp kịch bản (.json)
            <input type="file" accept=".json" onChange={handleImportJson} className="hidden" />
          </label>
          <button
            onClick={handleExportJson}
            className="flex items-center gap-1 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-lg border border-slate-300 transition-all cursor-pointer shadow-2xs"
          >
            <Download className="w-3.5 h-3.5" /> Xuất kịch bản (.json)
          </button>
        </div>
      </div>

      {mainData.length > 0 ? (
        <div className="space-y-6">
          {/* Live 10 Row Preview and Editor inside RulesStudio */}
          <LiveTenRowPreview
            data={mainData}
            columns={columns}
            onUpdateData={onUpdateMainData}
            highlightedIndices={violationResults.map(item => item.rowIdx)}
            highlightLabel="Dòng lỗi logic"
            title="BẢNG XEM NHANH & SỬA TRỰC TIẾP 10 DÒNG (RÀ SOÁT LOGIC)"
          />

          {/* TEMPLATE MANAGER SECTION */}
          <div className="bg-slate-50 rounded-2xl p-4.5 border border-slate-200 space-y-3">
            <h4 className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <FolderOpen className="w-4 h-4 text-amber-500" /> KỊCH BẢN QUÉT ĐANG LƯU ({templates.length})
            </h4>
            
            <div className="flex flex-wrap gap-2.5">
              {templates.map(t => (
                <button
                  key={t.name}
                  onClick={() => handleLoadTemplate(t.name)}
                  className={`px-3 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer border flex items-center gap-2 shadow-2xs active:scale-95 ${
                    activeTemplate === t.name
                      ? "bg-gradient-to-r from-amber-500 to-amber-600 text-white border-amber-600"
                      : "bg-white hover:bg-slate-100 text-slate-700 border-slate-200"
                  }`}
                >
                  <FileCheck className="w-3.5 h-3.5" />
                  {t.name}
                  <Trash2 
                    className={`w-3.5 h-3.5 ml-1 rounded-full p-0.5 hover:bg-red-500 hover:text-white transition-all`} 
                    onClick={(e) => handleDeleteTemplate(t.name, e)}
                  />
                </button>
              ))}
            </div>

            {/* Save current scenario form */}
            <div className="flex flex-col sm:flex-row items-center gap-3 pt-2">
              <input
                type="text"
                placeholder="Đặt tên để lưu kịch bản rà quét hiện tại..."
                value={newTemplateName}
                onChange={(e) => setNewTemplateName(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs outline-none focus:ring-1 focus:ring-amber-500 placeholder-slate-400"
              />
              <button
                onClick={handleSaveAsTemplate}
                className="bg-amber-500 hover:bg-amber-600 text-white font-bold text-xs py-2 px-5 rounded-xl cursor-pointer transition-all shadow-sm active:scale-95 flex items-center gap-1 border-0 shrink-0 self-stretch sm:self-auto justify-center"
              >
                <Save className="w-3.5 h-3.5" /> Lưu thành Kịch bản quét
              </button>
            </div>
          </div>

          {/* MASTER RULE GRID & SETTINGS (SIDE-BY-SIDE OR SPLIT) */}
          <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
            
            {/* Left side: Rules List Manager */}
            <div className="lg:col-span-3 space-y-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider font-mono">
                    📋 Cấu hình các quy tắc rà soát lỗi ({rules.length})
                  </h4>
                  <span className="text-[10px] text-indigo-600 font-bold">
                    ({rules.filter(r => r.active).length} đang bật)
                  </span>
                </div>
                <button
                  onClick={handleRestoreDefaults}
                  className="text-[10px] text-slate-500 hover:text-indigo-600 font-bold transition-all flex items-center gap-1 hover:underline cursor-pointer border-0 bg-transparent px-1.5 py-1"
                  title="Khôi phục danh sách các quy tắc hệ thống mặc định gốc"
                >
                  Khôi phục mặc định
                </button>
              </div>

              <div className="space-y-2.5 max-h-[420px] overflow-y-auto pr-1">
                {rules.map((rule, idx) => (
                  <div 
                    key={rule.id} 
                    className={`border rounded-xl p-3.5 transition-all flex items-start gap-3 shadow-2xs ${
                      rule.active ? "bg-white border-slate-200" : "bg-slate-50 border-slate-200 opacity-60"
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={rule.active}
                      onChange={() => handleToggleRule(rule.id)}
                      className="w-4 h-4 accent-indigo-600 cursor-pointer mt-0.5 shrink-0"
                    />
                    
                    <div className="space-y-1 flex-1">
                      <div className="flex items-start justify-between gap-2">
                        <h5 className="text-xs font-bold text-slate-800 leading-snug">{rule.name}</h5>
                        <span className="text-[9px] bg-slate-100 text-slate-500 border border-slate-200 px-1.5 py-0.5 rounded uppercase font-mono font-bold shrink-0">
                          {rule.category || "Rà soát"}
                        </span>
                      </div>
                      
                      {/* Rule formula brief */}
                      <div className="text-[10.5px] text-slate-500 font-mono leading-relaxed bg-slate-50/50 p-2 rounded-lg border border-slate-150">
                        {rule.type === "manual" ? (
                          <div className="space-y-0.5">
                            <div>• <strong>NẾU:</strong> {rule.ifRules?.map(r => `[${r.col}] ${r.op} '${r.val}'`).join(` ${rule.ifCombine} `)}</div>
                            {rule.thenRules && rule.thenRules.length > 0 ? (
                              <div>
                                • <strong>{rule.logicRuleMode === "conflict" ? "CÓ THÊM (MÂU THUẪN):" : "THÌ PHẢI:"}</strong> {rule.thenRules.map(r => `[${r.col}] ${r.op} '${r.val}'`).join(` ${rule.thenCombine} `)}
                              </div>
                            ) : (
                              <div>• <strong>Lọc trực tiếp độc lập</strong></div>
                            )}
                          </div>
                        ) : (
                          <div className="space-y-0.5">
                            <div>• <strong>AI Prompt:</strong> "{rule.prompt}"</div>
                            <div className="text-indigo-600 truncate max-w-full font-sans" title={rule.expression}>• <strong>JS:</strong> {rule.expression}</div>
                          </div>
                        )}
                      </div>
                    </div>

                    <button
                      onClick={() => handleDeleteRule(rule.id)}
                      className="text-slate-400 hover:text-red-500 p-1 rounded-lg hover:bg-red-50 transition-all shrink-0 cursor-pointer border-0"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}

                {rules.length === 0 && (
                  <div className="p-8 text-center text-slate-400 italic bg-slate-50 rounded-xl border border-dashed border-slate-250 text-xs">
                    Kịch bản quét trống! Hãy thêm quy tắc ở bên phải.
                  </div>
                )}
              </div>

              {/* Scan Trigger Button */}
              {rules.length > 0 && (
                <button
                  onClick={handleRunStudioScan}
                  disabled={isScanning}
                  className="w-full bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-bold text-xs py-3 rounded-xl cursor-pointer transition-all flex items-center justify-center gap-1.5 border-0 shadow-md active:scale-95 disabled:opacity-50"
                >
                  <Zap className="w-4 h-4 text-amber-300 animate-bounce" />
                  {isScanning ? `Đang rà quét kịch bản (${scanProgress}%)...` : "⚡ CHẠY QUÉT TOÀN BỘ KỊCH BẢN QUÉT"}
                </button>
              )}
            </div>

            {/* Right side: Rule Creator Form */}
            <div className="lg:col-span-2 bg-slate-50 border border-slate-200 p-4.5 rounded-2xl space-y-4">
              <div className="flex border-b border-slate-200">
                <button
                  onClick={() => setBuilderTab("manual")}
                  className={`flex-1 text-center py-2 text-xs font-bold border-b-2 transition-all cursor-pointer ${
                    builderTab === "manual" ? "border-indigo-600 text-indigo-700" : "border-transparent text-slate-500 hover:text-slate-800"
                  }`}
                >
                  Quy tắc thủ công
                </button>
                <button
                  onClick={() => setBuilderTab("ai")}
                  className={`flex-1 text-center py-2 text-xs font-bold border-b-2 transition-all cursor-pointer ${
                    builderTab === "ai" ? "border-indigo-600 text-indigo-700" : "border-transparent text-slate-500 hover:text-slate-800"
                  }`}
                >
                  Trợ lý dịch bằng AI
                </button>
              </div>

              {/* General Rule Fields */}
              <div className="space-y-3">
                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-700">Tên quy tắc rà soát:</label>
                  <input
                    type="text"
                    placeholder="Ví dụ: Lệch doanh thu và lao động..."
                    value={newRuleName}
                    onChange={(e) => setNewRuleName(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800"
                  />
                </div>

                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-700">Phân nhóm (Phân hệ):</label>
                  <select
                    value={newRuleCategory}
                    onChange={(e) => setNewRuleCategory(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs text-slate-800"
                  >
                    <option value="Rà soát chung">Rà soát chung</option>
                    <option value="Doanh thu &amp; Lao động">Doanh thu &amp; Lao động</option>
                    <option value="Cơ cấu hành chính">Cơ cấu hành chính</option>
                    <option value="Mã ngành VSIC">Mã ngành VSIC</option>
                    <option value="Dữ liệu biên">Dữ liệu biên</option>
                  </select>
                </div>
              </div>

              {/* TAB 1: MANUAL BUILDER */}
              {builderTab === "manual" && (
                <div className="space-y-4 animate-fade-in text-xs text-slate-700">
                  {/* IF Conditions */}
                  <div className="space-y-2">
                    <div className="flex justify-between items-center">
                      <span className="font-extrabold text-indigo-700 uppercase tracking-wider text-[10px]">VẾ NẾU (IF CONDITION)</span>
                      <button 
                        onClick={() => handleAddCondition("if")}
                        className="text-indigo-600 hover:text-indigo-800 font-bold flex items-center gap-0.5 border-0 bg-transparent cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" /> Thêm điều kiện
                      </button>
                    </div>

                    <div className="space-y-2">
                      {mIfRules.map((cond, idx) => (
                        <div key={idx} className="flex gap-2 items-center">
                          <select
                            value={cond.col}
                            onChange={(e) => {
                              const updated = [...mIfRules];
                              updated[idx].col = e.target.value;
                              setMIfRules(updated);
                            }}
                            className="bg-white border border-slate-300 rounded-lg p-1 w-28 truncate"
                          >
                            <option value="">-- Chọn cột --</option>
                            {columns.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>

                          <select
                            value={cond.op}
                            onChange={(e) => {
                              const updated = [...mIfRules];
                              updated[idx].op = e.target.value as any;
                              setMIfRules(updated);
                            }}
                            className="bg-white border border-slate-300 rounded-lg p-1 w-16"
                          >
                            <option value="==">==</option>
                            <option value="!=">!=</option>
                            <option value=">">&gt;</option>
                            <option value="<">&lt;</option>
                            <option value=">=">&gt;=</option>
                            <option value="<=">&lt;=</option>
                            <option value="chứa">chứa</option>
                            <option value="không chứa">không chứa</option>
                            <option value="trống">trống</option>
                            <option value="không trống">k.trống</option>
                          </select>

                          {cond.op !== "trống" && cond.op !== "không trống" && (
                            <input
                              type="text"
                              placeholder="Giá trị..."
                              value={cond.val}
                              onChange={(e) => {
                                const updated = [...mIfRules];
                                updated[idx].val = e.target.value;
                                setMIfRules(updated);
                              }}
                              className="bg-white border border-slate-300 rounded-lg p-1 flex-1 min-w-0"
                            />
                          )}

                          <button 
                            onClick={() => handleRemoveCondition("if", idx)}
                            className="text-slate-400 hover:text-red-500 border-0 bg-transparent cursor-pointer"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ))}
                    </div>

                    {mIfRules.length > 1 && (
                      <div className="flex items-center gap-4 pt-1">
                        <span className="text-[10px] text-slate-500 font-bold">Kết hợp logic:</span>
                        <div className="flex gap-2">
                          <label className="flex items-center gap-1 font-bold">
                            <input type="radio" checked={mIfCombine === "AND"} onChange={() => setMIfCombine("AND")} /> VÀ (AND)
                          </label>
                          <label className="flex items-center gap-1 font-bold">
                            <input type="radio" checked={mIfCombine === "OR"} onChange={() => setMIfCombine("OR")} /> HOẶC (OR)
                          </label>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* THEN Conditions */}
                  <div className="space-y-2 border-t border-slate-200 pt-3">
                    <div className="flex justify-between items-center">
                      <span className="font-extrabold text-amber-700 uppercase tracking-wider text-[10px]">VẾ THÌ PHẢI (THEN CONSTRAINT)</span>
                      <button 
                        onClick={() => handleAddCondition("then")}
                        className="text-amber-600 hover:text-amber-800 font-bold flex items-center gap-0.5 border-0 bg-transparent cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" /> Thêm điều kiện
                      </button>
                    </div>

                    {mThenRules.length > 0 ? (
                      <div className="space-y-2">
                        {mThenRules.map((cond, idx) => (
                          <div key={idx} className="flex gap-2 items-center">
                            <select
                              value={cond.col}
                              onChange={(e) => {
                                const updated = [...mThenRules];
                                updated[idx].col = e.target.value;
                                setMThenRules(updated);
                              }}
                              className="bg-white border border-slate-300 rounded-lg p-1 w-28 truncate"
                            >
                              <option value="">-- Chọn cột --</option>
                              {columns.map(c => <option key={c} value={c}>{c}</option>)}
                            </select>

                            <select
                              value={cond.op}
                              onChange={(e) => {
                                const updated = [...mThenRules];
                                updated[idx].op = e.target.value as any;
                                setMThenRules(updated);
                              }}
                              className="bg-white border border-slate-300 rounded-lg p-1 w-16"
                            >
                              <option value="==">==</option>
                              <option value="!=">!=</option>
                              <option value=">">&gt;</option>
                              <option value="<">&lt;</option>
                              <option value=">=">&gt;=</option>
                              <option value="<=">&lt;=</option>
                              <option value="chứa">chứa</option>
                              <option value="không chứa">không chứa</option>
                              <option value="trống">trống</option>
                              <option value="không trống">k.trống</option>
                            </select>

                            {cond.op !== "trống" && cond.op !== "không trống" && (
                              <input
                                type="text"
                                placeholder="Giá trị..."
                                value={cond.val}
                                onChange={(e) => {
                                  const updated = [...mThenRules];
                                  updated[idx].val = e.target.value;
                                  setMThenRules(updated);
                                }}
                                className="bg-white border border-slate-300 rounded-lg p-1 flex-1 min-w-0"
                              />
                            )}

                            <button 
                              onClick={() => handleRemoveCondition("then", idx)}
                              className="text-slate-400 hover:text-red-500 border-0 bg-transparent cursor-pointer"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ))}

                        {mThenRules.length > 1 && (
                          <div className="flex items-center gap-4 pt-1">
                            <span className="text-[10px] text-slate-500 font-bold">Kết hợp logic:</span>
                            <div className="flex gap-2">
                              <label className="flex items-center gap-1 font-bold">
                                <input type="radio" checked={mThenCombine === "AND"} onChange={() => setMThenCombine("AND")} /> VÀ (AND)
                              </label>
                              <label className="flex items-center gap-1 font-bold">
                                <input type="radio" checked={mThenCombine === "OR"} onChange={() => setMThenCombine("OR")} /> HOẶC (OR)
                              </label>
                            </div>
                          </div>
                        )}

                        {/* Logic violation Mode selection */}
                        <div className="flex items-center gap-4 pt-1">
                          <span className="text-[10px] text-slate-500 font-bold">Chế độ rà soát lỗi:</span>
                          <div className="flex gap-2">
                            <label className="flex items-center gap-1" title="Cảnh báo khi thỏa mãn NẾU nhưng KHÔNG đạt THÌ BẮT BUỘC PHẢI">
                              <input type="radio" checked={mLogicMode === "must_satisfy"} onChange={() => setMLogicMode("must_satisfy")} /> Phải thỏa mãn (Bắt buộc)
                            </label>
                            <label className="flex items-center gap-1" title="Cảnh báo khi thỏa mãn cả NẾU và đồng thời thỏa mãn vế CÓ THÊM (mâu thuẫn)">
                              <input type="radio" checked={mLogicMode === "conflict"} onChange={() => setMLogicMode("conflict")} /> Xung đột (Mâu thuẫn)
                            </label>
                          </div>
                        </div>
                      </div>
                    ) : (
                      <p className="text-[10.5px] text-slate-400 italic font-sans">
                        Để trống vế THÌ nếu bạn chỉ muốn lọc ra tất cả bản ghi thỏa mãn điều kiện NẾU độc lập.
                      </p>
                    )}
                  </div>

                  <button
                    onClick={handleSaveManualRule}
                    className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-2 rounded-xl border-0 transition-all cursor-pointer shadow-sm active:scale-95 text-center block"
                  >
                    + THÊM QUY TẮC VÀO KỊCH BẢN
                  </button>
                </div>
              )}

              {/* TAB 2: AI TRANSLATOR BUILDER */}
              {builderTab === "ai" && (
                <div className="space-y-4 animate-fade-in text-xs text-slate-700">
                  <div className="space-y-1.5">
                    <label className="font-bold flex items-center gap-1">
                      <Brain className="w-3.5 h-3.5 text-indigo-500" /> Nhập yêu cầu rà soát bằng tiếng Việt:
                    </label>
                    <textarea
                      rows={3}
                      value={aiPrompt}
                      onChange={(e) => setAiPrompt(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded-xl p-3 text-xs outline-none focus:ring-1 focus:ring-indigo-500 shadow-inner"
                      placeholder="Ví dụ: mô tả có chứa lúa gạo nhưng mã ngành lại không bắt đầu bằng 01..."
                    />
                  </div>

                  <button
                    onClick={handleTranslateAiRule}
                    disabled={isTranslatingAi}
                    className="w-full bg-indigo-100 hover:bg-indigo-200 text-indigo-800 font-bold py-2 rounded-xl border border-indigo-200 transition-all cursor-pointer flex items-center justify-center gap-1.5 disabled:opacity-50"
                  >
                    {isTranslatingAi ? "Đang nhờ AI dịch biểu thức..." : "🧠 DỊCH YÊU CẦU BẰNG AI GEMINI"}
                  </button>

                  {aiExpression && (
                    <div className="space-y-2 animate-fade-in bg-white border border-slate-200 p-3 rounded-xl shadow-2xs">
                      <span className="text-[10px] font-bold text-emerald-700 uppercase block">Biểu thức JavaScript sinh ra:</span>
                      <code className="text-[10px] font-mono text-slate-800 bg-slate-50 p-2 rounded block whitespace-pre-wrap break-all leading-relaxed border border-slate-150">
                        {aiExpression}
                      </code>
                      <button
                        onClick={handleSaveAiRule}
                        className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 rounded-xl border-0 cursor-pointer transition-all shadow-sm active:scale-95 text-center"
                      >
                        + ĐỒNG Ý &amp; CHÈN QUY TẮC NÀY
                      </button>
                    </div>
                  )}
                </div>
              )}

            </div>
          </div>

          {/* PROGRESS SCAN DISPLAY */}
          {isScanning && (
            <div className="bg-indigo-50 border border-indigo-100 rounded-2xl p-5 space-y-3 animate-pulse">
              <div className="flex justify-between items-center text-xs font-bold text-indigo-800">
                <span>Đang quét tệp dữ liệu so chéo toàn bộ quy tắc...</span>
                <span>{scanProgress}%</span>
              </div>
              <div className="w-full bg-indigo-200 rounded-full h-2.5 overflow-hidden">
                <div className="bg-indigo-600 h-2.5 rounded-full" style={{ width: `${scanProgress}%` }}></div>
              </div>
            </div>
          )}

          {/* SCAN DETAILED CONSOLIDATED DASHBOARD */}
          {hasScanned && (
            <div className="border border-slate-200 rounded-2xl p-5 space-y-5 animate-fade-in bg-white shadow-sm">
              
              {/* Scan Metrics */}
              {scanStats && (
                <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 bg-slate-50 border border-slate-200 p-4 rounded-xl text-center">
                  <div className="space-y-0.5">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Tổng số bản ghi quét</span>
                    <strong className="text-lg font-mono text-slate-800">{scanStats.totalRows.toLocaleString()}</strong>
                  </div>
                  <div className={`space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0 ${
                    scanStats.violatedRows > 0 ? "text-amber-600" : "text-emerald-600"
                  }`}>
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Số dòng vi phạm lỗi</span>
                    <strong className="text-lg font-mono font-bold">{scanStats.violatedRows.toLocaleString()}</strong>
                  </div>
                  <div className="space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Số dòng Đạt sạch lỗi</span>
                    <strong className="text-lg font-mono text-emerald-600 font-bold">{scanStats.passedRows.toLocaleString()}</strong>
                  </div>
                  <div className="space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Tổng số lỗi logic bắt giữ</span>
                    <strong className="text-lg font-mono text-indigo-700 font-bold">{scanStats.totalViolationsCount.toLocaleString()}</strong>
                  </div>
                </div>
              )}

              {/* Table list of violations */}
              <div className="space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                  <div>
                    <h5 className="text-[12px] font-extrabold text-slate-800 uppercase tracking-wider font-mono">
                      🚨 KẾT QUẢ RÀ QUÉT CHI TIẾT BẢN GHI LỖI
                    </h5>
                    <p className="text-[10.5px] text-slate-500">
                      Tất cả các hộ/bản ghi vi phạm ít nhất 1 quy tắc rà soát trong kịch bản quét được bật.
                    </p>
                  </div>

                  <div className="relative max-w-xs w-full sm:w-auto">
                    <input
                      type="text"
                      placeholder="Lọc lỗi theo từ khóa..."
                      value={resultSearch}
                      onChange={(e) => setResultSearch(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg pl-8 pr-3 py-1.5 text-xs outline-none focus:bg-white"
                    />
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  </div>
                </div>

                {filteredViolations.length === 0 ? (
                  <div className="p-10 text-center text-emerald-600 bg-emerald-50 rounded-2xl border border-dashed border-emerald-250 text-xs font-bold flex flex-col items-center justify-center gap-2">
                    <CheckCircle2 className="w-8 h-8 text-emerald-500 animate-bounce" />
                    KỊCH BẢN SẠCH SẼ! KHÔNG PHÁT HIỆN DÒNG DỮ LIỆU NÀO VI PHẠM RÀ SOÁT CỦA BẠN.
                  </div>
                ) : (
                  <div className="space-y-3 max-h-[360px] overflow-y-auto pr-1">
                    {filteredViolations.map((item, idx) => (
                      <div key={idx} className="bg-slate-50/60 border border-slate-200 rounded-xl p-3 flex flex-col sm:flex-row sm:items-start justify-between gap-4">
                        <div className="space-y-2 flex-1">
                          <div className="flex items-center gap-2">
                            <span className="bg-slate-200 text-slate-800 border border-slate-300 px-2 py-0.5 rounded font-mono font-bold text-[10.5px]">
                              Dòng #{item.rowNum}
                            </span>
                            
                            {/* Short data preview */}
                            <span className="text-[10.5px] text-slate-500 truncate max-w-xs block" title={JSON.stringify(item.row)}>
                              [ {columns.slice(0, 3).map(c => `${c}: ${item.row[c]}`).join(" | ")} ]
                            </span>
                          </div>

                          {/* Violations nested list */}
                          <div className="space-y-1.5 pl-2 border-l-2 border-amber-500">
                            {item.violations.map((v: any, vIdx: number) => (
                              <div key={vIdx} className="text-xs text-slate-800 leading-snug">
                                <span className="bg-amber-100 text-amber-800 border border-amber-200 px-1 py-0.25 rounded font-bold text-[9.5px] uppercase mr-1.5 shrink-0">
                                  {v.ruleName}
                                </span>
                                <span className="text-slate-600">{v.detail}</span>
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Focus filter button */}
                        {onFilterRows && (
                          <button
                            onClick={() => onFilterRows([item.rowIdx], `Dòng lỗi kịch bản #${item.rowNum}`)}
                            className="bg-white hover:bg-indigo-600 text-indigo-700 hover:text-white px-2.5 py-1.5 rounded-lg text-[10.5px] font-bold border border-slate-300 transition-all cursor-pointer shadow-3xs hover:shadow-2xs active:scale-95 whitespace-nowrap self-start sm:self-center"
                          >
                            🔍 Xem trên bảng
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                )}

                {/* Bottom Actions for results */}
                {violationResults.length > 0 && (
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-3 border-t border-slate-100">
                    <div className="text-[10.5px] text-slate-500">
                      * Nhấp nút lọc tập trung để nạp toàn bộ danh sách dòng vi phạm lỗi lên bảng chính phục vụ sửa thủ công.
                    </div>
                    
                    <div className="flex gap-2.5">
                      <button
                        onClick={handleExportViolationsExcel}
                        className="bg-white hover:bg-slate-100 text-slate-750 font-bold text-xs py-2 px-4 rounded-xl cursor-pointer border border-slate-300 transition-all flex items-center gap-1 shadow-3xs"
                      >
                        <Download className="w-3.5 h-3.5" /> Xuất tệp báo cáo (.xlsx)
                      </button>

                      {onFilterRows && (
                        <button
                          onClick={() => {
                            const indices = violationResults.map(v => v.rowIdx);
                            onFilterRows(indices, `Tất cả ${indices.length} dòng lỗi kịch bản quét`);
                          }}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-2 px-5 rounded-xl cursor-pointer border-0 transition-all flex items-center gap-1 shadow-md active:scale-95"
                        >
                          👁️ LỌC TẬP TRUNG TOÀN TẬP TRÊN BẢNG CHÍNH
                        </button>
                      )}
                    </div>
                  </div>
                )}

              </div>

            </div>
          )}
        </div>
      ) : (
        <div className="p-10 text-center text-slate-400 bg-slate-50 border border-dashed border-slate-300 rounded-2xl font-sans">
          <Info className="w-8 h-8 text-slate-300 mx-auto mb-2" />
          Chưa có dữ liệu nguồn được nạp. Hãy nạp file Excel/CSV ở trang chủ để kích hoạt chức năng rà quét.
        </div>
      )}
    </div>
  );
}
